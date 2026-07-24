Attribute VB_Name = "BatchInjectStandardStyles"
Option Explicit

' 所有样式统一使用：中文宋体，英文及其他西文 Times New Roman。
Private Const FONT_EAST_ASIA As String = "宋体"
Private Const FONT_LATIN As String = "Times New Roman"

' 中文字号对应的磅值：一号 26 磅，小四 12 磅，五号 10.5 磅。
Private Const COVER_FONT_SIZE As Single = 26
Private Const BODY_FONT_SIZE As Single = 12
Private Const TABLE_FONT_SIZE As Single = 10.5

' 新建的自定义段落样式名称。
Private Const STYLE_COVER_TITLE As String = "封面标题"
Private Const STYLE_TABLE_TITLE As String = "表格标题"
Private Const STYLE_TABLE_CONTENT As String = "表格内容"

Public Sub BatchInjectStandardStyles()
    Dim folderPath As String
    Dim fileName As String
    Dim filePath As Variant
    Dim wordFiles As Collection
    Dim result As String
    Dim detail As String
    Dim successCount As Long
    Dim skippedCount As Long
    Dim failedCount As Long
    Dim currentIndex As Long

    Dim oldScreenUpdating As Boolean
    Dim oldDisplayAlerts As WdAlertLevel
    Dim oldAutomationSecurity As Long
    Dim oldStatusBar As Variant
    Dim settingsCaptured As Boolean
    Dim fatalMessage As String

    On Error GoTo FatalError

    folderPath = PickTargetFolder()
    If Len(folderPath) = 0 Then Exit Sub

    If Right$(folderPath, 1) <> Application.PathSeparator Then
        folderPath = folderPath & Application.PathSeparator
    End If

    Set wordFiles = New Collection

    ' 只收集当前文件夹，不递归处理子文件夹。
    fileName = Dir$(folderPath & "*.*", _
                    vbNormal Or vbReadOnly Or vbHidden Or vbSystem)

    Do While Len(fileName) > 0
        If Left$(fileName, 2) <> "~$" Then
            If IsSupportedWordFile(fileName) Then
                wordFiles.Add folderPath & fileName
            End If
        End If
        fileName = Dir$
    Loop

    If wordFiles.Count = 0 Then
        MsgBox "所选文件夹中没有找到可处理的 Word 文档。", _
               vbExclamation, "没有 Word 文档"
        Exit Sub
    End If

    If MsgBox( _
        "即将原位修改并保存 " & wordFiles.Count & " 个 Word 文档。" & _
        vbCrLf & vbCrLf & folderPath & vbCrLf & vbCrLf & _
        "建议先备份文件。是否继续？", _
        vbQuestion + vbYesNo + vbDefaultButton2, _
        "确认注入标准样式") <> vbYes Then
        Exit Sub
    End If

    oldScreenUpdating = Application.ScreenUpdating
    oldDisplayAlerts = Application.DisplayAlerts
    oldAutomationSecurity = Application.AutomationSecurity
    oldStatusBar = Application.StatusBar
    settingsCaptured = True

    Application.ScreenUpdating = False
    Application.DisplayAlerts = wdAlertsNone

    ' 禁止待处理的 .docm 文档执行 AutoOpen 等宏。
    Application.AutomationSecurity = msoAutomationSecurityForceDisable

    For Each filePath In wordFiles
        currentIndex = currentIndex + 1
        Application.StatusBar = _
            "正在注入标准样式 " & currentIndex & "/" & wordFiles.Count & _
            "：" & CStr(filePath)
        DoEvents

        result = vbNullString
        detail = vbNullString

        ProcessOneDocument CStr(filePath), result, detail

        Select Case result
            Case "成功"
                successCount = successCount + 1
            Case "跳过"
                skippedCount = skippedCount + 1
            Case Else
                failedCount = failedCount + 1
        End Select

        If result <> "成功" Then
            Debug.Print result & " | " & CStr(filePath) & " | " & detail
        End If
    Next filePath

CleanExit:
    On Error Resume Next
    If settingsCaptured Then
        Application.ScreenUpdating = oldScreenUpdating
        Application.DisplayAlerts = oldDisplayAlerts
        Application.AutomationSecurity = oldAutomationSecurity
        Application.StatusBar = oldStatusBar
    End If
    On Error GoTo 0

    If Len(fatalMessage) > 0 Then
        MsgBox fatalMessage, vbCritical, "批量注入已中止"
    Else
        MsgBox "标准样式及页边距注入完成。" & vbCrLf & vbCrLf & _
               "成功：" & successCount & vbCrLf & _
               "跳过：" & skippedCount & vbCrLf & _
               "失败：" & failedCount & vbCrLf & vbCrLf & _
               "跳过或失败的详细信息可在 VBA 立即窗口中查看。", _
               IIf(failedCount = 0, vbInformation, vbExclamation), _
               "批量注入完成"
    End If
    Exit Sub

FatalError:
    fatalMessage = "处理过程中发生错误：" & vbCrLf & _
                   CStr(Err.Number) & " - " & Err.Description
    Resume CleanExit
End Sub

Private Function PickTargetFolder() As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "请选择需要注入标准样式的 Word 文件夹"

        If .Show <> -1 Then
            PickTargetFolder = vbNullString
        Else
            PickTargetFolder = .SelectedItems(1)
        End If
    End With
End Function

Private Sub ProcessOneDocument(ByVal filePath As String, _
                               ByRef result As String, _
                               ByRef detail As String)
    Dim doc As Document
    Dim documentWasOpened As Boolean
    Dim previousTrackRevisions As Boolean
    Dim trackRevisionsCaptured As Boolean

    On Error GoTo DocumentError

    ' 如果本宏保存在所选文件夹中的 .docm 内，避免修改正在运行的文档。
    If StrComp(filePath, ThisDocument.FullName, vbTextCompare) = 0 Then
        result = "跳过"
        detail = "该文件是当前宏所在文档"
        Exit Sub
    End If

    Set doc = Documents.Open( _
        FileName:=filePath, _
        ConfirmConversions:=False, _
        ReadOnly:=False, _
        AddToRecentFiles:=False, _
        Visible:=False, _
        OpenAndRepair:=True)
    documentWasOpened = True

    If doc.ReadOnly Then
        result = "跳过"
        detail = "文档以只读方式打开"
        GoTo CloseWithoutSaving
    End If

    If doc.ProtectionType <> wdNoProtection Then
        result = "跳过"
        detail = "文档受保护，ProtectionType=" & CStr(doc.ProtectionType)
        GoTo CloseWithoutSaving
    End If

    previousTrackRevisions = doc.TrackRevisions
    trackRevisionsCaptured = True
    doc.TrackRevisions = False

    ApplyStandardStyles doc
    ApplyTableContentStyleToExistingTables doc
    ApplyStandardMargins doc

    doc.TrackRevisions = previousTrackRevisions
    doc.Save
    doc.Close SaveChanges:=wdDoNotSaveChanges
    documentWasOpened = False

    result = "成功"
    detail = "封面/正文/标题/表格样式、现有表格内容及各节页边距已更新"
    Exit Sub

CloseWithoutSaving:
    doc.Close SaveChanges:=wdDoNotSaveChanges
    documentWasOpened = False
    Exit Sub

DocumentError:
    result = "失败"
    detail = CStr(Err.Number) & " - " & Err.Description

    On Error Resume Next
    If documentWasOpened Then
        If trackRevisionsCaptured Then
            doc.TrackRevisions = previousTrackRevisions
        End If
        doc.Close SaveChanges:=wdDoNotSaveChanges
    End If
    On Error GoTo 0
End Sub

Private Sub ApplyStandardStyles(ByVal doc As Document)
    Dim headingStyleIds As Variant
    Dim targetStyle As Style
    Dim index As Long

    CreateOrUpdateCoverTitleStyle doc

    ' 正文：中文宋体、西文 TNR、小四、不加粗、1.5 倍行距、段后 1 行。
    Set targetStyle = doc.Styles(wdStyleNormal)
    ApplyBilingualStyleFont targetStyle, BODY_FONT_SIZE, False

    With targetStyle.ParagraphFormat
        .LineSpacingRule = wdLineSpace1pt5
        .SpaceBeforeAuto = False
        .SpaceAfterAuto = False
        .LineUnitBefore = 0
        .LineUnitAfter = 1
    End With
    targetStyle.NextParagraphStyle = wdStyleNormal

    headingStyleIds = Array( _
        wdStyleHeading1, _
        wdStyleHeading2, _
        wdStyleHeading3, _
        wdStyleHeading4, _
        wdStyleHeading5, _
        wdStyleHeading6, _
        wdStyleHeading7, _
        wdStyleHeading8, _
        wdStyleHeading9)

    For index = LBound(headingStyleIds) To UBound(headingStyleIds)
        Set targetStyle = doc.Styles(headingStyleIds(index))

        ' 一级和二级标题加粗；三级至九级标题不加粗，字号均为小四。
        ApplyBilingualStyleFont targetStyle, BODY_FONT_SIZE, (index <= 1)

        With targetStyle.ParagraphFormat
            .SpaceBeforeAuto = False
            .SpaceAfterAuto = False
            .LineUnitBefore = 1
            .LineUnitAfter = 1
        End With
        targetStyle.NextParagraphStyle = wdStyleNormal
    Next index

    CreateOrUpdateTableTitleStyle doc
    CreateOrUpdateTableContentStyle doc
End Sub

Private Sub CreateOrUpdateCoverTitleStyle(ByVal doc As Document)
    Dim targetStyle As Style

    Set targetStyle = GetOrCreateParagraphStyle(doc, STYLE_COVER_TITLE)

    With targetStyle
        .BaseStyle = wdStyleNormal
        .AutomaticallyUpdate = False
        .NextParagraphStyle = STYLE_COVER_TITLE
    End With

    ApplyBilingualStyleFont targetStyle, COVER_FONT_SIZE, True

    With targetStyle.ParagraphFormat
        .Alignment = wdAlignParagraphCenter
        .LineSpacingRule = wdLineSpaceSingle
        .SpaceBeforeAuto = False
        .SpaceAfterAuto = False
        .LineUnitBefore = 0
        .LineUnitAfter = 0
    End With
End Sub

Private Sub CreateOrUpdateTableTitleStyle(ByVal doc As Document)
    Dim targetStyle As Style

    Set targetStyle = GetOrCreateParagraphStyle(doc, STYLE_TABLE_TITLE)

    With targetStyle
        .BaseStyle = wdStyleNormal
        .AutomaticallyUpdate = False
        .NextParagraphStyle = wdStyleNormal
    End With

    ApplyBilingualStyleFont targetStyle, TABLE_FONT_SIZE, True

    With targetStyle.ParagraphFormat
        .Alignment = wdAlignParagraphCenter
        .LineSpacingRule = wdLineSpaceSingle
        .SpaceBeforeAuto = False
        .SpaceAfterAuto = False
        .LineUnitBefore = 0
        ' 采用允许范围中的 0 行，使表格标题紧贴表格。
        .LineUnitAfter = 0
    End With
End Sub

Private Sub CreateOrUpdateTableContentStyle(ByVal doc As Document)
    Dim targetStyle As Style

    Set targetStyle = GetOrCreateParagraphStyle(doc, STYLE_TABLE_CONTENT)

    With targetStyle
        .BaseStyle = wdStyleNormal
        .AutomaticallyUpdate = False
        .NextParagraphStyle = STYLE_TABLE_CONTENT
    End With

    ApplyBilingualStyleFont targetStyle, TABLE_FONT_SIZE, False

    With targetStyle.ParagraphFormat
        .LineSpacingRule = wdLineSpaceSingle
        .SpaceBeforeAuto = False
        .SpaceAfterAuto = False
        .SpaceBefore = 0
        .SpaceAfter = 0
    End With
End Sub

Private Function GetOrCreateParagraphStyle(ByVal doc As Document, _
                                           ByVal styleName As String) As Style
    Dim targetStyle As Style

    On Error Resume Next
    Set targetStyle = doc.Styles(styleName)
    On Error GoTo 0

    If targetStyle Is Nothing Then
        Set targetStyle = doc.Styles.Add( _
            Name:=styleName, _
            Type:=wdStyleTypeParagraph)
    ElseIf targetStyle.Type <> wdStyleTypeParagraph And _
           targetStyle.Type <> wdStyleTypeLinked Then
        Err.Raise vbObjectError + 2101, _
                  "GetOrCreateParagraphStyle", _
                  "名称为“" & styleName & "”的现有样式不是段落样式。"
    End If

    ' 某些旧版 Word 不支持 QuickStyle；失败时不影响样式本身。
    On Error Resume Next
    targetStyle.QuickStyle = True
    On Error GoTo 0

    Set GetOrCreateParagraphStyle = targetStyle
End Function

Private Sub ApplyBilingualStyleFont(ByVal targetStyle As Style, _
                                    ByVal fontSize As Single, _
                                    ByVal useBold As Boolean)
    With targetStyle.Font
        .NameAscii = FONT_LATIN
        .NameOther = FONT_LATIN
        .NameBi = FONT_LATIN
        .NameFarEast = FONT_EAST_ASIA
        .Size = fontSize
        .Bold = useBold
    End With
End Sub

Private Sub ApplyTableContentStyleToExistingTables(ByVal doc As Document)
    Dim docTable As Table
    Dim tableParagraph As Paragraph
    Dim tableContentStyle As Style

    Set tableContentStyle = doc.Styles(STYLE_TABLE_CONTENT)

    For Each docTable In doc.Tables
        For Each tableParagraph In docTable.Range.Paragraphs
            tableParagraph.Range.Style = tableContentStyle

            ' 清除样式可能无法覆盖的直接字体格式，确保现有内容也符合要求。
            With tableParagraph.Range.Font
                .NameAscii = FONT_LATIN
                .NameOther = FONT_LATIN
                .NameBi = FONT_LATIN
                .NameFarEast = FONT_EAST_ASIA
                .Size = TABLE_FONT_SIZE
                .Bold = False
            End With

            With tableParagraph.Format
                .LineSpacingRule = wdLineSpaceSingle
                .SpaceBeforeAuto = False
                .SpaceAfterAuto = False
                .SpaceBefore = 0
                .SpaceAfter = 0
            End With
        Next tableParagraph
    Next docTable
End Sub

Public Sub AutoSetMarginsByOrientation()
    On Error GoTo MarginError

    If Documents.Count = 0 Then
        MsgBox "当前没有打开的 Word 文档。", _
               vbExclamation, "无法设置页边距"
        Exit Sub
    End If

    ApplyStandardMargins ActiveDocument

    MsgBox "全篇文档已按每个节的横竖方向完成页边距设置！", _
           vbInformation, "完成"
    Exit Sub

MarginError:
    MsgBox "设置页边距失败：" & vbCrLf & _
           CStr(Err.Number) & " - " & Err.Description, _
           vbCritical, "设置失败"
End Sub

Private Sub ApplyStandardMargins(ByVal doc As Document)
    Dim docSection As Section

    ' 逐节判断版式，支持横竖版混排文档。
    For Each docSection In doc.Sections
        With docSection.PageSetup
            If .Orientation = wdOrientPortrait Then
                .TopMargin = CentimetersToPoints(2.5)
                .BottomMargin = CentimetersToPoints(2.5)
                .LeftMargin = CentimetersToPoints(2)
                .RightMargin = CentimetersToPoints(2)
            ElseIf .Orientation = wdOrientLandscape Then
                .TopMargin = CentimetersToPoints(1.2)
                .BottomMargin = CentimetersToPoints(1.2)
                .LeftMargin = CentimetersToPoints(1.2)
                .RightMargin = CentimetersToPoints(1.2)
            End If
        End With
    Next docSection
End Sub

Private Function IsSupportedWordFile(ByVal fileName As String) As Boolean
    Dim dotPosition As Long
    Dim extensionName As String

    dotPosition = InStrRev(fileName, ".")
    If dotPosition = 0 Then Exit Function

    extensionName = LCase$(Mid$(fileName, dotPosition + 1))

    Select Case extensionName
        Case "doc", "docx", "docm", "docb"
            IsSupportedWordFile = True
    End Select
End Function
