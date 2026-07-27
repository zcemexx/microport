Attribute VB_Name = "BatchInjectStandardStyles"
Option Explicit

' ASCII-only VBA source file.
' This file can be imported into Word VBA without UTF-8/ANSI mojibake.
'
' The East Asian font name is built with Unicode code points:
'   U+5B8B U+4F53 = Songti (Chinese font)

Private Const FONT_LATIN As String = "Times New Roman"

' Chinese font sizes in points: No. 1 = 26, Small No. 4 = 12, No. 5 = 10.5.
Private Const COVER_FONT_SIZE As Single = 26
Private Const BODY_FONT_SIZE As Single = 12
Private Const TABLE_FONT_SIZE As Single = 10.5

' Custom paragraph style names are ASCII to remain import-safe in Word VBA.
Private Const STYLE_COVER_TITLE As String = "Cover Title"
Private Const STYLE_TABLE_TITLE As String = "Table Title"
Private Const STYLE_TABLE_CONTENT As String = "Table Content"

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

    ' Current folder only; subfolders are not included.
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
        MsgBox "No supported Word documents were found in the selected folder.", _
               vbExclamation, "No documents"
        Exit Sub
    End If

    If MsgBox( _
        "This will modify and save " & wordFiles.Count & _
        " Word document(s) in place." & vbCrLf & vbCrLf & _
        folderPath & vbCrLf & vbCrLf & _
        "Back up the files before continuing. Continue?", _
        vbQuestion + vbYesNo + vbDefaultButton2, _
        "Confirm batch formatting") <> vbYes Then
        Exit Sub
    End If

    oldScreenUpdating = Application.ScreenUpdating
    oldDisplayAlerts = Application.DisplayAlerts
    oldAutomationSecurity = Application.AutomationSecurity
    oldStatusBar = Application.StatusBar
    settingsCaptured = True

    Application.ScreenUpdating = False
    Application.DisplayAlerts = wdAlertsNone

    ' Prevent AutoOpen and other macros in processed .docm files from running.
    Application.AutomationSecurity = msoAutomationSecurityForceDisable

    For Each filePath In wordFiles
        currentIndex = currentIndex + 1
        Application.StatusBar = _
            "Applying styles " & currentIndex & "/" & wordFiles.Count & ": " & _
            CStr(filePath)
        DoEvents

        result = vbNullString
        detail = vbNullString

        ProcessOneDocument CStr(filePath), result, detail

        Select Case result
            Case "Success"
                successCount = successCount + 1
            Case "Skipped"
                skippedCount = skippedCount + 1
            Case Else
                failedCount = failedCount + 1
        End Select

        If result <> "Success" Then
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
        MsgBox fatalMessage, vbCritical, "Batch formatting stopped"
    Else
        MsgBox "Batch formatting finished." & vbCrLf & vbCrLf & _
               "Success: " & successCount & vbCrLf & _
               "Skipped: " & skippedCount & vbCrLf & _
               "Failed: " & failedCount & vbCrLf & vbCrLf & _
               "Details for skipped or failed files are in the Immediate window.", _
               IIf(failedCount = 0, vbInformation, vbExclamation), _
               "Batch formatting complete"
    End If
    Exit Sub

FatalError:
    fatalMessage = "An error occurred:" & vbCrLf & _
                   CStr(Err.Number) & " - " & Err.Description
    Resume CleanExit
End Sub

Private Function PickTargetFolder() As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select the folder containing Word documents"

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

    ' Avoid modifying the document that contains this macro.
    If StrComp(filePath, ThisDocument.FullName, vbTextCompare) = 0 Then
        result = "Skipped"
        detail = "The file is the document containing this macro."
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
        result = "Skipped"
        detail = "The document opened as read-only."
        GoTo CloseWithoutSaving
    End If

    If doc.ProtectionType <> wdNoProtection Then
        result = "Skipped"
        detail = "The document is protected."
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

    result = "Success"
    detail = "Styles, table content, and margins were updated."
    Exit Sub

CloseWithoutSaving:
    doc.Close SaveChanges:=wdDoNotSaveChanges
    documentWasOpened = False
    Exit Sub

DocumentError:
    result = "Failed"
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
    Dim styleIndex As Long

    CreateOrUpdateCoverTitleStyle doc

    ' Body: Songti for East Asian text, TNR for Latin text, 12 pt, 1.5 lines.
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

    For styleIndex = LBound(headingStyleIds) To UBound(headingStyleIds)
        Set targetStyle = doc.Styles(headingStyleIds(styleIndex))

        ' Heading 1 and 2 are bold. Heading 3 through 9 are not bold.
        ApplyBilingualStyleFont targetStyle, BODY_FONT_SIZE, (styleIndex <= 1)

        With targetStyle.ParagraphFormat
            .SpaceBeforeAuto = False
            .SpaceAfterAuto = False
            .LineUnitBefore = 1
            .LineUnitAfter = 1
        End With
        targetStyle.NextParagraphStyle = wdStyleNormal
    Next styleIndex

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
                  "The existing style is not a paragraph style: " & styleName
    End If

    ' Some old Word versions do not support QuickStyle.
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
        .NameFarEast = EastAsianFontName()
        .Size = fontSize
        .Bold = useBold
    End With
End Sub

Private Function EastAsianFontName() As String
    EastAsianFontName = ChrW(&H5B8B) & ChrW(&H4F53)
End Function

Private Sub ApplyTableContentStyleToExistingTables(ByVal doc As Document)
    Dim docTable As Table
    Dim tableParagraph As Paragraph
    Dim tableContentStyle As Style

    Set tableContentStyle = doc.Styles(STYLE_TABLE_CONTENT)

    For Each docTable In doc.Tables
        For Each tableParagraph In docTable.Range.Paragraphs
            tableParagraph.Range.Style = tableContentStyle

            ' Apply direct values too, so prior direct formatting cannot override the style.
            With tableParagraph.Range.Font
                .NameAscii = FONT_LATIN
                .NameOther = FONT_LATIN
                .NameBi = FONT_LATIN
                .NameFarEast = EastAsianFontName()
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
        MsgBox "No Word document is open.", _
               vbExclamation, "Cannot set margins"
        Exit Sub
    End If

    ApplyStandardMargins ActiveDocument

    MsgBox "Margins were set for every section based on its orientation.", _
           vbInformation, "Finished"
    Exit Sub

MarginError:
    MsgBox "Failed to set margins:" & vbCrLf & _
           CStr(Err.Number) & " - " & Err.Description, _
           vbCritical, "Error"
End Sub

Private Sub ApplyStandardMargins(ByVal doc As Document)
    Dim docSection As Section

    ' Each section is processed independently, supporting mixed orientations.
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
