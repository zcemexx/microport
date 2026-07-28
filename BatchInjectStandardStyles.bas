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
Private Const OUTPUT_FOLDER_NAME As String = "formatted_output"

Public Sub BatchInjectStandardStyles()
    Dim folderPath As String
    Dim filePath As Variant
    Dim outputPath As String
    Dim outputRootPath As String
    Dim wordFiles As Collection
    Dim fileSystem As Object
    Dim result As String
    Dim detail As String
    Dim successCount As Long
    Dim skippedCount As Long
    Dim failedCount As Long
    Dim currentIndex As Long

    Dim oldScreenUpdating As Boolean
    Dim oldDisplayAlerts As WdAlertLevel
    Dim oldAutomationSecurity As Long
    Dim settingsCaptured As Boolean
    Dim fatalMessage As String

    On Error GoTo FatalError

    folderPath = PickTargetFolder()
    If Len(folderPath) = 0 Then Exit Sub

    Set wordFiles = New Collection
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    folderPath = EnsureTrailingSeparator(fileSystem.GetAbsolutePathName(folderPath))
    outputRootPath = fileSystem.BuildPath(folderPath, OUTPUT_FOLDER_NAME)

    CollectWordFilesRecursive fileSystem.GetFolder(folderPath), _
                              outputRootPath, _
                              wordFiles, _
                              fileSystem

    If wordFiles.Count = 0 Then
        MsgBox "No supported Word documents were found in the selected folder.", _
               vbExclamation, "No documents"
        Exit Sub
    End If

    EnsureFolderExists outputRootPath, fileSystem

    oldScreenUpdating = Application.ScreenUpdating
    oldDisplayAlerts = Application.DisplayAlerts
    oldAutomationSecurity = Application.AutomationSecurity
    settingsCaptured = True

    Application.ScreenUpdating = False
    Application.DisplayAlerts = wdAlertsNone

    ' Prevent AutoOpen and other macros in processed .docm files from running.
    Application.AutomationSecurity = msoAutomationSecurityForceDisable

    For Each filePath In wordFiles
        currentIndex = currentIndex + 1
        DoEvents

        result = vbNullString
        detail = vbNullString

        outputPath = fileSystem.BuildPath( _
            outputRootPath, _
            RelativePathFromRoot(CStr(filePath), folderPath))

        ProcessOneDocument CStr(filePath), outputPath, fileSystem, result, detail

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
    End If
    On Error GoTo 0

    If Len(fatalMessage) > 0 Then
        MsgBox fatalMessage, vbCritical, "Batch formatting stopped"
    Else
        MsgBox "Batch formatting finished." & vbCrLf & vbCrLf & _
               "Success: " & successCount & vbCrLf & _
               "Skipped: " & skippedCount & vbCrLf & _
               "Failed: " & failedCount & vbCrLf & vbCrLf & _
               "Output folder: " & outputRootPath & vbCrLf & vbCrLf & _
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

Private Sub CollectWordFilesRecursive(ByVal sourceFolder As Object, _
                                      ByVal outputRootPath As String, _
                                      ByRef wordFiles As Collection, _
                                      ByVal fileSystem As Object)
    Dim sourceFile As Object
    Dim subFolder As Object

    For Each sourceFile In sourceFolder.Files
        If Left$(sourceFile.Name, 2) <> "~$" Then
            If IsSupportedWordFile(sourceFile.Name) Then
                wordFiles.Add sourceFile.Path
            End If
        End If
    Next sourceFile

    For Each subFolder In sourceFolder.SubFolders
        If StrComp(subFolder.Path, outputRootPath, vbTextCompare) <> 0 Then
            CollectWordFilesRecursive subFolder, outputRootPath, wordFiles, fileSystem
        End If
    Next subFolder
End Sub

Private Function EnsureTrailingSeparator(ByVal folderPath As String) As String
    If Right$(folderPath, 1) = Application.PathSeparator Then
        EnsureTrailingSeparator = folderPath
    Else
        EnsureTrailingSeparator = folderPath & Application.PathSeparator
    End If
End Function

Private Function RelativePathFromRoot(ByVal sourcePath As String, _
                                      ByVal rootPath As String) As String
    If StrComp(Left$(sourcePath, Len(rootPath)), rootPath, vbTextCompare) <> 0 Then
        Err.Raise vbObjectError + 2102, _
                  "RelativePathFromRoot", _
                  "The source path is outside the selected root folder."
    End If

    RelativePathFromRoot = Mid$(sourcePath, Len(rootPath) + 1)
End Function

Private Sub EnsureFolderExists(ByVal folderPath As String, ByVal fileSystem As Object)
    Dim parentPath As String

    If fileSystem.FolderExists(folderPath) Then Exit Sub

    parentPath = fileSystem.GetParentFolderName(folderPath)
    If Len(parentPath) = 0 Then
        Err.Raise vbObjectError + 2103, _
                  "EnsureFolderExists", _
                  "Cannot determine the parent folder for: " & folderPath
    End If

    EnsureFolderExists parentPath, fileSystem
    fileSystem.CreateFolder folderPath
End Sub

Private Function BuildTemporaryOutputPath(ByVal outputPath As String, _
                                          ByVal fileSystem As Object) As String
    Dim outputFolder As String
    Dim extensionName As String
    Dim temporaryName As String

    outputFolder = fileSystem.GetParentFolderName(outputPath)
    extensionName = fileSystem.GetExtensionName(outputPath)
    temporaryName = "~format_" & Replace(fileSystem.GetTempName, ".", "_")

    If Len(extensionName) > 0 Then
        temporaryName = temporaryName & "." & extensionName
    End If

    BuildTemporaryOutputPath = fileSystem.BuildPath(outputFolder, temporaryName)
End Function

Private Sub ProcessOneDocument(ByVal sourcePath As String, _
                               ByVal outputPath As String, _
                               ByVal fileSystem As Object, _
                               ByRef result As String, _
                               ByRef detail As String)
    Dim doc As Document
    Dim documentWasOpened As Boolean
    Dim previousTrackRevisions As Boolean
    Dim trackRevisionsCaptured As Boolean
    Dim sourceSaveFormat As Long
    Dim temporaryOutputPath As String

    On Error GoTo DocumentError

    ' Avoid modifying the document that contains this macro.
    If StrComp(sourcePath, ThisDocument.FullName, vbTextCompare) = 0 Then
        result = "Skipped"
        detail = "The file is the document containing this macro."
        Exit Sub
    End If

    Set doc = Documents.Open( _
        FileName:=sourcePath, _
        ConfirmConversions:=False, _
        ReadOnly:=False, _
        AddToRecentFiles:=False, _
        Visible:=False, _
        OpenAndRepair:=False)
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
    sourceSaveFormat = doc.SaveFormat
    doc.TrackRevisions = False

    ApplyStandardStyles doc
    ApplyStyleShortcuts doc
    ApplyHeadingFormattingToExistingParagraphs doc
    ApplyStandardMargins doc

    doc.TrackRevisions = previousTrackRevisions
    EnsureFolderExists fileSystem.GetParentFolderName(outputPath), fileSystem
    temporaryOutputPath = BuildTemporaryOutputPath(outputPath, fileSystem)
    doc.SaveAs2 FileName:=temporaryOutputPath, _
                FileFormat:=sourceSaveFormat, _
                AddToRecentFiles:=False
    doc.Close SaveChanges:=wdDoNotSaveChanges
    documentWasOpened = False

    If fileSystem.FileExists(outputPath) Then
        fileSystem.DeleteFile outputPath, True
    End If
    fileSystem.MoveFile temporaryOutputPath, outputPath

    result = "Success"
    detail = "Formatted copy was saved to: " & outputPath
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
    If Len(temporaryOutputPath) > 0 Then
        If fileSystem.FileExists(temporaryOutputPath) Then
            fileSystem.DeleteFile temporaryOutputPath, True
        End If
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
            .LineSpacingRule = wdLineSpace1pt5
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

Private Sub ApplyStyleShortcuts(ByVal doc As Document)
    Dim previousContext As Object
    Dim errorNumber As Long
    Dim errorDescription As String

    Set previousContext = Application.CustomizationContext
    On Error GoTo ShortcutError

    Application.CustomizationContext = doc

    AddStyleShortcut doc.Styles(wdStyleHeading1).NameLocal, _
                     Application.BuildKeyCode(wdKeyAlt, wdKey1)
    AddStyleShortcut doc.Styles(wdStyleHeading2).NameLocal, _
                     Application.BuildKeyCode(wdKeyAlt, wdKey2)
    AddStyleShortcut doc.Styles(wdStyleHeading3).NameLocal, _
                     Application.BuildKeyCode(wdKeyAlt, wdKey3)
    AddStyleShortcut STYLE_COVER_TITLE, _
                     Application.BuildKeyCode(wdKeyAlt, wdKeyQ)
    AddStyleShortcut STYLE_TABLE_TITLE, _
                     Application.BuildKeyCode(wdKeyAlt, wdKeyW)
    AddStyleShortcut STYLE_TABLE_CONTENT, _
                     Application.BuildKeyCode(wdKeyAlt, wdKeyE)

    Application.CustomizationContext = previousContext
    Exit Sub

ShortcutError:
    errorNumber = Err.Number
    errorDescription = Err.Description

    On Error Resume Next
    Application.CustomizationContext = previousContext
    On Error GoTo 0

    Err.Raise errorNumber, "ApplyStyleShortcuts", errorDescription
End Sub

Private Sub AddStyleShortcut(ByVal styleName As String, ByVal keyCode As Long)
    Application.KeyBindings.Add _
        KeyCategory:=wdKeyCategoryStyle, _
        Command:=styleName, _
        KeyCode:=keyCode
End Sub

Private Sub ApplyHeadingFormattingToExistingParagraphs(ByVal doc As Document)
    Dim documentParagraph As Paragraph

    For Each documentParagraph In doc.Paragraphs
        If IsHeadingParagraph(documentParagraph) Then
            With documentParagraph.Format
                .LineSpacingRule = wdLineSpace1pt5
                .SpaceBeforeAuto = False
                .SpaceAfterAuto = False
                .LineUnitBefore = 1
                .LineUnitAfter = 1
            End With
        End If
    Next documentParagraph
End Sub

Private Function IsHeadingParagraph(ByVal documentParagraph As Paragraph) As Boolean
    Select Case documentParagraph.OutlineLevel
        Case wdOutlineLevel1, _
             wdOutlineLevel2, _
             wdOutlineLevel3, _
             wdOutlineLevel4, _
             wdOutlineLevel5, _
             wdOutlineLevel6, _
             wdOutlineLevel7, _
             wdOutlineLevel8, _
             wdOutlineLevel9
            IsHeadingParagraph = True
    End Select
End Function

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
        .Alignment = wdAlignParagraphLeft
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
    ElseIf targetStyle.Type <> wdStyleTypeParagraph Then
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
