Attribute VB_Name = "AuWordReportExport"
Option Explicit

' Native Word VBA exporter for Au report blocks.
'
' Import this module into Word VBA, then run CreateAuWordReportFromWorkbook.
' It keeps Word export separate from the Python Excel workflow and uses Word's
' own PasteSpecial command for Microsoft Excel Worksheet Object output.

Private Const ORGANIZED_BLOCKS_SHEET_NAME As String = "Organized Blocks"
Private Const WORD_SAMPLE_FONT_SIZE As Long = 14
Private Const FIRST_PAGE_HEADER_RESERVED_POINTS As Double = 0#
Private Const APP_WORKBOOK_PATH As String = "__AU_WORKBOOK_PATH__"
Private Const APP_SAMPLE_TYPE As String = "__AU_SAMPLE_TYPE__"

Public Sub CreateAuWordReportFromWorkbook()
    CreateAuWordReportFromWorkbookCore "", ""
End Sub

Private Sub CreateAuWordReportFromWorkbookCore(ByVal workbookPathArgument As String, ByVal sampleTypeArgument As String)
    Dim excelApp As Object
    Dim workbook As Object
    Dim blockSheet As Object
    Dim workbookPath As String
    Dim outputPath As String
    Dim sampleName As String
    Dim orientationChoice As String
    Dim pageOrientation As WdOrientation
    Dim blockRanges As Collection
    Dim blockInfo As Variant
    Dim blockIndex As Long
    Dim blockRange As Object
    Dim exportSucceeded As Boolean

    On Error GoTo HandleError

    workbookPath = Trim$(workbookPathArgument)
    If Len(workbookPath) = 0 Then workbookPath = ConfiguredMacroValue(APP_WORKBOOK_PATH, "__AU_WORKBOOK_PATH__")
    If Len(workbookPath) = 0 Then workbookPath = PickExcelWorkbook()
    If Len(workbookPath) = 0 Then Exit Sub

    orientationChoice = Trim$(sampleTypeArgument)
    If Len(orientationChoice) = 0 Then orientationChoice = ConfiguredMacroValue(APP_SAMPLE_TYPE, "__AU_SAMPLE_TYPE__")
    If Len(orientationChoice) = 0 Then
        orientationChoice = InputBox( _
            "Select sample type for Word orientation:" & vbCrLf & vbCrLf & _
            "1 = Au+Ag (Portrait)" & vbCrLf & _
            "2 = Au+Ag+Cu (Portrait)" & vbCrLf & _
            "3 = Au+Ag+Cu+Hg (Landscape)" & vbCrLf & _
            "4 = Au+Ag+Hg (Portrait)", _
            "Au Word Report Sample Type", _
            "1")
    End If
    pageOrientation = OrientationFromSampleType(orientationChoice)

    sampleName = SampleNameFromPath(workbookPath)
    outputPath = ReplaceExtension(workbookPath, ".docx")

    Set excelApp = CreateObject("Excel.Application")
    excelApp.Visible = False
    Set workbook = excelApp.Workbooks.Open(workbookPath, 0, True)
    Set blockSheet = workbook.Worksheets(ORGANIZED_BLOCKS_SHEET_NAME)
    Set blockRanges = DetectBlockRanges(blockSheet)

    Documents.Add
    ActiveDocument.PageSetup.Orientation = pageOrientation
    ConfigureFirstPageSampleHeader ActiveDocument, sampleName

    For blockIndex = 1 To blockRanges.Count
        blockInfo = blockRanges(blockIndex)
        Selection.EndKey Unit:=wdStory
        If blockIndex > 1 Then
            Selection.TypeParagraph
            Selection.EndKey Unit:=wdStory
        End If
        Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter

        Set blockRange = blockSheet.Range(blockSheet.Cells(blockInfo(0), blockInfo(1)), blockSheet.Cells(blockInfo(2), blockInfo(3)))
        PasteBlockAsExcelWorksheetObject blockRange, ActiveDocument, (blockIndex = 1)
        excelApp.CutCopyMode = False
    Next blockIndex

    excelApp.CutCopyMode = False
    ActiveDocument.SaveAs2 FileName:=outputPath, FileFormat:=wdFormatXMLDocument
    exportSucceeded = True

CleanUp:
    On Error Resume Next
    If Not workbook Is Nothing Then workbook.Close SaveChanges:=False
    If Not excelApp Is Nothing Then excelApp.Quit
    On Error GoTo 0

    ' The Python app provides the Open Word Report button, so avoid an extra success dialog here.
    Exit Sub

HandleError:
    MsgBox "Word report export failed:" & vbCrLf & Err.Description, vbCritical, "Au Word Report"
    Resume CleanUp
End Sub

Private Function ConfiguredMacroValue(ByVal configuredValue As String, ByVal placeholderValue As String) As String
    If Trim$(configuredValue) = placeholderValue Then
        ConfiguredMacroValue = ""
    Else
        ConfiguredMacroValue = Trim$(configuredValue)
    End If
End Function

Private Function PickExcelWorkbook() As String
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = "Select the finished Au report workbook with Organized Blocks"
        .Filters.Clear
        .Filters.Add "Excel workbooks", "*.xlsx;*.xlsm"
        .AllowMultiSelect = False
        If .Show <> -1 Then
            PickExcelWorkbook = ""
        Else
            PickExcelWorkbook = .SelectedItems(1)
        End If
    End With
End Function

Private Function OrientationFromChoice(ByVal choice As String) As WdOrientation
    If Trim$(choice) = "3" Then
        OrientationFromChoice = wdOrientLandscape
    Else
        OrientationFromChoice = wdOrientPortrait
    End If
End Function

Private Function OrientationFromSampleType(ByVal sampleTypeOrChoice As String) As WdOrientation
    Dim normalizedValue As String
    normalizedValue = Replace$(UCase$(Trim$(sampleTypeOrChoice)), " ", "")

    If normalizedValue = "3" Or normalizedValue = "AU+AG+CU+HG" Or normalizedValue = "AU_AG_CU_HG" Then
        OrientationFromSampleType = wdOrientLandscape
    Else
        OrientationFromSampleType = wdOrientPortrait
    End If
End Function

Private Function DetectBlockRanges(ByVal blockSheet As Object) As Collection
    Dim ranges As New Collection
    Dim lastRow As Long
    Dim lastColumn As Long
    Dim rowNumber As Long
    Dim topRow As Long
    Dim bottomRow As Long
    Dim blockInfo(0 To 3) As Long

    lastRow = blockSheet.UsedRange.Row + blockSheet.UsedRange.Rows.Count - 1
    lastColumn = blockSheet.UsedRange.Column + blockSheet.UsedRange.Columns.Count - 1
    rowNumber = 1

    Do While rowNumber <= lastRow
        Do While rowNumber <= lastRow And Not RowHasContent(blockSheet, rowNumber, lastColumn)
            rowNumber = rowNumber + 1
        Loop
        If rowNumber > lastRow Then Exit Do

        topRow = rowNumber
        Do While rowNumber <= lastRow And RowHasContent(blockSheet, rowNumber, lastColumn)
            rowNumber = rowNumber + 1
        Loop
        bottomRow = rowNumber - 1

        blockInfo(0) = topRow
        blockInfo(1) = 1
        blockInfo(2) = bottomRow
        blockInfo(3) = lastColumn
        ranges.Add Array(blockInfo(0), blockInfo(1), blockInfo(2), blockInfo(3))
    Loop

    If ranges.Count = 0 Then Err.Raise vbObjectError + 1000, , "No report blocks were detected on '" & ORGANIZED_BLOCKS_SHEET_NAME & "'."
    Set DetectBlockRanges = ranges
End Function

Private Function RowHasContent(ByVal blockSheet As Object, ByVal rowNumber As Long, ByVal lastColumn As Long) As Boolean
    Dim columnNumber As Long
    For columnNumber = 1 To lastColumn
        If Len(CStr(blockSheet.Cells(rowNumber, columnNumber).Value)) > 0 Then
            RowHasContent = True
            Exit Function
        End If
    Next columnNumber
    RowHasContent = False
End Function

Private Sub ConfigureFirstPageSampleHeader(ByVal document As Document, ByVal sampleName As String)
    Dim headerRange As Range

    document.PageSetup.DifferentFirstPageHeaderFooter = True
    document.PageSetup.HeaderDistance = InchesToPoints(0.1)

    Set headerRange = document.Sections(1).Headers(wdHeaderFooterFirstPage).Range
    headerRange.Text = "Sample Name: " & sampleName
    headerRange.Font.Size = WORD_SAMPLE_FONT_SIZE
    headerRange.Font.Bold = True
    headerRange.ParagraphFormat.Alignment = wdAlignParagraphCenter
End Sub

Private Sub PasteBlockAsExcelWorksheetObject(ByVal blockRange As Object, ByVal document As Document, ByVal reserveHeaderSpace As Boolean)
    Dim beforeInlineCount As Long
    Dim pastedShape As InlineShape

    beforeInlineCount = document.InlineShapes.Count
    blockRange.Copy
    DoEvents
    Selection.PasteSpecial Link:=False, DataType:=wdPasteOLEObject, Placement:=wdInLine, DisplayAsIcon:=False

    If document.InlineShapes.Count <= beforeInlineCount Then
        Err.Raise vbObjectError + 1001, , "Word did not create an inline Excel worksheet object from Paste Special."
    End If

    Set pastedShape = document.InlineShapes(document.InlineShapes.Count)
    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    FitInlineShapeToPage pastedShape, document, reserveHeaderSpace
    Selection.EndKey Unit:=wdStory
End Sub

Private Sub FitInlineShapeToPage(ByVal pastedShape As InlineShape, ByVal document As Document, ByVal reserveHeaderSpace As Boolean)
    Dim availableWidth As Double
    Dim availableHeight As Double
    Dim widthScale As Double
    Dim heightScale As Double
    Dim scaleFactor As Double

    availableWidth = document.PageSetup.PageWidth - document.PageSetup.LeftMargin - document.PageSetup.RightMargin
    availableHeight = document.PageSetup.PageHeight - document.PageSetup.TopMargin - document.PageSetup.BottomMargin
    If reserveHeaderSpace Then availableHeight = availableHeight - FIRST_PAGE_HEADER_RESERVED_POINTS

    If pastedShape.Width <= 0 Or pastedShape.Height <= 0 Then Exit Sub

    widthScale = availableWidth / pastedShape.Width
    heightScale = availableHeight / pastedShape.Height
    scaleFactor = widthScale
    If heightScale < scaleFactor Then scaleFactor = heightScale

    If scaleFactor < 1 Then
        pastedShape.LockAspectRatio = msoTrue
        pastedShape.Width = pastedShape.Width * scaleFactor
    End If
End Sub

Private Function SampleNameFromPath(ByVal workbookPath As String) As String
    Dim fileName As String
    fileName = Mid$(workbookPath, InStrRev(workbookPath, "\") + 1)
    fileName = Left$(fileName, InStrRev(fileName, ".") - 1)
    fileName = Replace(fileName, "_Final", "")
    fileName = Replace(fileName, " Final", "")
    fileName = Replace(fileName, "_Inter", "")
    fileName = Replace(fileName, " Inter", "")
    fileName = Replace(fileName, "(Au SEM)", "")
    fileName = Replace(fileName, "(au sem)", "")
    SampleNameFromPath = Trim$(fileName)
End Function

Private Function ReplaceExtension(ByVal filePath As String, ByVal newExtension As String) As String
    ReplaceExtension = Left$(filePath, InStrRev(filePath, ".") - 1) & newExtension
End Function
