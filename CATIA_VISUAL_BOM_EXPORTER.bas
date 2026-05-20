' ================================================================
' CATIA_VISUAL_BOM_EXPORTER
' Native CATIA BOM -> Excel -> standalone-window thumbnails only.
' ================================================================

' ---------------------- CONFIGURATION ---------------------------
Const STANDALONE_CAPTURE_ONLY = True
Const PREFER_STANDALONE_PART_CAPTURE = True
Const FALLBACK_TO_ASSEMBLY_HIDE_SHOW = False
Const CLOSE_STANDALONE_DOCUMENT_AFTER_CAPTURE = True
Const NEVER_SAVE_CATIA_DOCUMENTS = True
Const SKIP_FASTENER_IMAGES = True
Const SKIP_FASTENER_ROWS = False
Const FASTENER_KEYWORDS = "vijak|vijci|zavrtanj|zavrtnji|screw|bolt|hex bolt|hexagon bolt|imbus|allen screw|navrtka|matica|nut|hex nut|podloska|washer|plain washer|spring washer|lock washer|DIN 125|DIN125|DIN 127|DIN127|DIN 933|DIN933|DIN 931|DIN931|DIN 934|DIN934|ISO 4014|ISO 4017|ISO 4032|ISO 7089|ISO 7090"
Const FAST_IMAGE_FORMAT = "jpg"
Const IMAGE_WIDTH = 1000
Const IMAGE_HEIGHT = 750
Const CREATE_THUMBNAIL_FILES = True
Const INSERT_THUMBNAIL_FILE_IN_EXCEL = True
Const INSERT_IMAGES_IN_EXCEL = True
Const THUMBNAIL_WIDTH = 160
Const THUMBNAIL_HEIGHT = 120
Const TEST_MODE = True
Const TEST_MAX_ROWS = 20
Const SAVE_EVERY_N_ROWS = 25
Const RESUME_MODE = True
Const SKIP_EXISTING_IMAGES = True
Const OUTPUT_TO_DESKTOP = True
Const USE_SHADED_WITH_EDGES = True
Const USE_WHITE_BACKGROUND = True
Const USE_PARALLEL_PROJECTION = True
Const IMAGE_CAPTURE_DELAY_SECONDS = 0.15

' CATIA constants used late-bound.
Const CAT_CAPTURE_FORMAT_JPEG = 2
Const CAT_RENDER_SHADING = 0
Const CAT_RENDER_SHADING_WITH_EDGES = 1
Const CAT_PROJECTION_CYLINDRIC = 0
Const CAT_FILE_TYPE_TXT = 0
Const CAT_FILE_TYPE_HTML = 2

' Excel constants used late-bound.
Const XL_OPENXML_WORKBOOK = 51
Const XL_CENTER = -4108
Const XL_LEFT = -4131
Const XL_TOP = -4160
Const XL_CONTINUOUS = 1
Const XL_THIN = 2
Const MSO_TRUE = -1
Const MSO_FALSE = 0

Dim gFSO
Dim gShell
Dim gRootProduct
Dim gActiveDoc
Dim gOutputFolder
Dim gImageFolder
Dim gThumbnailFolder
Dim gDebugLogPath
Dim gExcelPath
Dim gNativeBomPath
Dim gNativeHeaders
Dim gBomRows
Dim gSourceIndex
Dim gImageCache
Dim gExcelApp
Dim gWorkbook
Dim gWsBom
Dim gWsLog
Dim gWsSummary
Dim gWorkbookSaved
Dim gPartNumberColumnIndex
Dim gDescriptionColumnIndex
Dim gThumbnailColumnIndex
Dim gImagePathColumnIndex
Dim gThumbnailPathColumnIndex
Dim gExportStatusColumnIndex
Dim gImageSkipReasonColumnIndex
Dim gNextLogRow
Dim gProcessedImageRows
Dim gSuccessfulImageRows
Dim gSkippedFastenerRows
Dim gCurrentStandaloneDoc
Dim gCurrentStandaloneOpened
Dim gMainDocumentFullName
Dim gAbortMessage

Public Sub CATIA_VISUAL_BOM_EXPORTER()
    On Error Resume Next
    InitializeRuntime
    If Not ValidateActiveProductDocument() Then Exit Sub
    PrepareOutputFolders
    WriteDebugPhase "START", 0, "", "Native CATIA BOM export started."

    If Not ExportAndLoadNativeBom() Then
        AbortWithMessage "CATIA native Bill of Material export nije uspeo. Makro ne koristi rucnu rekonstrukciju BOM-a.", "CATIA_NATIVE_BOM_EXPORT", 0, ""
        Exit Sub
    End If

    If Not CreateExcelReportFromNativeBom() Then
        AbortWithMessage "Microsoft Excel nije dostupan ili Excel fajl nije mogao da se napravi.", "EXCEL_CREATED", 0, ""
        Exit Sub
    End If

    WriteDebugPhase "CATIA_FILE_INDEX_START", 0, "", "Building Part Number -> source file path index."
    BuildCatiaFileIndex
    WriteDebugPhase "CATIA_FILE_INDEX_DONE", 0, "", "Indexed source files: " & CStr(gSourceIndex.Count)

    If Not ProcessBomRowsForThumbnails() Then Exit Sub

    SaveExcelCheckpoint "FINISH", 0, ""
    UpdateSummarySheet
    CleanupCatiaSession
    WriteDebugPhase "FINISH", 0, "", "Macro finished."
    MsgBox "Visual BOM export je zavrsen." & vbCrLf & vbCrLf & _
           "Excel:" & vbCrLf & gExcelPath, vbInformation, "CATIA_VISUAL_BOM_EXPORTER"
End Sub

Private Sub InitializeRuntime()
    Set gFSO = CreateObject("Scripting.FileSystemObject")
    Set gShell = CreateObject("WScript.Shell")
    Set gNativeHeaders = NewList()
    Set gBomRows = NewList()
    Set gSourceIndex = CreateObject("Scripting.Dictionary")
    Set gImageCache = CreateObject("Scripting.Dictionary")
    gSourceIndex.CompareMode = 1
    gImageCache.CompareMode = 1
    Set gExcelApp = Nothing
    Set gWorkbook = Nothing
    Set gWsBom = Nothing
    Set gWsLog = Nothing
    Set gWsSummary = Nothing
    gWorkbookSaved = False
    Set gCurrentStandaloneDoc = Nothing
    gCurrentStandaloneOpened = False
    gPartNumberColumnIndex = 0
    gDescriptionColumnIndex = 0
    gThumbnailColumnIndex = 0
    gImagePathColumnIndex = 0
    gThumbnailPathColumnIndex = 0
    gExportStatusColumnIndex = 0
    gImageSkipReasonColumnIndex = 0
    gNextLogRow = 2
    gProcessedImageRows = 0
    gSuccessfulImageRows = 0
    gSkippedFastenerRows = 0
    gAbortMessage = ""
    Randomize
End Sub

Private Function ValidateActiveProductDocument()
    On Error Resume Next
    ValidateActiveProductDocument = False
    If CATIA.Documents.Count = 0 Then
        MsgBox "Nema otvorenog CATIA dokumenta. Otvorite CATProduct i pokrenite makro.", vbExclamation, "CATIA_VISUAL_BOM_EXPORTER"
        Exit Function
    End If
    Set gActiveDoc = CATIA.ActiveDocument
    If Not IsActiveDocumentProduct(gActiveDoc) Then
        MsgBox "Aktivni dokument mora biti CATProduct.", vbExclamation, "CATIA_VISUAL_BOM_EXPORTER"
        Exit Function
    End If
    Set gRootProduct = gActiveDoc.Product
    gMainDocumentFullName = GetDocumentFullName(gActiveDoc)
    ValidateActiveProductDocument = True
    Err.Clear
End Function

Private Function IsActiveDocumentProduct(doc)
    On Error Resume Next
    Dim p
    Set p = doc.Product
    IsActiveDocumentProduct = (Err.Number = 0 And Not p Is Nothing)
    Err.Clear
End Function

Private Sub PrepareOutputFolders()
    Dim rootName
    Dim baseFolder
    rootName = SafeFileName(FirstNonEmpty(GetProductPartNumber(gRootProduct), gRootProduct.Name, "CATProduct"))
    If rootName = "" Then rootName = "CATProduct"

    If OUTPUT_TO_DESKTOP Then
        baseFolder = gShell.SpecialFolders("Desktop")
    Else
        baseFolder = CStr(gActiveDoc.Path)
        If baseFolder = "" Then baseFolder = gShell.SpecialFolders("Desktop")
    End If

    gOutputFolder = JoinPath(baseFolder, rootName & "_VISUAL_BOM_EXPORT")
    gImageFolder = JoinPath(gOutputFolder, "IMAGES")
    gThumbnailFolder = JoinPath(gOutputFolder, "THUMBNAILS")
    gDebugLogPath = JoinPath(gOutputFolder, "DEBUG_PHASE_LOG.txt")
    gExcelPath = JoinPath(gOutputFolder, "VISUAL_BOM_EXPORT.xlsx")

    EnsureFolder gOutputFolder
    EnsureFolder gImageFolder
    EnsureFolder gThumbnailFolder
End Sub

Private Function ExportAndLoadNativeBom()
    On Error Resume Next
    ExportAndLoadNativeBom = False
    WriteDebugPhase "CATIA_NATIVE_BOM_EXPORT_START", 0, GetProductPartNumber(gRootProduct), "Trying Product.ExtractBOM."

    Dim fileTypes
    Dim fileExts
    Dim i
    Dim candidatePath
    fileTypes = Array(CAT_FILE_TYPE_TXT, 1, CAT_FILE_TYPE_HTML)
    fileExts = Array("txt", "txt", "html")

    For i = LBound(fileTypes) To UBound(fileTypes)
        ResetNativeBomData
        candidatePath = JoinPath(gOutputFolder, "~CATIA_NATIVE_BOM." & CStr(fileExts(i)))
        If gFSO.FileExists(candidatePath) Then gFSO.DeleteFile candidatePath, True
        Err.Clear
        gRootProduct.ExtractBOM CLng(fileTypes(i)), candidatePath
        If Err.Number = 0 And gFSO.FileExists(candidatePath) Then
            If LoadNativeBomFile(candidatePath) Then
                gNativeBomPath = candidatePath
                WriteDebugPhase "CATIA_NATIVE_BOM_EXPORT_DONE", 0, "", "Native BOM rows: " & CStr(ListCount(gBomRows)) & "; file=" & candidatePath
                WriteDebugPhase "BOM_HEADERS_READ", 0, "", JoinHeaderNames()
                gFSO.DeleteFile candidatePath, True
                ExportAndLoadNativeBom = True
                Exit Function
            End If
        End If
        Err.Clear
    Next
End Function

Private Sub ResetNativeBomData()
    Set gNativeHeaders = NewList()
    Set gBomRows = NewList()
End Sub

Private Function LoadNativeBomFile(filePath)
    On Error Resume Next
    LoadNativeBomFile = False
    Dim ts
    Dim txt
    If Not gFSO.FileExists(filePath) Then Exit Function
    Set ts = gFSO.OpenTextFile(filePath, 1, False)
    If Err.Number <> 0 Then
        Err.Clear
        Exit Function
    End If
    txt = ts.ReadAll
    ts.Close

    If InStr(1, LCase(filePath), ".htm", vbTextCompare) > 0 Then
        LoadNativeBomFile = LoadHtmlBomText(txt)
    Else
        LoadNativeBomFile = LoadDelimitedBomText(txt)
    End If
    LoadNativeBomFile = (LoadNativeBomFile And ListCount(gNativeHeaders) > 0 And ListCount(gBomRows) > 0)
    Err.Clear
End Function

Private Function LoadDelimitedBomText(txt)
    On Error Resume Next
    LoadDelimitedBomText = False
    Dim lines
    Dim i
    Dim lineText
    Dim delim
    Dim cells
    Dim headerFound
    lines = Split(Replace(txt, vbCr, vbLf), vbLf)
    headerFound = False

    For i = LBound(lines) To UBound(lines)
        lineText = Trim(CStr(lines(i)))
        If lineText <> "" Then
            delim = DetectDelimiter(lineText)
            If delim <> "" Then
                cells = ParseDelimitedLine(lineText, delim)
                If Not headerFound Then
                    If LooksLikeBomHeader(cells) Then
                        AddHeaderArray cells
                        headerFound = True
                    End If
                Else
                    AddBomRowArray cells
                End If
            End If
        End If
    Next

    If Not headerFound Then
        For i = LBound(lines) To UBound(lines)
            lineText = Trim(CStr(lines(i)))
            delim = DetectDelimiter(lineText)
            If delim <> "" Then
                cells = ParseDelimitedLine(lineText, delim)
                If CountUsefulCells(cells) >= 2 Then
                    AddHeaderArray cells
                    headerFound = True
                    Exit For
                End If
            End If
        Next
        If headerFound Then
            For i = i + 1 To UBound(lines)
                lineText = Trim(CStr(lines(i)))
                delim = DetectDelimiter(lineText)
                If delim <> "" Then AddBomRowArray ParseDelimitedLine(lineText, delim)
            Next
        End If
    End If

    LoadDelimitedBomText = (ListCount(gNativeHeaders) > 0 And ListCount(gBomRows) > 0)
    Err.Clear
End Function

Private Function LoadHtmlBomText(txt)
    On Error Resume Next
    LoadHtmlBomText = False
    Dim reTr
    Dim reCell
    Dim rows
    Dim rowMatch
    Dim cellMatches
    Dim cellMatch
    Dim values
    Dim headerFound

    Set reTr = CreateObject("VBScript.RegExp")
    reTr.Global = True
    reTr.IgnoreCase = True
    reTr.Pattern = "<tr[\s\S]*?</tr>"

    Set reCell = CreateObject("VBScript.RegExp")
    reCell.Global = True
    reCell.IgnoreCase = True
    reCell.Pattern = "<t[dh][^>]*>([\s\S]*?)</t[dh]>"

    Set rows = reTr.Execute(txt)
    headerFound = False
    For Each rowMatch In rows
        Set cellMatches = reCell.Execute(CStr(rowMatch.Value))
        If cellMatches.Count > 0 Then
            values = EmptyArray()
            For Each cellMatch In cellMatches
                values = ArrayPush(values, HtmlToText(CStr(cellMatch.SubMatches(0))))
            Next
            If Not headerFound Then
                If LooksLikeBomHeader(values) Or CountUsefulCells(values) >= 2 Then
                    AddHeaderArray values
                    headerFound = True
                End If
            Else
                AddBomRowArray values
            End If
        End If
    Next
    LoadHtmlBomText = (ListCount(gNativeHeaders) > 0 And ListCount(gBomRows) > 0)
    Err.Clear
End Function

Private Sub AddHeaderArray(cells)
    Dim i
    Dim h
    Set gNativeHeaders = NewList()
    For i = LBound(cells) To UBound(cells)
        h = Trim(CStr(cells(i)))
        If h = "" Then h = "Column " & CStr(i + 1)
        ListAddValue gNativeHeaders, h
    Next
    gPartNumberColumnIndex = FindHeaderIndexInList(gNativeHeaders, "Part Number|Number|PartNumber|Broj dela|Br. dela")
    gDescriptionColumnIndex = FindHeaderIndexInList(gNativeHeaders, "Nomenclature|Description|Keywords|Naziv")
End Sub

Private Sub AddBomRowArray(cells)
    On Error Resume Next
    If CountUsefulCells(cells) = 0 Then Exit Sub
    If IsSeparatorRow(cells) Then Exit Sub

    Dim values
    Dim i
    Set values = NewList()
    For i = 1 To ListCount(gNativeHeaders)
        If (i - 1) <= UBound(cells) Then
            ListAddValue values, Trim(CStr(cells(i - 1)))
        Else
            ListAddValue values, ""
        End If
    Next

    Dim rec
    Set rec = CreateObject("Scripting.Dictionary")
    Set rec.Item("Values") = values
    rec.Item("ExcelRow") = CLng(ListCount(gBomRows) + 2)
    rec.Item("ImagePath") = ""
    rec.Item("ThumbnailPath") = ""
    rec.Item("ExportStatus") = ""
    rec.Item("ImageSkipReason") = ""
    ListAddObject gBomRows, rec
    Err.Clear
End Sub

Private Function CreateExcelReportFromNativeBom()
    On Error Resume Next
    CreateExcelReportFromNativeBom = False
    Set gExcelApp = CreateObject("Excel.Application")
    If Err.Number <> 0 Or gExcelApp Is Nothing Then
        Err.Clear
        Exit Function
    End If
    gExcelApp.Visible = False
    gExcelApp.DisplayAlerts = False
    gExcelApp.ScreenUpdating = False

    Set gWorkbook = gExcelApp.Workbooks.Add
    Do While gWorkbook.Worksheets.Count < 3
        gWorkbook.Worksheets.Add
    Loop
    Do While gWorkbook.Worksheets.Count > 3
        gWorkbook.Worksheets(gWorkbook.Worksheets.Count).Delete
    Loop
    Set gWsBom = gWorkbook.Worksheets(1)
    Set gWsLog = gWorkbook.Worksheets(2)
    Set gWsSummary = gWorkbook.Worksheets(3)
    gWsBom.Name = "BOM"
    gWsLog.Name = "EXPORT_LOG"
    gWsSummary.Name = "SUMMARY"

    PrepareBomSheet
    PrepareLogSheet
    PrepareSummarySheet
    WriteNativeBomRowsToExcel
    SaveExcelCheckpoint "EXCEL_CREATED", 0, ""
    WriteDebugPhase "EXCEL_CREATED", 0, "", gExcelPath
    CreateExcelReportFromNativeBom = True
    Err.Clear
End Function

Private Sub PrepareBomSheet()
    Dim insertAfter
    Dim c
    Dim outCol
    Dim headerText

    insertAfter = gPartNumberColumnIndex
    If insertAfter <= 0 Then insertAfter = 1
    outCol = 1
    For c = 1 To ListCount(gNativeHeaders)
        headerText = CStr(ListValue(gNativeHeaders, c))
        gWsBom.Cells(1, outCol).Value = headerText
        outCol = outCol + 1
        If c = insertAfter Then
            gThumbnailColumnIndex = outCol
            gWsBom.Cells(1, outCol).Value = "Thumbnail"
            outCol = outCol + 1
        End If
    Next
    gImagePathColumnIndex = outCol
    gWsBom.Cells(1, outCol).Value = "Image Path"
    outCol = outCol + 1
    gThumbnailPathColumnIndex = outCol
    gWsBom.Cells(1, outCol).Value = "Thumbnail Path"
    outCol = outCol + 1
    gExportStatusColumnIndex = outCol
    gWsBom.Cells(1, outCol).Value = "Export Status"
    outCol = outCol + 1
    gImageSkipReasonColumnIndex = outCol
    gWsBom.Cells(1, outCol).Value = "Image Skip Reason"

    gWsBom.Rows(1).Font.Bold = True
    gWsBom.Rows(1).HorizontalAlignment = XL_CENTER
    gWsBom.Columns(gThumbnailColumnIndex).ColumnWidth = 24
    gWsBom.Columns(gImagePathColumnIndex).ColumnWidth = 45
    gWsBom.Columns(gThumbnailPathColumnIndex).ColumnWidth = 45
    gWsBom.Columns(gExportStatusColumnIndex).ColumnWidth = 22
    gWsBom.Columns(gImageSkipReasonColumnIndex).ColumnWidth = 40
End Sub

Private Sub PrepareLogSheet()
    gWsLog.Cells(1, 1).Value = "No."
    gWsLog.Cells(1, 2).Value = "Date/Time"
    gWsLog.Cells(1, 3).Value = "Excel Row"
    gWsLog.Cells(1, 4).Value = "Part Number"
    gWsLog.Cells(1, 5).Value = "Status"
    gWsLog.Cells(1, 6).Value = "Phase"
    gWsLog.Cells(1, 7).Value = "Message"
    gWsLog.Cells(1, 8).Value = "Image Path"
    gWsLog.Cells(1, 9).Value = "Thumbnail Path"
    gWsLog.Rows(1).Font.Bold = True
    gWsLog.Columns("A:I").ColumnWidth = 24
End Sub

Private Sub PrepareSummarySheet()
    gWsSummary.Cells(1, 1).Value = "CATIA VISUAL BOM EXPORTER"
    gWsSummary.Cells(3, 1).Value = "Main Assembly Part Number"
    gWsSummary.Cells(4, 1).Value = "Export Date/Time"
    gWsSummary.Cells(5, 1).Value = "Output Folder"
    gWsSummary.Cells(6, 1).Value = "Excel File"
    gWsSummary.Cells(7, 1).Value = "Native BOM Rows"
    gWsSummary.Cells(8, 1).Value = "Images Processed"
    gWsSummary.Cells(9, 1).Value = "Successful Images"
    gWsSummary.Cells(10, 1).Value = "Skipped Fasteners"
    gWsSummary.Cells(11, 1).Value = "Debug Log"
    gWsSummary.Cells(12, 1).Value = "Mode"
    gWsSummary.Range("A1:B1").Font.Bold = True
    gWsSummary.Columns("A").ColumnWidth = 28
    gWsSummary.Columns("B").ColumnWidth = 95
    UpdateSummarySheet
End Sub

Private Sub WriteNativeBomRowsToExcel()
    Dim r
    Dim c
    Dim outCol
    Dim rec
    For r = 1 To ListCount(gBomRows)
        Set rec = ListObject(gBomRows, r)
        outCol = 1
        For c = 1 To ListCount(gNativeHeaders)
            gWsBom.Cells(CLng(rec.Item("ExcelRow")), outCol).Value = GetNativeValue(rec, c)
            outCol = outCol + 1
            If outCol = gThumbnailColumnIndex Then outCol = outCol + 1
        Next
        gWsBom.Rows(CLng(rec.Item("ExcelRow"))).RowHeight = CLng(THUMBNAIL_HEIGHT * 0.75) + 12
    Next
    ApplyBomBorders
End Sub

Private Function ProcessBomRowsForThumbnails()
    On Error Resume Next
    ProcessBomRowsForThumbnails = False
    Dim i
    Dim rec
    Dim excelRow
    Dim partNumber
    Dim status
    Dim reason
    Dim sourcePath
    Dim imagePath
    Dim thumbPath
    Dim shouldMakeImage

    For i = 1 To ListCount(gBomRows)
        Set rec = ListObject(gBomRows, i)
        excelRow = CLng(rec.Item("ExcelRow"))
        partNumber = GetPartNumberFromRow(rec)
        CATIA.StatusBar = "BOM thumbnail row " & CStr(i) & " / " & CStr(ListCount(gBomRows)) & " - " & partNumber
        WriteDebugPhase "ROW_START", excelRow, partNumber, ""
        shouldMakeImage = True

        If SKIP_FASTENER_IMAGES And IsFastenerBomRow(rec) Then
            gSkippedFastenerRows = gSkippedFastenerRows + 1
            reason = "Standard fastener - retained in BOM, image skipped"
            SetBomUtilityValues rec, "", "", "SKIPPED_IMAGE_ONLY", reason
            WriteExportLog excelRow, partNumber, "SKIPPED_IMAGE_ONLY", "FASTENER_SKIPPED_IMAGE", reason, "", ""
            WriteDebugPhase "FASTENER_SKIPPED_IMAGE", excelRow, partNumber, reason
            shouldMakeImage = False
        End If

        If shouldMakeImage And TEST_MODE And gProcessedImageRows >= TEST_MAX_ROWS Then
            SetBomUtilityValues rec, "", "", "", "TEST_MODE limit reached; image not processed"
            shouldMakeImage = False
        End If

        If shouldMakeImage Then
            gProcessedImageRows = gProcessedImageRows + 1
            imagePath = BuildImagePath(partNumber, GetDescriptionFromRow(rec), excelRow)
            thumbPath = ThumbnailPathForImage(imagePath)

            If ReuseExistingOrCachedImage(partNumber, imagePath, thumbPath, rec) Then
                If Not InsertThumbnailForRow(excelRow, thumbPath) Then
                    AbortWithMessage "Excel ne moze da ubaci thumbnail za Part Number: " & partNumber, "EXCEL_THUMBNAIL_INSERTED", excelRow, partNumber
                    Exit Function
                End If
                SetBomUtilityValues rec, imagePath, thumbPath, "EXISTING_REUSED", ""
                WriteExportLog excelRow, partNumber, "EXISTING_REUSED", "EXCEL_THUMBNAIL_INSERTED", "Existing image/thumbnail reused.", imagePath, thumbPath
                gSuccessfulImageRows = gSuccessfulImageRows + 1
                MaybeSaveByProgress excelRow, partNumber
            Else
                sourcePath = SourcePathForPartNumber(partNumber)
                If sourcePath = "" Or Not gFSO.FileExists(sourcePath) Then
                    SetBomUtilityValues rec, "", "", "SOURCE_FILE_NOT_FOUND", "Source file not found"
                    WriteExportLog excelRow, partNumber, "SOURCE_FILE_NOT_FOUND", "SOURCE_FILE_NOT_FOUND", "Source file not found for Part Number: " & partNumber, "", ""
                    WriteDebugPhase "SOURCE_FILE_NOT_FOUND", excelRow, partNumber, "Source file not found for Part Number: " & partNumber
                    AbortWithMessage "Source file not found for Part Number: " & partNumber, "SOURCE_FILE_NOT_FOUND", excelRow, partNumber
                    Exit Function
                End If

                If Not IsSupportedCatiaSourceFile(sourcePath) Then
                    SetBomUtilityValues rec, "", "", "UNSUPPORTED_SOURCE_FILE", sourcePath
                    WriteExportLog excelRow, partNumber, "UNSUPPORTED_SOURCE_FILE", "UNSUPPORTED_SOURCE_FILE", sourcePath, "", ""
                    WriteDebugPhase "UNSUPPORTED_SOURCE_FILE", excelRow, partNumber, sourcePath
                    AbortWithMessage "Unsupported source file for Part Number: " & partNumber & vbCrLf & sourcePath, "UNSUPPORTED_SOURCE_FILE", excelRow, partNumber
                    Exit Function
                End If

                WriteDebugPhase "SOURCE_FILE_FOUND", excelRow, partNumber, sourcePath
                If Not CaptureStandaloneImage(partNumber, sourcePath, imagePath, thumbPath, excelRow) Then
                    SetBomUtilityValues rec, imagePath, thumbPath, "STANDALONE_CAPTURE_FAILED", "Standalone open/capture failed"
                    AbortWithMessage "Standalone capture failed for Part Number: " & partNumber, "STANDALONE_CAPTURE_FAILED", excelRow, partNumber
                    Exit Function
                End If

                SetBomUtilityValues rec, imagePath, thumbPath, "OK", ""
                CacheImage partNumber, imagePath, thumbPath
                If Not InsertThumbnailForRow(excelRow, thumbPath) Then
                    AbortWithMessage "Excel ne moze da ubaci thumbnail za Part Number: " & partNumber, "EXCEL_THUMBNAIL_INSERTED", excelRow, partNumber
                    Exit Function
                End If
                gSuccessfulImageRows = gSuccessfulImageRows + 1
                WriteExportLog excelRow, partNumber, "OK", "EXCEL_THUMBNAIL_INSERTED", "Thumbnail inserted.", imagePath, thumbPath
                MaybeSaveByProgress excelRow, partNumber
            End If
        End If

        DoEvents
    Next

    ProcessBomRowsForThumbnails = True
    Err.Clear
End Function

Private Sub BuildCatiaFileIndex()
    On Error Resume Next
    TraverseProductForSourceIndex gRootProduct
    Err.Clear
End Sub

Private Sub TraverseProductForSourceIndex(prod)
    On Error Resume Next
    Dim pn
    Dim sourcePath
    pn = GetProductPartNumber(prod)
    sourcePath = GetProductSourceFilePath(prod)
    If pn <> "" And sourcePath <> "" Then
        If Not SamePath(sourcePath, gMainDocumentFullName) Or NormalizePartNumber(pn) = NormalizePartNumber(GetProductPartNumber(gRootProduct)) Then
            If Not gSourceIndex.Exists(NormalizePartNumber(pn)) Then gSourceIndex.Item(NormalizePartNumber(pn)) = sourcePath
        End If
    End If

    Dim children
    Dim i
    Set children = prod.Products
    If Err.Number <> 0 Or children Is Nothing Then
        Err.Clear
        Exit Sub
    End If
    For i = 1 To children.Count
        TraverseProductForSourceIndex children.Item(i)
        If (gSourceIndex.Count Mod 250) = 0 Then DoEvents
    Next
    Err.Clear
End Sub

Private Function CaptureStandaloneImage(partNumber, sourcePath, imagePath, thumbPath, excelRow)
    On Error Resume Next
    CaptureStandaloneImage = False
    If gFSO.FileExists(imagePath) Then gFSO.DeleteFile imagePath, True
    If gFSO.FileExists(thumbPath) Then gFSO.DeleteFile thumbPath, True

    WriteDebugPhase "STANDALONE_OPEN_START", excelRow, partNumber, sourcePath
    If Not OpenStandaloneDocument(sourcePath) Then
        WriteExportLog excelRow, partNumber, "STANDALONE_OPEN_FAILED", "STANDALONE_OPEN_START", "Cannot open source file.", imagePath, thumbPath
        WriteDebugPhase "ERROR", excelRow, partNumber, "Cannot open source file: " & sourcePath
        CloseCurrentStandaloneDocument
        ActivateMainDocument
        Exit Function
    End If
    WriteDebugPhase "STANDALONE_OPEN_DONE", excelRow, partNumber, sourcePath

    WriteDebugPhase "STANDALONE_CAPTURE_START", excelRow, partNumber, imagePath
    ApplyIsoViewAndFit
    If Not CaptureActiveViewerToJpg(imagePath) Then
        WriteExportLog excelRow, partNumber, "STANDALONE_CAPTURE_FAILED", "STANDALONE_CAPTURE_START", "CaptureToFile failed.", imagePath, thumbPath
        WriteDebugPhase "ERROR", excelRow, partNumber, "CaptureToFile failed: " & imagePath
        CloseCurrentStandaloneDocument
        ActivateMainDocument
        Exit Function
    End If
    WriteDebugPhase "STANDALONE_CAPTURE_DONE", excelRow, partNumber, imagePath

    If Not CreateThumbnailFile(imagePath, thumbPath, THUMBNAIL_WIDTH, THUMBNAIL_HEIGHT) Then
        WriteExportLog excelRow, partNumber, "ERROR", "THUMBNAIL_CREATED", "Thumbnail creation failed.", imagePath, thumbPath
        WriteDebugPhase "ERROR", excelRow, partNumber, "Thumbnail creation failed: " & thumbPath
        CloseCurrentStandaloneDocument
        ActivateMainDocument
        Exit Function
    End If
    WriteDebugPhase "THUMBNAIL_CREATED", excelRow, partNumber, thumbPath

    CloseCurrentStandaloneDocument
    ActivateMainDocument
    CaptureStandaloneImage = True
    Err.Clear
End Function

Private Function OpenStandaloneDocument(sourcePath)
    On Error Resume Next
    OpenStandaloneDocument = False
    Set gCurrentStandaloneDoc = Nothing
    gCurrentStandaloneOpened = False

    If FindOpenDocumentByFullName(sourcePath, gCurrentStandaloneDoc) Then
        gCurrentStandaloneDoc.Activate
        OpenStandaloneDocument = True
        Exit Function
    End If

    Dim oldAlerts
    oldAlerts = CATIA.DisplayFileAlerts
    Err.Clear
    CATIA.DisplayFileAlerts = False
    Set gCurrentStandaloneDoc = CATIA.Documents.Open(sourcePath)
    CATIA.DisplayFileAlerts = oldAlerts
    If Err.Number <> 0 Or gCurrentStandaloneDoc Is Nothing Then
        CATIA.DisplayFileAlerts = oldAlerts
        Err.Clear
        Exit Function
    End If
    gCurrentStandaloneOpened = True
    gCurrentStandaloneDoc.Activate
    OpenStandaloneDocument = True
    Err.Clear
End Function

Private Sub CloseCurrentStandaloneDocument()
    On Error Resume Next
    If Not gCurrentStandaloneDoc Is Nothing Then
        If CLOSE_STANDALONE_DOCUMENT_AFTER_CAPTURE And gCurrentStandaloneOpened Then
            Dim oldAlerts
            oldAlerts = CATIA.DisplayFileAlerts
            Err.Clear
            CATIA.DisplayFileAlerts = False
            gCurrentStandaloneDoc.Close
            CATIA.DisplayFileAlerts = oldAlerts
            WriteDebugPhase "STANDALONE_CLOSE_DONE", 0, "", "Standalone document closed without saving."
        End If
    End If
    Set gCurrentStandaloneDoc = Nothing
    gCurrentStandaloneOpened = False
    Err.Clear
End Sub

Private Function CaptureActiveViewerToJpg(imagePath)
    On Error Resume Next
    CaptureActiveViewerToJpg = False
    Dim viewer
    Set viewer = CATIA.ActiveWindow.ActiveViewer
    viewer.Activate
    viewer.Update
    WaitSeconds IMAGE_CAPTURE_DELAY_SECONDS
    viewer.CaptureToFile CAT_CAPTURE_FORMAT_JPEG, imagePath
    WaitSeconds 0.05
    CaptureActiveViewerToJpg = gFSO.FileExists(imagePath)
    Err.Clear
End Function

Private Sub ApplyIsoViewAndFit()
    On Error Resume Next
    Dim viewer
    Dim vp
    Dim sight(2)
    Dim up(2)
    Set viewer = CATIA.ActiveWindow.ActiveViewer
    viewer.Activate
    If USE_SHADED_WITH_EDGES Then
        viewer.RenderingMode = CAT_RENDER_SHADING_WITH_EDGES
        If Err.Number <> 0 Then
            Err.Clear
            viewer.RenderingMode = CAT_RENDER_SHADING
        End If
    End If
    If USE_WHITE_BACKGROUND Then
        Dim bg(2)
        bg(0) = 1
        bg(1) = 1
        bg(2) = 1
        viewer.PutBackgroundColor bg
        Err.Clear
    End If
    Set vp = viewer.Viewpoint3D
    sight(0) = 1
    sight(1) = -1
    sight(2) = 1
    up(0) = -0.4082482905
    up(1) = 0.4082482905
    up(2) = 0.8164965809
    vp.PutSightDirection sight
    vp.PutUpDirection up
    If USE_PARALLEL_PROJECTION Then
        vp.ProjectionMode = CAT_PROJECTION_CYLINDRIC
        Err.Clear
    End If
    CATIA.ActiveWindow.Width = IMAGE_WIDTH
    CATIA.ActiveWindow.Height = IMAGE_HEIGHT
    viewer.Update
    viewer.Reframe
    viewer.Update
    viewer.Reframe
    viewer.Update
    DoEvents
    Err.Clear
End Sub

Private Function CreateThumbnailFile(sourcePath, thumbPath, maxWidth, maxHeight)
    On Error Resume Next
    CreateThumbnailFile = False
    Dim img
    Dim proc
    Dim thumb
    Dim tmpPath
    Set img = CreateObject("WIA.ImageFile")
    img.LoadFile sourcePath
    Set proc = CreateObject("WIA.ImageProcess")
    proc.Filters.Add proc.FilterInfos("Scale").FilterID
    proc.Filters(1).Properties("MaximumWidth").Value = CLng(maxWidth)
    proc.Filters(1).Properties("MaximumHeight").Value = CLng(maxHeight)
    proc.Filters(1).Properties("PreserveAspectRatio").Value = True
    Set thumb = proc.Apply(img)
    tmpPath = JoinPath(gThumbnailFolder, "~thumb_" & TimestampForFile() & "_" & CStr(Int(Rnd() * 100000)) & ".jpg")
    If gFSO.FileExists(tmpPath) Then gFSO.DeleteFile tmpPath, True
    thumb.SaveFile tmpPath
    If Err.Number = 0 And gFSO.FileExists(tmpPath) Then
        If gFSO.FileExists(thumbPath) Then gFSO.DeleteFile thumbPath, True
        gFSO.MoveFile tmpPath, thumbPath
        CreateThumbnailFile = gFSO.FileExists(thumbPath)
    End If
    Err.Clear
End Function

Private Function InsertThumbnailForRow(excelRow, thumbPath)
    On Error Resume Next
    InsertThumbnailForRow = False
    If Not INSERT_IMAGES_IN_EXCEL Then
        InsertThumbnailForRow = True
        Exit Function
    End If
    If thumbPath = "" Or Not gFSO.FileExists(thumbPath) Then Exit Function

    Dim cell
    Dim pic
    Set cell = gWsBom.Cells(CLng(excelRow), gThumbnailColumnIndex)
    gWsBom.Rows(CLng(excelRow)).RowHeight = CLng(THUMBNAIL_HEIGHT * 0.75) + 12
    Set pic = gWsBom.Shapes.AddPicture(thumbPath, MSO_FALSE, MSO_TRUE, cell.Left + 2, cell.Top + 2, -1, -1)
    pic.LockAspectRatio = MSO_TRUE
    If pic.Width > THUMBNAIL_WIDTH Then pic.Width = THUMBNAIL_WIDTH
    If pic.Height > THUMBNAIL_HEIGHT Then pic.Height = THUMBNAIL_HEIGHT
    pic.Left = cell.Left + ((cell.Width - pic.Width) / 2)
    pic.Top = cell.Top + ((cell.Height - pic.Height) / 2)
    WriteDebugPhase "EXCEL_THUMBNAIL_INSERTED", excelRow, GetPartNumberFromExcelRow(excelRow), thumbPath
    InsertThumbnailForRow = (Err.Number = 0)
    Err.Clear
End Function

Private Sub SetBomUtilityValues(rec, imagePath, thumbPath, statusText, skipReason)
    On Error Resume Next
    Dim rowIndex
    rowIndex = CLng(rec.Item("ExcelRow"))
    rec.Item("ImagePath") = imagePath
    rec.Item("ThumbnailPath") = thumbPath
    rec.Item("ExportStatus") = statusText
    rec.Item("ImageSkipReason") = skipReason
    gWsBom.Cells(rowIndex, gImagePathColumnIndex).Value = imagePath
    gWsBom.Cells(rowIndex, gThumbnailPathColumnIndex).Value = thumbPath
    gWsBom.Cells(rowIndex, gExportStatusColumnIndex).Value = statusText
    gWsBom.Cells(rowIndex, gImageSkipReasonColumnIndex).Value = skipReason
    Err.Clear
End Sub

Private Sub WriteExportLog(excelRow, partNumber, statusText, phase, messageText, imagePath, thumbPath)
    On Error Resume Next
    gWsLog.Cells(gNextLogRow, 1).Value = gNextLogRow - 1
    gWsLog.Cells(gNextLogRow, 2).Value = Now
    gWsLog.Cells(gNextLogRow, 3).Value = excelRow
    gWsLog.Cells(gNextLogRow, 4).Value = partNumber
    gWsLog.Cells(gNextLogRow, 5).Value = statusText
    gWsLog.Cells(gNextLogRow, 6).Value = phase
    gWsLog.Cells(gNextLogRow, 7).Value = messageText
    gWsLog.Cells(gNextLogRow, 8).Value = imagePath
    gWsLog.Cells(gNextLogRow, 9).Value = thumbPath
    gNextLogRow = gNextLogRow + 1
    Err.Clear
End Sub

Private Sub WriteDebugPhase(phase, excelRow, partNumber, messageText)
    On Error Resume Next
    Dim ts
    Set ts = gFSO.OpenTextFile(gDebugLogPath, 8, True)
    ts.WriteLine FormatDateTime(Now, 2) & " " & FormatDateTime(Now, 3) & " | " & phase & " | ExcelRow=" & CStr(excelRow) & " | PartNumber=" & CStr(partNumber) & " | " & CStr(messageText)
    ts.Close
    Err.Clear
End Sub

Private Sub SaveExcelCheckpoint(phase, excelRow, partNumber)
    On Error Resume Next
    If gWorkbook Is Nothing Then Exit Sub
    UpdateSummarySheet
    gExcelApp.DisplayAlerts = False
    If Not gWorkbookSaved Then
        gWorkbook.SaveAs gExcelPath, XL_OPENXML_WORKBOOK
        If Err.Number = 0 Then gWorkbookSaved = True
    Else
        gWorkbook.Save
    End If
    WriteDebugPhase "SAVE_CHECKPOINT", excelRow, partNumber, phase
    Err.Clear
End Sub

Private Sub MaybeSaveByProgress(excelRow, partNumber)
    If gSuccessfulImageRows = 1 Then
        SaveExcelCheckpoint "First successful image", excelRow, partNumber
    ElseIf SAVE_EVERY_N_ROWS > 0 Then
        If (gProcessedImageRows Mod SAVE_EVERY_N_ROWS) = 0 Then SaveExcelCheckpoint "Periodic save", excelRow, partNumber
    End If
End Sub

Private Sub AbortWithMessage(messageText, phase, excelRow, partNumber)
    On Error Resume Next
    gAbortMessage = messageText
    WriteDebugPhase "ERROR", excelRow, partNumber, phase & ": " & messageText
    WriteExportLog excelRow, partNumber, "ERROR", phase, messageText, "", ""
    SaveExcelCheckpoint phase, excelRow, partNumber
    CleanupCatiaSession
    MsgBox messageText, vbCritical, "CATIA_VISUAL_BOM_EXPORTER"
    Err.Clear
End Sub

Private Sub CleanupCatiaSession()
    On Error Resume Next
    CloseCurrentStandaloneDocument
    ActivateMainDocument
    If Not gActiveDoc Is Nothing Then gActiveDoc.Selection.Clear
    CATIA.RefreshDisplay = True
    CATIA.StatusBar = ""
    If Not gExcelApp Is Nothing Then
        gExcelApp.ScreenUpdating = True
        gExcelApp.DisplayAlerts = True
        gExcelApp.Visible = True
    End If
    Err.Clear
End Sub

Private Sub ActivateMainDocument()
    On Error Resume Next
    If Not gActiveDoc Is Nothing Then gActiveDoc.Activate
    Err.Clear
End Sub

Private Sub UpdateSummarySheet()
    On Error Resume Next
    If gWsSummary Is Nothing Then Exit Sub
    gWsSummary.Cells(3, 2).Value = GetProductPartNumber(gRootProduct)
    gWsSummary.Cells(4, 2).Value = Now
    gWsSummary.Cells(5, 2).Value = gOutputFolder
    gWsSummary.Cells(6, 2).Value = gExcelPath
    gWsSummary.Cells(7, 2).Value = ListCount(gBomRows)
    gWsSummary.Cells(8, 2).Value = gProcessedImageRows
    gWsSummary.Cells(9, 2).Value = gSuccessfulImageRows
    gWsSummary.Cells(10, 2).Value = gSkippedFastenerRows
    gWsSummary.Cells(11, 2).Value = gDebugLogPath
    gWsSummary.Cells(12, 2).Value = "TEST_MODE=" & CStr(TEST_MODE) & "; TEST_MAX_ROWS=" & CStr(TEST_MAX_ROWS) & "; STANDALONE_CAPTURE_ONLY=" & CStr(STANDALONE_CAPTURE_ONLY)
    Err.Clear
End Sub

Private Function SourcePathForPartNumber(partNumber)
    SourcePathForPartNumber = ""
    If gSourceIndex.Exists(NormalizePartNumber(partNumber)) Then SourcePathForPartNumber = CStr(gSourceIndex.Item(NormalizePartNumber(partNumber)))
End Function

Private Function ReuseExistingOrCachedImage(partNumber, imagePath, thumbPath, rec)
    On Error Resume Next
    ReuseExistingOrCachedImage = False
    Dim cacheKey
    Dim payload
    Dim parts
    cacheKey = NormalizePartNumber(partNumber)
    If gImageCache.Exists(cacheKey) Then
        payload = CStr(gImageCache.Item(cacheKey))
        parts = Split(payload, "|")
        If UBound(parts) >= 1 Then
            imagePath = CStr(parts(0))
            thumbPath = CStr(parts(1))
        End If
    End If
    If SKIP_EXISTING_IMAGES And gFSO.FileExists(imagePath) Then
        If Not gFSO.FileExists(thumbPath) Then CreateThumbnailFile imagePath, thumbPath, THUMBNAIL_WIDTH, THUMBNAIL_HEIGHT
        If gFSO.FileExists(thumbPath) Then
            CacheImage partNumber, imagePath, thumbPath
            ReuseExistingOrCachedImage = True
        End If
    End If
    Err.Clear
End Function

Private Sub CacheImage(partNumber, imagePath, thumbPath)
    If NormalizePartNumber(partNumber) <> "" Then gImageCache.Item(NormalizePartNumber(partNumber)) = imagePath & "|" & thumbPath
End Sub

Private Function BuildImagePath(partNumber, descriptionText, excelRow)
    Dim baseName
    baseName = SafeFileName(Trim(CStr(partNumber) & "_" & CStr(descriptionText)))
    If baseName = "" Or baseName = "_" Then baseName = "BOM_ROW_" & CStr(excelRow)
    If Len(baseName) > 140 Then baseName = Left(baseName, 140)
    BuildImagePath = JoinPath(gImageFolder, baseName & "." & FAST_IMAGE_FORMAT)
End Function

Private Function ThumbnailPathForImage(imagePath)
    ThumbnailPathForImage = JoinPath(gThumbnailFolder, gFSO.GetBaseName(imagePath) & ".jpg")
End Function

Private Function GetPartNumberFromRow(rec)
    If gPartNumberColumnIndex > 0 Then
        GetPartNumberFromRow = GetNativeValue(rec, gPartNumberColumnIndex)
    Else
        GetPartNumberFromRow = GetNativeValue(rec, 1)
    End If
End Function

Private Function GetDescriptionFromRow(rec)
    If gDescriptionColumnIndex > 0 Then GetDescriptionFromRow = GetNativeValue(rec, gDescriptionColumnIndex)
End Function

Private Function GetNativeValue(rec, nativeColumnIndex)
    On Error Resume Next
    Dim values
    Set values = rec.Item("Values")
    GetNativeValue = ""
    If nativeColumnIndex > 0 And nativeColumnIndex <= ListCount(values) Then GetNativeValue = CStr(ListValue(values, nativeColumnIndex))
    Err.Clear
End Function

Private Function GetPartNumberFromExcelRow(excelRow)
    Dim idx
    idx = CLng(excelRow) - 1
    If idx >= 1 And idx <= ListCount(gBomRows) Then GetPartNumberFromExcelRow = GetPartNumberFromRow(ListObject(gBomRows, idx))
End Function

Private Function IsFastenerBomRow(rec)
    On Error Resume Next
    IsFastenerBomRow = False
    Dim scanText
    Dim i
    Dim keywords
    Dim kw
    Dim normalizedKw
    scanText = ""
    For i = 1 To ListCount(gNativeHeaders)
        scanText = scanText & " " & GetNativeValue(rec, i)
    Next
    scanText = NormalizeFastenerText(scanText)
    keywords = Split(FASTENER_KEYWORDS, "|")
    For Each kw In keywords
        normalizedKw = NormalizeFastenerText(CStr(kw))
        If normalizedKw <> "" Then
            If InStr(1, scanText, normalizedKw, vbTextCompare) > 0 Then
                IsFastenerBomRow = True
                Exit Function
            End If
        End If
    Next
    Err.Clear
End Function

Private Function NormalizeFastenerText(valueText)
    Dim s
    s = LCase(CStr(valueText))
    s = Replace(s, ChrW(&H10D), "c")
    s = Replace(s, ChrW(&H107), "c")
    s = Replace(s, ChrW(&H161), "s")
    s = Replace(s, ChrW(&H111), "d")
    s = Replace(s, ChrW(&H17E), "z")
    s = Replace(s, ".", " ")
    s = Replace(s, "-", " ")
    s = Replace(s, "_", " ")
    s = Replace(s, "/", " ")
    s = Replace(s, "\", " ")
    s = Replace(s, vbTab, " ")
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    Do While InStr(1, s, "  ", vbTextCompare) > 0
        s = Replace(s, "  ", " ")
    Loop
    NormalizeFastenerText = Trim(s)
End Function

Private Function GetProductPartNumber(prod)
    On Error Resume Next
    GetProductPartNumber = Trim(CStr(prod.PartNumber))
    If GetProductPartNumber = "" Then
        Dim refProd
        Set refProd = prod.ReferenceProduct
        If Not refProd Is Nothing Then GetProductPartNumber = Trim(CStr(refProd.PartNumber))
    End If
    Err.Clear
End Function

Private Function GetProductSourceFilePath(prod)
    On Error Resume Next
    Dim refProd
    Dim pathText
    pathText = TryGetMasterShapePath(prod)
    If pathText <> "" Then
        GetProductSourceFilePath = pathText
        Exit Function
    End If
    Set refProd = prod.ReferenceProduct
    If Not refProd Is Nothing Then
        pathText = FindSourcePathInObjectChain(refProd)
        If pathText <> "" Then
            GetProductSourceFilePath = pathText
            Exit Function
        End If
        pathText = TryGetMasterShapePath(refProd)
        If pathText <> "" Then
            GetProductSourceFilePath = pathText
            Exit Function
        End If
    End If
    pathText = FindSourcePathInObjectChain(prod)
    If pathText <> "" Then GetProductSourceFilePath = pathText
    Err.Clear
End Function

Private Function TryGetMasterShapePath(obj)
    On Error Resume Next
    TryGetMasterShapePath = CStr(obj.GetMasterShapeRepresentationPathName)
    If Err.Number <> 0 Then
        TryGetMasterShapePath = ""
        Err.Clear
    End If
End Function

Private Function FindSourcePathInObjectChain(startObj)
    On Error Resume Next
    Dim cur
    Dim i
    Dim candidate
    Set cur = startObj
    For i = 1 To 10
        If cur Is Nothing Then Exit For
        candidate = ""
        candidate = CStr(cur.FullName)
        If LooksLikeFilePath(candidate) Then
            FindSourcePathInObjectChain = candidate
            Exit Function
        End If
        Err.Clear
        If CStr(cur.Path) <> "" And CStr(cur.Name) <> "" Then
            candidate = JoinPath(CStr(cur.Path), CStr(cur.Name))
            If LooksLikeFilePath(candidate) Then
                FindSourcePathInObjectChain = candidate
                Exit Function
            End If
        End If
        Err.Clear
        Set cur = cur.Parent
        If Err.Number <> 0 Then Exit For
    Next
    Err.Clear
End Function

Private Function LooksLikeFilePath(pathText)
    LooksLikeFilePath = (CStr(pathText) <> "" And gFSO.GetExtensionName(CStr(pathText)) <> "")
End Function

Private Function IsSupportedCatiaSourceFile(sourcePath)
    Dim ext
    ext = LCase(gFSO.GetExtensionName(CStr(sourcePath)))
    IsSupportedCatiaSourceFile = (ext = "catpart" Or ext = "catproduct")
End Function

Private Function FindOpenDocumentByFullName(sourcePath, foundDoc)
    On Error Resume Next
    FindOpenDocumentByFullName = False
    Dim i
    Dim doc
    For i = 1 To CATIA.Documents.Count
        Set doc = CATIA.Documents.Item(i)
        If SamePath(GetDocumentFullName(doc), sourcePath) Then
            Set foundDoc = doc
            FindOpenDocumentByFullName = True
            Exit Function
        End If
    Next
    Err.Clear
End Function

Private Function GetDocumentFullName(doc)
    On Error Resume Next
    GetDocumentFullName = CStr(doc.FullName)
    If Err.Number <> 0 Or GetDocumentFullName = "" Then
        Err.Clear
        If CStr(doc.Path) <> "" Then GetDocumentFullName = JoinPath(CStr(doc.Path), CStr(doc.Name))
    End If
    Err.Clear
End Function

Private Function SamePath(pathA, pathB)
    SamePath = (UCase(Trim(CStr(pathA))) = UCase(Trim(CStr(pathB))) And Trim(CStr(pathA)) <> "")
End Function

Private Function NormalizePartNumber(partNumber)
    NormalizePartNumber = UCase(Trim(CStr(partNumber)))
End Function

Private Function DetectDelimiter(lineText)
    Dim delims
    Dim bestDelim
    Dim bestCount
    Dim d
    Dim cnt
    delims = Array(vbTab, ";", "|", ",")
    bestDelim = ""
    bestCount = 0
    For Each d In delims
        cnt = CountSubstring(CStr(lineText), CStr(d))
        If cnt > bestCount Then
            bestCount = cnt
            bestDelim = CStr(d)
        End If
    Next
    If bestCount > 0 Then DetectDelimiter = bestDelim
End Function

Private Function ParseDelimitedLine(lineText, delim)
    Dim result
    Dim current
    Dim i
    Dim ch
    Dim inQuotes
    result = EmptyArray()
    current = ""
    inQuotes = False
    For i = 1 To Len(lineText)
        ch = Mid(lineText, i, 1)
        If ch = """" Then
            If inQuotes And i < Len(lineText) And Mid(lineText, i + 1, 1) = """" Then
                current = current & """"
                i = i + 1
            Else
                inQuotes = Not inQuotes
            End If
        ElseIf ch = delim And Not inQuotes Then
            result = ArrayPush(result, current)
            current = ""
        Else
            current = current & ch
        End If
    Next
    result = ArrayPush(result, current)
    ParseDelimitedLine = result
End Function

Private Function LooksLikeBomHeader(cells)
    Dim text
    Dim i
    text = ""
    For i = LBound(cells) To UBound(cells)
        text = text & " " & NormalizeHeaderName(CStr(cells(i)))
    Next
    LooksLikeBomHeader = (CountUsefulCells(cells) >= 2 And _
        (InStr(text, "partnumber") > 0 Or InStr(text, "number") > 0 Or InStr(text, "quantity") > 0 Or _
         InStr(text, "qty") > 0 Or InStr(text, "nomenclature") > 0 Or InStr(text, "description") > 0 Or _
         InStr(text, "brojdela") > 0 Or InStr(text, "kolicina") > 0))
End Function

Private Function CountUsefulCells(cells)
    Dim i
    CountUsefulCells = 0
    On Error Resume Next
    For i = LBound(cells) To UBound(cells)
        If Trim(CStr(cells(i))) <> "" Then CountUsefulCells = CountUsefulCells + 1
    Next
    Err.Clear
End Function

Private Function IsSeparatorRow(cells)
    Dim i
    Dim s
    IsSeparatorRow = True
    For i = LBound(cells) To UBound(cells)
        s = Replace(Replace(Trim(CStr(cells(i))), "-", ""), "=", "")
        If s <> "" Then
            IsSeparatorRow = False
            Exit Function
        End If
    Next
End Function

Private Function FindHeaderIndexInList(headers, possibleNames)
    Dim names
    Dim nm
    Dim i
    Dim h
    names = Split(possibleNames, "|")
    For i = 1 To ListCount(headers)
        h = NormalizeHeaderName(CStr(ListValue(headers, i)))
        For Each nm In names
            If h = NormalizeHeaderName(CStr(nm)) Then
                FindHeaderIndexInList = i
                Exit Function
            End If
        Next
    Next
End Function

Private Function NormalizeHeaderName(valueText)
    Dim s
    s = LCase(Trim(CStr(valueText)))
    s = Replace(s, ChrW(&H10D), "c")
    s = Replace(s, ChrW(&H107), "c")
    s = Replace(s, ChrW(&H161), "s")
    s = Replace(s, ChrW(&H111), "d")
    s = Replace(s, ChrW(&H17E), "z")
    s = Replace(s, ".", "")
    s = Replace(s, " ", "")
    s = Replace(s, "_", "")
    s = Replace(s, "-", "")
    NormalizeHeaderName = s
End Function

Private Function HtmlToText(htmlText)
    Dim re
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.IgnoreCase = True
    re.Pattern = "<[^>]+>"
    HtmlToText = re.Replace(CStr(htmlText), "")
    HtmlToText = Replace(HtmlToText, "&nbsp;", " ")
    HtmlToText = Replace(HtmlToText, "&amp;", "&")
    HtmlToText = Replace(HtmlToText, "&lt;", "<")
    HtmlToText = Replace(HtmlToText, "&gt;", ">")
    HtmlToText = Replace(HtmlToText, "&quot;", """")
    HtmlToText = Trim(HtmlToText)
End Function

Private Function JoinHeaderNames()
    Dim i
    Dim s
    s = ""
    For i = 1 To ListCount(gNativeHeaders)
        If s <> "" Then s = s & " | "
        s = s & CStr(ListValue(gNativeHeaders, i))
    Next
    JoinHeaderNames = s
End Function

Private Sub ApplyBomBorders()
    On Error Resume Next
    Dim lastRow
    Dim lastCol
    lastRow = ListCount(gBomRows) + 1
    lastCol = gImageSkipReasonColumnIndex
    gWsBom.Range(gWsBom.Cells(1, 1), gWsBom.Cells(lastRow, lastCol)).Borders.LineStyle = XL_CONTINUOUS
    gWsBom.Range(gWsBom.Cells(1, 1), gWsBom.Cells(lastRow, lastCol)).Borders.Weight = XL_THIN
    gWsBom.Range(gWsBom.Cells(1, 1), gWsBom.Cells(lastRow, lastCol)).VerticalAlignment = XL_TOP
    Err.Clear
End Sub

Private Function CountSubstring(textValue, token)
    Dim p
    Dim startAt
    If token = "" Then Exit Function
    startAt = 1
    Do
        p = InStr(startAt, textValue, token, vbBinaryCompare)
        If p <= 0 Then Exit Do
        CountSubstring = CountSubstring + 1
        startAt = p + Len(token)
    Loop
End Function

Private Function EmptyArray()
    EmptyArray = Array()
End Function

Private Function ArrayPush(arr, value)
    Dim n
    Dim tmp()
    Dim i
    On Error Resume Next
    n = UBound(arr) + 1
    If Err.Number <> 0 Then
        Err.Clear
        ReDim tmp(0)
        tmp(0) = value
        ArrayPush = tmp
        Exit Function
    End If
    ReDim tmp(n)
    For i = 0 To n - 1
        tmp(i) = arr(i)
    Next
    tmp(n) = value
    ArrayPush = tmp
End Function

Private Function NewList()
    Set NewList = CreateObject("Scripting.Dictionary")
End Function

Private Sub ListAddValue(listObj, valueText)
    listObj.Item(CStr(listObj.Count + 1)) = valueText
End Sub

Private Sub ListAddObject(listObj, obj)
    Set listObj.Item(CStr(listObj.Count + 1)) = obj
End Sub

Private Function ListCount(listObj)
    ListCount = listObj.Count
End Function

Private Function ListValue(listObj, indexNumber)
    ListValue = listObj.Item(CStr(indexNumber))
End Function

Private Function ListObject(listObj, indexNumber)
    Set ListObject = listObj.Item(CStr(indexNumber))
End Function

Private Function FirstNonEmpty(a, b, c)
    If Trim(CStr(a)) <> "" Then
        FirstNonEmpty = CStr(a)
    ElseIf Trim(CStr(b)) <> "" Then
        FirstNonEmpty = CStr(b)
    Else
        FirstNonEmpty = CStr(c)
    End If
End Function

Private Sub EnsureFolder(folderPath)
    If Not gFSO.FolderExists(folderPath) Then gFSO.CreateFolder folderPath
End Sub

Private Function JoinPath(folderPath, fileName)
    If Right(CStr(folderPath), 1) = "\" Then
        JoinPath = CStr(folderPath) & CStr(fileName)
    Else
        JoinPath = CStr(folderPath) & "\" & CStr(fileName)
    End If
End Function

Private Function SafeFileName(valueText)
    Dim s
    Dim badChars
    Dim ch
    s = Trim(CStr(valueText))
    badChars = Array("\", "/", ":", "*", "?", """", "<", ">", "|", vbTab, vbCr, vbLf)
    For Each ch In badChars
        s = Replace(s, CStr(ch), "_")
    Next
    Do While InStr(1, s, "__", vbTextCompare) > 0
        s = Replace(s, "__", "_")
    Loop
    SafeFileName = s
End Function

Private Function TimestampForFile()
    TimestampForFile = CStr(Year(Now)) & Pad2(Month(Now)) & Pad2(Day(Now)) & "_" & Pad2(Hour(Now)) & Pad2(Minute(Now)) & Pad2(Second(Now))
End Function

Private Function Pad2(n)
    If CLng(n) < 10 Then
        Pad2 = "0" & CStr(n)
    Else
        Pad2 = CStr(n)
    End If
End Function

Private Sub WaitSeconds(secondsValue)
    Dim startTime
    Dim currentTime
    startTime = Timer
    Do
        DoEvents
        currentTime = Timer
        If currentTime < startTime Then currentTime = currentTime + 86400
        If (currentTime - startTime) >= CDbl(secondsValue) Then Exit Do
    Loop
End Sub
