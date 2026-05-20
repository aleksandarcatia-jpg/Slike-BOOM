' ================================================================
' CATIA_VISUAL_BOM_EXPORTER
' CATIA BillOfMaterial.Print "XLS" -> Excel -> standalone thumbnails.
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
Const THUMBNAIL_WIDTH = 160
Const THUMBNAIL_HEIGHT = 120
Const CREATE_THUMBNAIL_FILES = True
Const INSERT_THUMBNAIL_FILE_IN_EXCEL = True
Const INSERT_IMAGES_IN_EXCEL = True
Const TEST_MODE = True
Const TEST_MAX_ROWS = 20
Const SAVE_EVERY_N_ROWS = 25
Const RESUME_MODE = True
Const SKIP_EXISTING_IMAGES = True
Const OUTPUT_TO_DESKTOP = False
Const USE_SHADED_WITH_EDGES = True
Const USE_WHITE_BACKGROUND = True
Const USE_PARALLEL_PROJECTION = True
Const IMAGE_CAPTURE_DELAY_SECONDS = 0.15

' CATIA constants used late-bound.
Const CAT_CAPTURE_FORMAT_JPEG = 2
Const CAT_RENDER_SHADING = 0
Const CAT_RENDER_SHADING_WITH_EDGES = 1
Const CAT_PROJECTION_CYLINDRIC = 0

' Excel constants used late-bound.
Const XL_OPENXML_WORKBOOK = 51
Const XL_LEFT = -4131
Const XL_TOP = -4160
Const XL_CENTER = -4108
Const XL_CONTINUOUS = 1
Const XL_THIN = 2
Const MSO_TRUE = -1
Const MSO_FALSE = 0

Dim gFSO
Dim gShell
Dim gProductDocument
Dim gProduct
Dim gAssemblyConvertor
Dim gMainDocumentFullName
Dim gExcelPath
Dim gOutputFolder
Dim gImageFolder
Dim gThumbnailFolder
Dim gDebugLogPath
Dim gExcelApp
Dim gWorkbook
Dim gWsBom
Dim gWsLog
Dim gWsSummary
Dim gHeaderRow
Dim gPartNumberColumnIndex
Dim gThumbnailColumnIndex
Dim gImagePathColumnIndex
Dim gThumbnailPathColumnIndex
Dim gExportStatusColumnIndex
Dim gImageSkipReasonColumnIndex
Dim gLastBomRow
Dim gSourceIndex
Dim gImageCache
Dim gNextLogRow
Dim gProcessedImageRows
Dim gSuccessfulImageRows
Dim gSkippedFastenerRows
Dim gErrorCount
Dim gCurrentStandaloneDoc
Dim gCurrentStandaloneOpened

Public Sub CATIA_VISUAL_BOM_EXPORTER()
    On Error Resume Next

    InitializeRuntime
    If Not ValidateActiveProductDocument() Then Exit Sub
    If Not PrepareBillOfMaterialFormat() Then Exit Sub
    If Not SelectExportPathAndPrepareFolders() Then Exit Sub

    WriteDebugPhase "START", 0, "", "", "", "CATIA_VISUAL_BOM_EXPORTER started."
    WriteDebugPhase "USER_SAVE_PATH_SELECTED", 0, "", "", gExcelPath, gExcelPath

    If Not PrintCatiaBomToXls() Then
        AbortWithMessage "CATIA BillOfMaterial Print XLS nije uspeo.", "CATIA_BOM_PRINT_XLS_START", 0, "", "", ""
        Exit Sub
    End If

    If Not OpenBomWorkbookAndPrepareSheets() Then
        AbortWithMessage "Ne mogu da otvorim Excel BOM fajl: " & gExcelPath, "EXCEL_OPENED", 0, "", "", ""
        Exit Sub
    End If

    WriteDebugPhase "CATIA_FILE_INDEX_START", 0, "", "", "", "Building Part Number -> source file path index."
    BuildCatiaFileIndex
    WriteDebugPhase "CATIA_FILE_INDEX_DONE", 0, "", "", "", "Indexed source files: " & CStr(gSourceIndex.Count)

    If Not ProcessBomRowsForThumbnails() Then Exit Sub

    SaveExcelCheckpoint "FINISH", 0, "", ""
    CleanupCatiaSession
    WriteDebugPhase "FINISH", 0, "", "", "", "Macro finished."
    MsgBox "CATIA BOM Excel export je zavrsen." & vbCrLf & vbCrLf & _
           "Excel:" & vbCrLf & gExcelPath, vbInformation, "CATIA_VISUAL_BOM_EXPORTER"
End Sub

Sub InitializeRuntime()
    Set gFSO = CreateObject("Scripting.FileSystemObject")
    Set gShell = CreateObject("WScript.Shell")
    Set gSourceIndex = CreateObject("Scripting.Dictionary")
    Set gImageCache = CreateObject("Scripting.Dictionary")
    gSourceIndex.CompareMode = 1
    gImageCache.CompareMode = 1
    Set gExcelApp = Nothing
    Set gWorkbook = Nothing
    Set gWsBom = Nothing
    Set gWsLog = Nothing
    Set gWsSummary = Nothing
    Set gCurrentStandaloneDoc = Nothing
    gCurrentStandaloneOpened = False
    gHeaderRow = 0
    gPartNumberColumnIndex = 0
    gThumbnailColumnIndex = 0
    gImagePathColumnIndex = 0
    gThumbnailPathColumnIndex = 0
    gExportStatusColumnIndex = 0
    gImageSkipReasonColumnIndex = 0
    gLastBomRow = 0
    gNextLogRow = 2
    gProcessedImageRows = 0
    gSuccessfulImageRows = 0
    gSkippedFastenerRows = 0
    gErrorCount = 0
    Randomize
End Sub

Function ValidateActiveProductDocument()
    On Error Resume Next
    ValidateActiveProductDocument = False
    If CATIA.Documents.Count = 0 Then
        MsgBox "Nema otvorenog CATIA dokumenta. Otvorite CATProduct i pokrenite makro.", vbExclamation, "CATIA_VISUAL_BOM_EXPORTER"
        Exit Function
    End If

    Set gProductDocument = CATIA.ActiveDocument
    Set gProduct = gProductDocument.Product
    If Err.Number <> 0 Or gProduct Is Nothing Then
        Err.Clear
        MsgBox "Aktivni dokument mora biti CATProduct.", vbExclamation, "CATIA_VISUAL_BOM_EXPORTER"
        Exit Function
    End If

    gMainDocumentFullName = GetDocumentFullName(gProductDocument)
    ValidateActiveProductDocument = True
    Err.Clear
End Function

Function PrepareBillOfMaterialFormat()
    On Error Resume Next
    PrepareBillOfMaterialFormat = False
    Set gAssemblyConvertor = gProduct.GetItem("BillOfMaterial")
    If Err.Number <> 0 Or gAssemblyConvertor Is Nothing Then
        Err.Clear
        MsgBox "CATIA BillOfMaterial objekat nije dostupan preko product.GetItem(""BillOfMaterial"").", vbCritical, "CATIA_VISUAL_BOM_EXPORTER"
        Exit Function
    End If

    Dim arrayOfVariantOfBSTR1(6)
    arrayOfVariantOfBSTR1(0) = "Nomenclature"
    arrayOfVariantOfBSTR1(1) = "Quantity"
    arrayOfVariantOfBSTR1(2) = "Part Number"
    arrayOfVariantOfBSTR1(3) = "Dimenzija"
    arrayOfVariantOfBSTR1(4) = "Material"
    arrayOfVariantOfBSTR1(5) = "Mass"
    arrayOfVariantOfBSTR1(6) = "Standard"

    gAssemblyConvertor.SetSecondaryFormat arrayOfVariantOfBSTR1
    If Err.Number <> 0 Then
        MsgBox "SetSecondaryFormat za CATIA BOM nije uspeo: " & Err.Description, vbCritical, "CATIA_VISUAL_BOM_EXPORTER"
        Err.Clear
        Exit Function
    End If

    PrepareBillOfMaterialFormat = True
    Err.Clear
End Function

Function SelectExportPathAndPrepareFolders()
    On Error Resume Next
    SelectExportPathAndPrepareFolders = False

    Dim defaultFolder
    Dim defaultName
    Dim defaultPath
    defaultFolder = CStr(gProductDocument.Path)
    If defaultFolder = "" Then defaultFolder = gShell.SpecialFolders("Desktop")
    defaultName = SafeFileName(GetProductPartNumber(gProduct))
    If defaultName = "" Then defaultName = SafeFileName(gProduct.Name)
    If defaultName = "" Then defaultName = "CATIA_BOM"
    defaultPath = JoinPath(defaultFolder, defaultName & "_VISUAL_BOM_EXPORT.xls")

    gExcelPath = AskUserForBomExcelSavePath(defaultPath)
    If gExcelPath = "" Then
        MsgBox "Export je otkazan od strane korisnika.", vbInformation, "CATIA_VISUAL_BOM_EXPORTER"
        Exit Function
    End If

    gOutputFolder = JoinPath(gFSO.GetParentFolderName(gExcelPath), gFSO.GetBaseName(gExcelPath) & "_FILES")
    gImageFolder = JoinPath(gOutputFolder, "IMAGES")
    gThumbnailFolder = JoinPath(gOutputFolder, "THUMBNAILS")
    gDebugLogPath = JoinPath(gOutputFolder, "DEBUG_PHASE_LOG.txt")
    EnsureFolder gOutputFolder
    EnsureFolder gImageFolder
    EnsureFolder gThumbnailFolder
    SelectExportPathAndPrepareFolders = True
    Err.Clear
End Function

Function AskUserForBomExcelSavePath(defaultPath)
    On Error Resume Next
    AskUserForBomExcelSavePath = ""
    Dim xl
    Dim selectedPath
    Set xl = CreateObject("Excel.Application")
    If Err.Number <> 0 Or xl Is Nothing Then
        Err.Clear
        MsgBox "Microsoft Excel nije dostupan za Save As dialog.", vbCritical, "CATIA_VISUAL_BOM_EXPORTER"
        Exit Function
    End If

    xl.Visible = False
    selectedPath = xl.GetSaveAsFilename(defaultPath, "Excel 97-2003 Workbook (*.xls), *.xls", , "Sacuvaj CATIA BOM Excel export")
    xl.Quit
    Set xl = Nothing

    If VarType(selectedPath) = vbBoolean Then
        AskUserForBomExcelSavePath = ""
    Else
        AskUserForBomExcelSavePath = CStr(selectedPath)
        If LCase(Right(AskUserForBomExcelSavePath, 4)) <> ".xls" Then AskUserForBomExcelSavePath = AskUserForBomExcelSavePath & ".xls"
    End If
    Err.Clear
End Function

Function PrintCatiaBomToXls()
    On Error Resume Next
    PrintCatiaBomToXls = False
    WriteDebugPhase "BOM_FORMAT_SET", 0, "", "", "", "Secondary BOM format set: Nomenclature|Quantity|Part Number|Dimenzija|Material|Mass|Standard"
    WriteDebugPhase "CATIA_BOM_PRINT_XLS_START", 0, "", "", gExcelPath, "assemblyConvertor.Print ""XLS"""

    If gFSO.FileExists(gExcelPath) Then gFSO.DeleteFile gExcelPath, True
    Err.Clear
    gAssemblyConvertor.Print "XLS", gExcelPath, gProduct
    If Err.Number = 0 And gFSO.FileExists(gExcelPath) Then
        WriteDebugPhase "CATIA_BOM_PRINT_XLS_DONE", 0, "", "", gExcelPath, "CATIA BOM XLS created."
        PrintCatiaBomToXls = True
    End If
    Err.Clear
End Function

Function OpenBomWorkbookAndPrepareSheets()
    On Error Resume Next
    OpenBomWorkbookAndPrepareSheets = False
    Set gExcelApp = CreateObject("Excel.Application")
    If Err.Number <> 0 Or gExcelApp Is Nothing Then
        Err.Clear
        Exit Function
    End If

    gExcelApp.Visible = False
    gExcelApp.DisplayAlerts = False
    gExcelApp.ScreenUpdating = False
    Set gWorkbook = gExcelApp.Workbooks.Open(gExcelPath)
    If Err.Number <> 0 Or gWorkbook Is Nothing Then
        Err.Clear
        Exit Function
    End If

    Set gWsBom = gWorkbook.Worksheets(1)
    gWsBom.Name = "BOM"
    If Not FindBomHeaderRowAndPartNumberColumn() Then
        SaveExcelCheckpoint "Part Number header missing", 0, "", ""
        MsgBox "Kolona 'Part Number' nije pronadjena u CATIA BOM Excel fajlu. Slikanje je prekinuto.", vbCritical, "CATIA_VISUAL_BOM_EXPORTER"
        Exit Function
    End If
    WriteDebugPhase "EXCEL_OPENED", 0, "", "", gExcelPath, "Workbook opened."
    WriteDebugPhase "BOM_HEADERS_READ", gHeaderRow, "", "", "", "Part Number column=" & CStr(gPartNumberColumnIndex)

    EnsureHelperColumns
    Set gWsLog = GetOrCreateWorksheet("EXPORT_LOG")
    PrepareLogSheet
    Set gWsSummary = GetOrCreateWorksheet("SUMMARY")
    PrepareSummarySheet
    gLastBomRow = LastUsedRow(gWsBom)
    SaveExcelCheckpoint "EXCEL_OPENED", 0, "", ""
    OpenBomWorkbookAndPrepareSheets = True
    Err.Clear
End Function

Function FindBomHeaderRowAndPartNumberColumn()
    On Error Resume Next
    FindBomHeaderRowAndPartNumberColumn = False
    Dim used
    Dim maxRow
    Dim maxCol
    Dim r
    Dim c
    Set used = gWsBom.UsedRange
    maxRow = used.Row + used.Rows.Count - 1
    maxCol = used.Column + used.Columns.Count - 1
    If maxRow > 30 Then maxRow = 30
    If maxCol > 120 Then maxCol = 120

    For r = 1 To maxRow
        For c = 1 To maxCol
            If IsPartNumberHeader(CStr(gWsBom.Cells(r, c).Value)) Then
                gHeaderRow = r
                gPartNumberColumnIndex = c
                FindBomHeaderRowAndPartNumberColumn = True
                Exit Function
            End If
        Next
    Next
    Err.Clear
End Function

Function IsPartNumberHeader(valueText)
    Dim h
    h = NormalizeHeaderName(valueText)
    IsPartNumberHeader = (h = "partnumber" Or h = "number" Or h = "partno" Or h = "brojdela" Or h = "brdela")
End Function

Sub EnsureHelperColumns()
    On Error Resume Next
    gThumbnailColumnIndex = FindHeaderColumn("Thumbnail")
    If gThumbnailColumnIndex = 0 Then
        gWsBom.Columns(gPartNumberColumnIndex + 1).Insert
        gWsBom.Cells(gHeaderRow, gPartNumberColumnIndex + 1).Value = "Thumbnail"
        gThumbnailColumnIndex = gPartNumberColumnIndex + 1
    End If

    gImagePathColumnIndex = EnsureEndHelperColumn("Image Path")
    gThumbnailPathColumnIndex = EnsureEndHelperColumn("Thumbnail Path")
    gExportStatusColumnIndex = EnsureEndHelperColumn("Export Status")
    gImageSkipReasonColumnIndex = EnsureEndHelperColumn("Image Skip Reason")

    gWsBom.Rows(gHeaderRow).Font.Bold = True
    gWsBom.Columns(gThumbnailColumnIndex).ColumnWidth = 24
    gWsBom.Columns(gImagePathColumnIndex).ColumnWidth = 45
    gWsBom.Columns(gThumbnailPathColumnIndex).ColumnWidth = 45
    gWsBom.Columns(gExportStatusColumnIndex).ColumnWidth = 22
    gWsBom.Columns(gImageSkipReasonColumnIndex).ColumnWidth = 42
    ApplyBomBorders
    Err.Clear
End Sub

Function EnsureEndHelperColumn(headerText)
    On Error Resume Next
    Dim existing
    existing = FindHeaderColumn(headerText)
    If existing > 0 Then
        EnsureEndHelperColumn = existing
    Else
        EnsureEndHelperColumn = LastUsedColumn(gWsBom) + 1
        gWsBom.Cells(gHeaderRow, EnsureEndHelperColumn).Value = headerText
    End If
    Err.Clear
End Function

Function FindHeaderColumn(headerText)
    On Error Resume Next
    Dim used
    Dim maxCol
    Dim c
    Set used = gWsBom.UsedRange
    maxCol = used.Column + used.Columns.Count - 1
    For c = 1 To maxCol
        If NormalizeHeaderName(CStr(gWsBom.Cells(gHeaderRow, c).Value)) = NormalizeHeaderName(headerText) Then
            FindHeaderColumn = c
            Exit Function
        End If
    Next
    Err.Clear
End Function

Function GetOrCreateWorksheet(sheetName)
    On Error Resume Next
    Dim ws
    Set ws = gWorkbook.Worksheets(sheetName)
    If Err.Number <> 0 Or ws Is Nothing Then
        Err.Clear
        Set ws = gWorkbook.Worksheets.Add(, gWorkbook.Worksheets(gWorkbook.Worksheets.Count))
        ws.Name = sheetName
    Else
        ws.Cells.Clear
    End If
    Set GetOrCreateWorksheet = ws
    Err.Clear
End Function

Sub PrepareLogSheet()
    gWsLog.Cells(1, 1).Value = "No."
    gWsLog.Cells(1, 2).Value = "Date/Time"
    gWsLog.Cells(1, 3).Value = "Excel Row"
    gWsLog.Cells(1, 4).Value = "Raw Part Number"
    gWsLog.Cells(1, 5).Value = "Normalized Part Number"
    gWsLog.Cells(1, 6).Value = "Status"
    gWsLog.Cells(1, 7).Value = "Phase"
    gWsLog.Cells(1, 8).Value = "Message"
    gWsLog.Cells(1, 9).Value = "Image Path"
    gWsLog.Cells(1, 10).Value = "Thumbnail Path"
    gWsLog.Cells(1, 11).Value = "Source Path"
    gWsLog.Rows(1).Font.Bold = True
    gWsLog.Columns("A:K").ColumnWidth = 24
End Sub

Sub PrepareSummarySheet()
    gWsSummary.Cells(1, 1).Value = "CATIA VISUAL BOM EXPORTER"
    gWsSummary.Cells(3, 1).Value = "Main Assembly Part Number"
    gWsSummary.Cells(4, 1).Value = "Export date/time"
    gWsSummary.Cells(5, 1).Value = "Excel path"
    gWsSummary.Cells(6, 1).Value = "Output folder"
    gWsSummary.Cells(7, 1).Value = "Total BOM rows"
    gWsSummary.Cells(8, 1).Value = "Images processed"
    gWsSummary.Cells(9, 1).Value = "Successful images"
    gWsSummary.Cells(10, 1).Value = "Skipped fasteners"
    gWsSummary.Cells(11, 1).Value = "Errors"
    gWsSummary.Cells(12, 1).Value = "Mode"
    gWsSummary.Cells(13, 1).Value = "Debug log path"
    gWsSummary.Range("A1:B1").Font.Bold = True
    gWsSummary.Columns("A").ColumnWidth = 28
    gWsSummary.Columns("B").ColumnWidth = 95
    UpdateSummarySheet
End Sub

Function ProcessBomRowsForThumbnails()
    On Error Resume Next
    ProcessBomRowsForThumbnails = False
    Dim rowIndex
    Dim rawPartNumber
    Dim normalizedPartNumber
    Dim sourcePath
    Dim imagePath
    Dim thumbPath
    Dim reason

    For rowIndex = gHeaderRow + 1 To gLastBomRow
        rawPartNumber = CStr(gWsBom.Cells(rowIndex, gPartNumberColumnIndex).Value)
        normalizedPartNumber = NormalizePartNumber(rawPartNumber)
        If Trim(rawPartNumber) <> "" Then
            CATIA.StatusBar = "BOM thumbnail row " & CStr(rowIndex - gHeaderRow) & " / " & CStr(gLastBomRow - gHeaderRow) & " - " & rawPartNumber
            WriteDebugPhase "ROW_START", rowIndex, rawPartNumber, normalizedPartNumber, "", "Raw BOM Part Number=" & rawPartNumber & "; Normalized BOM Part Number=" & normalizedPartNumber

            If SKIP_FASTENER_IMAGES And IsFastenerExcelRow(rowIndex) Then
                gSkippedFastenerRows = gSkippedFastenerRows + 1
                reason = "Standard fastener - retained in BOM, image skipped"
                SetRowUtilityValues rowIndex, "", "", "SKIPPED_IMAGE_ONLY", reason
                WriteExportLog rowIndex, rawPartNumber, normalizedPartNumber, "SKIPPED_IMAGE_ONLY", "FASTENER_SKIPPED_IMAGE", reason, "", "", ""
                WriteDebugPhase "FASTENER_SKIPPED_IMAGE", rowIndex, rawPartNumber, normalizedPartNumber, "", reason
            ElseIf TEST_MODE And gProcessedImageRows >= TEST_MAX_ROWS Then
                SetRowUtilityValues rowIndex, "", "", "NOT_PROCESSED_TEST_LIMIT", "TEST_MODE limit reached"
                WriteExportLog rowIndex, rawPartNumber, normalizedPartNumber, "NOT_PROCESSED_TEST_LIMIT", "ROW_START", "TEST_MODE limit reached.", "", "", ""
            Else
                gProcessedImageRows = gProcessedImageRows + 1
                imagePath = BuildImagePath(normalizedPartNumber, rowIndex)
                thumbPath = ThumbnailPathForImage(imagePath)

                If ReuseExistingOrCachedImage(normalizedPartNumber, imagePath, thumbPath) Then
                    SetRowUtilityValues rowIndex, imagePath, thumbPath, "EXISTING_REUSED", ""
                    If Not InsertThumbnailForRow(rowIndex, thumbPath) Then
                        AbortWithMessage "Excel ne moze da ubaci thumbnail za Part Number: " & rawPartNumber, "EXCEL_THUMBNAIL_INSERTED", rowIndex, rawPartNumber, normalizedPartNumber, ""
                        Exit Function
                    End If
                    gSuccessfulImageRows = gSuccessfulImageRows + 1
                    WriteExportLog rowIndex, rawPartNumber, normalizedPartNumber, "EXISTING_REUSED", "EXCEL_THUMBNAIL_INSERTED", "Existing image/thumbnail reused.", imagePath, thumbPath, ""
                    MaybeSaveByProgress rowIndex, rawPartNumber, normalizedPartNumber
                Else
                    sourcePath = SourcePathForPartNumber(rawPartNumber)
                    If sourcePath = "" Or Not gFSO.FileExists(sourcePath) Then
                        SetRowUtilityValues rowIndex, "", "", "SOURCE_FILE_NOT_FOUND", "Source file not found"
                        WriteExportLog rowIndex, rawPartNumber, normalizedPartNumber, "SOURCE_FILE_NOT_FOUND", "SOURCE_FILE_NOT_FOUND", "Source file not found for Part Number: " & rawPartNumber, "", "", sourcePath
                        WriteDebugPhase "SOURCE_FILE_NOT_FOUND", rowIndex, rawPartNumber, normalizedPartNumber, sourcePath, "Source file not found."
                        AbortWithMessage "Source file not found for Part Number: " & rawPartNumber, "SOURCE_FILE_NOT_FOUND", rowIndex, rawPartNumber, normalizedPartNumber, sourcePath
                        Exit Function
                    End If

                    If Not IsSupportedCatiaSourceFile(sourcePath) Then
                        SetRowUtilityValues rowIndex, "", "", "UNSUPPORTED_SOURCE_FILE", sourcePath
                        WriteExportLog rowIndex, rawPartNumber, normalizedPartNumber, "UNSUPPORTED_SOURCE_FILE", "UNSUPPORTED_SOURCE_FILE", sourcePath, "", "", sourcePath
                        AbortWithMessage "Unsupported source file for Part Number: " & rawPartNumber & vbCrLf & sourcePath, "UNSUPPORTED_SOURCE_FILE", rowIndex, rawPartNumber, normalizedPartNumber, sourcePath
                        Exit Function
                    End If

                    WriteDebugPhase "SOURCE_FILE_FOUND", rowIndex, rawPartNumber, normalizedPartNumber, sourcePath, sourcePath
                    If Not CaptureStandaloneImage(rawPartNumber, normalizedPartNumber, sourcePath, imagePath, thumbPath, rowIndex) Then
                        SetRowUtilityValues rowIndex, imagePath, thumbPath, "STANDALONE_CAPTURE_FAILED", "Standalone open/capture failed"
                        AbortWithMessage "Standalone capture failed for Part Number: " & rawPartNumber, "STANDALONE_CAPTURE_FAILED", rowIndex, rawPartNumber, normalizedPartNumber, sourcePath
                        Exit Function
                    End If

                    SetRowUtilityValues rowIndex, imagePath, thumbPath, "OK", ""
                    CacheImage normalizedPartNumber, imagePath, thumbPath
                    If Not InsertThumbnailForRow(rowIndex, thumbPath) Then
                        AbortWithMessage "Excel ne moze da ubaci thumbnail za Part Number: " & rawPartNumber, "EXCEL_THUMBNAIL_INSERTED", rowIndex, rawPartNumber, normalizedPartNumber, sourcePath
                        Exit Function
                    End If
                    gSuccessfulImageRows = gSuccessfulImageRows + 1
                    WriteExportLog rowIndex, rawPartNumber, normalizedPartNumber, "OK", "EXCEL_THUMBNAIL_INSERTED", "Thumbnail inserted.", imagePath, thumbPath, sourcePath
                    MaybeSaveByProgress rowIndex, rawPartNumber, normalizedPartNumber
                End If
            End If
        End If
        DoEvents
    Next

    ProcessBomRowsForThumbnails = True
    Err.Clear
End Function

Sub BuildCatiaFileIndex()
    On Error Resume Next
    TraverseProductForSourceIndex gProduct
    Err.Clear
End Sub

Sub TraverseProductForSourceIndex(prod)
    On Error Resume Next
    Dim rawPn
    Dim normalizedPn
    Dim sourcePath
    rawPn = GetProductPartNumber(prod)
    normalizedPn = NormalizePartNumber(rawPn)
    sourcePath = GetProductSourceFilePath(prod)
    If rawPn <> "" Then WriteDebugPhase "CATIA_FILE_INDEX_ITEM", 0, rawPn, normalizedPn, sourcePath, "Raw ProductTree Part Number=" & rawPn & "; Normalized ProductTree Part Number=" & normalizedPn
    If normalizedPn <> "" And sourcePath <> "" Then
        If Not SamePath(sourcePath, gMainDocumentFullName) Then
            If Not gSourceIndex.Exists(normalizedPn) Then gSourceIndex.Item(normalizedPn) = sourcePath
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
        If (i Mod 250) = 0 Then DoEvents
    Next
    Err.Clear
End Sub

Function CaptureStandaloneImage(rawPartNumber, normalizedPartNumber, sourcePath, imagePath, thumbPath, excelRow)
    On Error Resume Next
    CaptureStandaloneImage = False
    If gFSO.FileExists(imagePath) Then gFSO.DeleteFile imagePath, True
    If gFSO.FileExists(thumbPath) Then gFSO.DeleteFile thumbPath, True

    WriteDebugPhase "STANDALONE_OPEN_START", excelRow, rawPartNumber, normalizedPartNumber, sourcePath, sourcePath
    If Not OpenStandaloneDocument(sourcePath) Then
        WriteExportLog excelRow, rawPartNumber, normalizedPartNumber, "STANDALONE_OPEN_FAILED", "STANDALONE_OPEN_START", "Cannot open source file.", imagePath, thumbPath, sourcePath
        CloseCurrentStandaloneDocument
        ActivateMainDocument
        Exit Function
    End If
    WriteDebugPhase "STANDALONE_OPEN_DONE", excelRow, rawPartNumber, normalizedPartNumber, sourcePath, sourcePath

    WriteDebugPhase "STANDALONE_CAPTURE_START", excelRow, rawPartNumber, normalizedPartNumber, sourcePath, imagePath
    If Not CaptureActiveStandaloneViewer(imagePath) Then
        WriteExportLog excelRow, rawPartNumber, normalizedPartNumber, "STANDALONE_CAPTURE_FAILED", "STANDALONE_CAPTURE_START", "CaptureToFile failed.", imagePath, thumbPath, sourcePath
        CloseCurrentStandaloneDocument
        ActivateMainDocument
        Exit Function
    End If
    WriteDebugPhase "STANDALONE_CAPTURE_DONE", excelRow, rawPartNumber, normalizedPartNumber, sourcePath, imagePath

    If Not CreateThumbnailFile(imagePath, thumbPath, THUMBNAIL_WIDTH, THUMBNAIL_HEIGHT) Then
        WriteExportLog excelRow, rawPartNumber, normalizedPartNumber, "THUMBNAIL_FAILED", "THUMBNAIL_CREATED", "Thumbnail creation failed.", imagePath, thumbPath, sourcePath
        CloseCurrentStandaloneDocument
        ActivateMainDocument
        Exit Function
    End If
    WriteDebugPhase "THUMBNAIL_CREATED", excelRow, rawPartNumber, normalizedPartNumber, sourcePath, thumbPath

    CloseCurrentStandaloneDocument
    ActivateMainDocument
    CaptureStandaloneImage = True
    Err.Clear
End Function

Function OpenStandaloneDocument(sourcePath)
    On Error Resume Next
    OpenStandaloneDocument = False
    Set gCurrentStandaloneDoc = Nothing
    gCurrentStandaloneOpened = False

    If FindOpenDocumentByFullName(sourcePath, gCurrentStandaloneDoc) Then
        gCurrentStandaloneDoc.Activate
        OpenStandaloneDocument = True
        Exit Function
    End If

    CATIA.DisplayFileAlerts = False
    Set gCurrentStandaloneDoc = CATIA.Documents.Open(sourcePath)
    If Err.Number <> 0 Or gCurrentStandaloneDoc Is Nothing Then
        CATIA.DisplayFileAlerts = True
        Err.Clear
        Exit Function
    End If
    gCurrentStandaloneOpened = True
    gCurrentStandaloneDoc.Activate
    OpenStandaloneDocument = True
    Err.Clear
End Function

Sub CloseCurrentStandaloneDocument()
    On Error Resume Next
    If Not gCurrentStandaloneDoc Is Nothing Then
        If CLOSE_STANDALONE_DOCUMENT_AFTER_CAPTURE And gCurrentStandaloneOpened Then
            CATIA.DisplayFileAlerts = False
            gCurrentStandaloneDoc.Close
            CATIA.DisplayFileAlerts = True
            WriteDebugPhase "STANDALONE_CLOSE_DONE", 0, "", "", "", "Standalone document closed without saving."
        End If
    End If
    Set gCurrentStandaloneDoc = Nothing
    gCurrentStandaloneOpened = False
    Err.Clear
End Sub

Function CaptureActiveStandaloneViewer(imagePath)
    On Error Resume Next
    CaptureActiveStandaloneViewer = False
    Dim viewer
    Dim viewpoint
    Dim sight(2)
    Dim up(2)

    CATIA.ActiveWindow.Width = IMAGE_WIDTH
    CATIA.ActiveWindow.Height = IMAGE_HEIGHT

    Set viewer = CATIA.ActiveWindow.ActiveViewer
    viewer.Activate
    viewer.Reframe
    viewer.Update

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

    Set viewpoint = viewer.Viewpoint3D
    sight(0) = 1
    sight(1) = -1
    sight(2) = 1
    up(0) = -0.4082482905
    up(1) = 0.4082482905
    up(2) = 0.8164965809
    viewpoint.PutSightDirection sight
    viewpoint.PutUpDirection up
    If USE_PARALLEL_PROJECTION Then
        viewpoint.ProjectionMode = CAT_PROJECTION_CYLINDRIC
        Err.Clear
    End If

    viewer.Reframe
    viewer.Update
    WaitSeconds IMAGE_CAPTURE_DELAY_SECONDS
    viewer.CaptureToFile CAT_CAPTURE_FORMAT_JPEG, imagePath
    WaitSeconds 0.05
    CaptureActiveStandaloneViewer = gFSO.FileExists(imagePath)
    Err.Clear
End Function

Function CreateThumbnailFile(sourcePath, thumbPath, maxWidth, maxHeight)
    On Error Resume Next
    CreateThumbnailFile = False
    If Not CREATE_THUMBNAIL_FILES Then Exit Function
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

Function InsertThumbnailForRow(rowIndex, thumbPath)
    On Error Resume Next
    InsertThumbnailForRow = False
    If Not INSERT_IMAGES_IN_EXCEL Then
        InsertThumbnailForRow = True
        Exit Function
    End If
    If thumbPath = "" Or Not gFSO.FileExists(thumbPath) Then Exit Function

    Dim cell
    Dim pic
    Set cell = gWsBom.Cells(CLng(rowIndex), gThumbnailColumnIndex)
    DeletePicturesInCell gWsBom, cell
    gWsBom.Rows(CLng(rowIndex)).RowHeight = CLng(THUMBNAIL_HEIGHT * 0.75) + 12
    Set pic = gWsBom.Shapes.AddPicture(thumbPath, MSO_FALSE, MSO_TRUE, cell.Left + 2, cell.Top + 2, -1, -1)
    pic.LockAspectRatio = MSO_TRUE
    If pic.Width > THUMBNAIL_WIDTH Then pic.Width = THUMBNAIL_WIDTH
    If pic.Height > THUMBNAIL_HEIGHT Then pic.Height = THUMBNAIL_HEIGHT
    pic.Left = cell.Left + ((cell.Width - pic.Width) / 2)
    pic.Top = cell.Top + ((cell.Height - pic.Height) / 2)
    WriteDebugPhase "EXCEL_THUMBNAIL_INSERTED", rowIndex, CStr(gWsBom.Cells(rowIndex, gPartNumberColumnIndex).Value), NormalizePartNumber(gWsBom.Cells(rowIndex, gPartNumberColumnIndex).Value), "", thumbPath
    InsertThumbnailForRow = (Err.Number = 0)
    Err.Clear
End Function

Sub DeletePicturesInCell(ws, cell)
    On Error Resume Next
    Dim i
    Dim shp
    For i = ws.Shapes.Count To 1 Step -1
        Set shp = ws.Shapes.Item(i)
        If shp.Top >= cell.Top - 1 And shp.Top < cell.Top + cell.Height And shp.Left >= cell.Left - 1 And shp.Left < cell.Left + cell.Width Then shp.Delete
    Next
    Err.Clear
End Sub

Sub SetRowUtilityValues(rowIndex, imagePath, thumbPath, statusText, skipReason)
    On Error Resume Next
    gWsBom.Cells(rowIndex, gImagePathColumnIndex).Value = imagePath
    gWsBom.Cells(rowIndex, gThumbnailPathColumnIndex).Value = thumbPath
    gWsBom.Cells(rowIndex, gExportStatusColumnIndex).Value = statusText
    gWsBom.Cells(rowIndex, gImageSkipReasonColumnIndex).Value = skipReason
    Err.Clear
End Sub

Sub WriteExportLog(rowIndex, rawPartNumber, normalizedPartNumber, statusText, phase, messageText, imagePath, thumbPath, sourcePath)
    On Error Resume Next
    gWsLog.Cells(gNextLogRow, 1).Value = gNextLogRow - 1
    gWsLog.Cells(gNextLogRow, 2).Value = Now
    gWsLog.Cells(gNextLogRow, 3).Value = rowIndex
    gWsLog.Cells(gNextLogRow, 4).Value = rawPartNumber
    gWsLog.Cells(gNextLogRow, 5).Value = normalizedPartNumber
    gWsLog.Cells(gNextLogRow, 6).Value = statusText
    gWsLog.Cells(gNextLogRow, 7).Value = phase
    gWsLog.Cells(gNextLogRow, 8).Value = messageText
    gWsLog.Cells(gNextLogRow, 9).Value = imagePath
    gWsLog.Cells(gNextLogRow, 10).Value = thumbPath
    gWsLog.Cells(gNextLogRow, 11).Value = sourcePath
    gNextLogRow = gNextLogRow + 1
    If statusText = "SOURCE_FILE_NOT_FOUND" Or statusText = "UNSUPPORTED_SOURCE_FILE" Or statusText = "STANDALONE_OPEN_FAILED" Or statusText = "STANDALONE_CAPTURE_FAILED" Or statusText = "THUMBNAIL_FAILED" Or statusText = "ERROR" Then gErrorCount = gErrorCount + 1
    Err.Clear
End Sub

Sub WriteDebugPhase(phase, rowIndex, rawPartNumber, normalizedPartNumber, sourcePath, messageText)
    On Error Resume Next
    Dim ts
    If gDebugLogPath = "" Then Exit Sub
    Set ts = gFSO.OpenTextFile(gDebugLogPath, 8, True)
    ts.WriteLine FormatDateTime(Now, 2) & " " & FormatDateTime(Now, 3) & _
                 " | " & phase & _
                 " | ExcelRow=" & CStr(rowIndex) & _
                 " | Raw Part Number=" & CStr(rawPartNumber) & _
                 " | Normalized Part Number=" & CStr(normalizedPartNumber) & _
                 " | Source Path=" & CStr(sourcePath) & _
                 " | " & CStr(messageText)
    ts.Close
    Err.Clear
End Sub

Sub SaveExcelCheckpoint(phase, rowIndex, rawPartNumber, normalizedPartNumber)
    On Error Resume Next
    If gWorkbook Is Nothing Then Exit Sub
    UpdateSummarySheet
    gWorkbook.Save
    WriteDebugPhase "SAVE_CHECKPOINT", rowIndex, rawPartNumber, normalizedPartNumber, "", phase
    Err.Clear
End Sub

Sub MaybeSaveByProgress(rowIndex, rawPartNumber, normalizedPartNumber)
    If gSuccessfulImageRows = 1 Then
        SaveExcelCheckpoint "First successful image", rowIndex, rawPartNumber, normalizedPartNumber
    ElseIf SAVE_EVERY_N_ROWS > 0 Then
        If (gProcessedImageRows Mod SAVE_EVERY_N_ROWS) = 0 Then SaveExcelCheckpoint "Periodic save", rowIndex, rawPartNumber, normalizedPartNumber
    End If
End Sub

Sub AbortWithMessage(messageText, phase, rowIndex, rawPartNumber, normalizedPartNumber, sourcePath)
    On Error Resume Next
    WriteDebugPhase "ERROR", rowIndex, rawPartNumber, normalizedPartNumber, sourcePath, phase & ": " & messageText
    WriteExportLog rowIndex, rawPartNumber, normalizedPartNumber, "ERROR", phase, messageText, "", "", sourcePath
    SaveExcelCheckpoint phase, rowIndex, rawPartNumber, normalizedPartNumber
    CleanupCatiaSession
    MsgBox messageText, vbCritical, "CATIA_VISUAL_BOM_EXPORTER"
    Err.Clear
End Sub

Sub CleanupCatiaSession()
    On Error Resume Next
    CloseCurrentStandaloneDocument
    ActivateMainDocument
    If Not gProductDocument Is Nothing Then gProductDocument.Selection.Clear
    CATIA.RefreshDisplay = True
    CATIA.StatusBar = ""
    CATIA.DisplayFileAlerts = True
    If Not gExcelApp Is Nothing Then
        gExcelApp.ScreenUpdating = True
        gExcelApp.DisplayAlerts = True
        gExcelApp.Visible = True
    End If
    Err.Clear
End Sub

Sub ActivateMainDocument()
    On Error Resume Next
    If Not gProductDocument Is Nothing Then gProductDocument.Activate
    Err.Clear
End Sub

Sub UpdateSummarySheet()
    On Error Resume Next
    If gWsSummary Is Nothing Then Exit Sub
    gWsSummary.Cells(3, 2).Value = GetProductPartNumber(gProduct)
    gWsSummary.Cells(4, 2).Value = Now
    gWsSummary.Cells(5, 2).Value = gExcelPath
    gWsSummary.Cells(6, 2).Value = gOutputFolder
    gWsSummary.Cells(7, 2).Value = gLastBomRow - gHeaderRow
    gWsSummary.Cells(8, 2).Value = gProcessedImageRows
    gWsSummary.Cells(9, 2).Value = gSuccessfulImageRows
    gWsSummary.Cells(10, 2).Value = gSkippedFastenerRows
    gWsSummary.Cells(11, 2).Value = gErrorCount
    gWsSummary.Cells(12, 2).Value = "TEST_MODE=" & CStr(TEST_MODE) & "; TEST_MAX_ROWS=" & CStr(TEST_MAX_ROWS) & "; STANDALONE_CAPTURE_ONLY=" & CStr(STANDALONE_CAPTURE_ONLY)
    gWsSummary.Cells(13, 2).Value = gDebugLogPath
    Err.Clear
End Sub

Function SourcePathForPartNumber(partNumber)
    On Error Resume Next
    SourcePathForPartNumber = ""
    Dim normalizedPn
    Dim noRevPn
    Dim candidateKey
    normalizedPn = NormalizePartNumber(partNumber)
    If normalizedPn = "" Then Exit Function
    If gSourceIndex.Exists(normalizedPn) Then
        SourcePathForPartNumber = CStr(gSourceIndex.Item(normalizedPn))
        Exit Function
    End If
    noRevPn = PartNumberWithoutRevision(normalizedPn)
    If noRevPn <> "" And noRevPn <> normalizedPn Then
        If gSourceIndex.Exists(noRevPn) Then
            SourcePathForPartNumber = CStr(gSourceIndex.Item(noRevPn))
            Exit Function
        End If
    End If
    candidateKey = FindUniqueSourceIndexCandidate(normalizedPn)
    If candidateKey = "" And noRevPn <> "" And noRevPn <> normalizedPn Then candidateKey = FindUniqueSourceIndexCandidate(noRevPn)
    If candidateKey <> "" Then SourcePathForPartNumber = CStr(gSourceIndex.Item(candidateKey))
    Err.Clear
End Function

Function ReuseExistingOrCachedImage(normalizedPartNumber, imagePath, thumbPath)
    On Error Resume Next
    ReuseExistingOrCachedImage = False
    Dim payload
    Dim parts
    If gImageCache.Exists(normalizedPartNumber) Then
        payload = CStr(gImageCache.Item(normalizedPartNumber))
        parts = Split(payload, "|")
        If UBound(parts) >= 1 Then
            imagePath = CStr(parts(0))
            thumbPath = CStr(parts(1))
        End If
    End If
    If SKIP_EXISTING_IMAGES And gFSO.FileExists(imagePath) Then
        If Not gFSO.FileExists(thumbPath) Then CreateThumbnailFile imagePath, thumbPath, THUMBNAIL_WIDTH, THUMBNAIL_HEIGHT
        If gFSO.FileExists(thumbPath) Then
            CacheImage normalizedPartNumber, imagePath, thumbPath
            ReuseExistingOrCachedImage = True
        End If
    End If
    Err.Clear
End Function

Sub CacheImage(normalizedPartNumber, imagePath, thumbPath)
    If normalizedPartNumber <> "" Then gImageCache.Item(normalizedPartNumber) = imagePath & "|" & thumbPath
End Sub

Function BuildImagePath(normalizedPartNumber, rowIndex)
    Dim baseName
    baseName = SafeFileName(normalizedPartNumber)
    If baseName = "" Then baseName = "BOM_ROW_" & CStr(rowIndex)
    If Len(baseName) > 140 Then baseName = Left(baseName, 140)
    BuildImagePath = JoinPath(gImageFolder, baseName & "." & FAST_IMAGE_FORMAT)
End Function

Function ThumbnailPathForImage(imagePath)
    ThumbnailPathForImage = JoinPath(gThumbnailFolder, gFSO.GetBaseName(imagePath) & ".jpg")
End Function

Function IsFastenerExcelRow(rowIndex)
    On Error Resume Next
    IsFastenerExcelRow = False
    Dim used
    Dim maxCol
    Dim c
    Dim scanText
    Dim keywords
    Dim kw
    Dim normalizedKw
    Set used = gWsBom.UsedRange
    maxCol = used.Column + used.Columns.Count - 1
    scanText = ""
    For c = 1 To maxCol
        If c <> gThumbnailColumnIndex And c <> gImagePathColumnIndex And c <> gThumbnailPathColumnIndex And c <> gExportStatusColumnIndex And c <> gImageSkipReasonColumnIndex Then
            scanText = scanText & " " & CStr(gWsBom.Cells(rowIndex, c).Value)
        End If
    Next
    scanText = NormalizeFastenerText(scanText)
    keywords = Split(FASTENER_KEYWORDS, "|")
    For Each kw In keywords
        normalizedKw = NormalizeFastenerText(CStr(kw))
        If normalizedKw <> "" Then
            If InStr(1, scanText, normalizedKw, vbTextCompare) > 0 Then
                IsFastenerExcelRow = True
                Exit Function
            End If
        End If
    Next
    Err.Clear
End Function

Function NormalizeFastenerText(valueText)
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

Function GetProductPartNumber(prod)
    On Error Resume Next
    GetProductPartNumber = Trim(CStr(prod.PartNumber))
    If GetProductPartNumber = "" Then
        Dim refProd
        Set refProd = prod.ReferenceProduct
        If Not refProd Is Nothing Then GetProductPartNumber = Trim(CStr(refProd.PartNumber))
    End If
    Err.Clear
End Function

Function GetProductSourceFilePath(prod)
    On Error Resume Next
    Dim refProd
    Dim candidate
    Set refProd = prod.ReferenceProduct
    If Not refProd Is Nothing Then
        candidate = FindSourcePathInObjectChain(refProd.Parent)
        If candidate <> "" Then
            GetProductSourceFilePath = candidate
            Exit Function
        End If
    End If
    candidate = FindSourcePathInObjectChain(prod.Parent)
    If candidate <> "" Then
        GetProductSourceFilePath = candidate
        Exit Function
    End If
    candidate = TryGetMasterShapePath(prod)
    If candidate <> "" Then
        GetProductSourceFilePath = candidate
        Exit Function
    End If
    If Not refProd Is Nothing Then
        candidate = TryGetMasterShapePath(refProd)
        If candidate <> "" Then
            GetProductSourceFilePath = candidate
            Exit Function
        End If
    End If
    candidate = FindSourcePathInObjectChain(prod)
    If candidate <> "" Then GetProductSourceFilePath = candidate
    Err.Clear
End Function

Function TryGetMasterShapePath(obj)
    On Error Resume Next
    TryGetMasterShapePath = CStr(obj.GetMasterShapeRepresentationPathName)
    If Err.Number <> 0 Then
        TryGetMasterShapePath = ""
        Err.Clear
    End If
End Function

Function FindSourcePathInObjectChain(startObj)
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

Function LooksLikeFilePath(pathText)
    LooksLikeFilePath = (CStr(pathText) <> "" And gFSO.GetExtensionName(CStr(pathText)) <> "")
End Function

Function IsSupportedCatiaSourceFile(sourcePath)
    Dim ext
    ext = LCase(gFSO.GetExtensionName(CStr(sourcePath)))
    IsSupportedCatiaSourceFile = (ext = "catpart" Or ext = "catproduct")
End Function

Function FindOpenDocumentByFullName(sourcePath, foundDoc)
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

Function GetDocumentFullName(doc)
    On Error Resume Next
    GetDocumentFullName = CStr(doc.FullName)
    If Err.Number <> 0 Or GetDocumentFullName = "" Then
        Err.Clear
        If CStr(doc.Path) <> "" Then GetDocumentFullName = JoinPath(CStr(doc.Path), CStr(doc.Name))
    End If
    Err.Clear
End Function

Function SamePath(pathA, pathB)
    SamePath = (UCase(Trim(CStr(pathA))) = UCase(Trim(CStr(pathB))) And Trim(CStr(pathA)) <> "")
End Function

Function NormalizePartNumber(value)
    Dim s
    s = CStr(value)
    s = Replace(s, """", "")
    s = Replace(s, "'", "")
    s = Replace(s, vbTab, " ")
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    s = Trim(s)
    Do While Len(s) > 0 And IsTrailingPartNumberSeparator(Right(s, 1))
        s = Left(s, Len(s) - 1)
        s = Trim(s)
    Loop
    Do While InStr(1, s, "  ", vbTextCompare) > 0
        s = Replace(s, "  ", " ")
    Loop
    NormalizePartNumber = UCase(s)
End Function

Function IsTrailingPartNumberSeparator(ch)
    IsTrailingPartNumberSeparator = (ch = "," Or ch = ";" Or ch = ":" Or ch = "|")
End Function

Function PartNumberWithoutRevision(normalizedPartNumber)
    On Error Resume Next
    Dim s
    Dim re
    s = NormalizePartNumber(normalizedPartNumber)
    Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.IgnoreCase = True
    re.Pattern = "([\s_\-\.\/]+)(REV|REVISION|R)[\s_\-\.]*[A-Z0-9]+$"
    If re.Test(s) Then s = re.Replace(s, "")
    PartNumberWithoutRevision = NormalizePartNumber(s)
    Err.Clear
End Function

Function FindUniqueSourceIndexCandidate(normalizedPartNumber)
    On Error Resume Next
    Dim key
    Dim keyNoRev
    Dim wantedNoRev
    Dim matchCount
    Dim matchKey
    normalizedPartNumber = NormalizePartNumber(normalizedPartNumber)
    wantedNoRev = PartNumberWithoutRevision(normalizedPartNumber)
    matchCount = 0
    matchKey = ""
    For Each key In gSourceIndex.Keys
        keyNoRev = PartNumberWithoutRevision(CStr(key))
        If keyNoRev = wantedNoRev Then
            matchCount = matchCount + 1
            matchKey = CStr(key)
        ElseIf IsSafePartialPartNumberMatch(normalizedPartNumber, CStr(key)) Then
            matchCount = matchCount + 1
            matchKey = CStr(key)
        End If
        If matchCount > 1 Then
            matchKey = ""
            Exit For
        End If
    Next
    FindUniqueSourceIndexCandidate = matchKey
    Err.Clear
End Function

Function IsSafePartialPartNumberMatch(a, b)
    a = NormalizePartNumber(a)
    b = NormalizePartNumber(b)
    If a = "" Or b = "" Then Exit Function
    If Len(a) < 6 Or Len(b) < 6 Then Exit Function
    If Len(a) < Len(b) Then
        If Left(b, Len(a)) = a Then IsSafePartialPartNumberMatch = IsPartNumberBoundary(Mid(b, Len(a) + 1, 1))
    ElseIf Len(b) < Len(a) Then
        If Left(a, Len(b)) = b Then IsSafePartialPartNumberMatch = IsPartNumberBoundary(Mid(a, Len(b) + 1, 1))
    End If
End Function

Function IsPartNumberBoundary(ch)
    IsPartNumberBoundary = (ch = "" Or ch = " " Or ch = "_" Or ch = "-" Or ch = "." Or ch = "/" Or ch = "," Or ch = ";" Or ch = ":" Or ch = "|")
End Function

Function NormalizeHeaderName(valueText)
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

Function LastUsedRow(ws)
    On Error Resume Next
    Dim used
    Set used = ws.UsedRange
    LastUsedRow = used.Row + used.Rows.Count - 1
    Err.Clear
End Function

Function LastUsedColumn(ws)
    On Error Resume Next
    Dim used
    Set used = ws.UsedRange
    LastUsedColumn = used.Column + used.Columns.Count - 1
    Err.Clear
End Function

Sub ApplyBomBorders()
    On Error Resume Next
    Dim used
    Set used = gWsBom.UsedRange
    used.Borders.LineStyle = XL_CONTINUOUS
    used.Borders.Weight = XL_THIN
    used.VerticalAlignment = XL_TOP
    Err.Clear
End Sub

Function JoinPath(folderPath, fileName)
    If Right(CStr(folderPath), 1) = "\" Then
        JoinPath = CStr(folderPath) & CStr(fileName)
    Else
        JoinPath = CStr(folderPath) & "\" & CStr(fileName)
    End If
End Function

Sub EnsureFolder(folderPath)
    If Not gFSO.FolderExists(folderPath) Then gFSO.CreateFolder folderPath
End Sub

Function SafeFileName(valueText)
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

Function TimestampForFile()
    TimestampForFile = CStr(Year(Now)) & Pad2(Month(Now)) & Pad2(Day(Now)) & "_" & Pad2(Hour(Now)) & Pad2(Minute(Now)) & Pad2(Second(Now))
End Function

Function Pad2(n)
    If CLng(n) < 10 Then
        Pad2 = "0" & CStr(n)
    Else
        Pad2 = CStr(n)
    End If
End Function

Sub WaitSeconds(secondsValue)
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
