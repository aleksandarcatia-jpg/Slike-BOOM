' ================================================================
' CATIA_VISUAL_BOM_EXPORTER
' Two-phase model BOM Excel thumbnail workflow.
' ================================================================

' ---------------------- CONFIGURATION ---------------------------
Const WORK_ROOT = "C:\Temp\CATIA_VISUAL_BOM_WORK"
Const STATE_FILE_NAME = "VISUAL_BOM_STATE.txt"
Const SAVE_EVERY_N_ROWS = 25
Const SKIP_EXISTING_IMAGES = True
Const RESUME_MODE = True
Const TEST_MODE = True
Const TEST_MAX_ROWS = 20
Const SKIP_FASTENER_IMAGES = True
Const FASTENER_KEYWORDS = "vijak|vijci|zavrtanj|zavrtnji|screw|bolt|hex bolt|hexagon bolt|imbus|allen screw|navrtka|matica|nut|hex nut|podloska|washer|plain washer|spring washer|lock washer|DIN 125|DIN125|DIN 127|DIN127|DIN 933|DIN933|DIN 931|DIN931|DIN 934|DIN934|ISO 4014|ISO 4017|ISO 4032|ISO 7089|ISO 7090"
Const IMAGE_WIDTH = 1000
Const IMAGE_HEIGHT = 750
Const THUMBNAIL_WIDTH = 160
Const THUMBNAIL_HEIGHT = 120
Const IMAGE_CAPTURE_DELAY_SECONDS = 0.2
Const STANDALONE_CAPTURE_ONLY = True
Const FALLBACK_TO_ASSEMBLY_HIDE_SHOW = False
Const CLOSE_STANDALONE_DOCUMENT_AFTER_CAPTURE = True
Const NEVER_SAVE_CATIA_DOCUMENTS = True

' CATIA constants used late-bound.
Const CAT_CAPTURE_FORMAT_JPEG = 2
Const CAT_RENDER_SHADING = 0
Const CAT_RENDER_SHADING_WITH_EDGES = 1
Const CAT_PROJECTION_CYLINDRIC = 0

' Excel constants used late-bound.
Const XL_EXCEL8 = 56
Const XL_UP = -4162
Const XL_TO_LEFT = -4159
Const XL_TOP = -4160
Const XL_CONTINUOUS = 1
Const XL_THIN = 2
Const MSO_TRUE = -1
Const MSO_FALSE = 0

Dim gFSO
Dim gShell
Dim gProductDocument
Dim gProduct
Dim gRootPartNumber
Dim gMainDocumentFullName
Dim gStateFilePath
Dim gWorkbookFolder
Dim gImageFolder
Dim gThumbnailFolder
Dim gDebugLogPath
Dim gSelectedBomExcelPath
Dim gWorkExcelPath
Dim gFinalExcelPath
Dim gFinalFolder
Dim gExcelApp
Dim gWorkbook
Dim gWsBom
Dim gHeaderRow
Dim gPartNumberColumnIndex
Dim gImageColumnIndex
Dim gImagePathColumnIndex
Dim gExportStatusColumnIndex
Dim gImageSkipReasonColumnIndex
Dim gLastBomRow
Dim gSourceIndex
Dim gNeededPartNumbers
Dim gFoundNeededPartNumbers
Dim gImageCache
Dim gCurrentStandaloneDoc
Dim gCurrentStandaloneOpened
Dim gProcessedImageRows
Dim gSuccessfulImageRows
Dim gReusedImageRows
Dim gSkippedFastenerRows
Dim gRowsWithoutPartNumber
Dim gSourceNotFoundRows
Dim gImageCaptureFailedRows
Dim gErrorCount
Dim gAbortAlreadyHandled
Dim gPhase2Success
Dim gPendingDebugText

Public Sub CATIA_VISUAL_BOM_EXPORTER()
    On Error Resume Next

    InitializeRuntime
    If Not ValidateActiveProductDocument() Then Exit Sub
    PrepareWorkFolders

    WriteDebugPhase "START", 0, "", "", "", "CATIA_VISUAL_BOM_EXPORTER started."

    If IsWaitingForBomExcelState() Then
        RunPhase2
    Else
        RunPhase1
    End If

    CleanupCatiaSession
    Err.Clear
End Sub

Sub InitializeRuntime()
    Set gFSO = CreateObject("Scripting.FileSystemObject")
    Set gShell = CreateObject("WScript.Shell")
    Set gSourceIndex = CreateObject("Scripting.Dictionary")
    Set gNeededPartNumbers = CreateObject("Scripting.Dictionary")
    Set gFoundNeededPartNumbers = CreateObject("Scripting.Dictionary")
    Set gImageCache = CreateObject("Scripting.Dictionary")
    gSourceIndex.CompareMode = 1
    gNeededPartNumbers.CompareMode = 1
    gFoundNeededPartNumbers.CompareMode = 1
    gImageCache.CompareMode = 1
    Set gExcelApp = Nothing
    Set gWorkbook = Nothing
    Set gWsBom = Nothing
    Set gCurrentStandaloneDoc = Nothing
    gCurrentStandaloneOpened = False
    gProcessedImageRows = 0
    gSuccessfulImageRows = 0
    gReusedImageRows = 0
    gSkippedFastenerRows = 0
    gRowsWithoutPartNumber = 0
    gSourceNotFoundRows = 0
    gImageCaptureFailedRows = 0
    gErrorCount = 0
    gAbortAlreadyHandled = False
    gPhase2Success = False
    gPendingDebugText = ""
    Randomize
End Sub

Function ValidateActiveProductDocument()
    On Error Resume Next
    ValidateActiveProductDocument = False

    If CATIA.Documents.Count = 0 Then
        MsgBox "Nema otvorenog CATIA dokumenta. Otvorite glavni CATProduct i pokrenite makro.", vbExclamation, "CATIA_VISUAL_BOM_EXPORTER"
        Exit Function
    End If

    Set gProductDocument = CATIA.ActiveDocument
    Set gProduct = gProductDocument.Product
    If Err.Number <> 0 Or gProduct Is Nothing Then
        Err.Clear
        MsgBox "Aktivni dokument mora biti CATProduct.", vbExclamation, "CATIA_VISUAL_BOM_EXPORTER"
        Exit Function
    End If

    gRootPartNumber = GetProductPartNumber(gProduct)
    If gRootPartNumber = "" Then gRootPartNumber = SafeFileName(gProduct.Name)
    If gRootPartNumber = "" Then gRootPartNumber = "CATIA_PRODUCT"
    gMainDocumentFullName = GetDocumentFullName(gProductDocument)
    ValidateActiveProductDocument = True
    Err.Clear
End Function

Sub PrepareWorkFolders()
    On Error Resume Next
    EnsureFolder "C:\Temp"
    EnsureFolder WORK_ROOT
    gWorkbookFolder = JoinPath(WORK_ROOT, "WORKBOOK")
    gImageFolder = JoinPath(WORK_ROOT, "IMAGES")
    gThumbnailFolder = JoinPath(WORK_ROOT, "THUMBNAILS")
    gDebugLogPath = JoinPath(WORK_ROOT, "DEBUG_PHASE_LOG.txt")
    gStateFilePath = JoinPath(WORK_ROOT, STATE_FILE_NAME)
    EnsureFolder gWorkbookFolder
    EnsureFolder gImageFolder
    EnsureFolder gThumbnailFolder
    Err.Clear
End Sub

Function IsWaitingForBomExcelState()
    On Error Resume Next
    IsWaitingForBomExcelState = False
    If Not gFSO.FileExists(gStateFilePath) Then Exit Function
    IsWaitingForBomExcelState = (UCase(ReadStateValue("status")) = "WAITING_FOR_BOM_EXCEL")
    Err.Clear
End Function

Sub RunPhase1()
    On Error Resume Next

    WriteStateFile
    WriteDebugPhase "PHASE1_PREPARED", 0, "", "", gStateFilePath, "Phase 1 prepared."

    MsgBox "PHASE 1 je pripremljena." & vbCrLf & vbCrLf & _
           "Sada uradite:" & vbCrLf & _
           "1. Analyze -> Bill of Material" & vbCrLf & _
           "2. Define Bill of Material / Define formats" & vbCrLf & _
           "3. Podesite BOM kako zelite" & vbCrLf & _
           "4. Save as Excel" & vbCrLf & _
           "5. Ponovo pokrenite ovaj isti makro da doda slike.", _
           vbInformation, "CATIA_VISUAL_BOM_EXPORTER"

    Err.Clear
    CATIA.StartCommand "Bill of Material"
    If Err.Number = 0 Then
        WriteDebugPhase "BILL_OF_MATERIAL_COMMAND_STARTED_OR_SKIPPED", 0, "", "", "", "CATIA command started."
    Else
        WriteDebugPhase "BILL_OF_MATERIAL_COMMAND_STARTED_OR_SKIPPED", 0, "", "", "", "CATIA command could not be started; user must open it manually. Err.Number=" & CStr(Err.Number) & "; Err.Description=" & Err.Description
        Err.Clear
    End If

    CATIA.StatusBar = ""
    Err.Clear
End Sub

Sub RunPhase2()
    On Error Resume Next
    WriteDebugPhase "PHASE2_STARTED", 0, "", "", gStateFilePath, "Phase 2 started."

    gSelectedBomExcelPath = AskUserForExistingBomExcel()
    If gSelectedBomExcelPath = "" Then
        MsgBox "Export je otkazan od strane korisnika.", vbInformation, "CATIA_VISUAL_BOM_EXPORTER"
        WriteDebugPhase "ERROR", 0, "", "", "", "User cancelled existing BOM Excel selection."
        Exit Sub
    End If
    WriteDebugPhase "EXISTING_BOM_EXCEL_SELECTED", 0, "", "", gSelectedBomExcelPath, gSelectedBomExcelPath

    gFinalFolder = AskUserForOutputFolder(GetDefaultOutputFolder())
    If gFinalFolder = "" Then
        MsgBox "Export je otkazan od strane korisnika.", vbInformation, "CATIA_VISUAL_BOM_EXPORTER"
        WriteDebugPhase "ERROR", 0, "", "", "", "User cancelled final folder selection."
        Exit Sub
    End If
    WriteDebugPhase "FINAL_FOLDER_SELECTED", 0, "", "", gFinalFolder, gFinalFolder

    gFinalExcelPath = BuildFinalExcelPath(gFinalFolder, gSelectedBomExcelPath)
    gFinalExcelPath = ResolveExistingOutputFile(gFinalExcelPath)
    If gFinalExcelPath = "" Then
        MsgBox "Export je otkazan od strane korisnika.", vbInformation, "CATIA_VISUAL_BOM_EXPORTER"
        Exit Sub
    End If

    If Not CreateLocalWorkbookCopy(gSelectedBomExcelPath) Then
        AbortPhase2 "Ne mogu da napravim lokalnu kopiju BOM Excela. Proverite da li je fajl otvoren ili zakljucan.", "LOCAL_WORKBOOK_CREATED", 0, "", "", ""
        Exit Sub
    End If

    If Not OpenLocalWorkbookAndPrepareColumns() Then
        If Not gAbortAlreadyHandled Then AbortPhase2 "Ne mogu da pripremim Excel BOM za slike.", "EXCEL_OPENED", 0, "", "", ""
        Exit Sub
    End If

    BuildNeededPartNumbersFromBomRows
    BuildCatiaFileIndex
    ProcessBomRowsForThumbnails

    If Not gAbortAlreadyHandled Then
        SaveExcelCheckpoint "Before final save", 0, "", ""
        If FinalSaveWorkbook() Then
            gPhase2Success = True
            ClearStateFile
            WriteDebugPhase "FINISH", 0, "", "", gFinalExcelPath, "Phase 2 finished."
            MsgBox "Gotovo. Finalni Excel je snimljen:" & vbCrLf & vbCrLf & gFinalExcelPath, vbInformation, "CATIA_VISUAL_BOM_EXPORTER"
        End If
    End If
    Err.Clear
End Sub

Sub WriteStateFile()
    On Error Resume Next
    Dim ts
    Set ts = gFSO.OpenTextFile(gStateFilePath, 2, True)
    ts.WriteLine "status=WAITING_FOR_BOM_EXCEL"
    ts.WriteLine "root_full_path=" & gMainDocumentFullName
    ts.WriteLine "root_part_number=" & gRootPartNumber
    ts.WriteLine "phase1_datetime=" & CStr(Now)
    ts.WriteLine "work_folder=" & WORK_ROOT
    ts.Close
    WriteDebugPhase "STATE_FILE_WRITTEN", 0, "", "", gStateFilePath, "State file written."
    Err.Clear
End Sub

Function ReadStateValue(keyName)
    On Error Resume Next
    ReadStateValue = ""
    If Not gFSO.FileExists(gStateFilePath) Then Exit Function
    Dim ts
    Dim lineText
    Dim p
    Set ts = gFSO.OpenTextFile(gStateFilePath, 1, False)
    Do Until ts.AtEndOfStream
        lineText = ts.ReadLine
        p = InStr(1, lineText, "=", vbTextCompare)
        If p > 0 Then
            If LCase(Left(lineText, p - 1)) = LCase(keyName) Then
                ReadStateValue = Mid(lineText, p + 1)
                Exit Do
            End If
        End If
    Loop
    ts.Close
    Err.Clear
End Function

Sub ClearStateFile()
    On Error Resume Next
    If gFSO.FileExists(gStateFilePath) Then gFSO.DeleteFile gStateFilePath, True
    WriteDebugPhase "STATE_FILE_CLEARED", 0, "", "", gStateFilePath, "State file cleared."
    Err.Clear
End Sub

Function AskUserForExistingBomExcel()
    On Error Resume Next
    AskUserForExistingBomExcel = ""
    Dim selectedPath
    selectedPath = AskUserForExistingBomExcelWithExcelDialog()
    If selectedPath = "__CANCEL__" Then Exit Function
    If selectedPath <> "" Then
        AskUserForExistingBomExcel = selectedPath
        Exit Function
    End If
    selectedPath = AskUserForExistingBomExcelWithCatiaDialog()
    If selectedPath = "__CANCEL__" Then Exit Function
    If selectedPath <> "" Then
        AskUserForExistingBomExcel = selectedPath
        Exit Function
    End If
    AskUserForExistingBomExcel = AskUserForExistingBomExcelInputBox()
    Err.Clear
End Function

Function AskUserForExistingBomExcelWithExcelDialog()
    On Error Resume Next
    AskUserForExistingBomExcelWithExcelDialog = ""
    Dim xl
    Dim fd
    Set xl = CreateObject("Excel.Application")
    If Err.Number <> 0 Or xl Is Nothing Then
        Err.Clear
        Exit Function
    End If
    xl.Visible = False
    Set fd = xl.FileDialog(3)
    If Err.Number <> 0 Or fd Is Nothing Then
        xl.Quit
        Set xl = Nothing
        Err.Clear
        Exit Function
    End If
    fd.Title = "Izaberite CATIA BOM Excel koji ste rucno snimili"
    fd.AllowMultiSelect = False
    fd.Filters.Clear
    fd.Filters.Add "Excel files", "*.xls;*.xlsx;*.xlsm"
    If fd.Show = -1 Then
        AskUserForExistingBomExcelWithExcelDialog = CStr(fd.SelectedItems(1))
    Else
        AskUserForExistingBomExcelWithExcelDialog = "__CANCEL__"
    End If
    xl.Quit
    Set xl = Nothing
    Err.Clear
End Function

Function AskUserForExistingBomExcelWithCatiaDialog()
    On Error Resume Next
    AskUserForExistingBomExcelWithCatiaDialog = ""
    AskUserForExistingBomExcelWithCatiaDialog = CStr(CATIA.FileSelectionBox("Izaberite CATIA BOM Excel koji ste rucno snimili", "*.xls;*.xlsx;*.xlsm", 0))
    If Err.Number <> 0 Then
        AskUserForExistingBomExcelWithCatiaDialog = ""
        Err.Clear
    ElseIf Trim(AskUserForExistingBomExcelWithCatiaDialog) = "" Then
        AskUserForExistingBomExcelWithCatiaDialog = "__CANCEL__"
    End If
End Function

Function AskUserForExistingBomExcelInputBox()
    On Error Resume Next
    Dim p
    p = InputBox("Unesite punu putanju do CATIA BOM Excel fajla (*.xls, *.xlsx, *.xlsm):", "CATIA_VISUAL_BOM_EXPORTER")
    p = Trim(CStr(p))
    If p <> "" And gFSO.FileExists(p) Then
        AskUserForExistingBomExcelInputBox = p
    Else
        AskUserForExistingBomExcelInputBox = ""
    End If
    Err.Clear
End Function

Function AskUserForOutputFolder(defaultFolder)
    On Error Resume Next
    AskUserForOutputFolder = ""
    Dim selectedFolder
    selectedFolder = AskUserForOutputFolderExcelDialog(defaultFolder)
    If selectedFolder = "__CANCEL__" Then Exit Function
    If selectedFolder <> "" Then
        AskUserForOutputFolder = selectedFolder
        Exit Function
    End If
    selectedFolder = AskUserForOutputFolderShell(defaultFolder)
    If selectedFolder = "__CANCEL__" Then Exit Function
    If selectedFolder <> "" Then
        AskUserForOutputFolder = selectedFolder
        Exit Function
    End If
    AskUserForOutputFolder = AskUserForOutputFolderInputBox(defaultFolder)
    Err.Clear
End Function

Function AskUserForOutputFolderExcelDialog(defaultFolder)
    On Error Resume Next
    AskUserForOutputFolderExcelDialog = ""
    Dim xl
    Dim fd
    Set xl = CreateObject("Excel.Application")
    If Err.Number <> 0 Or xl Is Nothing Then
        Err.Clear
        Exit Function
    End If
    xl.Visible = False
    Set fd = xl.FileDialog(4)
    If Err.Number <> 0 Or fd Is Nothing Then
        xl.Quit
        Set xl = Nothing
        Err.Clear
        Exit Function
    End If
    fd.Title = "Izaberite folder za finalni Excel sa slikama"
    fd.AllowMultiSelect = False
    If CStr(defaultFolder) <> "" Then fd.InitialFileName = AddTrailingSlash(defaultFolder)
    If fd.Show = -1 Then
        AskUserForOutputFolderExcelDialog = CStr(fd.SelectedItems(1))
    Else
        AskUserForOutputFolderExcelDialog = "__CANCEL__"
    End If
    xl.Quit
    Set xl = Nothing
    Err.Clear
End Function

Function AskUserForOutputFolderShell(defaultFolder)
    On Error Resume Next
    AskUserForOutputFolderShell = ""
    Dim shellApp
    Dim folder
    Set shellApp = CreateObject("Shell.Application")
    If Err.Number <> 0 Or shellApp Is Nothing Then
        Err.Clear
        Exit Function
    End If
    Set folder = shellApp.BrowseForFolder(0, "Izaberite folder za finalni Excel sa slikama", 0, defaultFolder)
    If folder Is Nothing Then
        AskUserForOutputFolderShell = "__CANCEL__"
        Exit Function
    End If
    AskUserForOutputFolderShell = CStr(folder.Self.Path)
    Err.Clear
End Function

Function AskUserForOutputFolderInputBox(defaultFolder)
    On Error Resume Next
    Dim p
    p = InputBox("Unesite/potvrdite folder za finalni Excel sa slikama:", "CATIA_VISUAL_BOM_EXPORTER", defaultFolder)
    p = Trim(CStr(p))
    If p <> "" And gFSO.FolderExists(p) Then
        AskUserForOutputFolderInputBox = p
    Else
        AskUserForOutputFolderInputBox = ""
    End If
    Err.Clear
End Function

Function BuildFinalExcelPath(finalFolder, originalExcelPath)
    Dim baseName
    baseName = SafeFileName(gFSO.GetBaseName(originalExcelPath))
    If baseName = "" Then baseName = "BOM"
    BuildFinalExcelPath = JoinPath(finalFolder, baseName & "_WITH_IMAGES.xls")
End Function

Function ResolveExistingOutputFile(xlsPath)
    On Error Resume Next
    ResolveExistingOutputFile = xlsPath
    If Not gFSO.FileExists(xlsPath) Then Exit Function

    Dim answer
    answer = MsgBox("Fajl vec postoji:" & vbCrLf & xlsPath & vbCrLf & vbCrLf & _
                    "YES = zameni postojeci fajl" & vbCrLf & _
                    "NO = napravi novi fajl sa timestamp nastavkom" & vbCrLf & _
                    "CANCEL = prekini export", _
                    vbYesNoCancel + vbQuestion, "CATIA_VISUAL_BOM_EXPORTER")
    If answer = vbYes Then
        gFSO.DeleteFile xlsPath, True
        If Err.Number <> 0 Then
            MsgBox "Ne mogu da obrisem postojeci Excel. Zatvorite fajl ako je otvoren:" & vbCrLf & xlsPath, vbCritical, "CATIA_VISUAL_BOM_EXPORTER"
            Err.Clear
            ResolveExistingOutputFile = ""
        End If
    ElseIf answer = vbNo Then
        ResolveExistingOutputFile = Replace(xlsPath, ".xls", "_" & TimestampForFile() & ".xls")
    Else
        ResolveExistingOutputFile = ""
    End If
    Err.Clear
End Function

Function CreateLocalWorkbookCopy(sourceExcelPath)
    On Error Resume Next
    CreateLocalWorkbookCopy = False
    EnsureFolder gWorkbookFolder
    gWorkExcelPath = JoinPath(gWorkbookFolder, SafeFileName(gFSO.GetBaseName(sourceExcelPath)) & "_WORK.xls")
    If gFSO.FileExists(gWorkExcelPath) Then
        Err.Clear
        gFSO.DeleteFile gWorkExcelPath, True
        If Err.Number <> 0 Then
            Err.Clear
            Exit Function
        End If
    End If

    If LCase(gFSO.GetExtensionName(sourceExcelPath)) = "xls" Then
        gFSO.CopyFile sourceExcelPath, gWorkExcelPath, True
    Else
        If Not EnsureExcelApp() Then Exit Function
        Dim tempWb
        gExcelApp.DisplayAlerts = False
        Set tempWb = gExcelApp.Workbooks.Open(sourceExcelPath)
        If Err.Number <> 0 Or tempWb Is Nothing Then
            Err.Clear
            Exit Function
        End If
        tempWb.SaveAs gWorkExcelPath, XL_EXCEL8
        tempWb.Close False
        gExcelApp.DisplayAlerts = True
    End If

    If gFSO.FileExists(gWorkExcelPath) Then
        WriteDebugPhase "LOCAL_WORKBOOK_CREATED", 0, "", "", gWorkExcelPath, "Local workbook copy created."
        CreateLocalWorkbookCopy = True
    End If
    Err.Clear
End Function

Function OpenLocalWorkbookAndPrepareColumns()
    On Error Resume Next
    OpenLocalWorkbookAndPrepareColumns = False
    If Not EnsureExcelApp() Then
        MsgBox "Microsoft Excel nije dostupan.", vbCritical, "CATIA_VISUAL_BOM_EXPORTER"
        Exit Function
    End If

    gExcelApp.Visible = True
    gExcelApp.DisplayAlerts = False
    gExcelApp.ScreenUpdating = False
    Set gWorkbook = gExcelApp.Workbooks.Open(gWorkExcelPath)
    If Err.Number <> 0 Or gWorkbook Is Nothing Then
        Err.Clear
        Exit Function
    End If

    Set gWsBom = gWorkbook.Worksheets(1)
    WriteDebugPhase "EXCEL_OPENED", 0, "", "", gWorkExcelPath, "Local workbook opened."

    If Not FindBomHeaderRowAndPartNumberColumn() Then
        SaveExcelCheckpoint "Part Number column missing", 0, "", ""
        MsgBox "Kolona Part Number / PN / Oznaka nije pronadjena u BOM Excel fajlu. Slikanje je prekinuto.", vbCritical, "CATIA_VISUAL_BOM_EXPORTER"
        Exit Function
    End If
    WriteDebugPhase "PART_NUMBER_COLUMN_FOUND", gHeaderRow, "", "", "", "Column=" & CStr(gPartNumberColumnIndex)

    EnsureImageAndHelperColumns
    gLastBomRow = LastUsedRow(gWsBom)
    SaveExcelCheckpoint "EXCEL_OPENED", 0, "", ""
    OpenLocalWorkbookAndPrepareColumns = True
    Err.Clear
End Function

Function EnsureExcelApp()
    On Error Resume Next
    EnsureExcelApp = False
    If Not gExcelApp Is Nothing Then
        EnsureExcelApp = True
        Exit Function
    End If
    Set gExcelApp = CreateObject("Excel.Application")
    If Err.Number = 0 And Not gExcelApp Is Nothing Then EnsureExcelApp = True
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
    IsPartNumberHeader = (h = "partnumber" Or h = "partno" Or h = "pn" Or h = "number" Or h = "brojdela" Or h = "brdela" Or h = "oznaka")
End Function

Function IsImageHeader(valueText)
    Dim h
    h = NormalizeHeaderName(valueText)
    IsImageHeader = (h = "slika" Or h = "image" Or h = "thumbnail" Or h = "picture" Or h = "foto" Or h = "preview")
End Function

Sub EnsureImageAndHelperColumns()
    On Error Resume Next
    gImageColumnIndex = FindImageColumn()
    If gImageColumnIndex = 0 Then
        gWsBom.Columns(gPartNumberColumnIndex + 1).Insert
        gWsBom.Cells(gHeaderRow, gPartNumberColumnIndex + 1).Value = "Thumbnail"
        gImageColumnIndex = gPartNumberColumnIndex + 1
    End If
    gImagePathColumnIndex = EnsureEndHelperColumn("Image Path")
    gExportStatusColumnIndex = EnsureEndHelperColumn("Export Status")
    gImageSkipReasonColumnIndex = EnsureEndHelperColumn("Image Skip Reason")

    gWsBom.Columns(gImageColumnIndex).ColumnWidth = 20
    gWsBom.Columns(gImagePathColumnIndex).ColumnWidth = 45
    gWsBom.Columns(gExportStatusColumnIndex).ColumnWidth = 24
    gWsBom.Columns(gImageSkipReasonColumnIndex).ColumnWidth = 45
    gWsBom.Rows(gHeaderRow).Font.Bold = True
    ApplyBomBorders
    WriteDebugPhase "IMAGE_COLUMN_FOUND_OR_CREATED", gHeaderRow, "", "", "", "Image column=" & CStr(gImageColumnIndex)
    Err.Clear
End Sub

Function FindImageColumn()
    On Error Resume Next
    FindImageColumn = 0
    Dim used
    Dim maxCol
    Dim c
    Set used = gWsBom.UsedRange
    maxCol = used.Column + used.Columns.Count - 1
    For c = 1 To maxCol
        If IsImageHeader(CStr(gWsBom.Cells(gHeaderRow, c).Value)) Then
            FindImageColumn = c
            Exit Function
        End If
    Next
    Err.Clear
End Function

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
    FindHeaderColumn = 0
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

Sub BuildNeededPartNumbersFromBomRows()
    On Error Resume Next
    Set gNeededPartNumbers = CreateObject("Scripting.Dictionary")
    Set gFoundNeededPartNumbers = CreateObject("Scripting.Dictionary")
    gNeededPartNumbers.CompareMode = 1
    gFoundNeededPartNumbers.CompareMode = 1

    Dim rowIndex
    Dim rawPartNumber
    Dim normalizedPartNumber
    Dim candidateRows
    candidateRows = 0

    For rowIndex = gHeaderRow + 1 To gLastBomRow
        rawPartNumber = CStr(gWsBom.Cells(rowIndex, gPartNumberColumnIndex).Value)
        normalizedPartNumber = NormalizePartNumber(rawPartNumber)
        If normalizedPartNumber <> "" Then
            If Not (SKIP_FASTENER_IMAGES And IsFastenerExcelRow(rowIndex)) Then
                candidateRows = candidateRows + 1
                If Not (TEST_MODE And candidateRows > TEST_MAX_ROWS) Then
                    If Not ExistingImageAvailable(normalizedPartNumber, rowIndex) Then
                        If Not gNeededPartNumbers.Exists(normalizedPartNumber) Then gNeededPartNumbers.Item(normalizedPartNumber) = True
                    End If
                End If
            End If
        End If
        If (rowIndex Mod 250) = 0 Then DoEvents
    Next

    WriteDebugPhase "NEEDED_PART_NUMBERS_BUILT", 0, "", "", "", "Unique needed Part Numbers=" & CStr(gNeededPartNumbers.Count)
    Err.Clear
End Sub

Sub BuildCatiaFileIndex()
    On Error Resume Next
    WriteDebugPhase "CATIA_FILE_INDEX_START", 0, "", "", "", "Indexing source files for needed Part Numbers."
    CATIA.StatusBar = "Indexing source files: found 0 / " & CStr(gNeededPartNumbers.Count)
    If gNeededPartNumbers.Count > 0 Then TraverseProductForSourceIndex gProduct
    WriteDebugPhase "CATIA_FILE_INDEX_DONE", 0, "", "", "", "Found=" & CStr(gFoundNeededPartNumbers.Count) & " / Needed=" & CStr(gNeededPartNumbers.Count)
    Err.Clear
End Sub

Sub TraverseProductForSourceIndex(prod)
    On Error Resume Next
    If AllNeededPartNumbersFound() Then Exit Sub

    Dim rawPn
    Dim normalizedPn
    Dim sourcePath
    rawPn = GetProductPartNumber(prod)
    normalizedPn = NormalizePartNumber(rawPn)
    If ShouldIndexPartNumber(normalizedPn) Then
        sourcePath = GetProductSourceFilePath(prod)
        If normalizedPn <> "" And sourcePath <> "" Then
            If Not SamePath(sourcePath, gMainDocumentFullName) Then
                If Not gSourceIndex.Exists(normalizedPn) Then gSourceIndex.Item(normalizedPn) = sourcePath
                MarkNeededPartNumberFound normalizedPn, sourcePath
            End If
        End If
    End If

    If AllNeededPartNumbersFound() Then Exit Sub
    Dim children
    Dim i
    Set children = prod.Products
    If Err.Number <> 0 Or children Is Nothing Then
        Err.Clear
        Exit Sub
    End If
    For i = 1 To children.Count
        TraverseProductForSourceIndex children.Item(i)
        If AllNeededPartNumbersFound() Then Exit For
        If (i Mod 250) = 0 Then DoEvents
    Next
    Err.Clear
End Sub

Function ShouldIndexPartNumber(normalizedPn)
    normalizedPn = NormalizePartNumber(normalizedPn)
    ShouldIndexPartNumber = (MatchedNeededPartNumberKey(normalizedPn) <> "")
End Function

Sub MarkNeededPartNumberFound(normalizedPn, sourcePath)
    On Error Resume Next
    Dim matchedKey
    matchedKey = MatchedNeededPartNumberKey(normalizedPn)
    If matchedKey <> "" Then
        If Not gSourceIndex.Exists(matchedKey) Then gSourceIndex.Item(matchedKey) = sourcePath
        If Not gFoundNeededPartNumbers.Exists(matchedKey) Then gFoundNeededPartNumbers.Item(matchedKey) = True
        CATIA.StatusBar = "Indexing source files: found " & CStr(gFoundNeededPartNumbers.Count) & " / " & CStr(gNeededPartNumbers.Count)
    End If
    Err.Clear
End Sub

Function MatchedNeededPartNumberKey(normalizedPn)
    On Error Resume Next
    MatchedNeededPartNumberKey = ""
    Dim key
    Dim pnNoRev
    Dim keyNoRev
    normalizedPn = NormalizePartNumber(normalizedPn)
    If normalizedPn = "" Then Exit Function
    If gNeededPartNumbers.Exists(normalizedPn) Then
        MatchedNeededPartNumberKey = normalizedPn
        Exit Function
    End If
    pnNoRev = PartNumberWithoutRevision(normalizedPn)
    For Each key In gNeededPartNumbers.Keys
        keyNoRev = PartNumberWithoutRevision(CStr(key))
        If keyNoRev <> "" And keyNoRev = pnNoRev Then
            MatchedNeededPartNumberKey = CStr(key)
            Exit Function
        End If
        If IsSafePartialPartNumberMatch(CStr(key), normalizedPn) Then
            MatchedNeededPartNumberKey = CStr(key)
            Exit Function
        End If
    Next
    Err.Clear
End Function

Function AllNeededPartNumbersFound()
    On Error Resume Next
    AllNeededPartNumbersFound = False
    If gNeededPartNumbers.Count = 0 Then Exit Function
    AllNeededPartNumbersFound = (gFoundNeededPartNumbers.Count >= gNeededPartNumbers.Count)
    Err.Clear
End Function

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
    Dim candidateRows
    candidateRows = 0

    For rowIndex = gHeaderRow + 1 To gLastBomRow
        rawPartNumber = CStr(gWsBom.Cells(rowIndex, gPartNumberColumnIndex).Value)
        normalizedPartNumber = NormalizePartNumber(rawPartNumber)
        CATIA.StatusBar = "Processing BOM images: row " & CStr(rowIndex - gHeaderRow) & " / " & CStr(gLastBomRow - gHeaderRow)
        WriteDebugPhase "ROW_START", rowIndex, rawPartNumber, normalizedPartNumber, "", "Row start."

        If normalizedPartNumber = "" Then
            If RowHasAnyData(rowIndex) Then
                gRowsWithoutPartNumber = gRowsWithoutPartNumber + 1
                SetRowUtilityValues rowIndex, "", "NO_PART_NUMBER", "No Part Number in BOM row"
                WriteDebugPhase "ROW_START", rowIndex, rawPartNumber, normalizedPartNumber, "", "NO_PART_NUMBER"
            End If
        ElseIf SKIP_FASTENER_IMAGES And IsFastenerExcelRow(rowIndex) Then
            gSkippedFastenerRows = gSkippedFastenerRows + 1
            SetRowUtilityValues rowIndex, "", "SKIPPED_FASTENER", "Fastener image skipped"
            WriteDebugPhase "FASTENER_SKIPPED", rowIndex, rawPartNumber, normalizedPartNumber, "", "Fastener image skipped."
        Else
            candidateRows = candidateRows + 1
            If TEST_MODE And candidateRows > TEST_MAX_ROWS Then
                SetRowUtilityValues rowIndex, "", "NOT_PROCESSED_TEST_LIMIT", "TEST_MODE limit reached"
            Else
                imagePath = BuildImagePath(normalizedPartNumber, rowIndex)
                thumbPath = ThumbnailPathForImage(imagePath)

                If ReuseExistingOrCachedImage(normalizedPartNumber, imagePath, thumbPath) Then
                    SetRowUtilityValues rowIndex, imagePath, "EXISTING_REUSED", ""
                    If InsertThumbnailForRow(rowIndex, thumbPath) Then
                        gProcessedImageRows = gProcessedImageRows + 1
                        gSuccessfulImageRows = gSuccessfulImageRows + 1
                        gReusedImageRows = gReusedImageRows + 1
                        WriteDebugPhase "EXISTING_REUSED", rowIndex, rawPartNumber, normalizedPartNumber, "", thumbPath
                    Else
                        SetRowUtilityValues rowIndex, imagePath, "IMAGE_CAPTURE_FAILED", "Could not insert existing image in Excel"
                    End If
                Else
                    sourcePath = SourcePathForPartNumber(rawPartNumber)
                    If sourcePath = "" Or Not gFSO.FileExists(sourcePath) Then
                        gSourceNotFoundRows = gSourceNotFoundRows + 1
                        SetRowUtilityValues rowIndex, "", "SOURCE_FILE_NOT_FOUND", "Source CATPart/CATProduct not found for Part Number"
                        WriteDebugPhase "SOURCE_FILE_NOT_FOUND", rowIndex, rawPartNumber, normalizedPartNumber, sourcePath, "Source not found."
                    Else
                        WriteDebugPhase "SOURCE_FILE_FOUND", rowIndex, rawPartNumber, normalizedPartNumber, sourcePath, sourcePath
                        If CaptureStandaloneImage(rawPartNumber, normalizedPartNumber, sourcePath, imagePath, thumbPath, rowIndex) Then
                            SetRowUtilityValues rowIndex, imagePath, "OK", ""
                            CacheImage normalizedPartNumber, imagePath, thumbPath
                            If InsertThumbnailForRow(rowIndex, thumbPath) Then
                                gProcessedImageRows = gProcessedImageRows + 1
                                gSuccessfulImageRows = gSuccessfulImageRows + 1
                                WriteDebugPhase "IMAGE_INSERTED_IN_EXCEL", rowIndex, rawPartNumber, normalizedPartNumber, sourcePath, thumbPath
                            Else
                                gImageCaptureFailedRows = gImageCaptureFailedRows + 1
                                SetRowUtilityValues rowIndex, imagePath, "IMAGE_CAPTURE_FAILED", "Could not insert image in Excel"
                            End If
                        Else
                            gImageCaptureFailedRows = gImageCaptureFailedRows + 1
                            SetRowUtilityValues rowIndex, imagePath, "IMAGE_CAPTURE_FAILED", "Standalone image capture failed"
                        End If
                    End If
                End If
            End If
        End If

        If SAVE_EVERY_N_ROWS > 0 Then
            If ((rowIndex - gHeaderRow) Mod SAVE_EVERY_N_ROWS) = 0 Then SaveExcelCheckpoint "Periodic save", rowIndex, rawPartNumber, normalizedPartNumber
        End If
        DoEvents
    Next

    ProcessBomRowsForThumbnails = True
    Err.Clear
End Function

Function CaptureStandaloneImage(rawPartNumber, normalizedPartNumber, sourcePath, imagePath, thumbPath, excelRow)
    On Error Resume Next
    CaptureStandaloneImage = False
    If gFSO.FileExists(imagePath) Then gFSO.DeleteFile imagePath, True
    If gFSO.FileExists(thumbPath) Then gFSO.DeleteFile thumbPath, True

    WriteDebugPhase "STANDALONE_OPEN_START", excelRow, rawPartNumber, normalizedPartNumber, sourcePath, sourcePath
    If Not OpenStandaloneDocument(sourcePath) Then
        CloseCurrentStandaloneDocument
        ActivateMainDocument
        Exit Function
    End If

    If Not CaptureActiveStandaloneViewer(imagePath) Then
        CloseCurrentStandaloneDocument
        ActivateMainDocument
        Exit Function
    End If
    WriteDebugPhase "STANDALONE_CAPTURE_DONE", excelRow, rawPartNumber, normalizedPartNumber, sourcePath, imagePath

    If Not CreateThumbnailFile(imagePath, thumbPath, THUMBNAIL_WIDTH, THUMBNAIL_HEIGHT) Then
        Err.Clear
        gFSO.CopyFile imagePath, thumbPath, True
    End If

    CloseCurrentStandaloneDocument
    ActivateMainDocument
    CaptureStandaloneImage = gFSO.FileExists(imagePath) And gFSO.FileExists(thumbPath)
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

    CATIA.StatusBar = "Opening standalone file: " & gFSO.GetFileName(sourcePath)
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

    viewer.RenderingMode = CAT_RENDER_SHADING_WITH_EDGES
    If Err.Number <> 0 Then
        Err.Clear
        viewer.RenderingMode = CAT_RENDER_SHADING
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
    viewpoint.ProjectionMode = CAT_PROJECTION_CYLINDRIC
    Err.Clear

    viewer.Reframe
    viewer.Update
    WaitSeconds IMAGE_CAPTURE_DELAY_SECONDS
    viewer.CaptureToFile CAT_CAPTURE_FORMAT_JPEG, imagePath
    CaptureActiveStandaloneViewer = gFSO.FileExists(imagePath)
    Err.Clear
End Function

Function CreateThumbnailFile(sourcePath, thumbPath, maxWidth, maxHeight)
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
    WriteDebugPhase "THUMBNAIL_CREATED", 0, "", "", thumbPath, CStr(CreateThumbnailFile)
    Err.Clear
End Function

Function InsertThumbnailForRow(rowIndex, thumbPath)
    On Error Resume Next
    InsertThumbnailForRow = False
    If thumbPath = "" Or Not gFSO.FileExists(thumbPath) Then Exit Function

    Dim cell
    Dim pic
    Set cell = gWsBom.Cells(CLng(rowIndex), gImageColumnIndex)
    DeletePicturesInCell gWsBom, cell
    gWsBom.Rows(CLng(rowIndex)).RowHeight = 90
    Set pic = gWsBom.Shapes.AddPicture(thumbPath, MSO_FALSE, MSO_TRUE, cell.Left + 2, cell.Top + 2, -1, -1)
    pic.LockAspectRatio = MSO_TRUE
    If pic.Width > THUMBNAIL_WIDTH Then pic.Width = THUMBNAIL_WIDTH
    If pic.Height > THUMBNAIL_HEIGHT Then pic.Height = THUMBNAIL_HEIGHT
    pic.Left = cell.Left + ((cell.Width - pic.Width) / 2)
    pic.Top = cell.Top + ((cell.Height - pic.Height) / 2)
    InsertThumbnailForRow = (Err.Number = 0)
    If InsertThumbnailForRow Then WriteDebugPhase "IMAGE_INSERTED_IN_EXCEL", rowIndex, "", "", "", thumbPath
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

Sub SetRowUtilityValues(rowIndex, imagePath, statusText, skipReason)
    On Error Resume Next
    If gImagePathColumnIndex > 0 Then gWsBom.Cells(rowIndex, gImagePathColumnIndex).Value = imagePath
    If gExportStatusColumnIndex > 0 Then gWsBom.Cells(rowIndex, gExportStatusColumnIndex).Value = statusText
    If gImageSkipReasonColumnIndex > 0 Then gWsBom.Cells(rowIndex, gImageSkipReasonColumnIndex).Value = skipReason
    Err.Clear
End Sub

Sub SaveExcelCheckpoint(phase, rowIndex, rawPartNumber, normalizedPartNumber)
    On Error Resume Next
    If gWorkbook Is Nothing Then Exit Sub
    gWorkbook.Save
    WriteDebugPhase "SAVE_CHECKPOINT", rowIndex, rawPartNumber, normalizedPartNumber, "", phase
    Err.Clear
End Sub

Function FinalSaveWorkbook()
    On Error Resume Next
    FinalSaveWorkbook = False
    If gWorkbook Is Nothing Then Exit Function
    gWorkbook.Save
    gExcelApp.DisplayAlerts = False
    Err.Clear
    gWorkbook.SaveAs gFinalExcelPath, XL_EXCEL8
    If Err.Number = 0 And gFSO.FileExists(gFinalExcelPath) Then
        FinalSaveWorkbook = True
        WriteDebugPhase "FINAL_SAVE_DONE", 0, "", "", gFinalExcelPath, gFinalExcelPath
    Else
        Dim saveErr
        saveErr = Err.Description
        Err.Clear
        gWorkbook.Save
        gWorkbook.Close False
        Set gWorkbook = Nothing
        gFSO.CopyFile gWorkExcelPath, gFinalExcelPath, True
        If Err.Number = 0 And gFSO.FileExists(gFinalExcelPath) Then
            Set gWorkbook = gExcelApp.Workbooks.Open(gFinalExcelPath)
            Set gWsBom = gWorkbook.Worksheets(1)
            FinalSaveWorkbook = True
            WriteDebugPhase "FINAL_SAVE_DONE", 0, "", "", gFinalExcelPath, "Saved by CopyFile after SaveAs failure: " & saveErr
        Else
            AbortPhase2 "Finalno snimanje Excela nije uspelo: " & saveErr & " / " & Err.Description, "FINAL_SAVE_DONE", 0, "", "", gFinalExcelPath
            Err.Clear
        End If
    End If
    gExcelApp.DisplayAlerts = True
    Err.Clear
End Function

Sub AbortPhase2(messageText, phase, rowIndex, rawPartNumber, normalizedPartNumber, sourcePath)
    On Error Resume Next
    gAbortAlreadyHandled = True
    gErrorCount = gErrorCount + 1
    WriteDebugPhase "ERROR", rowIndex, rawPartNumber, normalizedPartNumber, sourcePath, phase & ": " & messageText & "; Err.Number=" & CStr(Err.Number) & "; Err.Description=" & Err.Description
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

Sub WriteDebugPhase(phase, rowIndex, rawPartNumber, normalizedPartNumber, sourcePath, messageText)
    On Error Resume Next
    Dim ts
    Dim lineText
    lineText = FormatDateTime(Now, 2) & " " & FormatDateTime(Now, 3) & _
               " | " & phase & _
               " | ExcelRow=" & CStr(rowIndex) & _
               " | Raw Part Number=" & CStr(rawPartNumber) & _
               " | Normalized Part Number=" & CStr(normalizedPartNumber) & _
               " | Source Path=" & CStr(sourcePath) & _
               " | " & CStr(messageText)
    If gDebugLogPath = "" Then
        gPendingDebugText = gPendingDebugText & lineText & vbCrLf
        Exit Sub
    End If
    Set ts = gFSO.OpenTextFile(gDebugLogPath, 8, True)
    If gPendingDebugText <> "" Then
        ts.Write gPendingDebugText
        gPendingDebugText = ""
    End If
    ts.WriteLine lineText
    ts.Close
    Err.Clear
End Sub

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
        If Not gFSO.FileExists(thumbPath) Then
            If Not CreateThumbnailFile(imagePath, thumbPath, THUMBNAIL_WIDTH, THUMBNAIL_HEIGHT) Then gFSO.CopyFile imagePath, thumbPath, True
        End If
        If gFSO.FileExists(thumbPath) Then
            CacheImage normalizedPartNumber, imagePath, thumbPath
            ReuseExistingOrCachedImage = True
        End If
    End If
    Err.Clear
End Function

Function ExistingImageAvailable(normalizedPartNumber, rowIndex)
    Dim imagePath
    imagePath = BuildImagePath(normalizedPartNumber, rowIndex)
    ExistingImageAvailable = (SKIP_EXISTING_IMAGES And gFSO.FileExists(imagePath))
End Function

Sub CacheImage(normalizedPartNumber, imagePath, thumbPath)
    If normalizedPartNumber <> "" Then gImageCache.Item(normalizedPartNumber) = imagePath & "|" & thumbPath
End Sub

Function BuildImagePath(normalizedPartNumber, rowIndex)
    Dim baseName
    baseName = SafeFileName(normalizedPartNumber)
    If baseName = "" Then baseName = "BOM_ROW_" & CStr(rowIndex)
    If Len(baseName) > 140 Then baseName = Left(baseName, 140)
    BuildImagePath = JoinPath(gImageFolder, baseName & ".jpg")
End Function

Function ThumbnailPathForImage(imagePath)
    ThumbnailPathForImage = JoinPath(gThumbnailFolder, gFSO.GetBaseName(imagePath) & ".jpg")
End Function

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
        If c <> gImageColumnIndex And c <> gImagePathColumnIndex And c <> gExportStatusColumnIndex And c <> gImageSkipReasonColumnIndex Then
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

Function LooksLikeFilePath(pathText)
    LooksLikeFilePath = (CStr(pathText) <> "" And gFSO.GetExtensionName(CStr(pathText)) <> "")
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
    LastUsedRow = ws.Cells(ws.Rows.Count, 1).End(XL_UP).Row
    Dim used
    Set used = ws.UsedRange
    If used.Row + used.Rows.Count - 1 > LastUsedRow Then LastUsedRow = used.Row + used.Rows.Count - 1
    Err.Clear
End Function

Function LastUsedColumn(ws)
    On Error Resume Next
    LastUsedColumn = ws.Cells(gHeaderRow, ws.Columns.Count).End(XL_TO_LEFT).Column
    Err.Clear
End Function

Function RowHasAnyData(rowIndex)
    On Error Resume Next
    RowHasAnyData = False
    Dim used
    Dim maxCol
    Dim c
    Set used = gWsBom.UsedRange
    maxCol = used.Column + used.Columns.Count - 1
    For c = 1 To maxCol
        If c <> gImageColumnIndex And c <> gImagePathColumnIndex And c <> gExportStatusColumnIndex And c <> gImageSkipReasonColumnIndex Then
            If Trim(CStr(gWsBom.Cells(rowIndex, c).Value)) <> "" Then
                RowHasAnyData = True
                Exit Function
            End If
        End If
    Next
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

Function GetDefaultOutputFolder()
    On Error Resume Next
    GetDefaultOutputFolder = CStr(gProductDocument.Path)
    If GetDefaultOutputFolder = "" Then GetDefaultOutputFolder = gShell.SpecialFolders("Desktop")
    Err.Clear
End Function

Sub EnsureFolder(folderPath)
    If Not gFSO.FolderExists(folderPath) Then gFSO.CreateFolder folderPath
End Sub

Function JoinPath(folderPath, fileName)
    If Right(CStr(folderPath), 1) = "\" Then
        JoinPath = CStr(folderPath) & CStr(fileName)
    Else
        JoinPath = CStr(folderPath) & "\" & CStr(fileName)
    End If
End Function

Function AddTrailingSlash(folderPath)
    If Right(CStr(folderPath), 1) = "\" Then
        AddTrailingSlash = CStr(folderPath)
    Else
        AddTrailingSlash = CStr(folderPath) & "\"
    End If
End Function

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
