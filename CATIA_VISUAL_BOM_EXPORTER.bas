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
Const FORCE_DEFAULT_BOM_COLUMNS = False
Const USE_LOCAL_WORK_FOLDER_FOR_CATIA_PRINT = True
Const SMOKE_TEST_BOM_ONLY = False
Const DEBUG_FORCE_RECORDER_STYLE_PATH = False
Const STOP_ON_SOURCE_NOT_FOUND = False
Const TEST_MODE = True
Const TEST_MAX_ROWS = 20
Const TEST_PHASE_TIMEOUT_SECONDS = 30
Const SAVE_EVERY_N_ROWS = 25
Const RESUME_MODE = True
Const SKIP_EXISTING_IMAGES = True
Const STOP_INDEX_SCAN_WHEN_ALL_FOUND = True
Const INDEX_ONLY_NEEDED_PART_NUMBERS = True
Const MAX_SOURCE_NOT_FOUND_BEFORE_WARNING = 20
Const USE_AUTO_CROP_WHITE_BACKGROUND = True
Const CROP_PADDING_PERCENT = 0.1
Const MIN_CROP_PADDING_PIXELS = 25
Const USE_SHADED_WITH_EDGES = True
Const USE_WHITE_BACKGROUND = True
Const USE_PARALLEL_PROJECTION = True
Const IMAGE_CAPTURE_DELAY_SECONDS = 0.2

' CATIA constants used late-bound.
Const CAT_CAPTURE_FORMAT_JPEG = 2
Const CAT_RENDER_SHADING = 0
Const CAT_RENDER_SHADING_WITH_EDGES = 1
Const CAT_PROJECTION_CYLINDRIC = 0

' Excel constants used late-bound.
Const XL_OPENXML_WORKBOOK = 51
Const XL_EXCEL8 = 56
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
Dim gWorkExcelPath
Dim gFinalExcelPath
Dim gOutputFolder
Dim gImageFolder
Dim gCroppedFolder
Dim gThumbnailFolder
Dim gFinalOutputFolder
Dim gFinalImageFolder
Dim gFinalCroppedFolder
Dim gFinalThumbnailFolder
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
Dim gCroppedImagePathColumnIndex
Dim gThumbnailPathColumnIndex
Dim gExportStatusColumnIndex
Dim gImageSkipReasonColumnIndex
Dim gLastBomRow
Dim gSourceIndex
Dim gNeededPartNumbers
Dim gFoundNeededPartNumbers
Dim gImageCache
Dim gNextLogRow
Dim gProcessedImageRows
Dim gSuccessfulImageRows
Dim gReusedImageRows
Dim gSkippedFastenerRows
Dim gRowsWithoutPartNumber
Dim gSourceNotFoundRows
Dim gSourceNotFoundWarningShown
Dim gErrorCount
Dim gCurrentStandaloneDoc
Dim gCurrentStandaloneOpened
Dim gAbortAlreadyHandled
Dim gFirstImageStatusShown
Dim gPendingDebugText

Public Sub CATIA_VISUAL_BOM_EXPORTER()
    On Error Resume Next

    InitializeRuntime
    If Not ValidateActiveProductDocument() Then Exit Sub
    If Not SelectExportPathAndPrepareFolders() Then Exit Sub

    WriteDebugPhase "START", 0, "", "", "", "CATIA_VISUAL_BOM_EXPORTER started."

    If Not PrintBomToXls(gWorkExcelPath) Then
        If gAbortAlreadyHandled Then Exit Sub
        AbortWithMessage "CATIA BillOfMaterial Print XLS nije uspeo.", "CATIA_BOM_PRINT_XLS_START", 0, "", "", ""
        Exit Sub
    End If

    If SMOKE_TEST_BOM_ONLY Then
        WriteDebugPhase "SMOKE_TEST_BOM_ONLY", 0, "", "", gWorkExcelPath, "BOM XLS export successful; stopping before Excel image workflow."
        CopySmokeTestWorkbookToFinal
        CleanupCatiaSession
        MsgBox "BOM XLS export uspesan" & vbCrLf & vbCrLf & gFinalExcelPath, vbInformation, "CATIA_VISUAL_BOM_EXPORTER"
        Exit Sub
    End If

    If Not OpenBomWorkbookAndPrepareSheets() Then
        If gAbortAlreadyHandled Then Exit Sub
        AbortWithMessage "Ne mogu da otvorim lokalni Excel BOM fajl: " & gWorkExcelPath, "EXCEL_OPENED_LOCAL_WORKBOOK", 0, "", "", ""
        Exit Sub
    End If

    BuildNeededPartNumbersFromBomRows

    Dim indexPhaseStart
    indexPhaseStart = StartTimedPhase("CATIA_FILE_INDEX", "4/6 Indeksiram CATIA partove...")
    WriteDebugPhase "CATIA_FILE_INDEX_START", 0, "", "", "", "Building Part Number -> source file path index for needed items: " & CStr(gNeededPartNumbers.Count)
    BuildCatiaFileIndex
    WriteDebugPhase "CATIA_FILE_INDEX_DONE", 0, "", "", "", "Indexed source files: " & CStr(gSourceIndex.Count) & " / Needed: " & CStr(gNeededPartNumbers.Count)
    If Not CheckTimedPhase("CATIA_FILE_INDEX", indexPhaseStart, 0, "", "", "") Then Exit Sub

    If Not ProcessBomRowsForThumbnails() Then Exit Sub

    SaveExcelCheckpoint "FINISH", 0, "", ""
    FinalizeExcelExport
    CleanupCatiaSession
    CATIA.StatusBar = "6/6 Gotovo."
    WriteDebugPhase "FINISH", 0, "", "", "", "Macro finished."
    CopyDebugLogToFinal
    MsgBox "CATIA BOM Excel export je zavrsen." & vbCrLf & vbCrLf & _
           "Excel:" & vbCrLf & gFinalExcelPath, vbInformation, "CATIA_VISUAL_BOM_EXPORTER"
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
    Set gWsLog = Nothing
    Set gWsSummary = Nothing
    Set gCurrentStandaloneDoc = Nothing
    gExcelPath = ""
    gWorkExcelPath = ""
    gFinalExcelPath = ""
    gOutputFolder = ""
    gFinalOutputFolder = ""
    gDebugLogPath = ""
    gCurrentStandaloneOpened = False
    gHeaderRow = 0
    gPartNumberColumnIndex = 0
    gThumbnailColumnIndex = 0
    gImagePathColumnIndex = 0
    gCroppedImagePathColumnIndex = 0
    gThumbnailPathColumnIndex = 0
    gExportStatusColumnIndex = 0
    gImageSkipReasonColumnIndex = 0
    gLastBomRow = 0
    gNextLogRow = 2
    gProcessedImageRows = 0
    gSuccessfulImageRows = 0
    gReusedImageRows = 0
    gSkippedFastenerRows = 0
    gRowsWithoutPartNumber = 0
    gSourceNotFoundRows = 0
    gSourceNotFoundWarningShown = False
    gErrorCount = 0
    gAbortAlreadyHandled = False
    gFirstImageStatusShown = False
    gPendingDebugText = ""
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

Function SelectExportPathAndPrepareFolders()
    On Error Resume Next
    SelectExportPathAndPrepareFolders = False

    Dim defaultFolder
    Dim selectedFolder
    Dim rootPartNumber

    rootPartNumber = GetProductPartNumber(gProduct)
    defaultFolder = GetDefaultOutputFolder()

    If DEBUG_FORCE_RECORDER_STYLE_PATH Then
        PrepareLocalWorkFolder
        gWorkExcelPath = "C:\Temp\CATIA_VISUAL_BOM_WORK\CATIA_BOM_TEST.xls"
        gFinalExcelPath = gWorkExcelPath
        gExcelPath = gWorkExcelPath
        PrepareWorkFolders rootPartNumber
        gFinalOutputFolder = gOutputFolder
        gFinalImageFolder = gImageFolder
        gFinalCroppedFolder = gCroppedFolder
        gFinalThumbnailFolder = gThumbnailFolder
        If gFSO.FileExists(gWorkExcelPath) Then
            Err.Clear
            gFSO.DeleteFile gWorkExcelPath, True
            If Err.Number <> 0 Then
                AbortWithMessage "Ne mogu da obrisem DEBUG XLS fajl. Zatvorite fajl ako je otvoren u Excelu: " & gWorkExcelPath, "OUTPUT_XLS_OVERWRITE", 0, "", "", gWorkExcelPath
                Err.Clear
                Exit Function
            End If
        End If
        WriteDebugPhase "LOCAL_WORK_FOLDER_PREPARED", 0, "", "", gOutputFolder, gOutputFolder
        WriteDebugPhase "LOCAL_WORK_XLS_PATH_PREPARED", 0, "", "", gWorkExcelPath, "DEBUG_FORCE_RECORDER_STYLE_PATH=True"
        SelectExportPathAndPrepareFolders = True
        Exit Function
    End If

    CATIA.StatusBar = "Cekam izbor foldera za Excel export..."
    WriteDebugPhase "USER_FOLDER_DIALOG_OPENED", 0, "", "", "", "Waiting for user folder selection - no timeout."
    selectedFolder = AskUserForOutputFolder(defaultFolder)
    If selectedFolder = "" Then
        CATIA.StatusBar = ""
        WriteDebugPhase "USER_FOLDER_CANCELLED", 0, "", "", "", "User cancelled folder selection."
        MsgBox "Export je otkazan od strane korisnika.", vbInformation, "CATIA_VISUAL_BOM_EXPORTER"
        Exit Function
    End If

    WriteDebugPhase "USER_FOLDER_SELECTED", 0, "", "", selectedFolder, selectedFolder
    gFinalExcelPath = BuildBomExcelPath(selectedFolder, rootPartNumber)
    gFinalExcelPath = ResolveExistingOutputFile(gFinalExcelPath)
    If gFinalExcelPath = "" Then
        CATIA.StatusBar = ""
        If Not gAbortAlreadyHandled Then
            WriteDebugPhase "OUTPUT_XLS_CANCELLED", 0, "", "", "", "User cancelled existing file decision."
            MsgBox "Export je otkazan od strane korisnika.", vbInformation, "CATIA_VISUAL_BOM_EXPORTER"
        End If
        Exit Function
    End If

    If USE_LOCAL_WORK_FOLDER_FOR_CATIA_PRINT Then
        PrepareLocalWorkPaths rootPartNumber
        If gAbortAlreadyHandled Then Exit Function
    Else
        gWorkExcelPath = gFinalExcelPath
        gExcelPath = gWorkExcelPath
        PrepareFoldersNextToExcel gWorkExcelPath
    End If
    PrepareFinalFolderVariables gFinalExcelPath
    WriteDebugPhase "FINAL_XLS_PATH_PREPARED", 0, "", "", gFinalExcelPath, gFinalExcelPath
    WriteDebugPhase "LOCAL_WORK_FOLDER_PREPARED", 0, "", "", gOutputFolder, gOutputFolder
    WriteDebugPhase "LOCAL_WORK_XLS_PATH_PREPARED", 0, "", "", gWorkExcelPath, gWorkExcelPath
    SelectExportPathAndPrepareFolders = True
    Err.Clear
End Function

Function AskUserForOutputFolder(defaultFolder)
    On Error Resume Next
    AskUserForOutputFolder = ""
    Dim excelDialogResult
    excelDialogResult = AskUserForOutputFolderExcelDialog(defaultFolder)
    If excelDialogResult = "__CANCEL__" Then Exit Function
    If excelDialogResult <> "" Then
        AskUserForOutputFolder = excelDialogResult
        Exit Function
    End If

    Dim shellApp
    Dim folder
    Dim folderPath
    Set shellApp = CreateObject("Shell.Application")
    If Err.Number <> 0 Or shellApp Is Nothing Then
        Err.Clear
        AskUserForOutputFolder = AskUserForOutputFolderInputBox(defaultFolder)
        Exit Function
    End If
    Err.Clear
    Set folder = shellApp.BrowseForFolder(0, "Izaberite folder gde zelite da sacuvate CATIA BOM Excel export", 0, defaultFolder)
    If Err.Number <> 0 Then
        Err.Clear
        AskUserForOutputFolder = AskUserForOutputFolderInputBox(defaultFolder)
        Exit Function
    End If
    If folder Is Nothing Then
        Err.Clear
        Exit Function
    End If
    folderPath = folder.Self.Path
    AskUserForOutputFolder = CStr(folderPath)
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
    fd.Title = "Izaberite folder gde zelite da sacuvate CATIA BOM Excel export"
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

Function AskUserForOutputFolderInputBox(defaultFolder)
    On Error Resume Next
    Dim folderPath
    folderPath = InputBox("Unesite/potvrdite folder za CATIA BOM Excel export:", "CATIA_VISUAL_BOM_EXPORTER", defaultFolder)
    folderPath = Trim(CStr(folderPath))
    If folderPath <> "" Then
        If gFSO.FolderExists(folderPath) Then
            AskUserForOutputFolderInputBox = folderPath
        Else
            MsgBox "Folder ne postoji: " & folderPath, vbExclamation, "CATIA_VISUAL_BOM_EXPORTER"
            AskUserForOutputFolderInputBox = ""
        End If
    Else
        AskUserForOutputFolderInputBox = ""
    End If
    Err.Clear
End Function

Function GetDefaultOutputFolder()
    On Error Resume Next
    GetDefaultOutputFolder = CStr(gProductDocument.Path)
    If GetDefaultOutputFolder = "" Then GetDefaultOutputFolder = gShell.SpecialFolders("Desktop")
    Err.Clear
End Function

Function BuildBomExcelPath(selectedFolder, rootPartNumber)
    Dim baseName
    baseName = SafeFileName(rootPartNumber)
    If baseName = "" Then baseName = SafeFileName(gProduct.Name)
    If baseName = "" Then baseName = "CATIA_BOM"
    BuildBomExcelPath = EnsureXlsExtension(JoinPath(selectedFolder, baseName & "_VISUAL_BOM_EXPORT.xls"))
End Function

Function ResolveExistingOutputFile(xlsPath)
    On Error Resume Next
    xlsPath = EnsureXlsExtension(xlsPath)
    ResolveExistingOutputFile = xlsPath
    If Not gFSO.FileExists(xlsPath) Then Exit Function

    WriteDebugPhase "OUTPUT_XLS_EXISTS", 0, "", "", xlsPath, xlsPath
    Dim answer
    answer = MsgBox("Fajl vec postoji:" & vbCrLf & xlsPath & vbCrLf & vbCrLf & _
                    "YES = zameni postojeci fajl" & vbCrLf & _
                    "NO = napravi novi fajl sa timestamp nastavkom" & vbCrLf & _
                    "CANCEL = prekini export", _
                    vbYesNoCancel + vbQuestion, "CATIA_VISUAL_BOM_EXPORTER")

    If answer = vbYes Then
        Err.Clear
        gFSO.DeleteFile xlsPath, True
        If Err.Number <> 0 Then
            AbortWithMessage "Ne mogu da obrisem postojeci XLS. Zatvorite fajl ako je otvoren u Excelu: " & xlsPath, "OUTPUT_XLS_OVERWRITE", 0, "", "", xlsPath
            Err.Clear
            ResolveExistingOutputFile = ""
            Exit Function
        End If
        WriteDebugPhase "OUTPUT_XLS_OVERWRITE_CONFIRMED", 0, "", "", xlsPath, xlsPath
        ResolveExistingOutputFile = xlsPath
    ElseIf answer = vbNo Then
        ResolveExistingOutputFile = Replace(EnsureXlsExtension(xlsPath), ".xls", "_" & TimestampForFile() & ".xls")
        WriteDebugPhase "OUTPUT_XLS_TIMESTAMP_CREATED", 0, "", "", ResolveExistingOutputFile, ResolveExistingOutputFile
    Else
        ResolveExistingOutputFile = ""
    End If
    Err.Clear
End Function

Sub PrepareFoldersNextToExcel(bomExcelPath)
    gOutputFolder = JoinPath(gFSO.GetParentFolderName(bomExcelPath), gFSO.GetBaseName(bomExcelPath) & "_FILES")
    gImageFolder = JoinPath(gOutputFolder, "IMAGES")
    gCroppedFolder = JoinPath(gOutputFolder, "CROPPED")
    gThumbnailFolder = JoinPath(gOutputFolder, "THUMBNAILS")
    gDebugLogPath = JoinPath(gOutputFolder, "DEBUG_PHASE_LOG.txt")
    EnsureFolder gOutputFolder
    EnsureFolder gImageFolder
    EnsureFolder gCroppedFolder
    EnsureFolder gThumbnailFolder
End Sub

Sub PrepareLocalWorkPaths(rootPartNumber)
    On Error Resume Next
    PrepareLocalWorkFolder
    gWorkExcelPath = JoinPath("C:\Temp\CATIA_VISUAL_BOM_WORK", SafeFileNameOrDefault(rootPartNumber, "CATIA_BOM") & "_VISUAL_BOM_EXPORT_WORK.xls")
    gWorkExcelPath = EnsureXlsExtension(gWorkExcelPath)
    gExcelPath = gWorkExcelPath
    If gFSO.FileExists(gWorkExcelPath) Then
        Err.Clear
        gFSO.DeleteFile gWorkExcelPath, True
        If Err.Number <> 0 Then
            AbortWithMessage "Ne mogu da obrisem lokalni work XLS. Zatvorite fajl ako je otvoren u Excelu: " & gWorkExcelPath, "LOCAL_WORK_XLS_PATH_PREPARED", 0, "", "", gWorkExcelPath
            Err.Clear
            Exit Sub
        End If
    End If
    PrepareWorkFolders rootPartNumber
    Err.Clear
End Sub

Sub PrepareLocalWorkFolder()
    EnsureFolder "C:\Temp"
    EnsureFolder "C:\Temp\CATIA_VISUAL_BOM_WORK"
End Sub

Sub PrepareWorkFolders(rootPartNumber)
    gOutputFolder = JoinPath("C:\Temp\CATIA_VISUAL_BOM_WORK", SafeFileNameOrDefault(rootPartNumber, "CATIA_BOM") & "_VISUAL_BOM_EXPORT_FILES")
    gImageFolder = JoinPath(gOutputFolder, "IMAGES")
    gCroppedFolder = JoinPath(gOutputFolder, "CROPPED")
    gThumbnailFolder = JoinPath(gOutputFolder, "THUMBNAILS")
    gDebugLogPath = JoinPath(gOutputFolder, "DEBUG_PHASE_LOG.txt")
    EnsureFolder gOutputFolder
    EnsureFolder gImageFolder
    EnsureFolder gCroppedFolder
    EnsureFolder gThumbnailFolder
End Sub

Sub PrepareFinalFolderVariables(finalExcelPath)
    gFinalOutputFolder = JoinPath(gFSO.GetParentFolderName(finalExcelPath), gFSO.GetBaseName(finalExcelPath) & "_FILES")
    gFinalImageFolder = JoinPath(gFinalOutputFolder, "IMAGES")
    gFinalCroppedFolder = JoinPath(gFinalOutputFolder, "CROPPED")
    gFinalThumbnailFolder = JoinPath(gFinalOutputFolder, "THUMBNAILS")
End Sub

Function PrintBomToXls(bomExcelPath)
    On Error Resume Next
    PrintBomToXls = False
    Dim productDocument1
    Dim product1
    Dim assemblyConvertor1
    Dim arrayOfVariantOfBSTR1(6)
    Dim parentFolder
    Dim phaseStart

    bomExcelPath = EnsureXlsExtension(bomExcelPath)
    gExcelPath = bomExcelPath
    parentFolder = gFSO.GetParentFolderName(bomExcelPath)
    WriteDebugPhase "BOM_PRINT_PATH_CHECK", 0, "", "", bomExcelPath, _
                    "Path=" & bomExcelPath & _
                    "; ParentFolder=" & parentFolder & _
                    "; FolderExists=" & CStr(gFSO.FolderExists(parentFolder)) & _
                    "; FileExistsBeforePrint=" & CStr(gFSO.FileExists(bomExcelPath)) & _
                    "; IsLocalWindowsPath=" & CStr(IsLocalWindowsPath(bomExcelPath))

    If parentFolder = "" Or Not gFSO.FolderExists(parentFolder) Then
        AbortWithMessage "Folder za XLS export ne postoji: " & parentFolder, "BOM_PRINT_PATH_CHECK", 0, "", "", bomExcelPath
        Exit Function
    End If

    If gFSO.FileExists(bomExcelPath) Then
        AbortWithMessage "XLS fajl i dalje postoji pre CATIA Print komande. Zatvorite ga ako je otvoren u Excelu: " & bomExcelPath, "BOM_PRINT_PATH_CHECK", 0, "", "", bomExcelPath
        Exit Function
    End If

    Set productDocument1 = CATIA.ActiveDocument
    Set product1 = productDocument1.Product
    Set assemblyConvertor1 = product1.GetItem("BillOfMaterial")
    Set gAssemblyConvertor = assemblyConvertor1

    If Err.Number <> 0 Or assemblyConvertor1 Is Nothing Then
        WriteDebugPhase "ERROR", 0, "", "", bomExcelPath, "GetItem(""BillOfMaterial"") failed. Err.Number=" & CStr(Err.Number) & "; Err.Description=" & Err.Description
        AbortWithMessage "CATIA BillOfMaterial objekat nije dostupan preko product.GetItem(""BillOfMaterial"").", "BOM_FORMAT_SET", 0, "", "", bomExcelPath
        Err.Clear
        Exit Function
    End If

    If FORCE_DEFAULT_BOM_COLUMNS Then
        arrayOfVariantOfBSTR1(0) = "Nomenclature"
        arrayOfVariantOfBSTR1(1) = "Quantity"
        arrayOfVariantOfBSTR1(2) = "Part Number"
        arrayOfVariantOfBSTR1(3) = "Dimenzija"
        arrayOfVariantOfBSTR1(4) = "Material"
        arrayOfVariantOfBSTR1(5) = "Mass"
        arrayOfVariantOfBSTR1(6) = "Standard"

        Err.Clear
        WriteDebugPhase "BOM_FORMAT_SET_START", 0, "", "", bomExcelPath, "SetSecondaryFormat start because FORCE_DEFAULT_BOM_COLUMNS=True."
        assemblyConvertor1.SetSecondaryFormat arrayOfVariantOfBSTR1
        If Err.Number <> 0 Then
            WriteDebugPhase "ERROR", 0, "", "", bomExcelPath, "SetSecondaryFormat failed. Err.Number=" & CStr(Err.Number) & "; Err.Description=" & Err.Description
            AbortWithMessage "CATIA BillOfMaterial.SetSecondaryFormat nije uspeo: " & Err.Description, "BOM_FORMAT_SET", 0, "", "", bomExcelPath
            Err.Clear
            Exit Function
        End If
        WriteDebugPhase "BOM_FORMAT_SET_DONE", 0, "", "", bomExcelPath, "SetSecondaryFormat OK."
    Else
        WriteDebugPhase "BOM_FORMAT_USER_DEFINED", 0, "", "", bomExcelPath, "Using CATIA user-defined BOM format; SetSecondaryFormat not called."
    End If

    If SMOKE_TEST_BOM_ONLY Then
        WriteDebugPhase "SMOKE_TEST_BOM_ONLY", 0, "", "", bomExcelPath, "Smoke test enabled; only BOM XLS export will run."
    End If

    WriteDebugPhase "CATIA_BOM_PRINT_XLS_START", 0, "", "", bomExcelPath, bomExcelPath
    CATIA.StatusBar = "2/6 Exportujem CATIA BOM u Excel..."
    phaseStart = Timer
    Err.Clear
    assemblyConvertor1.Print "XLS", bomExcelPath, product1
    If Err.Number <> 0 Then
        WriteDebugPhase "CATIA_BOM_PRINT_XLS_FAILED", 0, "", "", bomExcelPath, _
                        "Err.Number=" & CStr(Err.Number) & _
                        "; Err.Description=" & Err.Description & _
                        "; Path=" & bomExcelPath
        AbortWithMessage "CATIA BillOfMaterial.Print XLS nije uspeo." & vbCrLf & _
                         "Proverite da li je XLS fajl zatvoren i da li folder postoji." & vbCrLf & _
                         "Putanja: " & bomExcelPath & vbCrLf & _
                         "Detalji su u DEBUG_PHASE_LOG.txt.", _
                         "CATIA_BOM_PRINT_XLS_START", 0, "", "", bomExcelPath
        Err.Clear
        Exit Function
    End If

    If Not gFSO.FileExists(bomExcelPath) Then
        WriteDebugPhase "CATIA_BOM_PRINT_XLS_FAILED", 0, "", "", bomExcelPath, "Print returned no error, but XLS file does not exist: " & bomExcelPath
        AbortWithMessage "CATIA nije napravila XLS fajl: " & bomExcelPath, "CATIA_BOM_PRINT_XLS_DONE", 0, "", "", bomExcelPath
        Exit Function
    End If

    CATIA.StatusBar = "3/6 BOM Excel napravljen."
    WriteDebugPhase "CATIA_BOM_PRINT_XLS_DONE", 0, "", "", bomExcelPath, bomExcelPath
    If Not CheckTimedPhase("CATIA_BOM_PRINT_XLS", phaseStart, 0, "", "", bomExcelPath) Then Exit Function
    PrintBomToXls = True
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

    gExcelApp.Visible = True
    gExcelApp.DisplayAlerts = False
    gExcelApp.ScreenUpdating = False
    Set gWorkbook = gExcelApp.Workbooks.Open(gWorkExcelPath)
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
    WriteDebugPhase "EXCEL_OPENED_LOCAL_WORKBOOK", 0, "", "", gWorkExcelPath, "Local work workbook opened."
    WriteDebugPhase "BOM_HEADERS_READ", gHeaderRow, "", "", "", "Part Number column=" & CStr(gPartNumberColumnIndex)

    EnsureHelperColumns
    Set gWsLog = GetOrCreateWorksheet("EXPORT_LOG")
    PrepareLogSheet
    Set gWsSummary = GetOrCreateWorksheet("SUMMARY")
    PrepareSummarySheet
    gLastBomRow = LastUsedRow(gWsBom)
    SaveExcelCheckpoint "EXCEL_OPENED_LOCAL_WORKBOOK", 0, "", ""
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
    IsPartNumberHeader = (h = "partnumber" Or h = "number" Or h = "partno" Or h = "pn" Or h = "brojdela" Or h = "brdela")
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
    gCroppedImagePathColumnIndex = EnsureEndHelperColumn("Cropped Image Path")
    gThumbnailPathColumnIndex = EnsureEndHelperColumn("Thumbnail Path")
    gExportStatusColumnIndex = EnsureEndHelperColumn("Export Status")
    gImageSkipReasonColumnIndex = EnsureEndHelperColumn("Image Skip Reason")

    gWsBom.Rows(gHeaderRow).Font.Bold = True
    gWsBom.Columns(gThumbnailColumnIndex).ColumnWidth = 24
    gWsBom.Columns(gImagePathColumnIndex).ColumnWidth = 45
    gWsBom.Columns(gCroppedImagePathColumnIndex).ColumnWidth = 45
    gWsBom.Columns(gThumbnailPathColumnIndex).ColumnWidth = 45
    gWsBom.Columns(gExportStatusColumnIndex).ColumnWidth = 22
    gWsBom.Columns(gImageSkipReasonColumnIndex).ColumnWidth = 42
    ApplyBomBorders
    WriteDebugPhase "HELPER_COLUMNS_ADDED", gHeaderRow, "", "", "", "Thumbnail/Image/Cropped/Status helper columns prepared."
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
    gWsLog.Cells(1, 10).Value = "Cropped Image Path"
    gWsLog.Cells(1, 11).Value = "Thumbnail Path"
    gWsLog.Cells(1, 12).Value = "Source Path"
    gWsLog.Rows(1).Font.Bold = True
    gWsLog.Columns("A:L").ColumnWidth = 24
End Sub

Sub PrepareSummarySheet()
    gWsSummary.Cells(1, 1).Value = "CATIA VISUAL BOM EXPORTER"
    gWsSummary.Cells(3, 1).Value = "Main Assembly Part Number"
    gWsSummary.Cells(4, 1).Value = "Export date/time"
    gWsSummary.Cells(5, 1).Value = "Work Excel path"
    gWsSummary.Cells(6, 1).Value = "Final Excel path"
    gWsSummary.Cells(7, 1).Value = "Output folder"
    gWsSummary.Cells(8, 1).Value = "Total BOM rows"
    gWsSummary.Cells(9, 1).Value = "Unique needed Part Numbers"
    gWsSummary.Cells(10, 1).Value = "Rows without Part Number"
    gWsSummary.Cells(11, 1).Value = "Images processed"
    gWsSummary.Cells(12, 1).Value = "Successful images"
    gWsSummary.Cells(13, 1).Value = "Reused images"
    gWsSummary.Cells(14, 1).Value = "Skipped fasteners"
    gWsSummary.Cells(15, 1).Value = "Source files not found"
    gWsSummary.Cells(16, 1).Value = "Errors"
    gWsSummary.Cells(17, 1).Value = "Mode"
    gWsSummary.Cells(18, 1).Value = "Debug log path"
    gWsSummary.Cells(19, 1).Value = "FORCE_DEFAULT_BOM_COLUMNS"
    gWsSummary.Cells(20, 1).Value = "USE_LOCAL_WORK_FOLDER_FOR_CATIA_PRINT"
    gWsSummary.Cells(21, 1).Value = "TEST_MODE"
    gWsSummary.Range("A1:B1").Font.Bold = True
    gWsSummary.Columns("A").ColumnWidth = 28
    gWsSummary.Columns("B").ColumnWidth = 95
    UpdateSummarySheet
End Sub

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
                    If Not ExistingImageSetAvailable(normalizedPartNumber, rowIndex) Then
                        If Not gNeededPartNumbers.Exists(normalizedPartNumber) Then gNeededPartNumbers.Item(normalizedPartNumber) = True
                    End If
                End If
            End If
        End If
        If (rowIndex Mod 250) = 0 Then DoEvents
    Next

    WriteDebugPhase "NEEDED_PART_NUMBERS_BUILT", 0, "", "", "", "Unique needed Part Numbers=" & CStr(gNeededPartNumbers.Count) & "; TEST_MODE=" & CStr(TEST_MODE) & "; TEST_MAX_ROWS=" & CStr(TEST_MAX_ROWS)
    Err.Clear
End Sub

Function ProcessBomRowsForThumbnails()
    On Error Resume Next
    ProcessBomRowsForThumbnails = False
    Dim rowIndex
    Dim rawPartNumber
    Dim normalizedPartNumber
    Dim sourcePath
    Dim imagePath
    Dim croppedPath
    Dim thumbPath
    Dim reason
    Dim firstImagePhaseStart
    Dim watchFirstImagePhase
    Dim testCandidateRows
    testCandidateRows = 0

    For rowIndex = gHeaderRow + 1 To gLastBomRow
        watchFirstImagePhase = False
        rawPartNumber = CStr(gWsBom.Cells(rowIndex, gPartNumberColumnIndex).Value)
        normalizedPartNumber = NormalizePartNumber(rawPartNumber)
        If Trim(rawPartNumber) <> "" Then
            CATIA.StatusBar = "BOM thumbnail row " & CStr(rowIndex - gHeaderRow) & " / " & CStr(gLastBomRow - gHeaderRow) & " - " & rawPartNumber
            WriteDebugPhase "ROW_START", rowIndex, rawPartNumber, normalizedPartNumber, "", "Raw BOM Part Number=" & rawPartNumber & "; Normalized BOM Part Number=" & normalizedPartNumber

            If SKIP_FASTENER_IMAGES And IsFastenerExcelRow(rowIndex) Then
                gSkippedFastenerRows = gSkippedFastenerRows + 1
                reason = "Standard fastener - retained in BOM, image skipped"
                SetRowUtilityValues rowIndex, "", "", "", "SKIPPED_IMAGE_ONLY", reason
                WriteExportLog rowIndex, rawPartNumber, normalizedPartNumber, "SKIPPED_IMAGE_ONLY", "FASTENER_SKIPPED_IMAGE", reason, "", "", "", ""
                WriteDebugPhase "FASTENER_SKIPPED_IMAGE", rowIndex, rawPartNumber, normalizedPartNumber, "", reason
            Else
                testCandidateRows = testCandidateRows + 1
                If TEST_MODE And testCandidateRows > TEST_MAX_ROWS Then
                    SetRowUtilityValues rowIndex, "", "", "", "NOT_PROCESSED_TEST_LIMIT", "TEST_MODE limit reached"
                    WriteExportLog rowIndex, rawPartNumber, normalizedPartNumber, "NOT_PROCESSED_TEST_LIMIT", "ROW_START", "TEST_MODE limit reached.", "", "", "", ""
                    WriteDebugPhase "ROW_START", rowIndex, rawPartNumber, normalizedPartNumber, "", "TEST_MODE limit reached; row not processed."
                Else
                imagePath = BuildImagePath(normalizedPartNumber, rowIndex)
                croppedPath = CroppedPathForImage(imagePath)
                thumbPath = ThumbnailPathForImage(imagePath)

                If ReuseExistingOrCachedImage(normalizedPartNumber, imagePath, croppedPath, thumbPath) Then
                    SetRowUtilityValues rowIndex, imagePath, croppedPath, thumbPath, "EXISTING_REUSED", ""
                    If Not InsertThumbnailForRow(rowIndex, thumbPath) Then
                        AbortWithMessage "Excel ne moze da ubaci thumbnail za Part Number: " & rawPartNumber, "EXCEL_THUMBNAIL_INSERTED", rowIndex, rawPartNumber, normalizedPartNumber, ""
                        Exit Function
                    End If
                    gProcessedImageRows = gProcessedImageRows + 1
                    gSuccessfulImageRows = gSuccessfulImageRows + 1
                    gReusedImageRows = gReusedImageRows + 1
                    WriteDebugPhase "EXISTING_REUSED", rowIndex, rawPartNumber, normalizedPartNumber, "", thumbPath
                    WriteExportLog rowIndex, rawPartNumber, normalizedPartNumber, "EXISTING_REUSED", "EXCEL_THUMBNAIL_INSERTED", "Existing image/cropped/thumbnail reused.", imagePath, croppedPath, thumbPath, ""
                    MaybeSaveByProgress rowIndex, rawPartNumber, normalizedPartNumber
                Else
                    sourcePath = SourcePathForPartNumber(rawPartNumber)
                    If sourcePath = "" Or Not gFSO.FileExists(sourcePath) Then
                    gSourceNotFoundRows = gSourceNotFoundRows + 1
                    reason = "Source CATPart/CATProduct not found for Part Number"
                    SetRowUtilityValues rowIndex, "", "", "", "SOURCE_FILE_NOT_FOUND", reason
                    WriteExportLog rowIndex, rawPartNumber, normalizedPartNumber, "SOURCE_FILE_NOT_FOUND", "SOURCE_FILE_NOT_FOUND", reason, "", "", "", sourcePath
                    WriteDebugPhase "SOURCE_FILE_NOT_FOUND", rowIndex, rawPartNumber, normalizedPartNumber, sourcePath, reason
                    If Not gSourceNotFoundWarningShown And MAX_SOURCE_NOT_FOUND_BEFORE_WARNING > 0 And gSourceNotFoundRows >= MAX_SOURCE_NOT_FOUND_BEFORE_WARNING Then
                        gSourceNotFoundWarningShown = True
                        WriteDebugPhase "SOURCE_FILE_NOT_FOUND", rowIndex, rawPartNumber, normalizedPartNumber, sourcePath, "Warning threshold reached: " & CStr(gSourceNotFoundRows) & " source files not found."
                    End If
                    If STOP_ON_SOURCE_NOT_FOUND Then
                        AbortWithMessage "Source file not found for Part Number: " & rawPartNumber, "SOURCE_FILE_NOT_FOUND", rowIndex, rawPartNumber, normalizedPartNumber, sourcePath
                        Exit Function
                    End If
                ElseIf Not IsSupportedCatiaSourceFile(sourcePath) Then
                    reason = "Unsupported source file: " & sourcePath
                    SetRowUtilityValues rowIndex, "", "", "", "UNSUPPORTED_SOURCE_FILE", reason
                    WriteExportLog rowIndex, rawPartNumber, normalizedPartNumber, "UNSUPPORTED_SOURCE_FILE", "UNSUPPORTED_SOURCE_FILE", reason, "", "", "", sourcePath
                    WriteDebugPhase "UNSUPPORTED_SOURCE_FILE", rowIndex, rawPartNumber, normalizedPartNumber, sourcePath, reason
                Else
                    WriteDebugPhase "SOURCE_FILE_FOUND", rowIndex, rawPartNumber, normalizedPartNumber, sourcePath, sourcePath
                    If Not gFirstImageStatusShown Then
                        firstImagePhaseStart = StartTimedPhase("FIRST_IMAGE_STANDALONE_CAPTURE", "5/6 Otvaram prvi part za sliku...")
                        gFirstImageStatusShown = True
                        watchFirstImagePhase = True
                    End If
                    If Not CaptureStandaloneImage(rawPartNumber, normalizedPartNumber, sourcePath, imagePath, croppedPath, thumbPath, rowIndex) Then
                        If watchFirstImagePhase Then
                            If Not CheckTimedPhase("FIRST_IMAGE_STANDALONE_CAPTURE", firstImagePhaseStart, rowIndex, rawPartNumber, normalizedPartNumber, sourcePath) Then Exit Function
                        End If
                        SetRowUtilityValues rowIndex, imagePath, croppedPath, thumbPath, "STANDALONE_CAPTURE_FAILED", "Standalone open/capture failed"
                        AbortWithMessage "Standalone capture failed for Part Number: " & rawPartNumber, "STANDALONE_CAPTURE_FAILED", rowIndex, rawPartNumber, normalizedPartNumber, sourcePath
                        Exit Function
                    End If
                    If watchFirstImagePhase Then
                        If Not CheckTimedPhase("FIRST_IMAGE_STANDALONE_CAPTURE", firstImagePhaseStart, rowIndex, rawPartNumber, normalizedPartNumber, sourcePath) Then Exit Function
                    End If

                    SetRowUtilityValues rowIndex, imagePath, croppedPath, thumbPath, "OK", ""
                    CacheImage normalizedPartNumber, imagePath, croppedPath, thumbPath
                    If Not InsertThumbnailForRow(rowIndex, thumbPath) Then
                        AbortWithMessage "Excel ne moze da ubaci thumbnail za Part Number: " & rawPartNumber, "EXCEL_THUMBNAIL_INSERTED", rowIndex, rawPartNumber, normalizedPartNumber, sourcePath
                        Exit Function
                    End If
                    gProcessedImageRows = gProcessedImageRows + 1
                    gSuccessfulImageRows = gSuccessfulImageRows + 1
                    WriteExportLog rowIndex, rawPartNumber, normalizedPartNumber, "OK", "EXCEL_THUMBNAIL_INSERTED", "Thumbnail inserted.", imagePath, croppedPath, thumbPath, sourcePath
                    MaybeSaveByProgress rowIndex, rawPartNumber, normalizedPartNumber
                End If
                End If
                End If
            End If
        ElseIf RowHasAnyData(rowIndex) Then
            gRowsWithoutPartNumber = gRowsWithoutPartNumber + 1
            reason = "No Part Number in BOM row"
            SetRowUtilityValues rowIndex, "", "", "", "NO_PART_NUMBER", reason
            WriteExportLog rowIndex, rawPartNumber, normalizedPartNumber, "NO_PART_NUMBER", "NO_PART_NUMBER", reason, "", "", "", ""
            WriteDebugPhase "NO_PART_NUMBER", rowIndex, rawPartNumber, normalizedPartNumber, "", reason
        End If
        DoEvents
    Next

    ProcessBomRowsForThumbnails = True
    Err.Clear
End Function

Sub BuildCatiaFileIndex()
    On Error Resume Next
    CATIA.StatusBar = "Indexing source files: found 0 / " & CStr(gNeededPartNumbers.Count)
    If INDEX_ONLY_NEEDED_PART_NUMBERS Then
        If gNeededPartNumbers Is Nothing Then Exit Sub
        If gNeededPartNumbers.Count = 0 Then
            WriteDebugPhase "CATIA_FILE_INDEX_DONE", 0, "", "", "", "No source index needed; all required images are existing/reused or skipped."
            Exit Sub
        End If
    End If
    TraverseProductForSourceIndex gProduct
    Err.Clear
End Sub

Sub TraverseProductForSourceIndex(prod)
    On Error Resume Next
    If STOP_INDEX_SCAN_WHEN_ALL_FOUND And AllNeededPartNumbersFound() Then Exit Sub

    Dim rawPn
    Dim normalizedPn
    Dim sourcePath
    rawPn = GetProductPartNumber(prod)
    normalizedPn = NormalizePartNumber(rawPn)
    sourcePath = ""
    If ShouldIndexPartNumber(normalizedPn) Then
        sourcePath = GetProductSourceFilePath(prod)
        If rawPn <> "" Then WriteDebugPhase "CATIA_FILE_INDEX_ITEM", 0, rawPn, normalizedPn, sourcePath, "Raw ProductTree Part Number=" & rawPn & "; Normalized ProductTree Part Number=" & normalizedPn
        If normalizedPn <> "" And sourcePath <> "" Then
            If Not SamePath(sourcePath, gMainDocumentFullName) Then
                If Not gSourceIndex.Exists(normalizedPn) Then gSourceIndex.Item(normalizedPn) = sourcePath
                MarkNeededPartNumberFound normalizedPn, sourcePath
                CATIA.StatusBar = "Indexing source files: found " & CStr(gFoundNeededPartNumbers.Count) & " / " & CStr(gNeededPartNumbers.Count)
            End If
        End If
    End If
    If STOP_INDEX_SCAN_WHEN_ALL_FOUND And AllNeededPartNumbersFound() Then Exit Sub

    Dim children
    Dim i
    Set children = prod.Products
    If Err.Number <> 0 Or children Is Nothing Then
        Err.Clear
        Exit Sub
    End If
    For i = 1 To children.Count
        TraverseProductForSourceIndex children.Item(i)
        If STOP_INDEX_SCAN_WHEN_ALL_FOUND And AllNeededPartNumbersFound() Then Exit For
        If (i Mod 250) = 0 Then DoEvents
    Next
    Err.Clear
End Sub

Sub MarkNeededPartNumberFound(normalizedPn, sourcePath)
    On Error Resume Next
    Dim matchedKey
    matchedKey = MatchedNeededPartNumberKey(normalizedPn)
    If matchedKey <> "" Then
        If Not gSourceIndex.Exists(matchedKey) Then gSourceIndex.Item(matchedKey) = sourcePath
        If Not gFoundNeededPartNumbers.Exists(matchedKey) Then gFoundNeededPartNumbers.Item(matchedKey) = True
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
    If gNeededPartNumbers Is Nothing Then Exit Function
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

Function ShouldIndexPartNumber(normalizedPn)
    ShouldIndexPartNumber = False
    normalizedPn = NormalizePartNumber(normalizedPn)
    If normalizedPn = "" Then Exit Function
    If Not INDEX_ONLY_NEEDED_PART_NUMBERS Then
        ShouldIndexPartNumber = True
    ElseIf gNeededPartNumbers Is Nothing Then
        ShouldIndexPartNumber = True
    ElseIf gNeededPartNumbers.Count = 0 Then
        ShouldIndexPartNumber = False
    Else
        ShouldIndexPartNumber = (MatchedNeededPartNumberKey(normalizedPn) <> "")
    End If
End Function

Function AllNeededPartNumbersFound()
    On Error Resume Next
    AllNeededPartNumbersFound = False
    If gNeededPartNumbers Is Nothing Then Exit Function
    If gNeededPartNumbers.Count = 0 Then Exit Function
    If gFoundNeededPartNumbers Is Nothing Then Exit Function
    AllNeededPartNumbersFound = (gFoundNeededPartNumbers.Count >= gNeededPartNumbers.Count)
    Err.Clear
End Function

Function CaptureStandaloneImage(rawPartNumber, normalizedPartNumber, sourcePath, imagePath, croppedPath, thumbPath, excelRow)
    On Error Resume Next
    CaptureStandaloneImage = False
    If gFSO.FileExists(imagePath) Then gFSO.DeleteFile imagePath, True
    If gFSO.FileExists(croppedPath) Then gFSO.DeleteFile croppedPath, True
    If gFSO.FileExists(thumbPath) Then gFSO.DeleteFile thumbPath, True

    WriteDebugPhase "STANDALONE_OPEN_START", excelRow, rawPartNumber, normalizedPartNumber, sourcePath, sourcePath
    If Not OpenStandaloneDocument(sourcePath) Then
        WriteExportLog excelRow, rawPartNumber, normalizedPartNumber, "STANDALONE_OPEN_FAILED", "STANDALONE_OPEN_START", "Cannot open source file.", imagePath, croppedPath, thumbPath, sourcePath
        CloseCurrentStandaloneDocument
        ActivateMainDocument
        Exit Function
    End If
    WriteDebugPhase "STANDALONE_OPEN_DONE", excelRow, rawPartNumber, normalizedPartNumber, sourcePath, sourcePath

    WriteDebugPhase "STANDALONE_CAPTURE_START", excelRow, rawPartNumber, normalizedPartNumber, sourcePath, imagePath
    If Not CaptureActiveStandaloneViewer(imagePath) Then
        WriteExportLog excelRow, rawPartNumber, normalizedPartNumber, "STANDALONE_CAPTURE_FAILED", "STANDALONE_CAPTURE_START", "CaptureToFile failed.", imagePath, croppedPath, thumbPath, sourcePath
        CloseCurrentStandaloneDocument
        ActivateMainDocument
        Exit Function
    End If
    WriteDebugPhase "STANDALONE_CAPTURE_DONE", excelRow, rawPartNumber, normalizedPartNumber, sourcePath, imagePath

    If Not CreateCroppedImageFile(imagePath, croppedPath, excelRow, rawPartNumber, normalizedPartNumber, sourcePath) Then croppedPath = imagePath

    If Not CreateThumbnailFile(croppedPath, thumbPath, THUMBNAIL_WIDTH, THUMBNAIL_HEIGHT) Then
        WriteExportLog excelRow, rawPartNumber, normalizedPartNumber, "THUMBNAIL_FAILED", "THUMBNAIL_CREATED", "Thumbnail creation failed.", imagePath, croppedPath, thumbPath, sourcePath
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

Function CreateCroppedImageFile(imagePath, croppedPath, excelRow, rawPartNumber, normalizedPartNumber, sourcePath)
    On Error Resume Next
    CreateCroppedImageFile = False
    If imagePath = "" Or Not gFSO.FileExists(imagePath) Then Exit Function

    ' CATScript/VBScript has no stable built-in pixel scanner. Keep a safe cropped-file
    ' placeholder by copying the original, then create the Excel thumbnail from it.
    If gFSO.FileExists(croppedPath) Then gFSO.DeleteFile croppedPath, True
    gFSO.CopyFile imagePath, croppedPath, True
    If Err.Number = 0 And gFSO.FileExists(croppedPath) Then
        WriteDebugPhase "AUTO_CROP_FAILED_OR_SKIPPED", excelRow, rawPartNumber, normalizedPartNumber, sourcePath, "Auto-crop pixel scan not available in CATScript; original image copied to CROPPED."
        CreateCroppedImageFile = True
    Else
        WriteExportLog excelRow, rawPartNumber, normalizedPartNumber, "AUTO_CROP_FAILED_OR_SKIPPED", "AUTO_CROP_FAILED_OR_SKIPPED", "Could not create cropped fallback file.", imagePath, croppedPath, "", sourcePath
        WriteDebugPhase "AUTO_CROP_FAILED_OR_SKIPPED", excelRow, rawPartNumber, normalizedPartNumber, sourcePath, "Could not create cropped fallback file."
        Err.Clear
    End If
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

Sub SetRowUtilityValues(rowIndex, imagePath, croppedPath, thumbPath, statusText, skipReason)
    On Error Resume Next
    gWsBom.Cells(rowIndex, gImagePathColumnIndex).Value = imagePath
    gWsBom.Cells(rowIndex, gCroppedImagePathColumnIndex).Value = croppedPath
    gWsBom.Cells(rowIndex, gThumbnailPathColumnIndex).Value = thumbPath
    gWsBom.Cells(rowIndex, gExportStatusColumnIndex).Value = statusText
    gWsBom.Cells(rowIndex, gImageSkipReasonColumnIndex).Value = skipReason
    Err.Clear
End Sub

Sub WriteExportLog(rowIndex, rawPartNumber, normalizedPartNumber, statusText, phase, messageText, imagePath, croppedPath, thumbPath, sourcePath)
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
    gWsLog.Cells(gNextLogRow, 10).Value = croppedPath
    gWsLog.Cells(gNextLogRow, 11).Value = thumbPath
    gWsLog.Cells(gNextLogRow, 12).Value = sourcePath
    gNextLogRow = gNextLogRow + 1
    If statusText = "UNSUPPORTED_SOURCE_FILE" Or statusText = "STANDALONE_OPEN_FAILED" Or statusText = "STANDALONE_CAPTURE_FAILED" Or statusText = "THUMBNAIL_FAILED" Or statusText = "TIMEOUT" Or statusText = "ERROR" Then gErrorCount = gErrorCount + 1
    Err.Clear
End Sub

Sub WriteDebugPhase(phase, rowIndex, rawPartNumber, normalizedPartNumber, sourcePath, messageText)
    On Error Resume Next
    Dim ts
    Dim lineText
    lineText = BuildDebugLine(phase, rowIndex, rawPartNumber, normalizedPartNumber, sourcePath, messageText)
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

Function BuildDebugLine(phase, rowIndex, rawPartNumber, normalizedPartNumber, sourcePath, messageText)
    BuildDebugLine = FormatDateTime(Now, 2) & " " & FormatDateTime(Now, 3) & _
                     " | " & phase & _
                     " | ExcelRow=" & CStr(rowIndex) & _
                     " | Raw Part Number=" & CStr(rawPartNumber) & _
                     " | Normalized Part Number=" & CStr(normalizedPartNumber) & _
                     " | Source Path=" & CStr(sourcePath) & _
                     " | " & CStr(messageText)
End Function

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

Sub FinalizeExcelExport()
    On Error Resume Next
    If gWorkbook Is Nothing Then Exit Sub
    If gFinalExcelPath = "" Then gFinalExcelPath = gWorkExcelPath

    UpdateWorkbookPathsForFinalOutput
    UpdateSummarySheet
    gWorkbook.Save

    If Not SamePath(gWorkExcelPath, gFinalExcelPath) Then
        Err.Clear
        WriteDebugPhase "FINAL_SAVEAS_START", 0, "", "", gFinalExcelPath, gFinalExcelPath
        gExcelApp.DisplayAlerts = False
        gWorkbook.SaveAs gFinalExcelPath, XL_EXCEL8
        If Err.Number = 0 And gFSO.FileExists(gFinalExcelPath) Then
            gExcelPath = gFinalExcelPath
            WriteDebugPhase "FINAL_SAVEAS_DONE", 0, "", "", gFinalExcelPath, gFinalExcelPath
        Else
            Dim saveAsError
            saveAsError = Err.Description
            WriteDebugPhase "ERROR", 0, "", "", gFinalExcelPath, "Final SaveAs failed. Err.Number=" & CStr(Err.Number) & "; Err.Description=" & saveAsError
            Err.Clear
            WriteDebugPhase "FINAL_COPY_START", 0, "", "", gFinalExcelPath, "Copying local work XLS to final path."
            gWorkbook.Save
            gWorkbook.Close False
            Set gWorkbook = Nothing
            gFSO.CopyFile gWorkExcelPath, gFinalExcelPath, True
            If Err.Number = 0 And gFSO.FileExists(gFinalExcelPath) Then
                WriteDebugPhase "FINAL_COPY_DONE", 0, "", "", gFinalExcelPath, gFinalExcelPath
                Set gWorkbook = gExcelApp.Workbooks.Open(gFinalExcelPath)
                If Not gWorkbook Is Nothing Then
                    Set gWsBom = gWorkbook.Worksheets("BOM")
                    Set gWsLog = gWorkbook.Worksheets("EXPORT_LOG")
                    Set gWsSummary = gWorkbook.Worksheets("SUMMARY")
                    gExcelPath = gFinalExcelPath
                End If
            Else
                WriteDebugPhase "ERROR", 0, "", "", gFinalExcelPath, "Final CopyFile failed after SaveAs error. SaveAsError=" & saveAsError & "; Copy Err.Number=" & CStr(Err.Number) & "; Copy Err.Description=" & Err.Description
                Err.Clear
            End If
        End If
    End If

    CopyWorkOutputFilesToFinal
    If Not gWorkbook Is Nothing Then gWorkbook.Save
    gExcelApp.DisplayAlerts = True
    Err.Clear
End Sub

Sub CopySmokeTestWorkbookToFinal()
    On Error Resume Next
    If gFinalExcelPath = "" Then gFinalExcelPath = gWorkExcelPath
    If Not SamePath(gWorkExcelPath, gFinalExcelPath) Then
        WriteDebugPhase "FINAL_COPY_START", 0, "", "", gFinalExcelPath, "Smoke test copying local BOM XLS to final path."
        Err.Clear
        gFSO.CopyFile gWorkExcelPath, gFinalExcelPath, True
        If Err.Number = 0 And gFSO.FileExists(gFinalExcelPath) Then
            WriteDebugPhase "FINAL_COPY_DONE", 0, "", "", gFinalExcelPath, gFinalExcelPath
        Else
            WriteDebugPhase "ERROR", 0, "", "", gFinalExcelPath, "Smoke test final copy failed. Err.Number=" & CStr(Err.Number) & "; Err.Description=" & Err.Description
            Err.Clear
        End If
    End If
    CopyWorkOutputFilesToFinal
    CopyDebugLogToFinal
    Err.Clear
End Sub

Sub UpdateWorkbookPathsForFinalOutput()
    On Error Resume Next
    If gFinalOutputFolder = "" Or SamePath(gOutputFolder, gFinalOutputFolder) Then Exit Sub
    Dim rowIndex
    For rowIndex = gHeaderRow + 1 To gLastBomRow
        ReplaceCellPathPrefix rowIndex, gImagePathColumnIndex
        ReplaceCellPathPrefix rowIndex, gCroppedImagePathColumnIndex
        ReplaceCellPathPrefix rowIndex, gThumbnailPathColumnIndex
    Next
    UpdateSummarySheet
    Err.Clear
End Sub

Sub ReplaceCellPathPrefix(rowIndex, columnIndex)
    On Error Resume Next
    Dim oldValue
    oldValue = CStr(gWsBom.Cells(rowIndex, columnIndex).Value)
    If oldValue <> "" Then
        If UCase(Left(oldValue, Len(gOutputFolder))) = UCase(gOutputFolder) Then
            gWsBom.Cells(rowIndex, columnIndex).Value = gFinalOutputFolder & Mid(oldValue, Len(gOutputFolder) + 1)
        End If
    End If
    Err.Clear
End Sub

Sub CopyWorkOutputFilesToFinal()
    On Error Resume Next
    If gFinalOutputFolder = "" Or SamePath(gOutputFolder, gFinalOutputFolder) Then Exit Sub
    Err.Clear
    EnsureFolder gFinalOutputFolder
    EnsureFolder gFinalImageFolder
    EnsureFolder gFinalCroppedFolder
    EnsureFolder gFinalThumbnailFolder
    CopyFolderContents gImageFolder, gFinalImageFolder
    CopyFolderContents gCroppedFolder, gFinalCroppedFolder
    CopyFolderContents gThumbnailFolder, gFinalThumbnailFolder
    If Err.Number = 0 Then
        WriteDebugPhase "COPY_OUTPUT_FILES_DONE", 0, "", "", gFinalOutputFolder, gFinalOutputFolder
    Else
        WriteDebugPhase "COPY_OUTPUT_FILES_FAILED", 0, "", "", gFinalOutputFolder, "Err.Number=" & CStr(Err.Number) & "; Err.Description=" & Err.Description
        Err.Clear
    End If
    If gDebugLogPath <> "" And gFSO.FileExists(gDebugLogPath) Then gFSO.CopyFile gDebugLogPath, JoinPath(gFinalOutputFolder, "DEBUG_PHASE_LOG.txt"), True
End Sub

Sub CopyDebugLogToFinal()
    On Error Resume Next
    If gFinalOutputFolder = "" Or SamePath(gOutputFolder, gFinalOutputFolder) Then Exit Sub
    If gDebugLogPath <> "" And gFSO.FileExists(gDebugLogPath) Then gFSO.CopyFile gDebugLogPath, JoinPath(gFinalOutputFolder, "DEBUG_PHASE_LOG.txt"), True
    Err.Clear
End Sub

Sub CopyFolderContents(sourceFolder, targetFolder)
    On Error Resume Next
    If Not gFSO.FolderExists(sourceFolder) Then Exit Sub
    EnsureFolder targetFolder
    Dim fileObj
    For Each fileObj In gFSO.GetFolder(sourceFolder).Files
        gFSO.CopyFile fileObj.Path, JoinPath(targetFolder, fileObj.Name), True
    Next
    Err.Clear
End Sub

Function StartTimedPhase(phaseName, statusText)
    On Error Resume Next
    CATIA.StatusBar = statusText
    WriteDebugPhase phaseName, 0, "", "", "", "Phase started: " & statusText
    StartTimedPhase = Timer
    Err.Clear
End Function

Function CheckTimedPhase(phaseName, phaseStart, rowIndex, rawPartNumber, normalizedPartNumber, sourcePath)
    On Error Resume Next
    CheckTimedPhase = True
    If Not TEST_MODE Then Exit Function
    If CLng(TEST_PHASE_TIMEOUT_SECONDS) <= 0 Then Exit Function

    Dim elapsedSeconds
    elapsedSeconds = SecondsSince(phaseStart)
    If elapsedSeconds > CDbl(TEST_PHASE_TIMEOUT_SECONDS) Then
        HandlePhaseTimeout phaseName, elapsedSeconds, rowIndex, rawPartNumber, normalizedPartNumber, sourcePath
        CheckTimedPhase = False
    End If
    Err.Clear
End Function

Sub HandlePhaseTimeout(phaseName, elapsedSeconds, rowIndex, rawPartNumber, normalizedPartNumber, sourcePath)
    On Error Resume Next
    Dim messageText
    gAbortAlreadyHandled = True
    messageText = "TIMEOUT: faza '" & phaseName & "' traje duze od " & CStr(TEST_PHASE_TIMEOUT_SECONDS) & " sekundi. Proteklo: " & CStr(Round(elapsedSeconds, 1)) & " s."
    WriteDebugPhase "TIMEOUT", rowIndex, rawPartNumber, normalizedPartNumber, sourcePath, messageText
    If Not (gWsLog Is Nothing) Then WriteExportLog rowIndex, rawPartNumber, normalizedPartNumber, "TIMEOUT", phaseName, messageText, "", "", "", sourcePath
    SaveExcelCheckpoint phaseName, rowIndex, rawPartNumber, normalizedPartNumber
    CleanupCatiaSession
    MsgBox messageText, vbCritical, "CATIA_VISUAL_BOM_EXPORTER"
    Err.Clear
End Sub

Function SecondsSince(startValue)
    Dim nowValue
    nowValue = Timer
    If nowValue < CDbl(startValue) Then nowValue = nowValue + 86400
    SecondsSince = nowValue - CDbl(startValue)
End Function

Sub AbortWithMessage(messageText, phase, rowIndex, rawPartNumber, normalizedPartNumber, sourcePath)
    On Error Resume Next
    gAbortAlreadyHandled = True
    WriteDebugPhase "ERROR", rowIndex, rawPartNumber, normalizedPartNumber, sourcePath, phase & ": " & messageText
    WriteExportLog rowIndex, rawPartNumber, normalizedPartNumber, "ERROR", phase, messageText, "", "", "", sourcePath
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
    gWsSummary.Cells(5, 2).Value = gWorkExcelPath
    gWsSummary.Cells(6, 2).Value = gFinalExcelPath
    gWsSummary.Cells(7, 2).Value = gFinalOutputFolder
    gWsSummary.Cells(8, 2).Value = gLastBomRow - gHeaderRow
    If gNeededPartNumbers Is Nothing Then
        gWsSummary.Cells(9, 2).Value = 0
    Else
        gWsSummary.Cells(9, 2).Value = gNeededPartNumbers.Count
    End If
    gWsSummary.Cells(10, 2).Value = gRowsWithoutPartNumber
    gWsSummary.Cells(11, 2).Value = gProcessedImageRows
    gWsSummary.Cells(12, 2).Value = gSuccessfulImageRows
    gWsSummary.Cells(13, 2).Value = gReusedImageRows
    gWsSummary.Cells(14, 2).Value = gSkippedFastenerRows
    gWsSummary.Cells(15, 2).Value = gSourceNotFoundRows
    gWsSummary.Cells(16, 2).Value = gErrorCount
    gWsSummary.Cells(17, 2).Value = "TEST_MODE=" & CStr(TEST_MODE) & "; TEST_MAX_ROWS=" & CStr(TEST_MAX_ROWS) & "; STANDALONE_CAPTURE_ONLY=" & CStr(STANDALONE_CAPTURE_ONLY)
    gWsSummary.Cells(18, 2).Value = gDebugLogPath
    gWsSummary.Cells(19, 2).Value = FORCE_DEFAULT_BOM_COLUMNS
    gWsSummary.Cells(20, 2).Value = USE_LOCAL_WORK_FOLDER_FOR_CATIA_PRINT
    gWsSummary.Cells(21, 2).Value = TEST_MODE
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

Function ReuseExistingOrCachedImage(normalizedPartNumber, imagePath, croppedPath, thumbPath)
    On Error Resume Next
    ReuseExistingOrCachedImage = False
    Dim payload
    Dim parts
    If gImageCache.Exists(normalizedPartNumber) Then
        payload = CStr(gImageCache.Item(normalizedPartNumber))
        parts = Split(payload, "|")
        If UBound(parts) >= 2 Then
            imagePath = CStr(parts(0))
            croppedPath = CStr(parts(1))
            thumbPath = CStr(parts(2))
        End If
    End If
    If SKIP_EXISTING_IMAGES And gFSO.FileExists(imagePath) Then
        If Not gFSO.FileExists(croppedPath) Then
            Err.Clear
            gFSO.CopyFile imagePath, croppedPath, True
            Err.Clear
        End If
        If Not gFSO.FileExists(thumbPath) Then CreateThumbnailFile croppedPath, thumbPath, THUMBNAIL_WIDTH, THUMBNAIL_HEIGHT
        If gFSO.FileExists(thumbPath) Then
            CacheImage normalizedPartNumber, imagePath, croppedPath, thumbPath
            ReuseExistingOrCachedImage = True
        End If
    End If
    Err.Clear
End Function

Function ExistingImageSetAvailable(normalizedPartNumber, rowIndex)
    On Error Resume Next
    ExistingImageSetAvailable = False
    If Not SKIP_EXISTING_IMAGES Then Exit Function
    Dim imagePath
    Dim croppedPath
    Dim thumbPath
    imagePath = BuildImagePath(normalizedPartNumber, rowIndex)
    croppedPath = CroppedPathForImage(imagePath)
    thumbPath = ThumbnailPathForImage(imagePath)
    ExistingImageSetAvailable = gFSO.FileExists(imagePath)
    Err.Clear
End Function

Sub CacheImage(normalizedPartNumber, imagePath, croppedPath, thumbPath)
    If normalizedPartNumber <> "" Then gImageCache.Item(normalizedPartNumber) = imagePath & "|" & croppedPath & "|" & thumbPath
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

Function CroppedPathForImage(imagePath)
    CroppedPathForImage = JoinPath(gCroppedFolder, gFSO.GetBaseName(imagePath) & ".jpg")
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
        If c <> gThumbnailColumnIndex And c <> gImagePathColumnIndex And c <> gCroppedImagePathColumnIndex And c <> gThumbnailPathColumnIndex And c <> gExportStatusColumnIndex And c <> gImageSkipReasonColumnIndex Then
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

Function RowHasAnyData(rowIndex)
    On Error Resume Next
    RowHasAnyData = False
    Dim used
    Dim maxCol
    Dim c
    Set used = gWsBom.UsedRange
    maxCol = used.Column + used.Columns.Count - 1
    For c = 1 To maxCol
        If c <> gThumbnailColumnIndex And c <> gImagePathColumnIndex And c <> gCroppedImagePathColumnIndex And c <> gThumbnailPathColumnIndex And c <> gExportStatusColumnIndex And c <> gImageSkipReasonColumnIndex Then
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

Function EnsureXlsExtension(pathText)
    If LCase(Right(CStr(pathText), 4)) <> ".xls" Then
        EnsureXlsExtension = CStr(pathText) & ".xls"
    Else
        EnsureXlsExtension = CStr(pathText)
    End If
End Function

Function IsLocalWindowsPath(pathText)
    Dim s
    s = Trim(CStr(pathText))
    IsLocalWindowsPath = (Len(s) >= 3 And Mid(s, 2, 1) = ":" And (Mid(s, 3, 1) = "\" Or Mid(s, 3, 1) = "/"))
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

Function SafeFileNameOrDefault(valueText, defaultText)
    SafeFileNameOrDefault = SafeFileName(valueText)
    If SafeFileNameOrDefault = "" Then SafeFileNameOrDefault = SafeFileName(defaultText)
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
