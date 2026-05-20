CATIA_VISUAL_BOM_EXPORTER
=========================

Cilj:
CATIA Bill of Material export u Excel + thumbnail slike u istom Excel fajlu.

Fajlovi:
- CATIA_VISUAL_BOM_EXPORTER.CATScript
- CATIA_VISUAL_BOM_EXPORTER.bas

Glavni princip:
- CATIA BOM je izvor istine.
- Korisnik definise BOM u CATIA kroz Analyze > Bill of Material > Define Bill of Material / Define formats.
- Makro ne pravi BOM iz Product Tree-a.
- Makro ne racuna Quantity.
- Makro ne menja Part Number, Quantity, redosled redova ili postojece CATIA BOM vrednosti.
- Makro samo exportuje CATIA BOM u XLS, dodaje pomocne kolone i ubacuje thumbnail slike gde moze.

Osnovni tok:
1. Korisnik otvori glavni CATProduct.
2. Makro pita korisnika da izabere finalni folder za Excel.
3. Izbor foldera ide redom: Excel FileDialog folder picker, Shell BrowseForFolder, pa fallback InputBox.
4. Makro napravi finalnu XLS putanju:
   <RootPartNumber>_VISUAL_BOM_EXPORT.xls
5. Makro napravi lokalni work XLS u:
   C:\Temp\CATIA_VISUAL_BOM_WORK
6. Makro koristi CATIA BillOfMaterial objekat:
   product1.GetItem("BillOfMaterial")
7. Normalno ne poziva SetSecondaryFormat.
8. Makro poziva CATIA Print "XLS" samo na lokalni work fajl:
   assemblyConvertor1.Print "XLS", workExcelPath, product1
9. Makro otvara lokalni XLS:
   Excel.Workbooks.Open(workExcelPath)
10. Makro dodaje pomocne kolone:
   Thumbnail, Image Path, Cropped Image Path, Thumbnail Path, Export Status, Image Skip Reason.
11. Iz Excel BOM-a pravi listu potrebnih Part Number-a.
12. Product Tree se koristi samo za indeks:
   Part Number -> source CATPart/CATProduct path.
13. Za velike sklopove indeksira samo potrebne Part Number-e i prekida skeniranje kada ih nadje.
14. Slike se prave samo standalone metodom:
   CATIA.Documents.Open(sourcePath)
15. Isti Part Number se ne slika ponovo; koristi se image cache / postojeći thumbnail.
16. Na kraju se Excel snima/kopira na finalnu lokaciju koju je korisnik izabrao.
17. Excel ostaje otvoren korisniku.

Konfiguracija:
- FORCE_DEFAULT_BOM_COLUMNS=False
- USE_LOCAL_WORK_FOLDER_FOR_CATIA_PRINT=True
- SMOKE_TEST_BOM_ONLY=False
- DEBUG_FORCE_RECORDER_STYLE_PATH=False
- STOP_ON_SOURCE_NOT_FOUND=False
- TEST_MODE=True
- TEST_MAX_ROWS=20
- STANDALONE_CAPTURE_ONLY=True
- FALLBACK_TO_ASSEMBLY_HIDE_SHOW=False
- CLOSE_STANDALONE_DOCUMENT_AFTER_CAPTURE=True
- NEVER_SAVE_CATIA_DOCUMENTS=True
- SKIP_FASTENER_IMAGES=True
- SKIP_FASTENER_ROWS=False
- FAST_IMAGE_FORMAT="jpg"
- IMAGE_WIDTH=1000
- IMAGE_HEIGHT=750
- USE_WHITE_BACKGROUND=True
- USE_SHADED_WITH_EDGES=True
- USE_PARALLEL_PROJECTION=True
- IMAGE_CAPTURE_DELAY_SECONDS=0.20
- USE_AUTO_CROP_WHITE_BACKGROUND=True
- CROP_PADDING_PERCENT=0.10
- MIN_CROP_PADDING_PIXELS=25
- THUMBNAIL_WIDTH=160
- THUMBNAIL_HEIGHT=120
- SAVE_EVERY_N_ROWS=25
- RESUME_MODE=True
- SKIP_EXISTING_IMAGES=True
- STOP_INDEX_SCAN_WHEN_ALL_FOUND=True
- INDEX_ONLY_NEEDED_PART_NUMBERS=True
- MAX_SOURCE_NOT_FOUND_BEFORE_WARNING=20

FORCE_DEFAULT_BOM_COLUMNS:
- Default je False.
- Kada je False, makro koristi korisnicki definisan CATIA BOM format i ne poziva SetSecondaryFormat.
- Kada je True, makro postavlja default recorder kolone:
  Nomenclature, Quantity, Part Number, Dimenzija, Material, Mass, Standard.

SMOKE_TEST_BOM_ONLY:
- Default je False.
- Kada je True, makro samo exportuje BOM XLS preko BillOfMaterial.Print "XLS".
- Ne otvara partove, ne pravi slike i ne pravi Product Tree index.
- Ako XLS postoji, prikazuje:
  BOM XLS export uspesan

DEBUG_FORCE_RECORDER_STYLE_PATH:
- Default je False.
- Kada je True, makro koristi lokalni dijagnosticki fajl:
  C:\Temp\CATIA_VISUAL_BOM_WORK\CATIA_BOM_TEST.xls
- Ovo sluzi samo za dijagnostiku CATIA Print "XLS" problema.

Folder i fajlovi:
Ako korisnik izabere folder:
  D:\Project

Finalni Excel bude:
  D:\Project\3260.24.00.00_VISUAL_BOM_EXPORT.xls

CATIA Print "XLS" se prvo radi lokalno:
  C:\Temp\CATIA_VISUAL_BOM_WORK\3260.24.00.00_VISUAL_BOM_EXPORT_WORK.xls

Lokalni work folder je:
  C:\Temp\CATIA_VISUAL_BOM_WORK\3260.24.00.00_VISUAL_BOM_EXPORT_FILES

Na kraju makro kopira fajlove u finalni folder:
  D:\Project\3260.24.00.00_VISUAL_BOM_EXPORT_FILES

U njemu:
- IMAGES
- CROPPED
- THUMBNAILS
- DEBUG_PHASE_LOG.txt

Redovi bez Part Number-a:
- Makro ih ne preskace i ne zaustavlja export.
- Export Status = NO_PART_NUMBER
- Image Skip Reason = No Part Number in BOM row

Source file nije pronadjen:
- Default STOP_ON_SOURCE_NOT_FOUND=False.
- Red ostaje u Excelu.
- Export Status = SOURCE_FILE_NOT_FOUND
- Image Skip Reason = Source CATPart/CATProduct not found for Part Number
- Makro nastavlja dalje.

Fastener redovi:
- Vijci, navrtke, podloske i slicni standardni elementi ostaju u Excelu.
- Za njih se ne pravi slika.
- Export Status = SKIPPED_IMAGE_ONLY
- Image Skip Reason = Standard fastener - retained in BOM, image skipped

Crop:
- CATScript/VBScript nema stabilan built-in pixel scanner.
- Makro bezbedno kreira fajl u CROPPED folderu kopiranjem originalne slike.
- Thumbnail se pravi iz CROPPED fajla.
- DEBUG log dobija AUTO_CROP_FAILED_OR_SKIPPED kada je pixel auto-crop preskocen.

Optimizacija za velike sklopove:
- Makro prvo cita Part Number-e iz CATIA BOM Excel-a.
- Fastener redovi i redovi bez Part Number-a se ne ubacuju u listu za slike.
- U TEST_MODE obradjuje se samo prvih TEST_MAX_ROWS kandidata za slike.
- Product Tree scan trazi samo potrebne Part Number-e.
- Kada nadje sve potrebne source fajlove, prekida dalji scan.
- Image cache sprecava ponovno otvaranje i slikanje istog Part Number-a.
- Ako image/cropped/thumbnail vec postoje, koriste se ponovo.

DEBUG_PHASE_LOG.txt belezi faze:
START
SMOKE_TEST_BOM_ONLY
USER_FOLDER_DIALOG_OPENED
USER_FOLDER_SELECTED
USER_FOLDER_CANCELLED
FINAL_XLS_PATH_PREPARED
LOCAL_WORK_FOLDER_PREPARED
LOCAL_WORK_XLS_PATH_PREPARED
OUTPUT_XLS_EXISTS
OUTPUT_XLS_OVERWRITE_CONFIRMED
OUTPUT_XLS_TIMESTAMP_CREATED
BOM_PRINT_PATH_CHECK
BOM_FORMAT_USER_DEFINED
BOM_FORMAT_SET_START
BOM_FORMAT_SET_DONE
CATIA_BOM_PRINT_XLS_START
CATIA_BOM_PRINT_XLS_DONE
CATIA_BOM_PRINT_XLS_FAILED
EXCEL_OPENED_LOCAL_WORKBOOK
BOM_HEADERS_READ
HELPER_COLUMNS_ADDED
NEEDED_PART_NUMBERS_BUILT
CATIA_FILE_INDEX_START
CATIA_FILE_INDEX_ITEM
CATIA_FILE_INDEX_DONE
ROW_START
NO_PART_NUMBER
FASTENER_SKIPPED_IMAGE
SOURCE_FILE_FOUND
SOURCE_FILE_NOT_FOUND
STANDALONE_OPEN_START
STANDALONE_OPEN_DONE
STANDALONE_CAPTURE_START
STANDALONE_CAPTURE_DONE
AUTO_CROP_DONE
AUTO_CROP_FAILED_OR_SKIPPED
THUMBNAIL_CREATED
EXCEL_THUMBNAIL_INSERTED
EXISTING_REUSED
STANDALONE_CLOSE_DONE
SAVE_CHECKPOINT
FINAL_SAVEAS_START
FINAL_SAVEAS_DONE
FINAL_COPY_START
FINAL_COPY_DONE
COPY_OUTPUT_FILES_DONE
COPY_OUTPUT_FILES_FAILED
ERROR
FINISH

Zabranjeno u glavnom toku:
- CATIA TXT/HTML BOM export kao izvor BOM-a
- TXT/HTML parser
- rucno pravljenje BOM-a iz Product Tree-a
- rucno racunanje Quantity-ja
- hide/show image capture
- assembly visibility fallback za image workflow

Kako pokrenuti CATScript:
1. Tools > Macro > Macros...
2. Macro libraries... / Select...
3. Izaberite Directory kao tip biblioteke.
4. Pokazite na folder gde je CATIA_VISUAL_BOM_EXPORTER.CATScript.
5. Izaberite makro i Run.

Kako pokrenuti VBA verziju:
1. Tools > Macro > Macros...
2. Izaberite ili napravite VBA library (.catvba).
3. Otvorite VBA Editor.
4. Insert > Module.
5. Iskopirajte kompletan sadrzaj CATIA_VISUAL_BOM_EXPORTER.bas.
6. Pokrenite CATIA_VISUAL_BOM_EXPORTER.
