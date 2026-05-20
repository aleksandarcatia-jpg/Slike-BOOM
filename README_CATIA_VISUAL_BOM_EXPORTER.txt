CATIA_VISUAL_BOM_EXPORTER
=========================

Cilj:
CATIA CATProduct model workflow -> korisnicki snimljen BOM Excel -> thumbnail slike u istom Excel fajlu.

Fajlovi:
- CATIA_VISUAL_BOM_EXPORTER.CATScript
- CATIA_VISUAL_BOM_EXPORTER.bas

Glavni princip:
- CATIA BOM je izvor istine.
- Korisnik sam definise BOM u CATIA kroz Analyze > Bill of Material > Define Bill of Material / Define formats.
- Korisnik sam snima BOM kao Excel iz CATIA.
- Makro ne pravi BOM iz Product Tree-a.
- Makro ne racuna Quantity.
- Makro ne menja Part Number, Quantity, redosled redova ili postojece BOM vrednosti.
- Makro samo dodaje pomocne kolone i thumbnail slike.

Dvofazni tok
------------

PHASE 1:
1. Otvorite glavni CATProduct.
2. Pokrenite CATIA_VISUAL_BOM_EXPORTER.
3. Makro proverava CATProduct i pravi state fajl:
   C:\Temp\CATIA_VISUAL_BOM_WORK\VISUAL_BOM_STATE.txt
4. Makro prikazuje instrukciju.
5. U CATIA rucno otvorite Analyze > Bill of Material.
6. Definisite BOM kako zelite.
7. Snimite BOM kao Excel.
8. Makro se ne zadrzava i ne ceka Save As prozor.

PHASE 2:
1. Ponovo pokrenite isti makro.
2. Makro detektuje state fajl.
3. Makro pita da izaberete BOM Excel koji ste rucno snimili.
4. Makro pita za finalni folder.
5. Makro kopira izabrani BOM Excel u lokalni work folder.
6. Makro pronalazi Part Number kolonu.
7. Makro koristi postojecu Slika/Image/Thumbnail/Picture/Foto/Preview kolonu ili dodaje Thumbnail.
8. Makro indeksira otvoreni CATProduct samo za Part Number -> source CATPart/CATProduct path.
9. Makro otvara source fajlove standalone, pravi JPG slike i ubacuje ih u Excel.
10. Makro periodicki snima Excel.
11. Makro snima finalni Excel kao:
    <OriginalExcelName>_WITH_IMAGES.xls
12. Ako PHASE 2 uspe, state fajl se brise.
13. Ako PHASE 2 ne uspe, state fajl ostaje da mozete ponoviti drugi korak.

Work folder
-----------

Sav rad ide lokalno:

C:\Temp\CATIA_VISUAL_BOM_WORK

Podfolderi:
- WORKBOOK
- IMAGES
- THUMBNAILS
- DEBUG_PHASE_LOG.txt
- VISUAL_BOM_STATE.txt

Lokalna kopija workbooka:

C:\Temp\CATIA_VISUAL_BOM_WORK\WORKBOOK\<OriginalExcelName>_WORK.xls

Konfiguracija
-------------

Default vrednosti:
- SAVE_EVERY_N_ROWS=25
- SKIP_EXISTING_IMAGES=True
- RESUME_MODE=True
- TEST_MODE=True
- TEST_MAX_ROWS=20
- SKIP_FASTENER_IMAGES=True
- IMAGE_WIDTH=1000
- IMAGE_HEIGHT=750
- THUMBNAIL_WIDTH=160
- THUMBNAIL_HEIGHT=120
- STANDALONE_CAPTURE_ONLY=True
- FALLBACK_TO_ASSEMBLY_HIDE_SHOW=False
- CLOSE_STANDALONE_DOCUMENT_AFTER_CAPTURE=True
- NEVER_SAVE_CATIA_DOCUMENTS=True

TEST_MODE:
- Kada je True, makro pravi slike samo za prvih 20 redova koji imaju Part Number i nisu fastener.
- Ostali redovi dobijaju status NOT_PROCESSED_TEST_LIMIT.
- Kada je False, obradjuje se ceo BOM.

Excel pravila
-------------

Makro ne menja postojece BOM podatke.

Makro sme da doda/popuni samo:
- Thumbnail ili postojecu Slika/Image/Thumbnail/Picture/Foto/Preview kolonu
- Image Path
- Export Status
- Image Skip Reason

Part Number kolone koje se prepoznaju:
- Part Number
- PartNumber
- Part No.
- Part No
- PartNo
- Part-No
- PN
- P/N
- Number
- Broj dela
- Br. dela
- Oznaka
- Pozicija-DoN
- Position-DoN
- Item-DoN
- Drawing No
- Drawing Number
- Drw. No
- Drw No
- Broj crteza
- Br. crteza

Ako postoji vise mogucih Part Number kolona, prioritet je:
1. PARTNUMBER / PARTNO / PN
2. DRAWINGNO / DRWNO / BROJCRTEZA
3. OZNAKA / POZICIJA / ITEM
4. NUMBER

Ako Part Number kolona nije pronadjena:
- makro ne koristi prvu kolonu kao fallback
- slikanje se prekida jasnom porukom
- DEBUG_PHASE_LOG.txt sadrzi raw i normalized header-e

Redovi bez Part Number-a:
- ostaju u Excelu
- Export Status = NO_PART_NUMBER
- Image Skip Reason = No Part Number in BOM row

Fastener redovi:
- ostaju u Excelu
- slika se ne pravi
- Export Status = SKIPPED_FASTENER
- Image Skip Reason = Fastener image skipped

Source file nije pronadjen:
- red ostaje u Excelu
- Export Status = SOURCE_FILE_NOT_FOUND
- makro nastavlja dalje

Slikanje
--------

Slike se prave samo standalone metodom:
1. CATIA.Documents.Open(sourcePath)
2. Activate
3. ActiveViewer.Reframe
4. izometrija ako CATIA dozvoli
5. CaptureToFile 2, imagePath
6. thumbnail u Excel celiju
7. zatvaranje otvorenog dokumenta bez snimanja

Ako slikanje ne uspe:
- Export Status = IMAGE_CAPTURE_FAILED
- makro nastavlja dalje

Optimizacija za velike sklopove
-------------------------------

- Makro prvo cita Part Number-e iz izabranog BOM Excela.
- Pravi unique NeededPartNumbers.
- Preskace prazne Part Number redove i fastenere.
- Skenira CATProduct samo jednom.
- Ne otvara isti Part Number vise puta.
- Ne slika isti Part Number vise puta.
- Koristi image cache.
- Ako se isti Part Number ponovi, koristi postojeci thumbnail.
- Export Status = EXISTING_REUSED
- Excel se snima na svakih 25 redova.

Source izbor za slike
---------------------

Makro cuva kandidate iz Product Tree-a po normalized Part Number-u:
- raw Part Number
- normalized Part Number
- source path
- ekstenzija CATPart/CATProduct
- da li je root assembly
- da li ima children
- depth
- product name

Pravila izbora:
- exact normalized match ima prioritet
- zatim exact match bez REV nastavka
- safe unique partial match samo ako je jedinstven i dovoljno siguran
- CATPart se bira pre CATProduct kada postoji
- root assembly se ne koristi kao slika za red
- CATProduct se koristi samo kao podsklop ako nema CATPart kandidata i ako je dozvoljeno

Relevantne konfiguracije:
- PREFER_CATPART_OVER_CATPRODUCT=True
- ALLOW_ROOT_ASSEMBLY_IMAGE=False
- ALLOW_SUBASSEMBLY_IMAGE=True

DEBUG log za svaki source match pise:
- raw i normalized BOM Part Number
- matched ProductTree Part Number
- source extension
- CATPart/CATProduct
- root assembly
- has children
- match method: EXACT, NO_REV ili SAFE_UNIQUE_PARTIAL

DEBUG_PHASE_LOG.txt faze
------------------------

START
PHASE1_PREPARED
STATE_FILE_WRITTEN
BILL_OF_MATERIAL_COMMAND_STARTED_OR_SKIPPED
PHASE2_STARTED
EXISTING_BOM_EXCEL_SELECTED
FINAL_FOLDER_SELECTED
LOCAL_WORKBOOK_CREATED
EXCEL_OPENED
PART_NUMBER_COLUMN_FOUND
IMAGE_COLUMN_FOUND_OR_CREATED
NEEDED_PART_NUMBERS_BUILT
CATIA_FILE_INDEX_START
CATIA_FILE_INDEX_DONE
ROW_START
SOURCE_FILE_FOUND
SOURCE_FILE_NOT_FOUND
FASTENER_SKIPPED
STANDALONE_OPEN_START
STANDALONE_CAPTURE_DONE
IMAGE_INSERTED_IN_EXCEL
EXISTING_REUSED
SAVE_CHECKPOINT
FINAL_SAVE_DONE
STATE_FILE_CLEARED
ERROR
FINISH

Kako pokrenuti CATScript
------------------------

1. Tools > Macro > Macros...
2. Macro libraries... / Select...
3. Izaberite Directory kao tip biblioteke.
4. Pokazite na folder gde je CATIA_VISUAL_BOM_EXPORTER.CATScript.
5. Izaberite makro i Run.

Kako pokrenuti VBA verziju
--------------------------

1. Tools > Macro > Macros...
2. Izaberite ili napravite VBA library (.catvba).
3. Otvorite VBA Editor.
4. Insert > Module.
5. Iskopirajte kompletan sadrzaj CATIA_VISUAL_BOM_EXPORTER.bas.
6. Pokrenite CATIA_VISUAL_BOM_EXPORTER.
