CATIA_VISUAL_BOM_EXPORTER
=========================

Cilj:
CATIA Bill of Material export u Excel + thumbnail slike delova u istom Excel fajlu.

Fajlovi:
- CATIA_VISUAL_BOM_EXPORTER.CATScript
- CATIA_VISUAL_BOM_EXPORTER.bas

Osnovna logika:
1. Korisnik otvori glavni CATProduct u CATIA V5.
2. Makro proveri da je aktivni dokument CATProduct.
3. Makro uzme CATIA BOM objekat:
   product1.GetItem("BillOfMaterial")
4. Makro podesi BOM format:
   Nomenclature, Quantity, Part Number, Dimenzija, Material, Mass, Standard.
5. Makro pita korisnika gde zeli da sacuva Excel preko Save As dialoga.
6. Makro eksportuje CATIA BOM direktno u XLS:
   assemblyConvertor1.Print "XLS", selectedPath, product1
7. Makro otvara taj isti Excel fajl i dodaje samo pomocne kolone:
   Thumbnail, Image Path, Thumbnail Path, Export Status, Image Skip Reason.
8. Product Tree se koristi samo za indeks:
   Part Number -> source CATPart/CATProduct path.
9. Za svaki BOM red kome treba slika, makro otvara source CATPart/CATProduct kao standalone dokument/prozor.
10. Makro napravi JPG sliku, napravi mali thumbnail i ubaci thumbnail u Excel.
11. Makro zatvara samo dokument koji je sam otvorio za slikanje, bez snimanja.
12. Excel ostaje otvoren korisniku.

Sta makro NE radi:
- Ne pravi BOM iz Product Tree-a.
- Ne racuna Quantity.
- Ne menja redosled redova.
- Ne menja postojece CATIA BOM vrednosti.
- Ne koristi hide/show za slike.
- Ne koristi CATIA hide/show metodu za slike.
- Ne koristi fallback na assembly visibility capture.
- Ne snima CATPart/CATProduct dokumente.
- Ne zatvara glavni CATProduct.

Save As:
- Default naziv je:
  <RootPartNumber>_VISUAL_BOM_EXPORT.xls
- Default folder je folder aktivnog CATProduct-a, ako postoji.
- Ako aktivni CATProduct nema folder, default je Desktop.
- Save As dialog ceka korisnika neograniceno dugo i nije deo timeout merenja.
- Ako korisnik klikne Cancel, makro prekida rad i prikazuje:
  Export je otkazan od strane korisnika.

Output folder:
Ako korisnik izabere:
  D:\Project\MyBOM.xls

Makro pravi folder:
  D:\Project\MyBOM_FILES

U njemu:
- IMAGES
- THUMBNAILS
- DEBUG_PHASE_LOG.txt

Excel struktura:
- Prvi sheet je CATIA BOM sheet.
- Postojece CATIA BOM kolone ostaju netaknute.
- Thumbnail kolona se dodaje odmah posle Part Number kolone.
- Na kraj se dodaju:
  Image Path, Thumbnail Path, Export Status, Image Skip Reason.
- Dodaju se i sheetovi:
  EXPORT_LOG
  SUMMARY

Part Number lookup:
- Makro normalizuje Part Number iz Excel BOM-a i iz Product Tree index-a.
- Normalizacija radi:
  trim razmaka, uklanjanje navodnika, TAB/CR/LF, zavrsnih separatora
  comma, semicolon, colon, pipe, visestrukih razmaka i uppercase.
- Primer:
  "2417.01.14.01.04.01," -> "2417.01.14.01.04.01"
- Ako exact match ne uspe, makro pokusava match bez REV nastavka.
- Ako ni to ne uspe, pokusava safe unique partial match samo kada postoji jedan kandidat.

Slike:
- Slike se prave samo standalone metodom:
  CATIA.Documents.Open(sourcePath)
  standaloneDoc.Activate
  viewer.Reframe
  viewer.CaptureToFile 2, imagePath
  standaloneDoc.Close
- Full slike idu u IMAGES folder.
- Male thumbnail slike idu u THUMBNAILS folder.
- Excel ubacuje samo thumbnail sliku, ne full-size sliku.

Fastener elementi:
- Vijci, zavrtnji, navrtke/matice, podloske i slicni standardni elementi ostaju u BOM-u.
- Za njih se ne pravi slika.
- Export Status:
  SKIPPED_IMAGE_ONLY
- Image Skip Reason:
  Standard fastener - retained in BOM, image skipped

Default konfiguracija:
- TEST_MODE=True
- TEST_MAX_ROWS=20
- TEST_PHASE_TIMEOUT_SECONDS=30
- STANDALONE_CAPTURE_ONLY=True
- PREFER_STANDALONE_PART_CAPTURE=True
- FALLBACK_TO_ASSEMBLY_HIDE_SHOW=False
- CLOSE_STANDALONE_DOCUMENT_AFTER_CAPTURE=True
- NEVER_SAVE_CATIA_DOCUMENTS=True
- SKIP_FASTENER_IMAGES=True
- SKIP_FASTENER_ROWS=False
- FAST_IMAGE_FORMAT="jpg"
- IMAGE_WIDTH=1000
- IMAGE_HEIGHT=750
- THUMBNAIL_WIDTH=160
- THUMBNAIL_HEIGHT=120
- SAVE_EVERY_N_ROWS=25
- RESUME_MODE=True
- SKIP_EXISTING_IMAGES=True

TEST_MODE:
- CATIA eksportuje kompletan BOM Excel.
- Svi redovi ostaju u Excelu.
- Makro dodaje slike samo za prvih TEST_MAX_ROWS ne-fastener redova.
- Ostali neobradjeni redovi dobijaju:
  NOT_PROCESSED_TEST_LIMIT
- U TEST_MODE faze imaju watchdog od TEST_PHASE_TIMEOUT_SECONDS sekundi.
- Ako faza predje limit, DEBUG_PHASE_LOG.txt dobija TIMEOUT i makro prikazuje fazu koja je predugo trajala.

StatusBar kontrolne tacke:
- 1/6 Biranje Excel lokacije...
- 2/6 Exportujem CATIA BOM u Excel...
- 3/6 BOM Excel napravljen.
- 4/6 Indeksiram CATIA partove...
- 5/6 Otvaram prvi part za sliku...
- 6/6 Gotovo.

FULL MODE:
- Podesiti:
  Const TEST_MODE = False
- Makro obradjuje sve ne-fastener redove iz CATIA BOM-a.

DEBUG_PHASE_LOG.txt belezi:
START
BOM_FORMAT_SET
USER_SAVE_PATH_DIALOG_OPENED
USER_SAVE_PATH_SELECTED
USER_SAVE_CANCELLED
CATIA_BOM_PRINT_XLS_START
CATIA_BOM_PRINT_XLS_DONE
EXCEL_OPENED
BOM_HEADERS_READ
CATIA_FILE_INDEX_START
CATIA_FILE_INDEX_ITEM
CATIA_FILE_INDEX_DONE
ROW_START
FASTENER_SKIPPED_IMAGE
SOURCE_FILE_FOUND
SOURCE_FILE_NOT_FOUND
STANDALONE_OPEN_START
STANDALONE_OPEN_DONE
STANDALONE_CAPTURE_START
STANDALONE_CAPTURE_DONE
THUMBNAIL_CREATED
EXCEL_THUMBNAIL_INSERTED
STANDALONE_CLOSE_DONE
SAVE_CHECKPOINT
ERROR
TIMEOUT
FINISH

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
