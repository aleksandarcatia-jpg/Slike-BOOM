CATIA_VISUAL_BOM_EXPORTER - production visual BOM makro

Cilj:
CATIA CATProduct -> grupisani BOM Excel -> slike delova -> thumbnail slike direktno u Excelu.

Fajlovi:
- CATIA_VISUAL_BOM_EXPORTER.CATScript
- CATIA_VISUAL_BOM_EXPORTER.bas

Makro ne menja geometriju, ne snima CATProduct i ne zatvara CATIA.

Kako pokrenuti CATScript:
1. Otvorite glavni CATProduct u CATIA V5.
2. Tools > Macro > Macros...
3. Macro libraries... / Select...
4. Izaberite Directory kao tip biblioteke.
5. Pokazite na folder gde su fajlovi makroa.
6. Izaberite CATIA_VISUAL_BOM_EXPORTER.CATScript.
7. Run.

Kako pokrenuti VBA verziju:
1. Tools > Macro > Macros...
2. Izaberite ili napravite VBA library (.catvba).
3. Otvorite VBA Editor.
4. Insert > Module.
5. Iskopirajte kompletan sadrzaj CATIA_VISUAL_BOM_EXPORTER.bas.
6. Pokrenite CATIA_VISUAL_BOM_EXPORTER.

Izlazni folder:
<PartNumber>_VISUAL_BOM_EXPORT

U njemu se prave:
- VISUAL_BOM_EXPORT.xlsx
- IMAGES
- THUMBNAILS
- DEBUG_PHASE_LOG.txt
- MAIN_ASSEMBLY.jpg samo ako je EXPORT_MAIN_ASSEMBLY_IMAGE=True

Podrazumevana podesavanja za prvu proveru:
- TEST_MODE=True
- TEST_MAX_ITEMS=50
- STOP_SCAN_AFTER_TEST_ITEMS=True
- EXPORT_MAIN_ASSEMBLY_IMAGE=False
- HYBRID_PRODUCTION_MODE=True
- INSERT_IMAGES_IN_EXCEL=True
- CREATE_THUMBNAIL_FILES=True
- INSERT_THUMBNAIL_FILE_IN_EXCEL=True
- FAST_SNIP_MODE=True
- FAST_IMAGE_FORMAT="jpg"
- FAST_ZOOM_IN_STEPS=0
- FAST_CENTER_CROP_PERCENT=0.90
- RESUME_MODE=True
- SKIP_EXISTING_IMAGES=True
- SKIP_FASTENER_IMAGES=True
- SKIP_FASTENER_ROWS=False
- START_INDEX=1
- END_INDEX=0

Za full export:
1. U kodu podesite TEST_MODE=False.
2. Ostavite MAX_ITEMS_TO_EXPORT=0.
3. Ostavite MAX_BOM_SCAN_SECONDS=0 ako zelite da veliki sklop skenira koliko god treba.
4. Po potrebi koristite START_INDEX / END_INDEX za batch rad, npr. 1-200, 201-400.

Vazno za velike sklopove:
- Makro grupise po Part Number-u.
- Jedna BOM stavka dobija jednu sliku, bez slikanja svih instanci.
- U TEST_MODE=True makro zaustavlja BOM scan cim sakupi TEST_MAX_ITEMS jedinstvenih stavki, da ne cita ceo veliki sklop.
- MAX_BOM_SCAN_SECONDS=0 znaci bez timeout prekida. Ako je vece od 0, upisuje se WARNING, ali full export se ne prekida automatski.
- Quantity je broj pojavljivanja istog Part Number-a.
- Ako Part Number ne postoji, koristi se TreePath fallback.
- Excel se kreira odmah i snima checkpoint posle prve slike, posle prvih 10 slika i zatim periodicno.
- Ako slika vec postoji u IMAGES i thumbnail u THUMBNAILS, makro ih koristi ponovo.
- Ako postoji full slika bez thumbnail-a, makro pokusava da napravi thumbnail i nastavi.
- Standardni spojni elementi kao vijci, zavrtnji, matice/navrtke i podloske ostaju u BOM-u, ali dobijaju status SKIPPED_FASTENER i nemaju thumbnail kada je SKIP_FASTENER_IMAGES=True.

Excel:
- Sheet SUMMARY: osnovni podaci, folder, debug log.
- Sheet BOM: kolone iz CATIA Analyze > Bill of Material ako ExtractBOM uspe, plus Thumbnail, Image Path i Export Status.
- Sheet EXPORT_LOG: No., Date/Time, Item Index, Part Number, Status, Phase, Message, Image Path, Thumbnail Path, ISO_VIEW.

Debug:
DEBUG_PHASE_LOG.txt belezi faze:
START, EXCEL_CREATED, BOM_SCAN_START, BOM_SCAN_PROGRESS, BOM_SCAN_DONE,
UNIQUE_ITEMS_COUNT, MAIN_ASSEMBLY_SKIPPED, IMAGE_EXPORT_START, ITEM_START,
ITEM_VISIBLE_SET, ITEM_CAPTURE_START, ITEM_CAPTURE_DONE, THUMBNAIL_CREATED,
EXCEL_THUMBNAIL_INSERTED, SAVE_CHECKPOINT, ITEM_ERROR, ITEM_TIMEOUT, FINISH.

Napomene:
- CATIA CaptureToFile je COM poziv. Ako se sama CATIA potpuno blokira u tom pozivu, VBScript/VBA ne moze nasilno da ga prekine dok CATIA ne vrati kontrolu makrou.
- Hide/show zavisi od CATIA V5 konfiguracije, loaded/unloaded komponenti i CGR/lightweight stanja. Makro koristi batch Selection.VisProperties.SetShow i na kraju poziva SafeRestoreCatiaSession.
- Za stabilnije slike koristite belu pozadinu i shaded-with-edges prikaz u CATIA vieweru.
- JPG je podrazumevan u FAST_SNIP_MODE jer je brzi i dovoljan za Excel thumbnail.
