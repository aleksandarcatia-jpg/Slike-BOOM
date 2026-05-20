CATIA_VISUAL_BOM_EXPORTER - native CATIA BOM + standalone thumbnails

Cilj:
CATIA native Bill of Material -> Excel -> thumbnail slike direktno u Excelu.

Fajlovi:
- CATIA_VISUAL_BOM_EXPORTER.CATScript
- CATIA_VISUAL_BOM_EXPORTER.bas

Makro sada koristi CATIA native BOM kao jedini izvor BOM podataka.
Ne rekonstruise BOM iz Product Tree-a, ne racuna Quantity, ne grupise redove i ne menja redosled redova.

Osnovni tok:
1. Otvorite glavni CATProduct u CATIA V5.
2. U CATIA podesite BOM kroz Analyze > Bill of Material > Define formats / Define Bill of Material.
3. Pokrenite makro.
4. Makro poziva Product.ExtractBOM i cita CATIA native BOM kolone i redove.
5. Excel dobija iste BOM kolone, plus:
   - Thumbnail
   - Image Path
   - Thumbnail Path
   - Export Status
   - Image Skip Reason
6. Product Tree se koristi samo za indeks:
   Part Number -> source CATPart/CATProduct path.
7. Part Number lookup se radi preko normalizovane vrednosti: skidaju se navodnici, TAB/CR/LF, zavrsni separatori kao zarez/semicolon/colon/pipe i visestruki razmaci.
8. Ako exact match ne uspe, makro pokusava match bez REV nastavka i zatim jedinstveni safe partial match.
9. Slike se prave samo otvaranjem CATPart/CATProduct fajla kao standalone dokument/prozor.
10. Makro zatvara samo dokumente koje je sam otvorio za slikanje i nikada ih ne snima.

Izlazni folder:
ROOT_PARTNUMBER_VISUAL_BOM_EXPORT

U njemu:
- VISUAL_BOM_EXPORT.xlsx
- IMAGES
- THUMBNAILS
- DEBUG_PHASE_LOG.txt

Podrazumevana podesavanja:
- TEST_MODE=True
- TEST_MAX_ROWS=20
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

Vazno:
- Hide/show metoda nije deo glavnog toka.
- Makro ne poziva SetShow / NoShow radi slikanja.
- Ako native CATIA BOM export ne uspe, makro se zaustavlja i ne pravi rucni BOM fallback.
- Ako Part Number iz BOM reda nema source CATPart/CATProduct fajl, makro sacuva Excel do tada i stane sa jasnom porukom.
- Ako je source fajl CGR ili drugi nepodrzan format, makro staje sa statusom UNSUPPORTED_SOURCE_FILE.
- Fastener redovi ostaju u Excelu, ali nemaju sliku i dobijaju status SKIPPED_IMAGE_ONLY.

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

TEST_MODE:
- Kada je TEST_MODE=True, Excel sadrzi ceo CATIA native BOM.
- Makro pravi slike samo za prvih TEST_MAX_ROWS ne-fastener redova.
- Fastener redovi ostaju bez slike.

FULL MODE:
- Podesite TEST_MODE=False.
- Makro obradjuje sve ne-fastener redove iz CATIA native BOM-a.

DEBUG_PHASE_LOG.txt belezi:
START, CATIA_NATIVE_BOM_EXPORT_START, CATIA_NATIVE_BOM_EXPORT_DONE,
BOM_HEADERS_READ, EXCEL_CREATED, CATIA_FILE_INDEX_START,
CATIA_FILE_INDEX_DONE, ROW_START, FASTENER_SKIPPED_IMAGE,
SOURCE_FILE_FOUND, SOURCE_FILE_NOT_FOUND, UNSUPPORTED_SOURCE_FILE,
STANDALONE_OPEN_START, STANDALONE_OPEN_DONE, STANDALONE_CAPTURE_START,
STANDALONE_CAPTURE_DONE, THUMBNAIL_CREATED, EXCEL_THUMBNAIL_INSERTED,
STANDALONE_CLOSE_DONE, SAVE_CHECKPOINT, ERROR, FINISH.
