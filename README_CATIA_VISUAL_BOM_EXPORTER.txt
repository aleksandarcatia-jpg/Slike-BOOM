CATIA_VISUAL_BOM_EXPORTER - native BOM + standalone thumbnail workflow

VAZNO - koordinacija GPT / Codex / GitHub:
Ova grana menja postojeci CATIA_VISUAL_BOM_EXPORTER.CATScript prema poslednjem dogovoru sa Aleksandrom.
Postojeca main verzija se ne dira dok se PR ne pregleda/testira.
Codex treba da vidi ovaj README i PR pre daljih izmena.

Cilj:
CATIA native Bill of Material -> Excel -> thumbnail slike direktno u Excelu.

Dogovorena logika:
1. Korisnik u CATIA V5 otvori glavni CATProduct.
2. Korisnik u CATIA definise BOM kolone kroz:
   Analyze -> Bill of Material -> Define formats / Define Bill of Material
3. Makro koristi CATIA native BOM kao izvor podataka.
4. Makro ne rekonstruiše BOM iz Product Tree-a kao primarni izvor.
5. Makro ne menja Quantity, Part Number, Nomenclature, redosled ni BOM kolone.
6. Makro doda samo Thumbnail, Image Path, Thumbnail Path, Export Status i Image Skip Reason.
7. Product Tree se koristi samo za indeks:
   Part Number -> source CATPart/CATProduct file path
8. Slike se prave samo standalone metodom:
   otvori CATPart/CATProduct kao poseban dokument/prozor,
   namesti izometriju,
   uradi capture,
   napravi thumbnail,
   ubaci thumbnail u Excel,
   zatvori otvoreni dokument bez snimanja.
9. Hide/show metoda za slikanje je izbacena iz glavnog toka.
10. Nema SetShow / NoShow / HideAllComponents / SetVisibilityForListNoWait za image capture.
11. Vijci, navrtke i podloske ostaju u BOM Excelu, ali bez slike.

Fajlovi:
- CATIA_VISUAL_BOM_EXPORTER.CATScript  (glavni CATScript makro)
- CATIA_VISUAL_BOM_EXPORTER.bas        (VBA verzija - proveriti/sinhronizovati posle CATScript testa)

Output:
Makro pravi folder na Desktop-u:
<ROOT_PARTNUMBER>_VISUAL_BOM_EXPORT

U folderu:
- VISUAL_BOM_EXPORT.xlsx
- IMAGES
- THUMBNAILS
- DEBUG_PHASE_LOG.txt

Default test podesavanja u CATScript:
Const TEST_MODE = True
Const TEST_MAX_ROWS = 20

Prvo testirati u CATIA na 20 redova.
Ako radi, promeniti:
Const TEST_MODE = False

Napomena:
CATIA V5 Automation ne moze biti izvrsen/testiran na GitHub-u.
Makro mora biti testiran na Windows racunaru sa CATIA V5 i Microsoft Excel instalacijom.

Kriterijum uspeha testa:
- Excel ima kolone iz CATIA native BOM-a.
- Excel ima Thumbnail kolonu.
- Fastener redovi ostaju u Excelu bez slike.
- Prvih 20 ne-fastener redova dobijaju thumbnail slike.
- Otvoreni CATPart/CATProduct dokumenti se zatvaraju bez snimanja.
- Glavni CATProduct se ne zatvara i ne snima.
- Nema hide/show metode za slikanje.
