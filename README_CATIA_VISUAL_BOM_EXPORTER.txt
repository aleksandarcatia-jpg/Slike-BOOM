CATIA_VISUAL_BOM_EXPORTER - kratko uputstvo

Fajl makroa:
CATIA_VISUAL_BOM_EXPORTER.bas

Direktno selektabilni CATIA script fajlovi:
CATIA_VISUAL_BOM_EXPORTER.CATScript

Kako ubaciti u CATIA V5:
Opcija A - ako zelite direktno da selektujete skriptu:
1. Tools > Macro > Macros...
2. Macro libraries... / Select...
3. Izaberite Directory kao tip biblioteke.
4. Pokazite na folder gde je fajl:
   C:\Users\Aca Stojanovic\Documents\Codex\2026-05-18\napravi-odmah-funkcionalni-catia-v5-vba
5. Izaberite CATIA_VISUAL_BOM_EXPORTER.CATScript.
6. Run.

U aktuelnoj verziji Excel workbook se kreira odmah, BOM i LOG redovi se upisuju tokom rada, a fajl se snima posle svake obradjene stavke.

Opcija B - ako koristite VBA library:
1. Otvorite CATIA V5.
2. Tools > Macro > Macros...
3. Izaberite ili napravite VBA library (.catvba).
4. Otvorite VBA Editor.
5. Insert > Module.
6. Iskopirajte kompletan sadrzaj fajla CATIA_VISUAL_BOM_EXPORTER.bas u modul.
7. Sacuvajte VBA library.

Kako pokrenuti:
1. Otvorite glavni CATProduct sklop.
2. Aktivirajte prozor sklopa.
3. Tools > Macro > Macros...
4. Izaberite CATIA_VISUAL_BOM_EXPORTER.
5. Run.

Izlaz:
Makro pravi folder:
<PartNumber>_VISUAL_BOM_EXPORT

U njemu pravi:
- VISUAL_BOM_EXPORT.xlsx
- IMAGES folder
- MAIN_ASSEMBLY.png

Napomene:
- CATIA V5 Automation CaptureToFile uglavnom podrzava BMP/JPEG/TIFF, ne uvek PNG direktno. Makro zato hvata BMP pa konvertuje u PNG preko Windows WIA COM komponente.
- Ako WIA nije dostupna na racunaru, makro pravi JPG fallback i upisuje WARNING u EXPORT_LOG.
- Ako slike izlaze lose, pre pokretanja u CATIA podesite belu pozadinu i zeljeni render style u 3D vieweru. Makro koristi izometrijski Viewpoint3D, Reframe/Fit All i trenutni viewer.
- Hide/show se radi preko Selection.VisProperties.SetShow. Na nekim CATIA V5 konfiguracijama instance, CGR/lightweight komponente ili unloaded reference ne reaguju potpuno pouzdano. Makro ne brise, ne snima i ne menja geometriju, a na kraju pokusava da vrati prikaz celog sklopa.
- Ako zelite samo selektovane stavke, u kodu promenite:
  Const EXPORT_SELECTED_ONLY = True
- Za test bez slika podesite:
  Const EXPORT_IMAGES = False
- Za test samo prvih 10 stavki podeseno je:
  Const MAX_ITEMS_TO_EXPORT = 10
  Ako zelite sve stavke, stavite 0.
- Za prvi test samo CATPart delova podeseno je:
  Const SKIP_ASSEMBLIES = True
  Ako zelite i podsklopove, stavite False.
- Timeout po komponenti je:
  Const IMAGE_EXPORT_TIMEOUT_SECONDS = 10
  Napomena: ako se sam CATIA COM poziv potpuno blokira unutar CATIA, VBScript ne moze nasilno da ga prekine dok CATIA ne vrati kontrolu makrou.
