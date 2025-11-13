# Externé skripty (spúšťané mimo Enterprise Architect)
Tento priečinok obsahuje skripty, ktoré sa spúšťajú mimo Enterprise Architecta (EA) cez cscript alebo .cmd súbory. Externé skripty pracujú s bežiacou aplikáciou EA cez COM objekt EA.App a používajú aktuálne vybraný balík v Project Browseri. Dostupné skripty: ExportNRelToExcel.js – export konektorov s tagom N_RelationID do Excelu, run_export_nrel.cmd – jednoduchý spúšťač.

# 1. Odporúčaná lokálna inštalácia
Odporúčaný priečinok: C:\EA-Skripty\
Klon repozitára:
cd C:\EA-Skripty
git clone https://github.com/<organizacia>/Enterprise-Architect-scripts.git
Štruktúra:
C:\EA-Skripty\Enterprise-Architect-scripts\external\
C:\EA-Skripty\Enterprise-Architect-scripts\ea-scripts\
Aktualizácia:
git pull

# 2. ExportNRelToExcel.js
Skript sa pripojí ku bežiacemu EA, načíta balík vybraný v Project Browseri, vyhľadá konektory a tagy N_RelationID a OriginalName, vytvorí alebo aktualizuje Excel (.xlsx). Výstup obsahuje hlavný list EA Export, detailné listy NRel_<ID>, šablóny KFK/RAB/REST/SOAP a zvýraznenie neznámych riadkov.

# 3. Spustenie cez príkazový riadok
1. Spustiť Enterprise Architect
2. Vybrať balík v Project Browseri
3. Spustiť:
cd C:\EA-Skripty\Enterprise-Architect-scripts\external
run_export_nrel.cmd

# 4. Spúšťanie z EA (stub script)
V EA: Develop → Scripting → Project Browser → New Script (JScript). Vložiť inline kód:
!INC Local Scripts.EAConstants-JScript
function main() {
   var sh = new ActiveXObject("WScript.Shell");
   var cmd = "C:\\EA-Skripty\\Enterprise-Architect-scripts\\external\\run_export_nrel.cmd";
   sh.Run("\"" + cmd + "\"", 1, false);
}
main();

# 5. SharePoint / OneDrive
Ak je repozitár synchronizovaný cez OneDrive/SharePoint, cesta môže vyzerať:
C:\Users\<meno>\Firma\Enterprise-Architect-scripts\external\
V skripte treba upraviť cestu:
var cmd = "C:\\Users\\<meno>\\Firma\\Enterprise-Architect-scripts\\external\\run_export_nrel.cmd";

# 6. Údržba a zmeny
Odporúčaný postup: zmeny robiť cez Git (branch → PR → review → merge). Ostatní členovia aktualizujú lokálne pomocou:
git pull

