# Externé skripty pre Enterprise Architect (EA)

Tento priečinok obsahuje externé skripty, ktoré sa spúšťajú mimo Enterprise Architecta (EA) a pracujú s ním cez COM rozhranie. Skripty sa neukladajú do EA, sú verzované v Gite a spúšťajú sa z príkazového riadku.

## 1. Účel

Hlavný skript v tomto priečinku je ExportNRelToExcel.js. Jeho účel:

- pripojiť sa k bežiacej inštancii EA,
- zobrať balík aktuálne vybraný v Project Browseri,
- vyhľadať konektory a načítať tagy N_RelationID a OriginalName,
- vytvoriť alebo aktualizovať Excel súbor (.xlsx) s exportom.

Súčasťou priečinka je aj run_export_nrel.cmd, ktorý slúži ako jednoduchý spúšťač tohto skriptu.

## 2. Štruktúra priečinka

V priečinku external by mali byť minimálne tieto súbory:

- ExportNRelToExcel.js – hlavný JScript skript pre export,
- run_export_nrel.cmd – spúšťací .cmd súbor,
- README.md – tento súbor s popisom.

ExportNRelToExcel.js musí byť uložený v rovnakom priečinku ako run_export_nrel.cmd.

## 3. Predpoklady

Na správne fungovanie je potrebné:

- operačný systém Windows,
- nainštalovaný Enterprise Architect,
- nainštalovaný Microsoft Excel,
- povolené spúšťanie Windows Script Host (cscript),
- mať otvorený EA a v ňom načítaný model.

## 4. Inštalácia pre člena tímu

Odporúčaný adresár pre skripty na lokálnom disku:

C:\EA-Skripty\

Postup:

1. Vytvoriť adresár C:\EA-Skripty\ (ak ešte neexistuje).
2. Naklonovať repozitár:

   cd C:\EA-Skripty\
   git clone https://github.com/NASES-Slovakia/Enterprise-Architect-scripts.git

3. Externé skripty budú potom na ceste:

   C:\EA-Skripty\Enterprise-Architect-scripts\external\

Pri aktualizácii skriptov stačí:

   cd C:\EA-Skripty\Enterprise-Architect-scripts\
   git pull

## 5. Použitie – spustenie exportu

1. Spustiť Enterprise Architect.
2. V Project Browseri označiť balík, ktorý sa má exportovať.
3. Nechať EA otvorený (nesmie sa zatvoriť počas behu skriptu).
4. Spustiť príkazový riadok (Command Prompt).
5. Prejsť do priečinka external:

   cd C:\EA-Skripty\Enterprise-Architect-scripts\external\

6. Spustiť skript:

   run_export_nrel.cmd

Skript:

- pripojí sa k bežiacemu EA,
- zoberie aktuálne označený balík,
- vyžiada od používateľa priečinok a názov Excel súboru (podľa logiky v ExportNRelToExcel.js),
- vytvorí alebo aktualizuje .xlsx súbor.

Po dokončení skriptu zostane okno príkazového riadku otvorené (vďaka pause), aby bol viditeľný log.

## 6. Obsah súboru run_export_nrel.cmd

Súbor run_export_nrel.cmd má byť veľmi jednoduchý a predpokladá, že ExportNRelToExcel.js je v rovnakom priečinku:

@echo off
cscript //nologo "%~dp0ExportNRelToExcel.js"
pause

%~dp0 znamená adresár, v ktorom sa nachádza samotný .cmd súbor, takže nie je potrebné hardcodovať cestu typu C:\EA-Skripty\...

## 7. Typický scenár používania

- architekt si raz nastaví repozitár na C:\EA-Skripty\Enterprise-Architect-scripts\,
- pri každom exporte:

  - otvorí EA,
  - klikne na balík v Project Browseri,
  - spustí run_export_nrel.cmd z priečinka external.

Tým pádom všetci v tíme používajú rovnaký kód z Gitu a správanie je konzistentné.

## 8. Aktualizácia a údržba

Zmeny skriptu ExportNRelToExcel.js sa robia v Gite. Po schválení a mergnutí do hlavnej vetvy stačí, aby členovia tímu spustili:

cd C:\EA-Skripty\Enterprise-Architect-scripts\
git pull

Nie je potrebné nič meniť v Enterprise Architecte, všetko prebieha mimo EA.
