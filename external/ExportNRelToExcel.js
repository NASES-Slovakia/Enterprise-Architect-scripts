// ============================================
// Export N_RelationID väzieb z aktuálneho balíka v EA do Excelu
// Spúšťa sa mimo EA cez:  cscript //nologo ExportNRelToExcel.js
// Musí bežať Enterprise Architect a musí byť otvorený model.
// ============================================

// Pripojenie na otvorený EA (musí byť spustený)
var eaApp;
try {
    eaApp = GetObject("", "EA.App");
} catch (e) {
    WScript.Echo("EA.App nie je spusteny (Enterprise Architect).");
    WScript.Quit(1);
}
var Repository = eaApp.Repository;

// ===== Nastavenia =====
var OPEN_AFTER_SAVE = false;                 // ponechat Excel otvoreny po ulozeni?
var KEY_COL   = "N_RelationID";              // kluc pre upsert v hlavnom liste
var MAIN_SHEET = "EA Export";
var DETAIL_FIELDS = ["N_RelationID","Pattern"]; // do detailov zapisujeme len tieto dve
var META_LABEL = "EA_GUID";
var META_COL_L = 5;
var META_COL_V = 6;
var PROTECTED_COLOR = 22;   // svetlocervena pre bunky, ktore sa nemaju menit
var UNKNOWN_COLOR   = 36;   // zlta pre riadky bez detailneho harka
// ======================

// ---- logovanie ----
function log(msg){ 
    try{ WScript.Echo(msg); }catch(e){} 
}

// ---- XML helpery ----
function createXmlDoc() {
  var v=["MSXML2.DOMDocument.6.0","MSXML2.DOMDocument.4.0","MSXML2.DOMDocument.3.0","MSXML2.DOMDocument"];
  for (var i=0;i<v.length;i++){ 
      try { 
          var x=new ActiveXObject(v[i]); 
          x.async=false; 
          return x; 
      } catch(e){} 
  }
  return null;
}

function sanitizeXml(s){
  if (!s) return s;
  s = s.replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g,"");
  s = s.replace(/&(?!amp;|lt;|gt;|quot;|apos;|#\d+;)/g,"&amp;");
  return s;
}

// ---- priecinok ----
function pickFolder() {
  try {
    var sh = new ActiveXObject("Shell.Application");
    var folder = sh.BrowseForFolder(0, "Vyber priecinok pre XLSX", 0x0011);
    return folder ? folder.Self.Path : "";
  } catch(e) { return ""; }
}

function ensureXlsxName(name){
  if (!name || name=="") name = "EA_export";
  name = name.replace(/[\\\/:\*\?"<>\|]/g,"_");
  if (!/\.xlsx$/i.test(name)) name += ".xlsx";
  return name;
}

// ---- výber / vytvorenie XLSX cez Excel GetOpen/SaveAs + fallback ----
function pickExcelPathInteractive(initialFolder){
  try {
    log("* Excel GetOpen/SaveAs - vyber suboru...");
    var xl = new ActiveXObject("Excel.Application");
    try { xl.DisplayAlerts = false; } catch(e){}
    try { xl.Visible = true; } catch(e){}
    try { xl.ChangeFileOpenDirectory(initialFolder); } catch(e){}

    // 1) vybrat existujuci
    var openPath = xl.GetOpenFilename(
      "Excel subory (*.xlsx),*.xlsx,All files (*.*),*.*",
      1,
      "Vyber existujuci XLSX (Zrusit pre novy)"
    );
    if (openPath && openPath !== true && (""+openPath).toLowerCase()!=="false"){
      var p=""+openPath;
      log("-> Vybrany existujuci subor: " + p);
      try{ xl.Quit(); }catch(e){}
      return p;
    }

    // 2) zrusene -> SaveAs novy
    log("* Open zruseny, idem na SaveAs...");
    try { xl.ChangeFileOpenDirectory(initialFolder); } catch(e){}
    var def = initialFolder+"\\EA_export.xlsx";
    var savePath = xl.GetSaveAsFilename(
      def,
      "Excel subory (*.xlsx),*.xlsx",
      "Zadaj nazov noveho XLSX"
    );
    try{ xl.Quit(); }catch(e){}
    if (savePath && savePath !== true && (""+savePath).toLowerCase()!=="false"){
      var s=""+savePath;
      if(!/\.xlsx$/i.test(s)) s+=".xlsx";
      log("-> Zadany novy subor: " + s);
      return s;
    }

    // 3) zrusene aj SaveAs -> pojdeme na konzolové zadanie
    log("* Zrusene Open aj SaveAs - idem na manualne zadanie nazvu v konzole.");
  } catch(e){
    log("[WARN] Excel COM nie je dostupny: " + e.message + " - idem na manualne zadanie nazvu.");
  }

  // Fallback – manualne zadanie cez konzolu (cscript)
  try {
    WScript.StdOut.Write("Zadaj nazov existujuceho alebo noveho suboru (.xlsx) alebo nechaj prazdne pre zrusenie: ");
    var name = WScript.StdIn.ReadLine();
    if (!name) return "";
    name = ensureXlsxName(name);
    var full = initialFolder + "\\" + name;
    log("-> Manualne zadany subor: " + full);
    return full;
  } catch(e2) {
    log("[ERR] Nepodarilo sa nacitat vstup z konzoly: " + e2.message);
    return "";
  }
}

// ---- EA helpery ----
function collectPackageIds(pkg, arr) {
  if (!pkg) return;
  arr.push(pkg.PackageID);
  for (var i=0;i<pkg.Packages.Count;i++) collectPackageIds(pkg.Packages.GetAt(i), arr);
}

function safeSheetName(s, fallback, idx){
  if (!s || s=="") s = fallback + "_" + idx;
  s = s.replace(/[\\\/:\*\?\[\]]/g,"_");
  if (s.length > 31) s = s.substring(0,31);
  return s;
}

// ---------- šablóny (KFK, RAB, REST, SOAP, default) ----------
function buildTemplateKFK(ws){
  if (ws.Cells(1,1).Value) return;
  ws.Cells(1,1).Value="Atribut"; ws.Cells(1,2).Value="Poznamka"; ws.Cells(1,3).Value="Hodnota";
  ws.Range("A1:C1").Font.Bold=true;
  var attrs = [
    ["N_RelationID","", ""],
    ["Pattern","", ""],
    ["Comment","", ""],
    ["Service Name","", ""],
    ["KFK Topic","", ""],
    ["Env","FIX/UAT/PROD",""],
    ["Owner/Contact","mailova skupina",""],
    ["Net Source","FQDN",""],
    ["Net Target","FQDN",""],
    ["Target Ports/TLS","", ""]
  ];
  for (var i=0;i<attrs.length;i++){
    ws.Cells(i+2,1).Value=attrs[i][0];
    ws.Cells(i+2,2).Value=attrs[i][1];
  }
  var last=attrs.length+1;
  var rng=ws.Range(ws.Cells(1,1),ws.Cells(last,3));
  rng.Borders.Weight=2;
  ws.Range("A1:C1").AutoFilter();
  ws.Columns("A:C").AutoFit();
}

function buildTemplateRAB(ws){
  if (ws.Cells(1,1).Value) return;
  ws.Cells(1,1).Value="Atribut"; ws.Cells(1,2).Value="Poznamka"; ws.Cells(1,3).Value="Hodnota";
  ws.Range("A1:C1").Font.Bold=true;
  var attrs = [
    ["N_RelationID","", ""],
    ["Pattern","", ""],
    ["Comment","", ""],
    ["Service Name","", ""],
    ["Base URL","", ""],
    ["HTTP Method","GET/POST/...",""],
    ["Auth","Bearer/Basic/...",""],
    ["Timeouts/Retry","s",""],
    ["Owner/Contact","mail",""],
    ["Net Source","FQDN",""],
    ["Net Target","FQDN",""],
    ["Target Ports/TLS","", ""]
  ];
  for (var i=0;i<attrs.length;i++){
    ws.Cells(i+2,1).Value=attrs[i][0];
    ws.Cells(i+2,2).Value=attrs[i][1];
  }
  var last=attrs.length+1;
  var rng=ws.Range(ws.Cells(1,1),ws.Cells(last,3));
  rng.Borders.Weight=2;
  ws.Range("A1:C1").AutoFilter();
  ws.Columns("A:C").AutoFit();
}

function buildTemplateREST(ws){
  if (ws.Cells(1,1).Value) return;
  ws.Cells(1,1).Value="Atribut"; ws.Cells(1,2).Value="Poznamka"; ws.Cells(1,3).Value="Hodnota";
  ws.Range("A1:C1").Font.Bold=true;
  var attrs = [
    ["N_RelationID","", ""],
    ["Pattern","", ""],
    ["Comment","", ""],
    ["Service Name","", ""],
    ["Endpoint URL","", ""],
    ["HTTP Method","GET/POST/...",""],
    ["Request JSON","", ""],
    ["Response JSON","", ""],
    ["Auth","Bearer/Basic/...",""],
    ["Owner/Contact","mail",""],
    ["Net Source","FQDN",""],
    ["Net Target","FQDN",""]
  ];
  for (var i=0;i<attrs.length;i++){
    ws.Cells(i+2,1).Value=attrs[i][0];
    ws.Cells(i+2,2).Value=attrs[i][1];
  }
  var last=attrs.length+1;
  var rng=ws.Range(ws.Cells(1,1),ws.Cells(last,3));
  rng.Borders.Weight=2;
  ws.Range("A1:C1").AutoFilter();
  ws.Columns("A:C").AutoFit();
}

function buildTemplateSOAP(ws){
  if (ws.Cells(1,1).Value) return;
  ws.Cells(1,1).Value="Atribut"; ws.Cells(1,2).Value="Poznamka"; ws.Cells(1,3).Value="Hodnota";
  ws.Range("A1:C1").Font.Bold=true;
  var attrs = [
    ["N_RelationID","", ""],
    ["Pattern","", ""],
    ["Comment","", ""],
    ["Service Name","", ""],
    ["WSDL URL","", ""],
    ["SOAPAction/Binding","", ""],
    ["WS-Security","", ""],
    ["Owner/Contact","mail",""],
    ["Net Source","FQDN",""],
    ["Net Target","FQDN",""],
    ["Target Ports/TLS","", ""]
  ];
  for (var i=0;i<attrs.length;i++){
    ws.Cells(i+2,1).Value=attrs[i][0];
    ws.Cells(i+2,2).Value=attrs[i][1];
  }
  var last=attrs.length+1;
  var rng=ws.Range(ws.Cells(1,1),ws.Cells(last,3));
  rng.Borders.Weight=2;
  ws.Range("A1:C1").AutoFilter();
  ws.Columns("A:C").AutoFit();
}

function buildTemplateDefault(ws){
  if (ws.Cells(1,1).Value) return;
  ws.Cells(1,1).Value="Atribut"; ws.Cells(1,2).Value="Poznamka"; ws.Cells(1,3).Value="Hodnota";
  ws.Range("A1:C1").Font.Bold=true;
  var attrs=[["N_RelationID","", ""],["Pattern","", ""],["Comment","", ""]];
  for (var i=0;i<attrs.length;i++){
    ws.Cells(i+2,1).Value=attrs[i][0];
    ws.Cells(i+2,2).Value=attrs[i][1];
  }
  var last=attrs.length+1;
  var rng=ws.Range(ws.Cells(1,1),ws.Cells(last,3));
  rng.Borders.Weight=2;
  ws.Range("A1:C1").AutoFilter();
  ws.Columns("A:C").AutoFit();
}

// detekcia typu podla OriginalName
function detectSheetType(originalName){
  if (!originalName) return "";
  var up = (""+originalName).toUpperCase();
  if (up.indexOf("KFK-")  === 0) return "KFK";
  if (up.indexOf("RAB-")  === 0) return "RAB";
  if (up.indexOf("REST-") === 0) return "REST";
  if (up.indexOf("SOAP-") === 0) return "SOAP";
  return "";
}

function stampGuid(ws,guid){
  if (!ws.Cells(1,META_COL_L).Value) ws.Cells(1,META_COL_L).Value=META_LABEL;
  ws.Cells(1,META_COL_V).Value=guid;
  try{ ws.Columns(META_COL_L).Hidden=true; ws.Columns(META_COL_V).Hidden=true; }catch(e){}
}

function findDetailByGuid(wb,guid){
  for (var i=1;i<=wb.Worksheets.Count;i++){
    var ws=wb.Worksheets(i), nm=""+ws.Name;
    if (nm.indexOf("NRel_")==0){
      var val=ws.Cells(1,META_COL_V).Value, lbl=ws.Cells(1,META_COL_L).Value;
      if (lbl==META_LABEL && val && (""+val).toLowerCase()==(""+guid).toLowerCase())
        return ws;
    }
  }
  return null;
}

function findOrCreateRowByLabel(ws,label){
  var maxScan=200;
  for (var r=2;r<=maxScan;r++){
    var v=ws.Cells(r,1).Value;
    if (v && (""+v).toLowerCase()==(""+label).toLowerCase()) return r;
    if (!v && ws.Cells(r,2).Value==null && ws.Cells(r,3).Value==null){
      ws.Cells(r,1).Value=label; 
      return r;
    }
  }
  var end=ws.Cells(ws.Rows.Count,1).End(-4162/*xlUp*/).Row+1;
  ws.Cells(end,1).Value=label;
  return end;
}

// --------------------- MAIN ---------------------
function main(){
  log("== EA -> Excel export start ==");

  var sel=Repository.GetTreeSelectedPackage();
  if(!sel){ log("[ERR] Vyber balik v Project Browseri."); return; }

  var ids=[]; collectPackageIds(sel,ids);
  if(ids.length==0){ log("[ERR] Vetva balika je prazdna."); return; }

  var sql="SELECT "+
    "d.Name AS DiagramName, pd.Name AS DiagramPackage, c.Connector_ID, c.ea_guid AS ConnectorGUID, "+
    "c.Name AS ConnectorName, c.Connector_Type, src.Name AS SourceElement, trg.Name AS TargetElement, "+
    "psrc.Name AS SourcePackage, ptrg.Name AS TargetPackage, nrt.value AS N_RelationID, "+
    "ort.value AS OriginalName, dup.DuplicateCount AS DuplicateCount "+
    "FROM t_diagramlinks dl "+
    "JOIN t_diagram d ON d.Diagram_ID=dl.DiagramID "+
    "JOIN t_connector c ON c.Connector_ID=dl.ConnectorID "+
    "LEFT JOIN t_object src ON src.Object_ID=c.Start_Object_ID "+
    "LEFT JOIN t_object trg ON trg.Object_ID=c.End_Object_ID "+
    "LEFT JOIN t_package psrc ON psrc.Package_ID=src.Package_ID "+
    "LEFT JOIN t_package ptrg ON ptrg.Package_ID=trg.Package_ID "+
    "LEFT JOIN t_package pd ON pd.Package_ID=d.Package_ID "+
    "LEFT JOIN t_connectortag nrt ON nrt.elementid=c.Connector_ID AND lower(nrt.property)='n_relationid' "+
    "LEFT JOIN t_connectortag ort ON ort.elementid=c.Connector_ID AND lower(ort.property)='originalname' "+
    "LEFT JOIN (SELECT lower(value) AS relid, COUNT(*) AS DuplicateCount FROM t_connectortag "+
    "           WHERE lower(property)='n_relationid' GROUP BY lower(value)) dup ON dup.relid=lower(nrt.value) "+
    "WHERE pd.Package_ID IN ("+ids.join(",")+") "+
    "  AND dl.Hidden = 0 "+
    "ORDER BY dup.DuplicateCount DESC, pd.Name, d.Name, c.Connector_ID;";

  var folder=pickFolder(); 
  if(!folder){ log("[ERR] Zrusene (priecinok)."); return; }
  log("[OK] Zvoleny priecinok: " + folder);

  var path=pickExcelPathInteractive(folder);
  if(!path||path==""){ log("[ERR] Zrusene (subor)."); return; }
  if(!/\.xlsx$/i.test(path)) path = ensureXlsxName(path);
  log("[OK] Cielovy subor: " + path);

  var xmlRaw=Repository.SQLQuery(sql); 
  if(!xmlRaw||xmlRaw==""){ log("[ERR] SQL prazdny vysledok."); return; }

  var doc=createXmlDoc(); 
  if(!doc){ log("[ERR] MSXML nie je dostupny."); return; }

  if(!doc.loadXML(sanitizeXml(xmlRaw))){
    var fso0=new ActiveXObject("Scripting.FileSystemObject");
    var dbg=folder+"\\_debug_sqlquery.xml";
    var f0=fso0.CreateTextFile(dbg,true,true);
    f0.Write(xmlRaw); f0.Close();
    log("[ERR] Chyba parsovania XML. Dump: "+dbg);
    return;
  }

  var rows=doc.getElementsByTagName("Row"); 
  if(rows.length==0){ log("[INFO] Ziadne data (vo vetve nie su viditelne konektory)."); return; }
  log("[OK] Nacitane riadky: " + rows.length);

  var first=rows.item(0), colCount=first.childNodes.length;

  var excel=new ActiveXObject("Excel.Application");
  var wb, wsMain;
  var fso=new ActiveXObject("Scripting.FileSystemObject");

  if(fso.FileExists(path)){
    log("* Otvaram existujuci XLSX...");
    wb=excel.Workbooks.Open(path);
    try{ wsMain=wb.Worksheets(MAIN_SHEET); }
    catch(e){ wsMain=wb.Worksheets.Add(); wsMain.Name=MAIN_SHEET; }
  } else {
    log("* Vytvaram novy XLSX...");
    wb=excel.Workbooks.Add();
    wsMain=wb.Worksheets(1); 
    wsMain.Name=MAIN_SHEET;
    for(var c=0;c<colCount;c++){
      wsMain.Cells(1,c+1).Value=first.childNodes.item(c).nodeName;
    }
    wsMain.Range(wsMain.Cells(1,1),wsMain.Cells(1,colCount)).Font.Bold=true;
    wsMain.Range(wsMain.Cells(1,1),wsMain.Cells(1,colCount)).AutoFilter();
  }

  // mapovanie hlaviciek
  var existingCols={}, lastCol=wsMain.Cells(1,wsMain.Columns.Count).End(-4159/*xlToLeft*/).Column;
  for(var cc=1; cc<=lastCol; cc++){
    var hv=wsMain.Cells(1,cc).Value;
    if(hv) existingCols[(""+hv).toLowerCase()]=cc;
  }
  for(var c2=0;c2<colCount;c2++){
    var nm=first.childNodes.item(c2).nodeName.toLowerCase();
    if(!existingCols[nm]){
      lastCol++;
      wsMain.Cells(1,lastCol).Value=first.childNodes.item(c2).nodeName;
      existingCols[nm]=lastCol;
    }
  }
  var keyColIndex=existingCols[KEY_COL.toLowerCase()];
  if(!keyColIndex){
    lastCol++;
    wsMain.Cells(1,lastCol).Value=KEY_COL;
    keyColIndex=lastCol;
    existingCols[KEY_COL.toLowerCase()]=keyColIndex;
  }

  function idxOf(col){
    for(var i=0;i<colCount;i++){
      if(first.childNodes.item(i).nodeName.toLowerCase()==col.toLowerCase()) return i;
    }
    return -1;
  }
  var idxRel=idxOf("N_RelationID"),
      idxOrig=idxOf("OriginalName"),
      idxGUID=idxOf("ConnectorGUID");

  var lastRowMain=wsMain.Cells(wsMain.Rows.Count,keyColIndex).End(-4162/*xlUp*/).Row;
  var keyMap={};
  for(var r1=2;r1<=lastRowMain;r1++){
    var kv=wsMain.Cells(r1,keyColIndex).Value;
    if(kv!==null&&kv!=="") keyMap[(""+kv).toLowerCase()]=r1;
  }
  var nextRow=(lastRowMain>=2? lastRowMain+1 : 2);

  var keepSheets={};
  var oldAlerts=excel.DisplayAlerts; 
  excel.DisplayAlerts=false;

  for(var r=0;r<rows.length;r++){
    var row=rows.item(r);
    var nrel=(idxRel>=0)? row.childNodes.item(idxRel).text : "";
    var orig=(idxOrig>=0)? row.childNodes.item(idxOrig).text : "";
    var guid=(idxGUID>=0)? row.childNodes.item(idxGUID).text : "";
    var hasKey = nrel && (""+nrel).replace(/\s+/g,"")!=="";

    // hlavny harok - upsert
    var targetRow;
    if (hasKey && keyMap[(""+nrel).toLowerCase()]) {
      targetRow = keyMap[(""+nrel).toLowerCase()];
    } else {
      targetRow = nextRow;
      nextRow++;
    }
    for(var c3=0;c3<colCount;c3++){
      var colName=first.childNodes.item(c3).nodeName.toLowerCase();
      var tgt=existingCols[colName];
      if(tgt) wsMain.Cells(targetRow,tgt).Value = row.childNodes.item(c3).text;
    }
    if (hasKey && !keyMap[(""+nrel).toLowerCase()]) {
      keyMap[(""+nrel).toLowerCase()] = targetRow;
    }

    // typ podla OriginalName
    var sheetType = detectSheetType(orig);

    // highlight - ak nebude detailny harok
    if (!hasKey || !sheetType) {
      wsMain.Rows(targetRow).Interior.ColorIndex = UNKNOWN_COLOR;
    }

    if (!hasKey) continue;
    if (!sheetType) continue;

    // detailny harok
    var desiredName = safeSheetName("NRel_" + nrel, "NRel", r+1);
    var ws = guid ? findDetailByGuid(wb,guid) : null;
    if(!ws){
      try{ ws = wb.Worksheets(desiredName); }catch(e){ ws=null; }
      if(!ws){
        ws = wb.Worksheets.Add();
        ws.Move(null, wb.Worksheets(wb.Worksheets.Count));
        ws.Name=desiredName;
        if (sheetType=="KFK")        buildTemplateKFK(ws);
        else if (sheetType=="RAB")   buildTemplateRAB(ws);
        else if (sheetType=="REST")  buildTemplateREST(ws);
        else if (sheetType=="SOAP")  buildTemplateSOAP(ws);
        else                         buildTemplateDefault(ws);
      }
    } else {
      if (!ws.Cells(1,1).Value) {
        if (sheetType=="KFK")        buildTemplateKFK(ws);
        else if (sheetType=="RAB")   buildTemplateRAB(ws);
        else if (sheetType=="REST")  buildTemplateREST(ws);
        else if (sheetType=="SOAP")  buildTemplateSOAP(ws);
        else                         buildTemplateDefault(ws);
      }
    }

    // premenuj pri zmene N_RelationID
    if(ws.Name != desiredName){
      try{ ws.Name=desiredName; }
      catch(e){
        var base=desiredName, iTry=2;
        while(true){
          var candidate=safeSheetName(base+"_"+iTry,"NRel",r+1);
          try{ ws.Name=candidate; desiredName=candidate; break; }
          catch(e2){ iTry++; if(iTry>99) break; }
        }
      }
    }

    if(guid){
      stampGuid(ws,guid);
    }

    // do detailu iba N_RelationID a Pattern + cervene pozadie
    function setAllowed(label,value){
      for(var iA=0;iA<DETAIL_FIELDS.length;iA++){
        if(DETAIL_FIELDS[iA].toLowerCase()==label.toLowerCase()){
          var rr=findOrCreateRowByLabel(ws,label);
          ws.Cells(rr,3).Value=value;
          ws.Cells(rr,3).Interior.ColorIndex=PROTECTED_COLOR;
          return;
        }
      }
    }
    setAllowed("N_RelationID", nrel);
    setAllowed("Pattern",      orig);

    keepSheets[ws.Name]=true;
  }

  // odstranenie starych NRel_* harkov
  for(var i=wb.Worksheets.Count; i>=1; i--){
    var wsx=wb.Worksheets(i), nm=""+wsx.Name;
    if(nm.indexOf("NRel_")==0 && !keepSheets[nm]){
      try{ wsx.Delete(); }catch(e){}
    }
  }

  excel.DisplayAlerts=oldAlerts;

  wsMain.Columns.AutoFit();
  var rng2=wsMain.Range(wsMain.Cells(1,1), wsMain.Cells(wsMain.UsedRange.Rows.Count, wsMain.UsedRange.Columns.Count));
  rng2.Borders.Weight=2;

  try{
    wb.SaveAs(path,51);   // 51 = xlOpenXMLWorkbook (.xlsx)
  } catch(e2){
    if(!/\.xlsx$/i.test(path)){
      path=path+".xlsx";
      wb.SaveAs(path,51);
    } else {
      throw e2;
    }
  }

  if(OPEN_AFTER_SAVE){
    excel.Visible=true;
  } else {
    wb.Close(false); 
    excel.Quit();
  }

  log("[OK] Hotovo: export/aktualizacia do '"+path+"'.");
}

main();
