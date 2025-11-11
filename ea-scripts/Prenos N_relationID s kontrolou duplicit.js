// Put N_RelationID into connector Name + detect duplicates (diagram + global list)
!INC Local Scripts.EAConstants-JScript

// ==== Nastavenia ====
var TAG_NAME = "N_RelationID";
var PREFIX_START = "[RID:";     // začiatok prefixu
var PREFIX_END   = "] ";        // koniec prefixu (medzera je len vizuálna)
var HIGHLIGHT_DUPLICATES = true; // zvýrazni duplicity na diagrame
var DUP_COLOR = 0x0000FF;        // BGR (0x0000FF = červená)
// ====================

function Main(){
  var d = getActiveDiagram();
  if (!d){ Session.Prompt("Otvorte diagram a spustite znova.", promptOK); return; }

  var changed = 0;
  var map = {}; // normID -> [{dl, con, relIdRaw}]

  // 1) Doplnenie názvu a zber ID
  for (var e = new Enumerator(d.DiagramLinks); !e.atEnd(); e.moveNext()){
    var dl = e.item();
    var con = Repository.GetConnectorByID(dl.ConnectorID);
    if (!con) continue;

    var relIdRaw = getTag(con, TAG_NAME);
    if (!relIdRaw || trim(relIdRaw)=="") continue;

    var baseName = getOrSetOriginalName(con);
    var newName  = PREFIX_START + relIdRaw + PREFIX_END + baseName;

    if (con.Name != newName){
      con.Name = newName;
      con.Update();
      changed++;
    }

    var key = norm(relIdRaw);
    if (!map[key]) map[key] = [];
    map[key].push({ dl: dl, con: con, relIdRaw: relIdRaw });
  }

  // 2) Kontrola duplicit na DIAGRAME
  var dupTotal = 0;
  Session.Output("=== Duplicitné N_RelationID na diagrame: " + d.Name + " ===");
  for (var k in map){
    if (map[k].length > 1){
      dupTotal++;
      Session.Output("• " + map[k][0].relIdRaw + "  (počet na diagrame: " + map[k].length + ")");
      for (var i=0;i<map[k].length;i++){
        var it = map[k][i];
        if (HIGHLIGHT_DUPLICATES){
          it.dl.LineColor = DUP_COLOR;
          it.dl.Update();
        }
        Session.Output("    - ConnectorID=" + it.con.ConnectorID + " | " + it.con.Name + " | GUID=" + it.con.ConnectorGUID);
      }
    }
  }

  // 3) GLOBÁLNA kontrola (celý repozitár) pre ID z tohto diagramu
  Session.Output("=== Globálne výskyty N_RelationID v celej DB (pre ID prítomné na diagrame) ===");
  for (var k2 in map){
    var anyItem = map[k2][0];
    var relId   = anyItem.relIdRaw;
    var rows = listGlobalMatches(relId);  // všetky konektory s týmto ID v DB
    // odfiltruj duplicitné záznamy toho istého konektora (ak je na viacerých diagramoch)
    rows = dedupBy(rows, function(r){ return r.connector_id; });

    Session.Output("• ID '" + relId + "': celkom výskytov = " + rows.length);
    for (var j=0; j<rows.length; j++){
      var r = rows[j];
      Session.Output("    - ConnectorID=" + r.connector_id + " | Name=" + r.connector_name +
                     " | GUID=" + r.connector_guid + " | Diagram=" + (r.diagram_name || "<nepriradený>"));
    }
  }

  d.Update();
  Repository.ReloadDiagram(d.DiagramID);
  Session.Output("Hotovo. Aktualizovaných konektorov: " + changed + ". Duplicit na diagrame: " + dupTotal);
}

// --- Helpers ---

function getActiveDiagram(){
  var t = Repository.GetContextItemType && Repository.GetContextItemType();
  if (t==otDiagram) return Repository.GetContextObject();
  return Repository.GetCurrentDiagram();
}

function getTag(con, name){
  var it = new Enumerator(con.TaggedValues);
  for (; !it.atEnd(); it.moveNext()){
    var tv = it.item();
    if (tv.Name && tv.Name.toLowerCase() == name.toLowerCase())
      return tv.Value || "";
  }
  return "";
}

function setTag(con, name, val){
  var it = new Enumerator(con.TaggedValues);
  for (; !it.atEnd(); it.moveNext()){
    var tv = it.item();
    if (tv.Name && tv.Name.toLowerCase() == name.toLowerCase()){
      tv.Value = val; tv.Update(); con.TaggedValues.Refresh(); return;
    }
  }
  var tvNew = con.TaggedValues.AddNew(name, "");
  tvNew.Value = val; tvNew.Update(); con.TaggedValues.Refresh();
}

function getOrSetOriginalName(con){
  var orig = getTag(con, "OriginalName");
  if (orig && orig != "") return orig;

  var nameNow = con.Name || "";
  var base = nameNow;

  var start = nameNow.indexOf(PREFIX_START);
  var end = nameNow.indexOf(PREFIX_END);
  if (start == 0 && end > start){
    base = nameNow.substring(end + PREFIX_END.length);
  }
  setTag(con, "OriginalName", base);
  return base;
}

function trim(s){ return (s||"").replace(/^\s+|\s+$/g,""); }
function norm(s){ return trim(String(s)).toLowerCase(); }
function escSQL(s){ return String(s).replace(/'/g,"''"); }

function dedupBy(arr, keyFn){
  var seen = {};
  var res = [];
  for (var i=0;i<arr.length;i++){
    var k = keyFn(arr[i]);
    if (!seen[k]){ seen[k]=true; res.push(arr[i]); }
  }
  return res;
}

// Vráti zoznam všetkých konektorov v DB, ktoré majú N_RelationID == relId (case-insensitive)
// Vracia polia objektov: {connector_id, connector_guid, connector_name, diagram_name}
function listGlobalMatches(relId){
  var sql =
    "SELECT c.Connector_ID AS connector_id, c.ea_guid AS connector_guid, c.Name AS connector_name, d.Name AS diagram_name " +
    "FROM t_connectortag n " +
    "JOIN t_connector c ON c.Connector_ID = n.elementid " +            // POZOR: u vás je to elementid
    "LEFT JOIN t_diagramlinks dl ON dl.ConnectorID = c.Connector_ID " +
    "LEFT JOIN t_diagram d ON d.Diagram_ID = dl.DiagramID " +
    "WHERE lower(n.property) = 'n_relationid' AND lower(n.value) = '" + escSQL(relId).toLowerCase() + "'";

  var xml = Repository.SQLQuery(sql);
  return parseEAResult(xml, ["connector_id","connector_guid","connector_name","diagram_name"]);
}

// Parse EA SQL XML result to array of rows (by column names)
function parseEAResult(xml, cols){
  var rows = [];
  if (!xml) return rows;
  // EA vracia <Row>...</Row>; parsuj jednoducho cez regex
  var rowRe = /<Row[^>]*>([\s\S]*?)<\/Row>/ig;
  var m;
  while ( (m = rowRe.exec(xml)) ){
    var chunk = m[1];
    var obj = {};
    for (var i=0; i<cols.length; i++){
      var col = cols[i];
      var colRe = new RegExp("<" + col + ">([\\s\\S]*?)<\\/" + col + ">", "i");
      var mm = colRe.exec(chunk);
      obj[col] = mm ? mm[1] : "";
    }
    rows.push(obj);
  }
  return rows;
}

Main();
