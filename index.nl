<!DOCTYPE html>
<html lang="nl">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>ORCA blokschema ombouwen v28.08</title>

<script src="https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/tesseract.js@5/dist/tesseract.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/xlsx-js-style@1.2.0/dist/xlsx.bundle.js"></script>
<style>
*{box-sizing:border-box}
body{margin:0;font-family:Arial,sans-serif;background:#eef0f2;color:#222}
.toolbar{position:sticky;top:0;z-index:20;background:#fff;border-bottom:1px solid #ccc;padding:10px;display:flex;gap:8px;align-items:center;flex-wrap:wrap}
button,select,input{font-size:14px}
button,select{border:0;border-radius:6px;padding:9px 13px;cursor:pointer}
button{background:#198754;color:#fff}
button.gray{background:#6c757d}
button.red{background:#dc3545}
button.blue{background:#0d6efd}
button:disabled{opacity:.45;cursor:not-allowed}
/* Deze twee functies blijven in de code beschikbaar, maar de knoppen
   worden niet meer in de werkbalk getoond. */
#autoPni{display:none!important}
.status{font-size:14px;color:#555}
#workspace{padding:20px;overflow:auto;text-align:center}
#page{position:relative;display:inline-block;background:#fff;box-shadow:0 2px 12px #0003;line-height:0}
canvas{display:block}
#overlay{position:absolute;left:0;top:0;cursor:crosshair}
.info{max-width:1100px;margin:12px auto 0;background:#fff;border-radius:8px;padding:12px 15px;font-size:14px}
.step{font-weight:bold;font-size:16px}
.ok{color:#198754;font-weight:bold}
.wca{display:inline-block;margin:2px 5px;padding:4px 7px;background:#eee;border-radius:4px}

@media print{
  @page{size:auto;margin:0;}
  html,body{margin:0!important;padding:0!important;background:#fff!important;width:100%!important;}
  .toolbar,.info{display:none!important;}
  #workspace{padding:0!important;margin:0!important;overflow:visible!important;width:100%!important;text-align:left!important;}
  #page{display:block!important;position:relative!important;width:100%!important;max-width:100%!important;margin:0!important;box-shadow:none!important;background:#fff!important;}
  #page canvas{width:100%!important;height:auto!important;display:block!important;}
  #page #overlay{position:absolute!important;left:0!important;top:0!important;width:100%!important;height:100%!important;}
}
</style>
</head>
<body>
<div id="saveDirectoryTopRight" style="position:fixed;top:12px;right:12px;z-index:99999;">
  <button id="changeSaveDirectoryBtn" type="button" style="background:#f28c28;color:#fff;border:1px solid #d87512;border-radius:7px;padding:7px 12px;cursor:pointer;font-weight:600;white-space:nowrap;">Opslagmap wijzigen</button>
</div>


<div class="toolbar">
  <input id="pdfFile" type="file" accept="application/pdf">
  <button id="schemaFolderBtn" class="blue" type="button">Map met blokschema's kiezen</button>
  <input id="excelFile" type="file" accept=".xlsx,.xls" style="display:none">
  <button id="excelBtn" class="gray">Excel koppelen</button>
  <button id="extraMaterialBtn" class="blue">Extra materiaal</button>

  <button id="prev" class="gray" disabled>←</button>
  <span id="pageInfo">0 / 0</span>
  <button id="next" class="gray" disabled>→</button>

  <select id="daaSelect" disabled style="display:none">
    <option value="">Kies DAA...</option>
  </select>

  <button id="startDaa" disabled>DAA plaatsen</button>
  <button id="autoWca" class="blue" disabled>WCA → DAA automatisch</button>
  <button id="autoPni" class="blue" disabled>PNI → DAA automatisch</button>

  <button id="startDivider2" class="blue" disabled>2 verdeler</button>
  <button id="startDivider3" class="blue" disabled>3 verdeler</button>
  <button id="startDivider4" class="blue" disabled>4 verdeler</button>
  <button id="startExistingDivider" class="blue" disabled>Bestaande verdeler</button>
  <button id="startAftakker" class="blue" disabled>Aftakker</button>
  <button id="deleteDaaBtn" class="red" disabled style="display:none">DAA wissen</button>
  <button id="saveBtn" class="blue" disabled>Opslaan</button>
  <button id="eraserBtn" class="gray" disabled>Gum</button>
  <button id="clearBtn" class="red" disabled>Alles wissen</button>

  <span id="status" class="status">Open eerst een ORCA-PDF.</span>
</div>

<div class="info">
  <div id="step" class="step">Stap 1: upload de bestaande ORCA-PDF.</div>
  <div id="addressInfo"></div>

<div id="manualAddressModal" style="display:none;position:fixed;inset:0;z-index:250;background:rgba(0,0,0,.55);align-items:center;justify-content:center;padding:16px;">
  <div style="background:#fff;width:min(440px,94vw);padding:22px;border-radius:12px;box-shadow:0 8px 30px #0005;">
    <div style="font-size:19px;font-weight:700;margin-bottom:10px;">Adres handmatig invoeren</div>
    <div style="font-size:14px;line-height:1.45;margin-bottom:14px;">
      Het adres kon niet betrouwbaar uit de PDF worden gelezen. Vul hieronder het juiste adres in. Dit adres wordt linksboven op het opgeslagen blokschema geplaatst.
    </div>
    <label for="manualAddressInput" style="display:block;text-align:left;font-weight:600;margin-bottom:6px;">Straat, huisnummer, postcode en plaats</label>
    <input id="manualAddressInput" type="text" autocomplete="street-address"
      placeholder="Bijv. Polderland 77, 1112 RH Diemen"
      style="box-sizing:border-box;width:100%;padding:12px;font-size:16px;border:1px solid #aaa;border-radius:7px;margin-bottom:16px;">
    <div style="display:flex;gap:10px;flex-wrap:wrap;">
      <button id="manualAddressSave" class="blue" type="button" style="flex:1;min-width:130px;">Adres bevestigen</button>
      <button id="manualAddressSkip" class="gray" type="button" style="flex:1;min-width:130px;">Later invullen</button>
    </div>
  </div>
</div>

  <div id="found">Nog geen WCA-teksten uitgelezen.</div>
</div>


<div id="daaQuestion" style="
  display:none; position:fixed; inset:0; z-index:100;
  background:rgba(0,0,0,.45); align-items:center; justify-content:center;">
  <div style="
    background:#fff; width:min(420px,90vw); padding:24px; border-radius:12px;
    box-shadow:0 8px 30px #0005; text-align:center;">
    <div id="daaQuestionText" style="font-size:18px;font-weight:bold;margin-bottom:20px;">
      Wil je nog een DAA plaatsen?
    </div>
    <div style="display:flex;gap:12px;justify-content:center;">
      <button id="daaYes" class="blue" style="min-width:110px;">JA</button>
      <button id="daaNo" class="red" style="min-width:110px;">NEE</button>
    </div>
  </div>
</div>

<div id="dividerChoicePanel" style="display:none;position:fixed;inset:0;z-index:115;background:rgba(0,0,0,.45);align-items:center;justify-content:center;">
  <div style="background:#fff;width:min(400px,90vw);padding:24px;border-radius:12px;box-shadow:0 8px 30px #0005;text-align:center;">
    <div style="font-size:18px;font-weight:bold;margin-bottom:18px;">Welk type verdeler wil je plaatsen?</div>
    <button data-existing-divider-type="aftakker" class="blue" style="width:100%;margin-bottom:10px;display:none;">OT-VCX-??/PTP (t-o)</button>
    <button data-existing-divider-type="twee" class="blue" style="width:100%;margin-bottom:10px;display:none;">OS-VCX-02/PTP</button>
    <button data-existing-divider-type="een" class="blue" style="width:100%;margin-bottom:10px;display:none;">OT-VCX-00/PTP</button>
    <button data-main-divider="OS-VCX-02/TP" data-divider-count="2" class="blue" style="width:100%;margin-bottom:10px;">OS-VCX-02/TP</button>
    <button data-main-divider="OS-VCX-02/PTP" data-divider-count="2" class="blue" style="width:100%;margin-bottom:10px;">OS-VCX-02/PTP</button>
    <button data-main-divider="4OSP402" data-divider-count="2" class="blue" style="width:100%;margin-bottom:10px;">4OSP402</button>
    <button data-main-divider="OS-VCX-03/PTP" data-divider-count="3" class="blue" style="width:100%;margin-bottom:10px;display:none;">OS-VCX-03/PTP</button>
    <button data-main-divider="40SP603" data-divider-count="3" class="blue" style="width:100%;margin-bottom:10px;display:none;">40SP603</button>
    <button data-main-divider="40SP803" data-divider-count="3" class="blue" style="width:100%;margin-bottom:10px;display:none;">40SP803</button>
    <button data-main-divider="OS-VCX-04/PTP" data-divider-count="4" class="blue" style="width:100%;margin-bottom:10px;display:none;">OS-VCX-04/PTP</button>
    <button data-main-divider="40SP804" data-divider-count="4" class="blue" style="width:100%;margin-bottom:10px;display:none;">40SP804</button>
    <button id="dividerCustomType" class="blue" style="width:100%;margin-bottom:10px;">Ander type...</button>
    <button id="dividerChoiceCancel" class="red" style="width:100%;">Annuleer</button>
  </div>
</div>

<div id="extraMaterialPanel" style="display:none;position:fixed;inset:0;z-index:120;background:rgba(0,0,0,.45);align-items:center;justify-content:center;">
  <div style="background:#fff;width:min(500px,92vw);padding:24px;border-radius:12px;box-shadow:0 8px 30px #0005;text-align:center;">
    <div style="font-size:18px;font-weight:bold;margin-bottom:16px;">Extra materiaal</div>
    <select id="extraMaterialSelect" style="width:100%;padding:11px;font-size:16px;margin-bottom:12px;">
      <option value="">Kies materiaal...</option>
      <option>Kasten Sweepen</option>
      <option>Patchcord E2000-LC/PC (3mtr)</option>
      <option>Patchcord LC/APC-LC/PC (3mtr)</option>
      <option>Patchcord LC/PC-LC/PC DP(3mtr)</option>
      <option>Teleste 8Ch DWDM MUX</option>
      <option>SFP Ch 39+36 (node1)</option>
      <option>SFP Ch 33+28 (node2)</option>
      <option>SFP Ch 26+24 (node3)</option>
      <option>SFP Ch 22+21 (node4)</option>
      <option>SFP-10GE-BX10D-3327-i</option>
      <option>SFP-10GE-BX40D</option>
      <option>Ring PG11 -5/8&quot; tbv node</option>
      <option>Jumper 5/8&quot;-PIN 76cm</option>
      <option>Aansluitset televoeding</option>
      <option>Ring 5/8&quot; - PG11 tbv televoeding</option>
    </select>
    <input id="extraMaterialQty" type="number" min="1" step="1" placeholder="Aantal" style="width:100%;padding:11px;font-size:16px;margin-bottom:12px;">
    <div style="display:flex;gap:10px;">
      <button id="extraMaterialAdd" class="blue" style="flex:1;">Toevoegen</button>
      <button id="extraMaterialCancel" class="red" style="flex:1;">Annuleer</button>
    </div>
  </div>
</div>

<div id="manualDaaPanel" style="display:none;position:fixed;inset:0;z-index:110;background:rgba(0,0,0,.45);align-items:center;justify-content:center;">
  <div style="background:#fff;width:min(430px,92vw);padding:24px;border-radius:12px;box-shadow:0 8px 30px #0005;text-align:center;">
    <div style="font-size:18px;font-weight:bold;margin-bottom:16px;">DAA plaatsen</div>
    <div style="margin-bottom:10px;">Kies DAA:</div>
    <select id="manualDaaNumber" style="width:100%;padding:10px;font-size:16px;margin-bottom:14px;">
      <option value="DAA001">DAA001</option><option value="DAA002">DAA002</option><option value="DAA003">DAA003</option><option value="DAA004">DAA004</option>
      <option value="DAA005">DAA005</option><option value="DAA006">DAA006</option><option value="DAA007">DAA007</option><option value="DAA008">DAA008</option>
    </select>
    <div style="margin-bottom:8px;">Uitgang 1 wordt standaard geplaatst.</div>
    <div style="display:flex;gap:10px;justify-content:center;margin-bottom:16px;">
      <button id="manualOut1" class="blue" style="display:none">Uitgang 1</button>
      <button id="manualOut2" class="blue" style="display:none">Uitgang 2</button>
    </div>
    <button id="manualDaaContinue" class="blue" style="width:100%;margin-bottom:10px;">Start DAA plaatsen</button>
    <button id="manualDaaCancel" class="red" style="width:100%;margin-top:10px;">Annuleer</button>
  </div>
</div>

<div id="workspace">
  <div id="page">
    <canvas id="pdfCanvas"></canvas>
    <canvas id="overlay"></canvas>
  </div>
</div>

<script type="module">
import * as pdfjsLib from
  "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/4.4.168/pdf.min.mjs";

pdfjsLib.GlobalWorkerOptions.workerSrc =
  "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/4.4.168/pdf.worker.min.mjs";

const fileInput=document.getElementById("pdfFile");
const schemaFolderBtn=document.getElementById("schemaFolderBtn");
const excelFileInput=document.getElementById("excelFile");
const excelBtn=document.getElementById("excelBtn");
const extraMaterialBtn=document.getElementById("extraMaterialBtn");
const extraMaterialPanel=document.getElementById("extraMaterialPanel");
const extraMaterialSelect=document.getElementById("extraMaterialSelect");
const extraMaterialQty=document.getElementById("extraMaterialQty");
const extraMaterialAdd=document.getElementById("extraMaterialAdd");
const extraMaterialCancel=document.getElementById("extraMaterialCancel");
let extraMaterials=new Map();
const pdfCanvas=document.getElementById("pdfCanvas");
const overlay=document.getElementById("overlay");
let ctx=overlay.getContext("2d");
overlay.style.touchAction="none";
overlay.style.cursor="grab";

const pageBox=document.getElementById("page");
const pageInfo=document.getElementById("pageInfo");
const status=document.getElementById("status");
const stepEl=document.getElementById("step");
const addressInfo=document.getElementById("addressInfo");
const found=document.getElementById("found");

const daaSelect=document.getElementById("daaSelect");
const startDaa=document.getElementById("startDaa");
const autoWca=document.getElementById("autoWca");
const autoPni=document.getElementById("autoPni");
const startDividerBtn=document.getElementById("startDivider");
const startAftakkerBtn=document.getElementById("startAftakker");
const startDivider2Btn=document.getElementById("startDivider2");
const startDivider3Btn=document.getElementById("startDivider3");
const startDivider4Btn=document.getElementById("startDivider4");
const startExistingDividerBtn=document.getElementById("startExistingDivider");
const dividerChoicePanel=document.getElementById("dividerChoicePanel");
const dividerChoiceCancel=document.getElementById("dividerChoiceCancel");
const deleteDaaBtn=document.getElementById("deleteDaaBtn");
const eraserBtn=document.getElementById("eraserBtn");
const saveBtn=document.getElementById("saveBtn");
const clearBtn=document.getElementById("clearBtn");
const prevBtn=document.getElementById("prev");
const nextBtn=document.getElementById("next");

const daaQuestion=document.getElementById("daaQuestion");
const daaQuestionText=document.getElementById("daaQuestionText");
const daaYes=document.getElementById("daaYes");
const daaNo=document.getElementById("daaNo");
const manualDaaPanel=document.getElementById("manualDaaPanel");
const manualDaaNumber=document.getElementById("manualDaaNumber");
const manualOut1=document.getElementById("manualOut1");
const manualOut2=document.getElementById("manualOut2");
const manualDaaContinue=document.getElementById("manualDaaContinue");
const manualDaaCancel=document.getElementById("manualDaaCancel");
let manualSelectedOutput=1;


let pdf=null;
let currentPage=1;
let scale=1.5;

// Map met de originele blokschema's. Na opslaan van een schema wordt
// automatisch het volgende PDF-bestand uit deze map geopend.
let blokschemaDirectoryHandle=null;
let blokschemaFiles=[];
let currentBlokschemaIndex=-1;
// Type van de geselecteerde bronmap: ORCA gebruikt WCA-herkenning, PNI niet.
let blokschemaType="ORCA";

let textItems=[];
let wcaItems=[];
let addressText="";
let manualAddressConfirmed=false;
let manualAddressPopupOpen=false;
let fileNameText="";

// Gegevens uit het gekozen Excelbestand. De kolom Node koppelt aan de PDF-bestandsnaam.
let excelRowsByNode=new Map();
let excelHeaders=[];
let currentExcelRow=null;
let excelFileName="";
let excelFileHandle=null;

let changes=[];
// Volgorde waarin componenten zijn toegevoegd. Backspace verwijdert alleen
// het laatst geplaatste component in plaats van alle wijzigingen.
let componentHistory=[];

// Gum-status moet al bestaan voordat een Excel-import of renderfunctie
// drawOverlay() kan aanroepen. Anders kan de JavaScript TDZ-fout
// "Cannot access 'eraserStrokes' before initialization" ontstaan.
var eraserActive=false;
var eraserDrag=null;
var eraserStrokes=[];
var ERASER_RADIUS=22;

/*
  Workflow:
  1. PDF laden
  2. adres automatisch linksboven tekenen
  3. DAA kiezen
  4. klik op plaats waar DAA moet komen
  5. klik lijn uitgang 1
  6. klik lijn uitgang 2
  7. daarna worden oude rondjes "verwijderd" door een witte afdeklaag
*/

let workflow={
  active:false,
  daa:null,
  daaPoint:null,
  output1Point:null,
  output2Point:null,
  mode:null
};

// Handmatig corrigeren na automatische plaatsing.
// Een automatische DAA kan worden gesleept naar een andere WCA-lijn.
let dragDaa=null;
let dragDaaConnection=null;
let hoveredDaa=null;

// Alle DAA's worden op dezelfde horizontale hoogte geplaatst.
// De eerste geplaatste DAA bepaalt de hoogte.
let daaRowY = null;

// De bovenkant van het DAA-symbool staat standaard 2 cm
// vanaf de bovenkant van de PDF-pagina.
const DAA_TOP_CM = 1;

// Hoogte van de DAA-box boven het middelpunt in drawDAA().
const DAA_TOP_OFFSET_PX = 80;

function getDaaStandardY(){
  // PDF wordt gerenderd met scale=1.5.
  // 1 inch = 2.54 cm en PDF = 72 punten per inch.
  const pointsPerCm = 72 / 2.54;
  return (DAA_TOP_CM * pointsPerCm * scale) + DAA_TOP_OFFSET_PX;
}

// Eén vaste horizontale rij voor alle groene verbindingen.
// De eerste gekozen aansluitlijn bepaalt deze hoogte.
let greenRowY = null;

// Verdelers worden op dezelfde hoogte als de DAA's geplaatst.
let dividerRowY = null;
let dividerGreenRowY = null;
let dividers = [];
let aftakkers = [];


/* ---------------- EXCEL VULLEN MET GROENE VERDELERS ---------------- */
function normalizeExcelText(v){
  return String(v ?? "").trim().toLowerCase().replace(/\s+/g," ");
}

function getGreenDividerCounts(){
  const counts = new Map();
  // Alleen gewone/groene verdelers tellen. De zwarte "Bestaande verdeler"
  // heeft existingDivider=true en wordt bewust niet meegenomen.
  for(const d of dividers){
    if(!d || d.existingDivider) continue;
    const type = String(d.type || "").trim();
    if(!type) continue;
    counts.set(type,(counts.get(type)||0)+1);
  }
  return counts;
}

// De tekst in rij 1 van de materiaal-Excel is leidend.
// De typen in de app worden naar de juiste materiaal-kop gemapt.
// De koppen in rij 1 van de Excel zijn leidend.
// Er wordt NIET meer automatisch OS-VCX naar 4OSP402/40SP402 gemapt.
// Daardoor komt ieder type exact onder zijn eigen kolom terecht.
const excelDividerHeaderMap = {};

function findExcelMaterialHeader(headers, wantedHeader){
  const wanted=normalizeExcelText(wantedHeader);
  // Eerst exact genormaliseerd zoeken.
  for(let i=0;i<headers.length;i++){
    if(normalizeExcelText(headers[i])===wanted) return i;
  }
  // Kleine tolerantie voor meerdere spaties / hoofdletters.
  const compact=v=>normalizeExcelText(v).replace(/[^a-z0-9]+/g,"");
  const wc=compact(wantedHeader);
  for(let i=0;i<headers.length;i++){
    if(compact(headers[i])===wc) return i;
  }
  // In Excel kan achter het type nog extra tekst staan, bijvoorbeeld
  // "4OSP402 2WAY". Die extra tekst moet volledig worden genegeerd.
  // Ook kleine schrijf-/spatieverschillen in de kop mogen geen probleem zijn.
  const wantedCompact=compact(wantedHeader);
  // 4OSP402 wordt in sommige Excelbestanden als 40SP402 geschreven
  // (O en 0 lijken hier hetzelfde). Beide schrijfwijzen moeten naar
  // dezelfde kolom kunnen wijzen. Andere typen worden niet omgezet.
  const aliases = new Set([wantedCompact]);
  if(wantedCompact === "4osp402") aliases.add("40sp402");
  if(wantedCompact === "40sp402") aliases.add("4osp402");

  for(let i=0;i<headers.length;i++){
    const headerCompact=compact(headers[i]);
    for(const alias of aliases){
      if(headerCompact.startsWith(alias)) return i;
    }
  }
  return -1;
}

async function fillExcelWithGreenDividers(file, fileHandle=null, silent=false){
  const XLSXlib=window.XLSX;
  if(!XLSXlib || typeof XLSXlib.read!=="function" || !XLSXlib.utils){
    throw new Error("De Excel-bibliotheek is niet geladen. Ververs de pagina en probeer opnieuw.");
  }

  const counts=getGreenDividerCounts();
  const hasDaaMaterial=[...changes].some(c=>/^DAA\d{3}$/i.test(String(c?.daa||"").trim()));
  if(!counts.size && !hasDaaMaterial && !extraMaterials.size){
    if(!silent) alert("Er zijn nog geen verdelers, DAA's of extra materialen geplaatst.");
    return false;
  }

  const buffer=await file.arrayBuffer();
  const workbook=XLSXlib.read(buffer,{type:"array",cellDates:false,cellStyles:true});
  const sheetName=workbook.SheetNames[0];
  const sheet=workbook.Sheets[sheetName];
  // Bewaar rij 1 en de kolom-/rij-opmaak exact zoals die in het originele
  // Excelbestand staat. Deze mogen door de app nooit opnieuw worden opgebouwd.
  const originalHeaderCells={};
  const originalHeaderCols=sheet["!cols"] ? sheet["!cols"].map(c=>c ? {...c} : c) : sheet["!cols"];
  const originalHeaderRow=sheet["!rows"] && sheet["!rows"][0] ? {...sheet["!rows"][0]} : null;
  if(sheet["!ref"]){
    const rr=XLSXlib.utils.decode_range(sheet["!ref"]);
    for(let c=rr.s.c;c<=rr.e.c;c++){
      const a=XLSXlib.utils.encode_cell({r:0,c});
      if(sheet[a]) originalHeaderCells[a]={...sheet[a]};
    }
  }
  const raw=XLSXlib.utils.sheet_to_json(sheet,{header:1,defval:""});
  if(!raw.length || !raw[0].length) throw new Error("Het Excelbestand bevat geen gegevens in rij 1.");
  const headers=raw[0].map(v=>String(v??"").trim());
  // Zoek eerst of deze Node al in kolom A voorkomt.
  // Als de Node bestaat, werk dan uitsluitend de aantallen op diezelfde rij bij.
  // Alleen een nieuwe Node krijgt de eerstvolgende lege gegevensrij.
  const normalizeExcelNode=value=>String(value??'')
    .toUpperCase()
    .replace(/\.PDF$/i,'')
    .replace(/\s+/g,'')
    .replace(/[^A-Z0-9-]/g,'');

  const wantedNode=normalizeExcelNode(fileNameText);
  let targetRowIndex=-1;
  let existingNodeRow=false;

  if(wantedNode){
    for(let r=1;r<raw.length;r++){
      const excelNode=normalizeExcelNode(raw[r]?.[0]);
      if(excelNode && excelNode===wantedNode){
        targetRowIndex=r;
        existingNodeRow=true;
        break;
      }
    }
  }

  if(targetRowIndex<0){
    targetRowIndex=1;
    for(let r=1;r<raw.length;r++){
      const row=raw[r]||[];
      const occupied=row.some(v=>String(v??'').trim()!=='');
      if(!occupied){ targetRowIndex=r; break; }
      targetRowIndex=r+1;
    }
  }
  while(raw.length<=targetRowIndex) raw.push([]);

  // Bij een bestaande Node blijven kolom A (Node), kolom B (adres)
  // en alle overige niet-materiaalkolommen onaangeroerd.
  // Alleen bij een nieuwe Node vullen we Node en adres in.
  if(!existingNodeRow){
    if(fileNameText) raw[targetRowIndex][0]=String(fileNameText).trim();
    if(addressText) raw[targetRowIndex][1]=String(addressText).trim();
  }

  // Rij 1 bevat de materiaaltypen. Het aantal komt in rij 2 onder dezelfde kolom.
  const mappedCounts=new Map();
  const unknown=[];
  // Verdelers: exact dezelfde tekst zoeken als in rij 1 van Excel.
  for(const [appType,count] of counts.entries()){
    const col=findExcelMaterialHeader(headers,appType);
    if(col<0){ unknown.push(appType); continue; }
    const target=headers[col];
    mappedCounts.set(target,(mappedCounts.get(target)||0)+count);
  }
  // Commscope DC2182: één stuk per geplaatste DAA.
  // Dus DAA001 = 1, DAA001 + DAA002 = 2,
  // DAA001 + DAA002 + DAA003 = 3, enzovoort.
  const geplaatsteDaa = new Set();
  for(const c of changes){
    const daa=String(c?.daa||'').trim().toUpperCase();
    if(/^DAA\d{3}$/.test(daa)) geplaatsteDaa.add(daa);
  }
  const dc2182Count=geplaatsteDaa.size;
  if(dc2182Count>0){
    const col=findExcelMaterialHeader(headers,'Commscope DC2182');
    if(col<0) unknown.push('Commscope DC2182');
    else {
      const target=headers[col];
      mappedCounts.set(target,(mappedCounts.get(target)||0)+dc2182Count);
    }
  }

  // Extra materiaal wordt eveneens exact op de kop uit rij 1 gezet.
  for(const [material,count] of extraMaterials.entries()){
    const col=findExcelMaterialHeader(headers,material);
    if(col<0){ unknown.push(material); continue; }
    const target=headers[col];
    mappedCounts.set(target,(mappedCounts.get(target)||0)+count);
  }

  // Een nul is in de materiaallijst niet gewenst: laat nullen in de nieuwe rij leeg.
  for(let c=0;c<(raw[targetRowIndex]?.length||0);c++){
    const v=raw[targetRowIndex][c];
    if(v===0 || String(v).trim()==='0') raw[targetRowIndex][c]='';
  }

  // Bestaande Node: aantallen vervangen (niet optellen bij de oude telling).
  // Wis alleen kolommen die als materiaaltype in rij 1 zijn benoemd;
  // Node, adres en andere velden blijven behouden.
  if(existingNodeRow){
    const knownMaterialTypes=[
      ...counts.keys(),
      'Commscope DC2182',
      ...extraMaterials.keys()
    ];
    for(const type of knownMaterialTypes){
      const col=findExcelMaterialHeader(headers,type);
      if(col>=0) raw[targetRowIndex][col]='';
    }
  }

  let updated=0;
  const notFound=[];
  for(const [target,count] of mappedCounts.entries()){
    const col=findExcelMaterialHeader(headers,target);
    if(col<0){ notFound.push(target); continue; }
    raw[targetRowIndex][col]=count;
    updated++;
  }

  if(updated===0){
    throw new Error(
      "Geen passende materiaal-koppen gevonden in rij 1.\n\n"+
      "De app zoekt naar:\n"+[...mappedCounts.keys()].join("\n")+
      "\n\nVoorbeeld van de verwachte koppen:\n40SP402 2WAY\n40SP603 3WAY\n40SP803 unbalanced\n40SP804 4WAY"
    );
  }

  // Werk ALLEEN rij 2 bij. Rij 1 wordt hier bewust helemaal niet aangeraakt.
  // Daardoor blijven de bestaande opmaak, verticale tekst, tekstterugloop,
  // rijhoogte en kolombreedtes van rij 1 intact.
  const newSheet=sheet;
  const targetRow=raw[targetRowIndex]||[];
  for(let c=0;c<targetRow.length;c++){
    const value=targetRow[c];
    const addr=XLSXlib.utils.encode_cell({r:targetRowIndex,c});
    if(value===undefined || value==="") {
      // Alleen een bestaande waarde wissen als wij die bewust op leeg hebben gezet.
      if(newSheet[addr]) {
        newSheet[addr]={...newSheet[addr],v:"",t:"s"};
      }
      continue;
    }
    // Bestaande celstijl van rij 2 behouden; alleen de waarde wijzigen.
    newSheet[addr]={...(newSheet[addr]||{}),v:value,t:typeof value==="number"?"n":"s"};
  }
  // Zorg dat Excel rij 2 ook binnen het bestaande werkbladbereik (!ref) ziet.
  // Zonder dit kan een oorspronkelijk lege rij 2 onzichtbaar blijven na het opslaan.
  const currentRef = newSheet["!ref"];
  if(currentRef){
    const range = XLSXlib.utils.decode_range(currentRef);
    range.s.r = Math.min(range.s.r, 0);
    range.e.r = Math.max(range.e.r, targetRowIndex);
    range.e.c = Math.max(range.e.c, Math.max(0, headers.length-1));
    newSheet["!ref"] = XLSXlib.utils.encode_range(range);
  }else{
    newSheet["!ref"] = XLSXlib.utils.encode_range({s:{r:0,c:0},e:{r:targetRowIndex,c:Math.max(0,headers.length-1)}});
  }
  // Zet de originele rij-1-cellen en hun opmaak terug. Daardoor blijft o.a.
  // tekst 90 graden gedraaid, tekstterugloop, lettertype en uitlijning exact behouden.
  for(const [addr,cell] of Object.entries(originalHeaderCells)){
    newSheet[addr]={...cell};
  }
  if(originalHeaderCols) newSheet["!cols"]=originalHeaderCols.map(c=>c ? {...c} : c);
  if(originalHeaderRow){
    newSheet["!rows"]=newSheet["!rows"] || [];
    newSheet["!rows"][0]={...originalHeaderRow};
  }
  workbook.Sheets[sheetName]=newSheet;
  const outArray=XLSXlib.write(workbook,{bookType:"xlsx",type:"array",cellStyles:true});

  // Opslaan in hetzelfde gekozen Excelbestand via de File System Access API.
  // Dit voorkomt een aparte download en bewaart het originele bestand.
  if(!fileHandle){
    if(typeof window.showOpenFilePicker!=="function"){
      throw new Error("Deze browser ondersteunt rechtstreeks opslaan in het gekozen Excelbestand niet. Gebruik Microsoft Edge of Chrome.");
    }
    const handles=await window.showOpenFilePicker({
      multiple:false,
      types:[{description:"Excel-bestand",accept:{"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet":[".xlsx"],"application/vnd.ms-excel":[".xls"]}}]
    });
    fileHandle=handles[0];
  }

  if(fileHandle.queryPermission){
    let perm=await fileHandle.queryPermission({mode:"readwrite"});
    if(perm!=="granted") perm=await fileHandle.requestPermission({mode:"readwrite"});
    if(perm!=="granted") throw new Error("Geen toestemming om het Excelbestand op te slaan.");
  }
  const writable=await fileHandle.createWritable();
  await writable.write(outArray);
  await writable.close();

  const overzicht=[...counts.entries()].map(([type,aantal])=>`${type}: ${aantal}x`).join("\n");
  const materiaalOverzicht=[...mappedCounts.entries()].map(([type,aantal])=>`${type}: ${aantal}x`).join("\n");
  status.textContent=existingNodeRow
    ? `Bestaande Node ${fileNameText} bijgewerkt op rij ${targetRowIndex+1}: ${updated} materiaal-koppen aangepast.`
    : `Nieuwe Node ${fileNameText} toegevoegd op rij ${targetRowIndex+1}: ${updated} materiaal-koppen ingevuld.`;
  stepEl.textContent=existingNodeRow
    ? `Bestaande Excel-rij ${targetRowIndex+1} is bijgewerkt met de nieuwe aantallen.`
    : `Nieuwe Excel-rij ${targetRowIndex+1} is aangemaakt met de aantallen.`;

  let waarschuwing="";
  if(notFound.length) waarschuwing+=`\n\nNiet gevonden in rij 1:\n${notFound.join("\n")}`;
  if(unknown.length) waarschuwing+=`\n\nOnbekend app-type:\n${unknown.join("\n")}`;

  if(!silent){
    alert(`Excel is opgeslagen in het gekozen bestand. ${existingNodeRow?"Bestaande rij bijgewerkt":"Nieuwe rij toegevoegd"}: rij ${targetRowIndex+1}.\n\nNode: ${fileNameText||"(niet gevonden)"}\nAdres: ${addressText||"(niet gevonden)"}\n\nGetekende groene verdelers:\n${overzicht}\n\nIn Excel ingevuld:\n${materiaalOverzicht}${waarschuwing}`);
  }
  return true;
}

/* ---------------- EXCEL ADRESSEN / MATERIALEN ---------------- */

function normalizeNode(value){
  let v=String(value??"").trim();
  if(!v) return "";
  v=v.replace(/\.(pdf|PDF|xlsx|xlsm|xls)$/i,"");
  v=v.replace(/\s*\([^)]*\)\s*$/," ").trim();
  // Bij nieuwe Excel kunnen meerdere nodes in kolom A staan, bijvoorbeeld
  // "ASD-LC0101 ASD-WC0107". Voor de koppeling gebruiken we het WC-node.
  const matches=v.toUpperCase().match(/\b(?:[A-Z]{2,10}-)?WC\d{3,6}\b/g);
  if(matches && matches.length) return matches[matches.length-1];
  return v.toUpperCase().split(/\s+/)[0];
}

function getNodeAliases(value){
  const raw=String(value??"").trim();
  if(!raw) return [];
  const cleaned=raw.replace(/\.(pdf|PDF|xlsx|xlsm|xls)$/i,"").replace(/\s*\([^)]*\)\s*$/," ").trim();
  const aliases=[];
  const add=v=>{const k=String(v||"").trim().toUpperCase(); if(k && !aliases.includes(k)) aliases.push(k);};
  const matches=cleaned.toUpperCase().match(/\b(?:[A-Z]{2,10}-)?WC\d{3,6}\b/g) || [];
  matches.forEach(add);
  add(cleaned);
  return aliases;
}

function normalizeHeader(value){
  return String(value??"")
    .trim()
    .replace(/\s+/g," ")
    .toLowerCase();
}

function formatMaterialName(header){
  const h=String(header??"").trim();
  if(normalizeHeader(h)==="daa node") return "DAA";
  return h;
}

function formatMaterialValue(value){
  if(value===null || value===undefined || String(value).trim()==="") return "";
  if(typeof value==="number" && Number.isFinite(value)) return `${value} st`;
  const text=String(value).trim();
  if(/^\d+(?:[.,]\d+)?$/.test(text)) return `${text} st`;
  return text;
}

function getCurrentPdfName(){
  // Lees altijd de actuele file-input uit. Dit voorkomt dat bij het kiezen
  // van een tweede/derde PDF nog de Node van de vorige PDF wordt gebruikt.
  const f=fileInput && fileInput.files && fileInput.files[0];
  if(f && f.name) return f.name;
  return fileNameText || "";
}

function findExcelRowForPdf(){
  if(!excelRowsByNode.size) return {key:"", row:null};

  const rawName=getCurrentPdfName();
  const stem=String(rawName).replace(/\.[^.]+$/i,"").trim();
  const candidates=[];

  const addCandidate=v=>{
    const k=normalizeNode(v);
    if(k && !candidates.includes(k)) candidates.push(k);
  };

  // Haal de echte Node rechtstreeks uit de bestandsnaam.
  // Voorbeeld: "GV-WC3082 blokschema.pdf" -> "GV-WC3082".
  const nodeMatches=stem.toUpperCase().match(/\b(?:[A-Z]{2,10}-)?WC\d{3,6}\b/g) || [];
  nodeMatches.forEach(addCandidate);

  // Ook de tekst zonder "blokschema" proberen.
  const cleanStem=stem
    .replace(/\bblokschema\b/ig," ")
    .replace(/\s+/g," ")
    .trim();
  addCandidate(cleanStem);
  addCandidate(fileNameText);

  // Zoek eerst een exacte Node-match. Nooit een numerieke Excel-rij
  // zoals "30" als Node gebruiken wanneer de PDF een WC-node bevat.
  for(const [node,row] of excelRowsByNode.entries()){
    const n=normalizeNode(node);
    if(!n || /^\\d+$/.test(n)) continue;
    if(candidates.includes(n)) return {key:n,row};
  }

  // Als de PDF bijvoorbeeld WC3082 heet en Excel GV-WC3082 bevat,
  // koppel uitsluitend op de WC-code.
  const pdfWc=(stem.toUpperCase().match(/\bWC\d{3,6}\b/)||[])[0] || "";
  if(pdfWc){
    for(const [node,row] of excelRowsByNode.entries()){
      const n=normalizeNode(node);
      if(!n || /^\\d+$/.test(n)) continue;
      const excelWc=(n.match(/\bWC\d{3,6}\b/)||[])[0] || "";
      if(excelWc===pdfWc) return {key:n,row};
    }
  }

  return {key:candidates[0]||normalizeNode(stem),row:null};
}

function tryFindNodeInPdfText(){
  // Als de bestandsnaam geen directe match geeft, probeer de Node uit de
  // tekstlaag van de PDF te halen. Dit is alleen een fallback; Excel blijft
  // leidend.
  const re=/\b(?:[A-Z]{2,10}-)?WC\d{3,6}\b/ig;
  for(const item of textItems){
    const m=String(item.text||"").match(re);
    if(m){
      const k=normalizeNode(m[0]);
      if(excelRowsByNode.has(k)) return k;

      // Ook bij een kale PDF-node zoals WC3038 zoeken we naar
      // een Excel-node met dezelfde WC-code, bijvoorbeeld ASD-WC3038.
      const wc=(String(k).match(/\bWC\d{3,6}\b/)||[])[0] || "";
      if(wc){
        for(const [node] of excelRowsByNode.entries()){
          const excelNode=normalizeNode(node);
          const excelWc=(String(excelNode).match(/\bWC\d{3,6}\b/)||[])[0] || "";
          if(excelWc===wc) return excelNode;
        }
      }
    }
  }
  return "";
}

function applyExcelDataForCurrentPdf(){
  currentExcelRow=null;
  if(!excelRowsByNode.size) return false;

  const match=findExcelRowForPdf();
  let key=match.key;
  currentExcelRow=match.row;

  if(!currentExcelRow){
    const pdfNode=tryFindNodeInPdfText();
    if(pdfNode){
      key=pdfNode;
      currentExcelRow=excelRowsByNode.get(pdfNode)||null;
    }
  }

  if(!currentExcelRow){
    addressInfo.innerHTML=
      `<span class="warn">Geen Excel-koppeling</span><br>Node <b>${escapeHtml(key||"onbekend")}</b> staat niet in de Excel.`;
    found.innerHTML=`<b>Geen Excel-rij gevonden:</b> ${escapeHtml(key||"onbekend")}`;
    status.textContent=`PDF ${key||"onbekend"} is niet gevonden in de ingelezen Excel (${excelRowsByNode.size} rijen).`;
    return false;
  }

  // Nieuwe Excel-opmaak: kolom A = Node, kolom B = Adres.
  const addressHeader=excelHeaders[1] || "Adres";
  const addressValue=currentExcelRow[addressHeader] ?? "";

  if(addressValue!==null && addressValue!==undefined && String(addressValue).trim()){
    addressText=String(addressValue).trim();
    addressInfo.innerHTML=
      `<span class="ok">Adres uit Excel:</span><br>${escapeHtml(addressText)}`;
  }else{
    addressText="";
    addressInfo.innerHTML=
      `<span class="warn">Node gevonden, maar kolom B (Adres) is leeg.</span>`;
  }

  found.innerHTML=
    `<b>Excel gekoppeld:</b> ${escapeHtml(key)} — ${escapeHtml(addressText||"(geen adres)")}`;
  status.textContent=`Excel gekoppeld aan ${key}. Materialen: kolom C t/m BI.`;
  return true;
}

function drawExcelMaterials(){
  if(!currentExcelRow || currentPage!==1) return;

  const rows=[];
  // Materialen staan in de Excel vanaf kolom C t/m BI.
  // BI is kolom 61; slice gebruikt een exclusieve eindindex, dus 61.
  const materialHeaders=excelHeaders.slice(2,61);
  for(const header of materialHeaders){
    // Deze twee Excel-kolommen mogen niet in de materiaaltabel komen.
    // We vergelijken genormaliseerd zodat bijvoorbeeld "Bestaande WC",
    // "Bestaande WCA" of extra spaties ook worden uitgesloten.
    const normalizedMaterialHeader=normalizeHeader(header).replace(/[^a-z0-9]/g, "");
    // Sluit de kolom voor bestaand(e) WC/WCA altijd uit.
    // In de Excel komt dit bijvoorbeeld voor als:
    // "Bestaand WCA", "Bestaande WCA", "Bestaand WC en WCA"
    // of met extra spaties/tekens.
    const isExistingWcWca =
      /bestaand(e)?wc(?:enwca)?/.test(normalizedMaterialHeader) ||
      /bestaand(e)?wca/.test(normalizedMaterialHeader) ||
      normalizedMaterialHeader === "wca";
    if(isExistingWcWca || normalizedMaterialHeader.includes("lijstnodig")) continue;

    const value=currentExcelRow[header];
    if(value===null || value===undefined || String(value).trim()==="") continue;

    const formatted=formatMaterialValue(value);
    if(!formatted) continue;

    const numeric=typeof value==="number" ? value : Number(String(value).replace(",","."));
    if(Number.isFinite(numeric) && numeric===0) continue;

    rows.push({name:formatMaterialName(header),value:formatted});
  }

  if(!rows.length) return;

  // Materiaalstaat direct onder bestandsnaam/adres linksboven.
  // Alle materialen t/m BI worden verdeeld over meerdere kolommen, zodat
  // ook een lange materiaallijst binnen de PDF-pagina zichtbaar blijft.
  const margin=20;
  const rowH=20;
  const titleH=27;

  // Gebruik altijd de logische editor-afmetingen. Bij PDF-export wordt de
  // tekenlaag namelijk op 4x gerenderd; ctx.canvas.width/height zijn dan 4x
  // zo groot terwijl alle teken-coördinaten nog in editor-eenheden staan.
  // Zonder deze correctie verdwijnt de materiaaltabel bij export onder de
  // onderrand van de PDF.
  const logicalW=ctx.canvas._logicalWidth || ctx.canvas.width;
  const logicalH=ctx.canvas._logicalHeight || ctx.canvas.height;

  const cols=Math.min(3, Math.max(1, Math.floor((logicalW-2*margin)/180)));
  const rowsPerCol=Math.ceil(rows.length/cols);
  const tableW=Math.min(760,logicalW-2*margin);
  const colW=tableW/cols;
  const tableH=titleH+rowsPerCol*rowH+8;
  const x=margin;

  // Plaats de materiaaltabel linksonder op de PDF.
  // Eerst de uiteindelijke hoogte bepalen zodat de tabel netjes binnen de pagina blijft.
  const maxTableH=logicalH-2*margin;
  const effectiveRowH=Math.min(rowH, Math.max(12,(maxTableH-titleH-8)/Math.max(1,rowsPerCol)));
  const effectiveTableH=titleH+rowsPerCol*effectiveRowH+8;
  const safeY=Math.max(margin,logicalH-effectiveTableH-margin);

  ctx.save();
  ctx.fillStyle="#fff";
  ctx.strokeStyle="#555";
  ctx.lineWidth=1.5;
  ctx.fillRect(x,safeY,tableW,effectiveTableH);
  ctx.strokeRect(x,safeY,tableW,effectiveTableH);

  ctx.fillStyle="#111";
  ctx.font="bold 15px Arial";
  ctx.textAlign="left";
  ctx.textBaseline="middle";
  ctx.fillText("Materialen",x+10,safeY+titleH/2);

  ctx.strokeStyle="#aaa";
  ctx.beginPath();
  ctx.moveTo(x,safeY+titleH); ctx.lineTo(x+tableW,safeY+titleH);
  for(let c=1;c<cols;c++){
    const cx=x+c*colW;
    ctx.moveTo(cx,safeY); ctx.lineTo(cx,safeY+effectiveTableH);
  }
  ctx.stroke();

  ctx.font=`${effectiveRowH<16?9:10}px Arial`;
  for(let i=0;i<rows.length;i++){
    const col=Math.floor(i/rowsPerCol);
    const row=i%rowsPerCol;
    const itemX=x+col*colW;
    const cy=safeY+titleH+row*effectiveRowH+effectiveRowH/2;
    if(row>0){
      ctx.strokeStyle="#ddd";
      ctx.beginPath();
      ctx.moveTo(itemX,safeY+titleH+row*effectiveRowH);
      ctx.lineTo(itemX+colW,safeY+titleH+row*effectiveRowH);
      ctx.stroke();
    }
    const rowData=rows[i];
    ctx.fillStyle="#111";
    ctx.textAlign="left";
    // Houd de materiaalnaam leesbaar en voorkom dat lange namen over de waarde heen lopen.
    const maxNameW=colW-55;
    let name=rowData.name;
    while(name.length>3 && ctx.measureText(name).width>maxNameW) name=name.slice(0,-1);
    if(name!==rowData.name) name=name.slice(0,-1)+"…";
    ctx.fillText(name,itemX+6,cy);
    ctx.textAlign="right";
    ctx.fillText(rowData.value,itemX+colW-6,cy);
  }
  ctx.restore();
}

function restoreExcelFromMemory(){
  try{
    const headers=JSON.parse(localStorage.getItem("orcaExcelHeaders")||"null");
    const entries=JSON.parse(localStorage.getItem("orcaExcelRowsByNode")||"null");
    if(!Array.isArray(headers)||!Array.isArray(entries)||!entries.length) return false;
    excelHeaders=headers;
    excelRowsByNode=new Map(entries);
    excelFileName=localStorage.getItem("orcaExcelFileName")||"";
    status.textContent=`Excel actief: ${excelRowsByNode.size} rijen${excelFileName?" uit "+excelFileName:""}.`;
    return true;
  }catch(e){
    console.warn("Opgeslagen Excel kon niet worden hersteld",e);
    return false;
  }
}

// De ingelezen Excel blijft beschikbaar voor meerdere PDF's en ook na een refresh.
restoreExcelFromMemory();

excelBtn.addEventListener("click",async()=>{
  try{
    if(typeof window.showOpenFilePicker!=="function") throw new Error("Rechtstreeks opslaan in Excel wordt door deze browser niet ondersteund. Gebruik Microsoft Edge of Chrome.");
    if(!excelFileHandle){
      const handles=await window.showOpenFilePicker({
        multiple:false,
        types:[{description:"Excel-bestand",accept:{"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet":[".xlsx"]}}]
      });
      excelFileHandle=handles[0];
    }
    const file=await excelFileHandle.getFile();
    await fillExcelWithGreenDividers(file,excelFileHandle);
    excelBtn.textContent="Excel gekoppeld ✓";
    status.textContent="Excel gekoppeld ✓ – automatische update staat aan.";
    stepEl.textContent="Bij Opslaan wordt Excel automatisch bijgewerkt.";
  }catch(err){
    if(err && err.name==="AbortError") return;
    console.error(err);
    alert(`Excel kon niet worden opgeslagen.\n\n${err.message||err}`);
  }
});

/* ---------------- EXTRA MATERIAAL ---------------- */
extraMaterialBtn.addEventListener("click",()=>{
  extraMaterialPanel.style.display="flex";
  extraMaterialSelect.value="";
  extraMaterialQty.value="";
  setTimeout(()=>extraMaterialSelect.focus(),50);
});
extraMaterialCancel.addEventListener("click",()=>{ extraMaterialPanel.style.display="none"; });
extraMaterialAdd.addEventListener("click",()=>{
  const material=String(extraMaterialSelect.value||"").trim();
  const qty=Math.floor(Number(extraMaterialQty.value));
  if(!material){ alert("Kies eerst een materiaal."); return; }
  if(!Number.isFinite(qty)||qty<1){ alert("Vul een geldig aantal in."); return; }
  extraMaterials.set(material,(extraMaterials.get(material)||0)+qty);
  extraMaterialPanel.style.display="none";
  status.textContent=`Extra materiaal toegevoegd: ${material} (${qty}x).`;
  stepEl.textContent="Extra materiaal is toegevoegd. Het wordt meegenomen bij Excel opslaan.";
});

/* ---------------- PDF INLEZEN ---------------- */

async function loadPdfFile(file, fromBlokschemaMap=false){
  if(!file)return;
  // Wis direct het adres van de vorige PDF en maak eventuele lopende OCR ongeldig.
  // Zo kan een late OCR-uitkomst van het vorige bestand het nieuwe adres niet overschrijven.
  window.__addressRunToken=(window.__addressRunToken||0)+1;
  addressText="";
  addressInfo.textContent="Nieuw PDF-bestand geladen. Adres wordt opnieuw met OCR gezocht...";
  const manualAddressInputEl=document.getElementById("manualAddressInput");
  if(manualAddressInputEl) manualAddressInputEl.value="";
  manualAddressConfirmed=false;
  manualAddressPopupOpen=false;

  // Nieuwe PDF = altijd een volledig schone werklaag.
  // Gummodus en oude gumstreken mogen nooit naar het volgende blokschema meegaan.
  deactivateEraser();
  eraserActive=false;
  eraserDrag=null;
  eraserStrokes=[];

  try{
    pdf=await pdfjsLib.getDocument({
      data:new Uint8Array(await file.arrayBuffer())
    }).promise;

    currentPage=1;
    // Alleen de eerste Node uit de bestandsnaam gebruiken voor de kop.
    // De Excel-koppeling zelf gebruikt de volledige actuele bestandsnaam.
    // Gebruik voor de kop en de opgeslagen PDF altijd de echte Node uit de bestandsnaam.
    // Bijvoorbeeld: "UT-WC1006 (bestaande blokschema).pdf" -> "UT-WC1006".
    const fileStem=file.name.replace(/\.[^.]+$/i, "").trim();
    const nodeMatch=fileStem.match(/\b(?:[A-Z]{2,10}-)?WC\d{3,6}\b/i);
    fileNameText=nodeMatch ? nodeMatch[0].toUpperCase() : fileStem.replace(/\s*\([^)]*\)\s*$/, "").trim();
    changes=[];
    dividers=[];
    aftakkers=[];
    componentHistory=[];
    extraMaterials=new Map();
    daaRowY=null;
    greenRowY=null;
    dividerRowY=null;
    dividerGreenRowY=null;

    workflow={
      active:false,daa:null,daaPoint:null,
      output1Point:null,output2Point:null,mode:null
    };

    // Toon de PDF direct. OCR mag het laden van de PDF niet blokkeren.
    await renderPage();

    await readAllPdfText();
    // Bij PNI gebruiken we geen WCA-aanduidingen: die worden daar handmatig
    // bepaald. De tekst-/adresherkenning blijft wel actief.
    if(fromBlokschemaMap && blokschemaType==="PNI"){
      wcaItems=[];
      const foundEl=document.getElementById("found");
      if(foundEl) foundEl.innerHTML="PNI-tekening: WCA wordt niet automatisch gezocht. Plaats DAA's handmatig.";
      autoWca.disabled=true;
    }
    // Adres komt rechtstreeks uit de PDF. Excel is hiervoor niet meer nodig.
    // Materialen kunnen nog wel uit Excel worden geladen als die is gekoppeld.
    await findAddress();
    buildDaaList();
    daaSelect.value="DAA001";
    await renderPage();

    stepEl.textContent=
      fromBlokschemaMap
        ? `Blokschema ${currentBlokschemaIndex+1} van ${blokschemaFiles.length}: controleer het adres en maak je aanwijzingen.`
        : "Stap 2: controleer het adres linksboven en kies daarna een DAA.";

    status.textContent=fromBlokschemaMap
      ? `${blokschemaType}-blokschema geladen: ${file.name}`
      : "PDF geladen.";
  }catch(err){
    console.error(err);
    status.textContent="De PDF kon niet worden geopend.";
    alert("De PDF kon niet worden geopend.\n\n"+(err.message||err));
  }
}

fileInput.addEventListener("change",async()=>{
  const file=fileInput.files[0];
  if(!file)return;
  blokschemaDirectoryHandle=null;
  blokschemaFiles=[];
  currentBlokschemaIndex=-1;
  await loadPdfFile(file,false);
});

async function chooseBlokschemaFolder(){
  if(!window.showDirectoryPicker || !window.isSecureContext){
    alert("Een map met meerdere PDF's openen werkt alleen in Microsoft Edge of Chrome via HTTPS (bijvoorbeeld GitHub Pages).\n\nGebruik de knop 'PDF kiezen' als je maar één bestand wilt openen.");
    return;
  }

  try{
    const dir=await window.showDirectoryPicker({mode:"readwrite"});

    // Vraag direct na het kiezen van de map welk type tekening erin staat.
    const typeChoice=window.prompt(
      "Welk type tekeningen staan in deze map?\n\nTyp ORCA voor automatische WCA-herkenning.\nTyp PNI voor handmatige DAA-/verdelerplaatsing.",
      blokschemaType
    );
    if(typeChoice===null) return;
    const selectedType=String(typeChoice).trim().toUpperCase();
    if(!["ORCA","PNI"].includes(selectedType)){
      alert("Ongeldige keuze. Vul ORCA of PNI in.");
      return;
    }
    blokschemaType=selectedType;

    if(dir.requestPermission){
      const permission=await dir.requestPermission({mode:"readwrite"});
      if(permission!=="granted") return;
    }

    // Zoek recursief: PDF's mogen ook in onderliggende projectmappen staan.
    // De uitvoer-PDF's worden in de gekozen hoofdmap opgeslagen.
    const files=[];
    async function collectPdfFiles(folderHandle, relativeFolder=""){
      for await(const [name,handle] of folderHandle.entries()){
        if(handle.kind==="directory"){
          await collectPdfFiles(handle, relativeFolder ? `${relativeFolder}/${name}` : name);
          continue;
        }
        if(handle.kind!=="file" || !/\.pdf$/i.test(name)) continue;
        // Door de app gemaakte uitvoerbestanden niet opnieuw als bron gebruiken.
        if(/\s*\(nieuw\)\.pdf$/i.test(name)) continue;
        const file=await handle.getFile();
        // Bewaar het relatieve pad voor herkenbare sortering en statusmeldingen.
        files.push({file, relativePath: relativeFolder ? `${relativeFolder}/${name}` : name});
      }
    }
    await collectPdfFiles(dir);
    files.sort((a,b)=>a.relativePath.localeCompare(b.relativePath,undefined,{numeric:true,sensitivity:"base"}));
    const pdfFiles=files.map(item=>item.file);

    if(!pdfFiles.length){
      alert("In deze map en de onderliggende mappen zijn geen originele PDF-blokschema's gevonden.");
      return;
    }

    blokschemaDirectoryHandle=dir;
    blokschemaFiles=pdfFiles;
    currentBlokschemaIndex=0;

    // De gekozen bronmap wordt ook de standaard opslagmap voor de gemaakte PDF's.
    rememberedSaveDirectory=dir;

    status.textContent=`${pdfFiles.length} blokschema's gevonden, ook in submappen. Nummer 1 wordt geopend.`;
    stepEl.textContent=`Blokschema 1 van ${pdfFiles.length} wordt geladen...`;
    await loadPdfFile(pdfFiles[0],true);
  }catch(err){
    if(err && err.name==="AbortError") return;
    console.error(err);
    alert("De map kon niet worden geopend.\n\n"+(err.message||err));
  }
}

async function loadNextBlokschema(){
  if(!blokschemaDirectoryHandle || !blokschemaFiles.length) return false;

  const nextIndex=currentBlokschemaIndex+1;
  if(nextIndex>=blokschemaFiles.length){
    status.textContent=`Alle ${blokschemaFiles.length} blokschema's zijn verwerkt.`;
    stepEl.textContent="Klaar. Alle blokschema's uit de gekozen map zijn opgeslagen.";
    saveBtn.disabled=true;
    return false;
  }

  currentBlokschemaIndex=nextIndex;
  status.textContent=`PDF opgeslagen. Blokschema ${currentBlokschemaIndex+1} van ${blokschemaFiles.length} wordt geopend...`;
  stepEl.textContent=`Volgende blokschema wordt geladen: ${blokschemaFiles[currentBlokschemaIndex].name}`;
  await loadPdfFile(blokschemaFiles[currentBlokschemaIndex],true);
  return true;
}

schemaFolderBtn.addEventListener("click",chooseBlokschemaFolder);

async function readAllPdfText(){
  textItems=[];
  wcaItems=[];
  addressText="";

  for(let p=1;p<=pdf.numPages;p++){
    const page=await pdf.getPage(p);
    const content=await page.getTextContent();

    for(const item of content.items){
      const text=(item.str||"").trim();
      if(!text)continue;

      textItems.push({
        page:p,
        text,
        x:item.transform[4],
        y:item.transform[5],
        width:item.width||0,
        height:Math.abs(item.transform[3]||10),
        transform:item.transform.slice()
      });
    }
  }

  /*
   * Sommige ORCA-blokschema's zijn volledig als afbeelding opgeslagen.
   * Dan heeft PDF.js geen tekstlaag en zijn adres/WCA niet uitleesbaar.
   * Gebruik in dat geval OCR als fallback.
   */
  if(!textItems.length){
    status.textContent="PDF geladen. Tekstlaag ontbreekt; OCR zoekt WCA-labels en genummerde rondjes...";
    await readPdfWithOcr();
  }

  findWcaItems();

  /*
   * Adresherkenning wordt één keer gestart vanuit loadPdfFile(), nadat
   * alle PDF-tekst/OCR is ingelezen. Zo kunnen geen twee OCR-processen
   * tegelijk draaien en kan een latere OCR-pass geen handmatig adres wissen.
   * OCR op een gedraaid beeld leest de verticale WCA-labels veel beter.
   * Als de eerste OCR-pass nog geen WCA's oplevert, doen we die extra pass.
   */
  if(!wcaItems.length && pdf){
    status.textContent="Verticale WCA-labels worden met OCR gezocht...";
    await readVerticalWcaWithOcr();
    findWcaItems();
  }
}

async function ocrCanvas(canvas, options={}){
  if(!window.Tesseract) throw new Error("OCR-bibliotheek kon niet worden geladen.");

  const result=await Tesseract.recognize(canvas,"eng",{
    ...options,
    logger:m=>{
      if(m && m.status==="recognizing text" && typeof m.progress==="number"){
        status.textContent=`OCR bezig: ${Math.round(m.progress*100)}%`;
      }
    }
  });

  return result.data;
}

function addOcrWordsToTextItems(data,pageNumber,baseScale,rotation=0){
  const words=(data && data.words)||[];
  for(const w of words){
    const text=String(w.text||"").trim();
    if(!text)continue;

    const b=w.bbox;
    if(!b)continue;

    const cx=(b.x0+b.x1)/2;
    const cy=(b.y0+b.y1)/2;

    let pdfX,pdfY;

    if(rotation===0){
      pdfX=cx/baseScale;
      const pageHeight=ocrPageHeights.get(pageNumber);
      pdfY=pageHeight-(b.y1/baseScale);
    }else if(rotation===-90){
      /*
       * 90 graden tegen de klok in:
       * rotated(x,y) = (y, W-x)
       * terug naar origineel: x=W-rotatedY, y=rotatedX.
       */
      const originalWidth=ocrPageWidths.get(pageNumber);
      pdfX=originalWidth-(b.y1/baseScale);
      pdfY=b.x0/baseScale;
    }else{
      /* fallback voor oudere 90-graden clockwise OCR */
      const originalHeight=ocrPageHeights.get(pageNumber);
      pdfX=cy/baseScale;
      pdfY=originalHeight-(b.x1/baseScale);
    }

    textItems.push({
      page:pageNumber,
      text,
      x:pdfX,
      y:pdfY,
      width:Math.max(1,(b.x1-b.x0)/baseScale),
      height:Math.max(1,(b.y1-b.y0)/baseScale),
      transform:[1,0,0,1,pdfX,pdfY],
      ocr:true,
      confidence:typeof w.confidence==="number" ? w.confidence : 0
    });
  }
}

const ocrPageHeights=new Map();
const ocrPageWidths=new Map();

async function readPdfWithOcr(){
  for(let p=1;p<=pdf.numPages;p++){
    const page=await pdf.getPage(p);
    const pdfViewport=page.getViewport({scale:1});
    ocrPageHeights.set(p,pdfViewport.height);

    const ocrScale=2.5;
    const viewport=page.getViewport({scale:ocrScale});
    const canvas=document.createElement("canvas");
    canvas.width=Math.ceil(viewport.width);
    canvas.height=Math.ceil(viewport.height);

    await page.render({
      canvasContext:canvas.getContext("2d"),
      viewport
    }).promise;

    const data=await ocrCanvas(canvas);
    addOcrWordsToTextItems(data,p,ocrScale,0);
  }
}

async function readVerticalWcaWithOcr(){
  for(let p=1;p<=pdf.numPages;p++){
    const page=await pdf.getPage(p);
    const pdfViewport=page.getViewport({scale:1});
    const ocrPageHeight=pdfViewport.height;
    const ocrScale=2.5;

    const viewport=page.getViewport({scale:ocrScale});
    const source=document.createElement("canvas");
    source.width=Math.ceil(viewport.width);
    source.height=Math.ceil(viewport.height);

    await page.render({
      canvasContext:source.getContext("2d"),
      viewport
    }).promise;

    const rotated=document.createElement("canvas");
    rotated.width=source.height;
    rotated.height=source.width;

    // PNI-WCA teksten staan verticaal. Tegen de klok in draaien levert bij
    // deze tekeningen de juiste leesrichting op (WCA001A/WCA001B).
    const rctx=rotated.getContext("2d");
    rctx.translate(0,rotated.height);
    rctx.rotate(-Math.PI/2);
    rctx.drawImage(source,0,0);

    ocrPageHeights.set(p,ocrPageHeight);
    ocrPageWidths.set(p,pdfViewport.width);
    const data=await ocrCanvas(rotated);
    addOcrWordsToTextItems(data,p,ocrScale,-90);
  }
}

/*
  Het adres wordt uit de bovenste tekstregels gehaald.
  We nemen de eerste bruikbare adresregel en de bijbehorende regel
  eronder mee. De weergegeven tekst wordt linksboven geplaatst.
*/
async function findAddress(){
  // Voorkom dubbele OCR-aanroepen en overschrijven van handmatige invoer.
  if(manualAddressConfirmed || manualAddressPopupOpen) return addressText;
  const addressRunToken=(window.__addressRunToken=(window.__addressRunToken||0)+1);
  /*
   * Dynamische adresherkenning:
   * - OCR de volledige eerste pagina; geen vaste A4-hoogte of vaste uitsnede.
   * - Zoek de zelfstandige NODE-regel (niet NODE001 in de projectcode).
   * - Zoek alleen boven die NODE naar de postcode.
   * - Zoek vervolgens de straat + het huisnummer boven die postcode.
   * - Geen postcode-uitzonderingen of hardcoded adrescorrecties.
   */
  function norm(s){
    return String(s||'').replace(/\s+/g,' ').replace(/[|]/g,'I').trim();
  }

  // ORCA: gebruik eerst de ingebouwde PDF-tekstlaag. Bij de aangeleverde
  // ORCA's staat het adres als echte PDF-tekst direct naast het grote rondje
  // (bijv. 'Kweldam 6' en 'Vogelenzangseweg 16'). OCR maakte hiervan eerder
  // foutieve teksten zoals 'H ro, I 1'.
  function addressFromOrcaPdfText(){
    if(blokschemaType!=="ORCA" || !Array.isArray(textItems) || !textItems.length) return '';

    // PDF.js gebruikt PDF-coördinaten met de oorsprong linksonder.
    // Daardoor hebben teksten bovenaan de pagina juist een HOGERE y-waarde.
    // De hoofd-NODE's staan bovenaan; adressen van aftakkingen staan lager.
    const pageHeightCache=new Map();
    const grouped=[];
    for(const item of textItems.filter(v=>v.page===1 && !v.ocr)){
      const t=norm(item.text);
      if(!t) continue;
      let line=grouped.find(v=>Math.abs(v.y-item.y)<=1.6 && item.x>=v.minX-30 && item.x<=v.maxX+30);
      if(!line){line={y:item.y,minX:item.x,maxX:item.x,parts:[]};grouped.push(line);}
      line.minX=Math.min(line.minX,item.x);
      line.maxX=Math.max(line.maxX,item.x+(item.width||0));
      line.parts.push({x:item.x,text:t});
    }

    const lines=grouped.map(v=>({
      y:v.y,
      text:norm(v.parts.sort((a,b)=>a.x-b.x).map(x=>x.text).join(' '))
    }));

    const suffix=/(?:straat|laan|weg|plein|dreef|singel|kade|gracht|pad|dijk|hof|park|baan|steeg|ring|markt|plantsoen|weide|erf|veld|dam|vaart|land|stede|schelp|berg|broek|burg|dorp|horst|zoom|wal|waard|wetering|zijde|molen|reiger|hoep|kruislaan|avercampweg|hoofdenburg)$/i;
    const tech=/\b(?:NODE|WCA\d*|RF|DS|US|AC\d{3,}|DBC\s*\d{3,}|OT-VCX|OS-VCX|PI-SPL|C3-)\b/i;
    const candidates=[];

    for(const line of lines){
      const t=line.text.replace(/[,:;]+$/,'').trim();
      if(!t || tech.test(t) || t.length>90) continue;
      const m=t.match(/^([A-Za-zÀ-ÖØ-öø-ÿ' .-]{3,75}?)\s+(\d{1,5}[A-Za-z]?(?:[-/]\d+[A-Za-z]?)?)$/);
      if(!m) continue;
      const street=norm(m[1]);
      if(!suffix.test(street) || /\d/.test(street)) continue;
      candidates.push({street,house:m[2],y:line.y});
    }
    if(!candidates.length) return '';

    // Beperk de automatische keuze tot het bovenste deel van de tekening.
    // De hoofd-NODE-adressen staan daar naast de grote ronde NODE-symbolen;
    // downstream-adressen (zoals Kameraarsweide 14) mogen niet winnen.
    const pageHeight=pdf ? pdf.getPage(1).then(p=>p.getViewport({scale:1}).height) : Promise.resolve(842);
    return pageHeight.then(height=>{
      const topCandidates=candidates.filter(c=>c.y >= height*0.55);
      if(!topCandidates.length) return '';
      topCandidates.sort((a,b)=>b.y-a.y);
      const a=topCandidates[0];
      const code=(String(fileNameText||'').match(/^([A-Z0-9]+)/i)||[])[1]||'';
      const places={EVD:'Everdingen',HGSN:'Hagestein'};
      const place=places[code.toUpperCase()]||'';
      return `${a.street} ${a.house}${place?', '+place:''}`;
    });
  }
  if(blokschemaType==="ORCA"){
    const pdfAddress=await addressFromOrcaPdfText();
    if(pdfAddress){
      addressText=pdfAddress;
      addressInfo.innerHTML=`<span class="ok">Adres automatisch uit PDF-tekst gevonden:</span><br>${escapeHtml(addressText)}`;
      return addressText;
    }
  }

  function getLines(data){
    // Gebruik eerst de echte OCR-regels van Tesseract. Woorden zelf groeperen
    // kan regels onbedoeld samenvoegen wanneer tekstblokken dicht bij elkaar staan.
    const ocrLines=(data && data.lines)||[];
    if(ocrLines.length){
      const result=ocrLines.map(line=>{
        const b=line.bbox||{};
        return {text:norm(line.text),
          x0:Number(b.x0)||0,x1:Number(b.x1)||0,
          y0:Number(b.y0)||0,y1:Number(b.y1)||0};
      }).filter(l=>l.text && l.y1>=l.y0);
      if(result.length) return result.sort((a,b)=>a.y0-b.y0);
    }

    // Fallback voor OCR-uitvoer zonder line-objecten: groepeer woorden op
    // basis van hun middenhoogte, met een kleine tolerantie.
    const words=(data && data.words)||[];
    const usable=words.map(w=>{
      const b=w.bbox||{};
      return {text:norm(w.text),x0:Number(b.x0)||0,x1:Number(b.x1)||0,
        y0:Number(b.y0)||0,y1:Number(b.y1)||0};
    }).filter(w=>w.text && w.x1>=w.x0 && w.y1>=w.y0);
    if(!usable.length) return [];
    const heights=usable.map(w=>Math.max(1,w.y1-w.y0)).sort((a,b)=>a-b);
    const tolerance=Math.max(4,(heights[Math.floor(heights.length/2)]||10)*0.45);
    usable.sort((a,b)=>((a.y0+a.y1)/2)-((b.y0+b.y1)/2)||a.x0-b.x0);
    const grouped=[];
    for(const w of usable){
      const cy=(w.y0+w.y1)/2;
      let line=grouped.find(l=>Math.abs(l.cy-cy)<=tolerance);
      if(!line){line={cy,words:[]};grouped.push(line);}
      line.words.push(w);
      line.cy=line.words.reduce((sum,v)=>sum+(v.y0+v.y1)/2,0)/line.words.length;
    }
    return grouped.map(l=>{
      l.words.sort((a,b)=>a.x0-b.x0);
      return {text:norm(l.words.map(w=>w.text).join(' ')),
        x0:Math.min(...l.words.map(w=>w.x0)),x1:Math.max(...l.words.map(w=>w.x1)),
        y0:Math.min(...l.words.map(w=>w.y0)),y1:Math.max(...l.words.map(w=>w.y1))};
    }).sort((a,b)=>a.y0-b.y0);
  }

  function extractOrcaAddress(data){
    const lines=getLines(data);
    if(!lines.length) return null;
    // ORCA: het adres staat bij het grote ronde NODE-symbool. In de ORCA-
    // tekeningen staat dit adresblok in het bovenste deel van de blokschemaweergave.
    // Zoek daarom onafhankelijk van NODE-tekst naar een straat + huisnummer
    // in dat gebied; de OCR-regels worden niet aan een vaste x-positie gekoppeld.
    const pageHeight=Math.max(...lines.map(l=>l.y1),1);
    const streetSuffix=/(?:straat|laan|weg|plein|dreef|singel|kade|gracht|pad|dijk|hof|park|baan|steeg|ring|markt|plantsoen|weide|erf|veld|dam|vaart|land|stede|schelp|berg|broek|burg|dorp|horst|zoom|wal|waard|wetering|zijde|molen|reiger|hoep|kruislaan|avercampweg|hoofdenburg)$/i;
    const technical=/\b(?:NODE|WCA\d*|RF\s*DS|US\s*\d|AC\d{3,}|DBC\s*\d{3,}|OT-VCX|OS-VCX|PI-SPL|C3-[A-Z]+)\b/i;
    const candidates=[];
    for(const line of lines){
      const cy=(line.y0+line.y1)/2;
      if(cy>pageHeight*0.58) continue;
      const t=norm(line.text).replace(/[|]/g,'I');
      if(!t || technical.test(t)) continue;
      const m=t.match(/^(.{3,85}?)\s*,?\s+(\d{1,5}\s*[A-Za-z]?(?:[-/]\s*\d+[A-Za-z]?)?)\s*$/);
      if(!m) continue;
      const street=norm(m[1]);
      if(/\d/.test(street) || street.split(/\s+/).length>8) continue;
      const suffix=streetSuffix.test(street);
      candidates.push({street,house:norm(m[2]),y:cy,suffix,x:(line.x0+line.x1)/2});
    }
    if(!candidates.length) return null;
    // Het adres dat bij het ronde symbool hoort staat doorgaans het hoogst
    // in het schema; straatnaamsuffixen krijgen voorrang boven technische labels.
    candidates.sort((a,b)=>(a.suffix===b.suffix?0:(a.suffix?-1:1)) || a.y-b.y);
    const a=candidates[0];
    const code=(String(fileNameText||'').match(/^([A-Z0-9]+)/i)||[])[1]||'';
    const places={EVD:'Everdingen',HGSN:'Hagestein'};
    const place=places[code.toUpperCase()]||'';
    return `${a.street} ${a.house}${place?', '+place:''}`;
  }

  function extractFromOcr(data, requireNode=true){
    const lines=getLines(data);
    if(!lines.length) return null;

    // NODE is een voorkeur/anker, geen voorwaarde voor adresherkenning.
    // Als NODE ontbreekt of OCR het label niet herkent, zoeken we op de hele pagina.
    // Vind NODE als zelfstandig label; projectcodes zoals ASD-WC0002-NODE001 tellen niet.
    const nodeCandidates=lines.filter(l=>{
      const t=l.text.replace(/[|]/g,'I').replace(/[.,:;]/g,' ').replace(/\s+/g,' ').trim();
      // Alleen een los NODE-label telt; NODE001 in een projectcode niet.
      return /^(?:N\s*O\s*D\s*E|NODE)$/i.test(t);
    });
    if(requireNode && !nodeCandidates.length) return null;
    // Bij meerdere losse NODE-labels gebruiken we de meest waarschijnlijke: boven een adresblok.
    const node=nodeCandidates.sort((a,b)=>a.y0-b.y0)[0]||null;
    const nodeCenterY=node ? (node.y0+node.y1)/2 : Infinity;

    // Zoek postcode op basis van geometrie: elke horizontale positie is toegestaan,
    // maar de postcode moet boven NODE staan. Er is bewust geen maximale afstand.
    const postcodeCandidates=[];
    for(const line of lines){
      if(node && line.y1>=nodeCenterY) continue;
      const m=line.text.toUpperCase().replace(/[|]/g,'I').match(/(?:^|[^0-9])(\d{4})\s*([A-Z]{2})(?=$|[^A-Z])/i);
      if(!m) continue;
      postcodeCandidates.push({line,distance:nodeCenterY-(line.y0+line.y1)/2,
        postcode:(m[1]+' '+m[2]).toUpperCase(),
        remainder:norm(line.text.slice((m.index||0)+m[0].length))});
    }
    if(!postcodeCandidates.length) return null;
    postcodeCandidates.sort((a,b)=>a.distance-b.distance);
    const pc=postcodeCandidates[0];

    // Zoek straat + huisnummer boven de postcode. Ook hier geen vaste x-positie
    // of vaste verticale afstand: alleen de relatieve volgorde ten opzichte van NODE.
    const technical=/\b(?:NODE|PS\d*|WCA\d*|AD\d*|TB\d*|Teleste|AC\d{3,}|(?:ODK|DIM|DVT|ASZO)-WC\d+)\b/i;
    const streetSuffix=/(?:straat|laan|weg|plein|dreef|singel|kade|gracht|pad|dijk|hof|park|baan|steeg|ring|markt|plantsoen|weide|erf|veld|dam|vaart|land|stede|schelp|berg|broek|burg|dorp|horst|zoom|wal|waard|wetering|zijde|molen|reiger|hoep|kruislaan|avercampweg|hoofdenburg)$/i;
    const pcCenterY=(pc.line.y0+pc.line.y1)/2;
    const streetCandidates=[];
    for(const line of lines){
      const centerY=(line.y0+line.y1)/2;
      if(centerY>=pcCenterY || (node && centerY>=nodeCenterY)) continue;
      const t=norm(line.text);
      if(!t || technical.test(t)) continue;
      const m=t.match(/^(.{3,85}?)\s*,?\s+(\d{1,5}\s*[A-Za-z]?(?:[-/]\s*\d+[A-Za-z]?)?)\s*$/);
      if(!m) continue;
      const name=norm(m[1]);
      if(!/[A-Za-zÀ-ÖØ-öø-ÿ]/.test(name) || /\d/.test(name) || name.length<4 || name.length>75) continue;
      const words=name.split(/\s+/);
      if(words.length>8 || /^(?:ODK|DIM|DVT|ASZO|WCA|NODE|Teleste)/i.test(name)) continue;
      streetCandidates.push({name,house:m[2],distance:pcCenterY-centerY,
        suffix:streetSuffix.test(name)});
    }
    if(!streetCandidates.length) return null;
    streetCandidates.sort((a,b)=>(a.suffix===b.suffix?0:(a.suffix?-1:1)) || a.distance-b.distance);
    const street=streetCandidates[0].name;
    const house=streetCandidates[0].house;

    let city=pc.remainder.replace(/^[,;:\s-]+/,'');
    city=city.replace(/\b(?:NODE|ODK-WC\d+|DIM-WC\d+|DVT-WC\d+|ASZO-WC\d+)\b.*$/i,'').trim();
    if(!city){
      // Gemeente mag naast de postcode of op een andere horizontale positie staan,
      // zolang de regel tussen postcode en NODE ligt.
      const cityLines=lines.filter(l=>{
        const cy=(l.y0+l.y1)/2;
        return cy>pcCenterY && (!node || cy<nodeCenterY) &&
          /^[A-ZÀ-ÖØ-Ý][A-Za-zÀ-ÖØ-öø-ÿ' -]{2,40}$/.test(l.text) && !technical.test(l.text);
      }).sort((a,b)=>a.y0-b.y0);
      if(cityLines.length) city=cityLines[0].text;
    }
    if(!city) return `${street} ${house}, ${pc.postcode}`;
    return `${street} ${house}, ${pc.postcode} ${city}`;
  }

  let address='';
  if(pdf && window.Tesseract){
    try{
      const page=await pdf.getPage(1);
      // Hoge resolutie zodat kleine adresblokken leesbaar blijven,
      // ongeacht waar NODE op het A4-blad staat.
      const scale=5;
      const viewport=page.getViewport({scale});
      const canvas=document.createElement('canvas');
      canvas.width=Math.ceil(viewport.width);
      canvas.height=Math.ceil(viewport.height);
      await page.render({
        canvasContext:canvas.getContext('2d'),
        viewport
      }).promise;

      // PSM 11 vindt losse tekstblokken op technische blokschema's het best.
      // PSM 6 is een tweede poging voor compact uitgelijnde adresregels.
      if(blokschemaType==="PNI"){
        // PNI: extra tekst rechts van de NODE (zoals TB-/PS-aanduidingen)
        // mag niet in de adresherkenning terechtkomen. Zoek eerst de NODE op
        // de volledige pagina en OCR daarna alleen het smalle adresgebied
        // direct erboven, gecentreerd op de NODE.
        const clean=t=>norm(t).replace(/[|]/g,'I').replace(/[–—]/g,'-').trim();
        const technical=/\b(?:PS\d{1,3}|TB\d{3,}|TXFMR|WCA\d*|AC\d{3,}|NODE\d*|Teleste|SFP|RFCI)\b/i;
        const fullData=await ocrCanvas(canvas,{
          tessedit_pageseg_mode:'11',
          preserve_interword_spaces:'1'
        });
        const fullLines=getLines(fullData).map(l=>({...l,text:clean(l.text)})).filter(l=>l.text);
        const nodes=fullLines.filter(l=>/^(?:N\s*O\s*D\s*E|NODE)[.,:;]?$/i.test(l.text))
          .sort((a,b)=>a.y0-b.y0);

        // Diagnostiek voor PNI: toon de gevonden NODE, de exacte uitsnede
        // en de ruwe OCR-regels per poging, zodat de zoekzone controleerbaar is.
        let debugPanel=document.getElementById('pniOcrDebugPanel');
        if(!debugPanel){
          debugPanel=document.createElement('div');
          debugPanel.id='pniOcrDebugPanel';
          debugPanel.style.cssText='margin:12px 0;padding:12px;border:2px solid #d97706;border-radius:8px;background:#fffaf0;color:#222;font:13px/1.45 Arial,sans-serif;max-width:100%;overflow-wrap:anywhere;';
          const anchor=document.getElementById('addressInfo');
          if(anchor && anchor.parentNode) anchor.parentNode.insertBefore(debugPanel,anchor.nextSibling);
        }
        debugPanel.innerHTML='<b>PNI OCR-diagnose</b><br>Volledige pagina OCR (PSM 11): '+
          (fullLines.length?fullLines.map(l=>escapeHtml(l.text)).join(' | '):'geen tekstregels')+
          '<br>Losse NODE gevonden: '+nodes.length;

        for(const node of nodes){
          const nx=(node.x0+node.x1)/2;
          const ny=(node.y0+node.y1)/2;
          // PNI-adresgebied: 4% links en 6% rechts vanaf het NODE-midden.
          // Iets meer ruimte rechts om het volledige adresblok mee te nemen.
          const left=Math.max(0,Math.floor(nx-canvas.width*0.04));
          const right=Math.min(canvas.width,Math.ceil(nx+canvas.width*0.06));
          const top=Math.max(0,Math.floor(ny-canvas.height*0.115));
          // Lees uitsluitend boven de NODE: stop net vóór de bovenrand
          // van de herkende NODE-tekstregel, niet op/onder het midden.
          const bottom=Math.max(top+1,Math.min(canvas.height,
            Math.floor(node.y0-canvas.height*0.003)));
          debugPanel.innerHTML += `<hr><b>NODE-coördinaten</b>: x=${nx.toFixed(1)} (${(nx/canvas.width*100).toFixed(2)}%), y=${ny.toFixed(1)} (${(ny/canvas.height*100).toFixed(2)}%)`+
            `<br><b>Zoekgrenzen (4% links / 6% rechts)</b>: links=${left} (${(left/canvas.width*100).toFixed(2)}%), rechts=${right} (${(right/canvas.width*100).toFixed(2)}%)`+
            `<br><b>Verticale uitsnede</b>: boven=${top}, onder=${bottom} (stopt vóór NODE; onderliggende tekst uitgesloten)`;
          // Toon de 4%-grenzen zichtbaar over de volledige PDF-pagina.
          const guide=document.createElement('canvas');
          guide.width=canvas.width; guide.height=canvas.height;
          const guideCtx=guide.getContext('2d');
          guideCtx.drawImage(canvas,0,0);
          guideCtx.save();
          guideCtx.lineWidth=Math.max(3,canvas.width*0.0015);
          guideCtx.setLineDash([Math.max(10,canvas.width*0.006),Math.max(7,canvas.width*0.004)]);
          guideCtx.strokeStyle='#e00000';
          guideCtx.beginPath(); guideCtx.moveTo(left,0); guideCtx.lineTo(left,canvas.height);
          guideCtx.moveTo(right,0); guideCtx.lineTo(right,canvas.height); guideCtx.stroke();
          guideCtx.setLineDash([]);
          guideCtx.strokeStyle='#1769e0'; guideCtx.lineWidth=Math.max(2,canvas.width*0.001);
          guideCtx.beginPath(); guideCtx.moveTo(nx,Math.max(0,ny-canvas.height*.14)); guideCtx.lineTo(nx,Math.min(canvas.height,ny+canvas.height*.03)); guideCtx.stroke();
          // Groene horizontale lijn is de ondergrens: niets eronder gaat naar OCR.
          guideCtx.strokeStyle='#159447'; guideCtx.lineWidth=Math.max(3,canvas.width*0.0015);
          guideCtx.beginPath(); guideCtx.moveTo(left,bottom); guideCtx.lineTo(right,bottom); guideCtx.stroke();
          guideCtx.restore();
          debugPanel.innerHTML += `<br><b>PDF met zoekgrenzen</b> (rood = 4% links / 6% rechts, blauw = midden NODE, groen = ondergrens OCR)`+
            `<img alt="PDF met zichtbare 4 procent zoekgrenzen" style="display:block;width:100%;height:auto;max-height:680px;object-fit:contain;border:1px solid #999;background:#fff;margin:8px 0" src="${guide.toDataURL('image/png')}">`;
          if(right-left<80 || bottom-top<40){debugPanel.innerHTML+='<br>Uitsnede overgeslagen: te klein.';continue;}
          const crop=document.createElement('canvas');
          crop.width=right-left;
          crop.height=bottom-top;
          crop.getContext('2d').drawImage(canvas,left,top,crop.width,crop.height,0,0,crop.width,crop.height);

          // Verbeter alleen de PNI-adresuitsnede; de 4%-grenzen blijven gelijk.
          // Verdubbel de pixels en probeer daarnaast een contrastrijke zwart-witversie.
          const enlarged=document.createElement('canvas');
          enlarged.width=crop.width*2;
          enlarged.height=crop.height*2;
          const enlargedCtx=enlarged.getContext('2d',{willReadFrequently:true});
          enlargedCtx.imageSmoothingEnabled=true;
          enlargedCtx.imageSmoothingQuality='high';
          enlargedCtx.drawImage(crop,0,0,enlarged.width,enlarged.height);
          debugPanel.innerHTML += `<br><b>Daadwerkelijk OCR-beeld (4% links / 6% rechts, alleen boven NODE, vergroot)</b><br><img alt="PNI OCR-uitsnede" style="display:block;max-width:100%;max-height:420px;image-rendering:auto;border:1px solid #999;margin:6px 0" src="${enlarged.toDataURL('image/png')}">`;
          const contrast=document.createElement('canvas');
          contrast.width=enlarged.width;
          contrast.height=enlarged.height;
          const contrastCtx=contrast.getContext('2d',{willReadFrequently:true});
          contrastCtx.drawImage(enlarged,0,0);
          const pixels=contrastCtx.getImageData(0,0,contrast.width,contrast.height);
          // Lokale min/max-percentielen vermijden dat een paar donkere symbolen
          // de kleine grijze adresletters laten verdwijnen.
          const hist=new Uint32Array(256);
          for(let i=0;i<pixels.data.length;i+=4){
            const v=Math.round(.299*pixels.data[i]+.587*pixels.data[i+1]+.114*pixels.data[i+2]);
            hist[v]++;
          }
          const total=contrast.width*contrast.height;
          let lo=0,hi=255,acc=0;
          while(lo<255 && (acc+=hist[lo])<total*.01)lo++;
          acc=0;
          while(hi>0 && (acc+=hist[hi])<total*.01)hi--;
          if(hi<=lo+20){lo=0;hi=255;}
          for(let i=0;i<pixels.data.length;i+=4){
            const gray=.299*pixels.data[i]+.587*pixels.data[i+1]+.114*pixels.data[i+2];
            const v=Math.max(0,Math.min(255,Math.round((gray-lo)*255/(hi-lo))));
            pixels.data[i]=pixels.data[i+1]=pixels.data[i+2]=v;
          }
          contrastCtx.putImageData(pixels,0,0);
          // Eerst de gewone, vergrote versie; alleen bij mislukken de
          // contrastversie. Dat voorkomt extra OCR-tijd bij goede PDF's.
          for(const candidateCanvas of [enlarged,contrast]){
          for(const psm of ['6','11','4']){
            if(addressRunToken!==window.__addressRunToken || manualAddressConfirmed) return;
            const data=await ocrCanvas(candidateCanvas,{
              tessedit_pageseg_mode:psm,
              preserve_interword_spaces:'1',
              user_defined_dpi:'300'
            });
            const lines=getLines(data).map(l=>({...l,text:clean(l.text)})).filter(l=>l.text);
            debugPanel.innerHTML += `<br><b>OCR ${candidateCanvas===enlarged?'vergroot':'contrast'} / PSM ${psm}</b>: `+
              (lines.length?lines.map(l=>escapeHtml(l.text)).join(' | '):'geen tekstregels');
            if(!lines.length) continue;

            const pcCandidates=[];
            for(const l of lines){
              const m=l.text.toUpperCase().replace(/[|]/g,'I')
                .match(/(?:^|[^A-Z0-9])(\d{4})\s*([A-Z]{2})(?:\s+(.+))?$/i);
              if(m) pcCandidates.push({line:l,postcode:(m[1]+' '+m[2]).toUpperCase(),
                remainder:clean(m[3]||''),cy:(l.y0+l.y1)/2});
            }
            if(!pcCandidates.length){debugPanel.innerHTML+=' <i>→ geen postcode herkend</i>';continue;}
            debugPanel.innerHTML += ' <i>→ postcode: '+escapeHtml(pcCandidates.map(v=>v.postcode).join(', '))+'</i>';
            // Postcode staat normaal onder de straat en boven de NODE.
            pcCandidates.sort((a,b)=>a.cy-b.cy);
            const pc=pcCandidates[0];
            const streetSuffix=/(?:straat|laan|weg|plein|dreef|singel|kade|gracht|pad|dijk|hof|park|baan|steeg|ring|markt|plantsoen|weide|erf|veld|dam|vaart|land|stede|berg|broek|burg|dorp|horst|zoom|wal|waard|wetering|zijde|molen|reiger|hoep)$/i;
            const streets=lines.filter(l=>{
              const cy=(l.y0+l.y1)/2;
              if(cy>=pc.cy || technical.test(l.text)) return false;
              if(/\b[A-Z]{2,10}\s*-\s*WC\s*\d{3,6}\b/i.test(l.text)) return false;
              return /^([A-Za-zÀ-ÖØ-öø-ÿ' .-]{3,85}?)\s+(\d{1,5}\s*[A-Za-z]?(?:[-/]\s*\d+[A-Za-z]?)?)$/i.test(l.text);
            }).map(l=>{
              const m=l.text.match(/^([A-Za-zÀ-ÖØ-öø-ÿ' .-]{3,85}?)\s+(\d{1,5}\s*[A-Za-z]?(?:[-/]\s*\d+[A-Za-z]?)?)$/i);
              return {name:clean(m[1]),house:clean(m[2]),cy:(l.y0+l.y1)/2,
                suffix:streetSuffix.test(clean(m[1]))};
            }).filter(v=>v.name.length>=4 && !/\d/.test(v.name));
            if(!streets.length){debugPanel.innerHTML+=' <i>→ geen straat + huisnummer herkend</i>';continue;}
            debugPanel.innerHTML += ' <i>→ straatkandidaten: '+escapeHtml(streets.map(v=>v.name+' '+v.house).join(' / '))+'</i>';
            streets.sort((a,b)=>(Number(b.suffix)-Number(a.suffix)) || ((pc.cy-a.cy)-(pc.cy-b.cy)));
            const street=streets[0];
            let city=pc.remainder.replace(/^[,;:\s-]+/,'').trim();
            if(city) city=city.replace(/\b(?:NODE|PS\d+|TB\d+|TXFMR).*$/i,'').trim();
            if(!city){
              const cityLines=lines.filter(l=>{
                const cy=(l.y0+l.y1)/2;
                return cy>pc.cy && !technical.test(l.text) &&
                  !/\d{4}\s*[A-Z]{2}/i.test(l.text) &&
                  /^[A-ZÀ-ÖØ-Ý][A-Za-zÀ-ÖØ-öø-ÿ' -]{2,45}$/.test(l.text);
              }).sort((a,b)=>a.y0-b.y0);
              if(cityLines.length) city=cityLines[0].text;
            }
            address=`${street.name} ${street.house}, ${pc.postcode}${city?', '+city:''}`;
            debugPanel.innerHTML += '<br><b>Gekozen adres:</b> '+escapeHtml(address);
            break;
          }
          if(address) break;
          }
          if(address) break;
        }
      }else{
        // ORCA blijft exact de bestaande herkenning gebruiken.
        for(const psm of ['11','6','12','4']){
          const data=await ocrCanvas(canvas,{
            tessedit_pageseg_mode:psm,
            preserve_interword_spaces:'1'
          });
          address=extractOrcaAddress(data);
          if(address) break;
        }
      }
    }catch(err){
      console.warn('Adresherkenning via NODE is mislukt:',err);
    }
  }

  // Een handmatige invoer of een nieuw gestart PDF-proces heeft altijd voorrang.
  if(addressRunToken!==window.__addressRunToken || manualAddressConfirmed || manualAddressPopupOpen) return;

  if(address){
    addressText=address;
    addressInfo.innerHTML=`<span class="ok">Adres automatisch uit PDF gevonden:</span><br>${escapeHtml(addressText)}`;
    return;
  }

  addressText='';
  addressInfo.textContent='Adres niet automatisch gevonden. Vul het adres handmatig in.';
  openManualAddressPopup();
}

function openManualAddressPopup(){
  const modal=document.getElementById('manualAddressModal');
  const input=document.getElementById('manualAddressInput');
  const save=document.getElementById('manualAddressSave');
  const skip=document.getElementById('manualAddressSkip');
  if(!modal || !input || !save || !skip) return;
  if(manualAddressConfirmed || manualAddressPopupOpen) return;

  manualAddressPopupOpen=true;
  modal.style.display='flex';
  input.value=addressText||'';
  input.focus();

  save.onclick=async()=>{
    const value=String(input.value||'').trim();
    if(!value){
      input.focus();
      input.style.borderColor='#d33';
      return;
    }
    addressText=value;
    manualAddressConfirmed=true;
    manualAddressPopupOpen=false;
    addressInfo.innerHTML='<span class="ok">Adres handmatig ingevoerd:</span><br>'+escapeHtml(addressText);
    modal.style.display='none';
    input.style.borderColor='#aaa';
    if(typeof renderPage==='function' && pdf){
      try{ await renderPage(); }catch(err){ console.warn('Adresweergave bijwerken mislukt:',err); }
    }
  };

  skip.onclick=()=>{
    manualAddressPopupOpen=false;
    modal.style.display='none';
    addressInfo.textContent='Geen adres ingevuld. Je kunt verder werken; het adres kan later handmatig worden toegevoegd.';
  };

  input.onkeydown=event=>{
    if(event.key==='Enter'){
      event.preventDefault();
      save.click();
    }
  };
}
function escapeHtml(text){
  return text.replace(/[&<>"']/g,m=>({
    "&":"&amp;","<":"&lt;",">":"&gt;",
    '"':"&quot;","'":"&#039;"
  }[m]));
}

function findWcaItems(){
  wcaItems=[];

  // Eerst normale afzonderlijke tekstitems.
  const regex=/W\s*C\s*A\s*[-._]?\s*(\d{1,4})\s*[-._]?\s*([A-Z])/i;

  for(const item of textItems){
    const raw=String(item.text||"").replace(/\s+/g," ");
    const match=raw.match(regex);
    if(!match)continue;

    const nummer=match[1].padStart(3,"0");
    const letter=match[2].toUpperCase();
    const wca=`WCA${nummer}${letter}`;

    if(!wcaItems.some(x=>x.page===item.page&&x.wca===wca)){
      wcaItems.push({
        ...item,
        wca,
        daa:`DAA${nummer}`
      });
    }
  }

  // Tweede poging: PDF's waarin WCA / nummer / letter als losse
  // tekstobjecten zijn opgeslagen.
  for(let i=0;i<textItems.length;i++){
    const a=textItems[i];
    const ax=a.transform ? a.transform[4] : a.x;
    const ay=a.transform ? a.transform[5] : a.y;

    for(let j=0;j<textItems.length;j++){
      if(i===j)continue;
      const b=textItems[j];
      if(a.page!==b.page)continue;

      const bx=b.transform ? b.transform[4] : b.x;
      const by=b.transform ? b.transform[5] : b.y;

      if(Math.abs(ax-bx)>80 || Math.abs(ay-by)>35)continue;

      const combined=`${a.text||""}${b.text||""}`.replace(/\s+/g,"");
      const m=combined.match(/W\s*C\s*A[-._]?(\d{1,4})[-._]?([A-Z])/i);
      if(!m)continue;

      const nummer=m[1].padStart(3,"0");
      const letter=m[2].toUpperCase();
      const wca=`WCA${nummer}${letter}`;

      if(!wcaItems.some(x=>x.page===a.page&&x.wca===wca)){
        wcaItems.push({
          ...a,
          text:combined,
          wca,
          daa:`DAA${nummer}`
        });
      }
    }
  }

  // Derde poging: OCR kan "WCA", "001" en "A" als drie losse
  // woorden teruggeven. Combineer daarom ook korte groepjes die vlak
  // bij elkaar staan.
  for(let i=0;i<textItems.length;i++){
    const a=textItems[i];
    for(let j=i+1;j<textItems.length;j++){
      const b=textItems[j];
      if(a.page!==b.page)continue;

      const ax=a.transform ? a.transform[4] : a.x;
      const ay=a.transform ? a.transform[5] : a.y;
      const bx=b.transform ? b.transform[4] : b.x;
      const by=b.transform ? b.transform[5] : b.y;

      if(Math.hypot(ax-bx,ay-by)>120)continue;

      for(let k=j+1;k<textItems.length;k++){
        const c=textItems[k];
        if(c.page!==a.page)continue;

        const cx=c.transform ? c.transform[4] : c.x;
        const cy=c.transform ? c.transform[5] : c.y;

        if(Math.hypot(ax-cx,ay-cy)>120)continue;

        const parts=[a,b,c].sort((u,v)=>{
          const ux=u.transform ? u.transform[4] : u.x;
          const vx=v.transform ? v.transform[4] : v.x;
          return ux-vx;
        });

        const combined=parts.map(x=>String(x.text||""))
          .join("")
          .replace(/\s+/g,"");

        const m=combined.match(/W\s*C\s*A[-._]?(\d{1,4})[._-]?([A-Z])/i);
        if(!m)continue;

        const nummer=m[1].padStart(3,"0");
        const letter=m[2].toUpperCase();
        const wca=`WCA${nummer}${letter}`;

        if(!wcaItems.some(x=>x.page===a.page&&x.wca===wca)){
          wcaItems.push({
            ...a,
            text:parts.map(x=>x.text).join(" "),
            wca,
            daa:`DAA${nummer}`
          });
        }
        break;
      }
    }
  }

  // Sorteer op pagina en daarna op positie, zodat alle WCA's verwerkt worden.
  wcaItems.sort((a,b)=>{
    if(a.page!==b.page)return a.page-b.page;
    const ay=a.transform ? a.transform[5] : a.y;
    const by=b.transform ? b.transform[5] : b.y;
    return ay-by;
  });

  if(!wcaItems.length){
    found.innerHTML="Geen WCA-aanduidingen gevonden.";
  }else{
    found.innerHTML=
      "<b>Gevonden:</b> "+
      wcaItems.map(x=>
        `<span class="wca">${x.wca} → ${x.daa}</span>`
      ).join("");
  }
}


/* ---------------- AUTOMATISCH WCA -> DAA ---------------- */

function viewportPointForText(item,viewport){
  if(item && item.transform){
    return viewport.convertToViewportPoint(item.transform[4],item.transform[5]);
  }
  return viewport.convertToViewportPoint(item.x,item.y);
}

function findFlowerOnRenderedPage(canvas,item,viewport){
  /*
    BELANGRIJK:
    De WCA-tekst staat onderaan bij het bloemetje. We zoeken daarom
    NIET meer naar de sterkste verticale lijn in een groot gebied.
    Dat pakte soms een andere oude verticale lijn.

    Nieuwe werkwijze:
    1. Neem het midden van de WCA-tekst als referentie-X.
    2. Zoek direct BOVEN de WCA naar het bloemetje/aansluitpunt.
    3. Zoek daarna uitsluitend de verticale lijn die vanaf dat punt
       recht omhoog loopt.
    Zo wordt een andere verticale lijn verderop niet meer gekozen.
  */

  const pt=viewportPointForText(item,viewport);
  const x0=Math.round(pt[0]);
  const y0=Math.round(pt[1]);
  const w=canvas.width;
  const h=canvas.height;
  const image=canvas.getContext("2d").getImageData(0,0,w,h).data;

  function isDark(x,y){
    if(x<0||x>=w||y<0||y>=h)return false;
    const i=(y*w+x)*4;
    return image[i]<160 && image[i+1]<160 && image[i+2]<160;
  }

  function darkCount(x1,x2,y1,y2){
    x1=Math.max(0,Math.floor(x1));
    x2=Math.min(w-1,Math.ceil(x2));
    y1=Math.max(0,Math.floor(y1));
    y2=Math.min(h-1,Math.ceil(y2));
    let n=0;
    for(let yy=y1;yy<=y2;yy++){
      for(let xx=x1;xx<=x2;xx++){
        if(isDark(xx,yy))n++;
      }
    }
    return n;
  }

  // Zoek het verticale aansluitpunt direct boven de WCA-tekst.
  // De WCA-tekst zelf wordt niet meegenomen in deze zone.
  let bestX=null;
  let bestScore=-1;
  let bestTop=null;

  const minX=Math.max(0,x0-35);
  const maxX=Math.min(w-1,x0+35);
  const searchBottom=Math.max(0,y0-8);
  const searchTop=Math.max(0,y0-320);

  for(let x=minX;x<=maxX;x++){
    let score=0;
    let longest=0;
    let run=0;
    let firstY=null;

    for(let y=searchBottom;y>=searchTop;y--){
      // 3 pixels breed zodat dunne PDF-lijnen ook worden gevonden.
      const hit=isDark(x,y)||isDark(x-1,y)||isDark(x+1,y);
      if(hit){
        score++;
        run++;
        if(firstY===null)firstY=y;
        if(run>longest)longest=run;
      }else{
        // Kleine onderbrekingen in een gestippelde/dash-lijn toestaan.
        run=0;
      }
    }

    // Een echte verticale aansluitlijn heeft veel meer verticale
    // continuïteit dan losse tekst rondom de WCA.
    const weighted=score + longest*3;
    if(longest>=8 && weighted>bestScore){
      bestScore=weighted;
      bestX=x;
      bestTop=firstY;
    }
  }

  if(bestX===null){
    // Fallback: neem alleen de X van de WCA en zoek het dichtstbijzijnde
    // donkere aansluitpunt erboven. Geen willekeurige verre lijn.
    bestX=x0;
    bestTop=searchBottom;
  }

  // Zoek het bloemetje/aansluitpunt rond de onderkant van deze gevonden
  // verticale lijn. We zoeken uitsluitend rond bestX.
  let flowerY=null;
  let peakY=null;
  let peak=-1;

  const flowerTop=Math.max(0,y0-55);
  const flowerBottom=Math.min(h-1,y0+18);

  for(let y=flowerTop;y<=flowerBottom;y++){
    const count=darkCount(bestX-16,bestX+16,y,y);
    if(count>peak){
      peak=count;
      peakY=y;
    }

    // Bloemetje/triangle bestaat uit meerdere donkere pixels over
    // opeenvolgende rijen. Een enkele verticale lijn haalt dit niet.
    if(count>=8){
      let confirm=0;
      for(let yy=y;yy<=Math.min(flowerBottom,y+12);yy++){
        if(darkCount(bestX-16,bestX+16,yy,yy)>=5)confirm++;
      }
      if(confirm>=3){
        flowerY=y;
        break;
      }
    }
  }

  if(flowerY===null){
    // Als het bloemetje heel dun is, gebruik het dichtstbijzijnde
    // aansluitpunt vlak boven de WCA-tekst.
    flowerY=Math.max(flowerTop,Math.min(flowerBottom,peakY||searchBottom));
  }

  // Bepaal het midden van het bloemetje, maar blijf binnen de lokale zone.
  let sumY=0,sumW=0;
  for(let y=Math.max(0,flowerY-10);y<=Math.min(h-1,flowerY+18);y++){
    const count=darkCount(bestX-16,bestX+16,y,y);
    if(count>0){
      sumY+=y*count;
      sumW+=count;
    }
  }

  const fy=sumW?sumY/sumW:flowerY;

  return {
    x:bestX,
    y:fy,
    // Eindpunt net onder het bloemetje zodat het bloemetje zichtbaar blijft.
    connectionY:fy+18,
    page:item.page,
    wca:item.wca
  };
}


function findNumberedCircleNearWca(item,viewport){
  // In deze ORCA-variant staat WCA001A/B/C/D boven een rondje
  // met een nummer (1, 2, 3, ...). De DAA moet op dat rondje
  // aansluiten, niet op de positie van de WCA-tekst.
  const wcaPt=viewportPointForText(item,viewport);
  const candidates=[];

  for(const t of textItems){
    if(t.page!==item.page || !t.ocr) continue;
    const raw=String(t.text||"").trim().replace(/[Oo]/g,"0");
    if(!/^[1-9]$/.test(raw)) continue;

    const p=viewportPointForText(t,viewport);
    const dx=Math.abs(p[0]-wcaPt[0]);
    const dy=p[1]-wcaPt[1];

    // Het nummer hoort onder de WCA-tekst en dicht bij dezelfde verticale lijn.
    if(dx>32 || dy<8 || dy>130) continue;

    const confidence=t.confidence||0;
    const score=dx*2 + Math.abs(dy-48) + (confidence>=60 ? 0 : 20);
    candidates.push({t,p,score});
  }

  candidates.sort((a,b)=>a.score-b.score);
  if(candidates.length){
    const best=candidates[0];
    return {
      x:best.p[0],
      y:best.p[1],
      connectionY:best.p[1],
      number:best.t.text.trim(),
      page:item.page,
      wca:item.wca,
      fromNumberedCircle:true
    };
  }

  return null;
}


function findPniTopNumberedCircles(canvas,viewport){
  /*
   * PNI: eerst de fysieke genummerde rondjes zoeken, daarna pas OCR.
   * De twee geteste PNI-layouts hebben de rondjes op verschillende Y-posities
   * (bijv. ASD hoger dan DIM). Daarom geen vaste Y-band gebruiken.
   * We zoeken in het bovenste deel en kiezen de horizontale rij met de meeste
   * duidelijke ronde aansluitpunten. De WCA-tekst bij woningen komt hierdoor
   * nooit rechtstreeks in de selectie terecht.
   */
  const ctx=canvas.getContext('2d');
  const img=ctx.getImageData(0,0,canvas.width,canvas.height);
  const d=img.data,w=canvas.width,h=canvas.height;
  const dark=(x,y)=>{
    if(x<0||y<0||x>=w||y>=h)return 0;
    const i=(y*w+x)*4;
    return d[i]<190 && d[i+1]<190 && d[i+2]<190 ? 1:0;
  };
  const ringScore=(cx,cy,r)=>{
    let hit=0,total=0;
    for(let a=0;a<360;a+=15){
      const rad=a*Math.PI/180;
      let ok=0;
      for(let dd=-1;dd<=1;dd++){
        const x=Math.round(cx+Math.cos(rad)*r-Math.sin(rad)*dd);
        const y=Math.round(cy+Math.sin(rad)*r+Math.cos(rad)*dd);
        if(dark(x,y)){ok=1;break;}
      }
      hit+=ok; total++;
    }
    return total?hit/total:0;
  };
  const innerInk=(cx,cy,r)=>{
    let hit=0,total=0;
    const rr=Math.max(2,Math.floor(r*.48));
    for(let yy=-rr;yy<=rr;yy++)for(let xx=-rr;xx<=rr;xx++){
      if(xx*xx+yy*yy>rr*rr)continue;
      total++; hit+=dark(cx+xx,cy+yy);
    }
    return total?hit/total:1;
  };

  const yMin=Math.floor(h*0.16), yMax=Math.floor(h*0.55);
  const raw=[];
  const radii=[7,8,9,10,11,12];

  // Stap 4 px voorkomt dat deze scan onnodig zwaar wordt, maar is klein
  // genoeg om de kleine PNI-rondjes niet over te slaan.
  for(let cy=yMin;cy<=yMax;cy+=4){
    for(let cx=10;cx<w-10;cx+=4){
      let best=0,bestR=0;
      for(const r of radii){
        const sc=ringScore(cx,cy,r);
        if(sc>best){best=sc;bestR=r;}
      }
      if(best<0.78)continue;
      const inside=innerInk(cx,cy,bestR);
      // Een echt genummerd rondje heeft een ring en beperkte inkt binnenin.
      if(inside>0.42)continue;
      raw.push({x:cx,y:cy,r:bestR,ring:best,inside,
        score:best+(1-inside)*0.22});
    }
  }

  raw.sort((a,b)=>b.score-a.score);
  const clusters=[];
  for(const c of raw){
    const near=clusters.find(q=>Math.hypot(q.x-c.x,q.y-c.y)<20);
    if(near){
      if(c.score>near.score)Object.assign(near,c);
    }else{
      clusters.push({...c});
    }
  }
  if(!clusters.length)return [];

  // Groepeer rondjes per horizontale rij. De bovenste WCA-rij heeft in deze
  // schema's meerdere aansluitpunten; lagere technische rondjes vallen zo af.
  const rows=[];
  for(const c of clusters.sort((a,b)=>a.y-b.y)){
    let row=rows.find(r=>Math.abs(r.y-c.y)<=22);
    if(!row){row={y:c.y,items:[]};rows.push(row);}
    row.items.push(c);
    row.y=row.items.reduce((sum,q)=>sum+q.y,0)/row.items.length;
  }

  rows.sort((a,b)=>b.items.length-a.items.length || a.y-b.y);
  const chosenRow=rows[0];
  if(!chosenRow || chosenRow.items.length<2)return [];

  chosenRow.items.sort((a,b)=>a.x-b.x);
  const chosen=[];
  for(const c of chosenRow.items){
    if(chosen.some(q=>Math.abs(q.x-c.x)<24))continue;
    chosen.push(c);
    if(chosen.length>=12)break;
  }
  return chosen.map((c,i)=>({...c,index:i+1,page:currentPage}));
}


async function ocrPniWcaNearCircle(pageNumber, viewport, circle){
  /*
   * PNI WCA-OCR v27.76
   *
   * In the actual PNI PDF the WCA text is VERTICAL beside/above the
   * numbered circle.  The previous version treated it as horizontal.
   * We therefore OCR BOTH orientations and accept a WCA code only when
   * it is found in the small crop belonging to this physical circle.
   */
  try{
    const page=await pdf.getPage(pageNumber);
    const scaleLocal=6;
    const vp=page.getViewport({scale:scaleLocal});
    const src=document.createElement('canvas');
    src.width=Math.ceil(vp.width); src.height=Math.ceil(vp.height);
    await page.render({canvasContext:src.getContext('2d'),viewport:vp}).promise;

    const factor=scaleLocal/scale;
    const cx=circle.x*factor;
    const cy=circle.y*factor;

    // De WCA-code staat vlak boven/naast het rondje en loopt verticaal.
    // We nemen een ruime lokale rechthoek, maar nooit de hele pagina.
    const halfW=Math.round(95*factor);
    const top=Math.max(0,Math.round(cy-250*factor));
    const bottom=Math.min(src.height,Math.round(cy-5*factor));
    const sx=Math.max(0,Math.round(cx-halfW));
    const sy=top;
    const sw=Math.min(src.width-sx,Math.round(halfW*2));
    const sh=Math.max(40,bottom-top);

    const crop=document.createElement('canvas');
    crop.width=sw*2; crop.height=sh*2;
    const cctx=crop.getContext('2d');
    cctx.imageSmoothingEnabled=false;
    cctx.drawImage(src,sx,sy,sw,sh,0,0,crop.width,crop.height);

    function makeVariant(source,threshold){
      const c=document.createElement('canvas');
      c.width=source.width*2; c.height=source.height*2;
      const cc=c.getContext('2d');
      cc.imageSmoothingEnabled=false;
      cc.drawImage(source,0,0,c.width,c.height);
      if(threshold){
        const id=cc.getImageData(0,0,c.width,c.height);
        for(let i=0;i<id.data.length;i+=4){
          const g=id.data[i]*.299+id.data[i+1]*.587+id.data[i+2]*.114;
          const v=g<threshold?0:255;
          id.data[i]=id.data[i+1]=id.data[i+2]=v;
        }
        cc.putImageData(id,0,0);
      }
      return c;
    }

    // Zowel horizontaal als beide verticale richtingen proberen.
    const rotations=[0,-90,90];
    const variants=[];
    for(const rot of rotations){
      const r=document.createElement('canvas');
      const rr=r.getContext('2d');
      if(rot===0){
        r.width=crop.width; r.height=crop.height;
        rr.drawImage(crop,0,0);
      }else{
        r.width=crop.height; r.height=crop.width;
        // Correcte rotatie rond het midden/nieuwe canvas.
        // Bij -90° moet de afbeelding eerst over de volledige nieuwe hoogte
        // worden verschoven; anders valt de tekst buiten het canvas.
        if(rot===90){
          rr.translate(r.width,0);
          rr.rotate(Math.PI/2);
        }else{
          rr.translate(0,r.height);
          rr.rotate(-Math.PI/2);
        }
        rr.drawImage(crop,0,0);
      }
      variants.push(makeVariant(r,null));
      variants.push(makeVariant(r,185));
      variants.push(makeVariant(r,210));
    }

    function normaliseer(text){
      let t=String(text||'').toUpperCase()
        .replace(/[|!]/g,'I')
        .replace(/[“”'`]/g,'')
        .replace(/\s+/g,' ')
        .trim();
      t=t.replace(/W\s*C\s*A/g,'WCA');
      // OCR leest 0/O en 1/I/L soms door elkaar.
      t=t.replace(/WCA\s*[-._:]?\s*([0-9OIL]{1,4})\s*([A-D])/g,
        (all,num,letter)=>`WCA${num.replace(/O/g,'0').replace(/[IL]/g,'1')}${letter}`);
      return t;
    }

    const textsToTry=[];
    for(const c of variants){
      for(const psm of [6,7,11,12]){
        const data=await ocrCanvas(c,{
          tessedit_pageseg_mode:String(psm),
          preserve_interword_spaces:'1'
        });
        if(data?.text)textsToTry.push(data.text);
        for(const w of (data?.words||[])) if(w.text)textsToTry.push(w.text);
      }
    }

    for(const raw of textsToTry){
      const t=normaliseer(raw);
      const compact=t.replace(/[^A-Z0-9]/g,'');
      const candidates=[t,compact];
      for(const cand of candidates){
        let m=cand.match(/WCA([0-9OIL]{1,4})([A-D])/i);
        if(!m) m=cand.match(/WCA\s*([0-9OIL]{1,4})\s*([A-D])/i);
        if(!m)continue;
        const n=Number(m[1].replace(/O/g,'0').replace(/[IL]/g,'1'));
        const letter=m[2].toUpperCase();
        if(n>=1&&n<=8&&['A','B','C','D'].includes(letter)){
          return `WCA${String(n).padStart(3,'0')}${letter}`;
        }
      }
    }
  }catch(e){
    console.warn('PNI lokale WCA OCR mislukt',e);
  }
  return null;
}

function findPniLocalNumberedCircle(canvas, viewport, item){
  // PNI: zoek het kleine genummerde rondje DIRECT onder de WCA-tekst.
  // We zoeken lokaal rond de WCA, niet over de hele pagina. Daardoor kan
  // een ander rondje of de herhaalde WCA-tekst bij de woningen niet winnen.
  const p=viewportPointForText(item,viewport);
  const ctx=canvas.getContext('2d');
  const img=ctx.getImageData(0,0,canvas.width,canvas.height);
  const d=img.data, w=canvas.width, h=canvas.height;
  const dark=(x,y)=>{
    if(x<0||y<0||x>=w||y>=h)return 0;
    const i=(y*w+x)*4;
    return d[i]<215 && d[i+1]<215 && d[i+2]<215 ? 1:0;
  };
  const ringScore=(cx,cy,r)=>{
    let hit=0,total=0;
    for(let a=0;a<360;a+=12){
      const rad=a*Math.PI/180;
      let ok=0;
      for(let dd=-1;dd<=1;dd++){
        const x=Math.round(cx+Math.cos(rad)*r-Math.sin(rad)*dd);
        const y=Math.round(cy+Math.sin(rad)*r+Math.cos(rad)*dd);
        if(dark(x,y)){ok=1;break;}
      }
      hit+=ok; total++;
    }
    return hit/total;
  };
  const innerInk=(cx,cy,r)=>{
    let hit=0,total=0;
    const rr=Math.max(2,Math.floor(r*.48));
    for(let yy=-rr;yy<=rr;yy++)for(let xx=-rr;xx<=rr;xx++){
      if(xx*xx+yy*yy>rr*rr)continue;
      total++; hit+=dark(cx+xx,cy+yy);
    }
    return total?hit/total:1;
  };

  const x0=p[0], y0=p[1];
  const xMin=Math.max(8,Math.floor(x0-70)), xMax=Math.min(w-8,Math.ceil(x0+70));
  const yMin=Math.max(0,Math.floor(y0+8)), yMax=Math.min(h-8,Math.ceil(y0+190));
  const factor=1; // canvas en viewport hebben dezelfde schaal
  let best=null;

  for(let cy=yMin;cy<=yMax;cy+=2){
    for(let cx=xMin;cx<=xMax;cx+=2){
      const dx=Math.abs(cx-x0), dy=cy-y0;
      if(dx>70 || dy<8 || dy>190)continue;
      for(const r of [3,4,5,6,7,8,9,10,11,12]){
        const rs=ringScore(cx,cy,r);
        if(rs<0.55)continue;
        const inside=innerInk(cx,cy,r);
        // Een genummerd rondje heeft een duidelijke ring en wat inkt in het midden.
        // Lange rechte lijnen scoren slechter door de ringtest.
        const score=rs*100 + Math.min(inside,0.35)*18 - dx*1.0 - Math.abs(dy-55)*0.22 - Math.abs(r-7)*1.2;
        if(!best || score>best.score){best={x:cx,y:cy,r,score,rs,inside};}
      }
    }
  }

  if(!best || best.rs<0.55)return null;
  return {
    x:best.x, y:best.y, connectionY:best.y,
    page:item.page, wca:item.wca,
    fromPniCircle:true, circleNumberHint:true,
    circleScore:best.rs
  };
}

function findPniCircleForWca(item,viewport,circles){
  if(!circles||!circles.length)return null;
  const p=viewportPointForText(item,viewport);
  const candidates=circles.map(c=>({
    c,
    dx:Math.abs(c.x-p[0]),
    dy:c.y-p[1]
  }))
  // Alleen een WCA boven het bovenste aansluitgedeelte mag meedoen.
  .filter(v=>v.dy>0 && v.dy<Math.max(260,viewport.height*0.16))
  .sort((a,b)=>a.dx-b.dx || a.dy-b.dy);
  if(!candidates.length)return null;
  const best=candidates[0];
  // De WCA staat verticaal boven hetzelfde rondje. Een grote horizontale
  // afwijking betekent dat dit niet hetzelfde aansluitpunt is.
  if(best.dx>Math.max(55,viewport.width*0.045))return null;
  return {
    x:best.c.x,
    y:best.c.y,
    connectionY:best.c.y,
    page:item.page,
    wca:item.wca,
    fromPniCircle:true,
    circleNumberHint:true,
    circleScore:best.c.ring,
    circleIndex:best.c.index
  };
}

async function findPniAddressWithOcr(pageNumber=1){
  try{
    const page=await pdf.getPage(pageNumber);
    const localScale=4.0;
    const vp=page.getViewport({scale:localScale});
    const src=document.createElement('canvas');
    src.width=Math.ceil(vp.width); src.height=Math.ceil(vp.height);
    await page.render({canvasContext:src.getContext('2d'),viewport:vp}).promise;

    // Adres staat bovenaan. Beperk OCR daarom tot het bovenste deel.
    const crop=document.createElement('canvas');
    crop.width=src.width;
    crop.height=Math.min(src.height,Math.floor(src.height*0.55));
    crop.getContext('2d').drawImage(src,0,0,crop.width,crop.height,0,0,crop.width,crop.height);

    const data=await ocrCanvas(crop);
    const lines=(data.lines||[])
      .map(l=>String(l.text||'').replace(/\s+/g,' ').trim())
      .filter(Boolean);

    const streetWord=/(straat|weg|laan|hof|plein|dreef|singel|kade|gracht|dijk|park|boulevard|steeg|pad|baan|ring|markt|plantsoen|kamp|erf|zoom|weide|horst|dam)/i;
    const candidates=[];
    for(const line of lines){
      const cleaned=line
        .replace(/\b(1108|1111|1112|1113|1114)\s*(?:HP|G[Hh]|Amsterdam|Diemen)?\b/ig,' ')
        .replace(/\s+/g,' ').trim();
      const m=cleaned.match(/([A-Za-zÀ-ÖØ-öø-ÿ'’.-]+(?:\s+[A-Za-zÀ-ÖØ-öø-ÿ'’.-]+){0,4})\s+(\d{1,4}[A-Za-z]?)(?:\b|$)/);
      if(!m)continue;
      const street=m[1].trim();
      const number=m[2];
      if(!streetWord.test(street))continue;
      candidates.push(`${street} ${number}`.replace(/\s+/g,' '));
    }

    if(candidates.length){
      addressText=candidates[0];
      addressInfo.innerHTML=
        `<span class="ok">Adres automatisch gevonden:</span><br>${escapeHtml(addressText)}`;
      return addressText;
    }
  }catch(e){
    console.warn('PNI adres OCR mislukt',e);
  }
  return '';
}

async function autoPniToDaa(){
  if(!pdf||workflow.active)return;

  status.textContent='PNI: genummerde aansluitrondjes worden gezocht...';
  stepEl.textContent='PNI zoekt eerst de fysieke rondjes en leest daarna alleen de WCA-tekst direct erboven.';
  autoWca.disabled=true;
  if(autoPni)autoPni.disabled=true;

  try{
    // Bij deze afbeelding-PDF's is de PDF-tekstlaag leeg. Daarom géén globale
    // WCA-OCR gebruiken: die leest ook de WCA's bij de woningen en veroorzaakt
    // de foutmeldingen WCA019/WCA1111 enzovoort.
    const page=await pdf.getPage(currentPage);
    const viewport=page.getViewport({scale});
    const canvas=document.createElement('canvas');
    canvas.width=Math.ceil(viewport.width);
    canvas.height=Math.ceil(viewport.height);
    await page.render({canvasContext:canvas.getContext('2d'),viewport}).promise;

    const circles=findPniTopNumberedCircles(canvas,viewport);
    if(circles.length<2){
      status.textContent='PNI: geen betrouwbare rij met genummerde rondjes gevonden.';
      stepEl.textContent='Controleer of de bovenste aansluitrondjes zichtbaar zijn.';
      return;
    }

    // OCR alleen in een smalle strook rond ieder fysiek rondje.
    // Hierdoor kan de tekst onderaan bij de woningen niet worden meegenomen.
    const found=[];
    const missing=[];
    for(const circle of circles){
      const label=await ocrPniWcaNearCircle(currentPage,viewport,circle);
      if(!label){
        missing.push(`rondje ${circle.index}`);
        continue;
      }
      const m=label.match(/^WCA(\d{1,4})([A-D])$/i);
      if(!m)continue;
      const n=Number(m[1]);
      const letter=m[2].toUpperCase();
      if(n<1 || n>8)continue;
      found.push({circle,label:`WCA${String(n).padStart(3,'0')}${letter}`,letter});
    }

    if(!found.length){
      status.textContent='PNI: de rondjes zijn gevonden, maar de WCA-teksten konden niet worden gelezen.';
      stepEl.textContent='Controleer de leesbaarheid van de WCA-teksten boven de rondjes.';
      return;
    }

    // DAA's volgen de fysieke volgorde links → rechts. Dit is bewust NIET
    // gebaseerd op het WCA-nummer, want dezelfde WCA kan meerdere keren
    // voorkomen (zoals WCA001A op rondje 1 en 4 in ASD-WC0338).
    let added=0;
    for(const item of found.sort((a,b)=>a.circle.x-b.circle.x)){
      const daa=`DAA${String(added+1).padStart(3,'0')}`;
      const c={
        page:currentPage,
        daa,
        daaPoint:{page:currentPage,x:item.circle.x,y:getDaaStandardY()},
        wcaPoint:{page:currentPage,x:item.circle.x,y:item.circle.y},
        autoWca:true,
        wca:item.label,
        wcaLetter:item.letter,
        pni:true,
        pniCircle:true,
        pniCircleIndex:item.circle.index
      };

      if(item.letter==='A'||item.letter==='C'){
        c.output1={page:currentPage,x:item.circle.x,y:item.circle.y,autoWca:true,wca:item.label,wcaPoint:c.wcaPoint,pniCircle:true};
      }else{
        c.output2={page:currentPage,x:item.circle.x,y:item.circle.y,autoWca:true,wca:item.label,wcaPoint:c.wcaPoint,pniCircle:true};
      }

      changes.push(c);
      componentHistory.push({type:"daa",item:c});
      added++;
    }

    if(greenRowY===null && found.length){
      greenRowY=changes[changes.length-found.length].daaPoint.y+60*scale;
    }

    // Als de normale adresherkenning geen tekstlaag had, probeer alleen voor
    // PNI nog het adresblok bovenaan via OCR.
    if(!addressText){
      await findPniAddressWithOcr(currentPage);
    }

    await renderPage();
    status.textContent=`PNI klaar: ${added} fysieke aansluiting(en) gekoppeld.`;
    const lijst=found
      .sort((a,b)=>a.circle.x-b.circle.x)
      .map((x,i)=>`${x.label} → DAA${String(i+1).padStart(3,'0')}`)
      .join(' | ');
    stepEl.textContent=missing.length
      ? `Gevonden: ${lijst}. Niet gelezen: ${missing.join(', ')}.`
      : `Gevonden: ${lijst}. Alleen de bovenste WCA's zijn gebruikt.`;
  }catch(err){
    console.error(err);
    status.textContent='PNI automatisch koppelen is niet gelukt.';
    stepEl.textContent='Controleer de PDF en probeer opnieuw.';
  }finally{
    autoWca.disabled=!pdf||workflow.active||!wcaItems.length||(blokschemaDirectoryHandle&&blokschemaType==="PNI");
    if(autoPni)autoPni.disabled=!pdf||workflow.active;
    updateButtons();
  }
}

async function autoWcaToDaa(){
  if(!pdf||!wcaItems.length||workflow.active)return;

  status.textContent="WCA-lijnen worden automatisch gezocht...";
  stepEl.textContent="Iedere WCA-lijn wordt afzonderlijk gekoppeld aan een DAA-uitgang.";
  autoWca.disabled=true;
  if(autoPni) autoPni.disabled=true;

  try{
    let added=0;
    const missing=[];
    const pageCanvases=new Map();

    for(const item of wcaItems){
      const match=item.wca.replace(/\s+/g,"").match(/^WCA(\d{1,4})([A-Z])$/i);
      if(!match)continue;

      const nummer=match[1];
      const letter=match[2].toUpperCase();
      if(!["A","B","C","D"].includes(letter))continue;

      const daa=`DAA${nummer}`;

      // Iedere WCA wordt afzonderlijk geplaatst.
      // A en B hoeven dus niet naast elkaar te liggen.
      if(changes.some(c =>
        c.autoWca &&
        c.wca===item.wca &&
        c.page===item.page
      )){
        continue;
      }

      let rendered=pageCanvases.get(item.page);
      if(!rendered){
        const page=await pdf.getPage(item.page);
        const viewport=page.getViewport({scale});
        const canvas=document.createElement("canvas");
        canvas.width=viewport.width;
        canvas.height=viewport.height;

        await page.render({
          canvasContext:canvas.getContext("2d"),
          viewport
        }).promise;

        rendered={canvas,viewport};
        pageCanvases.set(item.page,rendered);
      }

      // Het zoeken van het bloemetje mag NOOIT de hele automatische
      // verwerking stoppen. Bij een afwijkende WCA-positie gebruiken we
      // altijd een fallback op basis van de tekstpositie.
      let flower=null;
      try{
        // Eerst zoeken we het genummerde rondje direct onder de WCA-tekst.
        // Dat is het daadwerkelijke aansluitpunt voor de DAA.
        flower=findNumberedCircleNearWca(item,rendered.viewport);

        // Als OCR het nummer niet heeft gelezen, gebruik de bestaande
        // beeldanalyse als fallback.
        if(!flower){
          flower=findFlowerOnRenderedPage(
            rendered.canvas,
            item,
            rendered.viewport
          );
        }
      }catch(flowerErr){
        console.warn(`Bloemetje zoeken mislukt bij ${item.wca}:`,flowerErr);
      }

      if(!flower){
        const pt=viewportPointForText(item,rendered.viewport);
        flower={
          x:pt[0],
          y:pt[1]+12,
          connectionY:pt[1]+12,
          page:item.page,
          wca:item.wca,
          fallback:true
        };
        missing.push(`${item.wca}: fallback gebruikt`);
      }

      // A/C = DAAxxx-01, B/D = DAAxxx-02.
      const connection={
        page:item.page,
        daa,
        daaPoint:{
          page:item.page,
          x:flower.x,
          y:getDaaStandardY()
        },
        wcaPoint:{
          page:item.page,
          x:flower.x,
          y:flower.y
        },
        autoWca:true,
        wca:item.wca,
        wcaLetter:letter
      };

      if(letter==="A" || letter==="C"){
        connection.output1={
          ...flower,
          y:flower.connectionY,
          autoWca:true,
          wca:item.wca,
          wcaPoint:connection.wcaPoint
        };
        connection.output2=null;
      }else{
        connection.output2={
          ...flower,
          y:flower.connectionY,
          autoWca:true,
          wca:item.wca,
          wcaPoint:connection.wcaPoint
        };
        connection.output1=null;
      }

      if(greenRowY===null){
        greenRowY=connection.daaPoint.y+60*scale;
      }

      changes.push(connection);
      componentHistory.push({type:"daa",item:connection});
      added++;
    }

    await renderPage();

    status.textContent=added
      ? `${added} WCA-lijn(en) automatisch gekoppeld.`
      : "Geen nieuwe WCA-lijnen gekoppeld.";

    stepEl.textContent=missing.length
      ? `Automatisch klaar. Controleer: ${missing.join(" | ")}`
      : "Automatisch WCA → DAA klaar.";

    if(missing.length)console.warn("WCA automatisch:",missing);

  }catch(err){
    console.error(err);
    status.textContent="Automatisch WCA → DAA is niet gelukt.";
    stepEl.textContent="Controleer de gevonden WCA-lijnen.";
    alert(
      "De automatische WCA → DAA koppeling kon niet worden uitgevoerd.\n\n"+
      err.message
    );
  }finally{
    autoWca.disabled=!pdf||workflow.active||!wcaItems.length||(blokschemaDirectoryHandle&&blokschemaType==="PNI");
  if(autoPni) autoPni.disabled=!pdf||workflow.active;
    updateButtons();
  }
}

autoWca.addEventListener("click",autoWcaToDaa);
if(autoPni) autoPni.addEventListener("click",autoPniToDaa);

function buildDaaList(){
  daaSelect.innerHTML='<option value="">Kies DAA...</option>';

  // We werken maximaal met DAA001 t/m DAA008.
  for(let i=1;i<=8;i++){
    const daa=`DAA${String(i).padStart(3,"0")}`;
    const option=document.createElement("option");
    option.value=daa;
    option.textContent=daa;
    daaSelect.appendChild(option);
  }

  daaSelect.disabled=false;
}

/* ---------------- RENDEREN ---------------- */

async function renderPage(){
  if(!pdf)return;

  const page=await pdf.getPage(currentPage);
  const viewport=page.getViewport({scale});

  pdfCanvas.width=viewport.width;
  pdfCanvas.height=viewport.height;
  overlay.width=viewport.width;
  overlay.height=viewport.height;

  pageBox.style.width=viewport.width+"px";
  pageBox.style.height=viewport.height+"px";

  await page.render({
    canvasContext:pdfCanvas.getContext("2d"),
    viewport
  }).promise;

  pageInfo.textContent=`${currentPage} / ${pdf.numPages}`;

  drawOverlay();
  updateButtons();
}

function deactivateEraser(){
  eraserActive=false;
  eraserBtn.classList.remove("red");
  eraserBtn.classList.add("gray");
  eraserBtn.textContent="Gum";
  overlay.style.cursor="grab";
}

eraserBtn.addEventListener("click",()=>{
  eraserActive=!eraserActive;
  eraserBtn.classList.toggle("red",eraserActive);
  eraserBtn.classList.toggle("gray",!eraserActive);
  eraserBtn.textContent=eraserActive?"Gum actief":"Gum";
  overlay.style.cursor=eraserActive?"crosshair":"grab";
  status.textContent=eraserActive
    ? "Gum actief: sleep over oude tekst of markeringen om ze weg te halen."
    : "Gum uitgeschakeld.";
});

function updateButtons(){
  prevBtn.disabled=!pdf||currentPage<=1;
  nextBtn.disabled=!pdf||currentPage>=pdf.numPages;

  startDaa.disabled=!pdf||workflow.active;
  autoWca.disabled=!pdf||workflow.active||!wcaItems.length||(blokschemaDirectoryHandle&&blokschemaType==="PNI");
  if(autoPni) autoPni.disabled=!pdf||workflow.active;
  startDivider2Btn.disabled=!pdf||workflow.active;
  startDivider3Btn.disabled=!pdf||workflow.active;
  startDivider4Btn.disabled=!pdf||workflow.active;
  startExistingDividerBtn.disabled=!pdf||workflow.active;
  startAftakkerBtn.disabled=!pdf||workflow.active;

  deleteDaaBtn.disabled=!pdf;
  eraserBtn.disabled=!pdf;
  saveBtn.disabled=!pdf;
  clearBtn.disabled=!(changes.length || dividers.length || aftakkers.length || eraserStrokes.length);
}

/* ---------------- WORKFLOW ---------------- */

daaSelect.addEventListener("change",()=>{
  updateButtons();
});


function beginDividerPlacement(type, outputCount=2, existing=false, existingStyle=null){
  if(!pdf || workflow.active)return;
  deactivateEraser();
  dividerChoicePanel.style.display="none";

  workflow={
    active:true, daa:null, daaPoint:null,
    output1Point:null, output2Point:null, outputPoints:[],
    mode:"dividerPlace", divider:type, dividerCount:outputCount, existingDivider:!!existing, existingStyle:existingStyle, dividerPoint:null,
    divider1Point:null, divider2Point:null, daaConnection:null
  };

  status.textContent=`${type} gekozen. Klik op de plaats waar de verdeler moet komen.`;
  stepEl.textContent=`${outputCount} verdeler ${type}: klik op de plaats waar de verdeler moet komen.`;
  updateButtons();
}

function beginAftakkerPlacement(){
  if(!pdf || workflow.active)return;

  // Eerst het type kiezen. ?? wordt daarna vervangen door het ingevoerde aantal DB.
  const typeInput=prompt("Type aftakker", "OT-VCX-??/PTP");
  if(typeInput===null)return;
  const gekozenType=String(typeInput).trim();
  if(!gekozenType){ alert("Vul een type aftakker in."); return; }

  const db=prompt("Hoeveel DB", "");
  if(db===null)return;
  const value=String(db).trim();
  if(!/^\d+(?:[.,]\d+)?$/.test(value)){ alert("Vul een getal in bij Hoeveel DB."); return; }

  const dbWaarde=value.replace(",",".");
  const aftakkerType=gekozenType.replace(/\?\?/g, dbWaarde);

  deactivateEraser();
  dividerChoicePanel.style.display="none";
  workflow={
    active:true, daa:null, daaPoint:null,
    output1Point:null, output2Point:null,
    mode:"aftakkerPlace", aftakkerPoint:null, aftakkerOutput1:null, aftakkerOutput2:null,
    aftakkerDb:value.replace(",","."), aftakkerType
  };
  status.textContent=`Aftakker ${workflow.aftakkerType}: klik op de plaats waar de aftakker moet komen.`;
  stepEl.textContent="Aftakker: klik op de plaats waar de aftakker moet komen.";
  updateButtons();
}

startAftakkerBtn.addEventListener("click",()=>{
  if(!pdf || workflow.active)return;
  beginAftakkerPlacement();
});

// Kies eerst welk type verdeler geplaatst moet worden.
function openDividerChoice(count, existing=false){
  if(!pdf || workflow.active)return;
  dividerChoicePanel.dataset.outputCount=String(count);
  dividerChoicePanel.dataset.existing=existing ? "1" : "0";

  dividerChoicePanel.querySelectorAll("button[data-main-divider]").forEach(btn=>{
    btn.style.display = !existing && btn.dataset.dividerCount===String(count) ? "block" : "none";
  });
  dividerChoicePanel.querySelectorAll("button[data-existing-divider-type]").forEach(btn=>{
    btn.style.display = existing ? "block" : "none";
  });
  dividerCustomType.style.display = existing ? "none" : "block";
  dividerChoicePanel.style.display="flex";
  status.textContent=existing ? "Kies het type bestaande verdeler." : `Kies de ${count} verdeler.`;
  stepEl.textContent=existing ? "Kies het type bestaande verdeler." : `Kies het type ${count} verdeler.`;
  updateButtons();
}

startDivider2Btn.addEventListener("click",()=>openDividerChoice(2,false));
startDivider3Btn.addEventListener("click",()=>openDividerChoice(3,false));
startDivider4Btn.addEventListener("click",()=>openDividerChoice(4,false));
startExistingDividerBtn.addEventListener("click",()=>openDividerChoice(2,true));

dividerChoiceCancel.addEventListener("click",()=>{
  dividerChoicePanel.style.display="none";
  updateButtons();
});

const dividerCustomType=document.getElementById("dividerCustomType");
if(dividerCustomType){
  dividerCustomType.addEventListener("click",()=>{
    const count=Number(dividerChoicePanel.dataset.outputCount||2);
    const existing=dividerChoicePanel.dataset.existing==="1";
    const type=prompt("Type verdeler", "");
    if(type===null)return;
    const customType=String(type).trim();
    if(!customType){ alert("Vul een type verdeler in."); return; }
    beginDividerPlacement(customType,count,existing);
  });
}

// De drie keuzes van de knop Bestaande verdeler.
dividerChoicePanel.querySelectorAll("button[data-existing-divider-type]").forEach(btn=>{
  btn.addEventListener("click",()=>{
    const style=btn.dataset.existingDividerType;
    if(style==="aftakker"){
      const value=prompt("Vul de ?? in voor OT-VCX-??/PTP (t-o)","");
      if(value===null)return;
      const code=String(value).trim();
      if(!/^[A-Za-z0-9]+$/.test(code)){
        alert("Vul alleen de code in die op de plek van ?? moet komen.");
        return;
      }
      beginDividerPlacement(`OT-VCX-${code}/PTP (t-o)`,2,true,"aftakker");
    }else if(style==="twee"){
      beginDividerPlacement("OS-VCX-02/PTP",2,true,"twee");
    }else{
      beginDividerPlacement("OT-VCX-00/PTP",1,true,"een");
    }
  });
});

dividerChoicePanel.querySelectorAll("button[data-main-divider]").forEach(btn=>{
  btn.addEventListener("click",()=>{
    const count=Number(dividerChoicePanel.dataset.outputCount||2);
    const existing=dividerChoicePanel.dataset.existing==="1";
    beginDividerPlacement(btn.dataset.mainDivider,count,existing);
  });
});

startDaa.addEventListener("click",()=>{
  if(!pdf || workflow.active)return;
  deactivateEraser();

  // Bij iedere handmatige start eerst expliciet vragen welke DAA geplaatst moet worden.
  // Zo kan ook een overgeslagen nummer (bijvoorbeeld DAA003) alsnog worden geplaatst.
  const used=new Set(changes.map(c=>c.daa).filter(Boolean));
  const free=[];
  for(let i=1;i<=8;i++){
    const d=`DAA${String(i).padStart(3,"0")}`;
    if(!used.has(d)) free.push(d);
  }
  if(!free.length){
    alert("DAA001 t/m DAA008 zijn allemaal geplaatst.");
    return;
  }

  // Toon de bestaande DAA-keuze. Uitgang 1 staat standaard geselecteerd.
  manualSelectedOutput=1;
  manualOut1.style.opacity="1";
  manualOut2.style.opacity=".55";
  manualDaaNumber.innerHTML=free.map(d=>`<option value="${d}">${d}</option>`).join("");
  manualDaaNumber.value=free[0];
  manualDaaPanel.style.display="flex";
  stepEl.textContent="Kies de DAA. Uitgang 1 wordt automatisch als eerste geplaatst.";
  status.textContent="Welke DAA wil je als eerste plaatsen? Uitgang 1 wordt automatisch geplaatst.";
  updateButtons();
});

manualOut1.addEventListener("click",()=>{ manualSelectedOutput=1; manualOut1.style.opacity="1"; manualOut2.style.opacity=".55"; });
manualOut2.addEventListener("click",()=>{ manualSelectedOutput=2; manualOut1.style.opacity=".55"; manualOut2.style.opacity="1"; });
manualDaaCancel.addEventListener("click",()=>{ manualDaaPanel.style.display="none"; });
let sequentialDaaNumber=null;
let sequentialDaaStage=null;

function getNextFreeDaa(){
  const used=new Set(changes.map(c=>c.daa).filter(Boolean));
  for(let i=1;i<=8;i++){
    const d=`DAA${String(i).padStart(3,"0")}`;
    if(!used.has(d)) return d;
  }
  return null;
}

function askOutput2Question(){
  daaQuestionText.textContent=`${sequentialDaaNumber} uitgang 1 is geplaatst. Wil je uitgang 2 plaatsen?`;
  daaYes.dataset.action="output2";
  daaNo.dataset.action="after-output1-no";
  daaQuestion.style.display="flex";
}

function askNextDaaQuestionSequential(){
  const next=getNextFreeDaa();
  if(!next){
    daaQuestionText.textContent="Alle 8 DAA's zijn geplaatst. Je kunt de PDF nu handmatig opslaan met de knop Opslaan.";
    daaYes.dataset.action="finish-daa";
    daaNo.dataset.action="finish-daa";
    daaQuestion.style.display="flex";
    return;
  }
  daaQuestionText.textContent=`${sequentialDaaNumber} is klaar. Wil je de volgende DAA plaatsen?`;
  daaYes.dataset.action="next-sequential";
  daaNo.dataset.action="finish-daa";
  daaQuestion.style.display="flex";
}

function startSequentialNextDaa(){
  // Na DAA001 komt altijd DAA002, daarna DAA003 enzovoort.
  // We lopen dus bewust oplopend door DAA001 t/m DAA008.
  const current=sequentialDaaNumber
    ? parseInt(sequentialDaaNumber.replace("DAA",""),10)
    : 0;
  const nextNumber=current+1;
  if(nextNumber>8){
    daaQuestionText.textContent="Alle DAA's DAA001 t/m DAA008 zijn geplaatst. Je kunt de PDF nu handmatig opslaan met de knop Opslaan.";
    daaYes.dataset.action="finish-daa";
    daaNo.dataset.action="finish-daa";
    daaQuestion.style.display="flex";
    return;
  }

  const next=`DAA${String(nextNumber).padStart(3,"0")}`;
  if(changes.some(c=>c.daa===next)){
    status.textContent=`${next} is al geplaatst.`;
    return;
  }

  sequentialDaaNumber=next;
  sequentialDaaStage="output1";
  workflow={active:true,daa:next,daaPoint:null,output1Point:null,output2Point:null,mode:"manualDaaLine",manualOutput:1};
  status.textContent=`Klik op de bestaande lijn voor ${next}-01.`;
  stepEl.textContent=`Plaats ${next} uitgang 1 op de bestaande WCA-lijn.`;
  updateButtons();
}

manualDaaContinue.addEventListener("click",()=>{
  const daa=manualDaaNumber.value;
  if(!daa)return;
  if(changes.some(c=>c.daa===daa)){ status.textContent=`${daa} is al geplaatst.`; return; }

  // Iedere nieuwe DAA begint altijd met uitgang 1.
  // Na het plaatsen volgt automatisch de vraag voor uitgang 2.
  sequentialDaaNumber=daa;
  sequentialDaaStage="output1";
  manualSelectedOutput=1;
  manualDaaPanel.style.display="none";
  workflow={active:true,daa,daaPoint:null,output1Point:null,output2Point:null,mode:"manualDaaLine",manualOutput:1};
  status.textContent=`Klik op de bestaande lijn voor ${daa}-01.`;
  stepEl.textContent=`Plaats ${daa} uitgang 1 op de bestaande WCA-lijn.`;
  updateButtons();
});
/*
  Eén klik doet afhankelijk van de stap iets anders:
  - DAA positie
  - uitgang 1 lijn
  - uitgang 2 lijn
*/

function getDaaConnectionAtPoint(x,y){
  for(let i=changes.length-1;i>=0;i--){
    const c=changes[i];
    if(c.page!==currentPage || !c.daaPoint || !c.autoWca) continue;
    const p=c.wcaPoint || c.output1 || c.output2;
    if(!p) continue;
    if(Math.hypot(x-p.x,y-p.y)<=16) return c;
  }
  return null;
}

function drawDaaConnectionHandle(c){
  const p=c && (c.wcaPoint || c.output1 || c.output2);
  if(!p || c.page!==currentPage) return;
  ctx.save();
  ctx.fillStyle="#18c900";
  ctx.strokeStyle="#fff";
  ctx.lineWidth=2;
  ctx.beginPath();
  ctx.arc(p.x,p.y,7,0,Math.PI*2);
  ctx.fill();
  ctx.stroke();
  ctx.restore();
}

function getAutoDaaAtPoint(x,y){
  // Alleen automatisch geplaatste DAA's zijn versleepbaar.
  for(let i=changes.length-1;i>=0;i--){
    const c=changes[i];
    if(c.page!==currentPage || !c.autoWca || !c.daaPoint)continue;

    const left=c.daaPoint.x-18;
    const right=c.daaPoint.x+18;
    const top=c.daaPoint.y-34;
    const bottom=c.daaPoint.y+24;

    if(x>=left && x<=right && y>=top && y<=bottom){
      return c;
    }
  }
  return null;
}

function findNearestAutoWcaLine(c,x,y){
  let best=null;
  let bestScore=Infinity;

  for(const other of changes){
    if(!other.autoWca || other.page!==currentPage || !other.wcaPoint)continue;
    if(other===c)continue;

    const lineX=other.wcaPoint.x;
    const dy=y-other.wcaPoint.y;
    const dx=x-lineX;

    // Alleen snappen wanneer de DAA echt boven/naast een bestaande
    // automatische WCA-lijn wordt neergezet.
    const score=Math.abs(dx)+Math.min(Math.abs(dy)*0.15,80);
    if(Math.abs(dx)<=35 && score<bestScore){
      bestScore=score;
      best=other;
    }
  }

  return best;
}

function finishDaaDrag(){
  if(!dragDaa)return;

  const c=dragDaa.c;
  const x=dragDaa.x;
  const y=getDaaStandardY();
  const target=findNearestAutoWcaLine(c,x,y);

  // De DAA blijft op de vaste DAA-hoogte.
  c.daaPoint.x=x;
  c.daaPoint.y=y;

  if(target){
    // Neem de bestaande WCA-lijn volledig over.
    // De DAA komt op de nieuwe X-positie, de groene verbinding loopt
    // daarna automatisch naar het nieuwe WCA-aansluitpunt.
    c.wca=target.wca;
    c.wcaLetter=target.wcaLetter;
    c.wcaPoint={...target.wcaPoint};

    if(c.wcaLetter==="A" || c.wcaLetter==="C"){
      c.output1={
        ...target.wcaPoint,
        page:currentPage,
        autoWca:true,
        wca:target.wca,
        wcaPoint:{...target.wcaPoint}
      };
      c.output2=null;
    }else{
      c.output2={
        ...target.wcaPoint,
        page:currentPage,
        autoWca:true,
        wca:target.wca,
        wcaPoint:{...target.wcaPoint}
      };
      c.output1=null;
    }

    status.textContent=`${c.daa} is verplaatst naar ${c.wca}.`;
    stepEl.textContent="DAA aangepast. De groene lijn is opnieuw gekoppeld.";
  }else{
    // Geen nieuwe WCA-lijn gevonden. De DAA mag dan vrij worden verplaatst,
    // maar het groene aansluitpunt blijft exact staan waar de gebruiker het
    // eerder heeft geplaatst. Alleen het groene punt zelf mag het
    // aansluitpunt veranderen.
    status.textContent=`${c.daa} is verplaatst.`;
    stepEl.textContent="Het groene aansluitpunt blijft staan. Sleep het groene bolletje zelf om de aansluiting te wijzigen.";
  }

  dragDaa=null;
  drawOverlay();
}

function getDividerAtPoint(x,y){
  for(let i=dividers.length-1;i>=0;i--){
    const d=dividers[i];
    if(d.page!==currentPage) continue;
    const cx=d.centerX;
    const ay=d.anchorY!=null ? d.anchorY : d.y+20;
    if(Math.hypot(x-cx,y-(ay-10))<=24 || (Math.abs(x-cx)<=14 && y>=ay-28 && y<=ay+10)) return d;
  }
  return null;
}

let dragDivider=null;
let dragDividerOutput=null;
let dragAftakkerOutput=null;
let dragAftakker=null;


function getDividerOutputAtPoint(x,y){
  for(let i=dividers.length-1;i>=0;i--){
    const d=dividers[i];
    if(d.page!==currentPage) continue;
    const maxOutputs=d.dividerCount || (d.output3 ? 3 : d.output4 ? 4 : 2);
    for(let output=1;output<=maxOutputs;output++){
      const p=d.outputs?.[output-1] || d[`output${output}`];
      if(!p) continue;
      if(Math.hypot(x-p.x,y-p.y)<=14){
        return {d,output,p};
      }
    }
  }
  return null;
}

function getAftakkerAtPoint(x,y){
  for(let i=aftakkers.length-1;i>=0;i--){
    const a=aftakkers[i];
    if(a.page!==currentPage) continue;
    const ay=a.anchorY!=null ? a.anchorY : a.y+20;
    const cx=a.centerX;
    // Pak het rode aftakkersymbool of de tekst er direct naast,
    // maar niet het losse groene aansluitpunt.
    if((Math.abs(x-cx)<=14 && y>=ay-34 && y<=ay+8) ||
       (x>=cx+10 && x<=cx+105 && y>=ay-32 && y<=ay+2)){
      return a;
    }
  }
  return null;
}

function getAftakkerOutputAtPoint(x,y){
  for(let i=aftakkers.length-1;i>=0;i--){
    const a=aftakkers[i];
    if(a.page!==currentPage) continue;
    for(const output of [1,2]){
      const p=output===1 ? a.output1 : a.output2;
      if(p && Math.hypot(x-p.x,y-p.y)<=14) return {a,p,output};
    }
  }
  return null;
}

function finishAftakkerOutputDrag(){
  if(!dragAftakkerOutput)return;
  const item=dragAftakkerOutput;
  const a=item.a;
  const x=item.x;
  const y=item.y;
  const point={page:currentPage,x,y};

  // Koppel alleen wanneer er daadwerkelijk een bestaande WCA-lijn onder het punt zit.
  const temp={page:currentPage,wcaPoint:{page:currentPage,x,y},wca:"HANDMATIG",wcaLetter:"A"};
  const target=findNearestAutoWcaLine(temp,x,y);
  if(target){
    point.x=target.wcaPoint.x;
    point.y=target.wcaPoint.y;
    status.textContent=`Aftakker is gekoppeld aan ${target.wca}.`;
  }else{
    status.textContent="Aftakker-uitgang is verplaatst.";
  }
  if(item.output===1) a.output1=point; else a.output2=point;
  dragAftakkerOutput=null;
  drawOverlay();
}

function finishDividerOutputDrag(){
  if(!dragDividerOutput)return;
  const item=dragDividerOutput;
  const d=item.d;
  const output=item.output;
  const x=item.x;
  const y=item.y;

  const temp={
    page:currentPage,
    autoWca:true,
    wcaPoint:{page:currentPage,x,y},
    wca:"HANDMATIG",
    wcaLetter:output===1 ? "A" : "B"
  };
  const target=findNearestAutoWcaLine(temp,x,y);

  const point={page:currentPage,x,y};
  if(target){
    point.x=target.wcaPoint.x;
    point.y=target.wcaPoint.y;
    status.textContent=`Uitgang ${output} van ${d.type} is gekoppeld aan ${target.wca}.`;
    stepEl.textContent=`Uitgang ${output} blijft verplaatsbaar via het groene rondje.`;
  }else{
    status.textContent=`Uitgang ${output} van ${d.type} is verplaatst.`;
    stepEl.textContent="Sleep het groene rondje naar de gewenste aansluiting.";
  }

  if(!d.outputs) d.outputs=[d.output1,d.output2].filter(Boolean);
  d.outputs[output-1]=point;
  d[`output${output}`]=point;

  dragDividerOutput=null;
  drawOverlay();
}

overlay.addEventListener("pointerdown",e=>{
  const rect=overlay.getBoundingClientRect();
  const x=e.clientX-rect.left;
  const y=e.clientY-rect.top;

  // Gummodus: klik en sleep om oude PDF-tekst/markeringen wit weg te halen.
  if(eraserActive){
    eraserDrag={pointerId:e.pointerId};
    overlay.setPointerCapture?.(e.pointerId);
    overlay.style.cursor="crosshair";
    addEraserStroke(x,y);
    return;
  }

  // Het groene aansluitpunt van een aftakker is afzonderlijk versleepbaar.
  if(!workflow.active){
    const aftakkerOutput=getAftakkerOutputAtPoint(x,y);
    if(aftakkerOutput){
      dragAftakkerOutput={
        a:aftakkerOutput.a,
        offsetX:x-aftakkerOutput.p.x,
        offsetY:y-aftakkerOutput.p.y,
        x:aftakkerOutput.p.x,
        y:aftakkerOutput.p.y
      };
      overlay.setPointerCapture?.(e.pointerId);
      overlay.style.cursor="grabbing";
      status.textContent="Aansluitpunt van de aftakker wordt verplaatst.";
      stepEl.textContent="Sleep het groene rondje naar de gewenste lijn.";
      return;
    }
  }

  // De aftakker zelf kan worden verplaatst. Het groene aansluitpunt
  // blijft op zijn bestaande positie staan, zodat alleen de aftakker
  // en de horizontaal/verticale verbinding meebewegen.
  if(!workflow.active){
    const aftakker=getAftakkerAtPoint(x,y);
    if(aftakker){
      const ay=aftakker.anchorY!=null ? aftakker.anchorY : aftakker.y+20;
      dragAftakker={a:aftakker,offsetX:x-aftakker.centerX,offsetY:y-ay};
      overlay.setPointerCapture?.(e.pointerId);
      overlay.style.cursor="grabbing";
      status.textContent="Aftakker wordt verplaatst.";
      stepEl.textContent="Sleep de aftakker naar de gewenste plaats. Het groene aansluitpunt blijft staan.";
      return;
    }
  }

  // De groene rondjes van de verdeler zijn afzonderlijk versleepbaar,
  // net als het groene aansluitpunt van een DAA.
  if(!workflow.active){
    const dividerOutput=getDividerOutputAtPoint(x,y);
    if(dividerOutput){
      dragDividerOutput={
        d:dividerOutput.d,
        output:dividerOutput.output,
        offsetX:x-dividerOutput.p.x,
        offsetY:y-dividerOutput.p.y,
        x:dividerOutput.p.x,
        y:dividerOutput.p.y
      };
      overlay.setPointerCapture?.(e.pointerId);
      overlay.style.cursor="grabbing";
      status.textContent=`Uitgang ${dividerOutput.output} van ${dividerOutput.d.type} wordt verplaatst.`;
      stepEl.textContent="Sleep het groene rondje naar de gewenste lijn.";
      return;
    }
  }

  // Een geplaatste verdeler kan net als een DAA worden verplaatst.
  // De gekozen uitgangspunten blijven aan hun bestaande lijnen gekoppeld.
  if(!workflow.active){
    const divider=getDividerAtPoint(x,y);
    if(divider){
      const ay=divider.anchorY!=null ? divider.anchorY : divider.y+20;
      dragDivider={d:divider,offsetX:x-divider.centerX,offsetY:y-ay};
      overlay.setPointerCapture?.(e.pointerId);
      overlay.style.cursor="grabbing";
      status.textContent=`${divider.type} wordt verplaatst.`;
      stepEl.textContent="Sleep de verdeler naar de gewenste plaats.";
      return;
    }
  }

  // Het onderste groene aansluitpunt is afzonderlijk versleepbaar.
  if(!workflow.active){
    const connection=getDaaConnectionAtPoint(x,y);
    if(connection){
      const p=connection.wcaPoint || connection.output1 || connection.output2;
      dragDaaConnection={
        c:connection,
        offsetX:x-p.x,
        offsetY:y-p.y,
        x:p.x,
        y:p.y
      };
      overlay.setPointerCapture?.(e.pointerId);
      overlay.style.cursor="grabbing";
      status.textContent=`Aansluitpunt van ${connection.daa} wordt verplaatst.`;
      stepEl.textContent="Sleep het groene punt naar de bestaande WCA-lijn.";
      return;
    }
  }

  // Na de automatische plaatsing kan de gebruiker een DAA slepen.
  // Sleep je hem boven een andere automatische WCA-lijn, dan wordt
  // de DAA bij loslaten aan die lijn gekoppeld.
  if(!workflow.active){
    const c=getAutoDaaAtPoint(x,y);
    if(c){
      dragDaa={
        c,
        offsetX:x-c.daaPoint.x,
        offsetY:y-c.daaPoint.y,
        x:c.daaPoint.x,
        y:c.daaPoint.y
      };
      overlay.setPointerCapture?.(e.pointerId);
      overlay.style.cursor="grabbing";
      status.textContent=`${c.daa} wordt verplaatst. Sleep naar een andere WCA-lijn.`;
      stepEl.textContent="Laat de DAA los boven de gewenste WCA-lijn.";
      return;
    }
    return;
  }

  const rect2=overlay.getBoundingClientRect();


  if(workflow.mode==="manualDaaLine"){
    // Eén klik op de bestaande lijn. De DAA komt op dezelfde vaste
    // hoogte als de automatisch geplaatste DAA's en de groene lijn
    // loopt vanuit het midden onder de DAA rechtstreeks naar het
    // aangeklikte punt. Er zijn geen tweede/derde klik of twee
    // afzonderlijke uitgangsverbindingen meer nodig.
    const yDaa=getDaaStandardY();
    const c={
      page:currentPage,
      daa:workflow.daa,
      daaPoint:{page:currentPage,x:x,y:yDaa},
      wcaPoint:{page:currentPage,x:x,y:y},
      output1:workflow.manualOutput===1 ? {page:currentPage,x:x,y:y,manual:true} : null,
      output2:workflow.manualOutput===2 ? {page:currentPage,x:x,y:y,manual:true} : null,
      autoWca:true,
      manualDaa:true,
      manualOutput:workflow.manualOutput,
      wca:`HANDMATIG`,
      wcaLetter:"A"
    };

    changes.push(c);
    componentHistory.push({type:"daa",item:c});

    // Eerste handmatige DAA bepaalt de vaste groene aansluitrij.
    if(greenRowY===null) greenRowY=c.daaPoint.y+60*scale;

    workflow={active:false,daa:null,daaPoint:null,
      output1Point:null,output2Point:null,mode:null};

    status.textContent=`${c.daa} geplaatst op de aangeklikte lijn.`;
    stepEl.textContent="DAA geplaatst.";
    drawOverlay();
    updateButtons();

    if(sequentialDaaStage==="output1"){
      sequentialDaaStage="waitingOutput2";
      askOutput2Question();
    }else if(sequentialDaaStage==="output2"){
      sequentialDaaStage="done";
      askNextDaaQuestionSequential();
    }
    return;
  }

  if(workflow.mode==="aftakkerPlace"){
    workflow.aftakkerPoint={page:currentPage,x:x,y:y,anchorY:y};
    workflow.mode="aftakkerOutput1";
    status.textContent="Aftakker staat. Kies nu uitgang 1.";
    stepEl.textContent="Aftakker: klik op de lijn voor uitgang 1.";
    drawOverlay();
    updateButtons();
    return;
  }

  if(workflow.mode==="aftakkerOutput1"){
    workflow.aftakkerOutput1={page:currentPage,x:x,y:y};
    workflow.mode="aftakkerOutput2";
    status.textContent="Uitgang 1 gekozen. Kies nu uitgang 2.";
    stepEl.textContent="Aftakker: klik op de lijn voor uitgang 2.";
    drawOverlay();
    updateButtons();
    return;
  }

  if(workflow.mode==="aftakkerOutput2"){
    workflow.aftakkerOutput2={page:currentPage,x:x,y:y};
    const p=workflow.aftakkerPoint;
    const nieuwAftakker={
      page:currentPage,
      centerX:p.x,
      y:p.y-20,
      anchorY:p.y,
      db:workflow.aftakkerDb,
      type:workflow.aftakkerType,
      output1:workflow.aftakkerOutput1,
      output2:workflow.aftakkerOutput2
    };
    aftakkers.push(nieuwAftakker);
    componentHistory.push({type:"aftakker",item:nieuwAftakker});
    workflow={active:false,daa:null,daaPoint:null,
      output1Point:null,output2Point:null,mode:null};
    status.textContent="Aftakker geplaatst met uitgang 1 en uitgang 2.";
    stepEl.textContent="Aftakker geplaatst.";
    drawOverlay();
    updateButtons();
    return;
  }

  if(workflow.mode==="dividerPlace"){
    // Klik 1 = plaats van de verdeler.
    // De bovenkant van de verdeler komt 1 cm onder de
    // horizontale groene DAA-lijn.
    const cmPx=(72/2.54)*scale;
    const dividerTopOffset=28; // bovenkant rood symbool t.o.v. center

    const basisY=greenRowY!==null
      ? greenRowY
      : (daaRowY!==null ? daaRowY : y);

    // De eerste klik bepaalt exact waar de verticale groene lijn eindigt.
    // Het verdelersymbool wordt direct boven dit aangeklikte punt gezet.
    dividerRowY=y;

    workflow.dividerPoint={
      page:currentPage,
      x:x,
      y:y,
      anchorY:y
    };
    workflow.mode="divider1";

    status.textContent="Verdeler staat. Klik nu op de bestaande lijn voor uitgang 1.";
    stepEl.textContent=`Verdeler ${workflow.divider}: klik op lijn 1.`;
    drawOverlay();
    updateButtons();
    return;
  }

  if(workflow.mode==="divider1"){
    // Bij een 2 verdeler blijft de bestaande werking behouden:
    // rechts = uitgang 1 en links = uitgang 2. Bij 3/4 verdelers
    // worden de uitgangen in de vaste volgorde gekozen.
    const firstOutput=(workflow.dividerCount===2 && workflow.existingStyle!=="aftakker")
      ? (x < workflow.dividerPoint.x ? 2 : 1)
      : 1;
    workflow.outputPoints=[{page:currentPage,x,y,output:firstOutput}];
    workflow.mode="dividerOutputs";
    if(workflow.dividerCount===2 && workflow.existingStyle!=="aftakker"){
      const andere=firstOutput===1 ? 2 : 1;
      status.textContent=`Uitgang ${firstOutput} gekozen. Klik nu ${andere===1 ? "rechts" : "links"} van de verdeler voor uitgang ${andere}.`;
      stepEl.textContent=`2 verdeler ${workflow.divider}: kies uitgang ${andere}.`;
    }else{
      status.textContent="Uitgang 1 gekozen (links). Klik nu op uitgang 2.";
      stepEl.textContent=`${workflow.dividerCount} verdeler ${workflow.divider}: klik op uitgang 2.`;
    }
    drawOverlay();
    updateButtons();
    return;
  }

  if(workflow.mode==="dividerOutputs"){
    const n=workflow.outputPoints.length+1;
    let outputNumber=n;
    if(workflow.dividerCount===2 && workflow.existingStyle!=="aftakker"){
      const first=workflow.outputPoints[0];
      const side=x < workflow.dividerPoint.x ? 2 : 1;
      if(side===first.output){
        status.textContent=`Klik aan de andere kant van de verdeler voor uitgang ${side===1?2:1}.`;
        stepEl.textContent="2 verdeler: kies de andere zijde.";
        return;
      }
      outputNumber=side;
    }
    workflow.outputPoints.push({page:currentPage,x,y,output:outputNumber});

    if(n < workflow.dividerCount){
      const labels=workflow.dividerCount===3
        ? ["links","midden onder","rechts"]
        : ["links","links schuin onder","rechts schuin onder","rechts"];
      status.textContent=`Uitgang ${n} gekozen (${labels[n-1]}). Klik nu op uitgang ${n+1} (${labels[n]}).`;
      stepEl.textContent=`${workflow.dividerCount} verdeler ${workflow.divider}: kies uitgang ${n+1} (${labels[n]}).`;
      drawOverlay();
      updateButtons();
      return;
    }

    // Alle uitgangen zijn gekozen. Bepaal het midden van de verdeler op basis
    // van de buitenste aansluitingen, net zoals bij de bestaande 2 verdeler.
    const xs=workflow.outputPoints.map(p=>p.x);
    const centerX=(Math.min(...xs)+Math.max(...xs))/2;
    const daaConnection=workflow.daaConnection || null;
    const nearestDaaOutput=daaConnection
      ? getNearestDaaOutputForDivider({centerX:centerX,y:workflow.dividerPoint.y},daaConnection)
      : null;

    const outputs={};
    workflow.outputPoints.forEach((p,i)=>outputs[`output${i+1}`]=p);

    const nieuweVerdeler={
      page:currentPage,
      type:workflow.divider,
      existingDivider:!!workflow.existingDivider,
      existingStyle:workflow.existingStyle || null,
      dividerCount:workflow.dividerCount,
      centerX:centerX,
      y:workflow.dividerPoint.anchorY-20,
      anchorY:workflow.dividerPoint.anchorY,
      greenY:dividerGreenRowY,
      output1:outputs.output1 || null,
      output2:outputs.output2 || null,
      output3:outputs.output3 || null,
      output4:outputs.output4 || null,
      outputs:workflow.outputPoints.slice(),
      daaConnection:daaConnection,
      daaOutputUsed:nearestDaaOutput ? nearestDaaOutput.output : null
    };
    dividers.push(nieuweVerdeler);
    componentHistory.push({type:"divider",item:nieuweVerdeler});

    const geplaatst=workflow.divider;
    const aantal=workflow.dividerCount;
    workflow={active:false,daa:null,daaPoint:null,
      output1Point:null,output2Point:null,mode:null};

    status.textContent=`${geplaatst} geplaatst met ${aantal} uitgangen.`;
    stepEl.textContent=`${aantal} verdeler geplaatst.`;
    drawOverlay();
    updateButtons();
    setTimeout(showNextDaaQuestion,100);
    return;
  }

  if(workflow.mode==="daa"){
    // De DAA staat standaard met de BOVENKANT 2 cm vanaf
    // de bovenkant van de PDF. De X-positie blijft door de klik bepaald.
    if(daaRowY === null){
      daaRowY = getDaaStandardY();
    }

    workflow.daaPoint={
      page:currentPage,
      x:x,
      y:daaRowY
    };

    workflow.mode="output1";

    status.textContent=
      `Klik nu op de bestaande lijn voor ${workflow.daa}-01.`;

    stepEl.textContent=
      `Stap 4: klik op de lijn voor ${workflow.daa}-01 (uitgang 1).`;

    drawOverlay();
    return;
  }

  if(workflow.mode==="output1"){
    workflow.output1Point={page:currentPage,x,y};

    // Alle DAA's gebruiken dezelfde horizontale groene aansluitrij.
    if(greenRowY === null){
      greenRowY = y;
    }

    workflow.mode="output2";

    status.textContent=
      `Uitgang 1 gekozen. Klik nu op de bestaande lijn voor ${workflow.daa}-02.`;

    stepEl.textContent=
      `Stap 5: klik op de lijn voor ${workflow.daa}-02 (uitgang 2).`;

    drawOverlay();
    return;
  }

  if(workflow.mode==="output2"){
    workflow.output2Point={page:currentPage,x,y};

    // Centreer het DAA-symbool exact tussen de twee gekozen uitgangen.
    // Links = uitgang 2, rechts = uitgang 1.
    workflow.daaPoint.x =
      (workflow.output1Point.x + workflow.output2Point.x) / 2;

    workflow.mode="done";

    status.textContent=
      `Beide lijnen gekozen. Het DAA-symbool is automatisch gecentreerd.`;

    stepEl.textContent=
      `Stap 6: kies daarna de gewenste verdeler.`;

    drawOverlay();
    updateButtons();
  }
});

overlay.addEventListener("pointermove",e=>{
  const rect=overlay.getBoundingClientRect();
  if(!dragDivider && !dragAftakker && !dragDaa && !dragDaaConnection && !workflow.active){
    const hx=e.clientX-rect.left;
    const hy=e.clientY-rect.top;
    const nextHover=getAutoDaaAtPoint(hx,hy);
    if(nextHover!==hoveredDaa){
      hoveredDaa=nextHover;
      deleteDaaBtn.disabled=!hoveredDaa || hoveredDaa.page!==currentPage;
    }
  }

  if(eraserDrag){
    const x=e.clientX-rect.left;
    const y=e.clientY-rect.top;
    addEraserStroke(x,y);
    return;
  }

  if(dragAftakkerOutput){
    const x=e.clientX-rect.left-dragAftakkerOutput.offsetX;
    const y=e.clientY-rect.top-dragAftakkerOutput.offsetY;
    dragAftakkerOutput.x=x;
    dragAftakkerOutput.y=y;
    const p=dragAftakkerOutput.output===1 ? dragAftakkerOutput.a.output1 : dragAftakkerOutput.a.output2;
    if(p){ p.x=x; p.y=y; }
    drawOverlay();
    return;
  }

  if(dragAftakker){
    const a=dragAftakker.a;
    a.centerX=e.clientX-rect.left-dragAftakker.offsetX;
    a.anchorY=e.clientY-rect.top-dragAftakker.offsetY;
    a.y=a.anchorY-20;
    drawOverlay();
    return;
  }

  if(dragDividerOutput){
    const x=e.clientX-rect.left-dragDividerOutput.offsetX;
    const y=e.clientY-rect.top-dragDividerOutput.offsetY;
    dragDividerOutput.x=x;
    dragDividerOutput.y=y;
    const p=dragDividerOutput.output===1 ? dragDividerOutput.d.output1 : dragDividerOutput.d.output2;
    if(p){
      p.x=x;
      p.y=y;
    }
    drawOverlay();
    return;
  }

  if(dragDivider){
    const d=dragDivider.d;
    d.centerX=e.clientX-rect.left-dragDivider.offsetX;
    d.anchorY=e.clientY-rect.top-dragDivider.offsetY;
    d.y=d.anchorY-20;
    drawOverlay();
    return;
  }

  if(dragDaaConnection){
    const x=e.clientX-rect.left-dragDaaConnection.offsetX;
    const y=e.clientY-rect.top-dragDaaConnection.offsetY;
    dragDaaConnection.x=x;
    dragDaaConnection.y=y;
    dragDaaConnection.c.wcaPoint={page:currentPage,x:x,y:y};
    drawOverlay();
    return;
  }

  if(!dragDaa)return;

  dragDaa.x=e.clientX-rect.left-dragDaa.offsetX;
  dragDaa.y=e.clientY-rect.top-dragDaa.offsetY;
  dragDaa.c.daaPoint.x=dragDaa.x;
  dragDaa.c.daaPoint.y=dragDaa.y;

  drawOverlay();
});

overlay.addEventListener("pointercancel",e=>{
  eraserDrag=null;
  overlay.style.cursor=eraserActive?"crosshair":"grab";
});

overlay.addEventListener("pointerup",e=>{
  if(eraserDrag){
    try{ overlay.releasePointerCapture?.(e.pointerId); }catch(err){}
    eraserDrag=null;
    overlay.style.cursor=eraserActive?"crosshair":"grab";
    status.textContent="Gum toegepast. Sleep opnieuw om meer oude inhoud te verwijderen.";
    return;
  }

  if(dragAftakkerOutput){
    try{ overlay.releasePointerCapture?.(e.pointerId); }catch(err){}
    overlay.style.cursor="grab";
    finishAftakkerOutputDrag();
    return;
  }

  if(dragAftakker){
    const a=dragAftakker.a;
    a.y=a.anchorY-20;
    try{ overlay.releasePointerCapture?.(e.pointerId); }catch(err){}
    dragAftakker=null;
    overlay.style.cursor="grab";
    status.textContent="Aftakker is verplaatst.";
    stepEl.textContent="Het groene aansluitpunt blijft op zijn bestaande plaats.";
    drawOverlay();
    return;
  }

  if(dragDividerOutput){
    try{ overlay.releasePointerCapture?.(e.pointerId); }catch(err){}
    overlay.style.cursor="grab";
    finishDividerOutputDrag();
    return;
  }

  if(dragDivider){
    const d=dragDivider.d;
    d.y=d.anchorY-20;
    try{ overlay.releasePointerCapture?.(e.pointerId); }catch(err){}
    dragDivider=null;
    overlay.style.cursor="grab";
    status.textContent=`${d.type} is verplaatst.`;
    stepEl.textContent="De groene verdelerlijn is meeverplaatst.";
    drawOverlay();
    return;
  }

  if(dragDaaConnection){
    const c=dragDaaConnection.c;
    const target=findNearestAutoWcaLine(c,dragDaaConnection.x,dragDaaConnection.y);

    if(target){
      c.wca=target.wca;
      c.wcaLetter=target.wcaLetter;
      c.wcaPoint={...target.wcaPoint};

      if(c.wcaLetter==="A" || c.wcaLetter==="C"){
        c.output1={...target.wcaPoint,page:currentPage,autoWca:true,wca:target.wca,wcaPoint:{...target.wcaPoint}};
        c.output2=null;
      }else{
        c.output2={...target.wcaPoint,page:currentPage,autoWca:true,wca:target.wca,wcaPoint:{...target.wcaPoint}};
        c.output1=null;
      }

      status.textContent=`${c.daa} is gekoppeld aan ${c.wca}.`;
      stepEl.textContent="Groene lijn gekoppeld aan de bestaande WCA-lijn.";
    }

    try{ overlay.releasePointerCapture?.(e.pointerId); }catch(err){}
    dragDaaConnection=null;
    overlay.style.cursor="grab";
    drawOverlay();
    return;
  }

  if(!dragDaa)return;
  try{ overlay.releasePointerCapture?.(e.pointerId); }catch(err){}
  finishDaaDrag();
});

overlay.addEventListener("pointercancel",e=>{
  if(dragAftakkerOutput){
    dragAftakkerOutput=null;
    overlay.style.cursor="grab";
    drawOverlay();
    return;
  }
  if(dragAftakker){
    dragAftakker=null;
    overlay.style.cursor="grab";
    drawOverlay();
    return;
  }
  if(dragDividerOutput){
    dragDividerOutput=null;
    overlay.style.cursor="grab";
    drawOverlay();
    return;
  }
  if(dragDaaConnection){
    dragDaaConnection=null;
    overlay.style.cursor="grab";
    drawOverlay();
    return;
  }
  if(!dragDaa)return;
  try{ overlay.releasePointerCapture?.(e.pointerId); }catch(err){}
  dragDaa=null;
  overlay.style.cursor="grab";
  drawOverlay();
});

overlay.addEventListener("pointerleave",()=>{
  if(!dragDivider && !dragDaa && !dragDaaConnection && !workflow.active){
    hoveredDaa=null;
    deleteDaaBtn.disabled=true;
  }
});

/* ---------------- UITLIJNEN ---------------- */

function alignAllBlocksForSave(){
  // Bij opslaan worden UITSLUITEND de DAA-symbolen uitgelijnd.
  // De groene eind-/aansluitbolletjes van de DAA's blijven exact staan
  // waar de gebruiker ze heeft geplaatst. Verdelers en aftakkers worden
  // helemaal niet verplaatst of uitgelijnd.
  const pageDaa=changes.filter(c=>c.page===currentPage && c.daaPoint);
  if(pageDaa.length){
    const targetDaaY=pageDaa[0].daaPoint.y;

    for(const c of pageDaa){
      if(c.daaPoint) c.daaPoint.y=targetDaaY;
      // BELANGRIJK: c.wcaPoint, c.output1 en c.output2 NIET aanpassen.
      // Deze zijn de afzonderlijke groene eindpunten en mogen hun eigen
      // hoogte behouden.
    }

    daaRowY=targetDaaY;
    greenRowY=targetDaaY+96*scale;
  }

  // Niets aanpassen aan groene eindpunten, verdelers of aftakkers.
  drawOverlay();
  status.textContent="Alleen de DAA-symbolen zijn uitgelijnd. Eindbolletjes, verdelers en aftakkers behouden hun eigen positie.";
  stepEl.textContent="De PDF wordt nu opgeslagen zonder eindbolletjes, verdelers of aftakkers uit te lijnen.";
}

/* ---------------- PDF OPSLAAN ---------------- */

/*
 * Maakt van de huidige ORCA-tekening een nieuwe PDF.
 * De oorspronkelijke PDF wordt opnieuw gerenderd en de DAA-wijzigingen
 * worden daar als definitieve laag bovenop geplaatst.
 */
saveBtn.addEventListener("click",async()=>{
  if(!pdf)return;

  // Een blokschema mag niet worden opgeslagen zonder gekoppeld Excelbestand,
  // omdat de Excel-registratie direct na het opslaan van de PDF moet gebeuren.
  if(!excelFileHandle){
    status.textContent="Excel niet gekoppeld – eerst Excel koppelen.";
    stepEl.textContent="Koppel eerst het Excelbestand voordat je het blokschema opslaat.";
    alert("Excel is niet gekoppeld.\n\nKoppel eerst het Excelbestand met de knop ‘Excel opslaan met verdelers’. Daarna kun je het blokschema opslaan en wordt Excel automatisch bijgewerkt.");
    return;
  }

  // Controleer ook of de browser nog schrijftoegang heeft.
  try{
    if(excelFileHandle.queryPermission){
      let permission=await excelFileHandle.queryPermission({mode:"readwrite"});
      if(permission!=="granted" && excelFileHandle.requestPermission){
        permission=await excelFileHandle.requestPermission({mode:"readwrite"});
      }
      if(permission!=="granted") throw new Error("Geen schrijftoestemming voor het gekoppelde Excelbestand.");
    }
  }catch(err){
    status.textContent="Excel niet gekoppeld – geen schrijftoegang.";
    stepEl.textContent="Koppel het Excelbestand opnieuw.";
    alert("Excel is niet gekoppeld of er is geen schrijftoegang meer.\n\nKoppel het Excelbestand opnieuw met de knop ‘Excel opslaan met verdelers’.\n\nHet PDF-bestand is nog niet opgeslagen.");
    excelFileHandle=null;
    excelBtn.textContent="Excel opslaan met verdelers";
    return;
  }

  alignAllBlocksForSave();
  await saveFinishedPdf();
});

let rememberedSaveDirectory = null;

async function chooseSaveDirectory(){
  if(!window.showDirectoryPicker || !window.isSecureContext) return null;
  try{
    const handle = await window.showDirectoryPicker({mode:"readwrite"});
    if(handle.requestPermission){
      const permission = await handle.requestPermission({mode:"readwrite"});
      if(permission !== "granted") return null;
    }
    rememberedSaveDirectory = handle;
    return handle;
  }catch(err){
    if(err && err.name === "AbortError") return null;
    console.warn("Mapkiezer mislukt:", err);
    return null;
  }
}

async function getRememberedSaveDirectory(){
  if(!rememberedSaveDirectory) return null;
  try{
    let permission = rememberedSaveDirectory.queryPermission
      ? await rememberedSaveDirectory.queryPermission({mode:"readwrite"})
      : "granted";
    if(permission !== "granted" && rememberedSaveDirectory.requestPermission){
      permission = await rememberedSaveDirectory.requestPermission({mode:"readwrite"});
    }
    if(permission !== "granted"){
      rememberedSaveDirectory = null;
      return null;
    }
    return rememberedSaveDirectory;
  }catch(err){
    rememberedSaveDirectory = null;
    return null;
  }
}

async function saveBlobToRememberedDirectory(blob, filename){
  let dirHandle = await getRememberedSaveDirectory();
  if(!dirHandle) dirHandle = await chooseSaveDirectory();
  if(!dirHandle) return false;

  try{
    const fileHandle = await dirHandle.getFileHandle(filename,{create:true});
    const writable = await fileHandle.createWritable();
    await writable.write(blob);
    await writable.close();
    return true;
  }catch(err){
    console.warn("Opslaan in map mislukt:",err);
    rememberedSaveDirectory = null;
    return false;
  }
}

async function saveFinishedPdf(){
  if(!pdf){
    alert("Er is geen PDF geopend.");
    return;
  }

  status.textContent="PDF wordt opgebouwd...";
  stepEl.textContent="PDF wordt klaargemaakt om op te slaan.";

  try{
    const { jsPDF } = window.jspdf;

    // BEHOUD HET ORIGINELE PAPIERFORMAAT VAN DE BRON-PDF.
    // Belangrijk: de render-schaal (2x) mag NIET worden gebruikt als
    // papiermaat. De vorige versie deed dat wel en maakte daardoor een
    // veel te groot document (bijvoorbeeld A0). PDF-punten bij scale=1
    // zijn de echte afmetingen van de originele PDF-pagina.
    const firstPage = await pdf.getPage(1);
    const firstSize = firstPage.getViewport({scale:1.0});

    const orientation =
      firstSize.width >= firstSize.height ? "landscape" : "portrait";

    const pdfOut = new jsPDF({
      orientation,
      unit:"pt",
      format:[firstSize.width, firstSize.height],
      compress:true
    });

    for(let pageNumber=1; pageNumber<=pdf.numPages; pageNumber++){
      const page = await pdf.getPage(pageNumber);
      // Dit is de echte PDF-pagina-afmeting in punten.
      const pageSize = page.getViewport({scale:1.0});
      // 4x wordt alleen gebruikt om de inhoud extra scherp te renderen.
      const viewport = page.getViewport({scale:4.0});

      const canvas = document.createElement("canvas");
      canvas.width = viewport.width;
      canvas.height = viewport.height;

      const pageCtx = canvas.getContext("2d");

      await page.render({
        canvasContext:pageCtx,
        viewport
      }).promise;

      /*
       * Kop linksboven: wis eerst de originele PDF-kop en teken daarna
       * precies één keer onze eigen kop: BESTANDSNAAM gevolgd door ADRES.
       * Zo blijft de bestandsnaam altijd als eerste zichtbaar en ontstaat
       * geen dubbele adresregel.
       */
      if(pageNumber===1 && (fileNameText || addressText)){
        pageCtx.save();
        pageCtx.fillStyle="#fff";
        pageCtx.fillRect(0,0,Math.min(900,viewport.width),105);
        pageCtx.restore();
      }

      /*
       * HOGE KWALITEIT VOOR DE BEWERKINGSLAGEN
       *
       * De editor werkt op schaal 1.5. In de oude export werd die laag eerst
       * op 1.5x gerenderd en daarna vergroot naar de 3x PDF-render. Vooral
       * bij WCA -> DAA automatisch waren de DAA's, teksten en groene lijnen
       * daardoor zichtbaar zachter dan de originele PDF.
       *
       * Teken de bewerkingslaag nu rechtstreeks op 3x. De coördinaten blijven
       * die van de editor (1.5x), maar de context schaalt ze naar 2x. Daardoor
       * blijven alle posities exact hetzelfde en wordt niets verschoven.
       */
      const editScale = 4.0;
      const editorScale = scale;
      const editViewport = page.getViewport({scale:editScale});
      const editCanvas = document.createElement("canvas");
      editCanvas.width = Math.ceil(editViewport.width);
      editCanvas.height = Math.ceil(editViewport.height);
      // Bewaar de logische editor-afmetingen voor functies die hun positie
      // berekenen op basis van de canvasgrootte, zoals de materiaaltabel.
      editCanvas._logicalWidth = pageSize.width * editorScale;
      editCanvas._logicalHeight = pageSize.height * editorScale;
      const editCtx = editCanvas.getContext("2d");

      const oldCtx = ctx;
      const oldCurrentPage = currentPage;
      ctx = editCtx;
      currentPage = pageNumber;

      // Zet de tekenlaag van 1.5x editor-coördinaten naar 2x export.
      const overlayFactor = editScale / editorScale;
      editCtx.setTransform(overlayFactor,0,0,overlayFactor,0,0);

      // Gebruik exact dezelfde tekenlaag als in de editor.
      // Teken de kop op de definitieve PDF één keer mee.
      drawOverlay(pageNumber===1);

      ctx = oldCtx;
      currentPage = oldCurrentPage;

      pageCtx.save();
      pageCtx.drawImage(
        editCanvas,
        0, 0, editCanvas.width, editCanvas.height,
        0, 0, viewport.width, viewport.height
      );
      pageCtx.restore();

      if(pageNumber>1){
        pdfOut.addPage(
          [pageSize.width, pageSize.height],
          pageSize.width >= pageSize.height ? "landscape" : "portrait"
        );
      }

      const image = canvas.toDataURL("image/png");
      // Plaats de 2x gerenderde afbeelding terug op de originele
      // PDF-pagina. Hierdoor blijft A4/A3/etc. exact behouden.
      pdfOut.addImage(
        image,
        "PNG",
        0,
        0,
        pageSize.width,
        pageSize.height
      );
    }

    // Laat de gebruiker zelf de doelmap kiezen en slaat daarna automatisch
    // de PDF met de voorgestelde bestandsnaam in die map op.
    const blob = pdfOut.output("blob");

    const origineleNaam =
      fileInput.files && fileInput.files[0]
        ? fileInput.files[0].name
        : "ORCA.pdf";

    const origineleStem = origineleNaam.replace(/\.pdf$/i, "").trim();
    const nodeMatch = origineleStem.match(/\b(?:ASD|DIM|UT)-WC\d{3,6}\b/i);
    const basisNaam = (nodeMatch ? nodeMatch[0].toUpperCase() : (fileNameText || origineleStem)).trim();
    const filename = basisNaam + " (nieuw).pdf";

    // Gebruik de mapkiezer wanneer de browser dit veilig ondersteunt.
    // Bij een lokaal geopend HTML-bestand kan showDirectoryPicker wel bestaan,
    // maar toch een SecurityError geven omdat de pagina niet in een secure
    // context draait. In dat geval moet Opslaan alsnog gewoon werken via de
    // normale browser-download.
    if(window.showDirectoryPicker && window.isSecureContext){
      const saved = await saveBlobToRememberedDirectory(blob, filename);
      if(saved){
        status.textContent=`PDF opgeslagen als ${filename}.`;
        stepEl.textContent="PDF opgeslagen.";

        // Als een Excelbestand één keer is gekoppeld, werk het automatisch
        // bij zodra dit blokschema als PDF is opgeslagen.
        if(excelFileHandle){
          try{
            const excelFile=await excelFileHandle.getFile();
            await fillExcelWithGreenDividers(excelFile,excelFileHandle,true);
          }catch(excelErr){
            console.error("Automatische Excel-update mislukt:",excelErr);
            alert("De PDF is opgeslagen, maar Excel kon niet automatisch worden bijgewerkt.\n\n"+(excelErr.message||excelErr));
          }
        }

        // Als een map met blokschema's is gekozen, ga na het opslaan
        // automatisch door naar het volgende originele PDF-bestand.
        if(blokschemaDirectoryHandle){
          await loadNextBlokschema();
        }else{
          stepEl.textContent="Klaar. De gekozen opslagmap wordt voor deze sessie onthouden.";
        }
        return;
      }

      if(!rememberedSaveDirectory){
        status.textContent="Opslaan geannuleerd.";
        stepEl.textContent="Geen opslagmap gekozen. Je kunt opnieuw op Opslaan klikken.";
        return;
      }
    }

    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    a.style.display = "none";
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(()=>URL.revokeObjectURL(url),10000);

    status.textContent=`PDF gedownload als ${filename}.`;
    stepEl.textContent="PDF opgeslagen. Excel wordt nu bijgewerkt.";

    // Ook bij de normale browser-download (zonder gekozen opslagmap)
    // moet de gekoppelde Excel direct worden bijgewerkt. Voorheen gebeurde
    // dit alleen in de map-opslagroute, waardoor Excel ongemoeid bleef.
    if(excelFileHandle){
      try{
        const excelFile=await excelFileHandle.getFile();
        const excelUpdated=await fillExcelWithGreenDividers(excelFile,excelFileHandle,true);
        if(excelUpdated){
          status.textContent=`PDF gedownload als ${filename}; Excel is bijgewerkt.`;
          stepEl.textContent="PDF opgeslagen en de gekoppelde Excel-rij is bijgewerkt.";
        }else{
          status.textContent=`PDF gedownload als ${filename}; geen materiaal-aantallen om in Excel te zetten.`;
          stepEl.textContent="Controleer of er verdelers, DAA's of extra materialen zijn geplaatst.";
        }
      }catch(excelErr){
        console.error("Automatische Excel-update mislukt:",excelErr);
        status.textContent=`PDF gedownload als ${filename}; Excel-update mislukt.`;
        stepEl.textContent="Excel kon niet worden bijgewerkt. Koppel Excel opnieuw en probeer het nogmaals.";
        alert("De PDF is opgeslagen, maar Excel kon niet automatisch worden bijgewerkt.\n\n"+(excelErr.message||excelErr));
      }
    }else{
      status.textContent=`PDF gedownload als ${filename}; Excel is niet gekoppeld.`;
      stepEl.textContent="Koppel het Excelbestand om de aantallen automatisch op te slaan.";
    }

    // Ook na de gewone browserdownload (fallback) automatisch verdergaan
    // met het volgende bronbestand uit de gekozen map. Voorheen gebeurde
    // dit uitsluitend wanneer de File System Access-opslagroute slaagde.
    if(blokschemaDirectoryHandle){
      await loadNextBlokschema();
    }

  }catch(err){
    console.error(err);

    if(err && err.name==="AbortError"){
      status.textContent="Opslaan geannuleerd.";
      stepEl.textContent="Opslaan geannuleerd. Je kunt opnieuw opslaan.";
      return;
    }

    status.textContent="Fout bij het maken van de PDF.";
    stepEl.textContent="Er ging iets mis tijdens het opslaan.";
    alert("De PDF kon niet worden opgeslagen.\\n\\n"+err.message);
  }
}

function exportCoverAutoWcaCorridors(pageNumber,ctx2){
  const items=changes
    .filter(c=>c.page===pageNumber && c.autoWca && c.daaPoint)
    .slice()
    .sort((a,b)=>a.daaPoint.x-b.daaPoint.x);

  if(!items.length)return;

  for(let i=0;i<items.length;i++){
    const c=items[i];
    const x=c.daaPoint.x;
    const bottom=c.daaPoint.y+40;
    const wcaY=c.wcaPoint ? c.wcaPoint.y : (c.output1||c.output2||{}).y;
    if(!Number.isFinite(wcaY))continue;

    let left=x-55;
    let right=x+55;
    if(i>0)left=(items[i-1].daaPoint.x+x)/2+3;
    if(i<items.length-1)right=(x+items[i+1].daaPoint.x)/2-3;

    const top=bottom+2;
    const height=Math.max(0,(wcaY-12)-top);
    const width=Math.max(0,right-left);
    if(width<=0 || height<=0)continue;

    ctx2.save();
    ctx2.fillStyle="#fff";
    ctx2.fillRect(left,top,width,height);
    ctx2.restore();
  }
}

function exportCoverOldPoint(p,ctx2,viewport){
  if(!p)return;

  // Bij automatische WCA-aansluitingen wordt de oude inhoud
  // uitsluitend door het corridor-gebied tussen DAA en WCA verwijderd.
  // Nooit meer vanaf de bovenkant van de pagina, zodat bestandsnaam
  // en straatnaam boven de DAA behouden blijven.
  if(p.autoWca || p.manual)return;

  const rowY=greenRowY!==null
    ? greenRowY
    : p.y;

  const left=p.x-60;
  const top=0;
  const width=120;
  const height=Math.max(0,rowY-5);

  ctx2.save();
  ctx2.fillStyle="#fff";
  ctx2.fillRect(left,top,width,height);
  ctx2.restore();
}



function exportCoverDividerOldArea(d,ctx2,viewport){
  if(!d)return;

  const rowY=d.greenY!==null ? d.greenY : d.y;
  const centerX=d.centerX;

  const width=260;
  const left=centerX-width/2;
  const top=0;
  const height=Math.max(0,rowY-5);

  ctx2.save();
  ctx2.fillStyle="#fff";
  ctx2.fillRect(left,top,width,height);
  ctx2.restore();
}

function exportCoverDividerOldPoint(p,ctx2,viewport){
  if(!p)return;

  const lineWidth=14;
  const left=p.x-lineWidth/2;
  const top=Math.max(0,p.y-8);
  const width=lineWidth;
  const height=16;

  ctx2.save();
  ctx2.fillStyle="#fff";
  ctx2.fillRect(left,top,width,height);
  ctx2.restore();
}

function exportDaaToDividerConnection(d,ctx2){
  const c=d.daaConnection;
  if(!c || !c.daaPoint || greenRowY===null)return;

  const nearest=getNearestDaaOutputForDivider(
    {centerX:d.centerX, y:d.y},
    c
  );
  if(!nearest)return;

  const daaX=nearest.x;
  const outputY=nearest.y;
  const rowY=greenRowY;
  const halfCm=(72/2.54)*scale/2;
  const dividerTop=rowY+halfCm;

  ctx2.save();
  ctx2.strokeStyle="#18c900";
  ctx2.lineWidth=4;
  ctx2.lineCap="round";
  ctx2.lineJoin="round";

  ctx2.beginPath();
  ctx2.moveTo(daaX,outputY);
  ctx2.lineTo(daaX,rowY);
  ctx2.lineTo(d.centerX,rowY);
  ctx2.lineTo(d.centerX,dividerTop);
  ctx2.stroke();

  ctx2.restore();
}

function exportDividerOutputDown(x1,y1,x2,y2,color,ctx2,ownerDivider=null,outputNumber=null){
  // 4 verdeler: de twee binnenste uitgangen lopen eerst omlaag,
  // daarna horizontaal naar buiten en vervolgens recht omlaag.
  ctx2.save();
  ctx2.strokeStyle=color;
  ctx2.lineWidth=4;
  ctx2.lineCap="round";
  ctx2.lineJoin="round";

  ctx2.beginPath();
  if(ownerDivider?.dividerCount===4 && (outputNumber===2 || outputNumber===3)){
    const bendY=y1+28;
    ctx2.moveTo(x1,y1);
    ctx2.lineTo(x1,bendY);
    ctx2.lineTo(x2,bendY);
    ctx2.lineTo(x2,y2);
  }else{
    // Bestaande werking van de buitenste uitgangen en 2/3 verdeler behouden.
    ctx2.moveTo(x1,y1);
    ctx2.lineTo(x2,y1);
    ctx2.lineTo(x2,y2);
  }
  ctx2.stroke();
  ctx2.restore();
}

function exportDrawDivider(d,ctx2,viewport){
  if(greenRowY!==null){
    d.y=greenRowY+(72/2.54)*scale/2+28;
  }

  const x=d.centerX;
  const y=d.y;

  ctx2.save();

  ctx2.fillStyle="#ff0000";
  ctx2.font="20px Arial";
  ctx2.textAlign="left";
  ctx2.textBaseline="middle";
  ctx2.fillText(d.type,x+32,y-2);

  ctx2.strokeStyle="#ff0000";
  ctx2.lineWidth=3;
  ctx2.lineJoin="round";
  ctx2.lineCap="round";

  exportDaaToDividerConnection(d,ctx2);

  ctx2.beginPath();
  ctx2.moveTo(x-24,y-28);
  ctx2.lineTo(x+24,y-28);
  ctx2.lineTo(x+24,y+2);
  ctx2.lineTo(x+14,y+15);
  ctx2.lineTo(x-14,y+15);
  ctx2.lineTo(x-24,y+2);
  ctx2.closePath();
  ctx2.stroke();

  const dCount=d.dividerCount || (d.output4 ? 4 : d.output3 ? 3 : 2);
  const dOutputs=d.outputs || [d.output1,d.output2,d.output3,d.output4].slice(0,dCount);
  for(let oi=0;oi<dCount;oi++){
    const p=dOutputs[oi];
    if(!p) continue;
    const sp=getDividerOutputStart(d,oi+1);
    exportDividerOutputDown(sp.x,sp.y,p.x,p.y,"#18c900",ctx2,d,oi+1);
  }

  ctx2.restore();
}

function exportDrawDAA(c,ctx2,viewport){
  const x=c.daaPoint.x;
  const y=c.daaPoint.y;
  const red="#ff0000";
  const green="#18c900";

  ctx2.save();

  ctx2.fillStyle="#111";
  ctx2.font="24px Arial";
  ctx2.textAlign="center";
  ctx2.textBaseline="alphabetic";
  ctx2.fillText(c.daa,x,y-115);

  const boxW=76;
  const boxH=158;
  const left=x-boxW/2;
  const top=y-80;

  ctx2.strokeStyle=red;
  ctx2.lineWidth=2;
  ctx2.strokeRect(left,top,boxW,boxH);

  ctx2.beginPath();
  ctx2.arc(x,top+43,30,0,Math.PI*2);
  ctx2.stroke();

  exportArrowSymbol(x-13,top+25,35,ctx2);
  exportArrowSymbol(x+8,top+25,35,ctx2);

  ctx2.fillStyle=red;
  ctx2.fillRect(left-10,top+68,10,9);

  ctx2.beginPath();
  ctx2.moveTo(x-27,top+90);
  ctx2.lineTo(x+27,top+90);
  ctx2.lineTo(x,top+139);
  ctx2.closePath();
  ctx2.stroke();

  const bottom=top+boxH;
  const leftStartX=x-15;
  const rightStartX=x+15;

  // Oude tekst/gele markeringen binnen de twee DAA-uitgangslijnen wissen.
  exportCoverDaaOutputInterior(c,ctx2);

  if(c.autoWca){
    const p=c.output1 || c.output2;
    if(p){
      ctx2.save();
      ctx2.strokeStyle=green;
      ctx2.lineWidth=2;
      ctx2.lineCap="round";
      ctx2.lineJoin="round";
      const wcaY=c.wcaPoint ? c.wcaPoint.y : p.y;
      const lineX=c.wcaPoint ? c.wcaPoint.x : p.x;
      ctx2.beginPath();
      ctx2.moveTo(x,bottom);
      ctx2.lineTo(x,wcaY);
      if(lineX!==x)ctx2.lineTo(lineX,wcaY);
      ctx2.stroke();
      ctx2.restore();
    }
  }else{
    if(c.output2){
      exportGreenConnection(
        leftStartX,bottom+1,
        c.output2.x,c.output2.y,
        green,ctx2,
        greenRowY,
        96*scale
      );
    }

    if(c.output1){
      exportGreenConnection(
        rightStartX,bottom+1,
        c.output1.x,c.output1.y,
        green,ctx2,
        greenRowY,
        96*scale
      );
    }
  }

  ctx2.fillStyle="#111";
  ctx2.font="18px Arial";
  ctx2.textAlign="center";

  if(c.output2){
    exportVerticalText(`${c.daa}-02`,leftStartX-28,bottom+62,ctx2);
  }

  if(c.output1){
    exportVerticalText(`${c.daa}-01`,rightStartX+28,bottom+62,ctx2);
  }

  ctx2.restore();
}

function exportGreenConnection(x1,y1,x2,y2,color,ctx2,sharedY,verticalLengthPx=null){
  const rowY=y1+(verticalLengthPx!==null
    ? verticalLengthPx
    : ((72/2.54)*scale));
  ctx2.save();
  ctx2.strokeStyle=color; ctx2.lineWidth=4; ctx2.lineCap="round"; ctx2.lineJoin="round";
  ctx2.beginPath();
  ctx2.moveTo(x1,y1); ctx2.lineTo(x1,rowY); ctx2.lineTo(x2,rowY); ctx2.lineTo(x2,y2);
  ctx2.stroke(); ctx2.restore();
}

function exportVerticalText(text,x,y,ctx2){
  ctx2.save();
  ctx2.translate(x,y);
  ctx2.rotate(-Math.PI/2);

  // Geen witte achtergrond: de groene verbinding moet volledig doorlopen.
  ctx2.fillStyle="#111";
  ctx2.textAlign="center";
  ctx2.textBaseline="alphabetic";
  ctx2.fillText(text,0,0);
  ctx2.restore();
}

function exportArrowSymbol(x,y,length,ctx2){
  ctx2.beginPath();
  ctx2.moveTo(x,y);
  ctx2.lineTo(x+length*.55,y+length*.55);
  ctx2.lineTo(x+length*.35,y+length*.55);
  ctx2.lineTo(x+length,y+length);
  ctx2.stroke();

  ctx2.beginPath();
  ctx2.moveTo(x+length*.55,y+length*.55);
  ctx2.lineTo(x+length*.35,y+length*.55);
  ctx2.stroke();
}


daaYes.addEventListener("click",async()=>{
  const action=daaYes.dataset.action;
  daaQuestion.style.display="none";

  if(action==="output2"){
    sequentialDaaStage="output2";
    workflow={active:true,daa:sequentialDaaNumber,daaPoint:null,output1Point:null,output2Point:null,mode:"manualDaaLine",manualOutput:2};
    status.textContent=`Klik op de bestaande lijn voor ${sequentialDaaNumber}-02.`;
    stepEl.textContent=`Plaats ${sequentialDaaNumber} uitgang 2 op de bestaande WCA-lijn.`;
    updateButtons();
    return;
  }

  if(action==="next-sequential"){
    startSequentialNextDaa();
    return;
  }

  if(action==="divider"){
    // Na het plaatsen van een DAA dezelfde gewone 2-verdelerkeuze tonen.
    openDividerChoice(2,false);
    return;
  }

  if(action==="next-daa"){
    const laatste=changes[changes.length-1];
    const nr=parseInt(laatste.daa.replace("DAA",""),10)+1;
    if(nr>8){
      await saveFinishedPdf();
      return;
    }

    const volgende=`DAA${String(nr).padStart(3,"0")}`;
    daaSelect.value=volgende;
    status.textContent=
      `Klik op "DAA plaatsen" om ${volgende} te plaatsen.`;
    stepEl.textContent=`Stap 3: ${volgende} plaatsen.`;
    updateButtons();
    return;
  }

  // Bij NEE op verdeler: ga naar de bestaande DAA-vraag.
  if(action==="no-divider"){
    const laatste=changes[changes.length-1];
    if(!laatste)return;

    const nr=parseInt(laatste.daa.replace("DAA",""),10);

    if(nr>=8){
      status.textContent="Geen verdeler. PDF wordt opgeslagen.";
      stepEl.textContent="PDF opslaan.";
      await saveFinishedPdf();
    }else{
      daaQuestionText.textContent=
        `${laatste.daa} is klaar. Wil je nog een DAA plaatsen?`;
      daaYes.dataset.action="next-daa";
      daaNo.dataset.action="finish-daa";
      daaQuestion.style.display="flex";
    }
    return;
  }

  if(action==="finish-daa"){
    status.textContent="DAA plaatsen afgerond. Je kunt handmatig op Opslaan klikken.";
    stepEl.textContent="DAA geplaatst. Opslaan gebeurt alleen via de knop Opslaan.";
    updateButtons();
    return;
  }

  if(action==="save"){
    await saveFinishedPdf();
  }
});

daaNo.addEventListener("click",async()=>{
  const action=daaNo.dataset.action;
  daaQuestion.style.display="none";

  if(action==="after-output1-no"){
    // Geen uitgang 2: vraag eerst of de volgende DAA geplaatst moet worden.
    askNextDaaQuestionSequential();
    return;
  }

  if(action==="next-daa"){
    startSequentialNextDaa();
    return;
  }

  if(action==="finish-daa"){
    status.textContent="DAA plaatsen afgerond. Je kunt handmatig op Opslaan klikken.";
    stepEl.textContent="DAA geplaatst. Opslaan gebeurt alleen via de knop Opslaan.";
    updateButtons();
    return;
  }

  if(action==="no-divider"){
    const laatste=changes[changes.length-1];
    if(!laatste)return;

    const nr=parseInt(laatste.daa.replace("DAA",""),10);

    if(nr>=8){
      await saveFinishedPdf();
    }else{
      daaQuestionText.textContent=
        `${laatste.daa} is klaar. Wil je nog een DAA plaatsen?`;
      daaYes.dataset.action="next-daa";
      daaNo.dataset.action="finish-daa";
      daaQuestion.style.display="flex";
    }
    return;
  }

  await saveFinishedPdf();
});

/* ---------------- TEKENING ---------------- */

function drawEraserStrokes(pageNumber, targetCtx=ctx){
  for(const s of eraserStrokes){
    if(s.page!==pageNumber) continue;
    targetCtx.save();
    targetCtx.fillStyle="#fff";
    targetCtx.beginPath();
    targetCtx.arc(s.x,s.y,s.r,0,Math.PI*2);
    targetCtx.fill();
    targetCtx.restore();
  }
}

function addEraserStroke(x,y){
  eraserStrokes.push({page:currentPage,x,y,r:ERASER_RADIUS});
  drawOverlay();
}

function drawOverlay(includeHeader=true){
  ctx.clearRect(0,0,ctx.canvas.width,ctx.canvas.height);

  // De gum wist alleen de oude PDF-inhoud. Nieuwe DAA/verdelers worden
  // daarna opnieuw getekend en blijven daardoor zichtbaar.
  drawEraserStrokes(currentPage,ctx);

  // Bij automatische WCA-plaatsing wordt de oude inhoud tussen en naast
  // de nieuwe groene verticale lijnen eerst wit afgedekt. De WCA zelf
  // blijft zichtbaar doordat de afdekking net boven de WCA stopt.
  coverAutoWcaCorridors(currentPage);

  for(const c of changes){
    if(c.page!==currentPage) continue;

    // Eerst oude ORCA-aansluitingen verwijderen.
    coverOldPoint(c.output1);
    coverOldPoint(c.output2);

    // Daarna DAA + nieuwe groene verbindingen tekenen.
    drawDAA(c);
    drawDaaConnectionHandle(c);
  }

  // Aftakkers tekenen.
  for(const a of aftakkers){
    if(a.page===currentPage){
      drawAftakker(a);
    }
  }

  // Verdelers tekenen.
  for(const d of dividers){
    if(d.page===currentPage){
      drawDivider(d);
    }
  }

  // Preview van de huidige workflow.
  if(workflow.active&&workflow.daaPoint){
    drawDAAPreview(workflow);
  }

  // Laat de verdeler tijdens het kiezen van de uitgangen continu zien.
  // Ook in dividerOutputs blijft het tijdelijke symbool zichtbaar, zodat
  // de gebruiker direct ziet waar de uitgangen op aansluiten.
  if(workflow.active&&(workflow.mode==="divider1"||workflow.mode==="divider2"||workflow.mode==="dividerOutputs")){
    drawDividerPreview(workflow);
  }

  if(workflow.active&&(workflow.mode==="aftakkerPlace"||workflow.mode==="aftakkerOutput1"||workflow.mode==="aftakkerOutput2")){
    drawAftakkerPreview(workflow);
  }

  // Teken de kop als LAATSTE. Afdekzones voor gum/DAA kunnen bovenaan
  // beginnen; door de kop daarna opnieuw te tekenen blijft de bestandsnaam
  // altijd zichtbaar en staat hij gegarandeerd vóór het adres.
  if(includeHeader && (fileNameText || addressText)){
    ctx.save();
    ctx.fillStyle="#fff";
    ctx.fillRect(0,0,Math.min(900,ctx.canvas.width),70);
    ctx.fillStyle="#111";
    ctx.textAlign="left";
    ctx.textBaseline="top";

    let nextX=25;
    if(fileNameText){
      // Alleen in de zichtbare PDF-kop het achtervoegsel -NODE001/-NODE... weglaten.
      // De volledige node blijft intern beschikbaar voor Excel en de koppeling.
      const displayNode=String(fileNameText).replace(/[-_\s]*NODE\s*\d+\b/i, '').trim();
      ctx.font="bold 20px Arial";
      ctx.fillText(displayNode,nextX,20);
      nextX += ctx.measureText(displayNode).width + 25;
    }
    if(addressText){
      ctx.font="20px Arial";
      const straat=addressText.split("\n")[0].trim();
      ctx.fillText(straat,nextX,20);
    }
    ctx.restore();
  }

  // Materiaaltabel uit Excel altijd als laatste tekenen, zodat geen DAA-,
  // gum- of afdeklaag de tabel bij het opslaan weer kan overschrijven.
  drawExcelMaterials();
}


function coverDaaOutputInterior(c){
  if(!c || !c.daaPoint || c.page!==currentPage)return;

  const x=c.daaPoint.x;
  const y=c.daaPoint.y;
  const bottom=y+78;
  const leftStart=x-15;
  const rightStart=x+15;

  const points=[];
  if(c.output2)points.push(c.output2);
  if(c.output1)points.push(c.output1);

  // Wis alles tussen de twee groene uitgangsroutes.
  // Ook het stuk onder de gezamenlijke horizontale rij wordt meegenomen.
  let minX=leftStart;
  let maxX=rightStart;
  let maxY=greenRowY!==null ? greenRowY : bottom;

  for(const p of points){
    if(p && p.page===currentPage){
      minX=Math.min(minX,p.x);
      maxX=Math.max(maxX,p.x);
      maxY=Math.max(maxY,p.y);
    }
  }

  // Kleine marge naar binnen zodat de groene lijnen zelf niet worden gewist.
  const left=minX+4;
  const top=bottom+2;
  const width=Math.max(0,(maxX-minX)-8);
  const height=Math.max(0,maxY-top+8);

  if(width<=0 || height<=0)return;

  ctx.save();
  ctx.fillStyle="#fff";
  ctx.fillRect(left,top,width,height);
  ctx.restore();
}

function exportCoverDaaOutputInterior(c,ctx2){
  if(!c || !c.daaPoint)return;

  const x=c.daaPoint.x;
  const y=c.daaPoint.y;
  const bottom=y+78;
  const leftStart=x-15;
  const rightStart=x+15;

  const points=[];
  if(c.output2)points.push(c.output2);
  if(c.output1)points.push(c.output1);

  let minX=leftStart;
  let maxX=rightStart;
  let maxY=greenRowY!==null ? greenRowY : bottom;

  for(const p of points){
    if(p){
      minX=Math.min(minX,p.x);
      maxX=Math.max(maxX,p.x);
      maxY=Math.max(maxY,p.y);
    }
  }

  const left=minX+4;
  const top=bottom+2;
  const width=Math.max(0,(maxX-minX)-8);
  const height=Math.max(0,maxY-top+8);

  if(width<=0 || height<=0)return;

  ctx2.save();
  ctx2.fillStyle="#fff";
  ctx2.fillRect(left,top,width,height);
  ctx2.restore();
}

function coverAutoWcaCorridors(pageNumber){
  const items=changes
    .filter(c=>c.page===pageNumber && c.autoWca && c.daaPoint)
    .slice()
    .sort((a,b)=>a.daaPoint.x-b.daaPoint.x);

  if(!items.length)return;

  for(let i=0;i<items.length;i++){
    const c=items[i];
    const x=c.daaPoint.x;
    const bottom=c.daaPoint.y+40;
    const wcaY=c.wcaPoint ? c.wcaPoint.y : (c.output1||c.output2||{}).y;
    if(!Number.isFinite(wcaY))continue;

    // Wis het volledige gebied links en rechts van de nieuwe groene lijn,
    // vanaf direct onder het DAA-symbool tot aan het WCA-aansluitpunt.
    // Bij een buur-DAA stoppen we halverwege beide lijnen, zodat de buur-
    // verbinding niet wordt geraakt. Aan de buitenzijde gebruiken we een
    // ruimere wiszone zodat ook oude teksten en gele markeringen verdwijnen.
    let left=i>0 ? (items[i-1].daaPoint.x+x)/2+4 : x-140;
    let right=i<items.length-1 ? (x+items[i+1].daaPoint.x)/2-4 : x+140;

    // Laat de uitgangsteksten direct onder de DAA staan.
    // Het oude PDF-materiaal wordt pas NA deze teksten gewist.
    // De uitgangstekst loopt ongeveer 96 px langs de verticale richting.
    const top=bottom+116;
    const height=Math.max(0,(wcaY-4)-top);
    const width=Math.max(0,right-left);
    if(width<=0 || height<=0)continue;

    ctx.save();
    ctx.fillStyle="#fff";
    ctx.fillRect(left,top,width,height);
    ctx.restore();
  }
}


function coverOldPoint(p){
  if(!p || p.page!==currentPage)return;

  // Bij automatisch WCA plaatsen wordt de oude inhoud tussen de
  // DAA's en de WCA's al door coverAutoWcaCorridors verwijderd.
  // Deze functie mag daarbij NIET meer vanaf de bovenkant wissen,
  // anders verdwijnen de bestandsnaam en straatnaam.
  if(p.autoWca || p.manual)return;

  const rowY=greenRowY!==null
    ? greenRowY
    : p.y;

  const left=p.x-60;
  const top=0;
  const width=120;
  const height=Math.max(0,rowY+16);

  ctx.save();
  ctx.fillStyle="#fff";
  ctx.fillRect(left,top,width,height);
  ctx.restore();
}



function coverDividerOldArea(d){
  if(!d || d.page!==currentPage)return;

  const rowY=d.greenY!==null ? d.greenY : d.y;
  const centerX=d.centerX;

  // Alles wat rond de verdeler BOVEN de nieuwe groene horizontale
  // lijn staat mag verdwijnen: oude tekst, oude lijnstukken en
  // oude aansluitingen. De nieuwe groene lijn zelf blijft staan.
  const width=260;
  const left=centerX-width/2;
  const top=0;
  const height=Math.max(0,rowY+16);

  ctx.save();
  ctx.fillStyle="#fff";
  ctx.fillRect(left,top,width,height);
  ctx.restore();
}

function coverDividerOldPoint(p){
  if(!p || p.page!==currentPage)return;

  const lineWidth=14;
  const left=p.x-lineWidth/2;
  const top=Math.max(0,p.y-8);
  const width=lineWidth;
  const height=16;

  ctx.save();
  ctx.fillStyle="#fff";
  ctx.fillRect(left,top,width,height);
  ctx.restore();
}

function getUsedDaaOutput(c){
  if(!c)return null;

  // Als een verdeler aan deze DAA gekoppeld is, is de uitgang
  // die naar de verdeler gaat de "gebruikte" uitgang.
  for(const d of dividers){
    if(d.daaConnection===c && d.daaOutputUsed){
      return d.daaOutputUsed;
    }
  }

  return null;
}

function getNearestDaaOutputForDivider(d,c){
  if(!d || !c || !c.daaPoint)return null;

  const left={x:c.daaPoint.x-15,y:c.daaPoint.y+78,output:2};
  const right={x:c.daaPoint.x+15,y:c.daaPoint.y+78,output:1};

  // De afstand wordt gemeten naar het echte aansluitpunt bovenaan
  // in het midden van de verdeler.
  const targetX=d.centerX;
  const targetY=d.y-28;

  const distLeft=Math.hypot(targetX-left.x,targetY-left.y);
  const distRight=Math.hypot(targetX-right.x,targetY-right.y);

  return distLeft<=distRight ? left : right;
}

function drawDaaToDividerConnection(d){
  const c=d.daaConnection;
  if(!c || !c.daaPoint || greenRowY===null)return;

  const nearest=getNearestDaaOutputForDivider(
    {centerX:d.centerX, y:d.y},
    c
  );
  if(!nearest)return;

  const daaX=nearest.x;
  const outputY=nearest.y;
  const rowY=greenRowY;

  // 1 cm tussen de horizontale DAA-rij en de bovenkant van de verdeler.
  const halfCm=(72/2.54)*scale/2;
  const dividerTop=rowY+halfCm;

  ctx.save();
  const lineColor=d.existingDivider ? "#000" : "#18c900";
  ctx.strokeStyle=lineColor;
  ctx.lineWidth=4;
  ctx.lineCap="round";
  ctx.lineJoin="round";

  ctx.beginPath();
  ctx.moveTo(daaX,outputY);
  ctx.lineTo(daaX,rowY);
  ctx.lineTo(d.centerX,rowY);
  ctx.lineTo(d.centerX,dividerTop);
  ctx.stroke();

  ctx.restore();
}

function drawDividerOutputCircle(point,color="#18c900"){
  if(!point)return;
  ctx.save();
  ctx.fillStyle="#fff";
  ctx.strokeStyle=color;
  ctx.lineWidth=3;
  ctx.beginPath();
  ctx.arc(point.x,point.y,7,0,Math.PI*2);
  ctx.fill();
  ctx.stroke();
  ctx.restore();
}

function getGreenSegmentsForDividerJump(excludeDivider){
  const segs=[];
  const add=(x1,y1,x2,y2)=>{
    if(![x1,y1,x2,y2].every(Number.isFinite)) return;
    if(Math.hypot(x2-x1,y2-y1)<2) return;
    segs.push({x1,y1,x2,y2});
  };

  // DAA-verbindingen zijn groene lijnen waar een verdeler overheen kan lopen.
  for(const c of changes){
    if(c.page!==currentPage || !c.daaPoint) continue;
    const x=c.daaPoint.x;
    const bottom=c.daaPoint.y+20;
    const row=greenRowY!==null ? greenRowY : c.daaPoint.y+60*scale;
    if(c.autoWca){
      const p=c.output1||c.output2;
      const lineX=c.wcaPoint ? c.wcaPoint.x : (p?p.x:x);
      const endY=c.wcaPoint ? c.wcaPoint.y : (p?p.y:row);
      add(x,bottom,x,row);
      add(x,row,lineX,row);
      add(lineX,row,lineX,endY);
    }else{
      if(c.output2){
        const sx=x-3.75;
        add(sx,bottom+1,sx,row); add(sx,row,c.output2.x,row); add(c.output2.x,row,c.output2.x,c.output2.y);
      }
      if(c.output1){
        const sx=x+3.75;
        add(sx,bottom+1,sx,row); add(sx,row,c.output1.x,row); add(c.output1.x,row,c.output1.x,c.output1.y);
      }
    }
  }

  // Aftakkers.
  for(const a of aftakkers){
    if(a.page!==currentPage) continue;
    const x=a.centerX, y=a.y, bottom=y+20;
    if(a.anchorY!=null) add(x,bottom,x,a.anchorY);
    for(const output of [a.output1,a.output2]){
      if(output){
        const sx=output===a.output1 ? x+9.5 : x-9.5, sy=y+8;
        add(sx,sy,output.x,sy); add(output.x,sy,output.x,output.y);
      }
    }
  }

  // Andere verdeleruitgangen.
  for(const d of dividers){
    if(d.page!==currentPage || d===excludeDivider) continue;
    const x=d.centerX, y=d.y, outY=y+8;
    if(d.anchorY!=null) add(x,y+20,x,d.anchorY);
    const dCount=d.dividerCount || (d.output4 ? 4 : d.output3 ? 3 : 2);
    const dOutputs=d.outputs || [d.output1,d.output2,d.output3,d.output4].slice(0,dCount);
    for(let oi=0;oi<dCount;oi++){
      const p=dOutputs[oi];
      if(!p) continue;
      const sp=getDividerOutputStart(d,oi+1);
      add(sp.x,sp.y,p.x,sp.y);
      add(p.x,sp.y,p.x,p.y);
    }
  }
  return segs;
}

function segmentIntersectionsForOrthogonalRoute(ax,ay,bx,by,segs){
  const hits=[];
  const horizontal=Math.abs(by-ay)<0.5;
  const vertical=Math.abs(bx-ax)<0.5;
  if(!horizontal && !vertical) return hits;
  for(const s of segs){
    const sh=Math.abs(s.y2-s.y1)<0.5;
    const sv=Math.abs(s.x2-s.x1)<0.5;
    if(horizontal && sv){
      const cx=s.x1, cy=ay;
      if(cx>=Math.min(ax,bx)+12 && cx<=Math.max(ax,bx)-12 && cy>=Math.min(s.y1,s.y2)-1 && cy<=Math.max(s.y1,s.y2)+1){
        hits.push({t:(cx-ax)/(bx-ax),x:cx,y:cy,dir:'h'});
      }
    }else if(vertical && sh){
      const cx=ax, cy=s.y1;
      if(cy>=Math.min(ay,by)+12 && cy<=Math.max(ay,by)-12 && cx>=Math.min(s.x1,s.x2)-1 && cx<=Math.max(s.x1,s.x2)+1){
        hits.push({t:(cy-ay)/(by-ay),x:cx,y:cy,dir:'v'});
      }
    }
  }
  hits.sort((a,b)=>a.t-b.t);
  const out=[];
  for(const h of hits){
    if(!out.length || Math.hypot(h.x-out[out.length-1].x,h.y-out[out.length-1].y)>8) out.push(h);
  }
  return out;
}

function strokeOrthogonalSegmentWithJumps(ax,ay,bx,by,segs){
  const hits=segmentIntersectionsForOrthogonalRoute(ax,ay,bx,by,segs);
  const horizontal=Math.abs(by-ay)<0.5;
  const r=7;
  ctx.moveTo(ax,ay);
  if(!hits.length){ ctx.lineTo(bx,by); return; }
  let cx=ax, cy=ay;
  for(const h of hits){
    if(horizontal){
      const dir=Math.sign(bx-ax)||1;
      const a=h.x-dir*r, b=h.x+dir*r;
      ctx.lineTo(a,ay);
      ctx.quadraticCurveTo(h.x,ay-7,b,ay);
      cx=b;
    }else{
      const dir=Math.sign(by-ay)||1;
      const a=h.y-dir*r, b=h.y+dir*r;
      ctx.lineTo(ax,a);
      ctx.quadraticCurveTo(ax+7,h.y,ax,b);
      cy=b;
    }
  }
  ctx.lineTo(bx,by);
}

function getDividerOutputStart(d, output){
  const x=d.centerX, y=d.y;
  const outY=y+8;
  const count=d.dividerCount || (d.output4 ? 4 : d.output3 ? 3 : 2);
  if(d.existingStyle==="aftakker"){
    return output===1 ? {x:x-9.5,y:y+1} : {x:x,y:y+8};
  }
  if(d.existingStyle==="een") return {x:x,y:y+8};
  if(count===3){
    if(output===1) return {x:x-9.5,y:outY};
    if(output===2) return {x:x,y:y+12};
    return {x:x+9.5,y:outY};
  }
  if(count===4){
    if(output===1) return {x:x-9.5,y:outY};
    if(output===2) return {x:x-5.5,y:y+8};
    if(output===3) return {x:x+5.5,y:y+8};
    return {x:x+9.5,y:outY};
  }
  return output===1 ? {x:x+9.5,y:outY} : {x:x-9.5,y:outY};
}

function drawDividerOutputDown(x1,y1,x2,y2,color,ownerDivider=null,outputNumber=null){
  // De buitenste uitgangen blijven horizontaal + omlaag werken.
  // De extra uitgangen van 3/4 verdelers hebben een eigen vaste richting:
  // 3 verdeler uitgang 2 = recht naar beneden,
  // 4 verdeler uitgang 2/3 = schuin naar beneden.
  const count=ownerDivider?.dividerCount || 2;

  ctx.save();
  ctx.strokeStyle=color;
  ctx.lineWidth=4;
  ctx.lineCap="round";
  ctx.lineJoin="round";
  ctx.beginPath();

  if(ownerDivider?.existingStyle==="een"){
    const segs=getGreenSegmentsForDividerJump(ownerDivider);
    strokeOrthogonalSegmentWithJumps(x1,y1,x2,y2,segs);
  }else if(ownerDivider?.existingStyle==="aftakker" && outputNumber===2){
    const segs=getGreenSegmentsForDividerJump(ownerDivider);
    strokeOrthogonalSegmentWithJumps(x1,y1,x1,y2,segs);
    strokeOrthogonalSegmentWithJumps(x1,y2,x2,y2,segs);
  }else if(count===3 && outputNumber===2){
    // Midden-uitgang van de 3 verdeler: eerst echt verticaal omlaag.
    const segs=getGreenSegmentsForDividerJump(ownerDivider);
    strokeOrthogonalSegmentWithJumps(x1,y1,x1,y2,segs);
    strokeOrthogonalSegmentWithJumps(x1,y2,x2,y2,segs);
  }else if(count===4 && (outputNumber===2 || outputNumber===3)){
    // Binnenste uitgangen van de 4 verdeler:
    // eerst recht omlaag, dan horizontaal naar links/rechts,
    // daarna recht omlaag naar het aansluitpunt.
    const bendY=y1+28;
    ctx.moveTo(x1,y1);
    ctx.lineTo(x1,bendY);
    ctx.lineTo(x2,bendY);
    ctx.lineTo(x2,y2);
  }else{
    // Bestaande werking van de buitenste uitgangen behouden:
    // eerst horizontaal, daarna verticaal omlaag.
    const segs=getGreenSegmentsForDividerJump(ownerDivider);
    strokeOrthogonalSegmentWithJumps(x1,y1,x2,y1,segs);
    strokeOrthogonalSegmentWithJumps(x2,y1,x2,y2,segs);
  }

  ctx.stroke();
  ctx.restore();
}

function drawAftakker(a){
  if(!a || a.page!==currentPage)return;
  const x=a.centerX, y=a.y, bottom=y+20;
  ctx.save();
  ctx.fillStyle="#ff0000"; ctx.strokeStyle="#ff0000";
  ctx.font="9px Arial"; ctx.textAlign="left"; ctx.textBaseline="middle";
  ctx.lineWidth=.75; ctx.lineJoin="round"; ctx.lineCap="round";
  ctx.fillText(a.type || `OT-VCX-${a.db||"??"}/PTP`,x+13,y-2);
  ctx.beginPath();
  ctx.moveTo(x-9.5,y-12); ctx.lineTo(x+9.5,y-12); ctx.lineTo(x+9.5,y+1);
  ctx.lineTo(x+5.5,y+8); ctx.lineTo(x-5.5,y+8); ctx.lineTo(x-9.5,y+1); ctx.closePath(); ctx.stroke();
  if(a.anchorY!=null){
    ctx.strokeStyle="#18c900"; ctx.lineWidth=4; ctx.lineCap="butt";
    ctx.beginPath(); ctx.moveTo(x,bottom); ctx.lineTo(x,a.anchorY); ctx.stroke();
  }
  // Uitgang 1 komt altijd vanaf de linkerzijde van de aftakker.
  if(a.output1){
    const startX=x-9.5, startY=y+1;
    const p=a.output1;
    ctx.strokeStyle="#18c900"; ctx.lineWidth=4; ctx.lineCap="round"; ctx.lineJoin="miter";
    ctx.beginPath();
    ctx.moveTo(startX,startY);
    ctx.lineTo(p.x,startY);
    ctx.lineTo(p.x,p.y);
    ctx.stroke();
    drawDividerOutputCircle(p);
  }
  // Uitgang 2 komt altijd vanaf de onderkant van de aftakker.
  if(a.output2){
    const startX=x, startY=y+8;
    const p=a.output2;
    ctx.strokeStyle="#18c900"; ctx.lineWidth=4; ctx.lineCap="round"; ctx.lineJoin="miter";
    ctx.beginPath();
    ctx.moveTo(startX,startY);
    ctx.lineTo(startX,p.y);
    ctx.lineTo(p.x,p.y);
    ctx.stroke();
    drawDividerOutputCircle(p);
  }
  ctx.restore();
}

function drawDivider(d){
  // De verdeler behoudt altijd zijn eigen positie.
  // Alleen de gebruiker mag hem verplaatsen.
  const x=d.centerX;
  const y=d.y;
  const bottom=y+20;
  const lineColor=d.existingDivider ? "#000" : "#18c900";
  const symbolColor=d.existingDivider ? "#000" : "#ff0000";

  ctx.save();

  // Verdeler even klein als de DAA.
  ctx.fillStyle=symbolColor;
  ctx.font="9px Arial";
  ctx.textAlign="left";
  ctx.textBaseline="middle";
  ctx.fillText(d.type,x+13,y-2);

  ctx.strokeStyle=symbolColor;
  ctx.lineWidth=.75;
  ctx.lineJoin="round";
  ctx.lineCap="round";

  ctx.beginPath();
  ctx.moveTo(x-9.5,y-12);
  ctx.lineTo(x+9.5,y-12);
  ctx.lineTo(x+9.5,y+1);
  ctx.lineTo(x+5.5,y+8);
  ctx.lineTo(x-5.5,y+8);
  ctx.lineTo(x-9.5,y+1);
  ctx.closePath();
  ctx.stroke();

  // Groene verbinding van DAA naar de bovenkant van de verdeler.
  drawDaaToDividerConnection(d);

  // De groene lijn op de eerste klik blijft altijd zichtbaar.
  if(d.anchorY!=null){
    ctx.strokeStyle=lineColor;
    ctx.lineWidth=4;
    ctx.lineCap="butt";
    ctx.beginPath();
    ctx.moveTo(x,bottom);
    ctx.lineTo(x,d.anchorY);
    ctx.stroke();
  }

  // Uitgangen. Bij 2 verdeler blijven de bestaande uitgangen links/rechts;
  // bij 3 en 4 verdelers komen de extra aansluitpunten volgens de gekozen vorm.
  const dCount=d.dividerCount || (d.output4 ? 4 : d.output3 ? 3 : 2);
  const dOutputs=d.outputs || [d.output1,d.output2,d.output3,d.output4].slice(0,dCount);
  for(let oi=0;oi<dCount;oi++){
    const p=dOutputs[oi];
    if(!p) continue;
    const sp=getDividerOutputStart(d,oi+1);
    drawDividerOutputDown(sp.x,sp.y,p.x,p.y,lineColor,d,oi+1);
    drawDividerOutputCircle(p,lineColor);
  }

  ctx.restore();
}
function drawDaaToDividerPreview(w){
  if(!w.dividerPoint || !w.daaConnection || !w.daaConnection.daaPoint)return;
  if(greenRowY===null)return;

  const nearest=getNearestDaaOutputForDivider(
    {centerX:w.dividerPoint.x, y:w.dividerPoint.y},
    w.daaConnection
  );
  if(!nearest)return;

  const daaX=nearest.x;
  const outputY=nearest.y;
  const rowY=greenRowY;
  const halfCm=(72/2.54)*scale/2;
  const dividerTop=rowY+halfCm;

  ctx.save();
  ctx.globalAlpha=.65;
  ctx.strokeStyle=w.existingDivider ? "#000" : "#18c900";
  ctx.lineWidth=4;
  ctx.lineCap="round";
  ctx.lineJoin="round";

  ctx.beginPath();
  ctx.moveTo(daaX,outputY);
  ctx.lineTo(daaX,rowY);
  ctx.lineTo(w.dividerPoint.x,rowY);
  ctx.lineTo(w.dividerPoint.x,dividerTop);
  ctx.stroke();

  ctx.restore();
}

function drawAftakkerPreview(w){
  if(!w.aftakkerPoint)return;
  const x=w.aftakkerPoint.x, y=w.aftakkerPoint.y-20, bottom=y+20;
  ctx.save(); ctx.globalAlpha=.65;
  ctx.fillStyle="#ff0000"; ctx.strokeStyle="#ff0000"; ctx.font="9px Arial";
  ctx.textAlign="left"; ctx.textBaseline="middle"; ctx.lineWidth=.75;
  ctx.beginPath();
  ctx.moveTo(x-9.5,y-12); ctx.lineTo(x+9.5,y-12); ctx.lineTo(x+9.5,y+1);
  ctx.lineTo(x+5.5,y+8); ctx.lineTo(x-5.5,y+8); ctx.lineTo(x-9.5,y+1); ctx.closePath(); ctx.stroke();
  ctx.fillText(w.aftakkerType || `OT-VCX-${w.aftakkerDb||"??"}/PTP`,x+13,y-2);
  ctx.strokeStyle="#18c900"; ctx.lineWidth=4;
  ctx.beginPath(); ctx.moveTo(x,bottom); ctx.lineTo(x,w.aftakkerPoint.y); ctx.stroke();
  // Voorbeeld: uitgang 1 links.
  if(w.aftakkerOutput1){
    const p=w.aftakkerOutput1;
    const startX=x-9.5, startY=y+1;
    ctx.strokeStyle="#18c900"; ctx.lineWidth=4; ctx.lineCap="round"; ctx.lineJoin="miter";
    ctx.beginPath();
    ctx.moveTo(startX,startY);
    ctx.lineTo(p.x,startY);
    ctx.lineTo(p.x,p.y);
    ctx.stroke();
    drawDividerOutputCircle(p);
  }
  // Voorbeeld: uitgang 2 vanaf de onderkant.
  if(w.aftakkerOutput2){
    const p=w.aftakkerOutput2;
    const startX=x, startY=y+8;
    ctx.strokeStyle="#18c900"; ctx.lineWidth=4; ctx.lineCap="round"; ctx.lineJoin="miter";
    ctx.beginPath();
    ctx.moveTo(startX,startY);
    ctx.lineTo(startX,p.y);
    ctx.lineTo(p.x,p.y);
    ctx.stroke();
    drawDividerOutputCircle(p);
  }
  ctx.restore();
}

function drawDividerPreview(w){
  if(!w.dividerPoint)return;

  const x=w.dividerPoint.x;
  const anchorY=w.dividerPoint.anchorY!=null ? w.dividerPoint.anchorY : w.dividerPoint.y;
  const y=anchorY-20;
  const outY=y+8;

  ctx.save();
  ctx.globalAlpha=.65;
  const lineColor=w.existingDivider ? "#000" : "#18c900";
  const symbolColor=w.existingDivider ? "#000" : "#ff0000";

  // Kleine verdeler, gelijk aan de DAA-grootte.
  ctx.strokeStyle=symbolColor;
  ctx.lineWidth=.75;
  ctx.lineJoin="round";
  ctx.lineCap="round";
  ctx.beginPath();
  ctx.moveTo(x-9.5,y-12);
  ctx.lineTo(x+9.5,y-12);
  ctx.lineTo(x+9.5,y+1);
  ctx.lineTo(x+5.5,y+8);
  ctx.lineTo(x-5.5,y+8);
  ctx.lineTo(x-9.5,y+1);
  ctx.closePath();
  ctx.stroke();

  ctx.fillStyle=symbolColor;
  ctx.font="9px Arial";
  ctx.textAlign="left";
  ctx.textBaseline="middle";
  ctx.fillText(w.divider,x+13,y-2);

  // De eerste klik is een vast ankerpunt: deze groene lijn blijft staan.
  ctx.strokeStyle=lineColor;
  ctx.lineWidth=4;
  ctx.lineCap="butt";
  ctx.beginPath();
  ctx.moveTo(x,outY);
  ctx.lineTo(x,anchorY);
  ctx.stroke();

  // Verbinding met DAA alleen als er al een DAA beschikbaar is.
  drawDaaToDividerPreview({...w,dividerPoint:{...w.dividerPoint,y:y}});

  const previewPoints=w.outputPoints || [w.divider1Point,w.divider2Point].filter(Boolean);
  const previewCount=w.dividerCount || 2;
  const previewD={centerX:x,y:y,dividerCount:previewCount};
  previewPoints.forEach((p,i)=>{
    const sp=getDividerOutputStart(previewD,i+1);
    drawDividerOutputDown(sp.x,sp.y,p.x,p.y,lineColor,{dividerCount:previewCount,existingDivider:!!w.existingDivider},i+1);
    drawDividerOutputCircle(p,lineColor);
  });

  ctx.restore();
}
function drawDAAPreview(w){
  const c={
    page:w.daaPoint.page,
    daa:w.daa,
    daaPoint:w.daaPoint,
    output1:w.output1Point,
    output2:w.output2Point
  };

  drawDAA(c,true);
}

function drawDAA(c,preview=false){
  if(c.page!==currentPage)return;

  const x=c.daaPoint.x;
  const y=c.daaPoint.y;
  const red="#ff0000";
  const green="#18c900";

  ctx.save();
  ctx.globalAlpha=preview?0.65:1;

  // DAA is nu 1/4 van de oorspronkelijke grootte.
  ctx.fillStyle="#111";
  ctx.font="9px Arial";
  ctx.textAlign="center";
  ctx.textBaseline="alphabetic";
  ctx.fillText(c.daa,x,y-28);

  const boxW=19;
  const boxH=40;
  const left=x-boxW/2;
  const top=y-20;

  ctx.strokeStyle=red;
  ctx.lineWidth=.75;
  ctx.strokeRect(left,top,boxW,boxH);

  ctx.beginPath();
  ctx.arc(x,top+11,7.5,0,Math.PI*2);
  ctx.stroke();

  drawArrowSymbol(x-3.25,top+6,8.75);
  drawArrowSymbol(x+2,top+6,8.75);

  ctx.fillStyle=red;
  ctx.fillRect(left-2.5,top+17,2.5,2.25);

  ctx.beginPath();
  ctx.moveTo(x-6.75,top+22.5);
  ctx.lineTo(x+6.75,top+22.5);
  ctx.lineTo(x,top+34.75);
  ctx.closePath();
  ctx.stroke();

  const bottom=top+boxH;
  const leftStartX=x-3.75;
  const rightStartX=x+3.75;

  // Bij automatisch WCA plaatsen loopt de groene lijn vanuit het
  // midden onder de DAA over de bestaande verticale lijn tot aan de
  // WCA-tekst. Onder het WCA-punt blijft de oude lijn staan.
  if(c.autoWca){
    const p=c.output1 || c.output2;
    if(p){
      ctx.save();
      ctx.strokeStyle=green;
      ctx.lineWidth=4;
      ctx.lineCap="butt";
      ctx.lineJoin="round";

      const aansluitingY=c.wcaPoint ? c.wcaPoint.y : p.y;
      const lineX=c.wcaPoint ? c.wcaPoint.x : p.x;
      // Alle DAA's gebruiken exact dezelfde horizontale aansluitrij.
      const rowY=greenRowY!==null ? greenRowY : (c.daaPoint.y+60*scale);

      ctx.beginPath();
      ctx.moveTo(x,bottom);
      ctx.lineTo(x,rowY);
      if(lineX!==x){
        ctx.lineTo(lineX,rowY);
      }
      ctx.lineTo(lineX,aansluitingY);
      ctx.stroke();
      ctx.restore();
    }
  }else{
    const usedOutput=getUsedDaaOutput(c);

    if(c.output2 && usedOutput!==2){
      drawGreenConnection(leftStartX,bottom+1,c.output2.x,c.output2.y,green,greenRowY,24*scale);
    }
    if(c.output1 && usedOutput!==1){
      drawGreenConnection(rightStartX,bottom+1,c.output1.x,c.output1.y,green,greenRowY,24*scale);
    }
  }

  // Kleine uitgangsteksten.
  ctx.fillStyle="#111";
  ctx.font="9px Arial";
  ctx.textAlign="center";

  if(c.output2){
    drawVerticalTextSmall(`${c.daa}-02`,leftStartX-13,bottom+30);
  }
  if(c.output1){
    drawVerticalTextSmall(`${c.daa}-01`,rightStartX+13,bottom+30);
  }

  ctx.restore();
}

function drawVerticalTextSmall(text,x,y){
  ctx.save();
  ctx.translate(x,y);
  ctx.rotate(-Math.PI/2);
  // Geen witte achtergrond: de groene verbinding moet volledig doorlopen.
  ctx.fillStyle="#111";
  ctx.font="9px Arial";
  ctx.textAlign="center";
  ctx.textBaseline="alphabetic";
  ctx.fillText(text,0,0);
  ctx.restore();
}

function drawGreenConnection(x1,y1,x2,y2,color,sharedY,verticalLengthPx=null){
  // Voor DAA-uitgangen gebruiken we een lijn die ongeveer even lang is
  // als de verticale uitgangstekst. Voor overige verbindingen blijft
  // de bestaande lengte van 1 cm behouden.
  const rowY=y1+(verticalLengthPx!==null
    ? verticalLengthPx
    : ((72/2.54)*scale));
  ctx.save();
  ctx.strokeStyle=color; ctx.lineWidth=4; ctx.lineCap="round"; ctx.lineJoin="round";
  ctx.beginPath();
  ctx.moveTo(x1,y1); ctx.lineTo(x1,rowY); ctx.lineTo(x2,rowY); ctx.lineTo(x2,y2);
  ctx.stroke(); ctx.restore();
}

function drawVerticalText(text,x,y){
  ctx.save();
  ctx.translate(x,y);
  ctx.rotate(-Math.PI/2);

  // Witte achtergrond achter de naam, zodat de groene lijn er niet
  // doorheen loopt en de uitgangsnaam goed leesbaar blijft.
  ctx.fillStyle="#fff";
  ctx.fillRect(-13,-48,26,96);

  ctx.fillStyle="#111";
  ctx.textAlign="center";
  ctx.textBaseline="alphabetic";
  ctx.fillText(text,0,0);
  ctx.restore();
}

function drawArrowSymbol(x,y,length){
  ctx.beginPath();
  ctx.moveTo(x,y);
  ctx.lineTo(x+length*.55,y+length*.55);
  ctx.lineTo(x+length*.35,y+length*.55);
  ctx.lineTo(x+length,y+length);
  ctx.stroke();

  ctx.beginPath();
  ctx.moveTo(x+length*.55,y+length*.55);
  ctx.lineTo(x+length*.35,y+length*.55);
  ctx.stroke();
}

/* ---------------- BEDIENING ---------------- */

prevBtn.addEventListener("click",async()=>{
  hoveredDaa=null;
  dragDaa=null;
  dragDivider=null;
  if(currentPage<=1)return;
  currentPage--;
  workflow={
    active:false,daa:null,daaPoint:null,
    output1Point:null,output2Point:null,mode:null
  };
  await renderPage();
});

nextBtn.addEventListener("click",async()=>{
  hoveredDaa=null;
  dragDaa=null;
  dragDivider=null;
  if(!pdf||currentPage>=pdf.numPages)return;
  currentPage++;
  workflow={
    active:false,daa:null,daaPoint:null,
    output1Point:null,output2Point:null,mode:null
  };
  await renderPage();
});

deleteDaaBtn.addEventListener("click",()=>{
  // DAA wissen werkt onafhankelijk van hover/selectie.
  // Toon alle DAA's die daadwerkelijk op de huidige pagina staan.
  const available=[...new Set(changes
    .filter(c=>c.page===currentPage && c.daaPoint && /^DAA\d{3}$/.test(String(c.daa||"")))
    .map(c=>String(c.daa).toUpperCase())
  )].sort((a,b)=>a.localeCompare(b,undefined,{numeric:true}));

  if(!available.length){
    alert("Er staat geen DAA op deze pagina om te wissen.");
    return;
  }

  const choice=prompt(
    `Welke DAA wil je wissen?\n\nBeschikbaar: ${available.join(", ")}`
  );
  if(choice===null)return;

  const daa=choice.trim().toUpperCase();
  if(!available.includes(daa)){
    alert(`DAA ${daa || "(leeg)"} staat niet op deze pagina.`);
    return;
  }

  if(!confirm(`Wil je ${daa} volledig wissen?`))return;

  // Verwijder alle records van deze DAA. Dit is bewust niet alleen de
  // eerste match, zodat ook een DAA met uitgang 1 + uitgang 2 volledig weg is.
  const removed=changes.filter(c=>
    c.page===currentPage && String(c.daa||"").toUpperCase()===daa
  );
  changes=changes.filter(c=>
    !(c.page===currentPage && String(c.daa||"").toUpperCase()===daa)
  );
  componentHistory=componentHistory.filter(h=>!(h.type==="daa" &&
    h.item && String(h.item.daa||"").toUpperCase()===daa && h.item.page===currentPage));

  // Verwijder koppelingen van verdelers naar deze DAA.
  for(const d of dividers){
    if(d.page!==currentPage)continue;
    if(removed.includes(d.daaConnection) ||
       (d.daaConnection && String(d.daaConnection.daa||"").toUpperCase()===daa)){
      d.daaConnection=null;
      d.daaOutputUsed=null;
    }
  }

  if(sequentialDaaNumber===daa){
    sequentialDaaNumber=null;
    sequentialDaaStage=null;
    workflow={active:false,daa:null,daaPoint:null,output1Point:null,output2Point:null,mode:null};
    daaQuestion.style.display="none";
  }

  hoveredDaa=null;
  status.textContent=`${daa} is gewist.`;
  stepEl.textContent="DAA verwijderd. Je kunt een nieuwe DAA plaatsen.";
  drawOverlay();
  updateButtons();
});

function undoLastComponent(){
  // Backspace werkt alleen buiten invoervelden en niet midden in een
  // plaatsingsworkflow. Zo wordt nooit per ongeluk de browserpagina
  // teruggestuurd en blijft een lopende plaatsing intact.
  if(workflow.active) return false;
  if(!componentHistory.length) return false;

  const last=componentHistory.pop();
  if(last.type==="daa") {
    const item=last.item;
    changes=changes.filter(c=>c!==item);

    // Verwijder eventuele verdelerkoppeling naar deze DAA.
    for(const d of dividers){
      if(d.daaConnection===item ||
         (d.daaConnection && item && d.daaConnection.daa===item.daa)){
        d.daaConnection=null;
        d.daaOutputUsed=null;
      }
    }

    if(sequentialDaaNumber===item.daa){
      sequentialDaaNumber=null;
      sequentialDaaStage=null;
      daaQuestion.style.display="none";
    }
  } else if(last.type==="divider") {
    dividers=dividers.filter(d=>d!==last.item);
  } else if(last.type==="aftakker") {
    aftakkers=aftakkers.filter(a=>a!==last.item);
  }

  // Als er geen DAA's meer zijn, mogen de vaste DAA-aansluitrijen opnieuw
  // worden bepaald bij de volgende plaatsing. Bij overgebleven DAA's blijft
  // de bestaande vaste hoogte behouden.
  if(!changes.length){
    daaRowY=null;
    greenRowY=null;
  }
  if(!dividers.length){
    dividerRowY=null;
    dividerGreenRowY=null;
  }

  hoveredDaa=null;
  status.textContent=`Laatste component (${last.type==="daa" ? (last.item.daa||"DAA") : last.type==="divider" ? (last.item.type||"verdeler") : "aftakker"}) is gewist.`;
  stepEl.textContent="Laatste component verwijderd. Druk nogmaals op Backspace om het vorige component te verwijderen.";
  drawOverlay();
  updateButtons();
  return true;
}

document.addEventListener("keydown",e=>{
  if(e.key!=="Backspace") return;

  const target=e.target;
  const tag=target && target.tagName ? target.tagName.toLowerCase() : "";
  if(tag==="input" || tag==="textarea" || tag==="select" || target?.isContentEditable) return;
  if(!pdf || workflow.active) return;

  if(undoLastComponent()){
    e.preventDefault();
    e.stopPropagation();
  }
});

clearBtn.addEventListener("click",()=>{
  // Ook alle gumstreken wissen: de oorspronkelijke PDF wordt daardoor
  // weer volledig zichtbaar op de plekken die met de gum waren weggehaald.
  eraserStrokes.length=0;
  eraserDrag=null;
  hoveredDaa=null;
  dragDaa=null;
  dragDivider=null;
  dragAftakker=null;
  dragAftakkerOutput=null;
  changes=[];
  dividers=[];
  aftakkers=[];
  componentHistory=[];
  daaRowY=null;
  greenRowY=null;
  dividerRowY=null;
  dividerGreenRowY=null;

  workflow={
    active:false,daa:null,daaPoint:null,
    output1Point:null,output2Point:null,mode:null
  };

  status.textContent="Alle wijzigingen zijn gewist.";
  stepEl.textContent="Stap 2: kies een DAA.";

  drawOverlay();
  updateButtons();
});
</script>
</body>
</html>


<script>
document.addEventListener("DOMContentLoaded",()=>{
  const changeBtn=document.getElementById("changeSaveDirectoryBtn");
  if(changeBtn){
    changeBtn.addEventListener("click",async()=>{
      const handle=await chooseSaveDirectory();
      if(handle){
        if(typeof status!=="undefined") status.textContent="Nieuwe opslagmap gekozen.";
        if(typeof stepEl!=="undefined") stepEl.textContent="Deze map wordt gebruikt voor volgende PDF's.";
      }
    });
  }
});
</script>
