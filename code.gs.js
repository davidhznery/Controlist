/************** CONFIG PRINCIPAL **************/
const SOS_PARENT_FOLDER_ID = "1QluhICfJkNjc6HBMpQbcZEN5dq3zIeW-";
const DLE_PARENT_FOLDER_ID = "1fc_HIWAJnX9pGTLcDMArp8r2uoC6RjCc";

const DRIVE_LINK_HEADER = "Drive Folder";

// (Opcional: ya no usamos Picker, pero puedes dejar esto)
const PICKER_CONFIG = {
  DEVELOPER_KEY: "AIzaSyAblQebzjcl55IH3B6xl7-AmSfU-8C_Aig",
  OAUTH_CLIENT_ID: "401635379104-1lge04bsij3blp9tcbg0c0tivs2q7m2n.apps.googleusercontent.com",
  PICKER_UPLOAD_FOLDER_ID: "1KNZ70nGI6z-Dw3Bmv9s-ixsko9vttl0q"
};

// Column map (1-based)
const COLMAP = {
  SOS: { SHEET_NAME: "SOS", PROJ: 1, NO: 2, CLIENT: 5, CLIENT_REF: 6, RECEIVED: 7, SUBJECT: 9, DEADLINE: 10, STATUS: 12, LEADER: 13 },
  DLE: { SHEET_NAME: "DLE", PROJ: 1, NO: 2, DIVISION: 3, CLIENT: 5, CLIENT_REF: 6, RECEIVED: 7, SUBJECT: 9, DEADLINE: 10, STATUS: 12, LEADER: 13 }
};

// Estructura de subcarpetas
const SUBFOLDERS = {
  SOS: [
    "1- Documents_&_Data_From_Client",
    "2- RFQ",
    "3- Supplier",
    "4- Estimation",
    "5- Contracts Agreements",
    "6- Native files",
    "7- Proposal",
    "8- Sent_to_the_Client"
  ],
  DLE: {
    ChampionX: [
      "1- Documents_&_Data_From_Client",
      "2- RFQ",
      "3- Supplier",
      "4- Estimation",
      "5- Contracts Agreements",
      "6- Native files",
      "7- Proposal",
      "8- Sent_to_the_Client"
    ],
    Procurement: [
      "1- Documents_&_Data_From_Client",
      "2- RFQ",
      "3- Supplier",
      "4- Estimation",
      "5- Contracts Agreements",
      "6- Native files",
      "7- Proposal",
      "8- Sent_to_the_Client"
    ],
    Trading: [
      "1- Documents_&_Data_From_Client",
      "2- RFQ",
      "3 - Suppliers",
      "4- Estimation",
      "5 Contract Agreement",
      "6 Native Files",
      "7 Proposals",
      "8- Sent_to_the_Client"
    ]
  }
};

/************** PLANTILLAS **************/
// Placeholders permitidos en "name": {{NO}}, {{SUBJECT}}, {{CLIENT}}

// === SOS ===
const SOS_TEMPLATE_FILES = {
  "2- RFQ": [
    { fileId: "1J72WbQTYx9E8MuwpEoiKmnMd-SpKZC7L", name: "RFQ SOS {{NO}} {{SUBJECT}}.docx" },
    { fileId: "1j493mcVrBGe5A154rMCkcJCLwBNxqG7X", name: "RFQ HMJ {{NO}} {{SUBJECT}}.docx" }
  ],
  "4- Estimation": [
    // ID corregido que me pasaste
    { fileId: "1_WuLIhvFCj5x4PkGwprYMQvqgL_nsRGD", name: "SOS {{NO}} Estimation cost.xlsx" }
    // Si quieres copiar una subcarpeta con imágenes, vuelve a activar:
    // ,{ folderId: "1i7mS6cskeo5j8yJWz_GzmHQjIWHCh8uJ", dest: "4- Estimation/Dimensions" }
  ],
  "7- Proposal": [
    { fileId: "16vdR3OvluNpaRJ4pCNwHSjUhD6QDdcWn", name: "SOS {{NO}} Technical Proposal.docx" },
    { fileId: "1PpAz5-NZ3BE9qyaWSlwy9h23hV1MkchV", name: "SOS {{NO}} Commercial Proposal.docx" }
  ]
};

// === DLE por división ===
const DLE_TEMPLATE_FILES = {
  ChampionX: {
    "2- RFQ": [
      { fileId: "1TEXpSwCzEG4qm73AjUfhFZFAL3ktLCrH", name: "RFQ DLE {{NO}} {{SUBJECT}}.docx" }
    ],
    "4- Estimation": [
      { fileId: "1BBTjKAJPZjH-71eI7EK3nShFPuhkO0eJ", name: "DLE {{NO}} Estimation cost.xlsx" }
    ],
    "6- Native files": [
      { folderId: "19dTU1cMPypdR97p1KbZzCZDrQUAGj5Q8", dest: "6- Native files" }
    ],
    "7- Proposal": [
      { fileId: "1usnsK2tnVPU7u0N6mLQ-jzI_nEdimIWt", name: "DLE {{NO}} Commercial Proposal.docx" },
      { fileId: "1527BI2I9hK756wywZdR28QKTKaLgDwlp", name: "DLE {{NO}} Technical Proposal.docx" }
    ]
  },
  Procurement: {
    "2- RFQ": [
      { fileId: "1WLiUhxaGnNgBjUPIKDrZrIkepxInD8Yn", name: "RFQ DLE {{NO}} {{SUBJECT}}.docx" }
    ],
    "4- Estimation": [
      { fileId: "12YozIWDtuOITqunrzJFkCE9JXFwD-Hb3", name: "DLE {{NO}} Estimation cost.xlsx" }
    ],
    "6- Native files": [
      { folderId: "16e6f2Rf2y8W7McaS64n6PihnCtLpogtB", dest: "6- Native files" }
    ],
    "7- Proposal": [
      { fileId: "1bFJyeBjB5QOiGTFCuWQeGotxzmrpQgCZ", name: "DLE {{NO}} Commercial Proposal.docx" },
      { fileId: "1SXbnmuTRINzdW1PVs3gPi3k-sM1LnJGa", name: "DLE {{NO}} Technical Proposal.docx" }
    ]
  },
  Trading: {
    "1- Documents_&_Data_From_Client": [
      { folderId: "1UI1zAkkCeTHoo3yLBAr8VppErCeg-4y2", dest: "1- Documents_&_Data_From_Client" }
    ],
    "3 - Suppliers": [
      { fileId: "1E3eaDJNizBM5hC2xN6-ff7C8KL7Vszm8", name: "DLE KYC Supplier.docx" },
      { fileId: "1V_E6hJrirh-EtINuVs3lRfusvLRUcbt4", name: "RECAP.docx" }
    ],
    "4- Estimation": [
      { fileId: "1Eq0oYdkeZtgF9BaoJlZTkkzRkmjCwdP0", name: "DLE {{NO}} Estimation cost.xlsx" }
    ],
    "5 Contract Agreement": [
      { fileId: "1OCDHmEtqQtAlOu-phMvivSZadJ993q30", name: "RECAP Template.docx" }
    ],
    "6 Native Files": [
      { folderId: "1P6DcRKSMXp7TWz2fuYQA5LKbw45VGjI_", dest: "6 Native Files" }
    ],
    "7 Proposals": [
      { fileId: "1tkQYVKpIOTwF0lHgCni2s2CnSudnEn3W", name: "Soft Corporate Proposal.docx" }
    ]
  }
};

/************** LISTAS: CLIENTS, LEADERS, STATUSES **************/
const CLIENTS_SHEET = "Clients";
const LEADERS_SHEET = "Leaders";
const STATUSES_SHEET = "Statuses";

const CLIENTS_SEED = [
  "AGOCO","ADNOC","AKAKUS","ALBERTA","AQUANI","AWDIA","AZZ","BOFFETTI","BLUE BORDER","BREGA","CPSU",
  "CNOOC IRAQ LIMITED","COCA COLA","CONEXUS BALTIC GROUP","CPSU-MALTA","DELIMARA","DELTA FZ LLE","ELICS ENGINEERING",
  "GREENSTREAM BV","HARRY PYE","HIGH TECH","HOO","INTECH","KE GROUP","LITRO GAS","LIEBHERR","MABRUK OIL OPERATIONS",
  "MEDICON SERVICES","MELITA MARINE GROUP","MIN TRANSP MT","MOG","NWD","ORLEN LIETUVA","ORLEN POLUDNIE","PALUMBO",
  "PKN ORLEN","QATAR CO","QATAR PE","RASCO","SABRI","SIRTE","SOS-Internal","TECHNO LOGIC LTD.","TRANSPORT MALTA",
  "UN- Ethiopia","UNIPETROL","UNILEVER","WAHA","WASTESERVE MALTA","WSC","WINTERSHALL","ZOC","GTCHS","AJIBADE","MOSD",
  "DSO","ENEMALTA","DAKAR","CAMEROON","SPAIN","NETHERLANDS","EASY GAS","D&S Holding","WOC","Lome","Air-JP"
];

const LEADERS_SEED = [
  "Abd alrhman","Alaa","Franklin Acevedo","Sabri Rezk","Cesar Jurado",
  "Younis","Dajana Zeka","Duygu","David Hernandez","Ahmed"
];

const STATUSES_SEED = [
  "Pending RFQ Sending",
  "Waiting RFQ Creation",
  "Waiting for offer by supplier",
  "Commercial proposal submitted",
  "Technical Submitted",
  "Technical Query",
  "Extension requested",
  "Awared",
  "No Quote – Deadline Passed",
  "Canceled by client",
  "Restriction from supplier",
  "Waiting for offer by Libya",
  "Waiting Client Reply – Tech Query",
  "Ready to quote",
  "No proposal - Comercial reasons",
  "ON Hold - WFMD",
  "No proposal - Technical reasons"
];

/************** UTILS **************/
function sanitize(s) {
  if (!s) return "";
  return String(s).replace(/[<>:"/\\|?*]/g, "-").replace(/\s+/g, " ").trim().substring(0, 180);
}
function getOrCreateHeaderColumn_(sheet, headerName) {
  const head = sheet.getRange(1,1,1, sheet.getLastColumn()).getValues()[0];
  let idx = head.indexOf(headerName) + 1;
  if (idx === 0) { idx = sheet.getLastColumn() + 1; sheet.getRange(1, idx).setValue(headerName); }
  return idx;
}
function findOrCreateFolderByName_(parent, name) {
  const it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}
function ensureSubPath_(baseFolder, path) {
  const parts = path.split("/").map(p => p.trim()).filter(Boolean);
  let f = baseFolder;
  parts.forEach(p => { f = findOrCreateFolderByName_(f, p); });
  return f;
}
function buildFolderName_(company, rowObj) {
  const left = [company, rowObj.no].filter(Boolean).join("-");
  const right = [rowObj.client, rowObj.subject].filter(Boolean).join(" - ");
  return sanitize([left, right].filter(Boolean).join(" - "));
}
function detectCompany_(sheetName){
  const n = sheetName.trim().toUpperCase();
  if (n === (COLMAP.SOS.SHEET_NAME||"SOS").toUpperCase() || /(^|[^A-Z])SOS([^A-Z]|$)/i.test(sheetName)) return "SOS";
  if (n === (COLMAP.DLE.SHEET_NAME||"DLE").toUpperCase() || /(^|[^A-Z])DLE([^A-Z]|$)/i.test(sheetName)) return "DLE";
  throw new Error("No puedo detectar la compania. Renombra las pestañas a 'SOS' y 'DLE' o ajusta COLMAP.*.SHEET_NAME.");
}
function getParentFor_(company) {
  return DriveApp.getFolderById(company === "SOS" ? SOS_PARENT_FOLDER_ID : DLE_PARENT_FOLDER_ID);
}
function subfoldersFor_(company, division) {
  if (company === "SOS") return SUBFOLDERS.SOS;
  return SUBFOLDERS.DLE[division] || SUBFOLDERS.DLE.Procurement;
}

/************** LISTAS (helpers) **************/
function ensureListSheet_(name, seed) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.hideSheet();
  }
  if (seed && seed.length && sh.getLastRow() === 0) {
    const unique = Array.from(new Set(seed.map(s => String(s).trim()).filter(Boolean))).sort();
    sh.getRange(1,1,unique.length,1).setValues(unique.map(v => [v]));
  }
  return sh;
}
function getListValues_(name, seed) {
  const sh = ensureListSheet_(name, seed);
  const last = sh.getLastRow();
  if (last === 0) return [];
  return sh.getRange(1,1,last,1).getValues().map(r => String(r[0]).trim()).filter(Boolean);
}
function appendUniqueToList_(name, seed, value) {
  const v = String(value || "").trim();
  if (!v) return;
  const sh = ensureListSheet_(name, seed);
  const all = getListValues_(name, seed);
  if (all.map(x => x.toLowerCase()).includes(v.toLowerCase())) return;
  sh.appendRow([v]);
}
/**
 * Returns the canonical value from a list that case-insensitively matches the input.
 * If no match is found, returns the original input.
 * Enhanced to handle fuzzy matching for common variations.
 */
function getCanonicalFromList_(value, listValues) {
  const v = String(value || "").trim();
  if (!v) return v;
  const lower = v.toLowerCase();

  // First, try exact case-insensitive match
  for (var i = 0; i < listValues.length; i++) {
    var candidate = String(listValues[i] || "").trim();
    if (candidate && candidate.toLowerCase() === lower) return candidate;
  }

  // Second, try fuzzy matching for common variations
  for (var i = 0; i < listValues.length; i++) {
    var candidate = String(listValues[i] || "").trim();
    if (!candidate) continue;

    const candidateLower = candidate.toLowerCase();

    // Check if they're similar (removing spaces, punctuation)
    const normalizedV = lower.replace(/[^a-z0-9]/g, '');
    const normalizedCandidate = candidateLower.replace(/[^a-z0-9]/g, '');

    if (normalizedV === normalizedCandidate) return candidate;

    // Check for partial matches (one contains the other)
    if (normalizedV.length > 3 && normalizedCandidate.length > 3) {
      if (normalizedV.includes(normalizedCandidate) || normalizedCandidate.includes(normalizedV)) {
        return candidate;
      }
    }
  }

  return v;
}
function getClients_()  { return getListValues_(CLIENTS_SHEET,  CLIENTS_SEED); }
function getLeaders_()  { return getListValues_(LEADERS_SHEET,  LEADERS_SEED); }
function getStatuses_() { return getListValues_(STATUSES_SHEET, STATUSES_SEED); }
function addClient_(name)  { appendUniqueToList_(CLIENTS_SHEET,  CLIENTS_SEED,  name);  return getClients_(); }
function addLeader_(name)  { appendUniqueToList_(LEADERS_SHEET,  LEADERS_SEED,  name);  return getLeaders_(); }
function addStatus_(name)  { appendUniqueToList_(STATUSES_SHEET, STATUSES_SEED, name);  return getStatuses_(); }

/************** SUGERIR PRÓXIMO NÚMERO (###-YY) **************/
function suggestNextProjectNo_(company) {
  const map = COLMAP[company];
  const sh = SpreadsheetApp.getActive().getSheetByName(map.SHEET_NAME);
  if (!sh) return "";
  const yearYY = (new Date().getFullYear() % 100).toString().padStart(2, "0");
  const colNo = map.NO;
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return "001-" + yearYY;

  const vals = sh.getRange(2, colNo, lastRow-1, 1).getValues().map(r => String(r[0]||"").trim());
  let maxN = 0;
  vals.forEach(v => {
    const m = v.match(/^(\d+)-(\d{2})$/);
    if (m && m[2] === yearYY) {
      const n = parseInt(m[1], 10);
      if (!isNaN(n) && n > maxN) maxN = n;
    }
  });
  return String(maxN + 1).padStart(3,"0") + "-" + yearYY;
}
function getFormDefaults(company) {
  const comp = (company === "DLE") ? "DLE" : "SOS";
  return {
    nextNo: suggestNextProjectNo_(comp),
    clients:  getClients_(),
    leaders:  getLeaders_(),
    statuses: getStatuses_()
  };
}

/************** MOTOR DE PLANTILLAS **************/
function applyPlaceholders_(pattern, rowObj) {
  return sanitize(String(pattern)
    .replace(/{{NO}}/g, rowObj.no || "")
    .replace(/{{SUBJECT}}/g, rowObj.subject || "")
    .replace(/{{CLIENT}}/g, rowObj.client || "")
  );
}
function templateMapFor_(company, division) {
  if (company === "SOS") return SOS_TEMPLATE_FILES;
  if (company === "DLE") return DLE_TEMPLATE_FILES[division] || {};
  return {};
}
function copyFile_(fileId, newName, targetFolder, errors) {
  try {
    const file = DriveApp.getFileById(fileId);
    file.makeCopy(newName || file.getName(), targetFolder);
  } catch (e) {
    errors.push("No pude copiar fileId=" + fileId + " -> " + (newName || "") + " (" + e.message + ")");
  }
}
function copyFolderContents_(folderId, targetFolder, errors) {
  try {
    const src = DriveApp.getFolderById(folderId);
    const files = src.getFiles();
    while (files.hasNext()) {
      const f = files.next();
      try {
        f.makeCopy(f.getName(), targetFolder);
      } catch (e) {
        errors.push("No pude copiar archivo '" + f.getName() + "' (" + e.message + ")");
      }
    }
  } catch (e) {
    errors.push("No pude copiar contenidos de folderId=" + folderId + " (" + e.message + ")");
  }
}
function copyTemplates_(projectFolder, company, division, rowObj) {
  const map = templateMapFor_(company, division);
  const errors = [];
  Object.keys(map).forEach(function(subfolderName){
    const baseTarget = ensureSubPath_(projectFolder, subfolderName);
    map[subfolderName].forEach(function(t){
      if (t.folderId) {
        const dest = t.dest ? ensureSubPath_(projectFolder, t.dest) : baseTarget;
        copyFolderContents_(t.folderId, dest, errors);
      } else if (t.fileId) {
        const newName = t.name ? applyPlaceholders_(t.name, rowObj) : null;
        copyFile_(t.fileId, newName, baseTarget, errors);
      }
    });
  });
  if (errors.length) {
    SpreadsheetApp.getUi().alert("Algunas plantillas no se copiaron:\n- " + errors.join("\n- "));
  }
}

/************** CORE **************/
function createProjectStructure_(company, division, rowObj) {
  const parent = getParentFor_(company);
  const folderName = buildFolderName_(company, rowObj);
  const projectFolder = findOrCreateFolderByName_(parent, folderName);
  subfoldersFor_(company, division).forEach(n => findOrCreateFolderByName_(projectFolder, n));
  copyTemplates_(projectFolder, company, division, rowObj);
  return { url: projectFolder.getUrl(), id: projectFolder.getId() };
}

/************** SHEET HELPERS **************/
function readRowToObject_(sheet, row, company) {
  const map = COLMAP[company];
  const vals = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
  return {
    no: sanitize(vals[map.NO - 1]),
    client: sanitize(vals[map.CLIENT - 1]),
    clientRef: sanitize(vals[map.CLIENT_REF - 1]),
    received: vals[map.RECEIVED - 1],
    subject: sanitize(vals[map.SUBJECT - 1]),
    deadline: vals[map.DEADLINE - 1],
    status: sanitize(vals[map.STATUS - 1]),
    leader: sanitize(vals[map.LEADER - 1]),
    division: map.DIVISION ? sanitize(vals[map.DIVISION - 1]) : ""
  };
}
function writeLinkBack_(sheet, row, url) {
  const col = getOrCreateHeaderColumn_(sheet, DRIVE_LINK_HEADER);
  sheet.getRange(row, col).setValue(url).setNote("Auto-created by Apps Script");
}

/************** COPIAR FÓRMULA EN COLUMNA D **************/
function copyDaysFormulaDown_(sheet, row) {
  var DAYS_COL = 4; // Columna D
  if (row <= 2) return; // no hay fila superior para copiar

  var src = sheet.getRange(row - 1, DAYS_COL);
  var dst = sheet.getRange(row, DAYS_COL);
  var hasFormula = !!src.getFormula();

  if (hasFormula) {
    src.copyTo(dst, SpreadsheetApp.CopyPasteType.PASTE_FORMULA, false);
  } else {
    // Fórmula de respaldo basada en columna J (Deadline)
    var formula = '=IF(J' + row + '="","",IF(J' + row + '<TODAY(),"Expired", J' + row + ' - TODAY()))';
    dst.setFormula(formula);
  }
}

/************** MENU & ACCIONES **************/
function setupMenu() {
  SpreadsheetApp.getUi()
    .createMenu("Projects")
    .addItem("New Project…", "openNewProjectSidebar")
    .addSeparator()
    .addItem("Create folder for selected row", "createForActiveRow")
    .addItem("Create folders for all rows (SOS & DLE)", "createForAllRows")
    .addItem("Diagnose active row", "diagnoseActiveRow")
    .addItem("Validate config (IDs/Sheets)", "validateConfig")
    .addItem("Fix invalid Clients (normalize)", "fixInvalidClients")
    .addToUi();
    
    SpreadsheetApp.getUi()
    .createMenu("Reports")
    .addItem("📊 Simple KPI Dashboard…", "openSimpleDashboard")
    .addItem("🧪 Simple Dashboard Test…", "simpleQuickTest")
    .addItem("🧪 KPI Test…", "kpiQuickTest")
    .addItem("🧪 Test Awarded Projects Fix…", "testAwardedProjectsFix")
    .addItem("🔍 Debug All Awarded Projects…", "debugAllAwardedProjects")
    .addItem("🔍 Debug Dashboard Data…", "debugDashboardData")
    .addItem("🔄 Clear Cache…", "clearCacheMenu")
    .addToUi();
}
function onOpen(){ setupMenu(); }

function createForActiveRow() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getActiveSheet();
    const row = sheet.getActiveCell().getRow();
    if (row <= 1) { SpreadsheetApp.getUi().alert("Selecciona una fila de datos (no el encabezado)."); return; }
    const company = detectCompany_(sheet.getName());
    const rowObj = readRowToObject_(sheet, row, company);
    const division = company === "DLE" ? (rowObj.division || "Procurement") : "";
    if (!rowObj.subject) { SpreadsheetApp.getUi().alert("La columna Subject esta vacia."); return; }
    const { url } = createProjectStructure_(company, division, rowObj);
    writeLinkBack_(sheet, row, url);
    SpreadsheetApp.getUi().alert("Carpeta creada:\n" + url);
  } catch (e){ SpreadsheetApp.getUi().alert("Error: " + e.message); }
}

function createForAllRows() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    ["SOS","DLE"].forEach(function(company){
      const sheet = ss.getSheetByName(COLMAP[company].SHEET_NAME);
      if (!sheet) return;
      const last = sheet.getLastRow();
      const linkCol = getOrCreateHeaderColumn_(sheet, DRIVE_LINK_HEADER);
      for (let r = 2; r <= last; r++) {
        const subj = sheet.getRange(r, COLMAP[company].SUBJECT).getValue();
        if (!subj) continue;
        const existing = sheet.getRange(r, linkCol).getValue();
        if (existing) continue;
        const rowObj = readRowToObject_(sheet, r, company);
        const division = company === "DLE" ? (rowObj.division || "Procurement") : "";
        const { url } = createProjectStructure_(company, division, rowObj);
        writeLinkBack_(sheet, r, url);
        // Copiar fórmula D también cuando generas en lote
        copyDaysFormulaDown_(sheet, r);
      }
    });
    SpreadsheetApp.getUi().alert("Proceso terminado.");
  } catch(e){ SpreadsheetApp.getUi().alert("Error: " + e.message); }
}

/************** DIÁLOGO ANCHO (sin Picker) **************/
function openNewProjectSidebar() {
  const html = HtmlService.createHtmlOutputFromFile("sidebar-new-project")
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .setWidth(960)
    .setHeight(720);
  SpreadsheetApp.getUi().showModelessDialog(html, "New Project");
}

/************** DIAGNÓSTICO / VALIDACIÓN **************/
function diagnoseActiveRow(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const row = sheet.getActiveCell().getRow();
  let company;
  try { company = detectCompany_(sheet.getName()); } catch(e){ company = "??"; }
  const obj = readRowToObject_(sheet, row, company === "??" ? "SOS" : company);
  SpreadsheetApp.getUi().alert(
    "Sheet: " + sheet.getName() + "\n" +
    "Company: " + company + "\n" +
    "No: " + obj.no + "\n" +
    "Client: " + obj.client + "\n" +
    "Division: " + obj.division + "\n" +
    "Subject: " + obj.subject + "\n" +
    "Deadline: " + obj.deadline
  );
}
function validateConfig(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const issues = [];
  if (!ss.getSheetByName(COLMAP.SOS.SHEET_NAME)) issues.push("No existe la hoja: " + COLMAP.SOS.SHEET_NAME);
  if (!ss.getSheetByName(COLMAP.DLE.SHEET_NAME)) issues.push("No existe la hoja: " + COLMAP.DLE.SHEET_NAME);
  try { DriveApp.getFolderById(SOS_PARENT_FOLDER_ID); } catch(e){ issues.push("SOS_PARENT_FOLDER_ID invalido o sin permisos."); }
  try { DriveApp.getFolderById(DLE_PARENT_FOLDER_ID); } catch(e){ issues.push("DLE_PARENT_FOLDER_ID invalido o sin permisos."); }
  SpreadsheetApp.getUi().alert(issues.length ? ("Problemas:\n- " + issues.join("\n- ")) : "Config OK");
}

/************** FORM (desde HTML) → CREA PROYECTO **************/
function getTargetSheet_(company) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const name = company === "SOS" ? (COLMAP.SOS.SHEET_NAME || "SOS") : (COLMAP.DLE.SHEET_NAME || "DLE");
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    const headers = company === "SOS"
      ? ["Proj.", "No.*", "Type", "Days for Deadline", "*Client*", "Client Ref No", "Received", "Month", "*Subject:*", "Deadline", "Month", "Status", "LEADER PJ"]
      : ["Proj.", "No.*", "Division", "Days for Deadline", "*Client*", "Client Ref No", "Received", "Month", "*Subject:*", "Deadline", "Month", "Status", "LEADER PJ"];
    sh.getRange(1,1,1,headers.length).setValues([headers]);
  }
  getOrCreateHeaderColumn_(sh, DRIVE_LINK_HEADER);
  return sh;
}

function createProjectFromForm(form) {
  // form = { company, division, no, client, clientOther, subject, receivedISO, deadlineISO, leader, leaderOther, status, statusOther }
  const company = form.company === "DLE" ? "DLE" : "SOS";
  const sheet = getTargetSheet_(company);

  // Resolve values from form
  let client = (form.client === "__OTHER__" ? form.clientOther : form.client) || "";
  let leader = (form.leader === "__OTHER__" ? form.leaderOther : form.leader) || "";
  let status = (form.status === "__OTHER__" ? form.statusOther : form.status) || "";

  // Enhanced client validation and normalization
  if (client) {
    // First, try to find an exact match (case insensitive)
    const clients = getClients_();
    const canonicalClient = getCanonicalFromList_(client, clients);

    if (canonicalClient === client) {
      // Exact match found, use it
      client = canonicalClient;
    } else {
      // No exact match, add as new client and use as-is
      client = client.trim();
      addClient_(client);
    }
  }

  // Enhanced leader validation and normalization
  if (leader) {
    const leaders = getLeaders_();
    const canonicalLeader = getCanonicalFromList_(leader, leaders);

    if (canonicalLeader === leader) {
      leader = canonicalLeader;
    } else {
      leader = leader.trim();
      addLeader_(leader);
    }
  }

  // Enhanced status validation and normalization
  if (status) {
    const statuses = getStatuses_();
    const canonicalStatus = getCanonicalFromList_(status, statuses);

    if (canonicalStatus === status) {
      status = canonicalStatus;
    } else {
      status = status.trim();
      addStatus_(status);
    }
  }

  const received = form.receivedISO ? new Date(form.receivedISO) : "";
  const deadline = form.deadlineISO ? new Date(form.deadlineISO) : "";

  if (company === "SOS") {
    // A..M: Proj, No, Type, Days, Client, ClientRef, Received, Month, Subject, Deadline, Month, Status, Leader
    const row = ["SOS", form.no || "", "", "", client, "", received, "", form.subject || "", deadline, "", status, leader];
    sheet.appendRow(row);
  } else {
    // A..M: Proj, No, Division, Days, Client, ClientRef, Received, Month, Subject, Deadline, Month, Status, Leader
    const row = ["DLE", form.no || "", form.division || "Procurement", "", client, "", received, "", form.subject || "", deadline, "", status, leader];
    sheet.appendRow(row);
  }

  const r = sheet.getLastRow();

  // 1) Copiar/poner fórmula en "Days for Deadline" (col D)
  copyDaysFormulaDown_(sheet, r);

  // 2) Crear estructura de carpetas + link
  const rowObj = readRowToObject_(sheet, r, company);
  const division = company === "DLE" ? (rowObj.division || "Procurement") : "";
  const { url /*, id*/ } = createProjectStructure_(company, division, rowObj);
  writeLinkBack_(sheet, r, url);

  // (Eliminado: manejo de RFQ / Upload)

  return { url };
}

/**
 * Normalize Client entries in both company sheets to match the canonical
 * casing found in the hidden `Clients` sheet. This fixes cells that show
 * the data-validation warning "Invalid: Input must fall within specified range".
 */
function fixInvalidClients() {
  var ss = SpreadsheetApp.getActive();
  var clients = getClients_();
  var companies = ["SOS", "DLE"];
  var totalFixed = 0;

  companies.forEach(function(company){
    var map = COLMAP[company];
    var sh = ss.getSheetByName(map.SHEET_NAME);
    if (!sh) return;
    var last = sh.getLastRow();
    if (last < 2) return;

    var rng = sh.getRange(2, map.CLIENT, last - 1, 1);
    var vals = rng.getValues();
    var changed = 0;

    for (var i = 0; i < vals.length; i++) {
      var original = String(vals[i][0] || "").trim();
      if (!original) continue;
      var canonical = getCanonicalFromList_(original, clients);
      if (canonical && canonical !== original) {
        vals[i][0] = canonical;
        changed++;
      }
    }

    if (changed > 0) {
      rng.setValues(vals);
      totalFixed += changed;
    }
  });

  SpreadsheetApp.getUi().alert("Clients normalized: " + totalFixed + " cell(s) updated.");
}
/************** DASHBOARD KPI 2025 — OPTIMIZED FOR SPEED **************/

/* Grupos de estado -> como están en tu hoja */
const STATUS_GROUPS = {
  rfq_pending_create: ['Waiting RFQ Creation'],
  rfq_pending_send: ['Pending RFQ Sending'],
  // Count specific waiting sources too (also part of waiting_for_quotation bucket)
  waiting_libya: ['Waiting for offer by Libya'],
  waiting_usa: ['Waiting for offer by USA'],
  waiting_for_quotation: [
    'Waiting for offer by supplier','Waiting for offer by Libya','Waiting for offer by USA'
  ],
  ready_to_quote: ['Ready to quote'],
  // Proposal sub-statuses
  creating_technical_proposal: ['Creating Technical proposal'],
  technical_proposal_pending: ['Technical proposal pending to approval'],
  creating_estimation: ['Creating Estimation Cost'],
  estimation_pending: ['Estimation pending to approval'],
  proposals_ready: ['Proposals ready to Submit'],
  offer_submitted: ['Commercial proposal submitted'],
  technical_query: ['Technical Query','Waiting Client Reply – Tech Query','Waiting Client Reply - Tech Query'],
  technical_review: ['Technical Review'],
  restriction_from_supplier: ['Restriction from supplier'],
  no_proposal_tech: ['No proposal - Technical reasons'],
  no_proposal_com: ['No proposal - Comercial reasons'],
  on_hold: ['ON Hold - WFMD'],
  canceled_by_client: ['Canceled by client','Canceled by the client'],
  no_quote_deadline: ['No Quote – Deadline Passed','No Quote - Deadline Passed'],
  awarded: [
    // Exact matches
    'Awared','Awarded','Awarded Projects','Awarded projects','AWARDED','AWARDED PROJECTS','awared','AWARED',
    // Common variations
    'Awarded Project','Awarded project','AWARDED PROJECT','awarded project','Award','award','AWARD','Awar',
    // With spaces and special chars
    'Awarded  Projects','Awarded-Projects','Awarded_Projects','Awarded (Projects)',
    // Partial matches will be handled separately
  ]
};

// Pre-normalized status groups for faster lookup
const STATUS_GROUPS_NORM = (function(){
  const out = {};
  Object.keys(STATUS_GROUPS).forEach(k=> out[k] = STATUS_GROUPS[k].map(s => String(s||'').replace(/\s+/g,' ').trim().toLowerCase()));
  return out;
})();

// Fast year detection with regex compilation
const YEAR_REGEX = /-(\d{2})\b/;
function _noIsYear_(noValue, year){
  if (!noValue) return false;
  const m = String(noValue).match(YEAR_REGEX);
  if (!m) return false;
  const yy = 2000 + parseInt(m[1], 10);
  return yy === year;
}

// Get all awarded projects regardless of year for debugging
function getAllAwardedProjects(company) {
  try {
    const ss = SpreadsheetApp.getActive();
    const map = COLMAP[company];
    const sh = ss.getSheetByName(map.SHEET_NAME);

    if (!sh) {
      return { error: 'Sheet not found' };
    }

    const last = sh.getLastRow();
    if (last < 2) {
      return { error: 'No data found' };
    }

    const values = sh.getRange(2, 1, last-1, sh.getLastColumn()).getValues();
    const awardedProjects = [];
    const yearCounts = {};

    values.forEach((r, i) => {
      const noVal = r[map.NO - 1];
      const subject = (r[map.SUBJECT - 1] || '').toString().trim();
      const status = String(r[map.STATUS - 1] || '').trim();

      if (!noVal || !subject) return;

      const statusLower = status.toLowerCase();
      const isAwarded = statusLower.includes('award') || statusLower.includes('awared') || statusLower.includes('awarded');

      if (isAwarded) {
        // Extract year from project number
        const yearMatch = String(noVal).match(/-(\d{2})$/);
        const year = yearMatch ? (2000 + parseInt(yearMatch[1], 10)) : 'unknown';
        
        yearCounts[year] = (yearCounts[year] || 0) + 1;
        
        awardedProjects.push({
          row: i + 2,
          no: noVal,
          status: status,
          year: year,
          client: r[map.CLIENT - 1] || '',
          subject: subject
        });
      }
    });

    return { 
      awardedProjects, 
      yearCounts,
      total: awardedProjects.length 
    };

  } catch (error) {
    return { error: error.message };
  }
}

// Optimized empty company structure - matches HTML expectations
function _emptyCompanyAgg(){
  return {
    rfq_pending_create: 0,
    rfq_pending_send: 0,
    waiting_for_quotation: 0,
    ready_to_quote: 0,
    creating_technical_proposal: 0,
    technical_proposal_pending: 0,
    creating_estimation: 0,
    estimation_pending: 0,
    proposals_ready: 0,
    offer_submitted: 0,
    technical_query: 0,
    technical_review: 0,
    restriction_from_supplier: 0,
    no_proposal_tech: 0,
    no_proposal_com: 0,
    waiting_libya: 0,
    waiting_usa: 0,
    on_hold: 0,
    canceled_by_client: 0,
    no_quote_deadline: 0,
    awarded: 0,
    total: 0,
    projects_not_quoted: 0,
    // Additional fields used internally/for debugging
    total_received: 0,
    overdue_projects: 0,
    high_priority_projects: 0,
    offers_submitted: 0,
    awarded_projects: 0,
    conversion_offers_sent: 0,
    offer_conversion_rate: 0
  };
}

// Fast batch processing function for multiple years
function _processBatchDataMultiYear_(values, map, years) {
    const comp = _emptyCompanyAgg();
  const today = new Date();
  const processedRows = [];

  // Process all rows in one pass
    values.forEach((r, i) => {
      const rowNum = i + 2; // Actual row number for debugging

      // Enhanced validation: skip completely empty rows
      const isEmptyRow = !r || r.every(cell => !cell || String(cell).trim() === '');
      if (isEmptyRow) return;

      const noVal = r[map.NO - 1];
      const subject = (r[map.SUBJECT - 1] || '').toString().trim();

      // Enhanced validation: require both project number and subject
      if (!noVal || !subject) return;

    // Check if project is in any of the specified years
    const projectYear = _getProjectYear_(noVal);
    if (!projectYear || !years.includes(projectYear)) return;

    const status = String(r[map.STATUS - 1] || '').replace(/\s+/g,' ').trim().toLowerCase();
    const client = r[map.CLIENT - 1] || '';
    const leader = r[map.LEADER - 1] || '';
    const deadline = r[map.DEADLINE - 1];

    // Fast priority calculation
    let priority = 'medium';
    let daysToDeadline = null;
    if (deadline && deadline instanceof Date) {
      const diffTime = deadline.getTime() - today.getTime();
      daysToDeadline = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
      if (daysToDeadline < 0) priority = 'overdue';
      else if (daysToDeadline <= 7) priority = 'high';
      else if (daysToDeadline <= 30) priority = 'medium';
      else priority = 'low';
    }

      const rowObj = {
        row: rowNum,
        no: noVal,
        year: projectYear,
      client: client,
      subject: subject,
      deadline: deadline,
      leader: leader,
      status: r[map.STATUS - 1] || '',
      daysToDeadline: daysToDeadline,
      priority: priority
      };

      // Count row towards totals
      comp.total++;
      comp.total_received++;

    if (priority === 'overdue') comp.overdue_projects++;
    if (priority === 'high') comp.high_priority_projects++;

    // Fast status grouping with pre-compiled lookup
    let statusMatched = false;
    for (const [key, patterns] of Object.entries(STATUS_GROUPS_NORM)) {
      if (patterns.includes(status)) {
        comp[key]++;
        statusMatched = true;
        // Do not break: allow a status to contribute to both a specific bucket
        // (e.g., waiting_libya) and a general one (waiting_for_quotation)
      }
    }

    // If no status matched, check for awarded projects with enhanced flexible matching
    if (!statusMatched) {
      const statusLower = status.toLowerCase().replace(/\s+/g, ' ').trim();

      // Enhanced awarded detection - check for various patterns
      const awardedPatterns = [
        'award', 'awarded', 'awared', 'award project', 'awarded project',
        'award-project', 'awarded-project', 'award_project', 'awarded_project'
      ];

      const isAwarded = awardedPatterns.some(pattern => statusLower.includes(pattern));

      if (isAwarded) {
        comp.awarded++;
        statusMatched = true;
      } else {
        comp.waiting_for_quotation++;
      }
    }

    processedRows.push(rowObj);
  });

  // Fast KPI calculations
    comp.offers_submitted = comp.offer_submitted;
    comp.awarded_projects = comp.awarded;
  comp.projects_not_quoted = comp.no_proposal_tech + comp.no_proposal_com +
                             comp.restriction_from_supplier + comp.on_hold +
                             comp.canceled_by_client + comp.no_quote_deadline;

    comp.conversion_offers_sent = comp.offers_submitted ? (comp.awarded_projects / comp.offers_submitted) : 0;
    // Use total that the UI expects
    comp.offer_conversion_rate = comp.total ? (comp.awarded_projects / comp.total) : 0;

  return comp;
}

// Helper function to extract year from project number
function _getProjectYear_(noValue) {
  if (!noValue) return null;
  const m = String(noValue).match(YEAR_REGEX);
  if (!m) return null;
  return 2000 + parseInt(m[1], 10);
}

// Original single year function for backward compatibility
function _processBatchData_(values, map, year) {
  return _processBatchDataMultiYear_(values, map, [year]);
}

// Cached data storage
const CACHE_DURATION = 300; // 5 minutes cache
function _getCachedData_(cacheKey) {
  try {
    const cache = CacheService.getScriptCache();
    const cached = cache.get(cacheKey);
    if (cached && cached[0]) {
      return JSON.parse(cached[0]);
    }
  } catch (e) {
    console.log('Cache read error:', e);
  }
  return null;
}

function _setCachedData_(cacheKey, data) {
  try {
    const cache = CacheService.getScriptCache();
    cache.put(cacheKey, JSON.stringify(data), CACHE_DURATION);
  } catch (e) {
    console.log('Cache write error:', e);
  }
}

// Main optimized KPI function - now supports multiple years
function getKPIDashboardData(year) {
  year = year || 2025;

  // Check cache first
  const cacheKey = `kpi_dashboard_${year}`;
  const cached = _getCachedData_(cacheKey);
  if (cached && false) { // Cache disabled for testing
    return cached;
  }

  const startTime = new Date();
  const ss = SpreadsheetApp.getActive();
  const out = { year, companies: {}, loadTime: 0 };

  try {
    ['SOS','DLE'].forEach(company => {
      const map = COLMAP[company];
      const sh = ss.getSheetByName(map.SHEET_NAME);
      if (!sh) { 
        out.companies[company] = _emptyCompanyAgg(); 
        return; 
      }

      const last = sh.getLastRow();
      if (last < 2) { 
        out.companies[company] = _emptyCompanyAgg(); 
        return; 
      }

      // Read data in one batch operation
      const values = sh.getRange(2, 1, last-1, sh.getLastColumn()).getValues();
      
      // Process data efficiently - now includes both 2024 and 2025
      const comp = _processBatchDataMultiYear_(values, map, [2024, 2025]);
    out.companies[company] = comp;
  });

    const endTime = new Date();
    out.loadTime = Math.round((endTime - startTime) / 1000);
    
    // Cache the result
    _setCachedData_(cacheKey, out);

  return out;
    
  } catch (error) {
    console.error('KPI Dashboard Error:', error);
    return { year, companies: {}, error: error.message, loadTime: 0 };
  }
}

// Fast monthly trends with minimal processing
function getMonthlyTrendData(year) {
  year = year || 2025;
  
  const cacheKey = `monthly_trends_${year}`;
  const cached = _getCachedData_(cacheKey);
  if (cached) return cached;

  const ss = SpreadsheetApp.getActive();
  const out = { year, monthlyData: {} };

  ['SOS','DLE'].forEach(company => {
    const map = COLMAP[company];
    const sh = ss.getSheetByName(map.SHEET_NAME);
    if (!sh) return;

    const last = sh.getLastRow();
    if (last < 2) return;

    // Only read necessary columns for speed
    const values = sh.getRange(2, map.NO, last-1, 1).getValues(); // Project numbers
    const receivedValues = sh.getRange(2, map.RECEIVED, last-1, 1).getValues(); // Received dates
    
    const monthlyCounts = Array(12).fill(0);
    
    values.forEach((r, i) => {
      const noVal = r[0];
      if (!_noIsYear_(noVal, year)) return;

      const received = receivedValues[i][0];
      if (received && received instanceof Date) {
        const month = received.getMonth();
        monthlyCounts[month]++;
      }
    });

    out.monthlyData[company] = monthlyCounts;
  });

  _setCachedData_(cacheKey, out);
  return out;
}

// Fast client stats
function getClientStats(year) {
  year = year || 2025;
  
  const cacheKey = `client_stats_${year}`;
  const cached = _getCachedData_(cacheKey);
  if (cached) return cached;

  const data = getKPIDashboardData(year);
  const out = { year, companies: {} };

  ['SOS','DLE'].forEach(company => {
    if (data.companies[company]) {
      const comp = data.companies[company];
      const clientCounts = {};
      
      // Count clients from all status groups
      Object.values(comp).forEach(statusGroup => {
        if (statusGroup.items && Array.isArray(statusGroup.items)) {
          statusGroup.items.forEach(item => {
            if (item.client) {
              clientCounts[item.client] = (clientCounts[item.client] || 0) + 1;
            }
          });
        }
      });

      // Get top 10 clients
      const topClients = Object.entries(clientCounts)
        .sort(([,a], [,b]) => b - a)
        .slice(0, 10)
        .reduce((obj, [client, count]) => {
          obj[client] = count;
          return obj;
        }, {});

      out.companies[company] = {
        topClients: topClients,
        totalClients: Object.keys(clientCounts).length,
        averageProjectsPerClient: comp.total_received / Object.keys(clientCounts).length || 0
      };
    }
  });

  _setCachedData_(cacheKey, out);
  return out;
}

// Force refresh function to clear cache
function refreshKPIDashboard() {
  try {
    const cache = CacheService.getScriptCache();
    cache.removeAll(['kpi_dashboard_2025', 'monthly_trends_2025', 'client_stats_2025']);
    return { success: true, message: 'Cache cleared successfully' };
  } catch (e) {
    return { success: false, error: e.message };
  }
}


/************** SIMPLE WORKING DASHBOARD **************/

// Simple function to get basic project counts
function getSimpleDashboardData(year) {
  year = year || 2025;
  
  try {
    const ss = SpreadsheetApp.getActive();
    const out = { year, companies: {} };
    
    ['SOS','DLE'].forEach(company => {
      const map = COLMAP[company];
      const sh = ss.getSheetByName(map.SHEET_NAME);
      
      if (!sh) {
        out.companies[company] = { total: 0, statuses: {} };
        return;
      }
      
      const last = sh.getLastRow();
      if (last < 2) {
        out.companies[company] = { total: 0, statuses: {} };
        return;
      }
      
      // Read only the columns we need
      const noCol = sh.getRange(2, map.NO, last-1, 1).getValues();
      const statusCol = sh.getRange(2, map.STATUS, last-1, 1).getValues();
      
      const statuses = {};
      let total = 0;
      
      for (let i = 0; i < noCol.length; i++) {
        const noVal = noCol[i][0];
        if (_noIsYear_(noVal, year)) {
          total++;
          const status = String(statusCol[i][0] || '').trim();
          statuses[status] = (statuses[status] || 0) + 1;
        }
      }
      
      out.companies[company] = { total, statuses };
    });
    
    return out;
    
  } catch (error) {
    return { year, companies: {}, error: error.message };
  }
}

// Removed duplicate function - keeping the one with status mapping below

// Simple dashboard opener
function openSimpleDashboard() {
  const html = HtmlService.createHtmlOutputFromFile('simple-dashboard')
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .setWidth(1200)
    .setHeight(800);
  SpreadsheetApp.getUi().showModelessDialog(html, '📊 Simple KPI Dashboard ' + new Date().getFullYear());
}

// Quick test function
function simpleQuickTest(){
  const startTime = new Date();
  const data = getSimpleDashboardData(2025);
  const endTime = new Date();
  const loadTime = Math.round((endTime - startTime) / 1000);
  
  if (data.error) {
    SpreadsheetApp.getUi().alert('❌ Error: ' + data.error);
    return;
  }
  
  const sos = data.companies.SOS;
  const dle = data.companies.DLE;
  
  let message = `📊 Simple Dashboard Test Results for 2025:\n\n`;
  message += `⏱️ Load Time: ${loadTime} seconds\n\n`;
  
  message += `🏢 SOS:\n`;
  message += `  • Total Projects: ${sos.total}\n`;
  message += `  • Statuses: ${Object.keys(sos.statuses).length}\n\n`;
  
  message += `🏭 DLE:\n`;
  message += `  • Total Projects: ${dle.total}\n`;
  message += `  • Statuses: ${Object.keys(dle.statuses).length}\n`;
  
  SpreadsheetApp.getUi().alert(message);
}

/************** KPI DASHBOARD FUNCTIONS **************/

// Create empty KPI structure
function createEmptyKPIStructure() {
  return {
    // RFQ Process
    rfq_pending_create: 0,
    rfq_pending_send: 0,
    waiting_for_quotation: 0,
    
    // Proposal Process
    ready_to_quote: 0,
    creating_technical_proposal: 0,
    technical_proposal_pending: 0,
    creating_estimation: 0,
    estimation_pending: 0,
    proposals_ready: 0,
    offer_submitted: 0,
    
    // Project Not Quoted
    technical_query: 0,
    technical_review: 0,
    restriction_from_supplier: 0,
    no_proposal_tech: 0,
    no_proposal_com: 0,
    waiting_libya: 0,
    waiting_usa: 0,
    on_hold: 0,
    canceled_by_client: 0,
    
    // Summary
    awarded: 0,
    total: 0,
    projects_not_quoted: 0
  };
}

// Categorize status into KPI structure
function categorizeStatus(status, kpiData) {
  const statusTrimmed = String(status || '').trim();

  // RFQ Process - exact matches
  if (statusTrimmed === 'Waiting RFQ Creation') {
    kpiData.rfq_pending_create++;
  } else if (statusTrimmed === 'Pending RFQ Sending') {
    kpiData.rfq_pending_send++;
  } else if (statusTrimmed === 'Waiting for offer by supplier' ||
             statusTrimmed === 'Waiting for offer by Libya' ||
             statusTrimmed === 'Waiting for offer by USA') {
    kpiData.waiting_for_quotation++;
  }

  // Proposal Process - exact matches
  else if (statusTrimmed === 'Ready to quote') {
    kpiData.ready_to_quote++;
  } else if (statusTrimmed === 'Creating Technical proposal') {
    kpiData.creating_technical_proposal++;
  } else if (statusTrimmed === 'Technical proposal pending to approval') {
    kpiData.technical_proposal_pending++;
  } else if (statusTrimmed === 'Creating Estimation Cost') {
    kpiData.creating_estimation++;
  } else if (statusTrimmed === 'Estimation pending to approval') {
    kpiData.estimation_pending++;
  } else if (statusTrimmed === 'Proposals ready to Submit') {
    kpiData.proposals_ready++;
  } else if (statusTrimmed === 'Commercial proposal submitted') {
    kpiData.offer_submitted++;
  }

  // Project Not Quoted - exact matches
  else if (statusTrimmed === 'Technical Query' || statusTrimmed === 'Waiting Client Reply – Tech Query' || statusTrimmed === 'Waiting Client Reply - Tech Query') {
    kpiData.technical_query++;
  } else if (statusTrimmed === 'Technical Review') {
    kpiData.technical_review++;
  } else if (statusTrimmed === 'Restriction from supplier') {
    kpiData.restriction_from_supplier++;
  } else if (statusTrimmed === 'No proposal - Technical reasons') {
    kpiData.no_proposal_tech++;
  } else if (statusTrimmed === 'No proposal - Comercial reasons') {
    kpiData.no_proposal_com++;
  } else if (statusTrimmed === 'Waiting for offer by Libya') {
    kpiData.waiting_libya++;
  } else if (statusTrimmed === 'Waiting for offer by USA') {
    kpiData.waiting_usa++;
  } else if (statusTrimmed === 'ON Hold - WFMD') {
    kpiData.on_hold++;
  } else if (statusTrimmed === 'Canceled by client' || statusTrimmed === 'Canceled by the client') {
    kpiData.canceled_by_client++;
  } else if (statusTrimmed === 'No Quote – Deadline Passed' || statusTrimmed === 'No Quote - Deadline Passed') {
    kpiData.no_quote_deadline++;
  }

  // Awarded projects - exact matches
  else if (statusTrimmed === 'Awared' || statusTrimmed === 'Awarded' || statusTrimmed === 'awared' || statusTrimmed === 'AWARED') {
    kpiData.awarded++;
  }

  // Awarded projects - partial matches (moved up for better detection)
  else if (statusTrimmed.toLowerCase().includes('award')) {
    kpiData.awarded++;
  }

  // If no match, try partial matches for backward compatibility
  else {
    const statusLower = statusTrimmed.toLowerCase();

    // RFQ Process - partial matches
    if (statusLower.includes('rfq') && statusLower.includes('create')) {
      kpiData.rfq_pending_create++;
    } else if (statusLower.includes('rfq') && statusLower.includes('send')) {
      kpiData.rfq_pending_send++;
    } else if (statusLower.includes('waiting') && statusLower.includes('quotation')) {
      kpiData.waiting_for_quotation++;
    }

    // Proposal Process - partial matches
    else if (statusLower.includes('ready') && statusLower.includes('quote')) {
      kpiData.ready_to_quote++;
    } else if (statusLower.includes('technical') && statusLower.includes('proposal') && statusLower.includes('create')) {
      kpiData.creating_technical_proposal++;
    } else if (statusLower.includes('technical') && statusLower.includes('proposal') && statusLower.includes('pending')) {
      kpiData.technical_proposal_pending++;
    } else if (statusLower.includes('estimation') && statusLower.includes('create')) {
      kpiData.creating_estimation++;
    } else if (statusLower.includes('estimation') && statusLower.includes('pending')) {
      kpiData.estimation_pending++;
    } else if (statusLower.includes('proposals') && statusLower.includes('ready')) {
      kpiData.proposals_ready++;
    } else if (statusLower.includes('commercial') && statusLower.includes('submitted')) {
      kpiData.offer_submitted++;
    }

    // Project Not Quoted - partial matches
    else if (statusLower.includes('technical') && statusLower.includes('query')) {
      kpiData.technical_query++;
    } else if (statusLower.includes('technical') && statusLower.includes('review')) {
      kpiData.technical_review++;
    } else if (statusLower.includes('restriction') && statusLower.includes('supplier')) {
      kpiData.restriction_from_supplier++;
    } else if (statusLower.includes('no proposal') && statusLower.includes('technical')) {
      kpiData.no_proposal_tech++;
    } else if (statusLower.includes('no proposal') && statusLower.includes('commercial')) {
      kpiData.no_proposal_com++;
    } else if (statusLower.includes('libya')) {
      kpiData.waiting_libya++;
    } else if (statusLower.includes('usa')) {
      kpiData.waiting_usa++;
    } else if (statusLower.includes('hold') && statusLower.includes('wfmd')) {
      kpiData.on_hold++;
    } else if (statusLower.includes('canceled') && statusLower.includes('client')) {
      kpiData.canceled_by_client++;
    } else if (statusLower.includes('deadline') && statusLower.includes('passed')) {
      kpiData.no_quote_deadline++;
    }

    // If still no match, add to waiting for quotation as fallback
    else {
      kpiData.waiting_for_quotation++;
    }
  }
}

// Calculate total projects not quoted
function calculateNotQuotedTotal(kpiData) {
  return kpiData.technical_query + 
         kpiData.technical_review + 
         kpiData.restriction_from_supplier + 
         kpiData.no_proposal_tech + 
         kpiData.no_proposal_com + 
         kpiData.waiting_libya + 
         kpiData.waiting_usa + 
         kpiData.on_hold + 
         kpiData.canceled_by_client;
}

// Removed duplicate function - keeping the cached version above

// Function to get projects by status for click functionality - now supports multiple years
function getProjectsByStatusSimple(company, status, year) {
  year = year || 2025;
  
  try {
    const ss = SpreadsheetApp.getActive();
    const map = COLMAP[company];
    const sh = ss.getSheetByName(map.SHEET_NAME);
    
    if (!sh) return { projects: [] };
    
    const last = sh.getLastRow();
    if (last < 2) return { projects: [] };
    
    const noCol = sh.getRange(2, map.NO, last-1, 1).getValues();
    const statusCol = sh.getRange(2, map.STATUS, last-1, 1).getValues();
    const clientCol = sh.getRange(2, map.CLIENT, last-1, 1).getValues();
    const subjectCol = sh.getRange(2, map.SUBJECT, last-1, 1).getValues();
    
    const projects = [];
    
    // Map display names to actual status names in the spreadsheet
    const statusMapping = {
      'RFQ Pending to create': 'Waiting RFQ Creation',
      'RFQ Pending to send': 'Pending RFQ Sending',
      'Waiting to receive the quotation': ['Waiting for offer by supplier', 'Waiting for offer by Libya', 'Waiting for offer by USA'],
      'Ready to quote': 'Ready to quote',
      'Creating Technical proposal': 'Creating Technical proposal',
      'Technical proposal pending to approval': 'Technical proposal pending to approval',
      'Creating Estimation Cost': 'Creating Estimation Cost',
      'Estimation pending to approval': 'Estimation pending to approval',
      'Proposals ready to Submit': 'Proposals ready to Submit',
      'Offer submitted': 'Commercial proposal submitted',
      'Technical Query': 'Technical Query',
      'Technical Review': 'Technical Review',
      'Restriction from supplier': 'Restriction from supplier',
      'No proposal - Technical reasons': 'No proposal - Technical reasons',
      'No proposal - Commercial reasons': 'No proposal - Comercial reasons',
      'Waiting for offer by Libya': 'Waiting for offer by Libya',
      'Waiting for offer by USA': 'Waiting for offer by USA',
      'on Hold WFMD': 'ON Hold - WFMD',
      'Canceled by the client': 'Canceled by client',
      'Awarded': ['Awared', 'Awarded', 'Awarded Projects', 'Awarded projects', 'AWARDED', 'AWARDED PROJECTS', 'awared', 'AWARED']
    };
    
    const actualStatuses = statusMapping[status];
    const statusList = Array.isArray(actualStatuses) ? actualStatuses : [actualStatuses];
    
    for (let i = 0; i < noCol.length; i++) {
      const noVal = noCol[i][0];
      const rowStatus = String(statusCol[i][0] || '').trim();
      
      // Check if project is from 2024 or 2025
      const projectYear = _getProjectYear_(noVal);
      const isTargetYear = projectYear === 2024 || projectYear === 2025;
      
      // For awarded projects, also check with flexible matching
      let isAwarded = false;
      if (status === 'Awarded') {
        const statusLower = rowStatus.toLowerCase();
        isAwarded = statusLower.includes('award') || statusLower.includes('awared') || statusLower.includes('awarded');
      }
      
      const statusMatches = statusList.includes(rowStatus) || (status === 'Awarded' && isAwarded);
      
      if (isTargetYear && statusMatches) {
        projects.push({
          no: noVal,
          year: projectYear,
          client: clientCol[i][0] || '',
          subject: subjectCol[i][0] || '',
          status: rowStatus
        });
      }
    }
    
    // Sort by year (2025 first, then 2024) and then by project number
    projects.sort((a, b) => {
      if (a.year !== b.year) {
        return b.year - a.year; // 2025 first
      }
      return a.no.localeCompare(b.no);
    });
    
    return { projects };
    
  } catch (error) {
    return { projects: [], error: error.message };
  }
}

// Function to get Ready to Quote sub-statuses
function getReadyToQuoteSubStatuses(company, year) {
  year = year || 2025;

  try {
    const ss = SpreadsheetApp.getActive();
    const map = COLMAP[company];
    const sh = ss.getSheetByName(map.SHEET_NAME);

    if (!sh) return { subStatuses: {} };

    const last = sh.getLastRow();
    if (last < 2) return { subStatuses: {} };

    const noCol = sh.getRange(2, map.NO, last-1, 1).getValues();
    const statusCol = sh.getRange(2, map.STATUS, last-1, 1).getValues();
    const clientCol = sh.getRange(2, map.CLIENT, last-1, 1).getValues();
    const subjectCol = sh.getRange(2, map.SUBJECT, last-1, 1).getValues();

    // Define sub-status categories for "Ready to quote"
    const subStatuses = {
      'Ready to quote': [],
      'Creating Technical proposal': [],
      'Technical proposal pending to approval': [],
      'Creating Estimation Cost': [],
      'Estimation pending to approval': [],
      'Proposals ready to Submit': []
    };

    // Status mapping for sub-statuses
    const subStatusMapping = {
      'Ready to quote': ['Ready to quote'],
      'Creating Technical proposal': ['Creating Technical proposal'],
      'Technical proposal pending to approval': ['Technical proposal pending to approval'],
      'Creating Estimation Cost': ['Creating Estimation Cost'],
      'Estimation pending to approval': ['Estimation pending to approval'],
      'Proposals ready to Submit': ['Proposals ready to Submit']
    };

    for (let i = 0; i < noCol.length; i++) {
      const noVal = noCol[i][0];
      const rowStatus = String(statusCol[i][0] || '').trim();

      if (_noIsYear_(noVal, year)) {
        const project = {
          no: noVal,
          client: clientCol[i][0] || '',
          subject: subjectCol[i][0] || '',
          status: rowStatus
        };

        // Check which sub-status category this project belongs to
        let categorized = false;
        for (const [category, statuses] of Object.entries(subStatusMapping)) {
          if (statuses.includes(rowStatus)) {
            subStatuses[category].push(project);
            categorized = true;
            break;
          }
        }

        // If not categorized into any specific sub-status, put it in "Ready to quote"
        if (!categorized && rowStatus === 'Ready to quote') {
          subStatuses['Ready to quote'].push(project);
        }
      }
    }

    return { subStatuses };

  } catch (error) {
    return { subStatuses: {}, error: error.message };
  }
}

// Function to update project status
function updateProjectStatus(company, projectNo, newStatus, year) {
  year = year || 2025;
  
  try {
    const ss = SpreadsheetApp.getActive();
    const map = COLMAP[company];
    const sh = ss.getSheetByName(map.SHEET_NAME);
    
  if (!sh) {
      return { success: false, error: 'Sheet not found' };
    }
    
    const last = sh.getLastRow();
    if (last < 2) {
      return { success: false, error: 'No data found' };
    }
    
    // Find the project row
    const noCol = sh.getRange(2, map.NO, last-1, 1).getValues();
    let projectRow = -1;
    
    for (let i = 0; i < noCol.length; i++) {
      const noVal = noCol[i][0];
      if (_noIsYear_(noVal, year) && String(noVal).trim() === String(projectNo).trim()) {
        projectRow = i + 2; // +2 because we start from row 2 and arrays are 0-indexed
        break;
      }
    }
    
    if (projectRow === -1) {
      return { success: false, error: 'Project not found' };
    }
    
    // Update the status
    sh.getRange(projectRow, map.STATUS).setValue(newStatus);
    
    return { success: true, message: `Project ${projectNo} status updated to "${newStatus}"` };
    
  } catch (error) {
    return { success: false, error: error.message };
  }
}

// Simple function to get ALL available statuses for editing
function getAllValidStatuses(company) {
  try {
    const ss = SpreadsheetApp.getActive();
    const map = COLMAP[company];
    const sh = ss.getSheetByName(map.SHEET_NAME);
    
    if (!sh) {
      return { statuses: [], error: 'Sheet not found' };
    }
    
    // Get the data validation for the status column
    const statusColumn = map.STATUS;
    const range = sh.getRange(2, statusColumn, Math.max(1, sh.getLastRow() - 1), 1);
    const validation = range.getDataValidation();
    
    if (validation && validation.getCriteriaType() === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) {
      // Get ALL valid values from data validation
      const validValues = validation.getCriteriaValues()[0];
      if (validValues && validValues.length > 0) {
        return { statuses: validValues.sort() };
      }
    }
    
    // If no data validation found, get all unique statuses from the data
    const last = sh.getLastRow();
    if (last >= 2) {
      const statusCol = sh.getRange(2, map.STATUS, last-1, 1).getValues();
      const uniqueStatuses = [...new Set(statusCol.map(row => String(row[0] || '').trim()).filter(status => status))].sort();
      return { statuses: uniqueStatuses };
    }
    
    return { statuses: [], error: 'No statuses found' };
    
  } catch (error) {
    return { statuses: [], error: error.message };
  }
}

// Debug function to see all statuses in spreadsheet
function debugAllStatuses() {
  try {
  const ss = SpreadsheetApp.getActive();
    let message = `🔍 All Statuses in Spreadsheet:\n\n`;

    ['SOS', 'DLE'].forEach(company => {
    const map = COLMAP[company];
    const sh = ss.getSheetByName(map.SHEET_NAME);
      
      if (sh) {
        message += `🏢 ${company}:\n`;
        
        // Check data validation first
        const statusColumn = map.STATUS;
        const range = sh.getRange(2, statusColumn, Math.max(1, sh.getLastRow() - 1), 1);
        const validation = range.getDataValidation();
        
        if (validation && validation.getCriteriaType() === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) {
          const validValues = validation.getCriteriaValues()[0];
          if (validValues && validValues.length > 0) {
            message += `  📋 Data Validation Rules (${validValues.length}):\n`;
            validValues.forEach(status => {
              message += `    • "${status}"\n`;
            });
          }
        } else {
          message += `  ⚠️ No data validation found\n`;
        }
        
        // Also show actual data
        const last = sh.getLastRow();
        if (last >= 2) {
          const statusCol = sh.getRange(2, map.STATUS, last-1, 1).getValues();
          const uniqueStatuses = [...new Set(statusCol.map(row => String(row[0] || '').trim()).filter(status => status))].sort();
          
          message += `  📊 Actual Data (${uniqueStatuses.length}):\n`;
          uniqueStatuses.forEach(status => {
            message += `    • "${status}"\n`;
          });
        }
        message += `\n`;
      }
    });
    
    SpreadsheetApp.getUi().alert(message);
    
  } catch (error) {
    SpreadsheetApp.getUi().alert('Error: ' + error.message);
  }
}

// Test function for all valid statuses
function testAllValidStatuses() {
  const sosStatuses = getAllValidStatuses('SOS');
  const dleStatuses = getAllValidStatuses('DLE');
  
  let message = `🧪 All Valid Statuses Test:\n\n`;
  message += `These are ALL statuses available in your spreadsheet:\n\n`;
  
  message += `🏢 SOS Statuses (${sosStatuses.statuses.length}):\n`;
  if (sosStatuses.statuses.length > 0) {
    sosStatuses.statuses.forEach(status => {
      message += `  • "${status}"\n`;
    });
  } else {
    message += `  ⚠️ No statuses found\n`;
  }
  
  message += `\n🏭 DLE Statuses (${dleStatuses.statuses.length}):\n`;
  if (dleStatuses.statuses.length > 0) {
    dleStatuses.statuses.forEach(status => {
      message += `  • "${status}"\n`;
    });
  } else {
    message += `  ⚠️ No statuses found\n`;
  }
  
  if (sosStatuses.error) {
    message += `\n❌ SOS Error: ${sosStatuses.error}`;
  }
  if (dleStatuses.error) {
    message += `\n❌ DLE Error: ${dleStatuses.error}`;
  }
  
  SpreadsheetApp.getUi().alert(message);
}

// Test the Ready to Quote projects
function testReadyToQuoteProjects() {
  try {
    const sosResult = getReadyToQuoteSubStatuses('SOS', 2025);
    const dleResult = getReadyToQuoteSubStatuses('DLE', 2025);
    
    let message = `🧪 Ready to Quote Projects Test:\n\n`;
    
    message += `🏢 SOS - Ready to Quote Projects:\n`;
    if (sosResult.subStatuses && sosResult.subStatuses['Ready to quote']) {
      const count = sosResult.subStatuses['Ready to quote'].length;
      message += `  Found ${count} projects\n`;
      if (count > 0) {
        sosResult.subStatuses['Ready to quote'].forEach(project => {
          message += `    • ${project.no}: ${project.subject}\n`;
        });
      }
    } else {
      message += `  ❌ No Ready to quote projects found\n`;
    }
    
    message += `\n🏭 DLE - Ready to Quote Projects:\n`;
    if (dleResult.subStatuses && dleResult.subStatuses['Ready to quote']) {
      const count = dleResult.subStatuses['Ready to quote'].length;
      message += `  Found ${count} projects\n`;
      if (count > 0) {
        dleResult.subStatuses['Ready to quote'].forEach(project => {
          message += `    • ${project.no}: ${project.subject}\n`;
        });
      }
    } else {
      message += `  ❌ No Ready to quote projects found\n`;
    }
    
    SpreadsheetApp.getUi().alert(message);
    
  } catch (error) {
    SpreadsheetApp.getUi().alert('Error in Ready to Quote test: ' + error.message);
  }
}

// Test function for Ready to Quote sub-statuses
function testReadyToQuoteSubStatuses() {
  try {
    const sosResult = getReadyToQuoteSubStatuses('SOS', 2025);
    const dleResult = getReadyToQuoteSubStatuses('DLE', 2025);

    let message = `🧪 Ready to Quote Sub-Statuses Test:\n\n`;

    message += `🏢 SOS - Ready to Quote Sub-Statuses:\n`;
    if (sosResult.error) {
      message += `  ❌ Error: ${sosResult.error}\n`;
    } else {
      const subStatuses = sosResult.subStatuses;
      let totalSOS = 0;
      Object.keys(subStatuses).forEach(category => {
        const count = subStatuses[category].length;
        totalSOS += count;
        message += `  • ${category}: ${count} projects\n`;
        // Show first project as example
        if (subStatuses[category].length > 0) {
          const project = subStatuses[category][0];
          message += `    - Example: ${project.no} - ${project.subject}\n`;
        }
      });
      message += `  📊 Total: ${totalSOS} projects\n`;
    }

    message += `\n🏭 DLE - Ready to Quote Sub-Statuses:\n`;
    if (dleResult.error) {
      message += `  ❌ Error: ${dleResult.error}\n`;
    } else {
      const subStatuses = dleResult.subStatuses;
      let totalDLE = 0;
      Object.keys(subStatuses).forEach(category => {
        const count = subStatuses[category].length;
        totalDLE += count;
        message += `  • ${category}: ${count} projects\n`;
        // Show first project as example
        if (subStatuses[category].length > 0) {
          const project = subStatuses[category][0];
          message += `    - Example: ${project.no} - ${project.subject}\n`;
        }
      });
      message += `  📊 Total: ${totalDLE} projects\n`;
    }

    SpreadsheetApp.getUi().alert(message);

  } catch (error) {
    SpreadsheetApp.getUi().alert('❌ Test Error: ' + error.message);
  }
}

// Debug function to test KPI dashboard data
function debugKPIDashboard() {
  try {
    const data = getKPIDashboardData(2025);

    let message = `🔍 KPI Dashboard Debug:\n\n`;

    if (data.error) {
      message += `❌ Error: ${data.error}\n`;
      SpreadsheetApp.getUi().alert(message);
      return;
    }

    message += `📊 Data Structure Check:\n`;
    message += `Year: ${data.year}\n`;
    message += `Companies: ${Object.keys(data.companies).join(', ')}\n\n`;

    ['SOS', 'DLE'].forEach(company => {
      if (data.companies[company]) {
        const comp = data.companies[company];
        message += `🏢 ${company}:\n`;
        message += `  • rfq_pending_create: ${comp.rfq_pending_create}\n`;
        message += `  • ready_to_quote: ${comp.ready_to_quote}\n`;
        message += `  • offer_submitted: ${comp.offer_submitted}\n`;
        message += `  • awarded: ${comp.awarded}\n`;
        message += `  • total: ${comp.total}\n\n`;
      } else {
        message += `🏢 ${company}: ❌ No data\n\n`;
      }
    });

    SpreadsheetApp.getUi().alert(message);

  } catch (error) {
    SpreadsheetApp.getUi().alert('❌ Debug Error: ' + error.message);
  }
}

// Debug function to check all awarded projects by year
function debugAllAwardedProjects() {
  try {
    const result = getAllAwardedProjects('DLE');
    
    if (result.error) {
      SpreadsheetApp.getUi().alert('❌ Error: ' + result.error);
      return;
    }

    let message = `🔍 All Awarded Projects in DLE:\n\n`;
    message += `📊 Total Awarded Projects: ${result.total}\n\n`;
    
    message += `📅 By Year:\n`;
    Object.entries(result.yearCounts)
      .sort(([a], [b]) => b - a)
      .forEach(([year, count]) => {
        message += `  • ${year}: ${count} projects\n`;
      });

    message += `\n📋 All Awarded Projects:\n`;
    result.awardedProjects.forEach(project => {
      message += `• Row ${project.row}: ${project.no} (${project.year}) - "${project.status}" - ${project.subject}\n`;
    });

    SpreadsheetApp.getUi().alert(message);

  } catch (error) {
    SpreadsheetApp.getUi().alert('❌ Debug Error: ' + error.message);
  }
}

// KPI Test function
function kpiQuickTest() {
  const startTime = new Date();
  const data = getKPIDashboardData(2025);
  const endTime = new Date();
  const loadTime = Math.round((endTime - startTime) / 1000);

  if (data.error) {
    SpreadsheetApp.getUi().alert('❌ Error: ' + data.error);
    return;
  }

  const sos = data.companies.SOS;
  const dle = data.companies.DLE;

  let message = `📊 KPI Dashboard Test Results for 2025:\n\n`;
  message += `⏱️ Load Time: ${loadTime} seconds\n\n`;

  message += `🏢 SOS:\n`;
  message += `  • Total Projects: ${sos.total}\n`;
  message += `  • Awarded: ${sos.awarded}\n`;
  message += `  • Not Quoted: ${sos.projects_not_quoted}\n\n`;

  message += `🏭 DLE:\n`;
  message += `  • Total Projects: ${dle.total}\n`;
  message += `  • Awarded: ${dle.awarded}\n`;
  message += `  • Not Quoted: ${dle.projects_not_quoted}\n`;

  SpreadsheetApp.getUi().alert(message);
}



// Menu wrapper for cache clearing
function clearCacheMenu() {
  try {
    const cache = CacheService.getScriptCache();
    cache.removeAll(['kpi_dashboard_2025', 'monthly_trends_2025', 'client_stats_2025']);
    
    // Force refresh by calling the function directly
    const data = getKPIDashboardData(2025);
    const dle = data.companies.DLE;
    
    let message = `✅ Cache cleared successfully!\n\n`;
    message += `📊 Current DLE Data (2024+2025):\n`;
    message += `  • Total Projects: ${dle.total}\n`;
    message += `  • Awarded Projects: ${dle.awarded}\n`;
    message += `  • Offer Submitted: ${dle.offer_submitted}\n\n`;
    message += `The dashboard will now show fresh data!`;
    
    SpreadsheetApp.getUi().alert(message);
  } catch (error) {
    SpreadsheetApp.getUi().alert('❌ Error clearing cache: ' + error.message);
  }
}

// Debug function to check dashboard data
function debugDashboardData() {
  try {
    const data = getKPIDashboardData(2025);
    if (data.error) {
      SpreadsheetApp.getUi().alert('❌ Error: ' + data.error);
      return;
    }

    const sos = data.companies.SOS;
    const dle = data.companies.DLE;

    let message = `🔍 Dashboard Data Debug (2024 + 2025):\n\n`;
    
    message += `🏢 SOS Results:\n`;
    message += `  • Total Projects: ${sos.total}\n`;
    message += `  • Awarded Projects: ${sos.awarded}\n`;
    message += `  • Offer Submitted: ${sos.offer_submitted}\n\n`;
    
    message += `🏭 DLE Results:\n`;
    message += `  • Total Projects: ${dle.total}\n`;
    message += `  • Awarded Projects: ${dle.awarded}\n`;
    message += `  • Offer Submitted: ${dle.offer_submitted}\n`;
    message += `  • Projects Not Quoted: ${dle.projects_not_quoted}\n`;
    message += `  • Conversion Rate: ${(dle.offer_conversion_rate * 100).toFixed(1)}%\n\n`;

    message += `📊 Status Breakdown (DLE):\n`;
    message += `  • RFQ Pending Create: ${dle.rfq_pending_create}\n`;
    message += `  • RFQ Pending Send: ${dle.rfq_pending_send}\n`;
    message += `  • Waiting for Quotation: ${dle.waiting_for_quotation}\n`;
    message += `  • Ready to Quote: ${dle.ready_to_quote}\n`;
    message += `  • Offer Submitted: ${dle.offer_submitted}\n`;
    message += `  • Awarded: ${dle.awarded}\n`;

    SpreadsheetApp.getUi().alert(message);

  } catch (error) {
    SpreadsheetApp.getUi().alert('❌ Debug Error: ' + error.message);
  }
}

// Test function to verify awarded projects fix
function testAwardedProjectsFix() {
  try {
    // Check the new multi-year dashboard data
    const data = getKPIDashboardData(2025);
    if (data.error) {
      SpreadsheetApp.getUi().alert('❌ Error: ' + data.error);
      return;
    }

    const dle = data.companies.DLE;

    let message = `🧪 Awarded Projects Analysis (2024 + 2025):\n\n`;
    
    message += `🏭 DLE Results (Combined 2024 & 2025):\n`;
    message += `  • Total Projects: ${dle.total}\n`;
    message += `  • Awarded Projects: ${dle.awarded}\n`;
    message += `  • Expected: 8+\n\n`;

    if (dle.awarded >= 7) {
      message += `✅ SUCCESS: Now showing ${dle.awarded} awarded projects!\n`;
      message += `   The dashboard now includes both 2024 and 2025 projects.\n\n`;
    } else {
      message += `❌ ISSUE: Still only showing ${dle.awarded} awarded projects.\n`;
      message += `   Check if there are more awarded projects in the data.\n\n`;
    }

    message += `📊 Other DLE KPIs:\n`;
    message += `  • Offer Submitted: ${dle.offer_submitted}\n`;
    message += `  • Projects Not Quoted: ${dle.projects_not_quoted}\n`;
    message += `  • Conversion Rate: ${(dle.offer_conversion_rate * 100).toFixed(1)}%\n`;

    SpreadsheetApp.getUi().alert(message);

  } catch (error) {
    SpreadsheetApp.getUi().alert('❌ Test Error: ' + error.message);
  }
}

