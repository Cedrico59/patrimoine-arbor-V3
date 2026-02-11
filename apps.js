/*******************************
 * CONFIG
 *******************************/
const SPREADSHEET_ID = "1Q-HaZs_nMcJRiH0lNu-NpYbdlvWFNWSK8dQPGl9vJNU";

const SHEET_TRAVAUX = "tableau_Elagages/Abattages";
const SHEET_PATRIMOINE = "Patrimoine_arboré";
const SHEET_HISTORIQUE = "Historique";

const DRIVE_FOLDER_ID = "1bC7CsCGBeQNp5ADelZ0SIXGjo12uhiUS";
const MAIRIE_LOGO_URL = "https://raw.githubusercontent.com/Cedrico59/patrimoine-arbor-V3/main/logo-marcq.png";







// =========================
// 📄 CONFIG PDF / QR
// =========================
const BASE_FICHE_URL = "https://cedrico59.github.io/patrimoine-arbor-V3"; // fiche en ligne
const PDF_MARGIN_INCH = 0.5;

// tailles police standard
const FONT_TITLE = 22;
const FONT_SUBTITLE = 19;
const FONT_SECTION = 18;
const FONT_BODY = 14;
const FONT_META = 14;
const FONT_LEGAL = 14;

// (garde le stub, utile si ton projet l'appelle)
function myFunction() {}

/*******************************
 * TEST
 *******************************/
function TEST_DRIVE_LINKED() {
  DriveApp.createFile("test_linked_drive.txt", "OK");
}

/* =========================
   📜 HISTORIQUE MODIFICATIONS
========================= */
function getOrCreateHistorySheet_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sh = ss.getSheetByName(SHEET_HISTORIQUE);
  if (!sh) {
    sh = ss.insertSheet(SHEET_HISTORIQUE);
    sh.appendRow([
      "timestamp",
      "login",
      "role",
      "secteurUser",
      "action",
      "treeId",
      "details"
    ]);
  }
  return sh;
}

function logHistory_(meta, action, treeId, detailsObj) {
  try {
    const hist = getOrCreateHistorySheet_();
    hist.appendRow([
      new Date(),
      (meta && meta.login) || "",
      (meta && meta.role) || "",
      (meta && meta.secteur) || "",
      action || "",
      treeId || "",
      JSON.stringify(detailsObj || {})
    ]);
  } catch (e) {
    Logger.log("Historique erreur: " + e);
  }
}

// récupère la ligne d’un arbre (avant modif) pour faire un diff
function getTreeRowAsObject_(sheet, treeId) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;

  const values = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    if (String(row[1]).trim() === String(treeId).trim()) {
      return {
        _rowIndex: i + 2,
        id: row[1],
        lat: row[2],
        lng: row[3],
        species: row[4],
        height: row[5],
        dbh: row[6],
        secteur: row[7],
        address: row[8],
        tags: row[9],
        historiqueInterventions: row[10],
        comment: row[11],
        photos: row[12],
        etat: row[13],
        updatedAt: row[14]
      };
    }
  }
  return null;
}

// diff simple avant/après
function diffObjects_(before, after) {
  if (!before) return [{ field: "__NEW__", from: null, to: (after && after.id) || "" }];

  const keys = [
    "lat","lng","species","height","dbh","secteur","address",
    "tags","historiqueInterventions","comment","photos","etat"
  ];
  const changes = [];

  keys.forEach(k => {
    const a = before[k];
    const b = after[k];
    const sa = (a === null || a === undefined) ? "" : String(a);
    const sb = (b === null || b === undefined) ? "" : String(b);
    if (sa !== sb) changes.push({ field: k, from: a, to: b });
  });

  return changes;
}

/* =========================
   🔐 AUTH MULTI-COMPTES
========================= */
const USERS = {
  admin: { password: "marcq2026", role: "admin", secteur: "" },

  "Hautes Loges - Briqueterie": { password: "HLB2026", role: "secteur", secteur: "Hautes Loges - Briqueterie" },
  "Bourg": { password: "BOURG2026", role: "secteur", secteur: "Bourg" },
  "Buisson - Delcencerie": { password: "BD2026", role: "secteur", secteur: "Buisson - Delcencerie" },
  "Mairie - Quesne": { password: "MQ2026", role: "secteur", secteur: "Mairie - Quesne" },
  "Pont - Plouich - Clémenceau": { password: "PPC2026", role: "secteur", secteur: "Pont - Plouich - Clémenceau" },
  "Cimetière Delcencerie": { password: "CD2026", role: "secteur", secteur: "Cimetière Delcencerie" },
  "Cimetière Pont": { password: "CP2026", role: "secteur", secteur: "Cimetière Pont" },
  "Hippodrome": { password: "HIP2026", role: "secteur", secteur: "Hippodrome" },
  "Ferme aux Oies": { password: "FAO2026", role: "secteur", secteur: "Ferme aux Oies" }
};

const TOKEN_STORE = PropertiesService.getScriptProperties();
const TOKEN_TTL_MS = 1000 * 60 * 60 * 12; // 12h

function createToken_() {
  const token = Utilities.getUuid();
  TOKEN_STORE.setProperty(token, String(Date.now()));
  return token;
}

function setTokenMeta_(token, meta) {
  if (!token || !meta) return;
  TOKEN_STORE.setProperty("meta_" + token, JSON.stringify(meta));
}

function getTokenMeta_(token) {
  if (!token) return null;
  const raw = TOKEN_STORE.getProperty("meta_" + token);
  if (!raw) return null;
  try { return JSON.parse(raw); } catch (e) { return null; }
}

function isValidToken_(token) {
  if (!token) return false;
  const ts = TOKEN_STORE.getProperty(token);
  if (!ts) return false;

  const age = Date.now() - Number(ts);
  if (!Number.isFinite(age) || age > TOKEN_TTL_MS) {
    TOKEN_STORE.deleteProperty(token);
    TOKEN_STORE.deleteProperty("meta_" + token);
    return false;
  }
  return true;
}

function authFail_() {
  return jsonResponse({ ok: false, error: "unauthorized" });
}

/* =========================
   GET – ROUTER
========================= */
function doGet(e) {
  const token = e && e.parameter && e.parameter.token;
  if (!isValidToken_(token)) return authFail_();

  if (e && e.parameter && e.parameter.action === "history") {
    return handleHistoryGet_(e);
  }

  return handleTreesAndTravauxGet_();
}

/* =========================
   📜 HISTORIQUE – GET
   GET?action=history&id=XXX&limit=50
========================= */
function handleHistoryGet_(e) {
  const treeId = String((e && e.parameter && e.parameter.id) || "").trim();
  const limit = Number((e && e.parameter && e.parameter.limit) || 50);

  if (!treeId) return jsonResponse({ ok: false, error: "id manquant" });

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const hist = ss.getSheetByName(SHEET_HISTORIQUE);
  if (!hist) return jsonResponse({ ok: true, history: [] });

  const last = hist.getLastRow();
  if (last < 2) return jsonResponse({ ok: true, history: [] });

  const rows = hist.getRange(2, 1, last - 1, hist.getLastColumn()).getValues();

  const out = [];
  for (let i = rows.length - 1; i >= 0; i--) {
    if (String(rows[i][5]).trim() === treeId) {
      out.push({
        timestamp: rows[i][0],
        login: rows[i][1],
        role: rows[i][2],
        secteurUser: rows[i][3],
        action: rows[i][4],
        treeId: rows[i][5],
        details: rows[i][6]
      });
      if (out.length >= limit) break;
    }
  }

  return jsonResponse({ ok: true, history: out });
}

/* =========================
   🌳 ARBRES + 🔧 TRAVAUX – GET
========================= */
function handleTreesAndTravauxGet_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_PATRIMOINE);
  const sheetTravaux = ss.getSheetByName(SHEET_TRAVAUX);

  if (!sheet) return jsonResponse({ ok: false, error: "SHEET_PATRIMOINE_INTROUVABLE" });

  /* ===== LECTURE TRAVAUX ===== */
  const travauxMap = {};
  if (sheetTravaux) {
    const lastT = sheetTravaux.getLastRow();
    if (lastT > 1) {
      const valuesT = sheetTravaux
        .getRange(2, 1, lastT - 1, sheetTravaux.getLastColumn())
        .getValues();

      valuesT.forEach(r => {
        const treeId = String(r[0]).trim();
        if (!treeId) return;

        travauxMap[treeId] = {
          etat: r[1] || "",
          secteur: r[2] || "",
          dateDemande: formatDateForInput(r[3]),
          natureTravaux: r[4] || "",
          address: r[5] || "",
          species: r[6] || "",
          dateDemandeDevis: formatDateForInput(r[7]),
          devisNumero: r[8] || "",
          montantDevis: r[9] || "",
          dateExecution: formatDateForInput(r[10]),
          remarquesTravaux: r[11] || "",
          numeroBDC: r[12] || "",
          numeroFacture: r[13] || ""
        };
      });
    }
  }

  /* ===== LECTURE ARBRES ===== */
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return ContentService.createTextOutput("[]")
      .setMimeType(ContentService.MimeType.JSON);
  }

  const values = sheet
    .getRange(2, 1, lastRow - 1, sheet.getLastColumn())
    .getValues();

  const trees = values.map(row => {
    const lat = Number(row[2]);
    const lng = Number(row[3]);
    const id = row[1];
    const travaux = travauxMap[String(id).trim()] || {};

    return {
      createdAt: row[0] && row[0].getTime ? row[0].getTime() : null,
      id,
      lat,
      lng,
      species: row[4],
      height: row[5] !== "" ? Number(row[5]) : null,
      dbh: row[6] !== "" ? Number(row[6]) : null,
      secteur: row[7],
      address: row[8],
      tags: row[9] ? String(row[9]).split(",").map(s => s.trim()).filter(Boolean) : [],
      historiqueInterventions: row[10] || "",
      comment: row[11],
      photos: (() => {
        if (!row[12]) return [];
        try { return JSON.parse(row[12]); } catch (e) { return []; }
      })(),
      etat: String(row[13] || "").trim(),
      updatedAt: row[14] ? Number(row[14]) : null,

      // ✅ TRAVAUX RENVOYÉS À L’APP
      secteurTravaux: travaux.secteur || "",
      dateDemande: travaux.dateDemande || "",
      natureTravaux: travaux.natureTravaux || "",
      dateDemandeDevis: travaux.dateDemandeDevis || "",
      devisNumero: travaux.devisNumero || "",
      montantDevis: travaux.montantDevis || "",
      dateExecution: travaux.dateExecution || "",
      remarquesTravaux: travaux.remarquesTravaux || "",
      numeroBDC: travaux.numeroBDC || "",
      numeroFacture: travaux.numeroFacture || ""
    };
  }).filter(t => t.id && Number.isFinite(t.lat) && Number.isFinite(t.lng));

  return ContentService
    .createTextOutput(JSON.stringify(trees))
    .setMimeType(ContentService.MimeType.JSON);
}

/* =========================
   DRIVE / PHOTOS
========================= */
// 📁 1 dossier par arbre
function getOrCreateTreeFolder(treeId) {
  const root = DriveApp.getFolderById(DRIVE_FOLDER_ID);
  const folders = root.getFoldersByName(treeId);
  return folders.hasNext() ? folders.next() : root.createFolder(treeId);
}

// 📸 upload photo base64 → Drive
function uploadPhoto(base64, filename, treeId) {
  if (!base64 || !base64.startsWith("data:")) return null;

  const folder = getOrCreateTreeFolder(treeId);
  const match = base64.match(/^data:(.*);base64,/);
  if (!match) return null;

  const contentType = match[1];
  const bytes = Utilities.base64Decode(base64.split(",")[1]);
  const blob = Utilities.newBlob(bytes, contentType, filename);

  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  return {
    driveId: file.getId(), // ⭐ CRITIQUE
    url: file.getUrl(),
    name: filename,
    addedAt: Date.now()
  };
}

function deletePhotoFromDrive(driveId) {
  try {
    if (!driveId) return false;
    DriveApp.getFileById(driveId).setTrashed(true);
    return true;
  } catch (e) {
    Logger.log("Erreur suppression photo Drive: " + e);
    return false;
  }
}

function deleteTreeFolder(treeId) {
  const root = DriveApp.getFolderById(DRIVE_FOLDER_ID);
  const folders = root.getFoldersByName(treeId);
  while (folders.hasNext()) {
    folders.next().setTrashed(true);
  }
}

function assertSheetAlive() {
  const file = DriveApp.getFileById(SPREADSHEET_ID);
  if (file.isTrashed()) {
    throw new Error("❌ Le Spreadsheet est dans la corbeille !");
  }
}

/* =========================
   POST – LOGIN / CREATE / UPDATE / DELETE / EXPORT / VALIDATE
========================= */
function doPost(e) {
  try {
    // 🔐 LOGIN (action=login & login=... & password=...)
    const actionParam = e && e.parameter && e.parameter.action;
    if (actionParam === "login") {
      const login = String((e && e.parameter && e.parameter.login) || "").trim();
      const pwd = String((e && e.parameter && e.parameter.password) || "");

      const user = USERS[login];
      if (!user || pwd !== user.password) return authFail_();

      const token = createToken_();
      setTokenMeta_(token, { role: user.role, secteur: user.secteur || "", login });

      return jsonResponse({ ok: true, token, role: user.role, secteur: user.secteur || "", login });
    }

    // 🔐 AUTH obligatoire pour tout le reste
    const token = e && e.parameter && e.parameter.token;
    if (!isValidToken_(token)) return authFail_();

    // ✅ META pour historique
    const meta = getTokenMeta_(token); // {role, secteur, login}

    // ✅ parse payload
    let data = {};
    if (e && e.parameter && Object.keys(e.parameter).length) {
      if (e.parameter.payload) {
        data = JSON.parse(e.parameter.payload);
      } else {
        data = { ...e.parameter };
      }
    } else if (e && e.postData && e.postData.contents) {
      data = JSON.parse(e.postData.contents);
    } else {
      throw new Error("Aucun payload reçu");
    }

    if (data && data.payload) data = data.payload;

    // retirer champs sensibles pour éviter effets de bord
    if (data && typeof data === "object") {
      delete data.token;
      delete data.password;
    }

    // =========================
    // 📄 EXPORT PDF (ADMIN UNIQUEMENT)
    // =========================
    if (data.action === "exportArbrePDF" && data.id) {
      const out = exportHistoriqueArbreToPDF_(String(data.id).trim(), meta);
      return jsonResponse(out);
    }
    if (data.action === "exportAnnuelPDF" && data.year) {
      const out = exportHistoriqueAnnuelToPDF_(Number(data.year), meta);
      return jsonResponse(out);
    }
   if (data.action === "exportSurveillancePDF") {
  const out = exportArbresASurveillerPDF_(meta);
  return jsonResponse(out);
}

if (data.action === "exportAbattagesPDF") {
  const out = exportAbattagesPDF_(meta);
  return jsonResponse(out);
}

if (data.action === "exportElagagesPDF") {
  const out = exportElagagesPDF_(meta);
  return jsonResponse(out);
}

    /* ===== VALIDATION INTERVENTION ===== */
    if (data.action === "validateIntervention" && data.id && data.intervention) {
      const sheetVI = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_PATRIMOINE);
      if (!sheetVI) return ok({ status: "SHEET_NOT_FOUND" });

      const lastVI = sheetVI.getLastRow();
      if (lastVI > 1) {
        const rowsVI = sheetVI.getRange(2, 1, lastVI - 1, sheetVI.getLastColumn()).getValues();
        for (let i = 0; i < rowsVI.length; i++) {
          if (String(rowsVI[i][1]).trim() === String(data.id).trim()) {
            const rowIndex = i + 2;
            const existing = String(rowsVI[i][10] || "").trim(); // col 11 Historique
            const sep = existing ? "\n" : "";
            const value = existing + sep + String(data.intervention);
            sheetVI.getRange(rowIndex, 11).setValue(value);
            sheetVI.getRange(rowIndex, 15).setValue(Date.now());
            SpreadsheetApp.flush();

            logHistory_(meta, "VALIDATE_INTERVENTION", data.id, { added: data.intervention });
            return ok({ status: "INTERVENTION_ADDED" });
          }
         
        }
      }
      return ok({ status: "NOT_FOUND" });
    }

    // 🔒 SÉCURITÉ SECTEUR : force le secteur du compte
    if (meta && String(meta.role || "").toLowerCase() === "secteur") {
      data.secteur = meta.secteur || data.secteur || "";
    }

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_PATRIMOINE);
    if (!sheet) throw new Error("SHEET_PATRIMOINE_INTROUVABLE");

    const lastRow = sheet.getLastRow();

    /* ===== SUPPRESSION PHOTO ===== */
    if (data.action === "deletePhoto" && data.photoDriveId && data.treeId) {
      logHistory_(meta, "DELETE_PHOTO", data.treeId, { photoDriveId: data.photoDriveId });

      const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
      for (let i = 0; i < rows.length; i++) {
        const sheetTreeId = String(rows[i][1]).trim();
        if (sheetTreeId === String(data.treeId).trim()) {
          let photos = [];
          try { photos = rows[i][12] ? JSON.parse(rows[i][12]) : []; } catch (err) { photos = []; }

          deletePhotoFromDrive(String(data.photoDriveId).trim());

          const newPhotos = photos.filter(p =>
            String(p.driveId || "").trim() !== String(data.photoDriveId).trim()
          );

          sheet.getRange(i + 2, 13).setValue(JSON.stringify(newPhotos)); // col 13 photos
          SpreadsheetApp.flush();

          return ok({ status: "PHOTO_DELETED", remaining: newPhotos.length });
        }
      }
      return ok({ status: "NOT_FOUND" });
    }

    /* ===== SUPPRESSION ARBRE ===== */
    if (data.action === "delete" && data.id) {
      if (lastRow < 2) return ok({ status: "NOT_FOUND" });

      const beforeObjDelete = getTreeRowAsObject_(sheet, data.id);
      logHistory_(meta, "DELETE", data.id, { deletedRow: beforeObjDelete || null });

      const rows = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
      for (let i = 0; i < rows.length; i++) {
        if (String(rows[i][1]).trim() === String(data.id).trim()) {
          deleteTreeFolder(String(data.id).trim());
          sheet.deleteRow(i + 2);
          sortArbresSheet_(sheet);
          SpreadsheetApp.flush();
          return ok({ status: "DELETED" });
        }
      }
      return ok({ status: "NOT_FOUND" });
    }

    // ✅ create/update -> id obligatoire
    if (!data.id) throw new Error("id manquant (create/update)");

    // ✅ conversions si on est passé par e.parameter (tout est string)
    if (typeof data.tags === "string") {
      try { data.tags = JSON.parse(data.tags); }
      catch (err) { data.tags = String(data.tags).split(",").map(s => s.trim()).filter(Boolean); }
    }
    if (typeof data.photos === "string") {
      try { data.photos = JSON.parse(data.photos); }
      catch (err) { data.photos = []; }
    }

    // ✅ HISTORIQUE : état avant update/create
    const beforeObj = getTreeRowAsObject_(sheet, data.id);

    /* ===== PHOTOS EXISTANTES ===== */
    let existingPhotos = [];
    if (lastRow > 1) {
      const rows = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
      for (let i = 0; i < rows.length; i++) {
        if (String(rows[i][1]).trim() === String(data.id).trim()) {
          try { existingPhotos = rows[i][12] ? JSON.parse(rows[i][12]) : []; } catch (err) { existingPhotos = []; }
          break;
        }
      }
    }

    /* ===== NOUVELLES PHOTOS ===== */
    let uploadedPhotos = [];
    if (Array.isArray(data.photos)) {
      uploadedPhotos = data.photos
        .map(p => uploadPhoto(
          p.dataUrl,
          `${Date.now()}_${p.name || "photo.jpg"}`,
          data.id
        ))
        .filter(Boolean);
    }

    const allPhotos = existingPhotos.concat(uploadedPhotos);

    /* ===== DONNÉES ===== */
    const rowData = [
      new Date(),                          // A createdAt
      data.id || "",                       // B id
      data.lat || "",                      // C lat
      data.lng || "",                      // D lng
      data.species || "",                  // E species
      data.height || "",                   // F height
      data.dbh || "",                      // G dbh
      data.secteur || "",                  // H secteur
      data.address || "",                  // I address
      (data.tags || []).join(","),         // J tags
      data.historiqueInterventions || "",  // K historique
      data.comment || "",                  // L comment
      JSON.stringify(allPhotos),           // M photos
      data.etat || "",                     // N etat
      data.updatedAt || Date.now()         // O updatedAt
    ];

    let isUpdate = false;

    /* ===== UPDATE ===== */
    if (lastRow > 1) {
      const ids = sheet.getRange(2, 2, lastRow - 1, 1).getValues(); // col B
      for (let i = 0; i < ids.length; i++) {
        if (String(ids[i][0]).trim() === String(data.id).trim()) {
          sheet.getRange(i + 2, 1, 1, rowData.length).setValues([rowData]);

          // ✅ tri après mise à jour
          sortArbresSheet_(sheet);

          // ✅ recolor fiable par ID (après tri)
          recolorOneArbreById_(sheet, data.id);

          isUpdate = true;
          break;
        }
      }
    }

    /* ===== CREATE ===== */
    if (!isUpdate) {
      sheet.appendRow(rowData);
      sortArbresSheet_(sheet);
      recolorOneArbreById_(sheet, data.id);
    }

    /* ===== TRAVAUX (sync si état déclencheur) ===== */
    syncTravaux_(data);
   
    
    SpreadsheetApp.flush();

    // ✅ HISTORIQUE : état après + diff
    const afterObj = {
      id: data.id,
      lat: data.lat || "",
      lng: data.lng || "",
      species: data.species || "",
      height: data.height || "",
      dbh: data.dbh || "",
      secteur: data.secteur || "",
      address: data.address || "",
      tags: (data.tags || []).join(","),
      historiqueInterventions: data.historiqueInterventions || "",
      comment: data.comment || "",
      photos: JSON.stringify(allPhotos || []),
      etat: data.etat || ""
    };

    const changes = diffObjects_(beforeObj, afterObj);
    logHistory_(meta, isUpdate ? "UPDATE" : "CREATE", data.id, { changes });

    return ok({ status: isUpdate ? "UPDATED" : "CREATED", photos: allPhotos });

  } catch (err) {
    return jsonResponse({ ok: false, error: String(err) });
  }
   // 👇 ICI, TOUT À LA FIN
return jsonResponse({
  ok: false,
  error: "Action inconnue"
});
}

/* =========================
   SYNC TRAVAUX
========================= */
function syncTravaux_(data) {
  const etatArbre = String(data.etat || "").trim();
  const ETATS_TRAVAUX = [
    "Dangereux (A abattre)",
    "A surveiller",
    "A élaguer (URGENT)",
    "A élaguer (Moyen)",
    "A élaguer (Faible)"
  ];
  const doitAllerTravaux = ETATS_TRAVAUX.includes(etatArbre);
  if (!doitAllerTravaux) return;

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheetTravaux = ss.getSheetByName(SHEET_TRAVAUX);
  if (!sheetTravaux) return;

  const travauxRow = [
    data.id || "",                    // A - Id
    etatArbre || "",                  // B - État de l’arbre
    data.secteur || "",               // C - Secteur
    data.dateDemande || "",           // D - Date de demande
    data.natureTravaux || "",         // E - Nature des travaux
    data.address || "",               // F - Adresse des travaux
    data.species || "",               // G - Espèce
    data.dateDemandeDevis || "",      // H - Date de demande de devis
    data.devisNumero || "",           // I - Devis n°
    data.montantDevis || "",          // J - Montant du devis (€)
    data.dateExecution || "",         // K - Date d’exécution des travaux
    data.remarquesTravaux || "",      // L - Remarques
    data.numeroBDC || "",             // M - N° bdc
    data.numeroFacture || ""          // N - N° Facture
  ];

  const lastTravaux = sheetTravaux.getLastRow();
  let foundTravaux = false;

  if (lastTravaux > 1) {
    const idsTravaux = sheetTravaux.getRange(2, 1, lastTravaux - 1, 1).getValues();
    for (let i = 0; i < idsTravaux.length; i++) {
      if (String(idsTravaux[i][0]).trim() === String(data.id).trim()) {
        const rowIndex = i + 2;
        sheetTravaux.getRange(rowIndex, 1, 1, travauxRow.length).setValues([travauxRow]);

        colorEtatTravaux(sheetTravaux, rowIndex, etatArbre);

        // ✅ tri désactivé (comme avant)
        sortTravauxSheet_(sheetTravaux);

        // ✅ recolor fiable par ID
        recolorOneTravauxById_(sheetTravaux, data.id);

        foundTravaux = true;
        break;
      }
    }
  }

  if (!foundTravaux) {
    sheetTravaux.appendRow(travauxRow);

    sortTravauxSheet_(sheetTravaux);

    const newRow = sheetTravaux.getLastRow();
    colorEtatTravaux(sheetTravaux, newRow, etatArbre);
    recolorOneTravauxById_(sheetTravaux, data.id);
  }
    /* =========================
     FEUILLES MÉTIER (AJOUT ICI)
  ========================= */
  syncSurveillance_(data);
  syncAbattages_(data);
  syncElagages_(data);
}

/* =========================
   UTIL JSON
========================= */
function ok(payload) {
  return jsonResponse({ ok: true, result: payload });
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function formatDateForInput(d) {
  if (!d) return "";
  if (Object.prototype.toString.call(d) !== "[object Date]") return "";
  if (isNaN(d.getTime())) return "";

  const yyyy = d.getFullYear();
  const mm = String(d.getMonth() + 1).padStart(2, "0");
  const dd = String(d.getDate()).padStart(2, "0");
  return `${yyyy}-${mm}-${dd}`;
}

/* =========================
   TRI & COULEURS
========================= */
// Secteur (col 8) -> Adresse (col 9) -> Espèce (col 5)
function sortArbresSheet_(sheet) {
  try {
    const lastRow = sheet.getLastRow();
    if (lastRow <= 2) return;

    sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).sort([
      { column: 8, ascending: true }, // secteur
      { column: 9, ascending: true }, // adresse (rue)
      { column: 5, ascending: true }  // espèce
    ]);
  } catch (e) {
    Logger.log("Tri arbres erreur: " + e);
  }
}

// tri travaux désactivé pour éviter propagation style
function sortTravauxSheet_(sheetTravaux) {
  return;
}

function colorRowByEtat(sheet, rowIndex, etat) {
  let color = null;

  if (etat === "Dangereux (A abattre)") color = "#f28b82";
  if (etat === "A surveiller") color = "#fbbc04";
  if (etat === "A élaguer (URGENT)") color = "#FFFF00";
  if (etat === "A élaguer (Moyen)") color = "#00FFFF";
  if (etat === "A élaguer (Faible)") color = "#ccff90";

  const range = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn());
  if (color) range.setBackground(color);
  else range.setBackground(null);
}

function colorEtatTravaux(sheet, rowIndex, etat) {
  let color = null;

  if (etat === "Dangereux (A abattre)") color = "#f28b82";
  if (etat === "A surveiller") color = "#fbbc04";
  if (etat === "A élaguer (URGENT)") color = "#FFFF00";
  if (etat === "A élaguer (Moyen)") color = "#00FFFF";
  if (etat === "A élaguer (Faible)") color = "#ccff90";

  const cell = sheet.getRange(rowIndex, 2); // col B
  if (color) {
    cell.setBackground(color);
    cell.setFontWeight("bold");
  } else {
    cell.setBackground(null);
    cell.setFontWeight("normal");
  }
}

// recolor travaux (optionnel)
function recolorEtatTravauxColumn_(sheetTravaux) {
  const lastRow = sheetTravaux.getLastRow();
  if (lastRow < 2) return;

  const etats = sheetTravaux.getRange(2, 2, lastRow - 1, 1).getValues();
  for (let i = 0; i < etats.length; i++) {
    const rowIndex = i + 2;
    const etat = String(etats[i][0] || "").trim();
    colorEtatTravaux(sheetTravaux, rowIndex, etat);
  }
}

// recolor arbres (optionnel)
function recolorArbresRows_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const etats = sheet.getRange(2, 13, lastRow - 1, 1).getValues(); // col 13
  for (let i = 0; i < etats.length; i++) {
    const rowIndex = i + 2;
    const etat = String(etats[i][0] || "").trim();
    colorRowByEtat(sheet, rowIndex, etat);
  }
}

// couleur travaux par ID (fiable)
function recolorTravauxById_(sheetTravaux) {
  const lastRow = sheetTravaux.getLastRow();
  if (lastRow < 2) return;

  const rows = sheetTravaux.getRange(2, 1, lastRow - 1, 2).getValues(); // A,B
  for (let i = 0; i < rows.length; i++) {
    const rowIndex = i + 2;
    const treeId = String(rows[i][0] || "").trim();
    const etat = String(rows[i][1] || "").trim();
    if (!treeId) continue;
    colorEtatTravaux(sheetTravaux, rowIndex, etat);
  }
}

// couleur arbres par ID (fiable)
function recolorArbresById_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const ids = sheet.getRange(2, 2, lastRow - 1, 1).getValues();   // col B
  const etats = sheet.getRange(2, 13, lastRow - 1, 1).getValues(); // col 13
  for (let i = 0; i < ids.length; i++) {
    const rowIndex = i + 2;
    const id = String(ids[i][0] || "").trim();
    const etat = String(etats[i][0] || "").trim();
    if (!id) continue;
    colorRowByEtat(sheet, rowIndex, etat);
  }
}

function recolorOneArbreById_(sheet, treeId) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const ids = sheet.getRange(2, 2, lastRow - 1, 1).getValues(); // col B
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]).trim() === String(treeId).trim()) {
      const rowIndex = i + 2;
      const etat = String(sheet.getRange(rowIndex, 13).getValue() || "").trim(); // col 13
      colorRowByEtat(sheet, rowIndex, etat);
      return;
    }
  }
}

function recolorOneTravauxById_(sheetTravaux, treeId) {
  const lastRow = sheetTravaux.getLastRow();
  if (lastRow < 2) return;

  const ids = sheetTravaux.getRange(2, 1, lastRow - 1, 1).getValues(); // col A
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]).trim() === String(treeId).trim()) {
      const rowIndex = i + 2;
      const etat = String(sheetTravaux.getRange(rowIndex, 2).getValue() || "").trim(); // col B
      colorEtatTravaux(sheetTravaux, rowIndex, etat);
      return;
    }
  }
}


// =========================
// 📏 UTIL PDF — hauteur de ligne selon texte
// =========================
function estimateRowHeight_(text, baseHeight, pxPerLine) {
  const t = String(text || "");
  const lines = Math.max(1, t.split("\n").length);
  const h = Math.round(Math.max(baseHeight, lines * pxPerLine));
  return h;
}

/* =========================
   🔐 EXPORT PDF (ADMIN UNIQUEMENT)
========================= */
function assertAdmin_(meta) {
  if (!meta || String(meta.role || "").toLowerCase() !== "admin") {
    throw new Error("ADMIN_ONLY");
  }
}

function writeCoverPage_(sheet, title, meta) {
  sheet.clear();

  // 📄 Mise en page A4 + impression
  sheet.setHiddenGridlines(true);
  try { sheet.setPageOrientation(SpreadsheetApp.PageOrientation.PORTRAIT); } catch (e) {}
  try { sheet.setPaperSize(SpreadsheetApp.PaperSize.A4); } catch (e) {}
  try {
    sheet.setMargins({
      top: PDF_MARGIN_INCH,
      bottom: PDF_MARGIN_INCH,
      left: PDF_MARGIN_INCH,
      right: PDF_MARGIN_INCH
    });
  } catch (e) {}

  // Largeurs pensées pour un PDF A4 portrait (A→E)
  sheet.setColumnWidths(1, 1, 255); // A ID / zone gauche
  sheet.setColumnWidths(2, 1, 128); // B
  sheet.setColumnWidths(3, 1, 160); // C
  sheet.setColumnWidths(4, 1, 200); // D
  sheet.setColumnWidths(5, 1, 420); // E (texte long)

  // Hauteurs de base (on ajuste ensuite les zones clés)
  sheet.setRowHeights(1, 60, 28);

 // 🖼️ LOGO VILLE — insertion réelle (PDF-safe)
if (MAIRIE_LOGO_URL) {
  try {
    const logoBlob = UrlFetchApp.fetch(MAIRIE_LOGO_URL).getBlob();
    sheet.insertImage(logoBlob, 1, 1) // colonne A, ligne 1
      .setWidth(180)
      .setHeight(180);
  } catch (e) {
    Logger.log("Logo PDF erreur: " + e);
  }
}


  // 🏛️ TITRES OFFICIELS
  sheet.getRange("C1")
    .setValue("   VILLE DE MARCQ-EN-BARŒUL")
    .setFontSize(FONT_TITLE)
    .setFontWeight("bold");

  sheet.getRange("C3")
    .setValue("Gestion du patrimoine arboré communal")
    .setFontSize(FONT_SUBTITLE)
    .setFontWeight("bold");

  sheet.getRange("B5")
    .setValue(           title)
    .setFontSize(FONT_SECTION)
    .setFontWeight("bold");

  // 📜 TEXTE RÉGLEMENTAIRE (fusion A8 → E12)
  const regRange = sheet.getRange("A8:E12");
  regRange.merge()
    .setValue(
      "DOCUMENT ADMINISTRATIF OFFICIEL\n\n" +
      "Ce document est généré automatiquement par le système d’information de la Ville.\n" +
      "Il constitue une extraction fidèle et figée des données enregistrées à la date indiquée.\n" +
      "Toute modification ultérieure des données sources n’affecte pas le présent document."
    )
    .setFontSize(FONT_LEGAL)
    .setFontStyle("italic")
    .setWrap(true)
    .setVerticalAlignment("top");

  // un peu plus d'air pour le texte réglementaire
  try { sheet.setRowHeights(8, 5, 40); } catch (e) {}

  // 📅 MÉTADONNÉES (sous le texte réglementaire)
  sheet.getRange("A14")
    .setValue(
      "Date de génération : " +
      Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm")
    )
    .setFontSize(FONT_META)
    .setFontStyle("italic");

  sheet.getRange("A16")
    .setValue("Généré par : " + ((meta && meta.login) ? meta.login : "Administrateur"))
    .setFontSize(FONT_META)
    .setFontStyle("italic");

  // ✍️ SIGNATURE (placée sous photo/QR, fusion A34 → E38)
  const signRange = sheet.getRange("A37:E41");
  signRange.merge()
    .setValue(
      "Service : Espaces verts \n\n" +
      "Responsable : ____________________________\n\n" +
      "Signature : ______________________________"
    )
    .setFontSize(FONT_BODY)
    .setFontWeight("bold")
    .setWrap(true)
    .setVerticalAlignment("top");

  try { sheet.setRowHeights(34, 5, 36); } catch (e) {}

  // séparateur discret
  sheet.getRange("A40").setValue("—");
}

function exportHistoriqueArbreToPDF_(treeId, meta) {
  assertAdmin_(meta);

  const id = String(treeId || "").trim();
  if (!id) throw new Error("ID_MANQUANT");

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const source = ss.getSheetByName(SHEET_PATRIMOINE);
  if (!source) throw new Error("PATRIMOINE_ARBORE_INTROUVABLE");

  const tmpName = "TMP_EXPORT_ARBRE";
  const old = ss.getSheetByName(tmpName);
  if (old) ss.deleteSheet(old);
  const tmp = ss.insertSheet(tmpName);

  writeCoverPage_(tmp, `Historique des travaux – Arbre ${id}`, meta);

  const lastRow = source.getLastRow();
  if (lastRow < 2) throw new Error("PATRIMOINE_VIDE");

  const data = source.getRange(2, 1, lastRow - 1, source.getLastColumn()).getValues();

  // =========================
  // 🌳 Récup arbre + 📷 photo + QR
  // =========================
  let treeObj = null;
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const rowId = String(row[1]).trim();
    if (rowId === id) {
      let photos = [];
      try { photos = row[12] ? JSON.parse(row[12]) : []; } catch (e) { photos = []; }
      treeObj = {
        id: rowId,
        species: row[4] || "",
        secteur: row[7] || "",
        address: row[8] || "",
        historique: String(row[10] || ""),
        photos
      };
      break;
    }
  }
  if (!treeObj) {
    ss.deleteSheet(tmp);
    throw new Error("ARBRE_INTROUVABLE");
  }

  // 📷 Photo (première photo Drive trouvée) — bloc gauche
  if (Array.isArray(treeObj.photos) && treeObj.photos.length > 0) {
    const p = treeObj.photos.find(x => x && x.driveId);
    if (p && p.driveId) {
      try {
        const photoBlob = DriveApp.getFileById(String(p.driveId)).getBlob();
        // position : A24
        tmp.insertImage(photoBlob, 1, 22).setWidth(560).setHeight(400);
        try { tmp.setRowHeight(24, 280); } catch (e) {}
      } catch (e) {
        Logger.log("Photo PDF: " + e);
      }
    }
  }

  // 📱 QR Code — bloc droit (E24) + libellé sous le QR
  if (BASE_FICHE_URL) {
    const ficheUrl = `${BASE_FICHE_URL}?id=${encodeURIComponent(treeObj.id)}`;
    const qrGenUrl =
      "https://api.qrserver.com/v1/create-qr-code/" +
      "?size=600x600&data=" + encodeURIComponent(ficheUrl);

    try {
      const qrBlob = UrlFetchApp.fetch(qrGenUrl).getBlob();
      tmp.insertImage(qrBlob, 5, 24).setWidth(150).setHeight(150);
      tmp.getRange("E30")
        .setValue("Scanner pour consulter\nla fiche en ligne")
        .setFontSize(FONT_META)
        .setHorizontalAlignment("center")
        .setVerticalAlignment("top")
        .setWrap(true);
    } catch (e) {
      Logger.log("QR PDF: " + e);
    }
  }

  // =========================
  // 📊 TABLEAU — zone contrôlée (évite que ça parte en bas / minuscule)
  // =========================
  const TABLE_START_ROW = 56; // ⬅️ déplace le tableau ici si besoin

  // En-tête
  tmp.getRange(`A${TABLE_START_ROW}:E${TABLE_START_ROW}`)
    .setValues([[
      "ID Arbre",
      "Espèce",
      "Secteur",
      "Adresse",
      "Historique des interventions"
    ]])
    .setFontWeight("bold")
    .setFontSize(FONT_BODY)
    .setWrap(true)
    .setVerticalAlignment("middle");

  // Ligne infos (A-D) + Historique sur la ligne suivante (A→E fusion)
  const infoRow = TABLE_START_ROW + 1;
  tmp.getRange(`A${infoRow}:D${infoRow}`)
    .setValues([[treeObj.id, treeObj.species, treeObj.secteur, treeObj.address]])
    .setFontSize(FONT_BODY)
    .setVerticalAlignment("middle")
    .setWrap(true);

  // Historique : grosse zone lisible
  const histRow = infoRow + 1;
  const histRange = tmp.getRange(`A${histRow}:E${histRow}`);
  histRange.merge()
    .setValue(treeObj.historique || "")
    .setFontSize(FONT_BODY) // ✅ même taille que le reste (plus "minuscule")
    .setWrap(true)
    .setVerticalAlignment("top");

  // hauteur auto selon nombre de lignes
  const h = estimateRowHeight_(treeObj.historique, 110, 26);
  try { tmp.setRowHeight(histRow, h); } catch (e) {}

  // un peu d'air au-dessus du tableau
  try { tmp.setRowHeights(1, TABLE_START_ROW - 1, 28); } catch (e) {}

  // si aucune intervention -> erreur
  if (!treeObj.historique) {
    ss.deleteSheet(tmp);
    throw new Error("AUCUNE_INTERVENTION_POUR_CET_ARBRE");
  }

  SpreadsheetApp.flush();

  const sheetId = tmp.getSheetId();
  const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd_HH-mm");
  const fileName = `Historique_Arbre_${id}_${now}.pdf`;

  const url =
    `https://docs.google.com/spreadsheets/d/${SPREADSHEET_ID}/export` +
    `?format=pdf` +
    `&gid=${sheetId}` +
    `&portrait=true` +
    `&fitw=true` +
    `&top_margin=${PDF_MARGIN_INCH}` +
    `&bottom_margin=${PDF_MARGIN_INCH}` +
    `&left_margin=${PDF_MARGIN_INCH}` +
    `&right_margin=${PDF_MARGIN_INCH}` +
    `&sheetnames=false` +
    `&printtitle=false` +
    `&fzr=false` +
    `&gridlines=false` +
    `&pagenumbers=true`;

  const blob = UrlFetchApp.fetch(url, {
    headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() }
  }).getBlob().setName(fileName);

  const file = DriveApp.getFolderById(DRIVE_FOLDER_ID).createFile(blob);

  ss.deleteSheet(tmp);

  return { ok: true, url: file.getUrl(), fileUrl: file.getUrl(), name: fileName };
}

function exportHistoriqueAnnuelToPDF_(year, meta) {
  assertAdmin_(meta);

  const y = Number(year);
  if (!Number.isFinite(y) || y < 2000 || y > 2100) throw new Error("ANNEE_INVALIDE");

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const source = ss.getSheetByName(SHEET_PATRIMOINE);
  if (!source) throw new Error("PATRIMOINE_ARBORE_INTROUVABLE");

  const tmpName = "TMP_EXPORT_ANNUEL";
  const old = ss.getSheetByName(tmpName);
  if (old) ss.deleteSheet(old);
  const tmp = ss.insertSheet(tmpName);

  writeCoverPage_(tmp, `                           Historique annuel des travaux – ${y}`, meta); //garder espaces devant Historique

  const lastRow = source.getLastRow();
  if (lastRow < 2) throw new Error("PATRIMOINE_VIDE");

  const data = source.getRange(2, 1, lastRow - 1, source.getLastColumn()).getValues();




  // =========================
  // 📊 TABLEAU — annuel
  // =========================
  const TABLE_START_ROW = 56;

  tmp.getRange(`A${TABLE_START_ROW}:F${TABLE_START_ROW}`)
    .setValues([[
      "ID Arbre",
      "Espèce",
      "Secteur",
      "Adresse",
      "Année",
      "Intervention"
    ]])
    .setFontWeight("bold")
    .setFontSize(FONT_BODY)
    .setWrap(true)
    .setVerticalAlignment("middle");

  // largeurs (A→F)
  tmp.setColumnWidths(1, 1, 120);
  tmp.setColumnWidths(2, 1, 140);
  tmp.setColumnWidths(3, 1, 160);
  tmp.setColumnWidths(4, 1, 220);
  tmp.setColumnWidths(5, 1, 80);
  tmp.setColumnWidths(6, 1, 420);

  let rowCursor = TABLE_START_ROW + 1;
  let count = 0;

  data.forEach(row => {
    const treeId = String(row[1]).trim();
    const historique = String(row[10] || "");
    if (!historique) return;

    historique.split("\n").forEach(line => {
      if (line.includes(String(y))) {
        tmp.getRange(rowCursor, 1, 1, 6)
          .setValues([[treeId, row[4] || "", row[7] || "", row[8] || "", y, line]])
          .setFontSize(FONT_BODY)
          .setWrap(true)
          .setVerticalAlignment("top");
        // hauteur selon texte d'intervention
        try { tmp.setRowHeight(rowCursor, estimateRowHeight_(line, 40, 24)); } catch (e) {}
        rowCursor++;
        count++;
      }
    });
  });

  if (count === 0) {
    ss.deleteSheet(tmp);
    throw new Error("AUCUNE_INTERVENTION_POUR_CETTE_ANNEE");
  }

  // un peu d'air au-dessus du tableau
  try { tmp.setRowHeights(1, TABLE_START_ROW - 1, 28); } catch (e) {}

  SpreadsheetApp.flush();

  const sheetId = tmp.getSheetId();
  const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd_HH-mm");
  const fileName = `Historique_Travaux_${y}_${now}.pdf`;

  const url =
    `https://docs.google.com/spreadsheets/d/${SPREADSHEET_ID}/export` +
    `?format=pdf` +
    `&gid=${sheetId}` +
    `&portrait=true` +
    `&fitw=true` +
    `&top_margin=${PDF_MARGIN_INCH}` +
    `&bottom_margin=${PDF_MARGIN_INCH}` +
    `&left_margin=${PDF_MARGIN_INCH}` +
    `&right_margin=${PDF_MARGIN_INCH}` +
    `&sheetnames=false` +
    `&printtitle=false` +
    `&fzr=false` +
    `&gridlines=false` +
    `&pagenumbers=true`;

  const blob = UrlFetchApp.fetch(url, {
    headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() }
  }).getBlob().setName(fileName);

  const file = DriveApp.getFolderById(DRIVE_FOLDER_ID).createFile(blob);

  ss.deleteSheet(tmp);

  return { ok: true, url: file.getUrl(), fileUrl: file.getUrl(), name: fileName, count };
}



/*************************************************
 * CONFIG FEUILLES
 *************************************************/
const SHEET_SURVEILLANCE = "Arbres_à_surveiller";
const SHEET_ABATTAGES = "Abattages";
const SHEET_ELAGAGES = "Elagages";

/*************************************************
 * COULEURS PAR ÉTAT (PASTILLES)
 *************************************************/
const ETAT_COLORS = {
  "A surveiller": "#FFE0B2",              // 🟠 orange
  "Dangereux (A abattre)": "#FFCDD2",     // 🔴 rouge
  "A élaguer (URGENT)": "#FFF9C4",        // 🟡 jaune
  "A élaguer (Moyen)": "#BBDEFB",         // 🔵 bleu
  "A élaguer (Faible)": "#C8E6C9"         // 🟢 vert
};

/*************************************************
 * UTILITAIRES
 *************************************************/
// 🔎 recherche ID colonne B (format Patrimoine)
function findRowByIdPatrimoine_(sheet, treeId) {
  const last = sheet.getLastRow();
  if (last < 2) return -1;

  const ids = sheet.getRange(2, 2, last - 1, 1).getValues(); // 👈 colonne B
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]).trim() === String(treeId).trim()) {
      return i + 2;
    }
  }
  return -1;
}

// trouver une ligne par ID (colonne A)
function findRowById_(sheet, treeId) {
  const last = sheet.getLastRow();
  if (last < 2) return -1;
  const ids = sheet.getRange(2, 1, last - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]).trim() === treeId) return i + 2;
  }
  return -1;
}

// colorer une ligne entière selon l’état
function colorRow_(sheet, rowIndex, etat) {
  const color = ETAT_COLORS[etat];
  if (!color) return;
  sheet
    .getRange(rowIndex, 1, 1, sheet.getLastColumn())
    .setBackground(color);
}

/*************************************************
 * 1️⃣ ARBRES À SURVEILLER
 * - format Patrimoine_arboré
 * - MAJ complète
 * - suppression si plus "A surveiller"
 *************************************************/
function syncSurveillance_(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_SURVEILLANCE);
  if (!sheet) return;

  const id = String(data.id || "").trim();
  const etat = String(data.etat || "").trim();
  if (!id) return;

  const rowIndex = findRowByIdPatrimoine_(sheet, id);


  // ❌ si plus à surveiller → suppression
  if (etat !== "A surveiller") {
    if (rowIndex !== -1) sheet.deleteRow(rowIndex);
    return;
  }

  // format identique à Patrimoine_arboré
  const row = [
    new Date(),                          // createdAt
    id,                                  // id
    data.lat || "",
    data.lng || "",
    data.species || "",
    data.height || "",
    data.dbh || "",
    data.secteur || "",
    data.address || "",
    (data.tags || []).join(","),
    data.historiqueInterventions || "",
    data.comment || "",
    JSON.stringify(data.photos || []),
    etat,
    Date.now()
  ];

  if (rowIndex === -1) {
    sheet.appendRow(row);
    colorRow_(sheet, sheet.getLastRow(), etat);
  } else {
    sheet.getRange(rowIndex, 1, 1, row.length).setValues([row]);
    colorRow_(sheet, rowIndex, etat);
  }
}

/*************************************************
 * 2️⃣ + 3️⃣ ABATTAGES & ÉLAGAGES
 * - format tableau_Elagages/Abattages
 * - ajout à l’entrée
 * - mise à jour UNIQUEMENT cellules vides
 * - jamais supprimé
 *************************************************/
function syncTravauxPartiel_(data, sheetName) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return;

  const id = String(data.id || "").trim();
  const etat = String(data.etat || "").trim();
  if (!id) return;

  const row = [
    id,
    etat,
    data.secteur || "",
    data.dateDemande || "",
    data.natureTravaux || "",
    data.address || "",
    data.species || "",
    data.dateDemandeDevis || "",
    data.devisNumero || "",
    data.montantDevis || "",
    data.dateExecution || "",
    data.remarquesTravaux || "",
    data.numeroBDC || "",
    data.numeroFacture || ""
  ];

  let rowIndex = findRowById_(sheet, id);

  // ➕ création
  if (rowIndex === -1) {
    sheet.appendRow(row);
    rowIndex = sheet.getLastRow();
  }
  // 🔁 MAJ partielle (seulement cellules vides)
  else {
    const range = sheet.getRange(rowIndex, 1, 1, row.length);
    const existing = range.getValues()[0];
    const merged = existing.map((cell, i) =>
      cell !== "" && cell !== null ? cell : (row[i] || "")
    );
    range.setValues([merged]);
  }

  const etatFinal = sheet.getRange(rowIndex, 2).getValue(); // colonne B
colorRow_(sheet, rowIndex, String(etatFinal || "").trim());

}
function recolorElagages_() {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_ELAGAGES);
  if (!sheet) return;

  const last = sheet.getLastRow();
  if (last < 2) return;

  const etats = sheet.getRange(2, 2, last - 1, 1).getValues();
  for (let i = 0; i < etats.length; i++) {
    colorRow_(sheet, i + 2, String(etats[i][0] || "").trim());
  }
}

/*************************************************
 * ABATTAGES
 * 🔴 Dangereux (A abattre)
 *************************************************/
function syncAbattages_(data) {
  if (String(data.etat || "").trim() !== "Dangereux (A abattre)") return;
  syncTravauxPartiel_(data, SHEET_ABATTAGES);
}

/*************************************************
 * ÉLAGAGESh
 * 🔵 🟡 🟢
 *************************************************/
function syncElagages_(data) {
  const ETATS = [
    "A élaguer (URGENT)",
    "A élaguer (Moyen)",
    "A élaguer (Faible)"
  ];
  if (!ETATS.includes(String(data.etat || "").trim())) return;
  syncTravauxPartiel_(data, SHEET_ELAGAGES);
}

/*************************************************
 * 📌 POINT DE BRANCHEMENT OBLIGATOIRE
 * À AJOUTER DANS doPost(), APRÈS syncTravaux_(data)
 *************************************************/

// syncSurveillance_(data);
// syncAbattages_(data);
// syncElagages_(data);
function exportSheetToPDF_(sheetName, filePrefix, meta) {
  assertAdmin_(meta);

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const source = ss.getSheetByName(sheetName);
  if (!source) throw new Error("SHEET_INTROUVABLE");

  const tmpName = "TMP_EXPORT_" + sheetName;
  const old = ss.getSheetByName(tmpName);
  if (old) ss.deleteSheet(old);

  const tmp = source.copyTo(ss).setName(tmpName);

  SpreadsheetApp.flush();

  const sheetId = tmp.getSheetId();
  const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd_HH-mm");
  const fileName = `${filePrefix}_${now}.pdf`;

  const url =
    `https://docs.google.com/spreadsheets/d/${SPREADSHEET_ID}/export` +
    `?format=pdf` +
    `&gid=${sheetId}` +
    `&portrait=true` +
    `&fitw=true` +
    `&sheetnames=false` +
    `&printtitle=false` +
    `&fzr=false` +
    `&gridlines=false`;

  const blob = UrlFetchApp.fetch(url, {
    headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() }
  }).getBlob().setName(fileName);

  const file = DriveApp.getFolderById(DRIVE_FOLDER_ID).createFile(blob);

  ss.deleteSheet(tmp);

  return {
    ok: true,
    fileUrl: file.getUrl(),
    name: fileName
  };
}


function exportArbresASurveillerPDF_(meta) {
  assertAdmin_(meta);
  return exportSheetToPDF_(SHEET_SURVEILLANCE, "Arbres_a_surveiller", meta);
}

function exportAbattagesPDF_(meta) {
  assertAdmin_(meta);
  return exportSheetToPDF_(SHEET_ABATTAGES, "Abattages", meta);
}

function exportElagagesPDF_(meta) {
  assertAdmin_(meta);
  return exportSheetToPDF_(SHEET_ELAGAGES, "Elagages", meta);
}

