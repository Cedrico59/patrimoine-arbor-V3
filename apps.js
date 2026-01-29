const SPREADSHEET_ID = "1Q-HaZs_nMcJRiH0lNu-NpYbdlvWFNWSK8dQPGl9vJNU";
function myFunction() {}
const SHEET_TRAVAUX = "tableau_Elagages/Abattages";
function TEST_DRIVE_LINKED() {
  DriveApp.createFile("test_linked_drive.txt", "OK");
}

/* =========================
   📜 HISTORIQUE MODIFICATIONS (AJOUT)
========================= */
const SHEET_HISTORIQUE = "Historique";

// crée l'onglet Historique s'il n'existe pas
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

// écrit une ligne d'historique
function logHistory_(meta, action, treeId, detailsObj) {
  try {
    const hist = getOrCreateHistorySheet_();
    hist.appendRow([
      new Date(),
      meta?.login || "",
      meta?.role || "",
      meta?.secteur || "",
      action,
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
  if (!before) return [{ field: "__NEW__", from: null, to: after?.id || "" }];

  const keys = ["lat","lng","species","height","dbh","secteur","address","tags","historiqueInterventions","comment","photos","etat"];
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
   🔐 AUTH (AJOUT)
========================= */
// =========================
// 🔐 AUTH MULTI-COMPTES
// =========================
// ✅ Admin : accès total
// ✅ Secteur : accès limité (filtrage côté front)
// ⚠️ Ici on ne change que la connexion / token

const USERS = {
  admin: { password: "marcq2026", role: "admin", secteur: "" },

  // 🔧 Remplace les mots de passe ci-dessous
  // Chaque secteur a son propre login + mot de passe
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
  try { return JSON.parse(raw); } catch { return null; }
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
   GET – ROUTER (CORRIGÉ: un seul doGet)
========================= */
function doGet(e) {
  // 🔐 AUTH
  const token = e?.parameter?.token;
  if (!isValidToken_(token)) return authFail_();

  // 📜 HISTORIQUE : GET?action=history&id=XXX
  if (e?.parameter?.action === "history") {
    return handleHistoryGet_(e);
  }

  // 🌳 ARBRES + 🔧 TRAVAUX
  return handleTreesAndTravauxGet_();
}

// 📜 HISTORIQUE – GET
function handleHistoryGet_(e) {
  const treeId = String(e?.parameter?.id || "").trim();
  const limit = Number(e?.parameter?.limit || 50);

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

// 🌳 ARBRES + 🔧 TRAVAUX – GET
function handleTreesAndTravauxGet_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName("Patrimoine_arboré");
  const sheetTravaux = ss.getSheetByName(SHEET_TRAVAUX);

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
    const travaux = travauxMap[id] || {};

    return {
      createdAt: row[0]?.getTime?.() || null,
      id,
      lat,
      lng,
      species: row[4],
      height: row[5] !== "" ? Number(row[5]) : null,
      dbh: row[6] !== "" ? Number(row[6]) : null,
      secteur: row[7],
      address: row[8],
      tags: row[9] ? String(row[9]).split(",") : [],
      historiqueInterventions: row[10] || "",
      comment: row[11],
      photos: (() => {
        if (!row[12]) return [];
        try { return JSON.parse(row[12]); }
        catch { return []; }
      })(),
      etat: String(row[13] || "").trim(),
      secteurTravaux: (travaux.secteur || ""),
      updatedAt: row[14] ? Number(row[14]) : null,

      // ✅ TRAVAUX RENVOYÉS À L’APP
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
   DRIVE
========================= */
const DRIVE_FOLDER_ID = "1bC7CsCGBeQNp5ADelZ0SIXGjo12uhiUS";

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

/* =========================
   POST – LOGIN / CREATE / UPDATE / DELETE
========================= */
function doPost(e) {
  try {
    // 🔐 LOGIN (action=login & password=...)
    const actionParam = e?.parameter?.action;
    if (actionParam === "login") {
      const login = String(e?.parameter?.login || "").trim();
      const pwd = String(e?.parameter?.password || "");

      const user = USERS[login];
      if (!user || pwd !== user.password) return authFail_();

      const token = createToken_();
      setTokenMeta_(token, { role: user.role, secteur: user.secteur || "", login });

      return ContentService
        .createTextOutput(JSON.stringify({ ok: true, token, role: user.role, secteur: user.secteur || "", login }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // 🔐 AUTH obligatoire pour tout le reste
    const token = e?.parameter?.token;
    if (!isValidToken_(token)) return authFail_();

    // ✅ META pour historique
    const meta = getTokenMeta_(token); // {role, secteur, login}

    let data = {};

    // ✅ Accepte :
    // - payload JSON (payload=...)
    // - paramètres directs (action=...&id=...)
    // - JSON brut dans le body
    if (e && e.parameter && Object.keys(e.parameter).length) {
      if (e.parameter.payload) {
        data = JSON.parse(e.parameter.payload);
      } else {
        // paramètres directs
        data = { ...e.parameter };
      }
    } else if (e && e.postData && e.postData.contents) {
      data = JSON.parse(e.postData.contents);
    } else {
      throw new Error("Aucun payload reçu");
    }

    // ✅ si on reçoit { payload: {...} }
    if (data && data.payload) data = data.payload;

    // (optionnel) on ne garde pas token/password dans data pour éviter effets de bord
    if (data && typeof data === "object") {
      delete data.token;
      delete data.password;
    }

    
    /* ===== VALIDATION INTERVENTION ===== */
    if (data.action === "validateIntervention" && data.id && data.intervention) {
      const sheetVI = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("Patrimoine_arboré");
      const lastVI = sheetVI.getLastRow();
      if (lastVI > 1) {
        const rowsVI = sheetVI.getRange(2, 1, lastVI - 1, sheetVI.getLastColumn()).getValues();
        for (let i = 0; i < rowsVI.length; i++) {
          if (String(rowsVI[i][1]).trim() === String(data.id).trim()) {
            const rowIndex = i + 2;
            const existing = String(rowsVI[i][10] || "").trim(); // col 11 Historique
            const sep = existing ? "\n" : "";
            const value = existing + sep + data.intervention;
            sheetVI.getRange(rowIndex, 11).setValue(value);
            sheetVI.getRange(rowIndex, 15).setValue(Date.now());
            SpreadsheetApp.flush();

            logHistory_(meta, "VALIDATE_INTERVENTION", data.id, {
              added: data.intervention
            });

            return ok({ status: "INTERVENTION_ADDED" });
          }
        }
      }
      return ok({ status: "NOT_FOUND" });
    }

    /* =========================
       📄 EXPORT PDF
       - exportHistoryPdf : historique interventions regroupé par année
       - exportTreePdf    : fiche arbre + historique complet
    ========================= */
    if (data.action === "exportHistoryPdf") {
      const year = String(data.year || "").trim();
      const pdf = exportHistoryPdf_(year);
      // ✅ HISTORIQUE (audit)
      logHistory_(meta, "EXPORT_HISTORY_PDF", year || "ALL", { year: year || "ALL" });
      return jsonResponse({ ok: true, url: pdf.url, fileId: pdf.fileId, filename: pdf.filename });
    }

    if (data.action === "exportTreePdf" && data.id) {
      const treeId = String(data.id).trim();
      const pdf = exportTreePdf_(treeId);
      // ✅ HISTORIQUE (audit)
      logHistory_(meta, "EXPORT_TREE_PDF", treeId, { treeId });
      return jsonResponse({ ok: true, url: pdf.url, fileId: pdf.fileId, filename: pdf.filename });
    }

// 🔒 SÉCURITÉ SECTEUR :
    // un compte secteur ne peut enregistrer que dans son secteur
    if (meta && meta.role === "secteur") {
      data.secteur = meta.secteur || data.secteur || "";
    }

    const sheet = SpreadsheetApp
      .openById(SPREADSHEET_ID)
      .getSheetByName("Patrimoine_arboré");

    const lastRow = sheet.getLastRow();

    /* ===== SUPPRESSION PHOTO ===== */
    if (data.action === "deletePhoto" && data.photoDriveId && data.treeId) {

      // ✅ HISTORIQUE
      logHistory_(meta, "DELETE_PHOTO", data.treeId, {
        photoDriveId: data.photoDriveId
      });

      const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();

      for (let i = 0; i < rows.length; i++) {
        const sheetTreeId = String(rows[i][1]).trim();
        if (sheetTreeId === String(data.treeId).trim()) {

          let photos = [];
          try {
            photos = rows[i][12] ? JSON.parse(rows[i][12]) : [];
          } catch (err) {
            photos = [];
          }

          // Drive
          deletePhotoFromDrive(String(data.photoDriveId).trim());

          // Sheets
          const newPhotos = photos.filter(p =>
            String(p.driveId || "").trim() !== String(data.photoDriveId).trim()
          );

          sheet.getRange(i + 2, 13).setValue(JSON.stringify(newPhotos));
          SpreadsheetApp.flush();

          return ok({ status: "PHOTO_DELETED", remaining: newPhotos.length });
        }
      }

      return ok({ status: "NOT_FOUND" });
    }

    /* ===== SUPPRESSION ARBRE ===== */
    if (data.action === "delete" && data.id) {
      if (lastRow < 2) return ok({ status: "NOT_FOUND" });

      // ✅ HISTORIQUE
      const beforeObjDelete = getTreeRowAsObject_(sheet, data.id);
      logHistory_(meta, "DELETE", data.id, {
        deletedRow: beforeObjDelete || null
      });

      const rows = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();

      for (let i = 0; i < rows.length; i++) {
        if (String(rows[i][1]).trim() === String(data.id).trim()) {
          deleteTreeFolder(String(data.id).trim());
          sheet.deleteRow(i + 2);
          // ✅ tri après suppression
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
      catch { data.tags = String(data.tags).split(",").map(s => s.trim()).filter(Boolean); }
    }
    if (typeof data.photos === "string") {
      try { data.photos = JSON.parse(data.photos); }
      catch { data.photos = []; }
    }

    // ✅ HISTORIQUE : état avant update/create
    const beforeObj = getTreeRowAsObject_(sheet, data.id);

    /* ===== PHOTOS EXISTANTES ===== */
    let existingPhotos = [];
    if (lastRow > 1) {
      const rows = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
      for (let i = 0; i < rows.length; i++) {
        if (String(rows[i][1]).trim() === String(data.id).trim() && rows[i][11]) {
          existingPhotos = JSON.parse(rows[i][12]);
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
      new Date(),
      data.id || "",
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
      JSON.stringify(allPhotos),
      data.etat || "",
      data.updatedAt || Date.now()
    ];

    let isUpdate = false;

    /* ===== UPDATE ===== */
    if (lastRow > 1) {
      const ids = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
      for (let i = 0; i < ids.length; i++) {
        if (String(ids[i][0]).trim() === String(data.id).trim()) {
          sheet.getRange(i + 2, 1, 1, rowData.length)
            .setValues([rowData]);

          // ✅ tri après mise à jour
          sortArbresSheet_(sheet);

          colorRowByEtat(sheet, i + 2, data.etat);
          recolorOneArbreById_(sheet, data.id);
          isUpdate = true;
          break;
        }
      }
    }



    /* ===== TRAVAUX (Élagages / Abattages) ===== */
    const etatArbre = String(data.etat || "").trim();
    const ETATS_TRAVAUX = [
      "Dangereux (A abattre)",
      "A surveiller",
      "A élaguer (URGENT)",
      "A élaguer (Moyen)",
      "A élaguer (Faible)"
    ];
    const doitAllerTravaux = ETATS_TRAVAUX.includes(etatArbre);

    if (doitAllerTravaux) {
      const sheetTravaux = SpreadsheetApp
        .openById(SPREADSHEET_ID)
        .getSheetByName(SHEET_TRAVAUX);

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

            sheetTravaux
              .getRange(rowIndex, 1, 1, travauxRow.length)
              .setValues([travauxRow]);

            colorEtatTravaux(sheetTravaux, rowIndex, etatArbre);
            // ✅ tri après mise à jour travaux
            sortTravauxSheet_(sheetTravaux);
            // ✅ recolor fiable par ID (après tri)
            recolorOneTravauxById_(sheetTravaux, data.id);
            foundTravaux = true;
            break;
          }
        }
      }

      if (!foundTravaux) {
        sheetTravaux.appendRow(travauxRow);
        // ✅ tri après création travaux
        sortTravauxSheet_(sheetTravaux);
        // ✅ recolor fiable par ID (après tri)
        recolorOneTravauxById_(sheetTravaux, data.id);
        const newRow = sheetTravaux.getLastRow();
        colorEtatTravaux(sheetTravaux, newRow, etatArbre);
        recolorOneTravauxById_(sheetTravaux, data.id);
      }
    }

    /* ===== CREATE ===== */
    if (!isUpdate) {
      sheet.appendRow(rowData);
      // ✅ tri après création
      sortArbresSheet_(sheet);
      const newRow = sheet.getLastRow();
      colorRowByEtat(sheet, newRow, data.etat);
      recolorOneArbreById_(sheet, data.id);
    }

    SpreadsheetApp.flush();

    // ✅ HISTORIQUE : état après + diff + log CREATE/UPDATE
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

    logHistory_(meta, isUpdate ? "UPDATE" : "CREATE", data.id, {
      changes
    });

    return ok({ status: "CREATED", photos: allPhotos });

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/* =========================
   UTIL
========================= */
function ok(payload) {
  const output = ContentService.createTextOutput(
    JSON.stringify({ ok: true, result: payload })
  );
  output.setMimeType(ContentService.MimeType.JSON);
  return output;
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
    const folder = folders.next();
    folder.setTrashed(true);
  }
}

function assertSheetAlive() {
  const file = DriveApp.getFileById(SPREADSHEET_ID);
  if (file.isTrashed()) {
    throw new Error("❌ Le Spreadsheet est dans la corbeille !");
  }
}

function colorRowByEtat(sheet, rowIndex, etat) {
  let color = null;

  if (etat === "Dangereux (A abattre)") color = "#f28b82"; // rouge clair
  if (etat === "A surveiller")  color = "#fbbc04"; // orange clair
  if (etat === "A élaguer (URGENT)")  color = "#FFFF00"; // jaune
  if (etat === "A élaguer (Moyen)")  color = "#00FFFF"; // beuc lair
  if (etat === "A élaguer (Faible)")  color = "#ccff90"; // vert clair

  const range = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn());

  if (color) {
    range.setBackground(color);
  } else {
    range.setBackground(null); // reset
  }
}

function colorEtatTravaux(sheet, rowIndex, etat) {
  let color = null;

  if (etat === "Dangereux (A abattre)") color = "#f28b82"; // rouge clair
  if (etat === "A surveiller")  color = "#fbbc04"; // orange clair
  if (etat === "A élaguer (URGENT)")  color = "#FFFF00"; // jaune
  if (etat === "A élaguer (Moyen)")  color = "#00FFFF"; // beuc lair
  if (etat === "A élaguer (Faible)")  color = "#ccff90"; // vert clair

  // 👉 UNIQUEMENT la colonne État (B)
  const cell = sheet.getRange(rowIndex, 2);

  if (color) {
    cell.setBackground(color);
    cell.setFontWeight("bold");
  } else {
    cell.setBackground(null);
    cell.setFontWeight("normal");
  }
}

// ✅ jsonResponse CORRIGÉ (ContentService ne supporte pas setHeader)
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

// =========================
// 📌 TRI AUTOMATIQUE FEUILLE ARBRES
// Secteur (col 8) -> Adresse (col 9) -> Espèce (col 5)
// =========================
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

// =========================
// 📌 TRI AUTOMATIQUE FEUILLE TRAVAUX
// =========================
function sortTravauxSheet_(sheetTravaux) {
  // ✅ Désactivé pour éviter les effets de style (couleur qui se propage)
  return;
}

// =========================
// 🎨 RECOLOR TRAVAUX APRÈS TRI
// =========================
function recolorEtatTravauxColumn_(sheetTravaux) {
  const lastRow = sheetTravaux.getLastRow();
  if (lastRow < 2) return;

  const etats = sheetTravaux.getRange(2, 2, lastRow - 1, 1).getValues(); // col B
  for (let i = 0; i < etats.length; i++) {
    const rowIndex = i + 2;
    const etat = String(etats[i][0] || "").trim();
    colorEtatTravaux(sheetTravaux, rowIndex, etat);
  }
}

// =========================
// 🎨 RECOLOR ARBRES APRÈS TRI
// =========================
function recolorArbresRows_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  // etat colonne 13
  const etats = sheet.getRange(2, 13, lastRow - 1, 1).getValues();
  for (let i = 0; i < etats.length; i++) {
    const rowIndex = i + 2;
    const etat = String(etats[i][0] || "").trim();
    colorRowByEtat(sheet, rowIndex, etat);
  }
}

// =========================
// 🎯 COULEUR TRAVAUX PAR ID (FIABLE)
// =========================
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

// =========================
// 🎯 COULEUR ARBRES PAR ID (FIABLE)
// =========================
function recolorArbresById_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const ids = sheet.getRange(2, 2, lastRow - 1, 1).getValues();  // col B
  const etats = sheet.getRange(2, 13, lastRow - 1, 1).getValues(); // col 13

  for (let i = 0; i < ids.length; i++) {
    const rowIndex = i + 2;
    const id = String(ids[i][0] || "").trim();
    const etat = String(etats[i][0] || "").trim();
    if (!id) continue;
    colorRowByEtat(sheet, rowIndex, etat);
  }
}

// =========================
// 🎯 RECOLOR 1 ARBRE PAR ID
// =========================
function recolorOneArbreById_(sheet, treeId) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const ids = sheet.getRange(2, 2, lastRow - 1, 1).getValues(); // col B = ID
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]).trim() === String(treeId).trim()) {
      const rowIndex = i + 2;
      const etat = String(sheet.getRange(rowIndex, 13).getValue() || "").trim(); // col 13 = etat
      colorRowByEtat(sheet, rowIndex, etat);
      return;
    }
  }
}

// =========================
// 🎯 RECOLOR 1 TRAVAUX PAR ID
// =========================
function recolorOneTravauxById_(sheetTravaux, treeId) {
  const lastRow = sheetTravaux.getLastRow();
  if (lastRow < 2) return;

  const ids = sheetTravaux.getRange(2, 1, lastRow - 1, 1).getValues(); // col A = ID
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]).trim() === String(treeId).trim()) {
      const rowIndex = i + 2;
      const etat = String(sheetTravaux.getRange(rowIndex, 2).getValue() || "").trim(); // col B = etat
      colorEtatTravaux(sheetTravaux, rowIndex, etat);
      return;
    }
  }
}

/* =========================
   📄 EXPORT PDF (HISTORIQUE + FICHE ARBRE)
   - Génération HTML -> PDF via HtmlService
   - Stockage dans un sous-dossier Drive (Exports_PDF)
========================= */

function getOrCreateExportsFolder_() {
  const root = DriveApp.getFolderById(DRIVE_FOLDER_ID);
  const name = "Exports_PDF";
  const it = root.getFoldersByName(name);
  return it.hasNext() ? it.next() : root.createFolder(name);
}

function safeFileName_(s) {
  return String(s || "")
    .replace(/[\\/:*?\"<>|]+/g, "-")
    .replace(/\s+/g, " ")
    .trim();
}

function driveFileToDataUrl_(driveId) {
  if (!driveId) return "";
  try {
    const f = DriveApp.getFileById(driveId);
    const blob = f.getBlob();
    const bytes = blob.getBytes();
    const b64 = Utilities.base64Encode(bytes);
    const mime = blob.getContentType() || f.getMimeType() || "application/octet-stream";
    return `data:${mime};base64,${b64}`;
  } catch (e) {
    Logger.log("driveFileToDataUrl_ error: " + e);
    return "";
  }
}

function exportPdfFromHtml_(html, filename) {
  const folder = getOrCreateExportsFolder_();

  const htmlBlob = HtmlService.createHtmlOutput(html).getBlob();
  const pdfBlob = htmlBlob.getAs(MimeType.PDF).setName(filename);

  const file = folder.createFile(pdfBlob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  return {
    fileId: file.getId(),
    url: file.getUrl(),
    filename
  };
}

function pdfBaseCss_() {
  return `
    <style>
      @page { margin: 18mm 14mm; }
      body { font-family: Arial, Helvetica, sans-serif; color:#111; font-size:12px; }
      h1 { font-size:18px; margin:0 0 8px; }
      h2 { font-size:14px; margin:18px 0 8px; border-bottom: 1px solid #e5e7eb; padding-bottom:6px; }
      .muted { color:#6b7280; }
      .meta { margin: 6px 0 0; }
      .box { border:1px solid #e5e7eb; border-radius:10px; padding:10px; margin:10px 0; }
      .grid { width:100%; border-collapse: collapse; }
      .grid td { padding:6px 8px; vertical-align: top; }
      .grid td:first-child{ width: 34%; color:#374151; }
      .badge { display:inline-block; padding:2px 8px; border-radius:999px; background:#f3f4f6; border:1px solid #e5e7eb; font-size:11px; }
      .photo { width:100%; max-height:260px; object-fit: contain; border-radius:10px; border:1px solid #e5e7eb; }
      .list { margin:0; padding-left:16px; }
      .list li { margin: 6px 0; }
      .sep { height:1px; background:#e5e7eb; margin:14px 0; }
      .page-break { page-break-before: always; }
      .small { font-size: 11px; }
      .k { font-weight:700; }
    </style>
  `;
}

function findTreeRow_(treeId) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("Patrimoine_arboré");
  const last = sheet.getLastRow();
  if (last < 2) return null;
  const values = sheet.getRange(2, 1, last - 1, sheet.getLastColumn()).getValues();
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][1]).trim() === String(treeId).trim()) {
      return { rowIndex: i + 2, row: values[i] };
    }
  }
  return null;
}

function parsePhotos_(raw) {
  if (!raw) return [];
  try {
    const arr = JSON.parse(raw);
    return Array.isArray(arr) ? arr : [];
  } catch {
    return [];
  }
}

function splitInterventions_(txt) {
  return String(txt || "")
    .split(/\r?\n/)
    .map(s => String(s || "").trim())
    .filter(Boolean);
}

function extractYearFromInterventionLine_(line) {
  // Ex: "🛠 29/01/2026 10:45:00 — ..." ou "29/01/2026 — ..."
  const m = String(line || "").match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (!m) return "";
  return m[3];
}

function extractDateLabel_(line) {
  // renvoie le morceau "dd/mm/yyyy ..." si présent
  const m = String(line || "").match(/(\d{1,2}\/\d{1,2}\/\d{4}(?:\s+\d{1,2}:\d{2}(?::\d{2})?)?)/);
  return m ? m[1] : "";
}

function exportTreePdf_(treeId) {
  const found = findTreeRow_(treeId);
  if (!found) {
    throw new Error("Arbre introuvable : " + treeId);
  }

  const r = found.row;

  const t = {
    id: r[1],
    lat: r[2],
    lng: r[3],
    species: r[4],
    height: r[5],
    dbh: r[6],
    secteur: r[7],
    address: r[8],
    tags: r[9],
    historiqueInterventions: r[10],
    comment: r[11],
    photos: parsePhotos_(r[12]),
    etat: r[13],
    updatedAt: r[14]
  };

  const firstPhotoId = t.photos?.[0]?.driveId || "";
  const photoDataUrl = firstPhotoId ? driveFileToDataUrl_(firstPhotoId) : "";

  const interventions = splitInterventions_(t.historiqueInterventions);
  const generatedAt = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");

  const html = `
    <html><head><meta charset="utf-8">${pdfBaseCss_()}</head><body>
      <h1>Fiche de l’arbre — ${safeFileName_(t.id)}</h1>
      <div class="muted small">Exporté le ${generatedAt}</div>

      ${photoDataUrl ? `<div class="box"><img class="photo" src="${photoDataUrl}" /></div>` : ``}

      <div class="box">
        <table class="grid">
          <tr><td class="k">ID</td><td>${t.id || "—"}</td></tr>
          <tr><td class="k">Espèce</td><td>${t.species || "—"}</td></tr>
          <tr><td class="k">Secteur</td><td>${t.secteur || "—"}</td></tr>
          <tr><td class="k">Adresse</td><td>${t.address || "—"}</td></tr>
          <tr><td class="k">Coordonnées</td><td>${t.lat || "—"}, ${t.lng || "—"}</td></tr>
          <tr><td class="k">Hauteur estimée</td><td>${(t.height !== "" && t.height !== null && t.height !== undefined) ? (t.height + " m") : "—"}</td></tr>
          <tr><td class="k">Diamètre tronc</td><td>${(t.dbh !== "" && t.dbh !== null && t.dbh !== undefined) ? (t.dbh + " cm") : "—"}</td></tr>
          <tr><td class="k">État</td><td>${t.etat || "—"}</td></tr>
          <tr><td class="k">Tags</td><td>${t.tags || "—"}</td></tr>
        </table>
      </div>

      <div class="box">
        <div class="k">Commentaire</div>
        <div class="muted" style="white-space:pre-wrap">${(t.comment || "—").toString().replace(/</g,'&lt;').replace(/>/g,'&gt;')}</div>
      </div>

      <h2>Historique des interventions</h2>
      ${interventions.length ? `
        <ul class="list">
          ${interventions.map(x => {
            const safe = x.toString().replace(/</g,'&lt;').replace(/>/g,'&gt;');
            return `<li>${safe}</li>`;
          }).join('')}
        </ul>
      ` : `<div class="muted">Aucune intervention enregistrée.</div>`}
    </body></html>
  `;

  const filename = safeFileName_(`Fiche_Arbre_${t.id}.pdf`);
  return exportPdfFromHtml_(html, filename);
}

function exportHistoryPdf_(year) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("Patrimoine_arboré");
  const last = sheet.getLastRow();
  if (last < 2) {
    throw new Error("Aucun arbre dans la base.");
  }

  const rows = sheet.getRange(2, 1, last - 1, sheet.getLastColumn()).getValues();
  const byYear = {}; // year -> entries[]

  rows.forEach(r => {
    const treeId = r[1];
    const species = r[4];
    const secteur = r[7];
    const address = r[8];
    const hist = r[10];

    const lines = splitInterventions_(hist);
    lines.forEach(line => {
      const y = extractYearFromInterventionLine_(line);
      if (!y) return;
      if (year && String(year).trim() && String(y) !== String(year).trim()) return;

      (byYear[y] = byYear[y] || []).push({
        treeId,
        species,
        secteur,
        address,
        dateLabel: extractDateLabel_(line),
        line
      });
    });
  });

  const years = Object.keys(byYear).sort((a, b) => Number(b) - Number(a));
  const generatedAt = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");

  const title = year && String(year).trim()
    ? `Historique des interventions — ${String(year).trim()}`
    : "Historique des interventions — Toutes années";

  let body = `
    <h1>${title}</h1>
    <div class="muted small">Exporté le ${generatedAt}</div>
  `;

  if (years.length === 0) {
    body += `<div class="box"><div class="muted">Aucune intervention trouvée${year ? " pour l’année " + year : ""}.</div></div>`;
  } else {
    years.forEach((y, idx) => {
      const entries = byYear[y] || [];
      if (!entries.length) return;

      // tri : d'abord par date (si possible), sinon par texte
      entries.sort((a, b) => {
        const da = a.dateLabel || "";
        const db = b.dateLabel || "";
        return db.localeCompare(da);
      });

      body += `
        ${idx ? '<div class="page-break"></div>' : ''}
        <h2>Année ${y} <span class="badge">${entries.length} intervention(s)</span></h2>
        <div class="box">
          <table class="grid" style="border-collapse:separate; border-spacing:0 6px;">
            <tr>
              <td class="k">Date</td>
              <td class="k">Arbre</td>
              <td class="k">Secteur / Adresse</td>
              <td class="k">Intervention</td>
            </tr>
            ${entries.map(e => {
              const safeLine = String(e.line || "").replace(/</g,'&lt;').replace(/>/g,'&gt;');
              const safeAddr = String((e.secteur || "") + (e.address ? " • " + e.address : "")).replace(/</g,'&lt;').replace(/>/g,'&gt;');
              const safeTree = String((e.treeId || "") + (e.species ? " • " + e.species : "")).replace(/</g,'&lt;').replace(/>/g,'&gt;');
              return `
                <tr>
                  <td>${(e.dateLabel || "—")}</td>
                  <td>${safeTree || "—"}</td>
                  <td>${safeAddr || "—"}</td>
                  <td style="white-space:pre-wrap">${safeLine}</td>
                </tr>
              `;
            }).join('')}
          </table>
        </div>
      `;
    });
  }

  const html = `<html><head><meta charset="utf-8">${pdfBaseCss_()}</head><body>${body}</body></html>`;
  const filename = safeFileName_(`Historique_Interventions_${year && String(year).trim() ? String(year).trim() : 'ALL'}.pdf`);
  return exportPdfFromHtml_(html, filename);
}

