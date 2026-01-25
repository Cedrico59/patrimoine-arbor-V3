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
   GET – LECTURE DES ARBRES
   (MODIF: ajout auth + param e)
========================= */

/* =========================
   GET – LECTURE (ARBRES + TRAVAUX + HISTORIQUE)
   ✅ 1 SEUL doGet (sinon Leaflet/app bug)
========================= */
function doGet(e) {
  // 🔐 AUTH
  const token = e?.parameter?.token;
  if (!isValidToken_(token)) return authFail_();
  const meta = getTokenMeta_(token) || {};

  // 📜 HISTORIQUE : GET?action=history&id=XXX
  if (e?.parameter?.action === "history") {
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

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName("Patrimoine_arboré");
  const sheetTravaux = ss.getSheetByName(SHEET_TRAVAUX);

  /* ===== LECTURE TRAVAUX =====
     Colonnes attendues dans tableau_Elagages/Abattages :
     A id | B etat | C secteur | D dateDemande | E natureTravaux | F address | G species
     H dateDemandeDevis | I devisNumero | J montantDevis | K dateExecution | L remarquesTravaux
     M numeroBDC | N numeroFacture
  */
  const travauxMap = {};
  if (sheetTravaux) {
    const lastT = sheetTravaux.getLastRow();
    if (lastT > 1) {
      const valuesT = sheetTravaux.getRange(2, 1, lastT - 1, sheetTravaux.getLastColumn()).getValues();
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

  /* ===== LECTURE ARBRES =====
     Colonnes attendues dans Patrimoine_arboré :
     A createdAt | B id | C lat | D lng | E species | F height | G dbh | H secteur | I address
     J tags | K historiqueInterventions | L comment | M photos | N etat | O updatedAt
  */
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return jsonArrayResponse_([]);

  const values = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();

  let trees = values.map(row => {
    const lat = Number(row[2]);
    const lng = Number(row[3]);
    const id = String(row[1] || "").trim();
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
      tags: row[9] ? String(row[9]).split(",").map(s=>s.trim()).filter(Boolean) : [],
      historiqueInterventions: row[10] || "",
      comment: row[11] || "",
      photos: (() => {
        if (!row[12]) return [];
        try { return JSON.parse(row[12]); } catch (e) { return []; }
      })(),
      etat: row[13] || "",
      updatedAt: row[14] ? Number(row[14]) : null,

      // ✅ TRAVAUX
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

  // 🔒 Filtrage serveur : un compte secteur ne voit que SON secteur
  if (meta.role === "secteur" && meta.secteur) {
    const s = String(meta.secteur).trim();
    trees = trees.filter(t => String(t.secteur || "").trim() === s);
  }

  return jsonArrayResponse_(trees);
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
   (MODIF: ajout login + auth)
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

    // ✅ META (AJOUT) pour historique
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

    // 🔒 SÉCURITÉ SECTEUR (AJOUT) :
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

      // ✅ HISTORIQUE (AJOUT)
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
          // ✅ Tri automatique : Secteur -> Adresse -> Espèce
    sortPatrimoineSheet_(sheet);

    SpreadsheetApp.flush();

          return ok({ status: "PHOTO_DELETED", remaining: newPhotos.length });
        }
      }

      return ok({ status: "NOT_FOUND" });
    }

    /* ===== SUPPRESSION ARBRE ===== */
    if (data.action === "delete" && data.id) {
      if (lastRow < 2) return ok({ status: "NOT_FOUND" });

      // ✅ HISTORIQUE (AJOUT)
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

    if (data.historyInterventions && !data.historiqueInterventions) data.historiqueInterventions = data.historyInterventions;

    // ✅ HISTORIQUE (AJOUT) : état avant update/create
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

    /* ===== DONNÉES =====
       ✅ Ajout updatedAt (col 14) pour correspondre à ton doGet row[13]
    */
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

    /* ===== TRAVAUX (Élagages / Abattages) =====
       ✅ NE PAS AJOUTER si pas de pastille Etat
    */
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
        data.id,
        etatArbre,
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

    // ✅ HISTORIQUE (AJOUT) : état après + diff + log CREATE/UPDATE
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


function sortPatrimoineSheet_(sheet) {
  try {
    if (!sheet) return;
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow <= 2) return; // rien à trier

    // Colonnes Patrimoine_arboré:
    // H (8) = secteur, I (9) = address, E (5) = species
    sheet.getRange(2, 1, lastRow - 1, lastCol).sort([
      { column: 8, ascending: true },
      { column: 9, ascending: true },
      { column: 5, ascending: true }
    ]);
  } catch (e) {
    Logger.log("Tri Patrimoine erreur: " + e);
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


function jsonArrayResponse_(arr) {
  return ContentService
    .createTextOutput(JSON.stringify(arr || []))
    .setMimeType(ContentService.MimeType.JSON)
    .setHeader("Access-Control-Allow-Origin", "*")
    .setHeader("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
    .setHeader("Access-Control-Allow-Headers", "Content-Type");
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON)
    .setHeader("Access-Control-Allow-Origin", "*")
    .setHeader("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
    .setHeader("Access-Control-Allow-Headers", "Content-Type");
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

    // ✅ IMPORTANT : ré-appliquer couleurs après tri
    
  } catch (e) {
    Logger.log("Tri arbres erreur: " + e);
  }
}



// =========================
// 📌 TRI AUTOMATIQUE FEUILLE TRAVAUX
// Secteur (col 3) -> Etat (col 2) -> Date demande (col 4)
// =========================
function sortTravauxSheet_(sheetTravaux) {
  // ✅ Désactivé pour éviter les effets de style (couleur qui se propage)
  // Si tu veux le tri, on pourra le remettre avec une approche "rebuild range"
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
