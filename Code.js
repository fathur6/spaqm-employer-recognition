/**
 * UniSZA Industrial & Employer Recognition Portal
 * Google Apps Script Integration Backend Code
 */

function doGet(e) {
  const activeUser = Session.getActiveUser().getEmail();
  const userDomain = activeUser.split('@')[1];

  // 1. Domain-Level Security Block
  if (userDomain !== 'unisza.edu.my') {
    return HtmlService.createHtmlOutput(
      '<div style="font-family: sans-serif; text-align: center; padding: 50px;">' +
      '<h2 style="color: #0c3c60;">Access Denied (403)</h2>' +
      '<p>This internal portal is restricted strictly to @unisza.edu.my verified accounts.</p>' +
      '<p>You are logged in as: <strong>' + activeUser + '</strong></p>' +
      '</div>'
    ).setTitle('UniSZA Portal - Access Denied');
  }

  // 2. Securely Fetch Secrets from Script Properties
  const scriptProperties = PropertiesService.getScriptProperties();
  const adminEmailsStr = scriptProperties.getProperty('ADMIN_EMAILS') || "";
  const ADMIN_EMAILS = adminEmailsStr.split(',').map(email => email.trim());
  const sheetUrl = scriptProperties.getProperty('SHEET_URL') || "#";

  // 3. Admin Role Verification
  const isAdmin = ADMIN_EMAILS.includes(activeUser);

  // 4. Serve Dynamic Template
  let template = HtmlService.createTemplateFromFile('index');
  template.userEmail = activeUser;
  template.isAdmin = isAdmin;
  template.sheetUrl = sheetUrl; // Inject secret URL directly to HTML

  return template.evaluate()
      .setTitle('UniSZA Employer Recognition & Research Portal')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Ensures the "Tindakan" (Actions) column is inserted right before "QS Nominee Consent"
 */
function ensureSheetStructure(sheet) {
  const lastCol = sheet.getLastColumn();
  
  if (lastCol === 0) {
    const headers = [
      "ID", "Syarikat", "Bidang Utama / Industri", "Bidang (THE WUR)", 
      "Cadangan Fakulti Utama", "Rentas Disiplin/Bidang", "Status Hubungan", "Tindakan",
      "QS Nominee Consent", "Nama Pegawai Hubungan", "E-mel Korporat", 
      "Tindakan Terakhir", "Tarikh Tindakan"
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    return;
  }

  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const actionIdx = headers.indexOf("Tindakan");
  const qsIdx = headers.indexOf("QS Nominee Consent");

  if (actionIdx === -1) {
    if (qsIdx !== -1) {
      sheet.insertColumnBefore(qsIdx + 1);
      sheet.getRange(1, qsIdx + 1).setValue("Tindakan");
    } else {
      sheet.getRange(1, lastCol + 1).setValue("Tindakan");
    }
  }
}

// --- BI-DIRECTIONAL TRANSLATION MAPS ---
function mapStatusToEn(bmStatus) {
  const map = {
    "Disasarkan": "Targeted",
    "Dalam Perbincangan": "In Discussion",
    "MoU/MoA Ditandatangani": "MoU/MoA Signed",
    "Penglibatan Aktif": "Active Engagement"
  };
  return map[bmStatus] || bmStatus || "Targeted";
}

function mapStatusToBm(enStatus) {
  const map = {
    "Targeted": "Disasarkan",
    "In Discussion": "Dalam Perbincangan",
    "MoU/MoA Signed": "MoU/MoA Ditandatangani",
    "Active Engagement": "Penglibatan Aktif"
  };
  return map[enStatus] || enStatus || "Disasarkan";
}

function mapActionToEn(bmAction) {
  const map = {
    "Hubungan Awal": "Initial Outreach",
    "Mesyuarat Dijadualkan": "Meeting Scheduled",
    "Merangka Perjanjian": "Drafting Agreement",
    "Memuktamadkan MoU": "Finalizing MoU",
    "Program Bersama": "Joint Program",
    "Tindakan Susulan": "Follow-up Required"
  };
  return map[bmAction] || bmAction || "Follow-up Required";
}

function mapActionToBm(enAction) {
  const map = {
    "Initial Outreach": "Hubungan Awal",
    "Meeting Scheduled": "Mesyuarat Dijadualkan",
    "Drafting Agreement": "Merangka Perjanjian",
    "Finalizing MoU": "Memuktamadkan MoU",
    "Joint Program": "Program Bersama",
    "Follow-up Required": "Tindakan Susulan"
  };
  return map[enAction] || enAction || "Tindakan Susulan";
}

function getMncData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName("MNC");
    if (!sheet) return [];
    
    ensureSheetStructure(sheet); 
    
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    
    if (lastRow <= 1 || lastCol === 0) return [];
    
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
    
    const safeIdx = (name) => headers.indexOf(name) !== -1 ? headers.indexOf(name) : -1;
    const getVal = (row, name, defaultVal = "") => safeIdx(name) !== -1 ? row[safeIdx(name)] : defaultVal;
    
    return data.map(row => ({
      id: parseInt(getVal(row, "ID", 0)) || 0,
      name: String(getVal(row, "Syarikat")),
      industry: String(getVal(row, "Bidang Utama / Industri")),
      wurField: String(getVal(row, "Bidang (THE WUR)")),
      primaryFaculty: String(getVal(row, "Cadangan Fakulti Utama")),
      crossFaculty: String(getVal(row, "Rentas Disiplin/Bidang")),
      status: mapStatusToEn(String(getVal(row, "Status Hubungan", "Disasarkan"))),
      action: mapActionToEn(String(getVal(row, "Tindakan", "Tindakan Susulan"))),
      qsNominated: String(getVal(row, "QS Nominee Consent", "No")).toLowerCase() === "yes",
      contactName: String(getVal(row, "Nama Pegawai Hubungan")),
      contactEmail: String(getVal(row, "E-mel Korporat")),
      lastAction: String(getVal(row, "Tindakan Terakhir")),
      lastActionDate: parseDateValue(getVal(row, "Tarikh Tindakan"))
    }));
  } catch (err) {
    Logger.log("Error in getMncData: " + err.toString());
    throw err;
  }
}

// --- SECURED ENDPOINT: Admin Only Save Function ---
function saveOrUpdateMnc(partner) {
  // Securely Fetch Secrets from Script Properties for Validation
  const scriptProperties = PropertiesService.getScriptProperties();
  const adminEmailsStr = scriptProperties.getProperty('ADMIN_EMAILS') || "";
  const ADMIN_EMAILS = adminEmailsStr.split(',').map(email => email.trim());
  
  const activeUser = Session.getActiveUser().getEmail();
  
  if (!ADMIN_EMAILS.includes(activeUser)) {
    throw new Error("Unauthorized Update: Admin privileges required.");
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName("MNC");
    if (!sheet) return false;
    
    ensureSheetStructure(sheet);
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    const idIdx = headers.indexOf("ID");
    
    const createRowData = (id) => {
      let rowData = new Array(headers.length).fill("");
      const setVal = (headerName, val) => {
        const idx = headers.indexOf(headerName);
        if (idx !== -1) rowData[idx] = val;
      };
      
      setVal("ID", id);
      setVal("Syarikat", partner.name);
      setVal("Bidang Utama / Industri", partner.industry);
      setVal("Bidang (THE WUR)", partner.wurField);
      setVal("Cadangan Fakulti Utama", partner.primaryFaculty);
      setVal("Rentas Disiplin/Bidang", partner.crossFaculty || "");
      setVal("Status Hubungan", mapStatusToBm(partner.status));
      setVal("Tindakan", mapActionToBm(partner.action));
      setVal("QS Nominee Consent", partner.qsNominated ? "Yes" : "No");
      setVal("Nama Pegawai Hubungan", partner.contactName || "");
      setVal("E-mel Korporat", partner.contactEmail || "");
      setVal("Tindakan Terakhir", partner.lastAction || "");
      setVal("Tarikh Tindakan", partner.lastActionDate || "");
      
      return rowData;
    };

    if (partner.id && lastRow > 1 && idIdx !== -1) {
      const ids = sheet.getRange(2, idIdx + 1, lastRow - 1, 1).getValues().map(r => parseInt(r[0]));
      const idx = ids.indexOf(parseInt(partner.id));
      if (idx !== -1) {
        sheet.getRange(idx + 2, 1, 1, headers.length).setValues([createRowData(partner.id)]);
        return true;
      }
    }
    return true;
  } catch (err) {
    Logger.log("Error in saveOrUpdateMnc: " + err.toString());
    throw err;
  }
}

function getResearchData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const getSheetData = (sheetName) => {
      const sheet = ss.getSheetByName(sheetName);
      if (!sheet) return [];
      const lastRow = sheet.getLastRow();
      const lastCol = sheet.getLastColumn();
      if (lastRow <= 1 || lastCol === 0) return [];
      return sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
    };

    const safeNum = (val) => { const num = Number(val); return isNaN(num) ? 0 : num; };

    return {
      rankUniversity: getSheetData('Researcher-RankUniversity').map(r => ({ org: String(r[0]||''), researcher: String(r[1]||''), sum: safeNum(r[2]) })),
      researcherCollaborator: getSheetData('Researcher-Collaborator').map(r => ({ country: String(r[0]||''), researcher: String(r[1]||''), univ: String(r[2]||''), count: safeNum(r[3]) })),
      uniCountryResearcher: getSheetData('UniCountry-Researcher').map(r => ({ researcher: String(r[0]||''), univ: String(r[1]||''), country: String(r[2]||''), sum: safeNum(r[3]) })),
      countryResearcher: getSheetData('Country-Researcher').map(r => ({ country: String(r[0]||''), researcher: String(r[1]||''), faculty: String(r[2]||''), count: safeNum(r[3]) })),
      universityCollaborator: getSheetData('UniversityCollaborator').map(r => ({ univ: String(r[0]||''), sum: safeNum(r[1]), country: String(r[2]||'') })),
      countryAffiliation: getSheetData('Country-Affiliation').map(r => ({ country: String(r[0]||''), sum: safeNum(r[1]) }))
    };
  } catch (e) {
    Logger.log("Error fetching research data: " + e.toString());
    throw e;
  }
}

function parseDateValue(dateVal) {
  if (dateVal instanceof Date) { return Utilities.formatDate(dateVal, Session.getScriptTimeZone(), "yyyy-MM-dd"); }
  return String(dateVal);
}