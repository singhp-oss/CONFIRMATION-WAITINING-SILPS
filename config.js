// ============================================================
//   PRADEEP CARGO - AWB AUTOMATION CONFIG
//   Sirf yahi file edit karo
// ============================================================

const path = require('path');

// Base folder — yahi ek jagah change karo, baaki sab isi se banega.
// IMPORTANT: Windows path me hamesha DOUBLE backslash (\\) likho.
const BASE_FOLDER = 'C:\\AWB_TOOLS\\WAITING-SLIPS_VERIFIER';

module.exports = {
    // WhatsApp Group ka exact naam
    WHATSAPP_GROUP_NAME: 'Pradeep Cargo Airport GRP',

    // Excel file ka full path (double backslash use karo)
    EXCEL_FILE_PATH: 'D:\\NEW_PC\\DAILY_WORKS\\PRADEEP_DAILYSHEET.xlsm',

    // Sheet name (exactly jaisa Excel me hai)
    EXCEL_SHEET_NAME: 'CCU_2025-26',

    // AWB column letter
    AWB_COLUMN: 'B',

    // Chrome ka path — agar yahan Chrome na mile to bundled Chromium use hoga
    CHROME_PATH: 'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe',

    // Base folder
    BASE_FOLDER,

    // Reports save folder (HTML dashboards)
    REPORTS_FOLDER: path.join(BASE_FOLDER, 'reports'),

    // Temp folder (working files)
    TEMP_FOLDER: path.join(BASE_FOLDER, 'temp'),

    // WhatsApp session folder (auth data — ise alag rakhna zaroori hai)
    SESSION_FOLDER: path.join(BASE_FOLDER, 'session'),

    // Airline name
    AIRLINE: 'IndiGo',

    // Fetch limit
    MESSAGE_FETCH_LIMIT: 500,

    // WhatsApp ready hone ke liye max wait (ms). Slow connection ke liye 4 min.
    WA_READY_TIMEOUT_MS: 240000
};
