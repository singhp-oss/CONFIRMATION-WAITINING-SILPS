// ============================================================
//   PRADEEP CARGO - AWB AUTOMATION SCRIPT
//   Version: Robust — Tesseract OCR + whatsapp-web.js
//   Flow: connect (state-machine) -> load Excel -> scan -> dashboard
// ============================================================

const { Client, LocalAuth } = require('whatsapp-web.js');
const qrcode = require('qrcode-terminal');
const XLSX = require('xlsx');
const fs = require('fs');
const readline = require('readline');
const Tesseract = require('tesseract.js');

const config = require('./config');
const { generateDashboard } = require('./dashboard');

// Folders ensure karo (ab teeno alag alag hain)
[config.TEMP_FOLDER, config.REPORTS_FOLDER, config.SESSION_FOLDER].forEach(f => {
    if (!fs.existsSync(f)) fs.mkdirSync(f, { recursive: true });
});

// ============================================================
// OCR - Text parse karo (pure function, worker se text milta hai)
// ============================================================
const AIRPORTS = ['CCU','DEL','BOM','MAA','BLR','HYD',
                  'IXS','IXB','IXA','PAT','GAU','IXC',
                  'IXZ','IMF','DIB','DMU','COK','AJL','BBI','BHO'];

const MONTHS = { jan:0,feb:1,mar:2,apr:3,may:4,jun:5,
                 jul:6,aug:7,sep:8,oct:9,nov:10,dec:11 };

// OCR se aaye text ko slip-info me convert karta hai
function parseSlipText(text) {
    const tl = text.toLowerCase();

    // Slip check
    const isSlip = (tl.includes('booking') &&
                   (tl.includes('confirmation') ||
                    tl.includes('reference') ||
                    tl.includes('pradccudom'))) ||
                   /\d{3}[-\s]\d{8}/.test(text);

    if (!isSlip) return null;

    // Booking Reference (312-28539044 -> last 8 digits)
    const refMatch = text.match(/\b(\d{3})[-\s]?(\d{8})\b/);
    const booking_ref = refMatch ? (refMatch[1] + refMatch[2]).slice(-8) : null;

    // Route (CCU-IXS)
    let route = 'CCU-?';
    const words = text.toUpperCase().replace(/[^A-Z\s]/g, ' ').split(/\s+/);
    const foundAirports = [...new Set(words.filter(w => AIRPORTS.includes(w)))];
    if (foundAirports.length >= 2) {
        const origin = foundAirports.includes('CCU') ? 'CCU' : foundAirports[0];
        const dest   = foundAirports.find(w => w !== origin) || foundAirports[1];
        route = `${origin}-${dest}`;
    }

    // Date (Wed, 27 May -> 27-May-2026)
    let flight_date = 'date unknown';
    const dateMatch = text.match(
        /\b(?:mon|tue|wed|thu|fri|sat|sun)[a-z]*[,.]?\s+(\d{1,2})\s+(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*/i
    );
    if (dateMatch) {
        const day = String(parseInt(dateMatch[1])).padStart(2, '0');
        const mon = dateMatch[2].charAt(0).toUpperCase() + dateMatch[2].slice(1, 3).toLowerCase();
        flight_date = `${day}-${mon}-${new Date().getFullYear()}`;
    }

    return { is_slip: true, booking_ref, flight_date, route };
}

// "27-May-2026" -> Date object (stop-date compare ke liye)
function slipDateToObj(dateStr) {
    try {
        const dp = dateStr.split('-');
        if (dp.length !== 3) return null;
        const mk = dp[1].toLowerCase().slice(0, 3);
        if (!(mk in MONTHS)) return null;
        return new Date(parseInt(dp[2]), MONTHS[mk], parseInt(dp[0]));
    } catch {
        return null;
    }
}

// ============================================================
// EXCEL LOAD
// ============================================================
function loadExcelAWB() {
    console.log('\n📂 Excel file load ho rahi hai...');
    try {
        const wb = XLSX.readFile(config.EXCEL_FILE_PATH, {
            cellDates: false,
            sheetStubs: false,
            dense: true
        });

        const ws = wb.Sheets[config.EXCEL_SHEET_NAME];
        if (!ws) {
            console.log(`❌ Sheet nahi mili: ${config.EXCEL_SHEET_NAME}`);
            process.exit(1);
        }

        const colIdx = config.AWB_COLUMN.toUpperCase().charCodeAt(0) - 65;
        const awbSet = new Set();
        const range  = XLSX.utils.decode_range(ws['!ref'] || 'A1');

        for (let r = 1; r <= range.e.r; r++) {
            try {
                const row  = ws[r];
                if (!row) continue;
                const cell = row[colIdx];
                if (!cell || !cell.v) continue;
                const val  = String(cell.v).trim().replace(/-/g, '').replace(/\s/g, '');
                if (val.length >= 8) awbSet.add(val.slice(-8));
            } catch { continue; }
        }

        console.log(`✅ Excel me ${awbSet.size} AWB numbers load hue`);
        return awbSet;

    } catch (e) {
        console.log(`❌ Excel error: ${e.message}`);
        process.exit(1);
    }
}

// ============================================================
// USER INPUT - Stop date
// ============================================================
function getStopDate() {
    return new Promise((resolve) => {
        const rl = readline.createInterface({ input: process.stdin, output: process.stdout });

        console.log('\n' + '='.repeat(55));
        console.log('   PRADEEP CARGO AWB AUTOMATION');
        console.log('='.repeat(55));

        const ask = () => {
            rl.question('\n  Stop date daalo (DD-MM-YYYY, jaise 28-05-2026): ', (ans) => {
                const parts = ans.trim().split('-');
                if (parts.length === 3) {
                    const d = new Date(parseInt(parts[2]), parseInt(parts[1]) - 1, parseInt(parts[0]));
                    if (!isNaN(d.getTime())) {
                        rl.close();
                        console.log(`  ✅ ${d.toDateString()} tak scan hoga`);
                        resolve(d);
                        return;
                    }
                }
                console.log('  ❌ Format galat — DD-MM-YYYY likhna hai');
                ask();
            });
        };
        ask();
    });
}

// ============================================================
// RESTART/EXIT PROMPT
// ============================================================
function askRestart() {
    return new Promise((resolve) => {
        const rl = readline.createInterface({ input: process.stdin, output: process.stdout });
        console.log('\n' + '='.repeat(40));
        console.log('  R = Restart (dobara scan)');
        console.log('  E = Exit');
        console.log('='.repeat(40));
        rl.question('  Choice (R/E): ', (ans) => {
            rl.close();
            resolve(ans.trim().toUpperCase() === 'R');
        });
    });
}

// ============================================================
// WHATSAPP CLIENT banao
// ============================================================
function createClient() {
    // Agar diya hua Chrome path exist nahi karta to bundled Chromium use karo
    const chromeExists = config.CHROME_PATH && fs.existsSync(config.CHROME_PATH);
    if (!chromeExists) {
        console.log('⚠️  Diya hua Chrome path nahi mila — bundled Chromium use hoga.');
    }

    return new Client({
        authStrategy: new LocalAuth({
            clientId: 'pccs-awb',
            dataPath: config.SESSION_FOLDER
        }),
        webVersionCache: {
            type: 'remote',
            remotePath: 'https://raw.githubusercontent.com/wppconnect-team/wa-version/main/html/2.3000.1015901765-alpha/index.html'
        },
        puppeteer: {
            headless: true,
            ...(chromeExists ? { executablePath: config.CHROME_PATH } : {}),
            args: [
                '--no-sandbox',
                '--disable-setuid-sandbox',
                '--disable-dev-shm-usage',
                '--disable-gpu',
                '--no-first-run',
                '--disable-extensions'
            ],
            timeout: config.WA_READY_TIMEOUT_MS
        },
        authTimeoutMs: config.WA_READY_TIMEOUT_MS,
        qrMaxRetries: 5,
        restartOnAuthFail: true
    });
}

// ============================================================
// WHATSAPP CONNECT — state machine + hard timeout + single-settle
// Ye promise SIRF tab resolve hota hai jab client 'ready' hota hai.
// ============================================================
function connectWhatsApp(client) {
    return new Promise((resolve, reject) => {
        let settled = false;
        let lastPct = -1;

        // Ek hi baar settle ho — double-resolve/late-event race se bachne ke liye
        const settle = (fn, arg) => {
            if (settled) return;
            settled = true;
            clearTimeout(hardTimer);
            fn(arg);
        };

        // Agar client atak jaye (ready hi na aaye) to yahan se cleanly nikal jao
        const hardTimer = setTimeout(() => {
            settle(reject, new Error(
                `WhatsApp ${Math.round(config.WA_READY_TIMEOUT_MS / 1000)}s me ready nahi hua ` +
                `(internet slow ya session issue). Dobara try karo.`
            ));
        }, config.WA_READY_TIMEOUT_MS);

        // --- Stage 1: QR (sirf pehli baar) ---
        client.on('qr', (qr) => {
            console.log('\n📱 QR CODE SCAN KARO:');
            console.log('   WhatsApp → ⋮ Menu → Linked Devices → Link a Device\n');
            qrcode.generate(qr, { small: true });
        });

        // --- Stage 2: Sync progress (yahi wo lag hai jo pehle dikhta nahi tha) ---
        client.on('loading_screen', (percent, message) => {
            const pct = parseInt(percent);
            if (pct !== lastPct) {
                lastPct = pct;
                console.log(`   ⏳ WhatsApp sync ho raha hai... ${pct}% ${message || ''}`);
            }
        });

        // --- Stage 3: Auth ho gaya (abhi ready nahi) ---
        client.on('authenticated', () => {
            console.log('   🔐 Authenticated — ab chats sync ho rahe hain...');
        });

        client.on('change_state', (state) => {
            console.log(`   🔄 State: ${state}`);
        });

        // --- Stage 4: READY (yahi se aage matching hogi) ---
        client.on('ready', () => {
            console.log('\n[WA] ✅ Connected & Ready!\n');
            settle(resolve);
        });

        // --- Failures ---
        client.on('auth_failure', (msg) => {
            settle(reject, new Error(
                `Auth fail: ${msg || 'unknown'}. Session folder delete karke dobara QR scan karo:\n` +
                `   ${config.SESSION_FOLDER}`
            ));
        });

        client.on('disconnected', (reason) => {
            // Ready se pehle disconnect = failure. Ready ke baad aane wala
            // late disconnect settled flag ki wajah se ignore ho jayega.
            settle(reject, new Error(`Disconnected before ready: ${reason}`));
        });

        console.log('\n🌐 WhatsApp connect ho raha hai...');
        console.log('   (Pehli baar QR scan karo, baad me auto-login)\n');

        client.initialize().catch((e) => settle(reject, e));
    });
}

// ============================================================
// GROUP SCAN — ready hone ke BAAD, ek reused OCR worker ke saath
// ============================================================
async function scanGroup(client, worker, stopDate, excelAWBs) {
    const confirmed = [];
    const pending   = [];
    const seen      = new Set();

    // Group dhundho
    const chats = await client.getChats();
    const group = chats.find(c => c.name === config.WHATSAPP_GROUP_NAME);

    if (!group) {
        console.log(`❌ Group nahi mila: "${config.WHATSAPP_GROUP_NAME}"`);
        console.log('   Available groups:');
        chats.filter(c => c.isGroup)
             .slice(0, 10)
             .forEach(c => console.log(`   - ${c.name}`));
        return { confirmed, pending };
    }

    console.log(`✅ Group mila: ${group.name}`);
    console.log(`📸 Messages fetch ho rahe hain...\n`);

    const messages = await group.fetchMessages({ limit: config.MESSAGE_FETCH_LIMIT });

    // Sirf image messages, latest pehle
    const imgMsgs = messages
        .filter(m => m.hasMedia && m.type === 'image')
        .reverse();

    console.log(`   Total image messages: ${imgMsgs.length}\n`);

    let stopCount = 0;
    let processed = 0;

    // NOTE: Sequential rakha hai kyunki stop-date logic ko order chahiye
    // (2 consecutive purani slips par ruk jao). Speed ka fayda reused
    // Tesseract worker se aata hai — har image par naya worker nahi banta.
    for (const msg of imgMsgs) {
        try {
            const media = await msg.downloadMedia();
            if (!media || !media.data) continue;

            const imgBuffer = Buffer.from(media.data, 'base64');

            // OCR — reused worker (bahut fast vs Tesseract.recognize har baar)
            const { data: { text } } = await worker.recognize(imgBuffer);
            const result = parseSlipText(text);

            if (!result || !result.is_slip) continue;

            const awb   = result.booking_ref;
            const date  = result.flight_date || 'date unknown';
            const route = result.route || 'CCU-?';

            // STOP CHECK — 2 consecutive purani slips par ruk jao
            const slipDate = slipDateToObj(date);
            if (slipDate && slipDate <= stopDate) {
                stopCount++;
                console.log(`   🔴 Stop date (${stopCount}/2): ${date}`);
                if (stopCount >= 2) {
                    console.log('\n🛑 SCAN BAND!\n');
                    break;
                }
                continue;
            } else {
                stopCount = 0;
            }

            if (!awb) continue;
            if (seen.has(awb)) continue;   // duplicate skip
            seen.add(awb);

            const entry = { awb, route, date };
            if (excelAWBs.has(awb)) {
                confirmed.push(entry);
                console.log(`   ✅ Confirmed: ${awb} | ${route} | ${date}`);
            } else {
                pending.push(entry);
                console.log(`   ⏳ Pending : ${awb} | ${route} | ${date}`);
            }

            processed++;

        } catch { continue; }
    }

    console.log(`\n✅ Scan complete! Processed: ${processed}`);
    return { confirmed, pending };
}

// ============================================================
// EK PURA RUN — connect -> excel -> ocr worker -> scan -> cleanup
// ============================================================
async function runOnce(stopDate) {
    const client = createClient();
    let worker = null;

    try {
        // 1) WhatsApp poori tarah ready hone tak EXPLICITLY wait karo
        await connectWhatsApp(client);

        // 2) Ab (ready hone ke baad) Excel load karo — matching gate yahi hai
        const excelAWBs = loadExcelAWB();

        // 3) Ek hi OCR worker banao, poore scan me reuse hoga
        console.log('🧠 OCR engine start ho raha hai...');
        worker = await Tesseract.createWorker('eng', 1, { logger: () => {} });

        // 4) Scan + match
        const { confirmed, pending } = await scanGroup(client, worker, stopDate, excelAWBs);

        return { confirmed, pending, excelSize: excelAWBs.size };

    } finally {
        // Cleanup — chahe success ho ya error, dono cheezein band karo
        if (worker) {
            try { await worker.terminate(); } catch {}
        }
        try { await client.destroy(); } catch {}
    }
}

// ============================================================
// MAIN
// ============================================================
async function main() {
    while (true) {
        const stopDate = await getStopDate();

        let confirmed = [], pending = [], excelSize = 0;
        try {
            const res = await runOnce(stopDate);
            confirmed = res.confirmed;
            pending   = res.pending;
            excelSize = res.excelSize;
        } catch (e) {
            console.log(`\n❌ Run fail: ${e.message}\n`);
            const retry = await askRestart();
            if (!retry) { console.log('\n   Goodbye! 👋\n'); process.exit(0); }
            console.log('\n🔄 Restart ho raha hai...\n');
            continue;
        }

        // Summary
        console.log('\n' + '='.repeat(55));
        console.log('   PRADEEP CARGO — AWB REPORT');
        console.log('='.repeat(55));
        console.log(`   Excel AWB Total : ${excelSize}`);
        console.log(`   Slips Scanned   : ${confirmed.length + pending.length}`);
        console.log(`   ✅ Confirmed    : ${confirmed.length}`);
        console.log(`   ⏳ Pending      : ${pending.length}`);
        console.log('='.repeat(55));

        if (pending.length > 0) {
            console.log(`\n   🔴 ${pending.length} PDF to be confirmed by ${config.AIRLINE}`);
        } else {
            console.log(`\n   🟢 Sab AWB confirmed hain!`);
        }

        // Dashboard
        generateDashboard(confirmed, pending, excelSize, stopDate);

        // Restart or Exit
        const restart = await askRestart();
        if (!restart) {
            console.log('\n   Goodbye! 👋\n');
            process.exit(0);
        }
        console.log('\n🔄 Restart ho raha hai...\n');
    }
}

main().catch(e => {
    console.error('❌ Fatal error:', e.message);
    process.exit(1);
});
