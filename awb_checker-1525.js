// ============================================================
//   PRADEEP CARGO - AWB AUTOMATION SCRIPT
//   Version: Final — Tesseract OCR + whatsapp-web.js
// ============================================================

const { Client, LocalAuth } = require('whatsapp-web.js');
const qrcode = require('qrcode-terminal');
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');
const readline = require('readline');
const Tesseract = require('tesseract.js');

const config = require('./config');
const { generateDashboard } = require('./dashboard');

// Folders ensure karo
[config.TEMP_FOLDER, config.REPORTS_FOLDER, config.SESSION_FOLDER].forEach(f => {
    if (!fs.existsSync(f)) fs.mkdirSync(f, { recursive: true });
});

// ============================================================
// OCR - Image analyze karo
// ============================================================
const AIRPORTS = ['CCU','DEL','BOM','MAA','BLR','HYD',
                  'IXS','IXB','IXA','PAT','GAU','IXC',
                  'IXZ','IMF','DIB','DMU','COK','AJL','BBI','BHO'];

async function analyzeImage(imgBuffer) {
    try {
        const { data: { text } } = await Tesseract.recognize(
            imgBuffer, 'eng', { logger: () => {} }
        );

        const tl = text.toLowerCase();

        // Slip check
        const isSlip = (tl.includes('booking') &&
                       (tl.includes('confirmation') ||
                        tl.includes('reference') ||
                        tl.includes('pradccudom'))) ||
                       /\d{3}[-\s]\d{8}/.test(text);

        if (!isSlip) return null;

        // Booking Reference (312-28539044 → last 8 digits)
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

        // Date (Wed, 27 May → 27-May-2026)
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
// WHATSAPP - Connect + Process
// ============================================================
async function runWhatsApp(stopDate, excelAWBs) {
    const confirmed = [];
    const pending   = [];
    const seen      = new Set();

    console.log('\n🌐 WhatsApp connect ho raha hai...');
    console.log('   (Pehli baar QR scan karo, baad me auto-login)\n');

    await new Promise(async (resolve) => {

        const client = new Client({
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
                executablePath: config.CHROME_PATH,
                args: [
                    '--no-sandbox',
                    '--disable-setuid-sandbox',
                    '--disable-dev-shm-usage',
                    '--disable-gpu',
                    '--no-first-run',
                    '--disable-extensions'
                ],
                timeout: 90000
            },
            authTimeoutMs: 90000,
            qrMaxRetries: 5
        });

        // QR Code
        client.on('qr', (qr) => {
            console.log('\n📱 QR CODE SCAN KARO:');
            console.log('   WhatsApp → ⋮ Menu → Linked Devices → Link a Device\n');
            qrcode.generate(qr, { small: true });
        });

        // Connected
        client.on('ready', async () => {
            console.log('[WA] ✅ Connected!\n');

            try {
                // Group dhundho
                const chats = await client.getChats();
                const group = chats.find(c => c.name === config.WHATSAPP_GROUP_NAME);

                if (!group) {
                    console.log(`❌ Group nahi mila: "${config.WHATSAPP_GROUP_NAME}"`);
                    console.log('   Available groups:');
                    chats.filter(c => c.isGroup)
                         .slice(0, 10)
                         .forEach(c => console.log(`   - ${c.name}`));
                    await client.destroy();
                    resolve();
                    return;
                }

                console.log(`✅ Group mila: ${group.name}`);
                console.log(`📸 Messages fetch ho rahe hain...\n`);

                // Messages fetch
                const messages = await group.fetchMessages({ limit: config.MESSAGE_FETCH_LIMIT });

                // Sirf image messages, latest pehle
                const imgMsgs = messages
                    .filter(m => m.hasMedia && m.type === 'image')
                    .reverse();

                console.log(`   Total image messages: ${imgMsgs.length}\n`);

                let stopCount = 0;
                let processed = 0;

                for (const msg of imgMsgs) {
                    try {
                        // Media download
                        const media = await msg.downloadMedia();
                        if (!media || !media.data) continue;

                        // Image buffer
                        const imgBuffer = Buffer.from(media.data, 'base64');

                        // OCR analyze
                        const result = await analyzeImage(imgBuffer);

                        if (!result || !result.is_slip) continue;

                        const awb   = result.booking_ref;
                        const date  = result.flight_date || 'date unknown';
                        const route = result.route || 'CCU-?';

                        // Date parse for stop check
                        let slipDate = null;
                        try {
                            const months = {
                                jan:0,feb:1,mar:2,apr:3,may:4,jun:5,
                                jul:6,aug:7,sep:8,oct:9,nov:10,dec:11
                            };
                            const dp = date.split('-');
                            if (dp.length === 3) {
                                const mk = dp[1].toLowerCase().slice(0, 3);
                                slipDate = new Date(parseInt(dp[2]), months[mk], parseInt(dp[0]));
                            }
                        } catch {}

                        // STOP CHECK
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

                        // Duplicate skip
                        if (seen.has(awb)) continue;
                        seen.add(awb);

                        // Excel match
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

            } catch (e) {
                console.log(`❌ Error: ${e.message}`);
            }

            await client.destroy();
            resolve();
        });

        client.on('auth_failure', () => {
            console.log('❌ Auth fail — session folder delete karke dobara try karo');
            console.log(`   Folder: ${config.SESSION_FOLDER}`);
            resolve();
        });

        client.on('disconnected', (reason) => {
            console.log(`⚠️  Disconnected: ${reason}`);
            resolve();
        });

        await client.initialize();
    });

    return { confirmed, pending };
}

// ============================================================
// MAIN
// ============================================================
async function main() {
    while (true) {
        const stopDate  = await getStopDate();
        const excelAWBs = loadExcelAWB();

        const { confirmed, pending } = await runWhatsApp(stopDate, excelAWBs);

        // Summary
        console.log('\n' + '='.repeat(55));
        console.log('   PRADEEP CARGO — AWB REPORT');
        console.log('='.repeat(55));
        console.log(`   Excel AWB Total : ${excelAWBs.size}`);
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
        generateDashboard(confirmed, pending, excelAWBs.size, stopDate);

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
