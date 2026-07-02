// ============================================================
//   PRADEEP CARGO - DASHBOARD GENERATOR
// ============================================================

const fs = require('fs');
const path = require('path');
const { exec } = require('child_process');
const config = require('./config');

function generateDashboard(confirmed, pending, totalExcel, stopDate) {
    const now = new Date();
    const ts = [
        String(now.getDate()).padStart(2,'0'),
        String(now.getMonth()+1).padStart(2,'0'),
        now.getFullYear()
    ].join('-') + '_' + [
        String(now.getHours()).padStart(2,'0'),
        String(now.getMinutes()).padStart(2,'0')
    ].join('-');

    const timeStr = now.toLocaleString('en-IN', {
        day:'2-digit', month:'short', year:'numeric',
        hour:'2-digit', minute:'2-digit', hour12:true
    });

    const stopStr = stopDate.toLocaleDateString('en-IN', {
        day:'2-digit', month:'short', year:'numeric'
    });

    const trackURL = (awb) =>
        `https://6ecargo.goindigo.in/FrmAWBTracking.aspx?awbno=312${awb}`;

    const pendingRows = pending.map(i => `
        <tr>
            <td><span class="badge badge-pending">⏳</span> ${i.awb}</td>
            <td>${i.route}</td>
            <td>${i.date}</td>
            <td><a href="${trackURL(i.awb)}" target="_blank" class="track-btn">🔍 Track</a></td>
        </tr>`).join('');

    const confirmedRows = confirmed.map(i => `
        <tr>
            <td><span class="badge badge-confirmed">✅</span> ${i.awb}</td>
            <td>${i.route}</td>
            <td>${i.date}</td>
            <td><a href="${trackURL(i.awb)}" target="_blank" class="track-btn track-btn-ok">🔍 Track</a></td>
        </tr>`).join('');

    const html = `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Pradeep Cargo — AWB Dashboard</title>
<style>
* { margin:0; padding:0; box-sizing:border-box; }
body { font-family:'Segoe UI',Tahoma,sans-serif; background:#eef2f7; padding:24px; color:#333; }

.header {
    background:linear-gradient(135deg,#1a237e,#283593);
    color:white; padding:24px 30px; border-radius:14px;
    margin-bottom:20px; display:flex;
    justify-content:space-between; align-items:center;
    box-shadow:0 4px 15px rgba(26,35,126,0.3);
}
.header h1 { font-size:20px; letter-spacing:0.5px; }
.header .sub { font-size:12px; opacity:0.75; margin-top:4px; }
.header .meta { text-align:right; font-size:12px; opacity:0.8; }

.alert { padding:16px 22px; border-radius:10px; margin-bottom:20px; font-weight:600; font-size:15px; }
.alert-pending { background:#fff3f3; border-left:5px solid #e53935; color:#c62828; }
.alert-ok { background:#f1fff4; border-left:5px solid #43a047; color:#2e7d32; }

.cards { display:grid; grid-template-columns:repeat(4,1fr); gap:16px; margin-bottom:20px; }
.card { background:white; border-radius:12px; padding:20px; text-align:center; box-shadow:0 2px 10px rgba(0,0,0,0.07); border-top:4px solid transparent; }
.card.c1 { border-color:#5c6bc0; }
.card.c2 { border-color:#42a5f5; }
.card.c3 { border-color:#66bb6a; }
.card.c4 { border-color:#ef5350; }
.card .num { font-size:38px; font-weight:700; margin-bottom:6px; }
.card.c1 .num { color:#3949ab; }
.card.c2 .num { color:#1e88e5; }
.card.c3 .num { color:#43a047; }
.card.c4 .num { color:#e53935; }
.card .lbl { font-size:11px; color:#888; text-transform:uppercase; letter-spacing:1.2px; }

.tables { display:grid; grid-template-columns:1fr 1fr; gap:16px; margin-bottom:20px; }
.box { background:white; border-radius:12px; padding:20px; box-shadow:0 2px 10px rgba(0,0,0,0.07); }
.box h2 { font-size:14px; font-weight:700; margin-bottom:14px; padding-bottom:10px; border-bottom:2px solid #f0f0f0; }
.box.pending h2 { color:#c62828; }
.box.confirmed h2 { color:#2e7d32; }

table { width:100%; border-collapse:collapse; font-size:13px; }
th { background:#f8f9fa; padding:10px 12px; text-align:left; font-weight:600; color:#555; font-size:11px; text-transform:uppercase; }
td { padding:9px 12px; border-bottom:1px solid #f5f5f5; }
tr:last-child td { border-bottom:none; }
tr:hover td { background:#fafafa; }

.badge { display:inline-block; padding:2px 8px; border-radius:20px; font-size:11px; font-weight:600; }
.badge-pending { background:#ffebee; color:#c62828; }
.badge-confirmed { background:#e8f5e9; color:#2e7d32; }

.track-btn {
    display:inline-block; padding:4px 12px;
    background:#1a237e; color:white; border-radius:20px;
    font-size:11px; font-weight:600; text-decoration:none;
    transition:all 0.2s;
}
.track-btn:hover { background:#283593; transform:translateY(-1px); }
.track-btn-ok { background:#2e7d32; }
.track-btn-ok:hover { background:#1b5e20; }

.empty { text-align:center; padding:40px; color:#bbb; font-size:14px; }

.reports-box { background:white; border-radius:12px; padding:20px; box-shadow:0 2px 10px rgba(0,0,0,0.07); margin-bottom:20px; }
.reports-box h2 { font-size:14px; font-weight:700; color:#1a237e; margin-bottom:10px; }
.reports-box code { background:#f5f5f5; padding:8px 14px; border-radius:6px; display:inline-block; font-size:12px; color:#333; margin-top:6px; }

.footer { text-align:center; color:#aaa; font-size:12px; margin-top:10px; }
</style>
</head>
<body>

<div class="header">
    <div>
        <h1>✈️ PRADEEP CARGO &amp; COURIER SERVICE</h1>
        <div class="sub">AWB Confirmation Dashboard &nbsp;|&nbsp; Stop Date: ${stopStr}</div>
    </div>
    <div class="meta">Generated: ${timeStr}</div>
</div>

${pending.length > 0
    ? `<div class="alert alert-pending">🔴 &nbsp;${pending.length} AWB ${config.AIRLINE} se confirm hona baaki — PDF ka wait karo</div>`
    : `<div class="alert alert-ok">🟢 &nbsp;Sab AWB ${config.AIRLINE} dwara confirmed hain!</div>`
}

<div class="cards">
    <div class="card c1"><div class="num">${totalExcel}</div><div class="lbl">Excel AWB Total</div></div>
    <div class="card c2"><div class="num">${confirmed.length + pending.length}</div><div class="lbl">Slips Scanned</div></div>
    <div class="card c3"><div class="num">${confirmed.length}</div><div class="lbl">✅ Confirmed</div></div>
    <div class="card c4"><div class="num">${pending.length}</div><div class="lbl">⏳ Pending</div></div>
</div>

<div class="tables">
    <div class="box pending">
        <h2>⏳ PENDING (${pending.length}) — ${config.AIRLINE} se confirm hona baaki</h2>
        ${pending.length > 0 ? `
        <table>
            <thead><tr><th>AWB Number</th><th>Route</th><th>Date</th><th>Status</th></tr></thead>
            <tbody>${pendingRows}</tbody>
        </table>` : '<div class="empty">🎉 Koi pending nahi!</div>'}
    </div>
    <div class="box confirmed">
        <h2>✅ CONFIRMED (${confirmed.length})</h2>
        ${confirmed.length > 0 ? `
        <table>
            <thead><tr><th>AWB Number</th><th>Route</th><th>Date</th><th>Status</th></tr></thead>
            <tbody>${confirmedRows}</tbody>
        </table>` : '<div class="empty">Koi confirmed nahi</div>'}
    </div>
</div>

<div class="reports-box">
    <h2>📁 Purani Reports</h2>
    <p style="font-size:13px;color:#666;">Pichli sabhi reports yahan saved hain:</p>
    <code>${config.REPORTS_FOLDER}</code>
</div>

<div class="footer">PCCS AWB Automation &nbsp;|&nbsp; Node.js + Tesseract OCR</div>

</body>
</html>`;

    try {
        // FIX: Ensure reports folder exists
        if (!fs.existsSync(config.REPORTS_FOLDER)) {
            fs.mkdirSync(config.REPORTS_FOLDER, { recursive: true });
        }

        // Save timestamped report in reports folder
        const reportFile = path.join(config.REPORTS_FOLDER, `report_${ts}.html`);
        fs.writeFileSync(reportFile, html, 'utf8');
        console.log(`\n✅ Report saved: ${reportFile}`);

        // FIX: Save to config folder instead of hardcoded D:\\AWB_Automation
        const dashFile = path.join(config.REPORTS_FOLDER, 'dashboard.html');
        fs.writeFileSync(dashFile, html, 'utf8');
        console.log(`✅ Dashboard updated: ${dashFile}`);

        // Try to open in default browser
        const command = process.platform === 'win32' 
            ? `start "" "${dashFile}"`
            : process.platform === 'darwin'
            ? `open "${dashFile}"`
            : `xdg-open "${dashFile}"`;
        
        exec(command, (err) => {
            if (!err) {
                console.log(`✅ Dashboard browser me khul gaya!`);
            } else {
                console.log(`⚠️  Dashboard file ready at: ${dashFile}`);
                console.log(`   Manually kholo browser me`);
            }
        });

    } catch (e) {
        console.log(`❌ Dashboard error: ${e.message}`);
    }
}

module.exports = { generateDashboard };
