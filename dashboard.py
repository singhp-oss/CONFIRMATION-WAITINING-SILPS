# ============================================================
#   PRADEEP CARGO - DASHBOARD
#   Version: 11.0 - Route Column Added
# ============================================================

import os
import webbrowser
from datetime import datetime


def generate_dashboard(confirmed, pending, excel_awbs, stop_date):
    """
    Dashboard HTML generate karo with Route column.
    confirmed / pending = list of dicts:
        {"awb": "28539044", "date": "26-May-2026", "route": "CCU → IXS (6E487)"}
    """

    gen_time  = datetime.now().strftime("%d %b %Y, %I:%M %p")
    stop_str  = stop_date.strftime("%d %b %Y")

    total_confirmed = len(confirmed)
    total_pending   = len(pending)
    total_excel     = len(excel_awbs)

    # ── Table rows helper ────────────────────────────────────
    def make_rows(entries, status_class, status_label):
        if not entries:
            return f"""
            <tr>
              <td colspan="4" style="text-align:center;color:#94a3b8;padding:28px 0;font-style:italic;">
                Koi entry nahi
              </td>
            </tr>"""
        rows = ""
        for i, e in enumerate(entries, 1):
            awb   = e.get("awb",   "—")
            date  = e.get("date",  "—")
            route = e.get("route", "N/A")
            rows += f"""
            <tr class="data-row">
              <td class="td-num">{i}</td>
              <td class="td-awb">312-{awb}</td>
              <td class="td-route">
                <span class="route-badge">{route}</span>
              </td>
              <td class="td-date">{date}</td>
            </tr>"""
        return rows

    confirmed_rows = make_rows(confirmed, "confirmed", "Confirmed")
    pending_rows   = make_rows(pending,   "pending",   "Pending")

    # ── Not-found AWBs (Excel me hai, slip me nahi) ──────────
    confirmed_set  = {e["awb"] for e in confirmed}
    pending_set    = {e["awb"] for e in pending}
    found_set      = confirmed_set | pending_set
    # Excel AWBs jo kisi bhi list me nahi
    not_found      = sorted(excel_awbs - found_set)
    not_found_html = ""
    if not_found:
        items = "".join(f'<span class="nf-chip">312-{a}</span>' for a in not_found)
        not_found_html = f"""
        <div class="card nf-card">
          <div class="card-title">⚠️ Excel AWBs — Slip Nahi Mili ({len(not_found)})</div>
          <div class="nf-chips">{items}</div>
        </div>"""

    html = f"""<!DOCTYPE html>
<html lang="hi">
<head>
<meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width,initial-scale=1"/>
<title>Pradeep Cargo Dashboard</title>
<style>
  /* ── Reset & Base ──────────────────────────────── */
  *, *::before, *::after {{ box-sizing: border-box; margin: 0; padding: 0; }}

  body {{
    font-family: 'Segoe UI', system-ui, sans-serif;
    background: #0f172a;
    color: #e2e8f0;
    min-height: 100vh;
    padding: 28px 20px 60px;
  }}

  /* ── Header ────────────────────────────────────── */
  .header {{
    text-align: center;
    margin-bottom: 32px;
  }}
  .header-logo {{
    font-size: 2rem;
    margin-bottom: 6px;
  }}
  .header h1 {{
    font-size: 1.6rem;
    font-weight: 700;
    color: #f8fafc;
    letter-spacing: .5px;
  }}
  .header-sub {{
    font-size: .82rem;
    color: #64748b;
    margin-top: 6px;
  }}

  /* ── Stats Row ─────────────────────────────────── */
  .stats {{
    display: flex;
    gap: 14px;
    justify-content: center;
    flex-wrap: wrap;
    margin-bottom: 32px;
  }}
  .stat-box {{
    background: #1e293b;
    border: 1px solid #334155;
    border-radius: 14px;
    padding: 18px 28px;
    text-align: center;
    min-width: 130px;
  }}
  .stat-box .s-num {{
    font-size: 2rem;
    font-weight: 800;
    line-height: 1;
  }}
  .stat-box .s-label {{
    font-size: .75rem;
    color: #94a3b8;
    margin-top: 5px;
    text-transform: uppercase;
    letter-spacing: .8px;
  }}
  .stat-box.green  .s-num {{ color: #4ade80; }}
  .stat-box.amber  .s-num {{ color: #fbbf24; }}
  .stat-box.blue   .s-num {{ color: #60a5fa; }}

  /* ── Card ──────────────────────────────────────── */
  .card {{
    background: #1e293b;
    border: 1px solid #334155;
    border-radius: 16px;
    padding: 24px;
    max-width: 860px;
    margin: 0 auto 22px;
  }}
  .card-title {{
    font-size: 1rem;
    font-weight: 700;
    margin-bottom: 18px;
    color: #f1f5f9;
  }}

  /* ── Table ─────────────────────────────────────── */
  table {{
    width: 100%;
    border-collapse: collapse;
  }}
  thead th {{
    font-size: .72rem;
    text-transform: uppercase;
    letter-spacing: .9px;
    color: #64748b;
    padding: 0 12px 12px;
    text-align: left;
    border-bottom: 1px solid #334155;
  }}
  .data-row td {{
    padding: 12px;
    border-bottom: 1px solid #1e293b;
    font-size: .88rem;
    vertical-align: middle;
  }}
  .data-row:last-child td {{ border-bottom: none; }}
  .data-row:hover {{ background: #263348; }}

  .td-num   {{ color: #475569; width: 42px; text-align: center; }}
  .td-awb   {{ font-family: 'Courier New', monospace; color: #f1f5f9; font-weight: 600; }}
  .td-date  {{ color: #94a3b8; white-space: nowrap; }}

  /* ── Route Badge ───────────────────────────────── */
  .route-badge {{
    display: inline-block;
    background: #0f2744;
    border: 1px solid #1d4ed8;
    color: #93c5fd;
    font-size: .8rem;
    font-weight: 600;
    padding: 3px 10px;
    border-radius: 20px;
    white-space: nowrap;
    font-family: 'Courier New', monospace;
    letter-spacing: .3px;
  }}
  .route-badge[data-na="true"] {{
    background: #1e1e2e;
    border-color: #334155;
    color: #475569;
  }}

  /* ── Not Found Chips ───────────────────────────── */
  .nf-card {{ border-color: #7c3aed44; }}
  .nf-chips {{ display: flex; flex-wrap: wrap; gap: 8px; }}
  .nf-chip {{
    background: #2d1b69;
    border: 1px solid #7c3aed;
    color: #c4b5fd;
    font-size: .78rem;
    font-family: 'Courier New', monospace;
    padding: 4px 10px;
    border-radius: 20px;
  }}

  /* ── Footer ────────────────────────────────────── */
  .footer {{
    text-align: center;
    font-size: .75rem;
    color: #334155;
    margin-top: 40px;
  }}
</style>
</head>
<body>

<!-- HEADER -->
<div class="header">
  <div class="header-logo">✈️</div>
  <h1>Pradeep Cargo &amp; Courier Service</h1>
  <div class="header-sub">Generated: {gen_time} &nbsp;|&nbsp; Stop Date: {stop_str}</div>
</div>

<!-- STATS -->
<div class="stats">
  <div class="stat-box green">
    <div class="s-num">{total_confirmed}</div>
    <div class="s-label">Confirmed</div>
  </div>
  <div class="stat-box amber">
    <div class="s-num">{total_pending}</div>
    <div class="s-label">Pending</div>
  </div>
  <div class="stat-box blue">
    <div class="s-num">{total_excel}</div>
    <div class="s-label">Excel Total</div>
  </div>
</div>

<!-- CONFIRMED TABLE -->
<div class="card">
  <div class="card-title">✅ Confirmed AWBs ({total_confirmed})</div>
  <table>
    <thead>
      <tr>
        <th>#</th>
        <th>AWB Number</th>
        <th>Route</th>
        <th>Flight Date</th>
      </tr>
    </thead>
    <tbody>
      {confirmed_rows}
    </tbody>
  </table>
</div>

<!-- PENDING TABLE -->
<div class="card">
  <div class="card-title">⏳ Pending AWBs ({total_pending})</div>
  <table>
    <thead>
      <tr>
        <th>#</th>
        <th>AWB Number</th>
        <th>Route</th>
        <th>Flight Date</th>
      </tr>
    </thead>
    <tbody>
      {pending_rows}
    </tbody>
  </table>
</div>

<!-- NOT FOUND -->
{not_found_html}

<div class="footer">Pradeep Cargo &amp; Courier Service — Auto Dashboard v11.0</div>

</body>
</html>"""

    # Save
    output_dir  = os.path.join(os.path.expanduser("~"), "PradeepCargo_Reports")
    os.makedirs(output_dir, exist_ok=True)
    filename    = f"dashboard_{datetime.now().strftime('%Y%m%d_%H%M%S')}.html"
    filepath    = os.path.join(output_dir, filename)

    with open(filepath, "w", encoding="utf-8") as f:
        f.write(html)

    return filepath


def open_dashboard(filepath):
    webbrowser.open(f"file://{filepath}")
