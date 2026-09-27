// ================================================================
//  65 BOYS — Backend Server
//  Express + Google Sheets API + Admin Dashboard
// ================================================================

require("dotenv").config();

const express    = require("express");
const cors       = require("cors");
const path       = require("path");
const { google } = require("googleapis");

const app  = express();
const PORT = process.env.PORT || 3000;

// ── GOOGLE SHEETS SETUP ───────────────────────────────────────────
const SHEET_ID   = process.env.GOOGLE_SHEET_ID;
const SHEET_TAB  = process.env.SHEET_TAB_NAME || "65 BOYS Registrations";
const CLIENT_EMAIL  = process.env.GOOGLE_CLIENT_EMAIL;
const PRIVATE_KEY   = (process.env.GOOGLE_PRIVATE_KEY || "").replace(/\\n/g, "\n");

let sheetsReady = false;
let sheetsClient;

async function initSheets() {
  try {
    const auth = new google.auth.JWT({
      email : CLIENT_EMAIL,
      key   : PRIVATE_KEY,
      scopes: ["https://www.googleapis.com/auth/spreadsheets"]
    });
    await auth.authorize();
    sheetsClient = google.sheets({ version: "v4", auth });
    sheetsReady  = true;
    console.log("✅  Google Sheets connected successfully!");
    await ensureHeaders();
  } catch (err) {
    console.error("❌  Google Sheets connection failed:", err.message);
    console.log("📝  Running WITHOUT Sheets — check your .env credentials.");
  }
}

// Create the header row if the sheet is empty
async function ensureHeaders() {
  try {
    const res = await sheetsClient.spreadsheets.values.get({
      spreadsheetId: SHEET_ID,
      range: `${SHEET_TAB}!A1:F1`
    });
    const rows = res.data.values || [];
    if (!rows.length) {
      await sheetsClient.spreadsheets.values.update({
        spreadsheetId: SHEET_ID,
        range: `${SHEET_TAB}!A1`,
        valueInputOption: "RAW",
        requestBody: {
          values: [["#", "Full Name", "Email Address", "Phone Number", "Registered At (IST)", "Registered At (ISO)"]]
        }
      });
      // Bold + orange background for headers via batchUpdate
      const sheetMeta = await sheetsClient.spreadsheets.get({ spreadsheetId: SHEET_ID });
      const sheet = sheetMeta.data.sheets.find(s => s.properties.title === SHEET_TAB);
      if (sheet) {
        await sheetsClient.spreadsheets.batchUpdate({
          spreadsheetId: SHEET_ID,
          requestBody: {
            requests: [
              {
                repeatCell: {
                  range: { sheetId: sheet.properties.sheetId, startRowIndex: 0, endRowIndex: 1 },
                  cell: {
                    userEnteredFormat: {
                      backgroundColor: { red: 1, green: 0.39, blue: 0 },
                      textFormat: { bold: true, foregroundColor: { red: 1, green: 1, blue: 1 }, fontSize: 11 }
                    }
                  },
                  fields: "userEnteredFormat(backgroundColor,textFormat)"
                }
              },
              {
                updateSheetProperties: {
                  properties: { sheetId: sheet.properties.sheetId, gridProperties: { frozenRowCount: 1 } },
                  fields: "gridProperties.frozenRowCount"
                }
              }
            ]
          }
        });
      }
      console.log("📋  Headers created in Google Sheet!");
    }
  } catch (err) {
    console.log("⚠️  Could not set headers:", err.message);
  }
}

// Append a new row to the sheet
async function appendToSheet(entry) {
  if (!sheetsReady) throw new Error("Google Sheets not connected.");
  const { id, name, email, phone, timestamp, createdAt } = entry;
  await sheetsClient.spreadsheets.values.append({
    spreadsheetId: SHEET_ID,
    range: `${SHEET_TAB}!A:F`,
    valueInputOption: "RAW",
    insertDataOption: "INSERT_ROWS",
    requestBody: {
      values: [[id, name, email, phone, timestamp, createdAt]]
    }
  });
}

// Read all rows from the sheet
async function readFromSheet() {
  if (!sheetsReady) return [];
  const res = await sheetsClient.spreadsheets.values.get({
    spreadsheetId: SHEET_ID,
    range: `${SHEET_TAB}!A2:F`   // skip header row
  });
  const rows = res.data.values || [];
  return rows.map(r => ({
    id        : r[0] || "",
    name      : r[1] || "",
    email     : r[2] || "",
    phone     : r[3] || "",
    timestamp : r[4] || "",
    createdAt : r[5] || ""
  }));
}

// ── HELPERS ───────────────────────────────────────────────────────
function getTimestamp() {
  return new Date().toLocaleString("en-IN", {
    timeZone : "Asia/Kolkata",
    day: "2-digit", month: "short", year: "numeric",
    hour: "2-digit", minute: "2-digit", hour12: true
  });
}

// ── MIDDLEWARE ────────────────────────────────────────────────────
app.use(cors());
app.use(express.json());
app.use(express.urlencoded({ extended: true }));
app.use(express.static(path.join(__dirname, "..")));

// ── API ROUTES ────────────────────────────────────────────────────

// ① Health check
app.get("/api/health", async (req, res) => {
  const data = await readFromSheet().catch(() => []);
  res.json({
    status       : "🔥 65 BOYS Backend is LIVE!",
    sheetsConnected : sheetsReady,
    totalMembers : data.length,
    time         : getTimestamp()
  });
});

// ② Register — POST /api/register
app.post("/api/register", async (req, res) => {
  const { name, email, phone } = req.body;

  // Validation
  const errors = [];
  if (!name  || name.trim().length < 2)
    errors.push("Name is required (minimum 2 characters).");
  if (!email || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email.trim()))
    errors.push("A valid email address is required.");
  if (!phone || phone.trim().length < 6)
    errors.push("A valid phone number is required.");

  if (errors.length)
    return res.status(400).json({ success: false, errors });

  try {
    // Check for duplicate email in sheet
    const existing = await readFromSheet();
    const duplicate = existing.find(r => r.email.toLowerCase() === email.trim().toLowerCase());
    if (duplicate)
      return res.status(409).json({ success: false, errors: ["This email is already registered!"] });

    const newEntry = {
      id        : existing.length + 1,
      name      : name.trim(),
      email     : email.trim().toLowerCase(),
      phone     : phone.trim(),
      timestamp : getTimestamp(),
      createdAt : new Date().toISOString()
    };

    await appendToSheet(newEntry);
    console.log(`✅  Saved to Google Sheets — #${newEntry.id}: ${newEntry.name} <${newEntry.email}>`);

    return res.status(201).json({
      success : true,
      message : `🔥 Welcome to 65 BOYS, ${newEntry.name}! You're officially one of us.`,
      data    : newEntry
    });

  } catch (err) {
    console.error("❌  Registration error:", err.message);
    return res.status(500).json({
      success : false,
      errors  : ["Server error. Please try again later."]
    });
  }
});

// ③ Get all registrations
app.get("/api/registrations", async (req, res) => {
  try {
    const data = await readFromSheet();
    res.json({ success: true, total: data.length, data });
  } catch (err) {
    res.status(500).json({ success: false, errors: [err.message] });
  }
});

// ④ Export as CSV
app.get("/api/export", async (req, res) => {
  try {
    const data = await readFromSheet();
    const csv = [
      ["#", "Full Name", "Email Address", "Phone Number", "Registered At"].join(","),
      ...data.map(r =>
        [r.id, `"${r.name}"`, r.email, r.phone, `"${r.timestamp}"`].join(",")
      )
    ].join("\n");
    res.setHeader("Content-Type", "text/csv");
    res.setHeader("Content-Disposition", 'attachment; filename="65boys-registrations.csv"');
    res.send(csv);
  } catch (err) {
    res.status(500).json({ success: false, errors: [err.message] });
  }
});

// ⑤ Admin Dashboard
app.get("/admin", (req, res) => res.send(adminHTML()));

// ── START ─────────────────────────────────────────────────────────
initSheets().then(() => {
  app.listen(PORT, () => {
    console.log("=".repeat(52));
    console.log("  🔥  65 BOYS Backend Server — RUNNING");
    console.log("=".repeat(52));
    console.log(`  🌐  Website  →  http://localhost:${PORT}`);
    console.log(`  📋  Admin    →  http://localhost:${PORT}/admin`);
    console.log(`  🔌  API      →  http://localhost:${PORT}/api/health`);
    console.log(`  📊  Sheets   →  ${sheetsReady ? "✅ Connected" : "❌ Not connected — check .env"}`);
    console.log("=".repeat(52));
  });
});

// ================================================================
//  ADMIN DASHBOARD
// ================================================================
function adminHTML() {
return `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8"/>
  <meta name="viewport" content="width=device-width,initial-scale=1.0"/>
  <title>65 BOYS — Admin Dashboard</title>
  <style>
    *,*::before,*::after{box-sizing:border-box;margin:0;padding:0}
    body{font-family:'Segoe UI',system-ui,sans-serif;background:#07070d;color:#ebebf5;min-height:100vh}
    ::-webkit-scrollbar{width:6px}::-webkit-scrollbar-thumb{background:#ff6400;border-radius:3px}
    a{color:inherit;text-decoration:none}

    nav{background:rgba(255,255,255,0.04);border-bottom:1px solid rgba(255,100,0,0.2);
      padding:1rem 2rem;display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:0.8rem}
    .logo{font-size:1.4rem;font-weight:900;
      background:linear-gradient(135deg,#ff6400,#ffb300);
      -webkit-background-clip:text;-webkit-text-fill-color:transparent;background-clip:text}
    .sheets-badge{display:inline-flex;align-items:center;gap:0.4rem;
      background:rgba(0,200,100,0.1);border:1px solid rgba(0,200,100,0.25);
      color:#00c878;padding:0.3rem 0.9rem;border-radius:50px;font-size:0.75rem;font-weight:700}
    .sheets-badge.off{background:rgba(255,60,60,0.1);border-color:rgba(255,60,60,0.25);color:#ff6060}
    .nav-right{display:flex;gap:0.7rem;flex-wrap:wrap}
    .btn{padding:0.55rem 1.2rem;border-radius:50px;font-size:0.82rem;font-weight:700;border:none;cursor:pointer;transition:transform .2s}
    .btn-orange{background:linear-gradient(135deg,#ff6400,#ff9500);color:#fff;box-shadow:0 3px 14px rgba(255,100,0,0.3)}
    .btn-orange:hover{transform:translateY(-2px)}
    .btn-ghost{background:rgba(255,255,255,0.06);color:#ebebf5;border:1px solid rgba(255,255,255,0.12)}
    .btn-ghost:hover{border-color:#ff6400}

    main{padding:2rem}

    /* STATS */
    .stats-row{display:grid;grid-template-columns:repeat(4,1fr);gap:1rem;margin-bottom:2rem}
    .sc{background:rgba(255,255,255,0.04);border:1px solid rgba(255,100,0,0.15);border-radius:16px;
      padding:1.5rem;text-align:center;transition:border-color .3s}
    .sc:hover{border-color:#ff6400}
    .sc-n{font-size:2.5rem;font-weight:900;line-height:1;
      background:linear-gradient(135deg,#ff6400,#ffb300);
      -webkit-background-clip:text;-webkit-text-fill-color:transparent;background-clip:text}
    .sc-l{font-size:0.7rem;color:rgba(235,235,245,0.45);letter-spacing:2px;text-transform:uppercase;margin-top:0.4rem}

    /* TOOLBAR */
    .toolbar{display:flex;align-items:center;justify-content:space-between;margin-bottom:1rem;flex-wrap:wrap;gap:0.8rem}
    .toolbar h2{font-size:1.15rem;font-weight:800;display:flex;align-items:center;gap:0.5rem}
    .search{background:rgba(255,255,255,0.05);border:1px solid rgba(255,255,255,0.1);
      border-radius:10px;padding:0.6rem 1rem;color:#ebebf5;font-size:0.9rem;outline:none;
      width:260px;transition:border-color .2s}
    .search:focus{border-color:#ff6400}
    .search::placeholder{color:rgba(235,235,245,0.3)}

    /* TABLE */
    .tw{background:rgba(255,255,255,0.03);border:1px solid rgba(255,100,0,0.15);
      border-radius:18px;overflow:hidden;overflow-x:auto}
    table{width:100%;border-collapse:collapse;min-width:720px}
    thead tr{background:linear-gradient(90deg,rgba(255,100,0,0.2),rgba(255,160,0,0.1))}
    th{padding:1rem 1.2rem;text-align:left;font-size:0.72rem;font-weight:800;letter-spacing:2px;
      text-transform:uppercase;color:#ff8c40;white-space:nowrap}
    tbody tr{border-top:1px solid rgba(255,255,255,0.05);transition:background .2s}
    tbody tr:hover{background:rgba(255,100,0,0.05)}
    td{padding:1rem 1.2rem;font-size:0.9rem}
    td.dim{color:rgba(235,235,245,0.5);font-size:0.82rem}
    .badge{display:inline-block;background:rgba(255,100,0,0.12);border:1px solid rgba(255,100,0,0.25);
      color:#ff8c40;padding:0.2rem 0.7rem;border-radius:50px;font-size:0.72rem;font-weight:700}
    #empty td{text-align:center;padding:3rem;color:rgba(235,235,245,0.3);font-style:italic}

    /* SHEETS LINK */
    .sheet-link{display:inline-flex;align-items:center;gap:0.5rem;
      background:rgba(52,168,83,0.1);border:1px solid rgba(52,168,83,0.25);
      color:#34a853;padding:0.5rem 1.1rem;border-radius:50px;font-size:0.82rem;font-weight:700;
      transition:background .2s}
    .sheet-link:hover{background:rgba(52,168,83,0.18)}

    /* TOAST */
    #toast{position:fixed;top:5rem;right:1.5rem;z-index:999;
      padding:0.9rem 1.4rem;border-radius:12px;font-weight:700;font-size:0.88rem;display:none;max-width:320px}
    #toast.s{background:rgba(0,200,100,0.15);border:1px solid rgba(0,200,100,0.35);color:#00d472}
    #toast.e{background:rgba(255,60,60,0.15);border:1px solid rgba(255,60,60,0.35);color:#ff5a5a}

    /* LAST UPDATED */
    .last-upd{font-size:0.75rem;color:rgba(235,235,245,0.3);text-align:right;margin-top:0.7rem}

    @media(max-width:700px){.stats-row{grid-template-columns:1fr 1fr}}
  </style>
</head>
<body>
<div id="toast"></div>

<nav>
  <div style="display:flex;align-items:center;gap:1rem;flex-wrap:wrap">
    <div class="logo">65 BOYS — Admin</div>
    <span class="sheets-badge" id="sheetsBadge">⏳ Checking Sheets…</span>
  </div>
  <div class="nav-right">
    <button class="btn btn-ghost" onclick="load()">🔄 Refresh</button>
    <a href="/api/export" class="btn btn-ghost">📥 Export CSV</a>
    <a id="sheetBtn" href="#" target="_blank" class="btn btn-ghost" style="display:none">
      📊 Open Google Sheet
    </a>
    <a href="/" class="btn btn-orange">🌐 Website</a>
  </div>
</nav>

<main>
  <div class="stats-row">
    <div class="sc"><div class="sc-n" id="sTotal">—</div><div class="sc-l">Total Members</div></div>
    <div class="sc"><div class="sc-n" id="sToday">—</div><div class="sc-l">Joined Today</div></div>
    <div class="sc"><div class="sc-n" id="sLatest">—</div><div class="sc-l">Latest Member</div></div>
    <div class="sc"><div class="sc-n">🔥</div><div class="sc-l">Squad Status</div></div>
  </div>

  <div class="toolbar">
    <h2>📋 All Registrations</h2>
    <div style="display:flex;gap:0.7rem;align-items:center;flex-wrap:wrap">
      <input class="search" id="searchBox" placeholder="🔍  Search name, email or phone…" oninput="filter()"/>
    </div>
  </div>

  <div class="tw">
    <table>
      <thead>
        <tr>
          <th>#</th>
          <th>Full Name</th>
          <th>Email Address</th>
          <th>Phone Number</th>
          <th>Registered At</th>
        </tr>
      </thead>
      <tbody id="tbody">
        <tr id="empty"><td colspan="5">Loading data from Google Sheets…</td></tr>
      </tbody>
    </table>
  </div>
  <div class="last-upd" id="lastUpd"></div>
</main>

<script>
  let allData = [];

  function toast(msg, type){
    const t=document.getElementById('toast');
    t.innerHTML=msg; t.className=type; t.style.display='block';
    setTimeout(()=>t.style.display='none', 4500);
  }

  async function load(){
    try{
      // health check (tells us if sheets is connected)
      const h = await fetch('/api/health').then(r=>r.json());
      const badge = document.getElementById('sheetsBadge');
      if(h.sheetsConnected){
        badge.textContent='✅ Google Sheets Connected';
        badge.className='sheets-badge';
      } else {
        badge.textContent='❌ Sheets Not Connected';
        badge.className='sheets-badge off';
      }

      const res  = await fetch('/api/registrations').then(r=>r.json());
      allData = res.data || [];
      render(allData);
      stats(allData);
      document.getElementById('lastUpd').textContent = 'Last refreshed: ' + new Date().toLocaleTimeString('en-IN');
    } catch(e){
      toast('❌ Could not reach backend. Is server running?','e');
    }
  }

  function stats(data){
    document.getElementById('sTotal').textContent = data.length;
    const today = new Date().toDateString();
    const t = data.filter(r=> r.createdAt && new Date(r.createdAt).toDateString()===today).length;
    document.getElementById('sToday').textContent = t;
    document.getElementById('sLatest').textContent = data.length ? data[data.length-1].name.split(' ')[0] : '—';
  }

  function render(data){
    const b = document.getElementById('tbody');
    if(!data.length){
      b.innerHTML='<tr id="empty"><td colspan="5">No registrations yet. Share your website! 🚀</td></tr>';
      return;
    }
    b.innerHTML = data.map((r,i)=>\`
      <tr>
        <td><span class="badge">#\${r.id||i+1}</span></td>
        <td><strong>\${esc(r.name)}</strong></td>
        <td class="dim">\${esc(r.email)}</td>
        <td class="dim">\${esc(r.phone)}</td>
        <td class="dim">\${esc(r.timestamp)}</td>
      </tr>
    \`).join('');
  }

  function filter(){
    const q=document.getElementById('searchBox').value.toLowerCase();
    render(allData.filter(r=>
      r.name.toLowerCase().includes(q)||
      r.email.toLowerCase().includes(q)||
      r.phone.toLowerCase().includes(q)
    ));
  }

  function esc(s){ return String(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;'); }

  load();
  setInterval(load, 30000);
</script>
</body>
</html>`;
}
