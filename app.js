// ================= GLOBAL STORAGE =================
let monitoringMap = {};   // key: FDTID HOTLIST
let masterFDTMap  = {};   // key: FDT Coding (misal FOT13700)

// ================= UTIL =================
function today() {
  return new Date().toISOString().slice(0, 10);
}

function stripBOM(s) {
  return s && s.charCodeAt(0) === 0xFEFF ? s.slice(1) : s;
}

function detectDelimiter(line) {
  const c = (line.match(/,/g) || []).length;
  const s = (line.match(/;/g) || []).length;
  const t = (line.match(/\t/g) || []).length;
  if (s >= c && s >= t) return ";";
  if (t >= c && t >= s) return "\t";
  return ",";
}

function splitCSV(line, d) {
  const out = [];
  let cur = "";
  let q = false;
  for (let i = 0; i < line.length; i++) {
    const ch = line[i];
    if (ch === '"') { q = !q; continue; }
    if (ch === d && !q) { out.push(cur); cur = ""; continue; }
    cur += ch;
  }
  out.push(cur);
  return out.map(x => x.trim());
}

function norm(h) {
  return String(h || "")
    .replace(/\u00A0/g, " ")
    .trim()
    .replace(/\s+/g, " ")
    .toUpperCase();
}

// ================= CSV PARSER (Monitoring) =================
function parseCSV(text) {
  text = stripBOM(text || "");
  const lines = text.split(/\r?\n/).filter(l => l.trim());
  if (!lines.length) return [];

  const d = detectDelimiter(lines[0]);
  let headerIdx = -1;

  for (let i = 0; i < Math.min(lines.length, 20); i++) {
    const cols = splitCSV(lines[i], d).map(norm);
    if (cols.includes("FDTID HOTLIST")) {
      headerIdx = i;
      break;
    }
  }
  if (headerIdx < 0) return [];

  const headers = splitCSV(lines[headerIdx], d).map(norm);
  const rows = [];

  for (let i = headerIdx + 1; i < lines.length; i++) {
    const cols = splitCSV(lines[i], d);
    if (cols.every(c => !c)) continue;
    const o = {};
    headers.forEach((h, idx) => o[h] = cols[idx] || "");
    rows.push(o);
  }
  return rows;
}

// ================= FAT → FDT =================
function fatToFdt(code) {
  return String(code || "").replace(/S\d+A\d+/i, "");
}

// ================= LOAD MONITORING =================
document.getElementById("monitoringFile").addEventListener("change", e => {
  const f = e.target.files[0];
  if (!f) return;

  const r = new FileReader();
  r.onload = () => {
    monitoringMap = {};
    const rows = parseCSV(r.result);
    rows.forEach(row => {
      const key = (row["FDTID HOTLIST"] || "").trim();
      if (key) monitoringMap[key] = row;
    });
    document.getElementById("status").innerText =
      `✅ Monitoring loaded (${Object.keys(monitoringMap).length} data)`;
  };
  r.readAsText(f);
});

// ================= LOAD MASTER POP UP (EXCEL) =================
document.getElementById("masterPopupFile").addEventListener("change", e => {
  const f = e.target.files[0];
  if (!f) return;

  const r = new FileReader();
  r.onload = evt => {
    const wb = XLSX.read(evt.target.result, { type: "binary" });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(ws, { defval: "" });

    masterFDTMap = {};

    rows.forEach(row => {
      // Ambil kode node / name (ini FIX sesuai file kamu)
      const code =
        row.NAME ||
        row.FDT ||
        row.FDT_CODE ||
        row.NODE ||
        row.NODE_CODE ||
        "";

      if (!code) return;

      const fdt = fatToFdt(code);

      if (!masterFDTMap[fdt]) {
        masterFDTMap[fdt] = {
          lat: null,
          lng: null,
          hpDesign: 0,
          hpHome: 0,
          hpBiz: 0
        };
      }

      // === Koordinat FDT (kode pendek, tanpa SxxAxx)
      if (!String(code).match(/S\d+A\d+/i)) {
        masterFDTMap[fdt].lat =
          row.LATITUDE || row.Latitude || row.Y || "";
        masterFDTMap[fdt].lng =
          row.LONGITUDE || row.Longitude || row.X || "";
      }

      // === Hitungan HP
      const houseNum = Number(row.HOUSE_NUMBER || 1);
      masterFDTMap[fdt].hpDesign += houseNum;

      const type = String(row.TYPE || row.HP_TYPE || "").toUpperCase();
      if (type === "HOME") masterFDTMap[fdt].hpHome++;
      if (type === "HOME-BIZ" || type === "BIZ") masterFDTMap[fdt].hpBiz++;
    });

    document.getElementById("status").innerText +=
      ` | Master Pop Up loaded (${Object.keys(masterFDTMap).length} FDT)`;
  };
  r.readAsBinaryString(f);
});

// ================= UI =================
function addRow() {
  const tb = document.querySelector("#dataTable tbody");
  const r = tb.rows[0].cloneNode(true);
  r.querySelectorAll("input").forEach(i => i.value = "");
  tb.appendChild(r);
}

// ================= GENERATE RFP =================
function generate() {
  const headers = [
    "No","Vendor RFP","Date Input","Project Type","City Town",
    "Tenant ID","Permit ID","Cluster ID APD",
    "FDT Coding","Drawing Number LM",
    "Nama Perumahan/ Kawasan","FDT Name/ Area Name",
    "Latitude","Longitude",
    "HP Plan","HP Survey",
    "HP Design (Breakdown Permit ID)","HP APD All",
    "HP Residential","Bizz Pass",
    "Type FDT","Kebutuhan Core BB","Jumlah Splitter","KM Strand LM (M)",
    "CIvil Work","Link Gdrive"
  ];

  const data = [headers];
  let no = 1;

  document.querySelectorAll("#dataTable tbody tr").forEach(tr => {
    const hot = tr.querySelector(".hotlist").value.trim();
    const fdt = tr.querySelector(".fdt").value.trim();
    const d5  = tr.querySelector(".draw5").value.trim();

    if (!hot || !fdt) return;

    const m = monitoringMap[hot];
    const fdtData = masterFDTMap[fdt];

    if (!m) {
      alert(`FDTID HOTLIST tidak ditemukan: ${hot}`);
      return;
    }
    if (!fdtData) {
      alert(`FDT Coding tidak ditemukan di Master Pop Up: ${fdt}`);
      return;
    }

    data.push([
      no++,
      "KESA",
      today(),
      "NRO B2S Longdrop",
      m["CITY TOWN"] || "",
      m["TENANT ID PAPAH"] || "",
      m["PERMIT ID PAPAH"] || "",
      m["CLUSTER ID"] || "",
      fdt,
      d5 ? `KESA_2_PC_${d5}_0` : "",
      m["FDT NAME"] || "",
      m["FDT NAME"] ? `${m["FDT NAME"]} ADD HP` : "",
      fdtData.lat,
      fdtData.lng,
      m["HP SURVEY"] || "",
      m["HP SURVEY"] || "",
      fdtData.hpDesign,
      fdtData.hpDesign,
      fdtData.hpHome,
      fdtData.hpBiz,
      "48C",
      "-","-","-",
      "AE",
      ""
    ]);
  });

  if (data.length === 1) {
    alert("Tidak ada data valid untuk digenerate");
    return;
  }

  const ws = XLSX.utils.aoa_to_sheet(data);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Submit DRM");
  XLSX.writeFile(wb, "RFP_SUBMIT_DRM.xlsx");

  document.getElementById("status").innerText =
    `✅ RFP berhasil digenerate (${data.length - 1} baris)`;
}
