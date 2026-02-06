let monitoringMap = {};

function today(){ return new Date().toISOString().slice(0,10); }

function stripBOM(s){
  return s && s.charCodeAt(0) === 0xFEFF ? s.slice(1) : s;
}

function detectDelimiter(line){
  const comma = (line.match(/,/g) || []).length;
  const semi  = (line.match(/;/g) || []).length;
  const tab   = (line.match(/\t/g) || []).length;
  if (semi >= comma && semi >= tab) return ";";
  if (tab >= comma && tab >= semi) return "\t";
  return ",";
}

function splitCSVLine(line, delim){
  const out = [];
  let cur = "";
  let inQuotes = false;

  for (let i=0; i<line.length; i++){
    const ch = line[i];

    if (ch === '"') {
      if (inQuotes && line[i+1] === '"') { cur += '"'; i++; }
      else { inQuotes = !inQuotes; }
      continue;
    }

    if (ch === delim && !inQuotes) {
      out.push(cur);
      cur = "";
      continue;
    }
    cur += ch;
  }
  out.push(cur);
  return out.map(x => x.trim());
}

// NORMALISASI HEADER (biar aman dari spasi/beda huruf)
function normHeader(h){
  return String(h || "")
    .replace(/\u00A0/g, " ")     // non-breaking space
    .trim()
    .replace(/\s+/g, " ")
    .toUpperCase();
}

// cari baris header yang mengandung "FDTID HOTLIST"
function findHeaderRowIndex(lines, delim){
  for (let i=0; i<Math.min(lines.length, 30); i++){
    const cols = splitCSVLine(lines[i], delim).map(normHeader);
    if (cols.includes("FDTID HOTLIST")) return i;
  }
  return -1;
}

function parseCSV(text){
  text = stripBOM(text || "");
  const lines = text.split(/\r?\n/).filter(l => l.trim() !== "");
  if (lines.length === 0) return { rows: [], headersNorm: [] };

  const delim = detectDelimiter(lines[0]);
  const headerIdx = findHeaderRowIndex(lines, delim);
  if (headerIdx === -1) return { rows: [], headersNorm: [] };

  const headersRaw = splitCSVLine(lines[headerIdx], delim);
  const headersNorm = headersRaw.map(normHeader);

  const rows = [];
  for (let i = headerIdx + 1; i < lines.length; i++){
    const cols = splitCSVLine(lines[i], delim);
    if (cols.every(c => c === "")) continue;

    const obj = {};
    headersNorm.forEach((h, idx) => {
      obj[h] = (cols[idx] ?? "").trim();
    });
    rows.push(obj);
  }

  return { rows, headersNorm };
}

function addRow(){
  const tb = document.querySelector("#dataTable tbody");
  const r = tb.rows[0].cloneNode(true);
  r.querySelectorAll("input").forEach(i=>i.value="");
  tb.appendChild(r);
}

function generate(){
  if (!Object.keys(monitoringMap).length){
    alert("Upload file Monitoring dulu (pastikan status loaded > 0 data)");
    return;
  }

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

  const trs = document.querySelectorAll("#dataTable tbody tr");
  for (const tr of trs){
    const hot = tr.querySelector(".hotlist")?.value?.trim() || "";
    if (!hot) continue;

    const m = monitoringMap[hot];
    if (!m){
      alert(`FDTID HOTLIST tidak ditemukan di Monitoring: ${hot}`);
      return;
    }

    const fdt = tr.querySelector(".fdt")?.value?.trim() || "";
    const d5  = tr.querySelector(".draw5")?.value?.trim() || "";
    const lat = tr.querySelector(".lat")?.value?.trim() || "";
    const lng = tr.querySelector(".lng")?.value?.trim() || "";
    const hpD = tr.querySelector(".hpDesign")?.value?.trim() || "";
    const hpR = tr.querySelector(".hpRes")?.value?.trim() || "";
    const biz = tr.querySelector(".bizz")?.value?.trim() || "";

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
      (m["FDT NAME"] ? `${m["FDT NAME"]} ADD HP` : ""),
      lat, lng,
      m["HP SURVEY"] || "",
      m["HP SURVEY"] || "",
      hpD,
      hpD,
      hpR,
      biz,
      "48C",
      "-", "-", "-",
      "AE",
      ""
    ]);
  }

  if (data.length === 1){
    alert("Isi minimal 1 FDTID HOTLIST");
    return;
  }

  const ws = XLSX.utils.aoa_to_sheet(data);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Submit DRM");
  XLSX.writeFile(wb, "RFP_SUBMIT_DRM.xlsx");

  document.getElementById("status").innerText =
    `✅ Berhasil generate (${data.length-1} baris)`;
}

document.addEventListener("DOMContentLoaded", () => {
  const el = document.getElementById("monitoringFile");
  if (!el) return;

  el.addEventListener("change", (e) => {
    const f = e.target.files[0];
    if (!f) return;

    const r = new FileReader();
    r.onload = () => {
      const parsed = parseCSV(r.result);
      const rows = parsed.rows;

      monitoringMap = {};
      rows.forEach(row => {
        const key = (row["FDTID HOTLIST"] || "").trim();
        if (key) monitoringMap[key] = row;
      });

      const count = Object.keys(monitoringMap).length;

      document.getElementById("status").innerText =
        `✅ Monitoring loaded (${count} data)`;

      if (count === 0){
        // tampilkan header yang kebaca biar gampang cek
        const preview = (parsed.headersNorm || []).slice(0, 20).join(" | ");
        alert(
          "Masih 0 data.\n" +
          "Header yang kebaca (20 pertama):\n" + preview
        );
      }
    };
    r.readAsText(f);
  });
});
