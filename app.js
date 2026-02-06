let monitoringMap = {};

function today(){ return new Date().toISOString().slice(0,10); }

// buang BOM (kadang CSV dari Excel ada \ufeff)
function stripBOM(s){
  return s && s.charCodeAt(0) === 0xFEFF ? s.slice(1) : s;
}

// deteksi delimiter paling mungkin: ; , \t
function detectDelimiter(line){
  const comma = (line.match(/,/g) || []).length;
  const semi  = (line.match(/;/g) || []).length;
  const tab   = (line.match(/\t/g) || []).length;
  if (semi >= comma && semi >= tab) return ";";
  if (tab >= comma && tab >= semi) return "\t";
  return ",";
}

// split 1 baris CSV dengan support "quote"
function splitCSVLine(line, delim){
  const out = [];
  let cur = "";
  let inQuotes = false;

  for (let i=0; i<line.length; i++){
    const ch = line[i];

    if (ch === '"') {
      // handle double quotes "" inside quotes
      if (inQuotes && line[i+1] === '"') {
        cur += '"';
        i++;
      } else {
        inQuotes = !inQuotes;
      }
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

// parse CSV text -> array of objects
function parseCSV(text){
  text = stripBOM(text || "");
  const lines = text.split(/\r?\n/).filter(l => l.trim() !== "");
  if (lines.length === 0) return [];

  const delim = detectDelimiter(lines[0]);
  const headers = splitCSVLine(lines[0], delim).map(h => h.trim());

  const rows = [];
  for (let i=1; i<lines.length; i++){
    const cols = splitCSVLine(lines[i], delim);
    if (cols.every(c => c === "")) continue;

    const obj = {};
    headers.forEach((h, idx) => {
      obj[h] = (cols[idx] ?? "").trim();
    });
    rows.push(obj);
  }
  return rows;
}

// load monitoring file
function bindMonitoringLoader(){
  const el = document.getElementById("monitoringFile");
  if (!el) return;

  el.addEventListener("change", (e) => {
    const f = e.target.files[0];
    if (!f) return;

    const r = new FileReader();
    r.onload = () => {
      const rows = parseCSV(r.result);

      monitoringMap = {};
      rows.forEach(row => {
        // header persis yang kamu kasih:
        // FDTID HOTLIST, City Town, Tenant ID PAPAH, Permit ID PAPAH, Cluster ID, FDT Name, HP Survey
        const key = (row["FDTID HOTLIST"] || "").trim();
        if (key) monitoringMap[key] = row;
      });

      const count = Object.keys(monitoringMap).length;
      document.getElementById("status").innerText =
        `✅ Monitoring loaded (${count} data)`;
      if (count === 0) {
        alert("File kebaca, tapi kolom 'FDTID HOTLIST' tidak ketemu / delimiter beda. Parser sudah auto, tapi pastikan headernya persis 'FDTID HOTLIST'.");
      }
    };
    r.readAsText(f);
  });
}

// === UI helpers ===
function addRow(){
  const tb = document.querySelector("#dataTable tbody");
  const r = tb.rows[0].cloneNode(true);
  r.querySelectorAll("input").forEach(i=>i.value="");
  tb.appendChild(r);
}

function generate(){
  const keys = Object.keys(monitoringMap);
  if (!keys.length){
    alert("Upload file Monitoring dulu (pastikan status Monitoring loaded > 0 data)");
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
      (m["City Town"] || "").trim(),
      (m["Tenant ID PAPAH"] || "").trim(),
      (m["Permit ID PAPAH"] || "").trim(),
      (m["Cluster ID"] || "").trim(),
      fdt,
      d5 ? `KESA_2_PC_${d5}_0` : "",
      (m["FDT Name"] || "").trim(),
      (m["FDT Name"] || "").trim() ? `${(m["FDT Name"] || "").trim()} ADD HP` : "",
      lat, lng,
      (m["HP Survey"] || "").trim(),
      (m["HP Survey"] || "").trim(),
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

// bind loader setelah DOM siap
document.addEventListener("DOMContentLoaded", bindMonitoringLoader);
