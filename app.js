let monitoringData = [];

document.getElementById("monitoringFile").addEventListener("change", e => {
  const reader = new FileReader();
  reader.onload = evt => {
    const lines = evt.target.result.trim().split(/\r?\n/);
    const headers = lines.shift().split(",");

    monitoringData = lines.map(l => {
      const v = l.split(",");
      let o = {};
      headers.forEach((h, i) => o[h.trim()] = (v[i] || "").trim());
      return o;
    });

    alert("Monitoring CSV berhasil dimuat: " + monitoringData.length + " baris");
  };
  reader.readAsText(e.target.files[0]);
});

function addRow() {
  const body = document.getElementById("areaBody");
  const r = body.rows[0].cloneNode(true);
  r.cells[0].innerText = body.rows.length + 1;
  r.querySelectorAll("input").forEach(i => i.value = "");
  body.appendChild(r);
}

function generateExcel() {
  if (!monitoringData.length) {
    alert("Upload Monitoring CSV dulu");
    return;
  }

  const rows = document.querySelectorAll("#areaBody tr");
  let output = [];

  rows.forEach((tr, i) => {
    const c = tr.querySelectorAll("input");
    const fdtid = c[0].value.trim();
    if (!fdtid) return;

    const m = monitoringData.find(d => d["FDTID HOTLIST"] === fdtid);

    output.push({
      "No": i + 1,
      "Vendor RFP": "KESA",
      "Date Input": new Date().toISOString().slice(0, 10),
      "Project Type": "NRO B2S Longdrop",
      "City Town": m?.["City Town"] || "",
      "Tenant ID": m?.["Tenant ID PAPAH"] || "",
      "Permit ID": m?.["Permit ID PAPAH"] || "",
      "Cluster ID APD": m?.["Cluster ID"] || "",
      "FDT Coding": fdtid,
      "Drawing Number LM": c[1].value,
      "HP Plan": m?.["HP Survey"] || "",
      "HP Survey": m?.["HP Survey"] || "",
      "HP Design": c[2].value,
      "HP APD All": c[3].value,
      "HP Residential": c[4].value,
      "Bizz Pass": c[5].value,
      "Type FDT": "48C",
      "Kebutuhan Core BB": "-",
      "Jumlah Splitter": "-",
      "KM Strand LM (M)": "-",
      "Civil Work": "AE",
      "Link Gdrive": ""
    });
  });

  if (!output.length) {
    alert("Tidak ada data untuk digenerate");
    return;
  }

  const ws = XLSX.utils.json_to_sheet(output);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "RFP");

  XLSX.writeFile(wb, "RFP_FINAL.xlsx");
}
