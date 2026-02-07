let monitoringData = [];

document.getElementById("monitoringFile").addEventListener("change", e => {
  const reader = new FileReader();
  reader.onload = evt => {
    const text = evt.target.result.trim();
    const rows = text.split(/\r?\n/).map(r => r.split(/,(?=(?:[^"]*"[^"]*")*[^"]*$)/));

    const headers = rows.shift().map(h => h.replace(/"/g, "").trim());

    monitoringData = rows.map(r => {
      let obj = {};
      headers.forEach((h, i) => {
        obj[h] = (r[i] || "").replace(/"/g, "").trim();
      });
      return obj;
    });

    alert(`Monitoring CSV loaded: ${monitoringData.length} baris`);
  };
  reader.readAsText(e.target.files[0]);
});

function addRow() {
  const tbody = document.getElementById("areaBody");
  const row = tbody.rows[0].cloneNode(true);
  row.cells[0].innerText = tbody.rows.length + 1;
  row.querySelectorAll("input").forEach(i => i.value = "");
  tbody.appendChild(row);
}

function generateExcel() {
  if (!monitoringData.length) {
    alert("Upload Monitoring CSV dulu!");
    return;
  }

  const output = [];
  const rows = document.querySelectorAll("#areaBody tr");

  rows.forEach((tr, idx) => {
    const i = tr.querySelectorAll("input");
    const fdtid = i[0].value.trim();
    if (!fdtid) return;

    const m = monitoringData.find(d =>
      (d["FDTID HOTLIST"] || "").toUpperCase() === fdtid.toUpperCase()
    );

    output.push({
      "No": idx + 1,
      "Vendor RFP": "KESA",
      "Date Input": new Date().toISOString().slice(0, 10),
      "Project Type": "NRO B2S Longdrop",
      "City Town": m?.["City Town"] || "",
      "Tenant ID": m?.["Tenant ID PAPAH"] || "",
      "Permit ID": m?.["Permit ID PAPAH"] || "",
      "Cluster ID APD": m?.["Cluster ID"] || "",
      "FDT Coding": fdtid,
      "Drawing Number LM": i[1].value,
      "Nama Perumahan/ Kawasan": m?.["Nama Perumahan"] || "",
      "FDT Name/ Area Name": m?.["FDT Name"] || "",
      "Latitude": m?.["Latitude"] || "",
      "Longitude": m?.["Longitude"] || "",
      "HP Plan": m?.["HP Survey"] || "",
      "HP Survey": m?.["HP Survey"] || "",
      "HP Design (Breakdown Permit ID)": i[2].value,
      "HP APD All": i[3].value,
      "HP Residential": i[4].value,
      "Bizz Pass": i[5].value,
      "Type FDT": "48C",
      "Kebutuhan Core BB": "-",
      "Jumlah Splitter": "-",
      "KM Strand LM (M)": "-",
      "Civil Work": "AE",
      "Link Gdrive": ""
    });
  });

  if (!output.length) {
    alert("Tidak ada data yang berhasil diproses");
    return;
  }

  const ws = XLSX.utils.json_to_sheet(output);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "RFP");

  XLSX.writeFile(wb, "RFP_FINAL.xlsx");
}
