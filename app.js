let monitoringData = [];

document.getElementById("monitoringFile").addEventListener("change", e => {
  const reader = new FileReader();
  reader.onload = evt => {
    const rows = evt.target.result.split("\n").map(r => r.split(","));
    const headers = rows.shift().map(h => h.trim());

    monitoringData = rows.map(r => {
      let obj = {};
      headers.forEach((h, i) => obj[h] = (r[i] || "").trim());
      return obj;
    });

    alert("Monitoring CSV loaded: " + monitoringData.length + " baris");
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
    const inputs = tr.querySelectorAll("input");
    const fdtid = inputs[0].value.trim();
    if (!fdtid) return;

    const match = monitoringData.find(m => (m["FDTID"] || "") === fdtid);

    output.push({
      "No": idx + 1,
      "Vendor RFP": "KESA",
      "Date Input": new Date().toISOString().slice(0,10),
      "Project Type": "NRO B2S Longdrop",
      "City Town": match?.CITY_TOWN || "",
      "Tenant ID": match?.TENANT_ID || "",
      "Permit ID": match?.PERMIT_ID || "",
      "Cluster ID APD": match?.CLUSTER_ID_APD || "",
      "FDT Coding": fdtid,
      "Drawing Number LM": inputs[1].value,
      "Nama Perumahan/ Kawasan": match?.PERUMAHAN || "",
      "FDT Name/ Area Name": match?.FDT_NAME || "",
      "Latitude": match?.LATITUDE || "",
      "Longitude": match?.LONGITUDE || "",
      "HP Plan": match?.HP_SURVEY || "",
      "HP Survey": match?.HP_SURVEY || "",
      "HP Design (Breakdown Permit ID)": inputs[2].value,
      "HP APD All": inputs[3].value,
      "HP Residential": inputs[4].value,
      "Bizz Pass": inputs[5].value,
      "Type FDT": "48C",
      "Kebutuhan Core BB": "-",
      "Jumlah Splitter": "-",
      "KM Strand LM (M)": "-",
      "Civil Work": "AE",
      "Link Gdrive": ""
    });
  });

  const ws = XLSX.utils.json_to_sheet(output);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "RFP");
  XLSX.writeFile(wb, "RFP_FINAL.xlsx");
}
