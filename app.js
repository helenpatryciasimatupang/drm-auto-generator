let monitoringData = [];

document.getElementById("monitoringFile").addEventListener("change", e => {
  const reader = new FileReader();
  reader.onload = evt => {
    const text = evt.target.result;
    const rows = text.split("\n").map(r => r.split(","));
    const headers = rows.shift();

    monitoringData = rows.map(r => {
      let obj = {};
      headers.forEach((h, i) => obj[h.trim()] = (r[i] || "").trim());
      return obj;
    });

    alert("Monitoring CSV loaded: " + monitoringData.length + " data");
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
    const fdtid = inputs[2].value.trim();

    if (!fdtid) return;

    const match = monitoringData.find(m =>
      (m["FDTID"] || "").trim() === fdtid
    );

    output.push({
      "No": idx + 1,
      "Vendor RFP": inputs[0].value,
      "Date Input": new Date().toISOString().slice(0, 10),
      "Project Type": inputs[1].value,
      "City Town": match ? match["CITY_TOWN"] : "",
      "Tenant ID": match ? match["TENANT_ID"] : "",
      "Permit ID": match ? match["PERMIT_ID"] : "",
      "Cluster ID APD": match ? match["CLUSTER_ID_APD"] : "",
      "FDT Coding": match ? (fdtid + (match["EXT"] || "")) : "",
      "Drawing Number LM": inputs[3].value,
      "Nama Perumahan/ Kawasan": match ? match["PERUMAHAN"] : "",
      "FDT Name/ Area Name": match ? match["FDT_NAME"] : "",
      "Latitude": match ? match["LATITUDE"] : "",
      "Longitude": match ? match["LONGITUDE"] : "",
      "HP Plan": inputs[4].value,
      "HP Survey": inputs[5].value,
      "HP Design (Breakdown Permit ID)": inputs[6].value,
      "HP APD All": inputs[7].value,
      "HP Residential": inputs[8].value,
      "Bizz Pass": inputs[9].value,
      "Type FDT": inputs[10].value,
      "Kebutuhan Core BB": inputs[11].value,
      "Jumlah Splitter": inputs[12].value,
      "KM Strand LM (M)": inputs[13].value,
      "CIvil Work": inputs[14].value,
      "Link Gdrive": inputs[15].value
    });
  });

  const ws = XLSX.utils.json_to_sheet(output);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "RFP");
  XLSX.writeFile(wb, "RFP_FINAL.xlsx");
}
