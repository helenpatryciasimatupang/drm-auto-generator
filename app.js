let rowCount = 0;
addRow();

function addRow() {
  rowCount++;
  const tbody = document.querySelector("#areaTable tbody");
  const tr = document.createElement("tr");

  tr.innerHTML = `
    <td>${rowCount}</td>
    <td><input class="fdt"></td>
    <td><input class="draw"></td>
  `;
  tbody.appendChild(tr);
}

function generate() {
  const file = document.getElementById("monitoringFile").files[0];
  if (!file) {
    alert("Upload Monitoring dulu");
    return;
  }

  readExcel(file).then(monitoring => {
    const map = {};
    monitoring.forEach(r => map[r["FDT ID"]] = r);

    const rows = document.querySelectorAll("#areaTable tbody tr");
    const output = [];

    rows.forEach((tr, i) => {
      const fdt = tr.querySelector(".fdt").value.trim();
      const draw = tr.querySelector(".draw").value.trim();
      if (!fdt || !draw) return;

      const m = map[fdt];
      if (!m) return;

      output.push({
        "No": i + 1,
        "City Town": m["City Town"],
        "Tenant ID": m["Tenant ID PAPAH"],
        "Permit ID": m["Permit ID PAPAH"],
        "Cluster ID APD": m["Cluster ID"],
        "FDT Coding": fdt + "EXT",
        "Drawing Number LM": `KESA_2_PC_${draw.padStart(5, "0")}_0`,
        "Nama Perumahan/ Kawasan": m["FDT Name"],
        "FDT Name/ Area Name": m["FDT Name"] + " ADD HP",
        "HP Plan": m["HP Survey"],
        "HP Survey": m["HP Survey"]
      });
    });

    if (output.length === 0) {
      alert("Tidak ada data yang match");
      return;
    }

    download(output);
  });
}

function readExcel(file) {
  return new Promise(res => {
    const reader = new FileReader();
    reader.onload = e => {
      const wb = XLSX.read(new Uint8Array(e.target.result), { type: "array" });
      res(XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]));
    };
    reader.readAsArrayBuffer(file);
  });
}

function download(data) {
  const ws = XLSX.utils.json_to_sheet(data);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "RFP");
  XLSX.writeFile(wb, "RFP_FINAL.xlsx");
}
