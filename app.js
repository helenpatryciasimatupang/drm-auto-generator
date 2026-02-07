function generate() {
  const monitoringFile = document.getElementById("monitoringFile").files[0];
  const hotlistFile = document.getElementById("hotlistFile").files[0];

  if (!monitoringFile || !hotlistFile) {
    alert("Upload Monitoring & Hotlist dulu");
    return;
  }

  Promise.all([
    readExcel(monitoringFile),
    readExcel(hotlistFile)
  ]).then(([monitoring, hotlist]) => {
    const monitoringMap = {};

    // 🔑 MAP MONITORING BY FDT ID
    monitoring.forEach(row => {
      monitoringMap[row["FDT ID"]] = row;
    });

    const output = [];

    hotlist.forEach((row, i) => {
      const fdtId = row["FDT ID HOTLIST"];
      const drawNum = row["Drawing Number"];

      const m = monitoringMap[fdtId];
      if (!m) return;

      output.push({
        "No": i + 1,
        "City Town": m["City Town"],
        "Tenant ID": m["Tenant ID PAPAH"],
        "Permit ID": m["Permit ID PAPAH"],
        "Cluster ID APD": m["Cluster ID"],
        "FDT Coding": fdtId + "EXT",
        "Drawing Number LM": "KESA_2_PC_" + String(drawNum).padStart(5, "0") + "_0",
        "Nama Perumahan / Kawasan": m["FDT Name"],
        "FDT Name / Area Name": m["FDT Name"] + " ADD HP",
        "HP Plan": m["HP Survey"],
        "HP Survey": m["HP Survey"]
      });
    });

    downloadExcel(output);
  });
}

function readExcel(file) {
  return new Promise(resolve => {
    const reader = new FileReader();
    reader.onload = e => {
      const data = new Uint8Array(e.target.result);
      const wb = XLSX.read(data, { type: "array" });
      const sheet = wb.Sheets[wb.SheetNames[0]];
      resolve(XLSX.utils.sheet_to_json(sheet));
    };
    reader.readAsArrayBuffer(file);
  });
}

function downloadExcel(data) {
  const ws = XLSX.utils.json_to_sheet(data);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "RFP");
  XLSX.writeFile(wb, "RFP_FINAL.xlsx");
}
