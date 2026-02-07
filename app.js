let monitoringData = [];
let masterData = [];

document.getElementById("monitoringFile").addEventListener("change", e => {
  readFile(e.target.files[0], data => monitoringData = data);
});

document.getElementById("masterFile").addEventListener("change", e => {
  readFile(e.target.files[0], data => masterData = data);
});

function readFile(file, cb) {
  const reader = new FileReader();
  reader.onload = e => {
    const wb = XLSX.read(e.target.result, { type: "binary" });
    const ws = wb.Sheets[wb.SheetNames[0]];
    cb(XLSX.utils.sheet_to_json(ws));
  };
  reader.readAsBinaryString(file);
}

function generate() {
  if (!monitoringData.length || !masterData.length) {
    alert("File belum lengkap");
    return;
  }

  const today = new Date().toISOString().slice(0,10);
  const fdtIds = document.getElementById("fdtList").value.split("\n").map(v => v.trim()).filter(Boolean);

  const output = [];

  fdtIds.forEach((fdt, i) => {
    const mon = monitoringData.find(r => r["FDTID HOTLIST"] == fdt);
    if (!mon) return;

    const masterRows = masterData.filter(r => r["FDT_CODE"] == fdt);

    const lat = masterRows[0]?.["Coordinate (Lat) NEW"] || "";
    const lon = masterRows[0]?.["Coordinate (Long) NEW"] || "";

    const hpDesign = masterRows.length;
    const hpRes = masterRows.filter(r => r["HOME/HOME-BIZ"] === "HOME").length;
    const bizz = masterRows.filter(r => r["HOME/HOME-BIZ"] === "HOME-BIZ").length;

    output.push({
      "No": i+1,
      "Vendor RFP": "KESA",
      "Date Input": today,
      "Project Type": "NRO B2S Longdrop",
      "City Town": mon["City Town"],
      "Tenant ID": mon["Tenant ID PAPAH"],
      "Permit ID": mon["Permit ID PAPAH"],
      "Cluster ID APD": mon["Cluster ID"],
      "FDT Coding": fdt + "EXT",
      "Drawing Number LM": "KESA_2_PC_00000_0",
      "Nama Perumahan/ Kawasan": mon["FDT Name"],
      "FDT Name/ Area Name": mon["FDT Name"] + " ADD HP",
      "Latitude": lat,
      "Longitude": lon,
      "HP Plan": mon["HP Survey"],
      "HP Survey": mon["HP Survey"],
      "HP Design (Breakdown Permit ID)": hpDesign,
      "HP APD All": hpDesign,
      "HP Residential": hpRes,
      "Bizz Pass": bizz,
      "Type FDT": "48C",
      "Kebutuhan Core BB": "-",
      "Jumlah Splitter": "-",
      "KM Strand LM (M)": "-",
      "CIvil Work": "AE",
      "Link Gdrive": ""
    });
  });

  const ws = XLSX.utils.json_to_sheet(output);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "SUBMIT_DRM_LM");
  XLSX.writeFile(wb, "RFP_SUBMIT_DRM_LM.xlsx");

  document.getElementById("status").innerText = "✅ Excel berhasil dibuat";
}
