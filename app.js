// CEK FILE KELOAD
console.log("✅ app.js berhasil diload");

// ============================
// HELPER
// ============================
function getCityFromTenant(tenant) {
  if (tenant.includes("SMG")) return "SEMARANG";
  if (tenant.includes("MDN")) return "MEDAN";
  if (tenant.includes("SLW") || tenant.includes("TGL")) return "TEGAL";
  if (tenant.includes("SBY")) return "SURABAYA";
  return "";
}

function today() {
  return new Date().toISOString().slice(0, 10);
}

// ============================
// MAIN FUNCTION (BUTTON)
// ============================
function processExcel() {
  console.log("🟢 tombol Generate RFP diklik");

  const fileInput = document.getElementById("file");

  if (!fileInput.files.length) {
    alert("❌ Upload file Excel dulu");
    return;
  }

  const reader = new FileReader();

  reader.onload = function (e) {
    const workbook = XLSX.read(e.target.result, { type: "binary" });
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(sheet);

    const output = data.map(row => {
      const tenantRaw = row["Tenant ID"] || "";
      const tenant = tenantRaw.replace(/^LN/, "");

      return {
        ...row,

        // OTOMATIS
        "Vendor RFP": "KESA",
        "Project Type": "NRO B2S Longdrop",
        "Date Input": today(),

        "Tenant ID": "LN" + tenant,
        "Permit ID": "LN" + tenant + "-001",
        "Cluster ID APD": tenant + "-001",

        "City Town": getCityFromTenant(tenant),

        "Type FDT": "48C",
        "Link GDrive": ""
      };
    });

    const newSheet = XLSX.utils.json_to_sheet(output);
    const newWorkbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(newWorkbook, newSheet, "RFP FINAL");

    XLSX.writeFile(newWorkbook, "RFP_FINAL.xlsx");

    document.getElementById("status").innerText =
      "✅ RFP berhasil digenerate";
  };

  reader.readAsBinaryString(fileInput.files[0]);
}
