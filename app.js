console.log("APP.JS LOADED");

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

function processExcel() {
  console.log("BUTTON CLICKED");

  const fileInput = document.getElementById("file");
  if (!fileInput.files.length) {
    alert("❌ Upload file Excel dulu");
    return;
  }

  const reader = new FileReader();

  reader.onload = function (e) {
    const wb = XLSX.read(e.target.result, { type: "binary" });
    const sheet = wb.Sheets[wb.SheetNames[0]];
    const data = XLSX.utils.sheet_to_json(sheet);

    const output = data.map(row => {
      const tenantRaw = row["Tenant ID"] || "";
      const tenant = tenantRaw.replace(/^LN/, "");

      return {
        ...row,
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

    const ws = XLSX.utils.json_to_sheet(output);
    const wbOut = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wbOut, ws, "RFP FINAL");
    XLSX.writeFile(wbOut, "RFP_FINAL.xlsx");

    document.getElementById("status").innerText = "✅ RFP BERHASIL DIGENERATE";
  };

  reader.readAsBinaryString(fileInput.files[0]);
}
