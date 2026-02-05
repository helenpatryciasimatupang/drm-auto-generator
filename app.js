console.log("✅ RFP Engine Loaded");

// =======================
// HELPER
// =======================
function today() {
  return new Date().toISOString().slice(0, 10);
}

function getCityFromTenant(tenant) {
  if (tenant.includes("SMG")) return "SEMARANG";
  if (tenant.includes("MDN")) return "MEDAN";
  if (tenant.includes("SLW") || tenant.includes("TGL")) return "TEGAL";
  if (tenant.includes("SBY")) return "SURABAYA";
  return "";
}

// =======================
// MAIN ENGINE
// =======================
function processExcel() {
  const fileInput = document.getElementById("file");

  if (!fileInput.files.length) {
    alert("Upload file Template Submit DRM dulu");
    return;
  }

  const reader = new FileReader();

  reader.onload = function (e) {
    const wb = XLSX.read(e.target.result, { type: "binary" });
    const sheetName = wb.SheetNames[0];
    const sheet = wb.Sheets[sheetName];

    const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });

    const output = rows.map(row => {
      const tenantRaw = row["Tenant"] || row["Tenant ID"] || "";
      const tenant = tenantRaw.replace(/^LN/, "");

      const kawasan = row["Nama Perumahan/ Kawasan"] || "";
      const fdt = row["FDT Coding"] || "";
      const hpPlan = row["HP Plan"] || "";
      const hpDesign = row["HP Design (Breakdown Permit ID)"] || "";

      return {
        ...row,

        // =====================
        // AUTO OVERRIDE
        // =====================
        "Vendor RFP": "KESA",
        "Date Input": today(),
        "Project Type": "NRO B2S Longdrop",

        "City Town": getCityFromTenant(tenant),

        "Tenant ID": "LN" + tenant,
        "Permit ID": "LN" + tenant + "-001",
        "Cluster ID APD": tenant + "-001",

        "FDT Coding": fdt ? fdt + "EXT" : "",
        "FDT Name/ Area Name": kawasan ? kawasan + " ADD HP" : "",

        "HP Survey": hpPlan,
        "HP APD All": hpDesign,

        "Type FDT": "48C",
        "Link Gdrive": ""
      };
    });

    const outSheet = XLSX.utils.json_to_sheet(output, { skipHeader: false });
    const outWB = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(outWB, outSheet, sheetName);

    XLSX.writeFile(outWB, "RFP_FINAL.xlsx");

    document.getElementById("status").innerText =
      "✅ RFP FINAL berhasil digenerate (siap submit)";
  };

  reader.readAsBinaryString(fileInput.files[0]);
}
