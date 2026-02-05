console.log("✅ RFP Auto Generator Loaded");

// ==========================
// HELPER
// ==========================
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

// ==========================
// MAIN
// ==========================
function processExcel() {
  const fileInput = document.getElementById("file");

  if (!fileInput.files.length) {
    alert("Upload Excel Submit DRM dulu");
    return;
  }

  const reader = new FileReader();

  reader.onload = function (e) {
    const wb = XLSX.read(e.target.result, { type: "binary" });
    const sheet = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet);

    const output = rows.map(row => {
      const tenantRaw = row["Tenant ID"] || "";
      const tenant = tenantRaw.replace(/^LN/, "");

      const kawasan = row["Nama Perumahan/ Kawasan"] || "";
      const fdt = row["FDT Coding"] || "";
      const hpPlan = row["HP Plan"] || "";
      const hpDesign = row["HP Design (Breakdown Permit ID)"] || "";

      return {
        // ======================
        // RFP AUTO
        // ======================
        "Vendor RFP": "KESA",
        "Date Input": today(),
        "Project Type": "NRO B2S Longdrop",

        "City Town": getCityFromTenant(tenant),

        "Tenant ID": "LN" + tenant,
        "Permit ID": "LN" + tenant + "-001",
        "Cluster ID APD": tenant + "-001",

        "FDT Coding": fdt ? fdt + "EXT" : "",
        "Drawing Number LM": row["Drawing Number LM"] || "",

        "Nama Perumahan/ Kawasan": kawasan,
        "FDT Name/ Area Name": kawasan ? kawasan + " ADD HP" : "",

        "Latitude": row["Latitude"] || "",
        "Longitude": row["Longitude"] || "",

        "HP Plan": hpPlan,
        "HP Survey": hpPlan,
        "HP Design (Breakdown Permit ID)": hpDesign,
        "HP APD All": hpDesign,

        "HP Residential": row["HP Residential"] || "",
        "Bizz Pass": row["Bizz Pass"] || "",

        "Type FDT": "48C",
        "Kebutuhan Core BB": row["Kebutuhan Core BB"] || "",
        "Jumlah Splitter": row["Jumlah Splitter"] || "",
        "KM Strand LM (M)": row["KM Strand LM (M)"] || "",
        "Civil Work": row["Civil Work"] || "",

        "Link GDrive": ""
      };
    });

    const outSheet = XLSX.utils.json_to_sheet(output);
    const outWB = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(outWB, outSheet, "RFP FINAL");

    XLSX.writeFile(outWB, "RFP_FINAL.xlsx");

    document.getElementById("status").innerText =
      "✅ RFP berhasil digenerate (auto isi)";
  };

  reader.readAsBinaryString(fileInput.files[0]);
}
