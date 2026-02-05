console.log("✅ app.js loaded");

function today() {
  return new Date().toISOString().slice(0, 10);
}

function cityFromTenant(t) {
  if (t.includes("SMG")) return "SEMARANG";
  if (t.includes("MDN")) return "MEDAN";
  if (t.includes("SLW") || t.includes("TGL")) return "TEGAL";
  if (t.includes("SBY")) return "SURABAYA";
  return "";
}

function v(id) {
  return document.getElementById(id)?.value || "";
}

function generate() {
  try {
    console.log("🟢 Generate diklik");

    const tenant = v("tenant").trim();
    if (!tenant) {
      alert("Tenant ID wajib diisi");
      return;
    }

    // HEADER DIKUNCI SESUAI SUBMIT DRM
    const headers = [
      "Vendor RFP",
      "Date Input",
      "Project Type",
      "City Town",
      "Tenant ID",
      "Permit ID",
      "Cluster ID APD",
      "FDT Coding",
      "Drawing Number LM",
      "Nama Perumahan/ Kawasan",
      "FDT Name/ Area Name",
      "Latitude",
      "Longitude",
      "HP Plan",
      "HP Survey",
      "HP Design (Breakdown Permit ID)",
      "HP APD All",
      "HP Residential",
      "Bizz Pass",
      "Type FDT",
      "Kebutuhan Core BB",
      "Jumlah Splitter",
      "KM Strand LM (M)",
      "Civil Work",
      "Link GDrive"
    ];

    const row = [
      "KESA",
      today(),
      "NRO B2S Longdrop",
      cityFromTenant(tenant),

      "LN" + tenant,
      "LN" + tenant + "-001",
      tenant + "-001",

      v("fdt") ? v("fdt") + "EXT" : "",
      v("drawing"),

      v("kawasan"),
      v("kawasan") ? v("kawasan") + " ADD HP" : "",

      v("lat"),
      v("lng"),

      v("hpPlan"),
      v("hpPlan"),
      v("hpDesign"),
      v("hpDesign"),

      v("hpRes"),
      v("bizz"),

      "48C",

      "-",   // Kebutuhan Core BB
      "-",   // Jumlah Splitter
      "-",   // KM Strand LM (M)
      "-",   // Civil Work

      ""     // Link GDrive
    ];

    const ws = XLSX.utils.aoa_to_sheet([headers, row]);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "RFP");

    XLSX.writeFile(wb, "RFP_FINAL.xlsx");

    document.getElementById("status").innerText =
      "✅ RFP berhasil digenerate & file terdownload";

    alert("RFP_FINAL.xlsx berhasil dibuat");

  } catch (err) {
    console.error(err);
    alert("❌ Gagal generate RFP, cek Console (F12)");
  }
}
