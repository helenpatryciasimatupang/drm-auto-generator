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
  return document.getElementById(id)?.value?.trim() || "";
}

// Alias untuk nama header yang sering beda-beda
function normHeader(h) {
  return String(h || "")
    .toLowerCase()
    .replace(/\s+/g, " ")
    .replace(/[^\w\s()/.-]/g, "")
    .trim();
}

function generate() {
  try {
    const tenant = v("tenant");
    if (!tenant) {
      alert("Tenant ID wajib diisi");
      return;
    }

    const kawasan = v("kawasan");
    const fdt = v("fdt");
    const drawing = v("drawing");
    const lat = v("lat");
    const lng = v("lng");
    const hpPlan = v("hpPlan");
    const hpDesign = v("hpDesign");
    const hpRes = v("hpRes");
    const bizz = v("bizz");

    // === KUNCI URUTAN + KUNCI NAMA KOLUMN (template-like) ===
    // Aku buat beberapa variasi header yang sering muncul.
    // Pilih yang mana? Kita pakai yang paling umum + kita “duplikasi” untuk yang beda nama.
    // Jadi walau template kamu pakai "CIvil Work", tetap keisi.

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

      // === Yang kamu minta selalu "-" ===
      "Kebutuhan Core BB",
      "Jumlah Splitter",
      "KM Strand LM (M)",
      "CIvil Work",      // sesuai tulisan kamu (huruf I besar)
      "Civil Work",      // jaga-jaga kalau template pakai normal

      // === Link gdrive harus kosong ===
      "LINK GDRIVE",
      "Link GDrive",
      "Link Gdrive"
    ];

    const row = [
      "KESA",
      today(),
      "NRO B2S Longdrop",
      cityFromTenant(tenant),

      "LN" + tenant,
      "LN" + tenant + "-001",
      tenant + "-001",

      fdt ? fdt + "EXT" : "",
      drawing,

      kawasan,
      kawasan ? kawasan + " ADD HP" : "",

      lat,
      lng,

      hpPlan,
      hpPlan,
      hpDesign,
      hpDesign,

      hpRes,
      bizz,

      "48C",

      "-", // Kebutuhan Core BB
      "-", // Jumlah Splitter
      "-", // KM Strand LM (M)
      "-", // CIvil Work
      "-", // Civil Work

      "",  // LINK GDRIVE
      "",  // Link GDrive
      ""   // Link Gdrive
    ];

    const ws = XLSX.utils.aoa_to_sheet([headers, row]);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "RFP");

    XLSX.writeFile(wb, "RFP_FINAL.xlsx");

    document.getElementById("status").innerText =
      "✅ RFP berhasil digenerate & file terdownload";

  } catch (err) {
    console.error(err);
    alert("❌ Gagal generate RFP. Buka Console (F12) untuk lihat error.");
  }
}
