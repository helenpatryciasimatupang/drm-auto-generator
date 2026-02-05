function today() {
  return new Date().toISOString().slice(0,10);
}

function cityFromTenant(t) {
  if (t.includes("SMG")) return "SEMARANG";
  if (t.includes("MDN")) return "MEDAN";
  if (t.includes("SLW") || t.includes("TGL")) return "TEGAL";
  if (t.includes("SBY")) return "SURABAYA";
  return "";
}

function v(id) {
  return document.getElementById(id).value || "";
}

function generate() {
  const tenant = v("tenant");
  if (!tenant) {
    alert("Tenant ID wajib diisi");
    return;
  }

  const data = [{
    "Vendor RFP": "KESA",
    "Date Input": today(),
    "Project Type": "NRO B2S Longdrop",
    "City Town": cityFromTenant(tenant),

    "Tenant ID": "LN" + tenant,
    "Permit ID": "LN" + tenant + "-001",
    "Cluster ID APD": tenant + "-001",

    "FDT Coding": v("fdt") + "EXT",
    "Drawing Number LM": v("drawing"),

    "Nama Perumahan/ Kawasan": v("kawasan"),
    "FDT Name/ Area Name": v("kawasan") + " ADD HP",

    "Latitude": v("lat"),
    "Longitude": v("lng"),

    "HP Plan": v("hpPlan"),
    "HP Survey": v("hpPlan"),
    "HP Design (Breakdown Permit ID)": v("hpDesign"),
    "HP APD All": v("hpDesign"),

    "HP Residential": v("hpRes"),
    "Bizz Pass": v("bizz"),

    "Type FDT": "48C",
    "Kebutuhan Core BB": v("core"),
    "Jumlah Splitter": v("splitter"),
    "KM Strand LM (M)": v("km"),
    "Civil Work": v("civil"),

    "Link GDrive": ""
  }];

  const ws = XLSX.utils.json_to_sheet(data);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "RFP");

  XLSX.writeFile(wb, "RFP_FINAL.xlsx");
  document.getElementById("status").innerText = "✅ RFP berhasil dibuat";
}
