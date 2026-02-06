console.log("✅ app.js loaded (multi area)");

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

// tambah baris baru
function addRow() {
  const tbody = document.querySelector("#dataTable tbody");
  const firstRow = tbody.querySelector("tr");
  const newRow = firstRow.cloneNode(true);
  newRow.querySelectorAll("input").forEach(inp => inp.value = "");
  tbody.appendChild(newRow);

  document.getElementById("status").innerText = "➕ Area ditambahkan";
}

// generate excel
function generate() {
  try {
    const rows = Array.from(document.querySelectorAll("#dataTable tbody tr"));

    const headers = [
      "Vendor RFP","Date Input","Project Type","City Town",
      "Tenant ID","Permit ID","Cluster ID APD",
      "FDT Coding","Drawing Number LM",
      "Nama Perumahan/ Kawasan","FDT Name/ Area Name",
      "Latitude","Longitude",
      "HP Plan","HP Survey",
      "HP Design (Breakdown Permit ID)","HP APD All",
      "HP Residential","Bizz Pass",
      "Type FDT",
      "Kebutuhan Core BB","Jumlah Splitter","KM Strand LM (M)",
      "CIvil Work","Civil Work",
      "Link GDrive"
    ];

    const data = [headers];

    rows.forEach((tr, idx) => {
      const tenant = tr.querySelector(".tenant")?.value?.trim() || "";
      if (!tenant) return; // skip baris kosong

      const kawasan = tr.querySelector(".kawasan")?.value?.trim() || "";
      const fdt = tr.querySelector(".fdt")?.value?.trim() || "";
      const drawing = tr.querySelector(".drawing")?.value?.trim() || "";
      const lat = tr.querySelector(".lat")?.value?.trim() || "";
      const lng = tr.querySelector(".lng")?.value?.trim() || "";
      const hpPlan = tr.querySelector(".hpPlan")?.value?.trim() || "";
      const hpDesign = tr.querySelector(".hpDesign")?.value?.trim() || "";
      const hpRes = tr.querySelector(".hpRes")?.value?.trim() || "";
      const bizz = tr.querySelector(".bizz")?.value?.trim() || "";

      data.push([
        "KESA",
        today(),
        "NRO B2S Longdrop",
        cityFromTenant(tenant),

        "LN" + tenant,
        "LN" + tenant + "-001",
        tenant + "-001",

        fdt ? (fdt + "EXT") : "",
        drawing,

        kawasan,
        kawasan ? (kawasan + " ADD HP") : "",

        lat,
        lng,

        hpPlan,
        hpPlan,
        hpDesign,
        hpDesign,

        hpRes,
        bizz,

        "48C",
        "-", "-", "-", "-", "-", "" // defaults + Link GDrive kosong
      ]);
    });

    if (data.length === 1) {
      alert("Tenant ID wajib diisi minimal 1 baris");
      return;
    }

    const ws = XLSX.utils.aoa_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "RFP");
    XLSX.writeFile(wb, "RFP_FINAL.xlsx");

    document.getElementById("status").innerText =
      "✅ Berhasil generate (" + (data.length - 1) + " baris)";

  } catch (e) {
    console.error(e);
    alert("❌ Error generate. Buka Console (F12).");
  }
}
