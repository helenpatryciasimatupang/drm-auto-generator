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

function addRow() {
  const tbody = document.querySelector("#dataTable tbody");
  const row = tbody.rows[0].cloneNode(true);
  row.querySelectorAll("input").forEach(i => i.value = "");
  tbody.appendChild(row);
}

function generate() {
  const rows = document.querySelectorAll("#dataTable tbody tr");

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

  rows.forEach(tr => {
    const c = tr.querySelectorAll("input");
    const tenant = c[0].value.trim();
    if (!tenant) return;

    data.push([
      "KESA",
      today(),
      "NRO B2S Longdrop",
      cityFromTenant(tenant),

      "LN"+tenant,
      "LN"+tenant+"-001",
      tenant+"-001",

      c[2].value ? c[2].value+"EXT" : "",
      c[3].value,

      c[1].value,
      c[1].value ? c[1].value+" ADD HP" : "",

      c[4].value,
      c[5].value,

      c[6].value,
      c[6].value,
      c[7].value,
      c[7].value,

      c[8].value,
      c[9].value,

      "48C",
      "-","-","-","-","-",""
    ]);
  });

  if (data.length === 1) {
    alert("Isi minimal 1 area");
    return;
  }

  const ws = XLSX.utils.aoa_to_sheet(data);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "RFP");
  XLSX.writeFile(wb, "RFP_FINAL.xlsx");

  document.getElementById("status").innerText =
    "✅ RFP berhasil dibuat ("+(data.length-1)+" baris)";
}
