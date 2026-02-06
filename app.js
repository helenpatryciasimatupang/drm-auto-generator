let monitoringMap = {};

function today(){ return new Date().toISOString().slice(0,10); }

function parseCSV(text){
  const lines = text.split(/\r?\n/).filter(l=>l.trim()!=="");
  const headers = lines.shift().split(",").map(h=>h.trim());
  return lines.map(l=>{
    const cols = l.split(",");
    const o={}; headers.forEach((h,i)=>o[h]= (cols[i]||"").trim());
    return o;
  });
}

// Load monitoring
document.getElementById("monitoringFile").addEventListener("change", (e)=>{
  const f = e.target.files[0];
  if(!f) return;
  const r = new FileReader();
  r.onload = ()=>{
    const rows = parseCSV(r.result);
    monitoringMap = {};
    rows.forEach(row=>{
      const key = row["FDTID HOTLIST"];
      if(key) monitoringMap[key] = row;
    });
    document.getElementById("status").innerText =
      `✅ Monitoring loaded (${Object.keys(monitoringMap).length} data)`;
  };
  r.readAsText(f);
});

function addRow(){
  const tb = document.querySelector("#dataTable tbody");
  const r = tb.rows[0].cloneNode(true);
  r.querySelectorAll("input").forEach(i=>i.value="");
  tb.appendChild(r);
}

function generate(){
  if(!Object.keys(monitoringMap).length){
    alert("Upload file Monitoring dulu");
    return;
  }

  const headers = [
    "No","Vendor RFP","Date Input","Project Type","City Town",
    "Tenant ID","Permit ID","Cluster ID APD",
    "FDT Coding","Drawing Number LM",
    "Nama Perumahan/ Kawasan","FDT Name/ Area Name",
    "Latitude","Longitude",
    "HP Plan","HP Survey",
    "HP Design (Breakdown Permit ID)","HP APD All",
    "HP Residential","Bizz Pass",
    "Type FDT","Kebutuhan Core BB","Jumlah Splitter","KM Strand LM (M)",
    "CIvil Work","Link Gdrive"
  ];

  const data=[headers];
  let no=1;

  document.querySelectorAll("#dataTable tbody tr").forEach(tr=>{
    const hot = tr.querySelector(".hotlist").value.trim();
    if(!hot) return;

    const m = monitoringMap[hot];
    if(!m){ alert(`FDTID HOTLIST tidak ditemukan: ${hot}`); return; }

    const fdt = tr.querySelector(".fdt").value.trim();
    const d5  = tr.querySelector(".draw5").value.trim();
    const lat = tr.querySelector(".lat").value.trim();
    const lng = tr.querySelector(".lng").value.trim();
    const hpD = tr.querySelector(".hpDesign").value.trim();
    const hpR = tr.querySelector(".hpRes").value.trim();
    const biz = tr.querySelector(".bizz").value.trim();

    data.push([
      no++,
      "KESA",
      today(),
      "NRO B2S Longdrop",
      m["City Town"],
      m["Tenant ID PAPAH"],
      m["Permit ID PAPAH"],
      m["Cluster ID"],
      fdt,
      d5 ? `KESA_2_PC_${d5}_0` : "",
      m["FDT Name"],
      m["FDT Name"] ? `${m["FDT Name"]} ADD HP` : "",
      lat, lng,
      m["HP Survey"],
      m["HP Survey"],
      hpD,
      hpD,
      hpR,
      biz,
      "48C",
      "-", "-", "-",
      "AE",
      ""
    ]);
  });

  if(data.length===1){
    alert("Isi minimal 1 FDTID HOTLIST");
    return;
  }

  const ws = XLSX.utils.aoa_to_sheet(data);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Submit DRM");
  XLSX.writeFile(wb, "RFP_SUBMIT_DRM.xlsx");

  document.getElementById("status").innerText =
    `✅ Berhasil generate (${data.length-1} baris)`;
}
