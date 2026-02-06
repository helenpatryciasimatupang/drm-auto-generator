// ================= GLOBAL =================
let monitoringMap = {};   // key: FDTID HOTLIST
let masterFDTMap  = {};   // key: FDT_CODE

// ================= UTIL =================
function today(){ return new Date().toISOString().slice(0,10); }
function stripBOM(s){ return s && s.charCodeAt(0)===0xFEFF ? s.slice(1):s; }
function detectDelimiter(l){
  const c=(l.match(/,/g)||[]).length, s=(l.match(/;/g)||[]).length, t=(l.match(/\t/g)||[]).length;
  if(s>=c && s>=t) return ";"; if(t>=c && t>=s) return "\t"; return ",";
}
function splitCSV(l,d){
  const o=[]; let c="", q=false;
  for(let i=0;i<l.length;i++){
    const ch=l[i];
    if(ch=='"'){ q=!q; continue; }
    if(ch===d && !q){ o.push(c); c=""; continue; }
    c+=ch;
  }
  o.push(c); return o.map(x=>x.trim());
}
function norm(h){ return String(h||"").replace(/\u00A0/g," ").trim().replace(/\s+/g," ").toUpperCase(); }

// ================= CSV PARSER (Monitoring) =================
function parseCSV(text){
  text=stripBOM(text||"");
  const lines=text.split(/\r?\n/).filter(l=>l.trim());
  if(!lines.length) return [];
  const d=detectDelimiter(lines[0]);
  let hi=-1;
  for(let i=0;i<Math.min(lines.length,20);i++){
    const cols=splitCSV(lines[i],d).map(norm);
    if(cols.includes("FDTID HOTLIST")){ hi=i; break; }
  }
  if(hi<0) return [];
  const headers=splitCSV(lines[hi],d).map(norm);
  const rows=[];
  for(let i=hi+1;i<lines.length;i++){
    const cols=splitCSV(lines[i],d);
    if(cols.every(c=>!c)) continue;
    const o={}; headers.forEach((h,ix)=>o[h]=cols[ix]||""); rows.push(o);
  }
  return rows;
}

// ================= LOAD MONITORING =================
document.getElementById("monitoringFile").addEventListener("change", e=>{
  const f=e.target.files[0]; if(!f) return;
  const r=new FileReader();
  r.onload=()=>{
    monitoringMap={};
    parseCSV(r.result).forEach(row=>{
      const k=(row["FDTID HOTLIST"]||"").trim();
      if(k) monitoringMap[k]=row;
    });
    document.getElementById("status").innerText =
      `✅ Monitoring loaded (${Object.keys(monitoringMap).length} data)`;
  };
  r.readAsText(f);
});

// ================= LOAD MASTER POP UP (EXCEL) =================
document.getElementById("masterPopupFile").addEventListener("change", e=>{
  const f=e.target.files[0]; if(!f) return;
  const r=new FileReader();
  r.onload=evt=>{
    const wb=XLSX.read(evt.target.result,{type:"binary"});
    const ws=wb.Sheets[wb.SheetNames[0]];
    const rows=XLSX.utils.sheet_to_json(ws,{defval:""});
    masterFDTMap={};

    rows.forEach(row=>{
      const fdt = String(row["FDT_CODE"]||"").trim();
      const fat = String(row["FAT_CODE"]||"").trim();
      if(!fdt) return;

      if(!masterFDTMap[fdt]){
        masterFDTMap[fdt]={ lat:null, lng:null, hpDesign:0, hpHome:0, hpBiz:0 };
      }

      // Koordinat dari BARIS FDT (FAT_CODE kosong)
      if(!fat){
        masterFDTMap[fdt].lat = row["BUILDING_LATITUDE"] || masterFDTMap[fdt].lat;
        masterFDTMap[fdt].lng = row["BUILDING_LONGITUDE"] || masterFDTMap[fdt].lng;
      }

      // Hitungan HP dari semua FAT
      const hn = Number(row["HOUSE_NUMBER"]||0);
      masterFDTMap[fdt].hpDesign += hn;

      const typ = String(row["HOME/HOME-BIZ"]||"").toUpperCase();
      if(typ==="HOME") masterFDTMap[fdt].hpHome++;
      if(typ==="HOME-BIZ") masterFDTMap[fdt].hpBiz++;
    });

    document.getElementById("status").innerText +=
      ` | Master Pop Up loaded (${Object.keys(masterFDTMap).length} FDT)`;
  };
  r.readAsBinaryString(f);
});

// ================= UI =================
function addRow(){
  const tb=document.querySelector("#dataTable tbody");
  const r=tb.rows[0].cloneNode(true);
  r.querySelectorAll("input").forEach(i=>i.value="");
  tb.appendChild(r);
}

// ================= GENERATE =================
function generate(){
  const headers=[
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
  const data=[headers]; let no=1;

  document.querySelectorAll("#dataTable tbody tr").forEach(tr=>{
    const hot=tr.querySelector(".hotlist").value.trim();
    const fdt=tr.querySelector(".fdt").value.trim();
    const d5 =tr.querySelector(".draw5").value.trim();
    if(!hot || !fdt) return;

    const m=monitoringMap[hot];
    const d=masterFDTMap[fdt];
    if(!m){ alert(`FDTID HOTLIST tidak ditemukan: ${hot}`); return; }
    if(!d){ alert(`FDT Coding tidak ada di Master Pop Up: ${fdt}`); return; }

    data.push([
      no++,
      "KESA",
      today(),
      "NRO B2S Longdrop",
      m["CITY TOWN"]||"",
      m["TENANT ID PAPAH"]||"",
      m["PERMIT ID PAPAH"]||"",
      m["CLUSTER ID"]||"",
      fdt,
      d5?`KESA_2_PC_${d5}_0`:"",
      m["FDT NAME"]||"",
      m["FDT NAME"]?`${m["FDT NAME"]} ADD HP`:"",
      d.lat, d.lng,
      m["HP SURVEY"]||"",
      m["HP SURVEY"]||"",
      d.hpDesign,
      d.hpDesign,
      d.hpHome,
      d.hpBiz,
      "48C","-","-","-","AE",""
    ]);
  });

  if(data.length===1){ alert("Tidak ada data valid"); return; }
  const ws=XLSX.utils.aoa_to_sheet(data);
  const wb=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb,ws,"Submit DRM");
  XLSX.writeFile(wb,"RFP_SUBMIT_DRM.xlsx");
  document.getElementById("status").innerText = `✅ RFP berhasil (${data.length-1} baris)`;
}
