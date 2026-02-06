let monitoringMap = {};
let masterFDTMap = {};

// ================= UTIL =================
function today(){ return new Date().toISOString().slice(0,10); }
function stripBOM(s){ return s && s.charCodeAt(0) === 0xFEFF ? s.slice(1) : s; }

function detectDelimiter(line){
  const c=(line.match(/,/g)||[]).length;
  const s=(line.match(/;/g)||[]).length;
  const t=(line.match(/\t/g)||[]).length;
  if(s>=c && s>=t) return ";";
  if(t>=c && t>=s) return "\t";
  return ",";
}

function splitCSV(line,d){
  const out=[],cur=[""]; let q=false;
  for(let i=0;i<line.length;i++){
    const ch=line[i];
    if(ch === '"'){ q=!q; continue; }
    if(ch===d && !q){ out.push(cur.join("")); cur.length=0; continue; }
    cur.push(ch);
  }
  out.push(cur.join(""));
  return out.map(x=>x.trim());
}

function norm(h){
  return String(h||"").replace(/\u00A0/g," ").trim().replace(/\s+/g," ").toUpperCase();
}

// ================= CSV PARSER =================
function parseCSV(text){
  text = stripBOM(text||"");
  const lines = text.split(/\r?\n/).filter(l=>l.trim());
  if(!lines.length) return [];
  const d = detectDelimiter(lines[0]);

  let headerIdx=-1;
  for(let i=0;i<20;i++){
    const cols = splitCSV(lines[i],d).map(norm);
    if(cols.includes("FDTID HOTLIST") || cols.includes("NODE_CODE")){
      headerIdx=i; break;
    }
  }
  if(headerIdx<0) return [];

  const headers = splitCSV(lines[headerIdx],d).map(norm);
  const rows=[];
  for(let i=headerIdx+1;i<lines.length;i++){
    const cols = splitCSV(lines[i],d);
    if(cols.every(c=>!c)) continue;
    const o={};
    headers.forEach((h,idx)=>o[h]=cols[idx]||"");
    rows.push(o);
  }
  return rows;
}

// ================= NORMALISASI FAT → FDT =================
function fatToFdt(code){
  return code.replace(/S\d+A\d+/i,"");
}

// ================= LOAD FILES =================
document.getElementById("monitoringFile").addEventListener("change",e=>{
  const r=new FileReader();
  r.onload=()=>{
    monitoringMap={};
    parseCSV(r.result).forEach(row=>{
      const k=row["FDTID HOTLIST"];
      if(k) monitoringMap[k]=row;
    });
    document.getElementById("status").innerText =
      `✅ Monitoring loaded (${Object.keys(monitoringMap).length})`;
  };
  r.readAsText(e.target.files[0]);
});

document.getElementById("masterHPFile").addEventListener("change",e=>{
  const r=new FileReader();
  r.onload=()=>{
    masterFDTMap={};

    parseCSV(r.result).forEach(row=>{
      const node=row["NODE_CODE"];
      if(!node) return;
      const fdt=fatToFdt(node);

      if(!masterFDTMap[fdt]){
        masterFDTMap[fdt]={
          lat:null,lng:null,
          hpDesign:0,hpHome:0,hpBiz:0
        };
      }

      // FDT coordinate (node tanpa SxxAxx)
      if(!node.match(/S\d+A\d+/i)){
        masterFDTMap[fdt].lat=row["LATITUDE"];
        masterFDTMap[fdt].lng=row["LONGITUDE"];
      }

      const hn = Number(row["HOUSE_NUMBER"]||0);
      masterFDTMap[fdt].hpDesign += hn;

      if(row["HP_TYPE"]==="HOME") masterFDTMap[fdt].hpHome++;
      if(row["HP_TYPE"]==="HOME-BIZ") masterFDTMap[fdt].hpBiz++;
    });

    document.getElementById("status").innerText +=
      ` | Master HP loaded (${Object.keys(masterFDTMap).length} FDT)`;
  };
  r.readAsText(e.target.files[0]);
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

  const data=[headers];
  let no=1;

  document.querySelectorAll("#dataTable tbody tr").forEach(tr=>{
    const hot=tr.querySelector(".hotlist").value.trim();
    if(!hot) return;

    const m=monitoringMap[hot];
    const fdtData=masterFDTMap[hot];
    if(!m||!fdtData) return;

    const fdt=tr.querySelector(".fdt").value.trim();
    const d5=tr.querySelector(".draw5").value.trim();

    data.push([
      no++,
      "KESA",
      today(),
      "NRO B2S Longdrop",
      m["CITY TOWN"],
      m["TENANT ID PAPAH"],
      m["PERMIT ID PAPAH"],
      m["CLUSTER ID"],
      fdt,
      d5 ? `KESA_2_PC_${d5}_0` : "",
      m["FDT NAME"],
      `${m["FDT NAME"]} ADD HP`,
      fdtData.lat,
      fdtData.lng,
      m["HP SURVEY"],
      m["HP SURVEY"],
      fdtData.hpDesign,
      fdtData.hpDesign,
      fdtData.hpHome,
      fdtData.hpBiz,
      "48C",
      "-","-","-",
      "AE",
      ""
    ]);
  });

  const ws=XLSX.utils.aoa_to_sheet(data);
  const wb=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb,ws,"Submit DRM");
  XLSX.writeFile(wb,"RFP_SUBMIT_DRM.xlsx");

  document.getElementById("status").innerText =
    `✅ Berhasil generate (${data.length-1} baris)`;
}
