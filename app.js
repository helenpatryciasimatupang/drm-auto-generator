function today() {
  const d = new Date();
  return d.toISOString().split("T")[0];
}

function addRow() {
  const tbody = document.getElementById("tbody");
  const tr = document.createElement("tr");
  tr.innerHTML = `
    <td>${tbody.children.length + 1}</td>
    <td><input></td>
    <td><input></td>
    <td><input></td>
    <td><input></td>
    <td><input class="hot"></td>
    <td><input class="draw"></td>
    <td><input></td>
    <td><input></td>
    <td><input></td>
    <td><input></td>
    <td><input></td>
    <td><input></td>
  `;
  tbody.appendChild(tr);
}

function generateExcel() {
  const header = [
    "No",
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
    "CIvil Work",
    "Link Gdrive"
  ];

  const rows = [];
  rows.push(header);

  const trs = document.querySelectorAll("#tbody tr");
  let no = 1;

  trs.forEach(tr => {
    const td = tr.querySelectorAll("td input");

    const city = td[0].value;
    const tenant = td[1].value;
    const permit = td[2].value;
    const cluster = td[3].value;
    const hot = td[4].value;
    const draw = td[5].value;
    const fdtName = td[6].value;
    const lat = td[7].value;
    const lng = td[8].value;
    const hpSurvey = td[9].value;
    const hpRes = td[10].value;
    const biz = td[11].value;

    rows.push([
      no++,
      "KESA",
      today(),
      "NRO B2S Longdrop",
      city,
      tenant,
      permit,
      cluster,
      hot ? hot + "EXT" : "",
      draw ? `KESA_2_PC_${draw}_0` : "",
      fdtName,
      fdtName ? fdtName + " ADD HP" : "",
      lat,
      lng,
      hpSurvey,
      hpSurvey,
      hpSurvey,
      hpSurvey,
      hpRes,
      biz,
      "48C",
      "-",
      "-",
      "-",
      "AE",
      ""
    ]);
  });

  const ws = XLSX.utils.aoa_to_sheet(rows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "RFP");

  XLSX.writeFile(wb, "RFP_Output.xlsx");
}
