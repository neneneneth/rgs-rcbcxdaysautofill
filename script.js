document.getElementById("excel-file").addEventListener("change", function (e) {
  const expectedHeaders = [
  "AGENCY", "ADDRESS VISITED", "DATETIME", "CAMPAIGN", "PRODUCT/LEVEL",
  "ACCOUNT NO", "FULLNAME", "ENDORSEMENT DATE", "ADDRESS 1", "AREA 1",
  "ADDRESS 2", "AREA 2", "REFERENCE NO. 1:", "AGENT NAME", "ADMIN TEAM",
  "STATUS", "FIELD VISITATION REMARKS", "CATEGORY CODE", "Contact Number",
  "Gathered Contact Number Identification", "NOTES", "FIELD WORK CATEGORY",
  "INFORMANT/THIRD PARTY", "INFORMANT/THIRD PARTY NAME", "BRANCH ENDO",
  "PTP DATE", "PTP AMOUNT", "AGENT/TL", "SEGMENT"
];
  const file = e.target.files[0];
  if (!file) return;
  document.getElementById("file-name").textContent = `Uploaded file: ${file.name}`;
  const reader = new FileReader();
  reader.onload = function (event) {
  const data = new Uint8Array(event.target.result);
  const workbook = XLSX.read(data, { type: "array" });
  const firstSheetName = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[firstSheetName];

  const headers = extractHeadersFromSheet(worksheet);

  const headersMatch =
    headers.length === expectedHeaders.length &&
    expectedHeaders.every((h, i) => h === headers[i]);

  if (!headersMatch) {
    alert("❌ Invalid template. Please use the correct format.");
    e.target.value = "";
    return;
  }
  
  const jsonData = XLSX.utils.sheet_to_json(worksheet, { defval: "" });
  displayTable(jsonData);
  
  if (jsonData.length === 0) {
    alert("⚠️ The Excel file has headers only and no data rows.");
    e.target.value = "";
    return;
  }
  
  e.target.value = "";
};
  reader.readAsArrayBuffer(file);
});

document.getElementById('download-template').addEventListener('click', function () {
  const headers = [
    "AGENCY", "ADDRESS VISITED", "DATETIME", "CAMPAIGN", "PRODUCT/LEVEL", "ACCOUNT NO",
    "FULLNAME", "ENDORSEMENT DATE", "ADDRESS 1", "AREA 1", "ADDRESS 2", "AREA 2",
    "REFERENCE NO. 1:", "AGENT NAME", "ADMIN TEAM", "STATUS", "FIELD VISITATION REMARKS",
    "CATEGORY CODE", "Contact Number", "Gathered Contact Number Identification", "NOTES",
    "FIELD WORK CATEGORY", "INFORMANT/THIRD PARTY", "INFORMANT/THIRD PARTY NAME",
    "BRANCH ENDO", "PTP DATE", "PTP AMOUNT", "AGENT/TL", "SEGMENT"
  ];

  const ws = XLSX.utils.aoa_to_sheet([headers]);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Template");
  XLSX.writeFile(wb, "template.xlsx");
});

function displayTable(data) {
  const container = document.getElementById("table-container");
  if (!data.length) {
    container.innerHTML = "<p>No data found.</p>";
    return;
  }

  const table = document.createElement("table");
  const thead = document.createElement("thead");
  const tbody = document.createElement("tbody");

  const headers = Object.keys(data[0]);

  const importantHeaders = [
    "FULLNAME", "ACCOUNT NO", "AGENCY", "SEGMENT", "ADDRESS VISITED",
    "ADDRESS 1", "AREA 1", "STATUS", "Contact Number", "NOTES", "AGENT NAME"
  ];

  const headerRow = document.createElement("tr");

  const numberHeader = document.createElement("th");
  numberHeader.textContent = "#";
  headerRow.appendChild(numberHeader);

  const formHeader = document.createElement("th");
  formHeader.textContent = "Fill Form";
  headerRow.appendChild(formHeader);

  headers.forEach((header) => {
    const th = document.createElement("th");
    th.textContent = header;
    if (importantHeaders.includes(header)) {
      th.classList.add("highlight-header");
    }
    headerRow.appendChild(th);
  });

  thead.appendChild(headerRow);

  data.forEach((row, index) => {
    const tr = document.createElement("tr");

    const numCell = document.createElement("td");
    numCell.textContent = index + 1;
    tr.appendChild(numCell);

    const formCell = document.createElement("td");
    const button = document.createElement("button");
    button.textContent = "Open GForm";
    button.className = "table-button";
    button.addEventListener("click", () => {
      const url = buildGoogleFormURL(row);
      window.open(url, "_blank");
    });
    formCell.appendChild(button);
    tr.appendChild(formCell);

    headers.forEach((header) => {
      const td = document.createElement("td");
      td.textContent = row[header] || "";
      tr.appendChild(td);
    });

    tbody.appendChild(tr);
  });

  table.appendChild(thead);
  table.appendChild(tbody);
  container.innerHTML = "";
  container.appendChild(table);
}
function extractHeadersFromSheet(sheet) {
  const range = XLSX.utils.decode_range(sheet["!ref"]);
  const headers = [];
  for (let col = range.s.c; col <= range.e.c; col++) {
    const cellAddress = XLSX.utils.encode_cell({ r: 0, c: col });
    const cell = sheet[cellAddress];
    const header = cell ? cell.v.toString().trim() : "";
    headers.push(header);
  }
  return headers;
}
function buildGoogleFormURL(row) {
  const get = (key) => encodeURIComponent(row[key]?.toString().trim() || "");

  const q2 = (row["STATUS"] || "").toString().toLowerCase();
  const q2Text = encodeURIComponent(q2 === "ptp" ? "POSITIVE" : row["STATUS"] || "");
  const q2Check = encodeURIComponent(
    q2 === "ptp" ? "YES" : (q2 === "positive" || q2 === "negative" ? "NO" : "")
  );

  const url =
    "https://docs.google.com/forms/d/e/1FAIpQLSer2au7OtUQk6-Sgo_peX63JZtxpx6IUVVgrU1IbjkUW8mxiQ/viewform?usp=pp_url" +
    "&entry.872122274=" + get("FULLNAME") +
    "&entry.966019820=" + get("ACCOUNT NO") +
    "&entry.577032155=" + get("AGENCY") +
    "&entry.2088386466=" + get("SEGMENT") +
    "&entry.652378398=" + get("ADDRESS VISITED") +
    "&entry.1803052466=" + get("ADDRESS 1") +
    "&entry.62304814=" + get("AREA 1") +
    "&entry.1523428726=" + q2Text +
    "&entry.705381761=" + get("Contact Number") +
    "&entry.1003829361=" + get("NOTES") +
    "&entry.341897994=" + get("AGENT NAME") +
    "&entry.95880473=" + q2Check;

  return url;
}

