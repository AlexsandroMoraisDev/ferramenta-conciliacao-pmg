const XLSX = require('xlsx');

function preview(file) {
  const workbook = XLSX.readFile(file);
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const json = XLSX.utils.sheet_to_json(sheet, { defval: '' });
  console.log(`\n--- File: ${file} ---`);
  if (json.length > 0) {
    // Find the first row with actual data to show keys
    let dataRow = json[0];
    for(let i=0; i<Math.min(json.length, 5); i++) {
        if(json[i]['Credor'] || json[i]['RAZÃO SOCIAL'] || json[i]['Status']) {
            dataRow = json[i];
            break;
        }
    }
    console.log("Keys:", Object.keys(dataRow));
    console.log("Sample Row:", dataRow);
  } else {
    console.log("No data");
  }
}

preview('../0.APROVADOS/TÍTULOS/ROMANEIO.xlsx');
preview('../0.APROVADOS/TÍTULOS/TÍTULOS SIENGE.xlsx');
preview('../0.APROVADOS/TÍTULOS/TÍTULOS ZEPP.xlsx');
