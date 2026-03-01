const XLSX = require('xlsx');

try {
    const workbook = XLSX.readFile(process.argv[2]);
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const json = XLSX.utils.sheet_to_json(sheet, { defval: null });

    if (json.length > 0) {
        console.log("Columns:");
        const headers = Object.keys(json[0]);
        headers.forEach(h => console.log(` - ${h}`));
        console.log("\nFirst row:");
        console.log(JSON.stringify(json[0], null, 2));
    } else {
        console.log("Empty sheet");
    }
} catch (e) {
    console.error(e.message);
}
