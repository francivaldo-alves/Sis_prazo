const fs = require('fs');
const XLSX = require('xlsx');

function normalizeSearch(str) {
    if (!str) return "";
    return String(str).normalize('NFD').replace(/[\u0300-\u036f]/g, "").toLowerCase().trim();
}

function getColValue(row, possibleNames) {
    for (let key in row) {
        const normKey = normalizeSearch(key);
        if (possibleNames.some(pn => normKey === normalizeSearch(pn))) {
            return row[key];
        }
    }
    return null;
}

const workbook = XLSX.readFile('cidades_atendidas_detalhado.xlsx');
const sheet = workbook.Sheets[workbook.SheetNames[0]];
const data = XLSX.utils.sheet_to_json(sheet, { defval: null });

const query = "cotia";
const matches = data.filter(row => {
    for (let key in row) {
        if (normalizeSearch(row[key])?.includes(query)) return true;
    }
    return false;
});
console.log("Matches count:", matches.length);

if (matches.length > 0) {
    console.log("First match:", matches[0]);

    const matchIsTecnico = matches[0] && getColValue(matches[0], ['TÉCNICO', 'Técnico', 'Tecnico']) && normalizeSearch(getColValue(matches[0], ['TÉCNICO', 'Técnico', 'Tecnico'])).includes(query);
    console.log("matchIsTecnico:", matchIsTecnico);

    const mapped = matches.map(m => `
            <div class="cidade-card">
                <div>
                    <span class="cidade-name" style="display:block; margin-bottom:0.25rem;">${getColValue(m, ['TÉCNICO', 'Técnico', 'Tecnico']) || 'N/A'}</span>
                    <span style="font-size: 0.75rem; color: #64748b;">De: ${getColValue(m, ['CIDADE BASE', 'Cidade Base']) || 'N/A'} (${getColValue(m, ['DISTÂNCIA (KM)', 'Distância']) || '0'}km)</span>
                </div>
                <span class="cidade-uf">${getColValue(m, ['UF ATENDIDA', 'UF Atendida', 'UF', 'Estado', 'Uf']) || 'N/A'}</span>
            </div>
    `);
    console.log("Mapped HTML sample:", mapped[0]);
} else {
    console.log("No match found for 'cotia'. Dump first 10 rows:");
    console.log(data.slice(0, 10));
}
