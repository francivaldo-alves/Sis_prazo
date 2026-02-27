let globalData = [];
let headers = [];

const searchInput = document.getElementById('search-input');
const btnSearch = document.getElementById('btn-search');
const statusIndicator = document.getElementById('status-indicator');
const statusText = statusIndicator.querySelector('.text');
const heroSection = document.getElementById('hero-section');
const dashboard = document.getElementById('dashboard');
const btnClear = document.getElementById('btn-clear');

const estabDetailsCard = document.getElementById('estab-details');
const bestMatchCard = document.getElementById('best-match-card');
const resultsTable = document.querySelector('#results-table tbody');

const uploadFallback = document.getElementById('upload-fallback');
const fileInput = document.getElementById('file-input');

// Define default excel file
const DEFAULT_FILENAME = 'Base_cadastros_Unificada_Final.xlsx';

function normalizeSearch(str) {
    if (!str) return "";
    return String(str).toLowerCase().trim().replace(/[-.,]/g, '');
}

window.addEventListener('DOMContentLoaded', async () => {
    try {
        const response = await fetch(DEFAULT_FILENAME);
        if (!response.ok) throw new Error('Network response was not ok');
        const arrayBuffer = await response.arrayBuffer();
        processExcelData(arrayBuffer);
    } catch (error) {
        console.warn("Could not load default file, showing upload option.", error);
        statusIndicator.className = 'status error';
        statusText.innerText = "Base local não encontrada";
        uploadFallback.classList.remove('hidden');
    }
});

fileInput.addEventListener('change', (e) => {
    if (e.target.files.length) {
        statusIndicator.className = 'status loading';
        statusText.innerText = "Processando arquivo...";
        uploadFallback.classList.add('hidden');

        const file = e.target.files[0];
        const reader = new FileReader();
        reader.onload = function (evt) {
            try {
                processExcelData(evt.target.result);
            } catch (err) {
                statusIndicator.className = 'status error';
                statusText.innerText = "Erro ao ler arquivo manual";
                uploadFallback.classList.remove('hidden');
            }
        };
        reader.readAsArrayBuffer(file);
    }
});

function processExcelData(dataBuffer) {
    try {
        const data = new Uint8Array(dataBuffer);
        const workbook = XLSX.read(data, { type: 'array' });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];

        const json = XLSX.utils.sheet_to_json(worksheet, { defval: null });

        if (json.length === 0) throw new Error("A planilha está vazia.");

        globalData = json;
        headers = Object.keys(json[0]);

        statusIndicator.className = 'status ready';
        statusText.innerText = `${globalData.length} registros prontos`;

        searchInput.disabled = false;
        searchInput.focus();
        searchInput.placeholder = "Digite o CEP ou Terminal (Ex: 06455-000)";

    } catch (error) {
        console.error("Error parsing Excel:", error);
        statusIndicator.className = 'status error';
        statusText.innerText = "Erro ao processar dados";
        uploadFallback.classList.remove('hidden');
    }
}

// Search Logic
searchInput.addEventListener('keypress', (e) => {
    if (e.key === 'Enter') performSearch();
});

btnSearch.addEventListener('click', performSearch);

btnClear.addEventListener('click', () => {
    searchInput.value = '';
    dashboard.classList.remove('visible');

    setTimeout(() => {
        heroSection.classList.remove('minimized');
    }, 100);

    setTimeout(() => { searchInput.focus(); }, 300);
});

function performSearch() {
    const rawQuery = searchInput.value.trim();
    if (!rawQuery) return;

    const query = normalizeSearch(rawQuery);

    let match = null;

    // First Priority: Try to match CEP exactly
    match = globalData.find(row => {
        const cepVal = normalizeSearch(row['CEP'] || row['cep']);
        if (cepVal && cepVal === query) return true;
        return false;
    });

    // Substring fallback if exact match fails
    if (!match) {
        match = globalData.find(row => {
            for (let key in row) {
                const cellVal = normalizeSearch(row[key]);
                if (cellVal && cellVal.includes(query)) {
                    return true;
                }
            }
            return false;
        });
    }

    if (!match) {
        searchInput.style.animation = 'shake 0.5s ease';
        setTimeout(() => searchInput.style.animation = '', 500);
        return;
    }

    renderEstablishmentInfo(match);
    analyzeDeliveryTimes(match);
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

function renderEstablishmentInfo(row) {
    // Try to find known columns using heuristic variations
    const nome = getColValue(row, ['Nome', 'Razao Social', 'Estabelecimento', 'Terminal']) || 'N/A';
    const endereco = getColValue(row, ['Endereço', 'Endereco', 'Logradouro', 'Rua']) || 'N/A';
    const cidade = getColValue(row, ['Cidade', 'Municipio']) || 'N/A';
    const uf = getColValue(row, ['UF', 'Estado']) || 'N/A';
    const grupo = getColValue(row, ['Grupo', 'Categoria', 'Setor', 'Rede']) || 'N/A';
    const cep = getColValue(row, ['CEP']) || 'N/A';

    estabDetailsCard.innerHTML = `
        <div class="estab-header">
            <h3>${nome}</h3>
            ${grupo !== 'N/A' ? `<span class="badge">${grupo}</span>` : ''}
        </div>
        <div class="estab-info-grid">
            <div class="estab-field">
                <span class="label">Endereço</span>
                <span class="val">${endereco} - ${cep}</span>
            </div>
            <div class="estab-field">
                <span class="label">Cidade</span>
                <span class="val">${cidade}</span>
            </div>
            <div class="estab-field">
                <span class="label">UF</span>
                <span class="val">${uf}</span>
            </div>
        </div>
    `;
}

function analyzeDeliveryTimes(row) {
    // We ignore typical descriptive columns AND explicitly ignore the columns user mentioned
    const ignoreKeywords = [
        'cep', 'terminal', 'cidade', 'estado', 'uf', 'região', 'regiao', 'id', 'codigo', 'origem',
        'endereco', 'endereço', 'estab novo', 'nome', 'grupo', 'razao social', 'logradouro'
    ];

    const deliveryOptions = [];

    for (let i = 0; i < headers.length; i++) {
        const colName = headers[i];
        const val = row[colName];

        const isIgnored = ignoreKeywords.some(kw => normalizeSearch(colName).includes(normalizeSearch(kw)));

        // Exact column skip based on user feedback (e.g., column might contain "Estab Novo")
        if (isIgnored) continue;

        if (val !== null && val !== undefined) {
            const numVal = parseInt(String(val).replace(/[^0-9]/g, ''), 10);

            // Ensure the value behaves like a delivery time (usually reasonable days, not huge IDs)
            // also ensure the string contains numbers and is not empty
            if (!isNaN(numVal) && String(val).trim() !== "" && numVal < 1000) {
                deliveryOptions.push({
                    provider: colName,
                    days: numVal
                });
            }
        }
    }

    deliveryOptions.sort((a, b) => a.days - b.days);

    renderResults(deliveryOptions);
}

function renderResults(options) {
    heroSection.classList.add('minimized');

    setTimeout(() => {
        dashboard.classList.add('visible');
    }, 100);

    resultsTable.innerHTML = '';

    if (options.length === 0) {
        bestMatchCard.innerHTML = `<div class="winner-name">Nenhum prazo encontrado</div>`;
        return;
    }

    const bestOption = options[0];

    // Recommend Card
    bestMatchCard.innerHTML = `
        <div class="trophy-icon">
            <svg xmlns="http://www.w3.org/2000/svg" width="32" height="32" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M13 10V3L4 14h7v7l9-11h-7z" />
            </svg>
        </div>
        <h3>Menor Prazo Estimado</h3>
        <div class="winner-name">${bestOption.provider}</div>
        <div>
            <span class="winner-time">${bestOption.days} dias</span>
        </div>
    `;

    // Table
    options.forEach((opt, index) => {
        const tr = document.createElement('tr');

        let rowClass = "";
        if (index === 0) rowClass = "row-success";
        else if (opt.days > bestOption.days * 1.5) rowClass = "row-warning";

        if (rowClass) tr.classList.add(rowClass);

        tr.innerHTML = `
            <td class="provider-name">${opt.provider}</td>
            <td class="time-badge">${opt.days} dias</td>
        `;
        resultsTable.appendChild(tr);
    });
}
