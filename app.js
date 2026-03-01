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
const highlightCards = document.getElementById('highlight-cards');
const delayWarningContainer = document.getElementById('delay-warning-container');
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
    const nome = getColValue(row, ['Nome', 'Razao Social', 'Estabelecimento']) || 'N/A';
    const terminal = getColValue(row, ['Terminal']) || '';
    const endereco = getColValue(row, ['Endereço', 'Endereco', 'Logradouro', 'Rua']) || 'N/A';
    const bairro = getColValue(row, ['Bairro']) || '';
    const cidade = getColValue(row, ['Cidade', 'Municipio']) || 'N/A';
    const uf = getColValue(row, ['UF', 'Estado']) || 'N/A';
    const grupo = getColValue(row, ['Grupo', 'Categoria', 'Setor', 'Rede']) || 'N/A';
    const cep = getColValue(row, ['CEP']) || 'N/A';

    const tipoOperacao = getColValue(row, ['Tipo Operacao', 'Tipo de Operacao', 'Tipo Operação', 'Operacao']) || '';
    const tecnico = getColValue(row, ['Técnico Mais Próximo', 'Tecnico Mais Proximo', 'Tecnico MAIS PROXIMO', 'Técnico']) || '';
    const melhorEnvio = getColValue(row, ['Melhor Forma de Envio', 'Melhor Forma', 'Melhor Envio', 'Forma de Envio']) || '';

    const displayName = terminal ? `${terminal} - ${nome}` : nome;

    estabDetailsCard.innerHTML = `
        <div class="estab-header">
            <h3>${displayName}</h3>
            <div class="badges-container">
                ${grupo !== 'N/A' ? `<span class="badge">${grupo}</span>` : ''}
                ${tipoOperacao ? `<span class="badge badge-op">${tipoOperacao}</span>` : ''}
            </div>
        </div>
        <div class="estab-info-grid">
            <div class="estab-field">
                <span class="label">Endereço</span>
                <span class="val">
                    ${endereco}${bairro ? ' - ' + bairro : ''} - ${cep}
                    <button class="btn-copy" onclick="copyToClipboard('${endereco}${bairro ? ' - ' + bairro : ''} - ${cep}', this)" title="Copiar Endereço">
                        <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                            <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M8 16H6a2 2 0 01-2-2V6a2 2 0 012-2h8a2 2 0 012 2v2m-6 12h8a2 2 0 002-2v-8a2 2 0 00-2-2h-8a2 2 0 00-2 2v8a2 2 0 002 2z" />
                        </svg>
                    </button>
                </span>
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

    // Render Highlights
    highlightCards.innerHTML = '';

    if (melhorEnvio) {
        highlightCards.innerHTML += `
            <div class="highlight-card method-card">
                <div class="highlight-icon method-icon">
                    <svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                        <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M13 10V3L4 14h7v7l9-11h-7z" />
                    </svg>
                </div>
                <div class="highlight-content">
                    <span class="label">Melhor Forma de Envio</span>
                    <span class="val">${melhorEnvio}</span>
                </div>
            </div>
        `;
    }

    if (tecnico && tecnico.toLowerCase() !== 'n/a') {
        highlightCards.innerHTML += `
            <div class="highlight-card tech-card">
                <div class="highlight-icon tech-icon">
                    <svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                        <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M16 7a4 4 0 11-8 0 4 4 0 018 0zM12 14a7 7 0 00-7 7h14a7 7 0 00-7-7z" />
                    </svg>
                </div>
                <div class="highlight-content">
                    <span class="label">Técnico Mais Próximo</span>
                    <span class="val">${tecnico}</span>
                </div>
            </div>
        `;
    }

    if (melhorEnvio || tecnico) {
        highlightCards.style.display = 'grid';
    } else {
        highlightCards.style.display = 'none';
    }
}

function analyzeDeliveryTimes(row) {
    // We ignore typical descriptive columns AND explicitly ignore the columns user mentioned
    const ignoreKeywords = [
        'cep', 'terminal', 'cidade', 'estado', 'uf', 'região', 'regiao', 'id', 'codigo', 'origem',
        'endereco', 'endereço', 'estab novo', 'nome', 'grupo', 'razao social', 'logradouro',
        'bairro', 'tipo operacao', 'técnico', 'tecnico', 'melhor forma', 'melhor envio'
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

function formatProviderName(name) {
    if (!name) return '';
    return name.replace(/_/g, ' ');
}

function copyToClipboard(text, btnElement) {
    navigator.clipboard.writeText(text).then(() => {
        const originalHTML = btnElement.innerHTML;
        btnElement.classList.add('copied');
        btnElement.innerHTML = `
            <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M5 13l4 4L19 7" />
            </svg>
        `;
        setTimeout(() => {
            btnElement.classList.remove('copied');
            btnElement.innerHTML = originalHTML;
        }, 1500);
    }).catch(err => {
        console.error('Failed to copy text: ', err);
    });
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

    // Check for delay warnings (e.g., any option taking more than 10 days)
    const hasLongDelay = options.some(opt => opt.days > 10);
    if (hasLongDelay) {
        delayWarningContainer.classList.remove('hidden');
    } else {
        delayWarningContainer.classList.add('hidden');
    }

    // Recommend Card
    bestMatchCard.innerHTML = `
        <div class="trophy-icon">
            <svg xmlns="http://www.w3.org/2000/svg" width="32" height="32" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M13 10V3L4 14h7v7l9-11h-7z" />
            </svg>
        </div>
        <h3>Menor Prazo Estimado</h3>
        <div class="winner-name">${formatProviderName(bestOption.provider)}</div>
        <div>
            <span class="winner-time">${bestOption.days} dias</span>
        </div>
    `;

    // Table
    const MAX_DAYS = Math.max(...options.map(o => o.days), 15); // for progress bar calculation

    options.forEach((opt, index) => {
        const tr = document.createElement('tr');

        let rowClass = "";
        let pbClass = "pb-success";

        if (index === 0) {
            rowClass = "row-success";
        } else if (opt.days > bestOption.days * 1.5) {
            rowClass = "row-warning";
            pbClass = opt.days > 10 ? "pb-error" : "pb-warning";
        }

        if (rowClass) tr.classList.add(rowClass);

        const widthPercentage = Math.min((opt.days / MAX_DAYS) * 100, 100).toFixed(0);

        tr.innerHTML = `
            <td class="provider-name">
                ${formatProviderName(opt.provider)}
                <div class="progress-container">
                    <div class="progress-bar ${pbClass}" style="width: 0%" data-target-width="${widthPercentage}"></div>
                </div>
            </td>
            <td class="time-badge">${opt.days} dias</td>
        `;
        resultsTable.appendChild(tr);

        // Animate progress bar after short delay
        setTimeout(() => {
            const pb = tr.querySelector('.progress-bar');
            if (pb) pb.style.width = pb.getAttribute('data-target-width') + '%';
        }, 150);
    });
}
