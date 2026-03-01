const dbInfo = {
    estabelecimentos: { filename: 'Base_cadastros_Unificada_Final.xlsx', data: [], headers: [], status: 'loading' },
    tecnicos: { filename: 'Base_TECNICOS_Prazos_Atualizados.xlsx', data: [], headers: [], status: 'loading' },
    cidades: { filename: 'cidades_atendidas_detalhado.xlsx', data: [], headers: [], status: 'loading' }
};

let currentTab = 'estabelecimentos';

const searchInput = document.getElementById('search-input');
const btnSearch = document.getElementById('btn-search');
const statusIndicator = document.getElementById('status-indicator');
const statusText = statusIndicator.querySelector('.text');
const heroSection = document.getElementById('hero-section');
const dashboard = document.getElementById('dashboard-results');
const btnClear = document.getElementById('btn-clear');

const estabDetailsCard = document.getElementById('estab-details');
const highlightCards = document.getElementById('highlight-cards');
const delayWarningContainer = document.getElementById('delay-warning-container');
const bestMatchCard = document.getElementById('best-match-card');
const resultsTable = document.querySelector('#results-table tbody');

const fileInput = document.getElementById('file-input');
const uploadFallback = document.getElementById('upload-fallback');

// Tabs
const tabBtns = document.querySelectorAll('.tab-btn');
const tabContents = document.querySelectorAll('.tab-content');

// Cidades elements
const cidadesHeader = document.getElementById('cidades-header');
const cidadesList = document.getElementById('cidades-list');

function normalizeSearch(str) {
    if (!str) return "";
    return String(str).toLowerCase().trim().replace(/[-.,]/g, '');
}

window.addEventListener('DOMContentLoaded', async () => {
    // Carregar todas as bases paralelamente
    await Promise.all([
        loadDatabase('estabelecimentos'),
        loadDatabase('tecnicos'),
        loadDatabase('cidades')
    ]);
    updateGlobalStatus();
});

async function loadDatabase(dbKey) {
    try {
        const response = await fetch(dbInfo[dbKey].filename);
        if (!response.ok) throw new Error('Network response was not ok');
        const arrayBuffer = await response.arrayBuffer();

        const data = new Uint8Array(arrayBuffer);
        const workbook = XLSX.read(data, { type: 'array' });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        const json = XLSX.utils.sheet_to_json(worksheet, { defval: null });

        if (json.length === 0) throw new Error("Planilha vazia.");

        dbInfo[dbKey].data = json;
        dbInfo[dbKey].headers = Object.keys(json[0]);
        dbInfo[dbKey].status = 'ready';
    } catch (error) {
        console.warn(`Could not load ${dbInfo[dbKey].filename}`, error);
        dbInfo[dbKey].status = 'error';
    }
}

function updateGlobalStatus() {
    const allReady = Object.values(dbInfo).every(db => db.status === 'ready');
    if (allReady) {
        statusIndicator.className = 'status ready';
        statusText.innerText = `Bases carregadas`;
        searchInput.disabled = false;
        searchInput.focus();
        updateSearchPlaceholder();
        uploadFallback.classList.add('hidden');
    } else {
        const allErrors = Object.values(dbInfo).every(db => db.status === 'error');
        if (allErrors) {
            statusIndicator.className = 'status error';
            statusText.innerText = "Falha ao carregar bases";
            uploadFallback.classList.remove('hidden');
        } else {
            statusIndicator.className = 'status warning';
            statusText.innerText = "Atenção: Nem todas as bases carregaram.";
            searchInput.disabled = false;
            updateSearchPlaceholder();
        }
    }
}

function updateSearchPlaceholder() {
    if (currentTab === 'estabelecimentos') {
        searchInput.placeholder = "Digite o CEP ou Terminal (Ex: 06455-000)";
    } else if (currentTab === 'tecnicos') {
        searchInput.placeholder = "Digite o nome do Técnico";
    } else if (currentTab === 'cidades') {
        searchInput.placeholder = "Digite a Cidade ou o Técnico";
    }
}

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

// Lógica das Abas
tabBtns.forEach(btn => {
    btn.addEventListener('click', () => {
        // Remover classes ativas
        tabBtns.forEach(b => b.classList.remove('active'));
        tabContents.forEach(c => c.classList.remove('active'));

        // Adicionar ativa na aba clicada
        btn.classList.add('active');
        currentTab = btn.getAttribute('data-tab');

        // Resetar interface
        searchInput.value = '';
        dashboard.classList.remove('visible');
        heroSection.classList.remove('minimized');
        updateSearchPlaceholder();

        // Ativar aba correspondente
        if (currentTab === 'estabelecimentos' || currentTab === 'tecnicos') {
            document.getElementById('prazos-results').classList.add('active');
        } else {
            document.getElementById('cidades-results').classList.add('active');
        }
    });
});


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

    if (dbInfo[currentTab].status !== 'ready') {
        alert("A base desta aba não pôde ser carregada.");
        return;
    }

    const query = normalizeSearch(rawQuery);
    const data = dbInfo[currentTab].data;

    // Tratamento para Estabelecimentos
    if (currentTab === 'estabelecimentos') {
        let match = data.find(row => normalizeSearch(row['CEP'] || row['cep']) === query);
        if (!match) {
            match = data.find(row => {
                for (let key in row) {
                    if (normalizeSearch(row[key])?.includes(query)) return true;
                }
                return false;
            });
        }

        if (!match) return showSearchError();

        renderEstablishmentInfo(match);
        analyzeDeliveryTimes(match, dbInfo[currentTab].headers);
    }
    // Tratamento para Técnicos
    else if (currentTab === 'tecnicos') {
        let match = data.find(row => normalizeSearch(row['TECNICO'] || row['tecnico'])?.includes(query));
        if (!match) {
            match = data.find(row => {
                for (let key in row) {
                    if (normalizeSearch(row[key])?.includes(query)) return true;
                }
                return false;
            });
        }

        if (!match) return showSearchError();

        renderTecnicoInfo(match);
        analyzeDeliveryTimes(match, dbInfo[currentTab].headers);
    }
    // Tratamento para Cidades Atendidas
    else if (currentTab === 'cidades') {
        const matches = data.filter(row => {
            for (let key in row) {
                if (normalizeSearch(row[key])?.includes(query)) return true;
            }
            return false;
        });

        if (matches.length === 0) return showSearchError();

        renderCidadesInfo(matches, query);
    }
}

function showSearchError() {
    searchInput.style.animation = 'shake 0.5s ease';
    setTimeout(() => searchInput.style.animation = '', 500);
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

function formatCnpj(val) {
    if (!val) return '';
    const str = String(val).replace(/\D/g, '');
    if (str.length === 14) {
        return str.replace(/^(\d{2})(\d{3})(\d{3})(\d{4})(\d{2})$/, "$1.$2.$3/$4-$5");
    }
    return val;
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

function renderTecnicoInfo(row) {
    const nome = getColValue(row, ['TECNICO', 'Tecnico']) || 'N/A';
    const bairro = getColValue(row, ['BAIRRO']) || '';
    const cidade = getColValue(row, ['CIDADE', 'Cidade ']) || 'N/A';
    const uf = getColValue(row, ['ESTADO', 'UF']) || 'N/A';
    const regiao = getColValue(row, ['REGIÃO', 'Regiao']) || 'N/A';

    const telefone = getColValue(row, ['TELEFONE']) || '';
    const cnpj = getColValue(row, ['CNPJ']) || '';

    estabDetailsCard.innerHTML = `
        <div class="estab-header">
            <h3>Técnico: ${nome}</h3>
            <div class="badges-container">
                <span class="badge" style="background: #dbeafe; color: #1e40af;">${regiao}</span>
            </div>
        </div>
        <div class="estab-info-grid">
            <div class="estab-field">
                <span class="label">Localização Base</span>
                <span class="val">${bairro ? bairro + ' - ' : ''}${cidade}/${uf}</span>
            </div>
            <div class="estab-field">
                <span class="label">Contato</span>
                <span class="val">${telefone || 'Indisponível'}</span>
            </div>
            <div class="estab-field">
                <span class="label">CNPJ</span>
                <span class="val">${formatCnpj(cnpj) || 'N/A'}</span>
            </div>
        </div>
    `;

    highlightCards.innerHTML = '';
    highlightCards.style.display = 'none';
}

function renderCidadesInfo(matches, query) {
    heroSection.classList.add('minimized');
    setTimeout(() => dashboard.classList.add('visible'), 100);

    const matchIsTecnico = matches[0] && getColValue(matches[0], ['TÉCNICO', 'Técnico', 'Tecnico']) && normalizeSearch(getColValue(matches[0], ['TÉCNICO', 'Técnico', 'Tecnico'])).includes(query);

    if (matchIsTecnico) {
        const nomeTecnico = getColValue(matches[0], ['TÉCNICO', 'Técnico', 'Tecnico']) || 'Sem Nome';
        cidadesHeader.innerHTML = `
            <h3>Técnico ${nomeTecnico}</h3>
            <p>Atende a <strong>${matches.length}</strong> cidades na região.</p>
        `;

        cidadesList.innerHTML = matches.map(m => `
            <div class="cidade-card">
                <div>
                    <span class="cidade-name" style="display:block; margin-bottom:0.25rem;">${getColValue(m, ['CIDADE ATENDIDA', 'Cidade Atendida', 'cidade atendida', 'Cidade']) || 'N/A'}</span>
                    <span style="font-size: 0.75rem; color: #64748b;">Distância: ${getColValue(m, ['DISTÂNCIA (KM)', 'Distância']) || '0'}km</span>
                </div>
                <span class="cidade-uf">${getColValue(m, ['UF ATENDIDA', 'UF Atendida', 'uf atendida', 'Estado', 'UF', 'Uf']) || 'N/A'}</span>
            </div>
        `).join('');
    } else {
        // Ordena os técnicos por distância
        const sortedMatches = matches.sort((a, b) => {
            const getVal = (row) => String(getColValue(row, ['DISTÂNCIA (KM)', 'Distância']) || '0').replace(',', '.');
            return parseFloat(getVal(a)) - parseFloat(getVal(b));
        });

        const rec = sortedMatches[0];
        const others = sortedMatches.slice(1);

        const recCityName = getColValue(rec, ['CIDADE ATENDIDA', 'Cidade Atendida', 'cidade atendida', 'Cidade']) || 'N/A';

        // Usuário procurou pela cidade
        cidadesHeader.innerHTML = `
            <h3>Cidade encontrada: ${recCityName}</h3>
            <p>Esta cidade é atendida por <strong>${matches.length}</strong> técnico(s):</p>
        `;

        let html = `
            <div class="cidade-card" style="grid-column: 1 / -1; border-color: #22c55e; border-width: 2px; background: #f0fdf4;">
                <div style="display: flex; align-items: center; justify-content: space-between; width: 100%;">
                    <div>
                        <div style="color: #16a34a; font-size: 0.75rem; font-weight: 800; margin-bottom: 0.25rem;">✨ TÉCNICO MAIS PRÓXIMO</div>
                        <span class="cidade-name" style="display:block; margin-bottom:0.25rem; font-size: 1.1rem; color: #166534;">${getColValue(rec, ['TÉCNICO', 'Técnico', 'Tecnico']) || 'N/A'}</span>
                        <span style="font-size: 0.85rem; color: #15803d;">De: ${getColValue(rec, ['CIDADE BASE', 'Cidade Base']) || 'N/A'} <strong>(${getColValue(rec, ['DISTÂNCIA (KM)', 'Distância']) || '0'}km)</strong></span>
                    </div>
                    <span class="cidade-uf" style="background: #bbf7d0; color: #166534;">${getColValue(rec, ['UF ATENDIDA', 'UF Atendida', 'UF', 'Estado', 'Uf']) || 'N/A'}</span>
                </div>
            </div>
        `;

        // Se houver mais técnicos
        if (others.length > 0) {
            html += `<h4 style="grid-column: 1 / -1; margin-top: 1rem; color: var(--text-secondary); font-size: 0.9rem;">Outras opções de técnicos:</h4>`;
            html += others.map(m => `
                <div class="cidade-card">
                    <div>
                        <span class="cidade-name" style="display:block; margin-bottom:0.25rem;">${getColValue(m, ['TÉCNICO', 'Técnico', 'Tecnico']) || 'N/A'}</span>
                        <span style="font-size: 0.75rem; color: #64748b;">De: ${getColValue(m, ['CIDADE BASE', 'Cidade Base']) || 'N/A'} (${getColValue(m, ['DISTÂNCIA (KM)', 'Distância']) || '0'}km)</span>
                    </div>
                    <span class="cidade-uf">${getColValue(m, ['UF ATENDIDA', 'UF Atendida', 'UF', 'Estado', 'Uf']) || 'N/A'}</span>
                </div>
            `).join('');
        }

        cidadesList.innerHTML = html;
    }
}

function analyzeDeliveryTimes(row, headersList) {
    // We ignore typical descriptive columns AND explicitly ignore the columns user mentioned
    const ignoreKeywords = [
        'cep', 'terminal', 'cidade', 'estado', 'uf', 'região', 'regiao', 'id', 'codigo', 'origem',
        'endereco', 'endereço', 'estab novo', 'nome', 'grupo', 'razao social', 'logradouro',
        'bairro', 'tipo operacao', 'técnico', 'tecnico', 'melhor forma', 'melhor envio', 'prazo correios',
        'empresa', 'treinamento', 'instalador', 'prospectador', 'parceiro', 'usuario', 'backup', 'cnpj', 'email', 'telefone', 'dat de'
    ];

    const deliveryOptions = [];

    for (let i = 0; i < headersList.length; i++) {
        const colName = headersList[i];
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
