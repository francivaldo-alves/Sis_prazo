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

function setupTabs() {
    tabBtns.forEach(btn => {
        btn.addEventListener('click', () => {
            // Remove active class from all buttons and contents
            tabBtns.forEach(b => b.classList.remove('active'));
            tabContents.forEach(c => c.classList.remove('active'));

            // Add active class to clicked button and corresponding content
            btn.classList.add('active');
            const targetId = btn.getAttribute('data-tab') + '-results';

            // Handle the specific ID mapping for Establishments/Tecnicos which share 'prazos-results'
            let contentId = targetId;
            if (btn.getAttribute('data-tab') === 'estabelecimentos' || btn.getAttribute('data-tab') === 'tecnicos') {
                contentId = 'prazos-results';
            }

            const targetContent = document.getElementById(contentId);
            if (targetContent) {
                targetContent.classList.add('active');
            }

            // Update current tab status for search
            currentTab = btn.getAttribute('data-tab');

            // Clear current search results
            document.getElementById('dashboard-results').classList.remove('visible');
            searchInput.value = '';
            document.getElementById('search-suggestions').style.display = 'none';
        });
    });
}

// Cidades elements
const cidadesHeader = document.getElementById('cidades-header');
const cidadesList = document.getElementById('cidades-list');

function normalizeSearch(str) {
    if (!str) return "";
    return String(str)
        .normalize("NFD")
        .replace(/[\u0300-\u036f]/g, "")
        .toLowerCase()
        .trim()
        .replace(/[-.,]/g, '');
}

window.addEventListener('DOMContentLoaded', async () => {
    // Inicializar os manipuladores de UI
    setupSearch();
    setupTabs();

    // Carregar todas as bases paralelamente
    await Promise.all([
        loadDatabase('estabelecimentos', dbInfo.estabelecimentos.filename),
        loadDatabase('tecnicos', dbInfo.tecnicos.filename),
        loadDatabase('cidades', dbInfo.cidades.filename)
    ]);
    updateGlobalStatus();
});

async function loadDatabase(key, filePath) {
    dbInfo[key].status = 'loading';
    // Assuming updateDashboardUI is a function that updates the UI based on dbInfo statuses
    // If it doesn't exist, this line might cause an error. For now, I'll assume it's intended.
    // If it's meant to be updateGlobalStatus, then that function should be called.
    // Given the context, updateGlobalStatus is likely the intended function to call here.
    updateGlobalStatus(); // Changed from updateDashboardUI to updateGlobalStatus for consistency

    try {
        const response = await fetch(filePath);
        if (!response.ok) throw new Error(`HTTP error! status: ${response.status}`);

        // Verifica idade da base de dados se for estabelecimentos ou tecnicos (as principais)
        if (key === 'estabelecimentos' || key === 'tecnicos') {
            const lastModified = response.headers.get('Last-Modified');
            if (lastModified) {
                const modDate = new Date(lastModified);
                const daysOld = (Date.now() - modDate.getTime()) / (1000 * 60 * 60 * 24);

                if (daysOld > 7) {
                    const warningBar = document.getElementById('delay-warning-container');
                    if (warningBar) {
                        warningBar.style.display = 'block';
                        warningBar.innerHTML = `
                            <div class="card row-warning" style="margin-bottom: 1rem; border: 1px solid #f59e0b; background: rgba(245, 158, 11, 0.1);">
                                <div style="display: flex; gap: 0.75rem; align-items: flex-start; color: #d97706;">
                                    <svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" fill="none" viewBox="0 0 24 24" stroke="currentColor" style="flex-shrink:0;">
                                        <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 9v2m0 4h.01m-6.938 4h13.856c1.54 0 2.502-1.667 1.732-3L13.732 4c-.77-1.333-2.694-1.333-3.464 0L3.34 16c-.77 1.333.192 3 1.732 3z"/>
                                    </svg>
                                    <div>
                                        <h4 style="margin-bottom: 0.25rem;">Atenção: Base Desatualizada</h4>
                                        <p style="font-size: 0.9rem;">A base de dados <strong>${key}</strong> tem mais de ${Math.floor(daysOld)} dias e pode conter informações não validadas recentemente.</p>
                                    </div>
                                </div>
                            </div>
                        `;
                    }
                }
            }
        }

        const arrayBuffer = await response.arrayBuffer();
        const data = new Uint8Array(arrayBuffer);
        const workbook = XLSX.read(data, { type: 'array' });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        const json = XLSX.utils.sheet_to_json(worksheet, { defval: null });

        if (json.length === 0) throw new Error("Planilha vazia.");

        dbInfo[key].data = json;
        dbInfo[key].headers = Object.keys(json[0]);
        dbInfo[key].status = 'ready';
    } catch (e) {
        console.error(`Erro ao carregar ${key}:`, e);
        dbInfo[key].status = 'error';
    } finally {
        updateGlobalStatus(); // Changed from updateDashboardUI to updateGlobalStatus for consistency
    }
}

function updateGlobalStatus() {
    const allReady = Object.values(dbInfo).every(db => db.status === 'ready');
    if (allReady) {
        statusIndicator.className = 'status ready';
        statusText.innerText = `Bases carregadas`;
        searchInput.disabled = false;
        btnSearch.disabled = false;
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
            btnSearch.disabled = false;
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

function setupSearch() {
    // Create a mini spinner for inside the search button
    const btnOriginalText = btnSearch.innerHTML;
    const btnLoadingText = `<svg class="spinner" viewBox="0 0 50 50" style="width:20px;height:20px;animation:rotate 2s linear infinite;"><circle cx="25" cy="25" r="20" fill="none" stroke="currentColor" stroke-width="5" stroke-linecap="round" style="stroke-dasharray: 1, 200; stroke-dashoffset: 0; animation: dash 1.5s ease-in-out infinite;"/></svg> Buscando...`;

    const searchSuggestions = document.getElementById('search-suggestions');

    const triggerSearch = (forceQuery) => {
        const query = (forceQuery || searchInput.value).trim();
        if (!query) return;

        searchSuggestions.style.display = 'none'; // hide suggestions
        searchInput.value = query; // fill input

        // Simula o tempo de busca
        btnSearch.innerHTML = btnLoadingText;
        btnSearch.disabled = true;
        document.getElementById('dashboard-results').style.opacity = '0.5';

        setTimeout(() => {
            handleSearch(query);
            btnSearch.innerHTML = btnOriginalText;
            btnSearch.disabled = false;
            document.getElementById('dashboard-results').style.opacity = '1';
        }, 400); // 400ms loading pra dar UX feeling
    };

    btnSearch.addEventListener('click', () => triggerSearch());

    let debounceTimeout;
    searchInput.addEventListener('input', (e) => {
        const val = e.target.value.trim();
        clearTimeout(debounceTimeout);

        if (val.length < 2 || !dbInfo[currentTab] || dbInfo[currentTab].status !== 'ready') {
            searchSuggestions.style.display = 'none';
            return;
        }

        debounceTimeout = setTimeout(() => {
            const queryNorm = normalizeSearch(val);
            const data = dbInfo[currentTab].data;
            const suggestions = [];

            // Popula sugestões baseado na aba atual
            for (let i = 0; i < data.length; i++) {
                if (suggestions.length >= 8) break; // max 8 suggestions
                const row = data[i];

                if (currentTab === 'estabelecimentos') {
                    const term = normalizeSearch(getColValue(row, ['Terminal', 'terminal', 'Terminal ', 'TERMINAL']) || '');
                    const cep = normalizeSearch(getColValue(row, ['CEP', 'cep']) || '');
                    const name = normalizeSearch(getColValue(row, ['Nome', 'Razao Social', 'Estabelecimento']) || '');

                    if (term.includes(queryNorm) || cep.includes(queryNorm) || name.includes(queryNorm)) {
                        const originalTerm = getColValue(row, ['Terminal', 'terminal', 'Terminal ', 'TERMINAL']) || '';
                        const originalName = getColValue(row, ['Nome', 'Razao Social', 'Estabelecimento']) || '';
                        suggestions.push(originalTerm ? `${originalTerm} - ${originalName}` : originalName);
                    }
                } else if (currentTab === 'tecnicos') {
                    const name = normalizeSearch(getColValue(row, ['TECNICO', 'Tecnico']) || '');
                    if (name.includes(queryNorm) && normalizeSearch(getColValue(row, ['TECNICO', 'Tecnico'])) !== 'n/a') {
                        const originalName = getColValue(row, ['TECNICO', 'Tecnico']);
                        if (!suggestions.includes(originalName)) suggestions.push(originalName);
                    }
                } else if (currentTab === 'cidades') {
                    const city = normalizeSearch(getColValue(row, ['CIDADE ATENDIDA', 'Cidade Atendida', 'cidade atendida', 'Cidade']) || '');
                    if (city.includes(queryNorm)) {
                        const originalCity = getColValue(row, ['CIDADE ATENDIDA', 'Cidade Atendida', 'cidade atendida', 'Cidade']);
                        if (!suggestions.includes(originalCity)) suggestions.push(originalCity);
                    }
                }
            }

            if (suggestions.length > 0) {
                searchSuggestions.innerHTML = suggestions.map(s => `
                    <div class="suggestion-item" style="padding: 10px 16px; cursor: pointer; border-bottom: 1px solid var(--card-border);" onmouseover="this.style.background='rgba(99, 102, 241, 0.1)'" onmouseout="this.style.background='transparent'" onclick="document.getElementById('search-input').value='${s.split(' - ')[0]}'; document.getElementById('btn-search').click();">
                        ${s}
                    </div>
                `).join('');
                searchSuggestions.style.display = 'flex';
            } else {
                searchSuggestions.style.display = 'none';
            }
        }, 300);
    });

    searchInput.addEventListener('keypress', (e) => {
        if (e.key === 'Enter') {
            clearTimeout(debounceTimeout);
            searchSuggestions.style.display = 'none';
            triggerSearch();
        }
    });

    // Fechar sugestoes ao clicar fora
    document.addEventListener('click', (e) => {
        if (!e.target.closest('.search-box')) {
            searchSuggestions.style.display = 'none';
        }
    });

    btnClear.addEventListener('click', () => {
        dashboard.classList.remove('visible');
        heroSection.classList.remove('minimized');
        searchInput.value = '';
        searchInput.focus();
        searchSuggestions.style.display = 'none';
    });
}

// Lógica das Abas
tabBtns.forEach(btn => {
    btn.addEventListener('click', () => {
        // Remover classes ativas
        tabBtns.forEach(b => b.classList.remove('active'));
        // tabContents.forEach(c => c.classList.remove('active')); // Original line
        btn.classList.add('active');

        document.querySelectorAll('.tab-content').forEach(tc => tc.classList.remove('active'));
        currentTab = btn.dataset.tab;

        const contentId = currentTab === 'estabelecimentos' ? 'prazos-results' :
            currentTab === 'tecnicos' ? 'prazos-results' :
                'cidades-results';

        document.getElementById(contentId).classList.add('active');

        if (contentId === 'prazos-results') {
            document.getElementById('estab-details').style.display = currentTab === 'tecnicos' ? 'block' : 'block';
            document.getElementById('highlight-cards').style.display = currentTab === 'tecnicos' ? 'none' : 'grid';
            document.getElementById('delay-warning-container').style.display = 'none';
        }

        searchInput.placeholder = currentTab === 'estabelecimentos' ? 'Digite o CEP ou Terminal (Ex: 04359-000 ou 35)' :
            currentTab === 'tecnicos' ? 'Digite o nome do Técnico (Ex: ALEX RAMOS COSTA)' :
                'Digite a Cidade Atendida (Ex: Carapicuíba)';

        if (dashboard.classList.contains('active')) {
            handleSearch(searchInput.value);
        }
    });
});

setTimeout(() => { searchInput.focus(); }, 300);
// The closing '});' for DOMContentLoaded was moved to the top.
function handleSearch(rawQuery) { // Renamed from performSearch and now accepts query
    if (!rawQuery) return;

    if (!dbInfo[currentTab] || dbInfo[currentTab].status !== 'ready') {
        return;
    }

    const query = normalizeSearch(rawQuery);
    const data = dbInfo[currentTab].data;

    // Tratamento para Estabelecimentos
    if (currentTab === 'estabelecimentos') {
        // Primeiro tenta encontrar correspondência EXATA de Terminal ou CEP (para buscas curtas ex: '35')
        let match = data.find(row => {
            const cepRaw = getColValue(row, ['CEP', 'cep']);
            const termRaw = getColValue(row, ['Terminal', 'terminal', 'Terminal ', 'TERMINAL']);

            const cep = normalizeSearch(cepRaw);
            const term = normalizeSearch(termRaw);

            return (cep && cep === query) || (term && term === query);
        });

        // Se não achar exato, tenta correspondência parcial pelo nome ou parte do terminal
        if (!match) {
            match = data.find(row => {
                const termRaw = getColValue(row, ['Terminal', 'terminal', 'Terminal ', 'TERMINAL']);
                const nameRaw = getColValue(row, ['Nome', 'Razao Social', 'Estabelecimento']);

                const term = normalizeSearch(termRaw);
                const name = normalizeSearch(nameRaw);

                if (term && term.includes(query)) return true;
                if (name && name.includes(query)) return true;
                return false;
            });
        }

        // Em último caso, procura a query em qualquer coluna da linha
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
        let techPhone = '';
        let techDist = '';

        // Try getting phone from DB 'tecnicos'
        const techData = dbInfo['tecnicos']?.data || [];
        const techMatch = techData.find(t => normalizeSearch(getColValue(t, ['TECNICO', 'Técnico', 'Tecnico'])) === normalizeSearch(tecnico));

        if (techMatch) {
            let phone = getColValue(techMatch, ['TELEFONE', 'Telefone', 'Celular']);
            if (phone) {
                let strPhone = String(phone).replace(/\D/g, '');
                if (strPhone.length === 11) {
                    techPhone = `(${strPhone.substring(0, 2)}) ${strPhone.substring(2, 7)}-${strPhone.substring(7)}`;
                } else if (strPhone.length === 10) {
                    techPhone = `(${strPhone.substring(0, 2)}) ${strPhone.substring(2, 6)}-${strPhone.substring(6)}`;
                } else {
                    techPhone = phone;
                }
            }
        }

        // Try getting dist from current row (usually columns like 'Distância', 'Distância (KM)')
        // Or if it's named 'KM', 'KM Tecnico'
        techDist = getColValue(row, ['DISTÂNCIA', 'Distância', 'Distancia', 'DISTÂNCIA (KM)', 'KM', 'KM TÉCNICO', 'KM TECNICO']);

        // If not found in the same row, cross-reference with 'cidades' database!
        if (!techDist) {
            const estabCity = getColValue(row, ['Cidade', 'Municipio', 'CIDADE']);
            const citiesData = dbInfo['cidades']?.data || [];
            if (estabCity && citiesData.length > 0) {
                // Find a match where the Technician Name and the Attended City match
                const cityMatch = citiesData.find(c => {
                    const cTech = getColValue(c, ['TÉCNICO', 'Técnico', 'Tecnico']);
                    const cCity = getColValue(c, ['CIDADE ATENDIDA', 'Cidade Atendida', 'cidade atendida', 'Cidade']);

                    return cTech && cCity &&
                        normalizeSearch(cTech) === normalizeSearch(tecnico) &&
                        normalizeSearch(cCity) === normalizeSearch(estabCity);
                });

                if (cityMatch) {
                    techDist = getColValue(cityMatch, ['DISTÂNCIA (KM)', 'Distância']);
                }
            }
        }

        let whatsAppLink = '';
        if (techPhone.includes('(')) {
            // Remove tudo exceto números para o link do whatsapp
            const justNumbers = techPhone.replace(/\D/g, '');
            whatsAppLink = `https://wa.me/55${justNumbers}`;
        }

        let mapsLink = '';
        if (techDist) {
            // Get estab city and tech base city for maps query
            const estabCity = getColValue(row, ['Cidade', 'Municipio', 'CIDADE']) || '';
            const techCity = (dbInfo['tecnicos']?.data || []).find(t => normalizeSearch(getColValue(t, ['TECNICO', 'Técnico', 'Tecnico'])) === normalizeSearch(tecnico));
            const baseCityName = techCity ? getColValue(techCity, ['CIDADE', 'Cidade ']) : '';
            if (estabCity && baseCityName) {
                mapsLink = `https://www.google.com/maps/dir/${encodeURIComponent(baseCityName)}/${encodeURIComponent(estabCity)}`;
            }
        }

        highlightCards.innerHTML += `
            <div class="highlight-card tech-card" style="display: flex; flex-direction: column; justify-content: space-between; gap: 1rem;">
                <div style="display: flex; align-items: center; gap: 1rem;">
                    <div class="highlight-icon tech-icon">
                        <svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                            <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M16 7a4 4 0 11-8 0 4 4 0 018 0zM12 14a7 7 0 00-7 7h14a7 7 0 00-7-7z" />
                        </svg>
                    </div>
                    <div class="highlight-content">
                        <span class="label">Técnico Mais Próximo</span>
                        <span class="val" style="font-size: 1.25rem;">${tecnico}</span>
                    </div>
                </div>
                
                <div style="display: flex; flex-wrap: wrap; gap: 0.75rem; align-items: center; justify-content: flex-start; margin-top: auto;">
                    ${techDist ? `
                    ${mapsLink ? `<a href="${mapsLink}" target="_blank" style="text-decoration: none;" title="Ver Rota no Mapa">` : ''}
                    <span style="font-size: 0.85rem; background: rgba(99, 102, 241, 0.1); color: #818cf8; padding: 4px 10px; border-radius: 6px; display: inline-flex; align-items: center; gap: 6px; border: 1px solid rgba(99, 102, 241, 0.2); transition: background 0.2s; cursor: ${mapsLink ? 'pointer' : 'default'};" onmouseover="this.style.background='rgba(99, 102, 241, 0.2)'" onmouseout="this.style.background='rgba(99, 102, 241, 0.1)'">
                        <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                            <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M17.657 16.657L13.414 20.9a1.998 1.998 0 01-2.827 0l-4.244-4.243a8 8 0 1111.314 0z"></path>
                            <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M15 11a3 3 0 11-6 0 3 3 0 016 0z"></path>
                        </svg>
                        <strong>${techDist}km</strong>
                    </span>
                    ${mapsLink ? `</a>` : ''}` : ''}
                    
                    ${techPhone ? `
                    <a href="${whatsAppLink}" target="_blank" style="text-decoration: none;" title="Abrir no WhatsApp">
                        <span style="font-size: 0.85rem; background: rgba(16, 185, 129, 0.1); color: #34d399; padding: 4px 10px; border-radius: 6px; display: inline-flex; align-items: center; gap: 6px; border: 1px solid rgba(16, 185, 129, 0.2); transition: background 0.2s; cursor: pointer;" onmouseover="this.style.background='rgba(16, 185, 129, 0.2)'" onmouseout="this.style.background='rgba(16, 185, 129, 0.1)'">
                            <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                                <path d="M22 16.92v3a2 2 0 0 1-2.18 2 19.79 19.79 0 0 1-8.63-3.07 19.5 19.5 0 0 1-6-6 19.79 19.79 0 0 1-3.07-8.67A2 2 0 0 1 4.11 2h3a2 2 0 0 1 2 1.72 12.84 12.84 0 0 0 .7 2.81 2 2 0 0 1-.45 2.11L8.09 9.91a16 16 0 0 0 6 6l1.27-1.27a2 2 0 0 1 2.11-.45 12.84 12.84 0 0 0 2.81.7A2 2 0 0 1 22 16.92z"></path>
                            </svg>
                            ${techPhone}
                        </span>
                    </a>` : ''}
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

    const telefoneRaw = getColValue(row, ['TELEFONE']) || '';
    const cnpj = getColValue(row, ['CNPJ']) || '';

    let telefone = telefoneRaw;
    let whatsAppLink = '';

    if (telefoneRaw) {
        let strPhone = String(telefoneRaw).replace(/\D/g, '');
        if (strPhone.length === 11) {
            telefone = `(${strPhone.substring(0, 2)}) ${strPhone.substring(2, 7)}-${strPhone.substring(7)}`;
        } else if (strPhone.length === 10) {
            telefone = `(${strPhone.substring(0, 2)}) ${strPhone.substring(2, 6)}-${strPhone.substring(6)}`;
        }
        whatsAppLink = `https://wa.me/55${strPhone}`;
    }

    const mapsLink = `https://www.google.com/maps/search/?api=1&query=${encodeURIComponent(`${bairro ? bairro + ' - ' : ''}${cidade} ${uf}`)}`;

    estabDetailsCard.innerHTML = `
        <div class="estab-header">
            <h3>Técnico: ${nome}</h3>
            <span class="badge" style="background: rgba(16, 185, 129, 0.2); color: #34d399; border: 1px solid rgba(16, 185, 129, 0.3); padding: 5px 12px; border-radius: 20px; font-weight: 500;">
                ${regiao}
            </span>
        </div>
        <div class="estab-grid">
            <div class="info-item">
                <span class="label">Localização Base</span>
                <a href="${mapsLink}" target="_blank" style="text-decoration: none; color: inherit;" title="Ver no Mapa">
                    <span class="val" style="display: flex; align-items: center; gap: 6px; cursor: pointer;">
                        <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" fill="none" viewBox="0 0 24 24" stroke="#818cf8">
                            <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M17.657 16.657L13.414 20.9a1.998 1.998 0 01-2.827 0l-4.244-4.243a8 8 0 1111.314 0z"/>
                        </svg>
                        ${bairro ? bairro + ' - ' : ''}${cidade}/${uf}
                    </span>
                </a>
            </div>
            ${telefone ? `
            <div class="info-item">
                <span class="label">Contato</span>
                <a href="${whatsAppLink}" target="_blank" style="text-decoration: none; color: inherit;" title="Abrir no WhatsApp">
                    <span class="val" style="display: inline-flex; align-items: center; gap: 6px; background: rgba(16, 185, 129, 0.1); color: #34d399; padding: 4px 10px; border-radius: 6px; border: 1px solid rgba(16, 185, 129, 0.2); transition: background 0.2s; cursor: pointer;" onmouseover="this.style.background='rgba(16, 185, 129, 0.2)'" onmouseout="this.style.background='rgba(16, 185, 129, 0.1)'">
                        <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                            <path d="M22 16.92v3a2 2 0 0 1-2.18 2 19.79 19.79 0 0 1-8.63-3.07 19.5 19.5 0 0 1-6-6 19.79 19.79 0 0 1-3.07-8.67A2 2 0 0 1 4.11 2h3a2 2 0 0 1 2 1.72 12.84 12.84 0 0 0 .7 2.81 2 2 0 0 1-.45 2.11L8.09 9.91a16 16 0 0 0 6 6l1.27-1.27a2 2 0 0 1 2.11-.45 12.84 12.84 0 0 0 2.81.7A2 2 0 0 1 22 16.92z"></path>
                        </svg>
                        ${telefone}
                    </span>
                </a>
            </div>` : ''}          <div class="info-item">
                <span class="label">CNPJ</span>
                <span class="val">${formatCnpj(cnpj) || 'N/A'}</span>
            </div>
        </div>
    `;

    highlightCards.innerHTML = '';
    highlightCards.style.display = 'none';
}

function renderCidadesInfo(rawMatches, query) {
    let isTecnico = false;
    let filteredMatches = [];

    // Priorizamos buscar pelo nome do técnico (técnico pode ser pesquisa parcial)
    const techMatches = rawMatches.filter(row => {
        const val = getColValue(row, ['TÉCNICO', 'Técnico', 'Tecnico']);
        return val && normalizeSearch(val).includes(query);
    });

    if (techMatches.length > 0) {
        isTecnico = true;
        filteredMatches = techMatches;
    } else {
        // Se não for técnico, filtra explicitamente para que a "Cidade Atendida" seja EXATAMENTE a pesquisada
        filteredMatches = rawMatches.filter(row => {
            const val = getColValue(row, ['CIDADE ATENDIDA', 'Cidade Atendida', 'cidade atendida', 'Cidade']);
            return val && normalizeSearch(val) === query;
        });

        // Caso a busca por Cidade Atendida exata não retorne nada, tenta buscar por inclusão (comportamento original)
        // Isso ajuda caso o usuário digite parte do nome da cidade sem querer e não seja técnico
        if (filteredMatches.length === 0) {
            filteredMatches = rawMatches.filter(row => {
                const val = getColValue(row, ['CIDADE ATENDIDA', 'Cidade Atendida', 'cidade atendida', 'Cidade']);
                return val && normalizeSearch(val).includes(query);
            });
            // Se ainda assim for 0, volta pra match bruto
            if (filteredMatches.length === 0) {
                filteredMatches = rawMatches;
            }
        }
    }

    // Deduplicar técnicos para a MESMA "Cidade Atendida", mantendo a menor distância
    const uniqueMatchesMap = new Map();

    filteredMatches.forEach(row => {
        // Se estivermos buscando um técnico, a chave de desduplicação será o Técnico + A Cidade que ele atende
        // Se estivermos buscando uma cidade, a chave será apenas o Técnico (pois já filtramos pela cidade acima)

        const tecnico = getColValue(row, ['TÉCNICO', 'Técnico', 'Tecnico']) || 'N/A';
        const cidadeAtendida = getColValue(row, ['CIDADE ATENDIDA', 'Cidade Atendida', 'cidade atendida', 'Cidade']) || 'N/A';
        const distRaw = getColValue(row, ['DISTÂNCIA (KM)', 'Distância']) || '0';
        const dist = parseFloat(String(distRaw).replace(',', '.')) || 0;

        const key = isTecnico ?
            `${normalizeSearch(tecnico)}||${normalizeSearch(cidadeAtendida)}` :
            `${normalizeSearch(tecnico)}`;

        if (!uniqueMatchesMap.has(key)) {
            uniqueMatchesMap.set(key, row);
        } else {
            const existingRow = uniqueMatchesMap.get(key);
            const existingDistRaw = getColValue(existingRow, ['DISTÂNCIA (KM)', 'Distância']) || '0';
            const existingDist = parseFloat(String(existingDistRaw).replace(',', '.')) || 0;

            if (dist < existingDist) {
                uniqueMatchesMap.set(key, row);
            }
        }
    });

    const matches = Array.from(uniqueMatchesMap.values());

    heroSection.classList.add('minimized');
    setTimeout(() => dashboard.classList.add('visible'), 100);

    // Função auxiliar para buscar telefone do técnico
    const getTechPhone = (techName) => {
        if (!techName) return '';
        const techData = dbInfo['tecnicos'].data;
        if (!techData || techData.length === 0) return '';
        const techMatch = techData.find(t => normalizeSearch(getColValue(t, ['TECNICO', 'Técnico', 'Tecnico'])) === normalizeSearch(techName));
        if (techMatch) {
            let phone = getColValue(techMatch, ['TELEFONE', 'Telefone', 'Celular']);
            if (phone) {
                // Remove tudo que não for dígito
                let strPhone = String(phone).replace(/\D/g, '');
                // Formatação: (11) 95871-1836
                if (strPhone.length === 11) {
                    return `(${strPhone.substring(0, 2)}) ${strPhone.substring(2, 7)}-${strPhone.substring(7)}`;
                } else if (strPhone.length === 10) {
                    return `(${strPhone.substring(0, 2)}) ${strPhone.substring(2, 6)}-${strPhone.substring(6)}`;
                }
                return phone; // retorna original se não bater o tamanho
            }
        }
        return '';
    };

    if (isTecnico) {
        const nomeTecnico = getColValue(matches[0], ['TÉCNICO', 'Técnico', 'Tecnico']) || 'Sem Nome';
        const telefone = getTechPhone(nomeTecnico);

        // Buscar informações extras na base de Técnicos
        const techData = dbInfo['tecnicos']?.data || [];
        const techRow = techData.find(t => normalizeSearch(getColValue(t, ['TECNICO', 'Técnico', 'Tecnico'])) === normalizeSearch(nomeTecnico));

        let baseRegiaoHtml = '';
        if (techRow) {
            const cidadeBase = getColValue(techRow, ['CIDADE', 'Cidade ']) || '';
            const ufBase = getColValue(techRow, ['ESTADO', 'UF']) || '';
            const regiao = getColValue(techRow, ['REGIÃO', 'Regiao']) || '';

            let parts = [];
            if (cidadeBase) parts.push(`Base: <strong>${cidadeBase}${ufBase ? ` - ${ufBase}` : ''}</strong>`);
            if (regiao) parts.push(`Região: <strong>${regiao}</strong>`);

            if (parts.length > 0) {
                baseRegiaoHtml = `<p style="margin-top: 0.5rem; font-size: 0.95rem; color: var(--text-secondary);">${parts.join(' | ')}</p>`;
            }
        }

        cidadesHeader.innerHTML = `
            <h3>Técnico ${nomeTecnico}</h3>
            ${telefone ? `<div><span style="font-size: 0.85rem; background: #e2e8f0; color: #475569; padding: 2px 8px; border-radius: 4px; display: inline-flex; align-items: center; gap: 4px;"><svg xmlns="http://www.w3.org/2000/svg" width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M22 16.92v3a2 2 0 0 1-2.18 2 19.79 19.79 0 0 1-8.63-3.07 19.5 19.5 0 0 1-6-6 19.79 19.79 0 0 1-3.07-8.67A2 2 0 0 1 4.11 2h3a2 2 0 0 1 2 1.72 12.84 12.84 0 0 0 .7 2.81 2 2 0 0 1-.45 2.11L8.09 9.91a16 16 0 0 0 6 6l1.27-1.27a2 2 0 0 1 2.11-.45 12.84 12.84 0 0 0 2.81.7A2 2 0 0 1 22 16.92z"></path></svg> ${telefone}</span></div>` : ''}
            ${baseRegiaoHtml}
            <p style="margin-top: 0.25rem;">Atende a <strong>${matches.length}</strong> cidades na região.</p>
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
        const recUfName = getColValue(rec, ['UF ATENDIDA', 'UF Atendida', 'UF', 'Estado', 'Uf']) || 'N/A';

        // Usuário procurou pela cidade
        cidadesHeader.innerHTML = `
            <h3>Cidade encontrada: ${recCityName} - ${recUfName}</h3>
            <p>Esta cidade é atendida por <strong>${matches.length}</strong> técnico(s):</p>
        `;

        const recTech = getColValue(rec, ['TÉCNICO', 'Técnico', 'Tecnico']) || 'N/A';
        const recPhone = getTechPhone(recTech);

        let html = `
            <div class="cidade-card" style="grid-column: 1 / -1; border-color: #22c55e; border-width: 2px; background: #f0fdf4;">
                <div style="display: flex; align-items: center; justify-content: space-between; width: 100%;">
                    <div>
                        <div style="color: #16a34a; font-size: 0.75rem; font-weight: 800; margin-bottom: 0.25rem;">✨ TÉCNICO MAIS PRÓXIMO</div>
                        <span class="cidade-name" style="display:inline-block; margin-bottom:0.25rem; font-size: 1.1rem; color: #166534; font-weight: bold;">${recTech}</span>
                        ${recPhone ? `<span style="font-size: 0.8rem; background: #bbf7d0; color: #166534; padding: 2px 6px; border-radius: 4px; margin-left: 8px;"><svg xmlns="http://www.w3.org/2000/svg" width="10" height="10" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" style="margin-right:2px; vertical-align:middle;"><path d="M22 16.92v3a2 2 0 0 1-2.18 2 19.79 19.79 0 0 1-8.63-3.07 19.5 19.5 0 0 1-6-6 19.79 19.79 0 0 1-3.07-8.67A2 2 0 0 1 4.11 2h3a2 2 0 0 1 2 1.72 12.84 12.84 0 0 0 .7 2.81 2 2 0 0 1-.45 2.11L8.09 9.91a16 16 0 0 0 6 6l1.27-1.27a2 2 0 0 1 2.11-.45 12.84 12.84 0 0 0 2.81.7A2 2 0 0 1 22 16.92z"></path></svg>${recPhone}</span>` : ''}
                        <div style="font-size: 0.85rem; color: #15803d; margin-top: 4px;">De: ${getColValue(rec, ['CIDADE BASE', 'Cidade Base']) || 'N/A'} <strong>(${getColValue(rec, ['DISTÂNCIA (KM)', 'Distância']) || '0'}km)</strong></div>
                    </div>
                    <span class="cidade-uf" style="background: #bbf7d0; color: #166534;">${getColValue(rec, ['UF ATENDIDA', 'UF Atendida', 'UF', 'Estado', 'Uf']) || 'N/A'}</span>
                </div>
            </div>
        `;

        // Se houver mais técnicos
        if (others.length > 0) {
            html += `<h4 style="grid-column: 1 / -1; margin-top: 1rem; color: var(--text-secondary); font-size: 0.9rem;">Outras opções de técnicos:</h4>`;
            html += others.map(m => {
                const techName = getColValue(m, ['TÉCNICO', 'Técnico', 'Tecnico']) || 'N/A';
                const techPhone = getTechPhone(techName);
                return `
                <div class="cidade-card">
                    <div>
                        <span class="cidade-name" style="display:inline-block; margin-bottom:0.25rem; font-weight: bold;">${techName}</span>
                        ${techPhone ? `<div style="font-size: 0.75rem; color: #475569; margin-bottom: 2px;"><svg xmlns="http://www.w3.org/2000/svg" width="10" height="10" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" style="margin-right:2px; vertical-align:middle;"><path d="M22 16.92v3a2 2 0 0 1-2.18 2 19.79 19.79 0 0 1-8.63-3.07 19.5 19.5 0 0 1-6-6 19.79 19.79 0 0 1-3.07-8.67A2 2 0 0 1 4.11 2h3a2 2 0 0 1 2 1.72 12.84 12.84 0 0 0 .7 2.81 2 2 0 0 1-.45 2.11L8.09 9.91a16 16 0 0 0 6 6l1.27-1.27a2 2 0 0 1 2.11-.45 12.84 12.84 0 0 0 2.81.7A2 2 0 0 1 22 16.92z"></path></svg>${techPhone}</div>` : ''}
                        <div style="font-size: 0.75rem; color: #64748b;">De: ${getColValue(m, ['CIDADE BASE', 'Cidade Base']) || 'N/A'} (${getColValue(m, ['DISTÂNCIA (KM)', 'Distância']) || '0'}km)</div>
                    </div>
                    <span class="cidade-uf">${getColValue(m, ['UF ATENDIDA', 'UF Atendida', 'UF', 'Estado', 'Uf']) || 'N/A'}</span>
                </div>
            `;
            }).join('');
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

    // Popular Tabela
    const resultsTableBody = document.querySelector('#results-table tbody');
    resultsTableBody.innerHTML = '';

    // Configuração de Ordenação Global para a tabela atual
    window.currentTableSort = { column: 'prazo', asc: true };

    const renderTableRows = (sortedOptions) => {
        resultsTableBody.innerHTML = '';
        sortedOptions.forEach(opt => {
            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td style="font-weight: 500;">
                    ${formatProviderName(opt.provider)}
                    ${opt.provider === bestOption.provider ? ' <span title="Recomendado" style="font-size:1.1rem">✨</span>' : ''}
                </td>
                <td>
                    <span class="time-badge">${opt.days} ${opt.days === 1 ? 'dia' : 'dias'}</span>
                </td>
            `;
            resultsTableBody.appendChild(tr);
        });
    };

    // Render inicial
    // Default sort by Prazo (asc)
    const sortedInit = [...options].sort((a, b) => a.days - b.days);
    renderTableRows(sortedInit);

    // Adiciona listener de clique nos headers da tabela para ordenação
    const theadTr = document.querySelector('#results-table thead tr');
    theadTr.innerHTML = `
        <th data-sort="provider" style="cursor: pointer; user-select: none;" title="Ordenar por Transportadora">Transportadora <span class="sort-icon"></span></th>
        <th data-sort="days" style="cursor: pointer; user-select: none;" title="Ordenar por Prazo">Prazo <span class="sort-icon">🔼</span></th>
    `;

    theadTr.querySelectorAll('th').forEach(th => {
        th.addEventListener('click', () => {
            const column = th.getAttribute('data-sort');
            if (window.currentTableSort.column === column) {
                window.currentTableSort.asc = !window.currentTableSort.asc;
            } else {
                window.currentTableSort.column = column;
                window.currentTableSort.asc = true; // default asc on new column
            }

            // Atualiza ícones
            theadTr.querySelectorAll('.sort-icon').forEach(icon => icon.innerHTML = '');
            th.querySelector('.sort-icon').innerHTML = window.currentTableSort.asc ? '🔼' : '🔽';

            // Ordena os dados
            const newSorted = [...options].sort((a, b) => {
                if (column === 'days') {
                    return window.currentTableSort.asc ? a.days - b.days : b.days - a.days;
                } else { // column is 'provider'
                    return window.currentTableSort.asc ? a.provider.localeCompare(b.provider) : b.provider.localeCompare(a.provider);
                }
            });

            renderTableRows(newSorted);
        });
    });
}


