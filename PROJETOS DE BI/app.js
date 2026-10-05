let fullData = [];

// Event Listeners
document.getElementById('excelFile').addEventListener('change', handleFile, false);
document.getElementById('searchInput').addEventListener('input', filterTable, false);

// Controle de Navegação da Sidebar
const navItems = document.querySelectorAll('.nav-item');
const dashboardViews = document.querySelectorAll('.dashboard-view');

navItems.forEach(item => {
    item.addEventListener('click', () => {
        navItems.forEach(i => i.classList.remove('active'));
        item.classList.add('active');

        const targetId = item.getAttribute('data-target');
        dashboardViews.forEach(view => {
            if (view.id === targetId) {
                view.classList.add('active-view');
            } else {
                view.classList.remove('active-view');
            }
        });
    });
});

function handleFile(e) {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function(event) {
        try {
            const data = new Uint8Array(event.target.result);
            const workbook = XLSX.read(data, { 
                type: 'array',
                cellDates: true,
                dateNF: 'dd/mm/yyyy hh:mm:ss'
            });

            let worksheet = null;
            for (let name of workbook.SheetNames) {
                const sheet = workbook.Sheets[name];
                if (sheet && sheet['!ref']) {
                    worksheet = sheet;
                    break;
                }
            }

            if (!worksheet) {
                alert("Nenhuma aba com dados foi encontrada no arquivo.");
                return;
            }

            fullData = XLSX.utils.sheet_to_json(worksheet, { defval: "" });

            if (fullData.length === 0) {
                alert("O arquivo foi lido, mas a tabela está vazia.");
                return;
            }

            const searchInput = document.getElementById('searchInput');
            searchInput.disabled = false;
            searchInput.value = '';

            renderDashboard(fullData);
        } catch (err) {
            console.error("Erro ao ler o arquivo:", err);
            alert("Erro ao ler o arquivo Excel. Verifique o console para mais detalhes.");
        }
    };
    reader.readAsArrayBuffer(file);
}

function renderDashboard(data) {
    const tableBody = document.getElementById('tableBody');
    const statusSummaryContainer = document.getElementById('status-summary-container');
    
    tableBody.innerHTML = '';
    statusSummaryContainer.innerHTML = '';

    if (data.length === 0) {
        tableBody.innerHTML = '<tr><td colspan="7" class="empty-msg">Nenhum registro encontrado.</td></tr>';
        document.getElementById('kpi-abertos').innerText = 0;
        document.getElementById('kpi-dentro-sla').innerText = 0;
        document.getElementById('kpi-fora-sla').innerText = 0;
        document.getElementById('kpi-dentro-pct').innerText = '(0%)';
        document.getElementById('kpi-fora-pct').innerText = '(0%)';
        return;
    }

    // --- CLASSIFICAÇÃO POR AGEING (MAIOR PARA O MENOR) ---
    const sortedData = [...data].sort((a, b) => {
        const getAgeing = (row) => {
            const foundKey = Object.keys(row).find(k => k.trim().toUpperCase() === 'AGEING');
            if (!foundKey) return 0;
            const val = parseFloat(String(row[foundKey]).replace(',', '.').replace(/[^0-9.-]/g, ''));
            return isNaN(val) ? 0 : val;
        };

        const ageingA = getAgeing(a);
        const ageingB = getAgeing(b);

        return ageingB - ageingA;
    });

    let totalAbertos = 0;
    let dentroSLA = 0;
    let foraSLA = 0;
    const statusCounts = {}; // Mapeador de contagem por status

    sortedData.forEach(row => {
        const getValue = (keyName) => {
            const foundKey = Object.keys(row).find(k => 
                k.trim().toUpperCase() === keyName.trim().toUpperCase()
            );
            return foundKey ? row[foundKey] : '-';
        };

        const dataPed = getValue('DATA');
        const pedido = getValue('PEDIDO');
        const franquia = getValue('FRANQUIA');
        const status = getValue('STATUS');
        const limiteSLA = getValue('LIMITE SLA');
        const ageing = getValue('AGEING');
        const statusSLA = getValue('STATUS SLA');

        if (pedido === '-' && status === '-') return;

        totalAbertos++;

        // Contabiliza ocorrências por status do pedido
        const statusKey = String(status).trim();
        statusCounts[statusKey] = (statusCounts[statusKey] || 0) + 1;

        const slaStr = String(statusSLA).toUpperCase();
        const isVencido = slaStr.includes('VENCIDO') || slaStr.includes('FORA') || slaStr.includes('ATRASADO');

        if (isVencido) {
            foraSLA++;
        } else if (slaStr.includes('DENTRO') || slaStr.includes('OK')) {
            dentroSLA++;
        }

        const badgeClass = isVencido ? 'badge-late' : 'badge-ok';

        const dataFormatada = (dataPed instanceof Date) 
            ? dataPed.toLocaleString('pt-BR') 
            : dataPed;

        const limiteFormatado = (limiteSLA instanceof Date) 
            ? limiteSLA.toLocaleString('pt-BR') 
            : limiteSLA;

        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td>${dataFormatada}</td>
            <td><strong>${pedido}</strong></td>
            <td>${franquia}</td>
            <td>${status}</td>
            <td>${limiteFormatado}</td>
            <td><span class="badge ${badgeClass}">${statusSLA}</span></td>
            <td><strong>${ageing}</strong></td>
        `;
        tableBody.appendChild(tr);
    });

    // --- MONTAGEM DO RESUMO DE STATUS DO PEDIDO ---
    Object.keys(statusCounts).forEach(statusName => {
        const chip = document.createElement('div');
        chip.className = 'status-chip';
        chip.innerHTML = `
            <span class="status-chip-label">${statusName}</span>
            <span class="status-chip-count">${statusCounts[statusName]}</span>
        `;
        statusSummaryContainer.appendChild(chip);
    });

    // --- CÁLCULO DOS PERCENTUAIS DOS KPIs ---
    const pctDentro = totalAbertos > 0 ? ((dentroSLA / totalAbertos) * 100).toFixed(1) : 0;
    const pctFora = totalAbertos > 0 ? ((foraSLA / totalAbertos) * 100).toFixed(1) : 0;

    document.getElementById('kpi-abertos').innerText = totalAbertos;
    document.getElementById('kpi-dentro-sla').innerText = dentroSLA;
    document.getElementById('kpi-fora-sla').innerText = foraSLA;
    document.getElementById('kpi-dentro-pct').innerText = `(${pctDentro}%)`;
    document.getElementById('kpi-fora-pct').innerText = `(${pctFora}%)`;
}

function filterTable() {
    const searchTerm = document.getElementById('searchInput').value.toLowerCase().trim();

    const filteredData = fullData.filter(row => {
        const getValue = (keyName) => {
            const foundKey = Object.keys(row).find(k => 
                k.trim().toUpperCase() === keyName.trim().toUpperCase()
            );
            return foundKey ? String(row[foundKey]).toLowerCase() : '';
        };

        const pedido = getValue('PEDIDO');
        const franquia = getValue('FRANQUIA');
        const status = getValue('STATUS');
        const statusSLA = getValue('STATUS SLA');

        return pedido.includes(searchTerm) || franquia.includes(searchTerm) || statusSLA.includes(searchTerm) || status.includes(searchTerm);
    });

    renderDashboard(filteredData);
}