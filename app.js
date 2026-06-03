// ================= GLOBAL =================
let globalData = [];
let filteredData = [];
let charts = {};

// ================= CONFIG =================
const chartSeriesConfig = {
    celulose: [
        { field: 'BHKP_IM', label: 'IM', color: '#CD853F' },
        { field: 'BHKP_EU', label: 'EU', color: '#6B8E23' },
        { field: 'BHKP_CN', label: 'CN', color: '#708090' }
    ],
    tio2: [
        { field: 'TIO2_IM', label: 'IM', color: '#4A148C' },
        { field: 'TIO2_CN', label: 'CN', color: '#CD853F' }
    ],
    metanol: [
        { field: 'MET_GPC', label: 'GPC', color: '#708090' },
        { field: 'MET_MN', label: 'MN', color: '#CD853F' }
    ],
    ureia: [
        { field: 'URE_GPC', label: 'GPC', color: '#6B8E23' },
        { field: 'URE_ME', label: 'ME', color: '#CD853F' }
    ],
    melamina: [
        { field: 'MEL_GPC', label: 'GPC', color: '#CD853F' },
        { field: 'MEL_CN', label: 'CN', color: '#708090' }
    ],
    resinas: [
        { field: 'RES_UF', label: 'UF', color: '#6B8E23' },
        { field: 'RES_MF', label: 'MF', color: '#CD853F' },
        { field: 'USDBRL_GPC', label: 'USD', color: '#708090' }
    ],
    moedas: [
        { field: 'USDBRL', label: 'USD', color: '#6B8E23' },
        { field: 'EURBRL', label: 'EUR', color: '#708090' },
        { field: 'CNYBRL', label: 'CNY', color: '#CD853F' }
    ],
    freteimport: [
        { field: 'CNT_EU_EUR', label: 'EU', color: '#6B8E23' },
        { field: 'CNT_CN_USD', label: 'CN', color: '#CD853F' }
    ],
    freteexport: [
        { field: 'CNT_GQ_USD', label: 'GQ', color: '#6B8E23' },
        { field: 'CNT_CG_USD', label: 'CG', color: '#8B4513' },
        { field: 'CNT_VC_USD', label: 'VC', color: '#CD853F' }
    ]
};

// ================= FIX DATA (CRÍTICO) =================
function parseDateBR(value) {

    // Excel number → manter UTC
    if (typeof value === 'number') {
        const utc_days = Math.floor(value - 25569);
        return new Date(Date.UTC(1970, 0, utc_days));
    }

    if (typeof value === 'string') {
        const p = value.split('/');
        if (p.length === 3) {
            return new Date(Date.UTC(p[2], p[1] - 1, p[0]));
        }
    }

    return null;
}

// ================= PROCESS =================
function processData(data) {
    return data.map(r => {
        Object.keys(r).forEach(k => {
            if (k !== 'Data') r[k] = parseFloat(r[k]) || 0;
        });

        r.Data = parseDateBR(r.Data);
        return r;

    }).filter(r => r.Data)
      .sort((a, b) => a.Data - b.Data);
}

function updateDateInputs() {

    if (!filteredData.length) return;

    const first = filteredData[0].Data;
    const last = filteredData[filteredData.length - 1].Data;

    const format = (date) => {
        const d = String(date.getUTCDate()).padStart(2, '0');
        const m = String(date.getUTCMonth() + 1).padStart(2, '0');
        const y = date.getUTCFullYear();
        return `${d}/${m}/${y}`;
    };

    const startInput = document.getElementById('startDate');
    const endInput = document.getElementById('endDate');

    if (startInput) startInput.value = format(first);
    if (endInput) endInput.value = format(last);
}

// ================= LOAD =================
function loadDatabaseFile() {
    fetch('./app_scm_data.xlsx')
        .then(r => r.arrayBuffer())
        .then(data => {

            const wb = XLSX.read(data, { type: 'array' });
            const ws = wb.Sheets['Final'] || wb.Sheets[wb.SheetNames[0]];
            const json = XLSX.utils.sheet_to_json(ws);

            globalData = processData(json);
            filteredData = [...globalData];

            console.log("✅ AUTO:", globalData.length);
            
            updateDateInputs();

            updateAll();
        });
}

// ================= KPI =================
function updateKPIs() {
    const d = filteredData.at(-1);
    if (!d) return;

    const safe = v => (v && !isNaN(v)) ? v : 0;

    setKPI('celuloseKPI', [
        { label: 'IM', val: safe(d.BHKP_IM) },
        { label: 'EU', val: safe(d.BHKP_EU) }
    ]);

    setKPI('celuloseCNKPI', [
        { label: 'CN', val: safe(d.BHKP_CN) }
    ]);

    setKPI('tio2KPI', [
        { label: 'IM', val: safe(d.TIO2_IM) },
        { label: 'CN', val: safe(d.TIO2_CN) }
    ]);

    setKPI('resinasKPI', [
        { label: 'UF', val: safe(d.RES_UF) },
        { label: 'MF', val: safe(d.RES_MF) }
    ]);

    setKPI('freteImportKPI', [
        { label: 'EU', val: safe(d.CNT_EU_EUR) },
        { label: 'CN', val: safe(d.CNT_CN_USD) }
    ]);

    setKPI('freteExportKPI', [
        { label: 'GQ', val: safe(d.CNT_GQ_USD) },
        { label: 'CG', val: safe(d.CNT_CG_USD) },
        { label: 'VC', val: safe(d.CNT_VC_USD) }
    ]);
}

function setKPI(id, arr) {
    const el = document.getElementById(id);

    el.innerHTML = arr.map(i => `
        <div class="kpi-item">
            <span>${i.label}</span>
            <span class="value">${i.val.toLocaleString('pt-BR')}</span>
        </div>
    `).join('');
}

// ================= CHART =================
function createChart(id, config) {

    const ctx = document.getElementById(id);

    if (charts[id]) charts[id].destroy();

    charts[id] = new Chart(ctx, {
        type: 'line',
        data: {
            labels: filteredData.map(d => formatDateBR(d.Data)),
            datasets: config.map(s => ({
                label: s.label,
                data: filteredData.map(d => d[s.field]),
                borderColor: s.color,
                tension: 0.3
            }))
        },
        options: {
            responsive: true,
            maintainAspectRatio: false, // ✅ ESSENCIAL
        }
    });
}

function updateCharts() {
    Object.keys(chartSeriesConfig).forEach(k => {
        createChart(k + 'Chart', chartSeriesConfig[k]);
    });
}

// ================= ALL =================
function updateAll() {
    updateCharts();
    updateKPIs();
    updateAllMetrics(); // ✅ VOLTOU
}

function calculateChange(oldVal, newVal) {
    if (!oldVal || oldVal === 0) return 0;
    return ((newVal - oldVal) / oldVal) * 100;
}

function getYearStartData(data) {
    if (!data.length) return [];
    const year = data[data.length - 1].Data.getUTCFullYear();
    return data.filter(d => d.Data.getUTCFullYear() === year);
}

function updateChartMetrics(chartType) {

    const container = document.getElementById(chartType + "Metrics");
    if (!container) return;

    const series = chartSeriesConfig[chartType];
    if (!series) return;

    const data = filteredData;
    if (data.length < 2) return;

    const last = data[data.length - 1];
    const prev = data[data.length - 2];

    const ytdData = getYearStartData(data);

    let html = '';

    // HEADER (nomes das séries)
    html += '<div class="metrics-header">';
    html += '<div></div>';
    series.forEach(s => {
        html += `<div class="series-header">${s.label}</div>`;
    });
    html += '</div>';

    // TIPOS DE MÉTRICA
    const metrics = [
        { key: 'mom', label: 'MOM' },
        { key: 'ytd', label: 'YTD' },
        { key: 'all', label: 'ALL' }
    ];

    metrics.forEach(metric => {

        html += '<div class="metric-row">';
        html += `<div class="metric-row-label">${metric.label}</div>`;

        series.forEach(s => {

            let value = 0;

            if (metric.key === 'mom') {
                value = calculateChange(prev[s.field], last[s.field]);
            }

            if (metric.key === 'ytd' && ytdData.length > 1) {
                value = calculateChange(
                    ytdData[0][s.field],
                    ytdData[ytdData.length - 1][s.field]
                );
            }

            if (metric.key === 'all') {
                value = calculateChange(
                    data[0][s.field],
                    last[s.field]
                );
            }

            const formatted = `${value > 0 ? '+' : ''}${value.toFixed(1)}%`;

            html += `
                <div class="metric-cell">
                    <span class="metric-value">
                        ${formatted}
                    </span>
                </div>
            `;
        });

        html += '</div>';
    });

    container.innerHTML = html;
}

function updateAllMetrics() {
    Object.keys(chartSeriesConfig).forEach(chart => {
        updateChartMetrics(chart);
    });
}

function formatDateBR(date) {
    const d = String(date.getUTCDate()).padStart(2, '0');
    const m = String(date.getUTCMonth() + 1).padStart(2, '0');
    const y = date.getUTCFullYear();
    return `${d}/${m}/${y}`;
}


// ================= INIT =================
document.addEventListener('DOMContentLoaded', () => {
    loadDatabaseFile();
});
``
