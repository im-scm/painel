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

// ================= UTIL =================
function parseNumber(v) {
    return typeof v === 'number' ? v : parseFloat(v) || 0;
}

// ✅ CORREÇÃO DATA (CRÍTICO)
function parseDateBR(value) {
    if (typeof value === 'number') {
        const utc_days = Math.floor(value - 25569);
        const date = new Date(utc_days * 86400 * 1000);
        return new Date(date.getFullYear(), date.getMonth(), date.getDate());
    }

    if (typeof value === 'string') {
        const parts = value.split('/');
        if (parts.length === 3) {
            return new Date(parts[2], parts[1] - 1, parts[0]);
        }
    }

    return null;
}

// ================= PROCESS =================
function processData(data) {
    return data.map(r => {
        Object.keys(r).forEach(k => {
            if (k !== 'Data') r[k] = parseNumber(r[k]);
        });
        r.Data = parseDateBR(r.Data);
        return r;
    }).sort((a, b) => a.Data - b.Data);
}

// ================= LOAD AUTO =================
function loadDatabaseFile() {
    fetch('./app_scm_data.xlsx')
        .then(r => {
            if (!r.ok) throw new Error();
            return r.arrayBuffer();
        })
        .then(data => {
            const wb = XLSX.read(data, { type: 'array' });
            const ws = wb.Sheets['Final'] || wb.Sheets[wb.SheetNames[0]];
            const json = XLSX.utils.sheet_to_json(ws);

            globalData = processData(json);
            filteredData = [...globalData];

            console.log("✅ AUTO LOAD:", globalData.length);

            updateAll();
        })
        .catch(e => {
            console.warn("⚠️ Auto load falhou", e);
        });
}

// ================= UPLOAD =================
function handleFileSelect(e) {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = e => {
        const wb = XLSX.read(new Uint8Array(e.target.result), { type: 'array' });
        const ws = wb.Sheets['Final'] || wb.Sheets[wb.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json(ws);

        globalData = processData(json);
        filteredData = [...globalData];

        console.log("✅ UPLOAD:", globalData.length);

        updateAll();
    };
    reader.readAsArrayBuffer(file);
}

// ✅ FALTAVAM ESSAS FUNÇÕES
function dropHandler(e) {
    e.preventDefault();
    handleFileSelect({ target: { files: e.dataTransfer.files } });
}

function dragOverHandler(e) { e.preventDefault(); }
function dragEnterHandler(e) { e.preventDefault(); }
function dragLeaveHandler(e) { e.preventDefault(); }

// ================= KPIs =================
function updateKPIs() {
    const d = filteredData.at(-1);
    if (!d) return;

    setKPI('celuloseKPI', [
        { label: 'IM', val: d.BHKP_IM },
        { label: 'EU', val: d.BHKP_EU }
    ]);

    setKPI('celuloseCNKPI', [{ label: 'CN', val: d.BHKP_CN }]);

    setKPI('tio2KPI', [
        { label: 'IM', val: d.TIO2_IM },
        { label: 'CN', val: d.TIO2_CN }
    ]);

    setKPI('resinasKPI', [
        { label: 'UF', val: d.RES_UF },
        { label: 'MF', val: d.RES_MF }
    ]);

    setKPI('freteImportKPI', [
        { label: 'EU', val: d.CNT_EU_EUR },
        { label: 'CN', val: d.CNT_CN_USD }
    ]);

    setKPI('freteExportKPI', [
        { label: 'GQ', val: d.CNT_GQ_USD },
        { label: 'CG', val: d.CNT_CG_USD },
        { label: 'VC', val: d.CNT_VC_USD }
    ]);
}

function setKPI(id, arr) {
    const el = document.getElementById(id);
    if (!el) return;

    el.innerHTML = arr.map(i => `
        <div class="kpi-item">
            <span>${i.label}</span>
            <strong>${i.val?.toFixed(0) || '--'}</strong>
        </div>
    `).join('');
}

// ================= CHART =================
function createChart(id, cfg) {
    const ctx = document.getElementById(id);
    if (!ctx) return;

    if (charts[id]) charts[id].destroy();

    charts[id] = new Chart(ctx, {
        type: 'line',
        data: {
            labels: filteredData.map(d => d.Data.toLocaleDateString()),
            datasets: cfg.map(s => ({
                label: s.label,
                data: filteredData.map(d => d[s.field]),
                borderColor: s.color,
                tension: 0.3
            }))
        }
    });
}

function updateCharts() {
    Object.keys(chartSeriesConfig).forEach(k => {
        createChart(k + 'Chart', chartSeriesConfig[k]);
    });
}

// ================= UPDATE =================
function updateAll() {
    updateCharts();
    updateKPIs();
}

// ================= INIT =================
document.addEventListener('DOMContentLoaded', () => {
    console.log("XLSX disponível?", typeof XLSX);
    loadDatabaseFile();
});

window.handleFileSelect = handleFileSelect;
window.dropHandler = dropHandler;
window.dragOverHandler = dragOverHandler;
