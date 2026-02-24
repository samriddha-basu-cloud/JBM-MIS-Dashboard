import { initializeApp } from 'https://www.gstatic.com/firebasejs/10.7.1/firebase-app.js';
import { getFirestore, collection, getDocs, query, orderBy } from 'https://www.gstatic.com/firebasejs/10.7.1/firebase-firestore.js';

const firebaseConfig = {
    apiKey: "AIzaSyDTczWACmsAcYKY743wLqNe0DiRAy5k6tk",
    authDomain: "mis-dashboard-f1688.firebaseapp.com",
    projectId: "mis-dashboard-f1688",
    storageBucket: "mis-dashboard-f1688.firebasestorage.app",
    messagingSenderId: "476511626207",
    appId: "1:476511626207:web:d3eb5f41cbec14b7325a6a"
};

const app = initializeApp(firebaseConfig);
const db  = getFirestore(app);

let allData      = [];
let filteredData = [];
let charts       = {};

// ── Neon palette ───────────────────────────────────────────────────────────
const C = {
    cyan:      '#00e5ff', cyanDim:  'rgba(0,229,255,0.13)', cyanMid:  'rgba(0,229,255,0.38)',
    green:     '#00e676', greenDim: 'rgba(0,230,118,0.13)', greenMid: 'rgba(0,230,118,0.38)',
    amber:     '#ffab00', amberDim: 'rgba(255,171,0,0.13)', amberMid: 'rgba(255,171,0,0.38)',
    red:       '#ff1744', redDim:   'rgba(255,23,68,0.13)', redMid:   'rgba(255,23,68,0.38)',
    violet:    '#d500f9', violetDim:'rgba(213,0,249,0.13)',
    blue:      '#2979ff', orange:   '#ff6d00', teal:    '#1de9b6',
    grid:      'rgba(255,255,255,0.06)',
    tick:      'rgba(255,255,255,0.5)',
    legend:    'rgba(255,255,255,0.7)',
    tooltip:   'rgba(8,12,24,0.96)',
};

const PALETTE = [C.cyan,C.green,C.amber,C.red,C.violet,C.blue,C.orange,C.teal,'#ff80ab','#80d8ff','#ccff90','#ffd180'];

const darkOpts = {
    plugins: {
        legend: { labels: { color: C.legend, usePointStyle: true, padding: 14,
            font: { size: 11, family: "'JetBrains Mono', monospace" } } },
        tooltip: { backgroundColor: C.tooltip, borderColor: 'rgba(0,229,255,0.25)', borderWidth: 1,
            titleColor: C.cyan, bodyColor: '#c8d6e5', padding: 12, cornerRadius: 8,
            titleFont: { family:"'JetBrains Mono',monospace", size:11 },
            bodyFont:  { family:"'JetBrains Mono',monospace", size:11 } }
    },
    scales: {
        x: { ticks: { color: C.tick, font: { size:10, family:"'JetBrains Mono',monospace" }, maxRotation:45 }, grid: { color: C.grid } },
        y: { ticks: { color: C.tick, font: { size:10, family:"'JetBrains Mono',monospace" } }, grid: { color: C.grid }, beginAtZero: true }
    }
};

function dMerge(t, s) {
    const o = { ...t };
    for (const k in s) {
        if (s[k] && typeof s[k] === 'object' && !Array.isArray(s[k])) o[k] = dMerge(t[k]||{}, s[k]);
        else o[k] = s[k];
    }
    return o;
}
const mo = (...args) => args.reduce((a, b) => dMerge(a, b), {});

// Safe number parse — handles strings, Firestore numbers, undefined, null, NaN
function n(v) { const x = parseFloat(v); return isFinite(x) ? x : 0; }

// ── Data loading ──────────────────────────────────────────────────────────
async function loadData() {
    try {
        const today = new Date();
        const dateStrings = [];
        for (let i = 0; i < 365; i++) {
            const d = new Date(today);
            d.setDate(today.getDate() - i);
            dateStrings.push(
                `${d.getUTCFullYear()}-${String(d.getUTCMonth()+1).padStart(2,'0')}-${String(d.getUTCDate()).padStart(2,'0')}`
            );
        }

        allData = [];

        const fetches = dateStrings.map(async (dateStr) => {
            try {
                const ref      = collection(db, 'daily_entries', dateStr, 'data');
                const snapshot = await getDocs(query(ref, orderBy('__name__')));
                if (snapshot.empty) return;
                snapshot.forEach(docSnap => {
                    const d = docSnap.data();
                    allData.push({
                        id: docSnap.id,
                        ...d,
                        dateString: dateStr,
                        date: parseDateUTC(dateStr)
                    });
                });
            } catch { /* day has no data — skip */ }
        });

        await Promise.all(fetches);
        allData.sort((a, b) => b.dateString.localeCompare(a.dateString));
        filteredData = [...allData];
        console.log(`Loaded ${allData.length} records from ${new Set(allData.map(d=>d.dateString)).size} dates`);

    } catch (error) {
        console.error('Error loading data:', error);
        showToast('Error loading data: ' + error.message, 'danger');
    }
}

function parseDateUTC(str) {
    if (!str) return null;
    const [y, m, d] = str.split('-').map(Number);
    return new Date(Date.UTC(y, m - 1, d));
}

// ── Init ──────────────────────────────────────────────────────────────────
document.addEventListener('DOMContentLoaded', async () => {
    showLoading();
    try {
        await loadData();
        populateFilters();
        setDefaultDates();
        updateDashboard();
        initCharts();
    } catch(e) {
        console.error('Init error:', e);
        showToast('Initialisation error: ' + e.message, 'danger');
    } finally {
        hideLoading();
    }

    document.getElementById('btnApplyFilter')?.addEventListener('click', applyFilters);
    document.getElementById('btnExport')?.addEventListener('click', exportToExcel);
    document.getElementById('btnRefreshDashboard')?.addEventListener('click', async () => {
        showLoading();
        try { await loadData(); applyFilters(); } finally { hideLoading(); }
    });
    document.getElementById('btnExportPDF')?.addEventListener('click', exportToPDF);

    document.getElementById('prodTrendLine')?.addEventListener('change', () => {
        if (charts.dailyProduction) { charts.dailyProduction.config.type = 'line'; charts.dailyProduction.update(); }
    });
    document.getElementById('prodTrendBar')?.addEventListener('change', () => {
        if (charts.dailyProduction) { charts.dailyProduction.config.type = 'bar'; charts.dailyProduction.update(); }
    });
});

function setDefaultDates() {
    const today = new Date();
    const from  = new Date(today);
    from.setDate(today.getDate() - 30);
    const fmt = d => `${d.getUTCFullYear()}-${String(d.getUTCMonth()+1).padStart(2,'0')}-${String(d.getUTCDate()).padStart(2,'0')}`;
    const elFrom = document.getElementById('filterFromDate');
    const elTo   = document.getElementById('filterToDate');
    if (elFrom) elFrom.value = fmt(from);
    if (elTo)   elTo.value   = fmt(today);
}

function populateFilters() {
    const months = ['January','February','March','April','May','June',
                    'July','August','September','October','November','December'];
    const monthSel = document.getElementById('filterMonth');
    if (monthSel) {
        months.forEach((m, i) => {
            const o = document.createElement('option');
            o.value = i + 1; o.textContent = m;
            monthSel.appendChild(o);
        });
    }
    const years = [...new Set(allData.map(d => d.date?.getUTCFullYear()).filter(Boolean))];
    const yearSel = document.getElementById('filterYear');
    if (yearSel) {
        years.sort((a,b)=>b-a).forEach(y => {
            const o = document.createElement('option');
            o.value = y; o.textContent = y;
            yearSel.appendChild(o);
        });
    }
    const models = [...new Set(allData.map(d => d.model).filter(Boolean))].sort();
    const modelSel = document.getElementById('filterModel');
    if (modelSel) {
        models.forEach(m => {
            const o = document.createElement('option');
            o.value = m; o.textContent = m;
            modelSel.appendChild(o);
        });
    }
}

function applyFilters() {
    const month    = document.getElementById('filterMonth')?.value;
    const year     = document.getElementById('filterYear')?.value;
    const model    = document.getElementById('filterModel')?.value;
    const fromDate = document.getElementById('filterFromDate')?.value;
    const toDate   = document.getElementById('filterToDate')?.value;

    filteredData = allData.filter(item => {
        const d = item.date;
        if (fromDate && d && d < new Date(fromDate)) return false;
        if (toDate   && d && d > new Date(toDate + 'T23:59:59')) return false;
        if (month    && d && d.getUTCMonth() + 1 !== parseInt(month)) return false;
        if (year     && d && d.getUTCFullYear()   !== parseInt(year))  return false;
        if (model    && item.model !== model) return false;
        return true;
    });

    updateDashboard();
    updateCharts();
}

// ── OEE Calculation ───────────────────────────────────────────────────────
// Proper OEE = (Availability% / 100) × (Performance% / 100) × (Quality% / 100) × 100
// Falls back to stored oeePct if components are zero.
function computeOEE(item) {
    const stored = n(item.oeePct);
    const avail  = n(item.availabilityPct);
    const perf   = n(item.performancePct);
    const qual   = n(item.qualityPct);
    if (avail > 0 && perf > 0 && qual > 0) {
        return (avail / 100) * (perf / 100) * (qual / 100) * 100;
    }
    return stored;
}

// ── KPIs ──────────────────────────────────────────────────────────────────
function updateDashboard() {
    if (filteredData.length === 0) {
        ['kpiProduction','kpiOEE','kpiQuality','kpiPerformance','kpiAvailability',
         'kpiTotalPlan','kpiQualityOK','kpiRejections','kpiTotalEntries','kpiModelsCount',
         'kpiAvgLoss','kpiCapacityUtil'].forEach(id => s(id, id.includes('Pct')||id.includes('OEE')||id.includes('Quality')||id.includes('Perf')||id.includes('Avail')||id.includes('Util') ? '0%' : '0'));
        sv('kpiPlanPct','0%','badge bg-danger');
        sv('kpiRejRate','0%','badge bg-success');
        bar('kpiOEEBar',0); bar('kpiQualityBar',0); bar('kpiPerformanceBar',0);
        updateStatistics(); updateTable();
        return;
    }

    const cnt = filteredData.length;
    const tp  = filteredData.reduce((s,d) => s + n(d.totalProduction), 0);
    const tpl = filteredData.reduce((s,d) => s + n(d.todaysPlan),      0);
    const qOK = filteredData.reduce((s,d) => s + n(d.qualityOK),       0);
    const rej = filteredData.reduce((s,d) => s + n(d.rejections),      0);

    const oee  = filteredData.reduce((s,d) => s + computeOEE(d), 0) / cnt;
    const prf  = filteredData.reduce((s,d) => s + n(d.performancePct),      0) / cnt;
    const avl  = filteredData.reduce((s,d) => s + n(d.availabilityPct),     0) / cnt;
    const cap  = filteredData.reduce((s,d) => s + n(d.capacityUtilization), 0) / cnt;

    const LOSS_FIELDS = ['lossStartup','lossSetup','lossFixture','lossHR','lossPress','lossStore',
        'lossQC','lossMaintRobotProg','lossMaintFault','lossMaintShank','lossMaintClamp',
        'lossMaintLogic','lossMaintUtility','lossMaintSensor','lossMaintMig','lossMaintTucker',
        'lossMaintSSW','lossMaintSPM','lossPPC','lossMgmt'];
    const avgLoss = filteredData.reduce((s,d) => s + LOSS_FIELDS.reduce((ls,f) => ls + n(d[f]), 0), 0) / cnt;

    const qual = tp  > 0 ? (qOK / tp  * 100) : 0;
    const rr   = tp  > 0 ? (rej / tp  * 100) : 0;
    const pa   = tpl > 0 ? (tp  / tpl * 100) : 0;
    const uMdl = [...new Set(filteredData.map(d => d.model).filter(Boolean))];

    s('kpiProduction',  fmt(tp));
    sv('kpiPlanPct', pa.toFixed(1)+'%', pa>=100 ? 'badge bg-success' : pa>=90 ? 'badge bg-warning' : 'badge bg-danger');
    s('kpiOEE',         oee.toFixed(1)+'%');
    s('kpiQuality',     qual.toFixed(1)+'%');
    sv('kpiRejRate', rr.toFixed(2)+'%', rr<2 ? 'badge bg-success' : rr<5 ? 'badge bg-warning' : 'badge bg-danger');
    s('kpiPerformance', prf.toFixed(1)+'%');
    bar('kpiOEEBar', oee); bar('kpiQualityBar', qual); bar('kpiPerformanceBar', prf);
    s('kpiTotalPlan',   fmt(tpl));
    s('kpiQualityOK',   fmt(qOK));
    s('kpiRejections',  fmt(rej));
    s('kpiAvailability',avl.toFixed(1)+'%');
    s('kpiTotalEntries',fmt(filteredData.length));
    s('kpiModelsCount', uMdl.length);
    s('kpiAvgLoss',     avgLoss.toFixed(1));
    s('kpiCapacityUtil',cap.toFixed(1)+'%');

    updateStatistics();
    updateTable();
}

function s(id, v)  { const e=document.getElementById(id); if(e) e.textContent=v; }
function sv(id,v,c){ const e=document.getElementById(id); if(!e)return; e.textContent=v; if(c) e.className=c; }
function bar(id,p) { const e=document.getElementById(id); if(e) e.style.width=Math.min(100,Math.max(0,p))+'%'; }
function fmt(x)    { return new Intl.NumberFormat('en-IN').format(Math.round(n(x))||0); }
function fmtDate(d){ if(!d) return '-'; const dt=new Date(d); return `${dt.getUTCFullYear()}-${String(dt.getUTCMonth()+1).padStart(2,'0')}-${String(dt.getUTCDate()).padStart(2,'0')}`; }
function fmtN(v,dec=1){ return n(v).toFixed(dec); }

function updateStatistics() {
    if (filteredData.length === 0) return;
    const mOEE={}, mQual={}, daily={};
    filteredData.forEach(item => {
        const m = item.model || 'Unknown';
        mOEE[m]  = mOEE[m]  || { sum:0, count:0 };
        mOEE[m].sum += computeOEE(item); mOEE[m].count++;
        mQual[m] = mQual[m] || { tp:0, qOK:0 };
        mQual[m].tp  += n(item.totalProduction);
        mQual[m].qOK += n(item.qualityOK);
        const dt = item.dateString || fmtDate(item.date);
        daily[dt] = (daily[dt]||0) + n(item.totalProduction);
    });

    const keys = Object.keys(mOEE);
    if (!keys.length) return;

    const bm  = keys.reduce((a,b) => mOEE[a].sum/mOEE[a].count > mOEE[b].sum/mOEE[b].count ? a : b);
    const bqm = Object.keys(mQual).reduce((a,b) => {
        const qa = mQual[a].tp > 0 ? mQual[a].qOK/mQual[a].tp : 0;
        const qb = mQual[b].tp > 0 ? mQual[b].qOK/mQual[b].tp : 0;
        return qa > qb ? a : b;
    });
    const dKeys = Object.keys(daily);
    const bd  = dKeys.length ? dKeys.reduce((a,b) => daily[a]>daily[b] ? a : b) : '-';
    const tp  = filteredData.reduce((s,d) => s+n(d.totalProduction), 0);
    const tpl = filteredData.reduce((s,d) => s+n(d.todaysPlan), 0);
    const pa  = tpl > 0 ? tp/tpl*100 : 0;

    s('statBestModel',     bm);
    s('statBestModelOEE',  `OEE: ${(mOEE[bm].sum/mOEE[bm].count).toFixed(1)}%`);
    s('statBestQuality',   bqm);
    const q = mQual[bqm].tp > 0 ? mQual[bqm].qOK/mQual[bqm].tp*100 : 0;
    s('statBestQualityPct',`Quality: ${q.toFixed(1)}%`);
    if (bd !== '-') { s('statBestDay', bd); s('statBestDayQty', `Qty: ${fmt(daily[bd])}`); }
    s('statPlanAchievement', pa.toFixed(1)+'%');
    s('statPlanDetails',     `${fmt(tp)} / ${fmt(tpl)}`);
}

function updateTable() {
    const tbody = document.getElementById('tableBody');
    const rc    = document.getElementById('recordCount');
    if (!tbody) return;
    if (filteredData.length === 0) {
        tbody.innerHTML = `<tr><td colspan="12" class="text-center py-5" style="color:rgba(0,229,255,0.4);font-family:'Syne',sans-serif;">
            <i class="fas fa-satellite-dish fa-3x mb-3 d-block" style="opacity:.2;color:var(--neon-cyan);"></i>
            No data found. Adjust filters and click <strong style="color:var(--neon-cyan);">Apply Filter</strong>.</td></tr>`;
        if(rc) rc.textContent='0';
        return;
    }
    tbody.innerHTML = filteredData.slice(0,100).map(item => {
        const oee = computeOEE(item);
        const cls = oee>=70 ? 'status-good' : oee>=50 ? 'status-warning' : 'status-danger';
        const tp  = n(item.totalProduction);
        const rej = n(item.rejections);
        return `<tr>
            <td>${item.dateString||fmtDate(item.date)}</td>
            <td>${item.partNo||'-'}</td>
            <td>${item.model||'-'}</td>
            <td class="text-end">${fmt(item.todaysPlan)}</td>
            <td class="text-end">${fmt(tp)}</td>
            <td class="text-end">${fmt(item.qualityOK)}</td>
            <td class="text-end">${fmt(rej)}</td>
            <td class="text-center">${(tp>0 ? rej/tp*100 : 0).toFixed(2)}%</td>
            <td class="text-center ${cls}">${oee.toFixed(1)}%</td>
            <td class="text-center">${fmtN(item.performancePct)}%</td>
            <td class="text-center">${fmtN(item.qualityPct)}%</td>
            <td class="text-center">${fmtN(item.availabilityPct)}%</td>
        </tr>`;
    }).join('');
    if(rc) rc.textContent = fmt(filteredData.length);
}

// ── Charts ────────────────────────────────────────────────────────────────
function initCharts() {
    const mk = (id, type, data, extra={}) => {
        const ctx = document.getElementById(id);
        if (!ctx) return null;
        return new Chart(ctx, {
            type,
            data,
            options: mo({ responsive:true, maintainAspectRatio:false }, darkOpts, extra)
        });
    };

    charts.dailyProduction        = mk('chartDailyProduction',        'line',     getDailyProductionData());
    charts.modelBreakdown         = mk('chartModelBreakdown',         'doughnut', getModelBreakdownData(),
        { cutout:'65%', plugins:{ legend:{position:'bottom'} } });
    charts.oeeByModel             = mk('chartOEEByModel',             'bar',      getOEEByModelData(),
        { scales:{ y:{ max:100, ticks:{ callback:v=>v+'%', color:C.tick } } } });
    charts.qualityTrend           = mk('chartQualityTrend',           'line',     getQualityTrendData(),
        { scales:{ y:{ max:100, ticks:{ callback:v=>v+'%', color:C.tick } } } });
    charts.performanceDist        = mk('chartPerformanceDist',        'doughnut', getPerformanceDistributionData(),
        { cutout:'65%', plugins:{ legend:{position:'bottom'} } });
    charts.availabilityPerformance= mk('chartAvailabilityPerformance','line',     getAvailabilityPerformanceData(),
        { scales:{ y:{ max:100, ticks:{ callback:v=>v+'%', color:C.tick } } } });
    charts.rejectionByModel       = mk('chartRejectionByModel',       'bar',      getRejectionByModelData(),
        { indexAxis:'y', plugins:{ legend:{ display:false } } });
    charts.planVsActual           = mk('chartPlanVsActual',           'bar',      getPlanVsActualData());
    charts.productionHeatmap      = mk('chartProductionHeatmap',      'bar',      getProductionHeatmapData(),
        { scales:{ y:{ stacked:true, grid:{ display:false }, ticks:{color:C.tick} },
                   x:{ stacked:true, grid:{ display:false }, ticks:{color:C.tick} } },
          plugins:{ legend:{ position:'right' } } });
    charts.pareto                 = mk('chartPareto',                 'bar',      getParetoData(),
        { scales:{ y: { type:'linear', position:'left',  beginAtZero:true, grid:{color:C.grid}, ticks:{color:C.tick} },
                   y1:{ type:'linear', position:'right', beginAtZero:true, max:100,
                        grid:{ display:false }, ticks:{ callback:v=>v+'%', color:C.tick } } } });
    charts.lossCategories         = mk('chartLossCategories',         'bar',      getLossCategoriesData(),
        { indexAxis:'y', plugins:{ legend:{ display:false } } });
    charts.shiftProduction        = mk('chartShiftProduction',        'bar',      getShiftProductionData());
    charts.shopArea               = mk('chartShopArea',               'bar',      getShopAreaData(),
        { indexAxis:'y', plugins:{ legend:{ display:false } } });
    charts.processType            = mk('chartProcessType',            'doughnut', getProcessTypeData(),
        { cutout:'60%', plugins:{ legend:{ position:'bottom' } } });
}

function updateCharts() {
    const map = {
        dailyProduction: getDailyProductionData,
        modelBreakdown: getModelBreakdownData,
        oeeByModel: getOEEByModelData,
        qualityTrend: getQualityTrendData,
        performanceDist: getPerformanceDistributionData,
        availabilityPerformance: getAvailabilityPerformanceData,
        rejectionByModel: getRejectionByModelData,
        planVsActual: getPlanVsActualData,
        productionHeatmap: getProductionHeatmapData,
        pareto: getParetoData,
        lossCategories: getLossCategoriesData,
        shiftProduction: getShiftProductionData,
        shopArea: getShopAreaData,
        processType: getProcessTypeData
    };
    Object.keys(map).forEach(k => {
        if (charts[k]) { charts[k].data = map[k](); charts[k].update(); }
    });
}

// ── Chart data builders ───────────────────────────────────────────────────

function getDailyProductionData() {
    const dd = {};
    filteredData.forEach(item => {
        const dt = item.dateString || fmtDate(item.date);
        dd[dt] = dd[dt] || { production:0, plan:0 };
        dd[dt].production += n(item.totalProduction);
        dd[dt].plan       += n(item.todaysPlan);
    });
    const dates = Object.keys(dd).sort().slice(-30);
    return { labels: dates, datasets: [
        { label:'Production', data:dates.map(d=>dd[d].production), borderColor:C.cyan,  backgroundColor:C.cyanDim,       borderWidth:2.5, fill:true,  tension:0.4, pointRadius:3, pointBackgroundColor:C.cyan  },
        { label:'Plan',       data:dates.map(d=>dd[d].plan),       borderColor:C.green, backgroundColor:'transparent',   borderWidth:2,   fill:false, tension:0.4, borderDash:[6,4], pointRadius:2, pointBackgroundColor:C.green }
    ]};
}

function getModelBreakdownData() {
    const md = {};
    filteredData.forEach(item => { const m=item.model||'Unknown'; md[m]=(md[m]||0)+n(item.totalProduction); });
    return { labels:Object.keys(md), datasets:[{ data:Object.values(md), backgroundColor:PALETTE, borderColor:'#060a12', borderWidth:2 }]};
}

function getOEEByModelData() {
    const md = {};
    filteredData.forEach(item => {
        const m = item.model||'Unknown';
        md[m] = md[m] || { avail:0, perf:0, qual:0, oee:0, count:0 };
        md[m].avail += n(item.availabilityPct);
        md[m].perf  += n(item.performancePct);
        md[m].qual  += n(item.qualityPct);
        md[m].oee   += computeOEE(item);
        md[m].count++;
    });
    const models = Object.keys(md).sort((a,b)=>md[b].oee/md[b].count - md[a].oee/md[a].count).slice(0,12);
    return { labels: models, datasets: [
        { label:'Availability %', data:models.map(m => md[m].count ? md[m].avail/md[m].count : 0), backgroundColor:C.cyanMid,  borderColor:C.cyan,  borderWidth:2 },
        { label:'Performance %',  data:models.map(m => md[m].count ? md[m].perf/md[m].count  : 0), backgroundColor:C.amberMid, borderColor:C.amber, borderWidth:2 },
        { label:'Quality %',      data:models.map(m => md[m].count ? md[m].qual/md[m].count  : 0), backgroundColor:C.greenMid, borderColor:C.green, borderWidth:2 },
        { label:'OEE %',          data:models.map(m => md[m].count ? md[m].oee/md[m].count   : 0), backgroundColor:C.redMid,   borderColor:C.red,   borderWidth:2 },
    ]};
}

function getQualityTrendData() {
    const dd = {};
    filteredData.forEach(item => {
        const dt = item.dateString || fmtDate(item.date);
        dd[dt] = dd[dt] || { tp:0, qOK:0, rej:0 };
        dd[dt].tp  += n(item.totalProduction);
        dd[dt].qOK += n(item.qualityOK);
        dd[dt].rej += n(item.rejections);
    });
    const dates = Object.keys(dd).sort().slice(-30);
    return { labels: dates, datasets: [
        { label:'Quality %',   data:dates.map(d => dd[d].tp>0 ? dd[d].qOK/dd[d].tp*100 : 0), borderColor:C.green, backgroundColor:C.greenDim, borderWidth:2.5, fill:true, tension:0.4 },
        { label:'Rejection %', data:dates.map(d => dd[d].tp>0 ? dd[d].rej/dd[d].tp*100 : 0), borderColor:C.red,   backgroundColor:C.redDim,   borderWidth:2.5, fill:true, tension:0.4 }
    ]};
}

function getPerformanceDistributionData() {
    let ex=0, go=0, fa=0, po=0;
    filteredData.forEach(item => {
        const p = computeOEE(item);
        if (p>85) ex++; else if (p>=70) go++; else if (p>=55) fa++; else po++;
    });
    return { labels:['World Class (>85%)','Good (70-85%)','Fair (55-70%)','Poor (<55%)'],
             datasets:[{ data:[ex,go,fa,po], backgroundColor:[C.green,C.cyan,C.amber,C.red], borderColor:'#060a12', borderWidth:2 }]};
}

function getAvailabilityPerformanceData() {
    const dd = {};
    filteredData.forEach(item => {
        const dt = item.dateString || fmtDate(item.date);
        dd[dt] = dd[dt] || { avail:0, perf:0, oee:0, count:0 };
        dd[dt].avail += n(item.availabilityPct);
        dd[dt].perf  += n(item.performancePct);
        dd[dt].oee   += computeOEE(item);
        dd[dt].count++;
    });
    const dates = Object.keys(dd).sort().slice(-30);
    return { labels: dates, datasets: [
        { label:'Availability %', data:dates.map(d => dd[d].count ? dd[d].avail/dd[d].count : 0), borderColor:C.cyan,  backgroundColor:C.cyanDim,  borderWidth:2.5, fill:true,  tension:0.4 },
        { label:'Performance %',  data:dates.map(d => dd[d].count ? dd[d].perf/dd[d].count  : 0), borderColor:C.amber, backgroundColor:C.amberDim, borderWidth:2.5, fill:false, tension:0.4 },
        { label:'OEE %',          data:dates.map(d => dd[d].count ? dd[d].oee/dd[d].count   : 0), borderColor:C.green, backgroundColor:C.greenDim, borderWidth:2,   fill:false, tension:0.4, borderDash:[4,3] }
    ]};
}

function getRejectionByModelData() {
    const md = {};
    filteredData.forEach(item => { const m=item.model||'Unknown'; md[m]=(md[m]||0)+n(item.rejections); });
    const sorted = Object.keys(md).sort((a,b)=>md[b]-md[a]).slice(0,10);
    return { labels:sorted, datasets:[{ label:'Rejections', data:sorted.map(m=>md[m]), backgroundColor:C.redMid, borderColor:C.red, borderWidth:2 }]};
}

function getPlanVsActualData() {
    const md = {};
    filteredData.forEach(item => {
        const m = item.model||'Unknown';
        md[m] = md[m] || { production:0, plan:0 };
        md[m].production += n(item.totalProduction);
        md[m].plan       += n(item.todaysPlan);
    });
    const models = Object.keys(md).sort((a,b)=>md[b].production-md[a].production).slice(0,10);
    return { labels:models, datasets:[
        { label:'Actual', data:models.map(m=>md[m].production), backgroundColor:C.cyanMid,  borderColor:C.cyan,  borderWidth:2 },
        { label:'Plan',   data:models.map(m=>md[m].plan),       backgroundColor:C.greenMid, borderColor:C.green, borderWidth:2 }
    ]};
}

function getProductionHeatmapData() {
    const daily = {};
    filteredData.forEach(item => {
        const dt = item.dateString||fmtDate(item.date);
        daily[dt] = (daily[dt]||0) + n(item.totalProduction);
    });
    const dates    = Object.keys(daily).sort().slice(-28);
    const weeks    = [];
    const weekData = [];
    for(let i=0; i<dates.length; i+=7) {
        const sl = dates.slice(i, i+7);
        weeks.push(`W${Math.floor(i/7)+1}`);
        weekData.push(sl.map(d=>daily[d]||0));
    }
    const maxLen = weekData.length ? Math.max(...weekData.map(w=>w.length)) : 0;
    const cols   = ['rgba(0,229,255,0.10)','rgba(0,229,255,0.20)','rgba(0,229,255,0.32)',
                    'rgba(0,229,255,0.44)','rgba(0,229,255,0.56)','rgba(0,229,255,0.70)','rgba(0,229,255,0.85)'];
    if (maxLen === 0) return { labels:[], datasets:[] };
    return { labels: weeks, datasets: Array.from({length:maxLen}, (_,di) => ({
        label: ['Mon','Tue','Wed','Thu','Fri','Sat','Sun'][di] || `D${di+1}`,
        data:  weekData.map(w => w[di]||0),
        backgroundColor: cols[di%cols.length],
        borderColor: 'rgba(0,229,255,0.1)',
        borderWidth: 1
    }))};
}

function getParetoData() {
    const md = {};
    filteredData.forEach(item => { const m=item.model||'Unknown'; md[m]=(md[m]||0)+n(item.totalProduction); });
    const sorted = Object.keys(md).map(m=>({model:m,production:md[m]}))
        .sort((a,b)=>b.production-a.production).slice(0,10);
    const total = sorted.reduce((s,i)=>s+i.production,0);
    let sum=0;
    const cum = sorted.map(i => { sum+=i.production; return total>0 ? sum/total*100 : 0; });
    return { labels: sorted.map(i=>i.model), datasets:[
        { label:'Production',   data:sorted.map(i=>i.production), backgroundColor:C.cyanMid, borderColor:C.cyan, borderWidth:2, yAxisID:'y' },
        { label:'Cumulative %', data:cum, type:'line', borderColor:C.red, backgroundColor:'transparent', borderWidth:2.5, yAxisID:'y1', pointRadius:4, pointBackgroundColor:C.red }
    ]};
}

// ── Loss Categories ───────────────────────────────────────────────────────
function getLossCategoriesData() {
    const lossMap = {
        'Startup':       'lossStartup',
        'Setup':         'lossSetup',
        'Fixture':       'lossFixture',
        'HR':            'lossHR',
        'Press':         'lossPress',
        'Store':         'lossStore',
        'QC':            'lossQC',
        'Maint-Robot':   'lossMaintRobotProg',
        'Maint-Fault':   'lossMaintFault',
        'Maint-Shank':   'lossMaintShank',
        'Maint-Clamp':   'lossMaintClamp',
        'Maint-Logic':   'lossMaintLogic',
        'Maint-Utility': 'lossMaintUtility',
        'Maint-Sensor':  'lossMaintSensor',
        'Maint-MIG':     'lossMaintMig',
        'Maint-Tucker':  'lossMaintTucker',
        'Maint-SSW':     'lossMaintSSW',
        'Maint-SPM':     'lossMaintSPM',
        'PPC':           'lossPPC',
        'Mgmt':          'lossMgmt',
    };
    const totals = {};
    filteredData.forEach(item => {
        Object.entries(lossMap).forEach(([label, field]) => {
            totals[label] = (totals[label]||0) + n(item[field]);
        });
    });
    const sorted  = Object.entries(totals).filter(([,v])=>v>0).sort((a,b)=>b[1]-a[1]);
    const labels  = sorted.map(([k])=>k);
    const data    = sorted.map(([,v])=>v);
    const bgColors  = PALETTE.slice(0, labels.length).map(c => c + 'aa');
    const brdColors = PALETTE.slice(0, labels.length);
    return { labels, datasets:[{
        label:'Loss Time (min)', data,
        backgroundColor: bgColors.length ? bgColors : ['rgba(255,23,68,0.5)'],
        borderColor:     brdColors.length ? brdColors : [C.red],
        borderWidth: 1.5
    }]};
}

// ── Shift Production ──────────────────────────────────────────────────────
function getShiftProductionData() {
    const dd = {};
    filteredData.forEach(item => {
        const dt = item.dateString || fmtDate(item.date);
        dd[dt] = dd[dt] || { A:0, B:0, C:0 };
        dd[dt].A += n(item.aShiftProduction);
        dd[dt].B += n(item.bShiftProduction);
        dd[dt].C += n(item.cShiftProduction);
    });
    const dates = Object.keys(dd).sort().slice(-14);
    return { labels: dates, datasets:[
        { label:'A Shift', data:dates.map(d=>dd[d].A), backgroundColor:'rgba(0,229,255,0.55)',  borderColor:C.cyan,  borderWidth:1.5 },
        { label:'B Shift', data:dates.map(d=>dd[d].B), backgroundColor:'rgba(0,230,118,0.55)',  borderColor:C.green, borderWidth:1.5 },
        { label:'C Shift', data:dates.map(d=>dd[d].C), backgroundColor:'rgba(255,171,0,0.55)',  borderColor:C.amber, borderWidth:1.5 }
    ]};
}

// ── Shop Area ─────────────────────────────────────────────────────────────
function getShopAreaData() {
    const md = {};
    filteredData.forEach(item => {
        const area = (item.prodShopArea || 'Unknown').trim();
        md[area] = (md[area]||0) + n(item.totalProduction);
    });
    const sorted  = Object.entries(md).filter(([,v])=>v>0).sort((a,b)=>b[1]-a[1]).slice(0,10);
    const labels  = sorted.map(([k])=>k);
    const data    = sorted.map(([,v])=>v);
    return { labels, datasets:[{
        label:'Production',
        data,
        backgroundColor: PALETTE.slice(0,labels.length).map(c=>c+'88'),
        borderColor:     PALETTE.slice(0,labels.length),
        borderWidth: 1.5
    }]};
}

// ── Process Type ──────────────────────────────────────────────────────────
function getProcessTypeData() {
    const md = {};
    filteredData.forEach(item => {
        const pt = (item.processType || 'Unknown').trim();
        md[pt] = (md[pt]||0) + n(item.totalProduction);
    });
    return { labels:Object.keys(md), datasets:[{
        data: Object.values(md),
        backgroundColor: PALETTE,
        borderColor: '#060a12',
        borderWidth: 2
    }]};
}

// ── Export ────────────────────────────────────────────────────────────────
const LOSS_FIELDS = ['lossStartup','lossSetup','lossFixture','lossHR','lossPress','lossStore',
    'lossQC','lossMaintRobotProg','lossMaintFault','lossMaintShank','lossMaintClamp',
    'lossMaintLogic','lossMaintUtility','lossMaintSensor','lossMaintMig','lossMaintTucker',
    'lossMaintSSW','lossMaintSPM','lossPPC','lossMgmt'];

function exportToExcel() {
    const data = filteredData.map(item => ({
        'Date': item.dateString||fmtDate(item.date),
        'Part No': item.partNo,
        'Model': item.model,
        'Shop Area': item.prodShopArea,
        'Process': item.processType,
        'Plan': n(item.todaysPlan),
        'Production': n(item.totalProduction),
        'A Shift': n(item.aShiftProduction),
        'B Shift': n(item.bShiftProduction),
        'C Shift': n(item.cShiftProduction),
        'Quality OK': n(item.qualityOK),
        'Rejections': n(item.rejections),
        'OEE %': computeOEE(item).toFixed(2),
        'Performance %': n(item.performancePct).toFixed(2),
        'Quality %': n(item.qualityPct).toFixed(2),
        'Availability %': n(item.availabilityPct).toFixed(2),
        'Capacity Util %': n(item.capacityUtilization).toFixed(2),
        'Total Loss (min)': LOSS_FIELDS.reduce((s,f)=>s+n(item[f]),0).toFixed(1)
    }));
    const ws = XLSX.utils.json_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Production Data');
    XLSX.writeFile(wb, `JBM_MIS_Export_${fmtDate(new Date())}.xlsx`);
}

function exportToPDF() {
    try {
        const { jsPDF } = window.jspdf;
        const doc = new jsPDF('l','mm','a4');
        doc.setFontSize(18); doc.text('JBM MIS Dashboard Report', 14, 15);
        doc.setFontSize(11); doc.text(`Generated: ${fmtDate(new Date())}`, 14, 22);
        const tableData = filteredData.slice(0,100).map(item => [
            item.dateString||fmtDate(item.date),
            item.partNo||'-',
            item.model||'-',
            fmt(item.todaysPlan),
            fmt(item.totalProduction),
            computeOEE(item).toFixed(1)+'%',
            n(item.performancePct).toFixed(1)+'%',
            n(item.qualityPct).toFixed(1)+'%',
            n(item.availabilityPct).toFixed(1)+'%'
        ]);
        doc.autoTable({
            startY: 30,
            head: [['Date','Part No','Model','Plan','Production','OEE%','Perf%','Qual%','Avail%']],
            body: tableData,
            theme: 'grid',
            headStyles: { fillColor:[0,180,220], textColor:[0,0,0], fontSize:9 },
            bodyStyles: { fontSize:8 },
            margin: { top:30 }
        });
        doc.save(`JBM_MIS_Report_${fmtDate(new Date())}.pdf`);
        showToast('PDF exported successfully', 'success');
    } catch(e) { console.error(e); showToast('Error exporting PDF: '+e.message, 'danger'); }
}

// ── Utilities ─────────────────────────────────────────────────────────────
function showLoading() { document.getElementById('loadingOverlay')?.classList.add('active'); }
function hideLoading() { document.getElementById('loadingOverlay')?.classList.remove('active'); }

function showToast(message, type='info') {
    const t = document.createElement('div');
    t.className = `alert alert-${type} alert-dismissible fade show position-fixed`;
    t.style.cssText = 'top:20px;right:20px;z-index:10001;min-width:300px;';
    t.innerHTML = `${message}<button type="button" class="btn-close" data-bs-dismiss="alert"></button>`;
    document.body.appendChild(t);
    setTimeout(() => { t.classList.remove('show'); setTimeout(()=>t.remove(), 150); }, 5000);
}
