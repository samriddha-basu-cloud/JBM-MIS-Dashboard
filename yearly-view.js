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

const MONTH_NAMES = [
    'January','February','March','April','May','June',
    'July','August','September','October','November','December'
];
const MONTH_SHORT = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];

let allData  = [];   // raw rows for the year
let trendChart = null;

// ─────────────────────────────────────────────────────────
// Init
// ─────────────────────────────────────────────────────────
document.addEventListener('DOMContentLoaded', () => {
    populateYearSelector();

    document.getElementById('btnLoadYear').addEventListener('click', loadYearData);
    document.getElementById('btnExportYear').addEventListener('click', exportYear);
    document.getElementById('selectGroupBy').addEventListener('change', renderTable);

    const yr = new Date().getFullYear();
    document.getElementById('selectYear').value = yr;
    loadYearData();
});

function populateYearSelector() {
    const sel = document.getElementById('selectYear');
    const yr  = new Date().getFullYear();
    for (let y = yr; y >= yr - 10; y--) {
        const o = document.createElement('option');
        o.value = y;
        o.textContent = y;
        sel.appendChild(o);
    }
}

// ─────────────────────────────────────────────────────────
// Load — iterate every day of the year, fetch subcollections
// Structure: daily_entries/{YYYY-MM-DD}/data/{docId}
// ─────────────────────────────────────────────────────────
async function loadYearData() {
    showLoading();
    try {
        const year = parseInt(document.getElementById('selectYear').value);

        // Build every calendar date string for the year
        const dateStrings = [];
        for (let m = 0; m < 12; m++) {
            const daysInMonth = new Date(year, m + 1, 0).getDate();
            for (let d = 1; d <= daysInMonth; d++) {
                const mm = String(m + 1).padStart(2, '0');
                const dd = String(d).padStart(2, '0');
                dateStrings.push(`${year}-${mm}-${dd}`);
            }
        }

        allData = [];

        // Parallel fetch — one promise per calendar day
        const fetches = dateStrings.map(async (dateStr) => {
            try {
                const ref      = collection(db, 'daily_entries', dateStr, 'data');
                const snapshot = await getDocs(query(ref, orderBy('__name__')));
                snapshot.forEach(docSnap => {
                    allData.push({ dateString: dateStr, id: docSnap.id, ...docSnap.data() });
                });
            } catch {
                // day has no data — skip silently
            }
        });

        await Promise.all(fetches);

        // Sort by date → srNo
        allData.sort((a, b) => {
            if (a.dateString < b.dateString) return -1;
            if (a.dateString > b.dateString) return  1;
            return (a.srNo || 0) - (b.srNo || 0);
        });

        renderTable();
        updateKPIs();
        renderTrendChart();

        document.getElementById('recordCount').textContent = `${allData.length} records`;
    } catch (err) {
        console.error('Error loading year:', err);
        showToast('Error loading data: ' + err.message, 'danger');
    } finally {
        hideLoading();
    }
}

// ─────────────────────────────────────────────────────────
// Render Table — groups by selected grouping
// ─────────────────────────────────────────────────────────
function renderTable() {
    const tbody   = document.getElementById('tableBody');
    const groupBy = document.getElementById('selectGroupBy').value;
    tbody.innerHTML = '';

    if (allData.length === 0) {
        tbody.innerHTML = `
            <tr>
              <td colspan="30" class="text-center text-muted py-4" style="font-size:13px;">
                <i class="fas fa-inbox me-2"></i>No data found for this year.
              </td>
            </tr>`;
        document.getElementById('tableFooter').style.display = 'none';
        return;
    }

    // Grand totals accumulators
    let GT = newAcc();
    let globalRow = 1;

    if (groupBy === 'month') {
        renderGroupedByMonth(tbody, GT, globalRow);
    } else if (groupBy === 'part') {
        renderGroupedByKey(tbody, GT, globalRow, r => r.sapCode || '(no SAP)', 'Part (SAP Code)');
    } else {
        renderGroupedByKey(tbody, GT, globalRow, r => r.prodShopArea || '(no area)', 'Shop / Area');
    }

    // Write grand-total footer
    writeFooter(GT);
    document.getElementById('tableFooter').style.display = '';
}

// ── Group by Month ──────────────────────────────────────
function renderGroupedByMonth(tbody, GT, globalRow) {
    // Build month buckets
    const buckets = {};
    for (let m = 0; m < 12; m++) buckets[m] = [];
    allData.forEach(item => {
        const m = parseInt((item.dateString || '').split('-')[1] || '1') - 1;
        buckets[m].push(item);
    });

    for (let m = 0; m < 12; m++) {
        const rows = buckets[m];

        // Month header separator
        const sep = document.createElement('tr');
        sep.className = 'month-group-row';
        sep.innerHTML = `<td colspan="30">
            <i class="fas fa-calendar-alt me-1"></i>
            <strong>${MONTH_NAMES[m]}</strong>
            <span style="font-weight:400;margin-left:6px;font-size:10px;">(${rows.length} entries)</span>
        </td>`;
        tbody.appendChild(sep);

        if (rows.length === 0) {
            const empty = document.createElement('tr');
            empty.innerHTML = `<td colspan="30" class="text-center text-muted" style="font-size:10px;padding:4px;">No data for ${MONTH_NAMES[m]}</td>`;
            tbody.appendChild(empty);
            continue;
        }

        // Month accumulator
        const MA = newAcc();

        rows.forEach(item => {
            tbody.appendChild(buildDataRow(item, globalRow++, MA, GT));
        });

        // Month subtotal row
        tbody.appendChild(buildSubtotalRow(MA, MONTH_NAMES[m] + ' SUBTOTAL'));
    }
}

// ── Group by arbitrary key ──────────────────────────────
function renderGroupedByKey(tbody, GT, globalRow, keyFn, label) {
    // Collect unique keys preserving insertion order
    const keyOrder = [];
    const buckets  = {};
    allData.forEach(item => {
        const k = keyFn(item);
        if (!buckets[k]) { buckets[k] = []; keyOrder.push(k); }
        buckets[k].push(item);
    });
    keyOrder.sort();

    keyOrder.forEach(k => {
        const rows = buckets[k];
        const sep  = document.createElement('tr');
        sep.className = 'month-group-row';
        sep.innerHTML = `<td colspan="30">
            <i class="fas fa-layer-group me-1"></i>
            <strong>${k}</strong>
            <span style="font-weight:400;margin-left:6px;font-size:10px;">(${rows.length} entries)</span>
        </td>`;
        tbody.appendChild(sep);

        const GA = newAcc();
        rows.forEach(item => {
            tbody.appendChild(buildDataRow(item, globalRow++, GA, GT));
        });
        tbody.appendChild(buildSubtotalRow(GA, k + ' SUBTOTAL'));
    });
}

// ── Build a single data row ─────────────────────────────
function buildDataRow(item, rowNum, MA, GT) {
    const totalProd = parseFloat(item.totalProduction)    || 0;
    const qualOK    = parseFloat(item.qualityOK)          || 0;
    const rej       = parseFloat(item.rejections)         || 0;
    const rejPct    = totalProd > 0 ? (rej / totalProd * 100) : 0;
    const capUtil   = parseFloat(item.capacityUtilization)|| 0;
    const avail     = parseFloat(item.availabilityPct)    || 0;
    const perf      = parseFloat(item.performancePct)     || 0;
    const qual      = parseFloat(item.qualityPct)         || 0;
    const oee       = parseFloat(item.oeePct)             || 0;

    accumulate(MA, item, totalProd, qualOK, rej, avail, perf, qual, oee);
    accumulate(GT, item, totalProd, qualOK, rej, avail, perf, qual, oee);

    const tr = document.createElement('tr');
    tr.innerHTML = `
        <td class="fixed-col fc-1" style="text-align:center;">${rowNum}</td>
        <td class="fixed-col fc-2 col-fixed-assy" style="font-weight:600;color:#5b21b6;">${item.dateString || ''}</td>
        <td class="fixed-col fc-3 col-fixed-sap">${item.assyCategory || ''}</td>
        <td class="fixed-col fc-4 col-fixed-partno" title="${item.sapCode || ''}">${item.sapCode || ''}</td>
        <td class="fixed-col fc-5 col-fixed-desc" title="${item.partDescription || ''}">${item.partDescription || ''}</td>
        <td class="fixed-col fc-6 col-fixed-model">${item.model || ''}</td>
        <td class="fixed-col fc-7">${item.prodShopArea || ''}</td>
        <td>${fmtN(item.todaysPlan)}</td>
        <td>${fmtD(item.capacityPerDay)}</td>
        <td>${fmtN(item.mpA)}</td>
        <td>${fmtN(item.mpB)}</td>
        <td>${fmtN(item.mpC)}</td>
        <td>${fmtD(item.totalMdays)}</td>
        <td>${fmtN(item.aShiftProduction)}</td>
        <td>${fmtN(item.bShiftProduction)}</td>
        <td>${fmtN(item.cShiftProduction)}</td>
        <td style="font-weight:700;">${fmtN(totalProd)}</td>
        <td class="${getCapClass(capUtil)}">${fmtD(capUtil)}%</td>
        <td>${fmtN(qualOK)}</td>
        <td style="color:${rej>0?'#dc3545':'inherit'};font-weight:${rej>0?'700':'400'};">${fmtN(rej)}</td>
        <td class="${rejPct>5?'st-bad':rejPct>2?'st-warn':''}">${rejPct.toFixed(1)}%</td>
        <td>${fmtN(item.totalLossTime)}</td>
        <td>${fmtN(item.noPlanTime)}</td>
        <td title="${item.rejectionReason||''}" style="max-width:80px;overflow:hidden;text-overflow:ellipsis;">${item.rejectionReason||''}</td>
        <td class="${getStatusClass(avail)}">${fmtD(avail)}%</td>
        <td class="${getStatusClass(perf)}">${fmtD(perf)}%</td>
        <td class="${getStatusClass(qual)}">${fmtD(qual)}%</td>
        <td class="${getStatusClass(oee)}" style="font-weight:700;">${fmtD(oee)}%</td>
        <td>${fmtD(item.stdManhead)}</td>
        <td>${fmtD(item.cycleTime)}</td>
    `;
    return tr;
}

// ── Build subtotal row ──────────────────────────────────
function buildSubtotalRow(A, label) {
    const n       = A.count || 1;
    const rejPct  = A.prod > 0 ? (A.rej  / A.prod  * 100) : 0;
    const capU    = A.prod > 0 ? (A.prod / (A.capSum) * 100) : 0;

    const tr = document.createElement('tr');
    tr.className = 'month-subtotal-row';
    tr.innerHTML = `
        <td class="fixed-col fc-1">∑</td>
        <td class="fixed-col fc-2 col-fixed-assy" style="font-size:10px;">${label}</td>
        <td class="fixed-col fc-3 col-fixed-sap">—</td>
        <td class="fixed-col fc-4 col-fixed-partno">—</td>
        <td class="fixed-col fc-5 col-fixed-desc">—</td>
        <td class="fixed-col fc-6 col-fixed-model">—</td>
        <td class="fixed-col fc-7">—</td>
        <td>${formatNumber(A.plan)}</td>
        <td>—</td>
        <td>—</td><td>—</td><td>—</td>
        <td>${A.mdays.toFixed(0)}</td>
        <td>—</td><td>—</td><td>—</td>
        <td style="font-weight:700;">${formatNumber(A.prod)}</td>
        <td class="${getCapClass(capU)}">${capU.toFixed(1)}%</td>
        <td>${formatNumber(A.qOK)}</td>
        <td>${formatNumber(A.rej)}</td>
        <td class="${rejPct>5?'st-bad':rejPct>2?'st-warn':''}">${rejPct.toFixed(1)}%</td>
        <td>${formatNumber(A.loss)}</td>
        <td>—</td>
        <td>—</td>
        <td class="${getStatusClass(A.avail/n)}">${(A.avail/n).toFixed(1)}%</td>
        <td>${(A.perf/n).toFixed(1)}%</td>
        <td>${(A.qual/n).toFixed(1)}%</td>
        <td class="${getStatusClass(A.oee/n)}" style="font-weight:700;">${(A.oee/n).toFixed(1)}%</td>
        <td>—</td>
        <td>—</td>
    `;
    return tr;
}

// ─────────────────────────────────────────────────────────
// KPI Cards
// ─────────────────────────────────────────────────────────
function updateKPIs() {
    const n       = allData.length || 1;
    const totProd = allData.reduce((s,i) => s+(parseFloat(i.totalProduction)||0), 0);
    const totPlan = allData.reduce((s,i) => s+(parseFloat(i.todaysPlan)     ||0), 0);
    const totRej  = allData.reduce((s,i) => s+(parseFloat(i.rejections)     ||0), 0);
    const avgOEE  = allData.reduce((s,i) => s+(parseFloat(i.oeePct)         ||0), 0) / n;
    const avgQual = allData.reduce((s,i) => s+(parseFloat(i.qualityPct)     ||0), 0) / n;
    const avgPerf = allData.reduce((s,i) => s+(parseFloat(i.performancePct) ||0), 0) / n;
    const avgAvail= allData.reduce((s,i) => s+(parseFloat(i.availabilityPct)||0), 0) / n;
    const rejPct  = totProd > 0 ? (totRej / totProd * 100) : 0;

    // Unique working days
    const uniqueDays = new Set(allData.map(i => i.dateString)).size;

    document.getElementById('kpiProduction').textContent  = formatNumber(totProd);
    document.getElementById('kpiOEE').textContent         = avgOEE.toFixed(1) + '%';
    document.getElementById('kpiQuality').textContent     = avgQual.toFixed(1) + '%';
    document.getElementById('kpiPerformance').textContent = avgPerf.toFixed(1) + '%';
    document.getElementById('kpiPlan').textContent        = formatNumber(totPlan);
    document.getElementById('kpiRejections').textContent  = formatNumber(totRej);
    document.getElementById('kpiRejPct').textContent      = rejPct.toFixed(1) + '%';
    document.getElementById('kpiAvail').textContent       = avgAvail.toFixed(1) + '%';
    document.getElementById('kpiDays').textContent        = uniqueDays;
}

// ─────────────────────────────────────────────────────────
// Trend Chart — monthly production bars + OEE line
// ─────────────────────────────────────────────────────────
function renderTrendChart() {
    // Aggregate per month
    const byMonth = Array.from({length:12}, () => ({ prod:0, oeeSum:0, count:0 }));

    allData.forEach(item => {
        const m = parseInt((item.dateString || '').split('-')[1] || '1') - 1;
        byMonth[m].prod    += parseFloat(item.totalProduction) || 0;
        byMonth[m].oeeSum  += parseFloat(item.oeePct)         || 0;
        byMonth[m].count++;
    });

    const prodData = byMonth.map(m => m.prod);
    const oeeData  = byMonth.map(m => m.count > 0 ? +(m.oeeSum / m.count).toFixed(1) : 0);

    if (trendChart) { trendChart.destroy(); trendChart = null; }

    const ctx = document.getElementById('trendChart');
    trendChart = new Chart(ctx, {
        data: {
            labels: MONTH_SHORT,
            datasets: [
                {
                    type: 'bar',
                    label: 'Production',
                    data: prodData,
                    backgroundColor: 'rgba(25,118,210,0.7)',
                    borderColor: '#1565C0',
                    borderWidth: 1,
                    yAxisID: 'yProd',
                    order: 2
                },
                {
                    type: 'line',
                    label: 'Avg OEE %',
                    data: oeeData,
                    borderColor: '#059669',
                    backgroundColor: 'rgba(5,150,105,0.12)',
                    borderWidth: 2,
                    pointRadius: 3,
                    pointBackgroundColor: '#059669',
                    tension: 0.35,
                    fill: false,
                    yAxisID: 'yOEE',
                    order: 1
                }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            interaction: { mode: 'index', intersect: false },
            plugins: {
                legend: {
                    position: 'right',
                    labels: { font: { size: 10 }, boxWidth: 12, padding: 8 }
                },
                tooltip: { bodyFont: { size: 10 }, titleFont: { size: 10 } }
            },
            scales: {
                yProd: {
                    type: 'linear',
                    position: 'left',
                    title: { display: true, text: 'Production', font: { size: 9 } },
                    ticks: { font: { size: 9 } },
                    grid: { color: 'rgba(0,0,0,0.05)' }
                },
                yOEE: {
                    type: 'linear',
                    position: 'right',
                    min: 0, max: 100,
                    title: { display: true, text: 'OEE %', font: { size: 9 } },
                    ticks: { font: { size: 9 } },
                    grid: { drawOnChartArea: false }
                },
                x: {
                    ticks: { font: { size: 9 } },
                    grid: { color: 'rgba(0,0,0,0.04)' }
                }
            }
        }
    });
}

// ─────────────────────────────────────────────────────────
// Export — full field set identical to daily-entry export
// ─────────────────────────────────────────────────────────
function exportYear() {
    if (allData.length === 0) { showToast('No data to export', 'warning'); return; }

    const year = document.getElementById('selectYear').value;

    const exportData = allData.map((item, idx) => ({
        'Sr. No.':           idx + 1,
        'Date':              item.dateString,
        'Assy Cat.':         item.assyCategory,
        'SAP Code':          item.sapCode,
        'Part No.':          item.partNo,
        'Part Description':  item.partDescription,
        'Model':             item.model,
        'Variant':           item.variant,
        'Prod Shop/Area':    item.prodShopArea,
        'Ownership':         item.ownership,
        'Process Type':      item.processType,
        'Cell No.':          item.cellNo,
        'OP':                item.operation,
        'Std. ManHead':      item.stdManhead,
        'CT (Min.)':         item.cycleTime,
        'JPH':               item.jph,
        'Cap/day':           item.capacityPerDay,
        'Tgt Man/part':      item.targetManPerPart,
        'Plan':              item.todaysPlan,
        'Req. mdays':        item.reqMandays,
        'A Shift MP':        item.mpA,
        'B Shift MP':        item.mpB,
        'C Shift MP':        item.mpC,
        'Total Mdays':       item.totalMdays,
        'A Prod':            item.aShiftProduction,
        'B Prod':            item.bShiftProduction,
        'C Prod':            item.cShiftProduction,
        'Total Prod':        item.totalProduction,
        'Q. OK':             item.qualityOK,
        'NG/Rej':            item.rejections,
        'Rej Reason':        item.rejectionReason,
        'Ach Man/Part':      item.achievedManPerPart,
        'Startup':           item.lossStartup,
        'Setup Chg':         item.lossSetup,
        'Fixture':           item.lossFixture,
        'HR Loss':           item.lossHR,
        'Press':             item.lossPress,
        'BOP':               item.lossStore,
        'QC':                item.lossQC,
        'Robot Prog':        item.lossMaintRobotProg,
        'Robot Fault':       item.lossMaintFault,
        'Shank':             item.lossMaintShank,
        'Clamp':             item.lossMaintClamp,
        'Thickness':         item.lossMaintLogic,
        'Air/Gas':           item.lossMaintUtility,
        'Sensor':            item.lossMaintSensor,
        'Mig':               item.lossMaintMig,
        'Tucker':            item.lossMaintTucker,
        'SSW':               item.lossMaintSSW,
        'SPM':               item.lossMaintSPM,
        'PPC':               item.lossPPC,
        'Mgmt':              item.lossMgmt,
        'Total Loss':        item.totalLossTime,
        'No Plans':          item.noPlanTime,
        'Shift':             item.lossShift,
        'Duration':          item.lossDuration,
        'Cap Util%':         item.capacityUtilization,
        'Act Man/Part':      item.actualManPerPart,
        'Avail%':            item.availabilityPct,
        'Perf%':             item.performancePct,
        'Qual%':             item.qualityPct,
        'OEE%':              item.oeePct
    }));

    const ws = XLSX.utils.json_to_sheet(exportData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, `Year ${year}`);
    XLSX.writeFile(wb, `JBM_MIS_Year_${year}.xlsx`);
    showToast(`Exported ${allData.length} rows for ${year}`, 'success');
}

// ─────────────────────────────────────────────────────────
// Footer writer
// ─────────────────────────────────────────────────────────
function writeFooter(GT) {
    const n      = GT.count || 1;
    const rejPct = GT.prod  > 0 ? (GT.rej  / GT.prod  * 100) : 0;
    const capU   = GT.prod  > 0 ? (GT.prod / GT.capSum * 100) : 0;

    document.getElementById('ftPlan').textContent      = formatNumber(GT.plan);
    document.getElementById('ftTotalMdays').textContent = GT.mdays.toFixed(0);
    document.getElementById('ftTotalProd').textContent  = formatNumber(GT.prod);
    document.getElementById('ftCapUtil').textContent    = capU.toFixed(1) + '%';
    document.getElementById('ftQualOK').textContent     = formatNumber(GT.qOK);
    document.getElementById('ftRej').textContent        = formatNumber(GT.rej);
    document.getElementById('ftRejPct').textContent     = rejPct.toFixed(1) + '%';
    document.getElementById('ftTotalLoss').textContent  = formatNumber(GT.loss);
    document.getElementById('ftAvail').textContent      = (GT.avail/n).toFixed(1) + '%';
    document.getElementById('ftPerf').textContent       = (GT.perf /n).toFixed(1) + '%';
    document.getElementById('ftQual').textContent       = (GT.qual /n).toFixed(1) + '%';
    document.getElementById('ftOEE').textContent        = (GT.oee  /n).toFixed(1) + '%';
}

// ─────────────────────────────────────────────────────────
// Accumulator helpers
// ─────────────────────────────────────────────────────────
function newAcc() {
    return { plan:0, mdays:0, prod:0, capSum:0, qOK:0, rej:0, loss:0, avail:0, perf:0, qual:0, oee:0, count:0 };
}

function accumulate(A, item, totalProd, qualOK, rej, avail, perf, qual, oee) {
    A.plan   += parseFloat(item.todaysPlan)     || 0;
    A.mdays  += parseFloat(item.totalMdays)     || 0;
    A.prod   += totalProd;
    A.capSum += parseFloat(item.capacityPerDay) || 0;
    A.qOK    += qualOK;
    A.rej    += rej;
    A.loss   += parseFloat(item.totalLossTime)  || 0;
    A.avail  += avail;
    A.perf   += perf;
    A.qual   += qual;
    A.oee    += oee;
    A.count++;
}

// ─────────────────────────────────────────────────────────
// Formatting helpers
// ─────────────────────────────────────────────────────────
function formatNumber(num) {
    return new Intl.NumberFormat('en-IN').format(Math.round(num) || 0);
}

function fmtN(val) {
    const n = parseFloat(val);
    if (!n || n === 0) return '<span style="color:#cbd5e0;">—</span>';
    return new Intl.NumberFormat('en-IN').format(Math.round(n));
}

function fmtD(val) {
    const n = parseFloat(val);
    if (n === undefined || n === null || isNaN(n)) return '<span style="color:#cbd5e0;">—</span>';
    return n.toFixed(1);
}

function getStatusClass(value) {
    if (value >= 70) return 'st-good';
    if (value >= 50) return 'st-warn';
    return 'st-bad';
}

function getCapClass(value) {
    if (value >= 85) return 'st-good';
    if (value >= 60) return 'st-warn';
    return 'st-bad';
}

function showLoading() {
    document.getElementById('loadingOverlay').classList.add('active');
}

function hideLoading() {
    document.getElementById('loadingOverlay').classList.remove('active');
}

function showToast(message, type = 'info') {
    const toast = document.createElement('div');
    toast.className = `alert alert-${type} alert-dismissible fade show position-fixed`;
    toast.style.cssText = 'top: 20px; right: 20px; z-index: 10000; min-width: 300px;';
    toast.innerHTML = `${message}<button type="button" class="btn-close" data-bs-dismiss="alert"></button>`;
    document.body.appendChild(toast);
    setTimeout(() => {
        toast.classList.remove('show');
        setTimeout(() => toast.remove(), 150);
    }, 5000);
}
