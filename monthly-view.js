import { initializeApp } from 'https://www.gstatic.com/firebasejs/10.7.1/firebase-app.js';
import { getFirestore, collection, getDocs, orderBy, query } from 'https://www.gstatic.com/firebasejs/10.7.1/firebase-firestore.js';

const firebaseConfig = {
    apiKey: "AIzaSyDTczWACmsAcYKY743wLqNe0DiRAy5k6tk",
    authDomain: "mis-dashboard-f1688.firebaseapp.com",
    projectId: "mis-dashboard-f1688",
    storageBucket: "mis-dashboard-f1688.firebasestorage.app",
    messagingSenderId: "476511626207",
    appId: "1:476511626207:web:d3eb5f41cbec14b7325a6a"
};

const app = initializeApp(firebaseConfig);
const db = getFirestore(app);

let allData = []; // all raw rows loaded for the month

// ─────────────────────────────────────────────
// Init
// ─────────────────────────────────────────────
document.addEventListener('DOMContentLoaded', () => {
    populateSelectors();
    document.getElementById('btnLoadMonth').addEventListener('click', loadMonthData);
    document.getElementById('btnExportMonth').addEventListener('click', exportMonth);

    // Default to current month/year
    const now = new Date();
    document.getElementById('selectMonth').value = now.getMonth() + 1;
    document.getElementById('selectYear').value  = now.getFullYear();

    loadMonthData();
});

// ─────────────────────────────────────────────
// Selectors
// ─────────────────────────────────────────────
function populateSelectors() {
    const months = [
        'January','February','March','April','May','June',
        'July','August','September','October','November','December'
    ];
    const monthSel = document.getElementById('selectMonth');
    months.forEach((m, i) => {
        const o = document.createElement('option');
        o.value = i + 1;
        o.textContent = m;
        monthSel.appendChild(o);
    });

    const yearSel = document.getElementById('selectYear');
    const yr = new Date().getFullYear();
    for (let y = yr; y >= yr - 5; y--) {
        const o = document.createElement('option');
        o.value = y;
        o.textContent = y;
        yearSel.appendChild(o);
    }
}

// ─────────────────────────────────────────────
// Load Data
// Daily Entry stores: daily_entries/{YYYY-MM-DD}/data/{docId}
// We iterate each calendar day in the month and fetch sub-collections.
// ─────────────────────────────────────────────
async function loadMonthData() {
    showLoading();
    try {
        const month = parseInt(document.getElementById('selectMonth').value);
        const year  = parseInt(document.getElementById('selectYear').value);

        // Generate all dates in the month
        const daysInMonth = new Date(year, month, 0).getDate();
        const dateStrings = [];
        for (let d = 1; d <= daysInMonth; d++) {
            const mm = String(month).padStart(2, '0');
            const dd = String(d).padStart(2, '0');
            dateStrings.push(`${year}-${mm}-${dd}`);
        }

        allData = [];

        // Fetch each day's sub-collection in parallel
        const fetchPromises = dateStrings.map(async (dateStr) => {
            try {
                const dataRef  = collection(db, 'daily_entries', dateStr, 'data');
                const snapshot = await getDocs(query(dataRef, orderBy('__name__')));
                snapshot.forEach(docSnap => {
                    allData.push({ dateString: dateStr, id: docSnap.id, ...docSnap.data() });
                });
            } catch {
                // Day has no data — skip silently
            }
        });

        await Promise.all(fetchPromises);

        // Sort by date then by srNo / document order
        allData.sort((a, b) => {
            if (a.dateString < b.dateString) return -1;
            if (a.dateString > b.dateString) return  1;
            return (a.srNo || 0) - (b.srNo || 0);
        });

        renderTable();
        updateKPIs();
        document.getElementById('recordCount').textContent = `${allData.length} records`;
    } catch (error) {
        console.error('Error loading month data:', error);
        showToast('Error loading data: ' + error.message, 'danger');
    } finally {
        hideLoading();
    }
}

// ─────────────────────────────────────────────
// Render Table — row-per-entry, grouped by date
// ─────────────────────────────────────────────
function renderTable() {
    const tbody = document.getElementById('tableBody');
    tbody.innerHTML = '';

    if (allData.length === 0) {
        tbody.innerHTML = `
            <tr>
                <td colspan="30" class="text-center text-muted py-4" style="font-size:13px;">
                    <i class="fas fa-inbox me-2"></i>No data found for the selected month.
                </td>
            </tr>`;
        document.getElementById('tableFooter').style.display = 'none';
        return;
    }

    let globalRow = 1;
    let lastDate  = null;

    // Totals for footer
    let totPlan=0, totProd=0, totQOK=0, totRej=0, totLoss=0, totMdays=0;
    let sumAvail=0, sumPerf=0, sumQual=0, sumOEE=0, countRows=0;

    allData.forEach(item => {
        // Date group separator row
        if (item.dateString !== lastDate) {
            lastDate = item.dateString;
            const sep = document.createElement('tr');
            sep.className = 'date-group-row';
            sep.innerHTML = `<td colspan="30">
                <i class="fas fa-calendar-day me-1"></i>
                <strong>${formatDisplayDate(item.dateString)}</strong>
            </td>`;
            tbody.appendChild(sep);
        }

        const row = document.createElement('tr');
        const totalProd = parseFloat(item.totalProduction) || 0;
        const qualOK    = parseFloat(item.qualityOK)       || 0;
        const rej       = parseFloat(item.rejections)      || 0;
        const rejPct    = totalProd > 0 ? (rej / totalProd * 100) : 0;
        const capUtil   = parseFloat(item.capacityUtilization) || 0;
        const avail     = parseFloat(item.availabilityPct)     || 0;
        const perf      = parseFloat(item.performancePct)      || 0;
        const qual      = parseFloat(item.qualityPct)          || 0;
        const oee       = parseFloat(item.oeePct)              || 0;

        totPlan   += parseFloat(item.todaysPlan)  || 0;
        totProd   += totalProd;
        totQOK    += qualOK;
        totRej    += rej;
        totLoss   += parseFloat(item.totalLossTime) || 0;
        totMdays  += parseFloat(item.totalMdays)    || 0;
        sumAvail  += avail;
        sumPerf   += perf;
        sumQual   += qual;
        sumOEE    += oee;
        countRows++;

        row.innerHTML = `
            <td class="fixed-col fc-1" style="text-align:center;">${globalRow++}</td>
            <td class="fixed-col fc-2 col-fixed-assy" style="font-weight:600;color:#0369a1;">${item.dateString || ''}</td>
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
            <td class="${rejPct > 5 ? 'st-bad' : rejPct > 2 ? 'st-warn' : ''}">${rejPct.toFixed(1)}%</td>
            <td>${fmtN(item.totalLossTime)}</td>
            <td>${fmtN(item.noPlanTime)}</td>
            <td title="${item.rejectionReason || ''}" style="max-width:80px;overflow:hidden;text-overflow:ellipsis;">${item.rejectionReason || ''}</td>
            <td class="${getStatusClass(avail)}">${fmtD(avail)}%</td>
            <td class="${getStatusClass(perf)}">${fmtD(perf)}%</td>
            <td class="${getStatusClass(qual)}">${fmtD(qual)}%</td>
            <td class="${getStatusClass(oee)}" style="font-weight:700;">${fmtD(oee)}%</td>
            <td>${fmtD(item.stdManhead)}</td>
            <td>${fmtD(item.cycleTime)}</td>
        `;
        tbody.appendChild(row);
    });

    // Footer totals
    const avgAvail = countRows > 0 ? sumAvail / countRows : 0;
    const avgPerf  = countRows > 0 ? sumPerf  / countRows : 0;
    const avgQual  = countRows > 0 ? sumQual  / countRows : 0;
    const avgOEE   = countRows > 0 ? sumOEE   / countRows : 0;
    const totRejP  = totProd > 0 ? (totRej / totProd * 100) : 0;
    const totCapU  = totProd > 0 ? (totProd / (allData.reduce((s,i)=>s+(parseFloat(i.capacityPerDay)||0),0)) * 100) : 0;

    document.getElementById('ftPlan').textContent       = fmtN(totPlan);
    document.getElementById('ftTotalMdays').textContent  = fmtD(totMdays);
    document.getElementById('ftTotalProd').textContent   = fmtN(totProd);
    document.getElementById('ftCapUtil').textContent     = totCapU.toFixed(1) + '%';
    document.getElementById('ftQualOK').textContent      = fmtN(totQOK);
    document.getElementById('ftRej').textContent         = fmtN(totRej);
    document.getElementById('ftRejPct').textContent      = totRejP.toFixed(1) + '%';
    document.getElementById('ftTotalLoss').textContent   = fmtN(totLoss);
    document.getElementById('ftAvail').textContent       = avgAvail.toFixed(1) + '%';
    document.getElementById('ftPerf').textContent        = avgPerf.toFixed(1) + '%';
    document.getElementById('ftQual').textContent        = avgQual.toFixed(1) + '%';
    document.getElementById('ftOEE').textContent         = avgOEE.toFixed(1) + '%';

    document.getElementById('tableFooter').style.display = '';
}

// ─────────────────────────────────────────────
// Update KPI Cards
// ─────────────────────────────────────────────
function updateKPIs() {
    const totProd = allData.reduce((s, i) => s + (parseFloat(i.totalProduction) || 0), 0);
    const totPlan = allData.reduce((s, i) => s + (parseFloat(i.todaysPlan) || 0), 0);
    const totRej  = allData.reduce((s, i) => s + (parseFloat(i.rejections) || 0), 0);
    const n       = allData.length || 1;
    const avgOEE  = allData.reduce((s, i) => s + (parseFloat(i.oeePct)          || 0), 0) / n;
    const avgQual = allData.reduce((s, i) => s + (parseFloat(i.qualityPct)       || 0), 0) / n;
    const avgPerf = allData.reduce((s, i) => s + (parseFloat(i.performancePct)   || 0), 0) / n;
    const avgAvail= allData.reduce((s, i) => s + (parseFloat(i.availabilityPct)  || 0), 0) / n;
    const rejPct  = totProd > 0 ? (totRej / totProd * 100) : 0;

    document.getElementById('kpiProduction').textContent = formatNumber(totProd);
    document.getElementById('kpiOEE').textContent        = avgOEE.toFixed(1) + '%';
    document.getElementById('kpiQuality').textContent    = avgQual.toFixed(1) + '%';
    document.getElementById('kpiPerformance').textContent= avgPerf.toFixed(1) + '%';
    document.getElementById('kpiPlan').textContent       = formatNumber(totPlan);
    document.getElementById('kpiRejections').textContent = formatNumber(totRej);
    document.getElementById('kpiRejPct').textContent     = rejPct.toFixed(1) + '%';
    document.getElementById('kpiAvail').textContent      = avgAvail.toFixed(1) + '%';
}

// ─────────────────────────────────────────────
// Export — all fields matching daily-entry export
// ─────────────────────────────────────────────
function exportMonth() {
    if (allData.length === 0) {
        showToast('No data to export', 'warning');
        return;
    }

    const monthEl = document.getElementById('selectMonth');
    const monthName = monthEl.options[monthEl.selectedIndex].text;
    const year = document.getElementById('selectYear').value;

    const exportData = allData.map((item, idx) => ({
        'Sr. No.':            idx + 1,
        'Date':               item.dateString,
        'Assy Cat.':          item.assyCategory,
        'SAP Code':           item.sapCode,
        'Part No.':           item.partNo,
        'Part Description':   item.partDescription,
        'Model':              item.model,
        'Variant':            item.variant,
        'Prod Shop/Area':     item.prodShopArea,
        'Ownership':          item.ownership,
        'Process Type':       item.processType,
        'Cell No.':           item.cellNo,
        'OP':                 item.operation,
        'Std. ManHead':       item.stdManhead,
        'CT (Min.)':          item.cycleTime,
        'JPH':                item.jph,
        'Cap/day':            item.capacityPerDay,
        'Tgt Man/part':       item.targetManPerPart,
        'Plan':               item.todaysPlan,
        'Req. mdays':         item.reqMandays,
        'A Shift MP':         item.mpA,
        'B Shift MP':         item.mpB,
        'C Shift MP':         item.mpC,
        'Total Mdays':        item.totalMdays,
        'A Prod':             item.aShiftProduction,
        'B Prod':             item.bShiftProduction,
        'C Prod':             item.cShiftProduction,
        'Total Prod':         item.totalProduction,
        'Q. OK':              item.qualityOK,
        'NG/Rej':             item.rejections,
        'Rej Reason':         item.rejectionReason,
        'Ach Man/Part':       item.achievedManPerPart,
        'Startup':            item.lossStartup,
        'Setup Chg':          item.lossSetup,
        'Fixture':            item.lossFixture,
        'HR Loss':            item.lossHR,
        'Press':              item.lossPress,
        'BOP':                item.lossStore,
        'QC':                 item.lossQC,
        'Robot Prog':         item.lossMaintRobotProg,
        'Robot Fault':        item.lossMaintFault,
        'Shank':              item.lossMaintShank,
        'Clamp':              item.lossMaintClamp,
        'Thickness':          item.lossMaintLogic,
        'Air/Gas':            item.lossMaintUtility,
        'Sensor':             item.lossMaintSensor,
        'Mig':                item.lossMaintMig,
        'Tucker':             item.lossMaintTucker,
        'SSW':                item.lossMaintSSW,
        'SPM':                item.lossMaintSPM,
        'PPC':                item.lossPPC,
        'Mgmt':               item.lossMgmt,
        'Total Loss':         item.totalLossTime,
        'No Plans':           item.noPlanTime,
        'Shift':              item.lossShift,
        'Duration':           item.lossDuration,
        'Cap Util%':          item.capacityUtilization,
        'Act Man/Part':       item.actualManPerPart,
        'Avail%':             item.availabilityPct,
        'Perf%':              item.performancePct,
        'Qual%':              item.qualityPct,
        'OEE%':               item.oeePct
    }));

    const ws = XLSX.utils.json_to_sheet(exportData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Monthly Data');
    XLSX.writeFile(wb, `JBM_MIS_${monthName}_${year}.xlsx`);
    showToast(`Exported ${allData.length} rows for ${monthName} ${year}`, 'success');
}

// ─────────────────────────────────────────────
// Helpers
// ─────────────────────────────────────────────
function formatDisplayDate(dateStr) {
    if (!dateStr) return '';
    const [y, m, d] = dateStr.split('-');
    const months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
    return `${d} ${months[parseInt(m)-1]} ${y}`;
}

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
    if (!n && n !== 0) return '<span style="color:#cbd5e0;">—</span>';
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
