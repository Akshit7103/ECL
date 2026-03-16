/* ═══════════════════════════════════════════════════════════════════════
   ECL Automation — Frontend Application
   ═══════════════════════════════════════════════════════════════════════ */

/* ── Math Utilities (Normal CDF / Inverse / Vasicek) ───────────────── */
function normCDF(x) {
    const a1=0.254829592, a2=-0.284496736, a3=1.421413741, a4=-1.453152027, a5=1.061405429, p=0.3275911;
    const sign = x < 0 ? -1 : 1;
    const ax = Math.abs(x) / Math.SQRT2;
    const t = 1 / (1 + p * ax);
    const y = 1 - (((((a5*t+a4)*t)+a3)*t+a2)*t+a1)*t*Math.exp(-ax*ax);
    return 0.5 * (1 + sign * y);
}

function normInv(p) {
    if (p <= 0) return -Infinity;
    if (p >= 1) return Infinity;
    if (p === 0.5) return 0;
    const a=[-3.969683028665376e1,2.209460984245205e2,-2.759285104469687e2,1.383577518672690e2,-3.066479806614716e1,2.506628277459239e0];
    const b=[-5.447609879822406e1,1.615858368580409e2,-1.556989798598866e2,6.680131188771972e1,-1.328068155288572e1];
    const c=[-7.784894002430293e-3,-3.223964580411365e-1,-2.400758277161838e0,-2.549732539343734e0,4.374664141464968e0,2.938163982698783e0];
    const d=[7.784695709041462e-3,3.224671290700398e-1,2.445134137142996e0,3.754408661907416e0];
    const pLow=0.02425, pHigh=1-pLow;
    let q, r;
    if (p < pLow) {
        q = Math.sqrt(-2*Math.log(p));
        return (((((c[0]*q+c[1])*q+c[2])*q+c[3])*q+c[4])*q+c[5]) / ((((d[0]*q+d[1])*q+d[2])*q+d[3])*q+1);
    } else if (p <= pHigh) {
        q = p - 0.5; r = q*q;
        return (((((a[0]*r+a[1])*r+a[2])*r+a[3])*r+a[4])*r+a[5])*q / (((((b[0]*r+b[1])*r+b[2])*r+b[3])*r+b[4])*r+1);
    } else {
        q = Math.sqrt(-2*Math.log(1-p));
        return -(((((c[0]*q+c[1])*q+c[2])*q+c[3])*q+c[4])*q+c[5]) / ((((d[0]*q+d[1])*q+d[2])*q+d[3])*q+1);
    }
}

function vasicekPD(ttc, rho, z) {
    if (ttc >= 1) return 1;
    return normCDF((normInv(ttc) - Math.sqrt(rho) * z) / Math.sqrt(1 - rho));
}

/* ── Color Utilities ───────────────────────────────────────────────── */
function heatColor(value, min, max) {
    if (max === min) return 'transparent';
    const ratio = Math.max(0, Math.min(1, (value - min) / (max - min)));
    const r = ratio < 0.5 ? Math.round(255 * ratio * 2) : 255;
    const g = ratio < 0.5 ? 255 : Math.round(255 * (1 - (ratio - 0.5) * 2));
    return `rgba(${r}, ${g}, 0, 0.22)`;
}

function zColor(z) {
    const abs = Math.min(Math.abs(z), 3) / 3;
    if (z < 0) return `rgba(229, 62, 62, ${abs * 0.35})`;
    return `rgba(56, 161, 105, ${abs * 0.35})`;
}

/* ── State ──────────────────────────────────────────────────────────── */
let DATA = null;
let downloadUrl = null;
let charts = {};
let runHistory = [];

/* ── File Upload ───────────────────────────────────────────────────── */
document.getElementById('dpd-file').addEventListener('change', e => handleFile(e, 'dpd'));
document.getElementById('weo-file').addEventListener('change', e => handleFile(e, 'weo'));

['dpd-zone', 'weo-zone'].forEach(id => {
    const zone = document.getElementById(id);
    const type = id.replace('-zone', '');
    zone.addEventListener('dragover', e => { e.preventDefault(); zone.style.borderColor = '#4299e1'; });
    zone.addEventListener('dragleave', () => { zone.style.borderColor = ''; });
    zone.addEventListener('drop', e => {
        e.preventDefault(); zone.style.borderColor = '';
        if (e.dataTransfer.files[0]) {
            document.getElementById(type + '-file').files = e.dataTransfer.files;
            handleFile({ target: { files: [e.dataTransfer.files[0]] } }, type);
        }
    });
});

function handleFile(e, type) {
    const file = e.target.files[0];
    const badge = document.getElementById(type + '-badge');
    const zone = document.getElementById(type + '-zone');
    if (file) { badge.textContent = file.name; zone.classList.add('has-file'); }
    else { badge.textContent = ''; zone.classList.remove('has-file'); }
    checkReady();
}

function checkReady() {
    const ready = document.getElementById('dpd-file').files.length > 0 &&
                  document.getElementById('weo-file').files.length > 0;
    document.getElementById('run-btn').disabled = !ready;
}

/* ── Dark Mode ─────────────────────────────────────────────────────── */
function toggleDarkMode() {
    document.body.classList.toggle('dark');
    const isDark = document.body.classList.contains('dark');
    document.getElementById('icon-sun').style.display  = isDark ? 'none' : 'block';
    document.getElementById('icon-moon').style.display = isDark ? 'block' : 'none';
    localStorage.setItem('ecl-dark', isDark ? '1' : '0');
    Object.values(charts).forEach(c => { if (c) c.update(); });
}
if (localStorage.getItem('ecl-dark') === '1') toggleDarkMode();

/* ── Config Toggle ─────────────────────────────────────────────────── */
function toggleConfig() {
    document.getElementById('config-body').classList.toggle('open');
    document.getElementById('config-chevron').classList.toggle('open');
}

/* ── Progress Overlay ──────────────────────────────────────────────── */
let progressTimer = null;
function showProgress() {
    const overlay = document.getElementById('progress-overlay');
    overlay.classList.add('active');
    document.querySelectorAll('.pstep').forEach(s => { s.classList.remove('active', 'done'); });
    document.getElementById('progress-fill').style.width = '0%';
    let step = 0;
    const total = 8;
    progressTimer = setInterval(() => {
        if (step > 0) document.getElementById('pstep-' + step).classList.replace('active', 'done');
        step++;
        if (step <= total) {
            document.getElementById('pstep-' + step).classList.add('active');
            document.getElementById('progress-fill').style.width = (step / total * 90) + '%';
        }
    }, 600);
}
function hideProgress() {
    clearInterval(progressTimer);
    document.querySelectorAll('.pstep').forEach(s => { s.classList.remove('active'); s.classList.add('done'); });
    document.getElementById('progress-fill').style.width = '100%';
    setTimeout(() => { document.getElementById('progress-overlay').classList.remove('active'); }, 500);
}

/* ── Run Computation ───────────────────────────────────────────────── */
async function runComputation() {
    const btn = document.getElementById('run-btn');
    const errBox = document.getElementById('error-box');
    btn.classList.add('loading');
    document.getElementById('run-text').textContent = 'Computing...';
    errBox.style.display = 'none';
    showProgress();

    const form = new FormData();
    form.append('dpd_file', document.getElementById('dpd-file').files[0]);
    form.append('weo_file', document.getElementById('weo-file').files[0]);
    form.append('shock', parseFloat(document.getElementById('shock').value) / 100);
    form.append('tm_start_year', document.getElementById('tm-start').value);
    form.append('hist_cutoff', document.getElementById('hist-cutoff').value);

    try {
        const resp = await fetch('/api/compute', { method: 'POST', body: form });
        const json = await resp.json();
        if (json.error) { errBox.textContent = 'Error: ' + json.error; errBox.style.display = 'block'; return; }
        DATA = json;
        downloadUrl = json.download_url;
        runHistory.push({ data: JSON.parse(JSON.stringify(json)), time: new Date().toLocaleTimeString() });
        if (runHistory.length > 5) runHistory.shift();
        renderResults();
    } catch (err) {
        errBox.textContent = 'Request failed: ' + err.message; errBox.style.display = 'block';
    } finally {
        hideProgress();
        btn.classList.remove('loading');
        document.getElementById('run-text').textContent = 'Run ECL Computation';
    }
}

/* ── Render Results ────────────────────────────────────────────────── */
function renderResults() {
    document.getElementById('results-section').style.display = 'block';
    setTimeout(() => document.getElementById('results-section').scrollIntoView({ behavior: 'smooth', block: 'start' }), 100);
    renderMetrics();
    renderODRChart();
    renderTTCChart();
    renderCorrChart();
    renderODRHistChart();
    renderTMDropdown();
    renderTMTable();
    renderODRTable();
    renderODRGradeHeatmap();
    renderMAVTable();
    renderFanChart();
    renderZFactorChart();
    renderPDTable('Base');
    renderPDChart('Base');
    renderSensitivity();
    renderComparison();
}

/* ── Metrics ───────────────────────────────────────────────────────── */
function renderMetrics() {
    const odr = DATA.odr_summary;
    const active = odr.filter(r => r.status !== 'no_data');
    const odrs = active.filter(r => r.odr !== null).map(r => r.odr);
    const avg = odrs.length ? odrs.reduce((a, b) => a + b, 0) / odrs.length : 0;
    const latest = odrs.length ? odrs[odrs.length - 1] : 0;
    document.getElementById('metrics-grid').innerHTML = `
        <div class="metric-card"><div class="metric-label">Total Periods</div><div class="metric-value">${odr.length}</div><div class="metric-sub">${active.length} with active data</div></div>
        <div class="metric-card green"><div class="metric-label">Average ODR</div><div class="metric-value">${(avg*100).toFixed(2)}%</div><div class="metric-sub">Across active periods</div></div>
        <div class="metric-card amber"><div class="metric-label">Latest ODR</div><div class="metric-value">${(latest*100).toFixed(2)}%</div><div class="metric-sub">${active.length ? active[active.length-1].period : '-'}</div></div>
        <div class="metric-card red"><div class="metric-label">Risk Grades</div><div class="metric-value">${DATA.ttc_rho.length}</div><div class="metric-sub">DPD bucket categories</div></div>`;
}

/* ── Helper: destroy & create chart ────────────────────────────────── */
function makeChart(key, canvas, cfg) {
    if (charts[key]) charts[key].destroy();
    charts[key] = new Chart(document.getElementById(canvas), cfg);
}

/* ── ODR Trend Chart ───────────────────────────────────────────────── */
function renderODRChart() {
    const active = DATA.odr_summary.filter(r => r.odr !== null && r.status !== 'no_data');
    makeChart('odrTrend', 'chart-odr-trend', {
        type: 'line',
        data: {
            labels: active.map(r => r.period),
            datasets: [{ label: 'ODR', data: active.map(r => r.odr*100), borderColor: '#2b6cb0', backgroundColor: 'rgba(43,108,176,.1)', fill: true, tension: .35, pointRadius: 5, pointBackgroundColor: active.map(r => r.status==='partial'?'#d69e2e':'#2b6cb0'), borderWidth: 2.5 }]
        },
        options: { responsive: true, plugins: { legend: { display: false }, tooltip: { callbacks: { label: c => `ODR: ${c.parsed.y.toFixed(4)}%` } } }, scales: { y: { title: { display: true, text: 'ODR (%)' }, beginAtZero: true } } }
    });
}

/* ── TTC Chart ─────────────────────────────────────────────────────── */
function renderTTCChart() {
    const g = DATA.ttc_rho.filter(r => r.ttc < 1);
    makeChart('ttc', 'chart-ttc', {
        type: 'bar',
        data: { labels: g.map(r=>r.grade), datasets: [{ label:'TTC PD', data: g.map(r=>r.ttc*100), backgroundColor:['#bee3f8','#90cdf4','#63b3ed','#4299e1'], borderColor:'#2b6cb0', borderWidth: 1.5, borderRadius: 6 }] },
        options: { responsive: true, plugins: { legend:{display:false}, tooltip:{callbacks:{label:c=>`TTC: ${c.parsed.y.toFixed(4)}%`}} }, scales: { y: { title:{display:true,text:'TTC PD (%)'}, beginAtZero:true } } }
    });
}

/* ── Asset Correlation Curve ───────────────────────────────────────── */
function renderCorrChart() {
    const cc = DATA.corr_curve;
    const pts = cc.filter((_, i) => i % 3 === 0 || i < 20);
    makeChart('corr', 'chart-corr', {
        type: 'line',
        data: {
            labels: pts.map(p => (p.pd * 100).toFixed(1) + '%'),
            datasets: [{
                label: 'Asset Correlation', data: pts.map(p => p.rho),
                borderColor: '#805ad5', backgroundColor: 'rgba(128,90,213,.08)', fill: true, tension: .4, pointRadius: 0, borderWidth: 2.5
            }]
        },
        options: {
            responsive: true,
            plugins: { legend:{display:false}, tooltip:{callbacks:{label:c=>`rho: ${c.parsed.y.toFixed(4)}`}} },
            scales: {
                x: { title:{display:true,text:'PD'}, ticks: { maxTicksLimit: 10 } },
                y: { title:{display:true,text:'Correlation (rho)'}, min: 0.02, max: 0.17 }
            }
        }
    });
}

/* ── ODR Histogram ─────────────────────────────────────────────────── */
function renderODRHistChart() {
    const odrs = DATA.odr_summary.filter(r => r.odr !== null && r.odr > 0).map(r => r.odr * 100);
    const bins = [0, 0.5, 1.0, 1.5, 2.0, 2.5, 3.0];
    const counts = bins.slice(0, -1).map((lo, i) => odrs.filter(v => v >= lo && v < bins[i+1]).length);
    counts.push(odrs.filter(v => v >= bins[bins.length-1]).length);
    const labels = bins.slice(0,-1).map((lo,i) => `${lo}-${bins[i+1]}%`);
    labels.push(`${bins[bins.length-1]}%+`);
    makeChart('odrHist', 'chart-odr-hist', {
        type: 'bar',
        data: { labels, datasets: [{ label:'Periods', data:counts, backgroundColor:'rgba(43,108,176,.6)', borderRadius:4 }] },
        options: { responsive:true, plugins:{legend:{display:false}}, scales:{ y:{title:{display:true,text:'Count'},beginAtZero:true,ticks:{stepSize:1}}, x:{title:{display:true,text:'ODR Range'}} } }
    });
}

/* ── Transition Matrix Viewer ──────────────────────────────────────── */
function renderTMDropdown() {
    const sel = document.getElementById('tm-period-select');
    sel.innerHTML = '';
    Object.keys(DATA.odr_matrices).forEach(period => {
        sel.innerHTML += `<option value="${period}">${period}</option>`;
    });
}

function renderTMTable() {
    const period = document.getElementById('tm-period-select').value;
    const mat = DATA.odr_matrices[period];
    if (!mat) { document.getElementById('tm-table-wrap').innerHTML = '<p>No data for this period.</p>'; return; }
    const fromB = Object.keys(mat);
    const toB = ["0","1-30","31-60","61-90","90+","WO","ARC","Closed"];
    let allVals = [];
    fromB.forEach(fb => toB.forEach(tb => allVals.push(mat[fb][tb])));
    const maxVal = Math.max(...allVals);

    let html = '<table class="data-table"><thead><tr><th>From \\ To</th>';
    toB.forEach(tb => html += `<th>${tb}</th>`);
    html += '<th>Total</th><th>Default %</th></tr></thead><tbody>';
    fromB.forEach(fb => {
        const row = mat[fb];
        const total = toB.reduce((s, tb) => s + row[tb], 0);
        const dft = (row["90+"]||0) + (row["WO"]||0) + (row["ARC"]||0);
        const dftPct = total ? (dft/total*100) : 0;
        html += `<tr><td>${fb}</td>`;
        toB.forEach(tb => {
            const v = row[tb];
            html += `<td class="heat-cell" style="background:${heatColor(v, 0, maxVal)}">${v}</td>`;
        });
        html += `<td><strong>${total}</strong></td><td class="heat-cell" style="background:${heatColor(dftPct,0,100)}">${dftPct.toFixed(2)}%</td></tr>`;
    });
    html += '</tbody></table>';
    document.getElementById('tm-table-wrap').innerHTML = html;

    // Heatmap chart
    const chartData = fromB.map(fb => toB.map(tb => mat[fb][tb]));
    makeChart('tmHeat', 'chart-tm-heat', {
        type: 'bar',
        data: {
            labels: fromB,
            datasets: toB.map((tb, i) => ({
                label: tb, data: chartData.map(r => r[i]),
                backgroundColor: ['#bee3f8','#90cdf4','#63b3ed','#4299e1','#2b6cb0','#fc8181','#f6ad55','#cbd5e0'][i]
            }))
        },
        options: { responsive:true, plugins:{legend:{position:'top'}}, scales:{ x:{stacked:true,title:{display:true,text:'From Bucket'}}, y:{stacked:true,title:{display:true,text:'Count'}} } }
    });
}

/* ── ODR Table ─────────────────────────────────────────────────────── */
function renderODRTable() {
    let html = '<table class="data-table"><thead><tr><th>Period</th><th>ODR</th><th>Months</th><th>Observations</th><th>Status</th></tr></thead><tbody>';
    DATA.odr_summary.forEach(r => {
        const bc = r.status==='full'?'badge-full':r.status==='partial'?'badge-partial':'badge-nodata';
        const lb = r.status==='full'?'Full':r.status==='partial'?'Partial':'No Data';
        html += `<tr><td>${r.period}</td><td>${r.odr!==null?(r.odr*100).toFixed(4)+'%':'N/A'}</td><td>${r.months}</td><td>${r.total_obs.toLocaleString()}</td><td><span class="badge ${bc}">${lb}</span></td></tr>`;
    });
    document.getElementById('odr-table-wrap').innerHTML = html + '</tbody></table>';
}

/* ── ODR by Grade Heatmap ──────────────────────────────────────────── */
function renderODRGradeHeatmap() {
    const grades = Object.keys(DATA.odr_by_grade);
    const periods = DATA.odr_by_grade[grades[0]]?.map(r => r.period) || [];
    let allOdrs = [];
    grades.forEach(g => DATA.odr_by_grade[g].forEach(r => { if(r.odr>0) allOdrs.push(r.odr*100); }));
    const maxO = Math.max(...allOdrs, 1);

    let html = '<table class="data-table"><thead><tr><th>Grade</th>';
    periods.forEach(p => html += `<th>${p}</th>`);
    html += '</tr></thead><tbody>';
    grades.forEach(g => {
        html += `<tr><td>${g}</td>`;
        DATA.odr_by_grade[g].forEach(r => {
            const v = r.odr * 100;
            html += `<td class="heat-cell" style="background:${heatColor(v, 0, maxO)}">${v.toFixed(2)}%</td>`;
        });
        html += '</tr>';
    });
    document.getElementById('odr-grade-table-wrap').innerHTML = html + '</tbody></table>';
}

/* ── MAV Table ─────────────────────────────────────────────────────── */
function renderMAVTable() {
    let html = '<table class="data-table"><thead><tr><th>Macroeconomic Variable</th><th>Code</th><th>LTM</th><th>SD</th></tr></thead><tbody>';
    DATA.mav_params.forEach(r => { html += `<tr><td>${r.mev}</td><td>${r.code}</td><td>${r.ltm.toFixed(2)}</td><td>${r.sd.toFixed(2)}</td></tr>`; });
    document.getElementById('mav-table-wrap').innerHTML = html + '</tbody></table>';
}

/* ── Fan Chart (GDP Z-Factor Scenarios) ────────────────────────────── */
function renderFanChart() {
    const s = DATA.scenarios;
    const fyrs = s.years.filter(y => parseInt(y) >= 2020);
    const idx = fyrs.map(y => s.years.indexOf(y));
    makeChart('fan', 'chart-fan', {
        type: 'line',
        data: {
            labels: fyrs,
            datasets: [
                { label: 'Upturn', data: idx.map(i=>s.Upturn[i]), borderColor:'rgba(56,161,105,.4)', backgroundColor:'rgba(56,161,105,.08)', fill:'+1', tension:.3, pointRadius:0, borderWidth:1, borderDash:[4,3] },
                { label: 'Base', data: idx.map(i=>s.Base[i]), borderColor:'#2b6cb0', backgroundColor:'rgba(43,108,176,.06)', fill:false, tension:.3, pointRadius:3, borderWidth:2.5 },
                { label: 'Downturn', data: idx.map(i=>s.Downturn[i]), borderColor:'rgba(229,62,62,.4)', backgroundColor:'rgba(229,62,62,.08)', fill:'-1', tension:.3, pointRadius:0, borderWidth:1, borderDash:[4,3] },
            ]
        },
        options: { responsive:true, plugins:{legend:{position:'top'},tooltip:{mode:'index',intersect:false}}, scales:{y:{title:{display:true,text:'Z-Factor'}}} }
    });
}

/* ── All MEV Z-Factor Chart ────────────────────────────────────────── */
function renderZFactorChart() {
    const colors = ['#e53e3e','#2b6cb0','#d69e2e','#38a169','#805ad5'];
    const mevs = Object.keys(DATA.z_factors);
    const yrs = Object.keys(DATA.z_factors[mevs[0]]).filter(y => parseInt(y) >= 2020);
    makeChart('zfactors', 'chart-zfactors', {
        type: 'line',
        data: {
            labels: yrs,
            datasets: mevs.map((m, i) => ({
                label: m, data: yrs.map(y => DATA.z_factors[m][y]),
                borderColor: colors[i], fill: false, tension: .3, borderWidth: 2, pointRadius: 2
            }))
        },
        options: { responsive:true, plugins:{legend:{position:'top'},tooltip:{mode:'index',intersect:false}}, scales:{y:{title:{display:true,text:'Z-Factor'}}} }
    });
}

/* ── PD Table with Heatmap ─────────────────────────────────────────── */
function renderPDTable(scenario) {
    const pd = DATA.vasicek_pd;
    const yrs = pd.years;
    let allPDs = [];
    DATA.ttc_rho.filter(r=>r.ttc<1).forEach(r => pd[scenario][r.grade].forEach(v => allPDs.push(v*100)));
    const maxPD = Math.max(...allPDs, 0.1);

    let html = '<table class="data-table"><thead><tr><th>Grade</th><th>TTC</th>';
    yrs.forEach(y => html += `<th>${y}</th>`);
    html += '</tr></thead><tbody>';
    DATA.ttc_rho.forEach(row => {
        html += `<tr><td>${row.grade}</td><td>${(row.ttc*100).toFixed(4)}%</td>`;
        pd[scenario][row.grade].forEach(v => {
            const pct = v * 100;
            html += `<td class="heat-cell" style="background:${heatColor(pct, 0, maxPD)}">${pct.toFixed(2)}%</td>`;
        });
        html += '</tr>';
    });
    document.getElementById('pd-table-wrap').innerHTML = html + '</tbody></table>';
}

/* ── PD Chart ──────────────────────────────────────────────────────── */
function renderPDChart(scenario) {
    const pd = DATA.vasicek_pd;
    const yrs = pd.years;
    const colors = ['#2b6cb0','#4299e1','#d69e2e','#e53e3e','#718096'];
    const grades = DATA.ttc_rho.filter(r=>r.ttc<1).map(r=>r.grade);
    makeChart('pd', 'chart-pd', {
        type: 'line',
        data: { labels:yrs, datasets: grades.map((g,i)=>({ label:g, data:pd[scenario][g].map(v=>v*100), borderColor:colors[i], fill:false, tension:.3, borderWidth:2.5, pointRadius:4 })) },
        options: { responsive:true, plugins:{legend:{position:'top'},tooltip:{callbacks:{label:c=>`${c.dataset.label}: ${c.parsed.y.toFixed(4)}%`}}}, scales:{y:{title:{display:true,text:`PD (%) - ${scenario}`},beginAtZero:true}} }
    });
}

/* ── Scenario Selector ─────────────────────────────────────────────── */
function selectScenario(scenario) {
    document.querySelectorAll('.scen-btn').forEach(b => b.classList.remove('active'));
    document.querySelector(`.scen-btn[data-scen="${scenario}"]`).classList.add('active');
    renderPDTable(scenario);
    renderPDChart(scenario);
}

/* ── Sensitivity Analysis ──────────────────────────────────────────── */
function onSensitivityChange() {
    const shock = parseInt(document.getElementById('sens-shock').value) / 100;
    document.getElementById('sens-shock-val').textContent = (shock * 100).toFixed(0) + '%';
    renderSensitivity(shock);
}

function renderSensitivity(shockOverride) {
    if (!DATA) return;
    const shock = shockOverride !== undefined ? shockOverride : DATA.config_used.shock;
    const scenFilter = document.getElementById('sens-scenario').value;
    const zRaw = DATA.gdp_z_raw;
    const yrs = Object.keys(zRaw);
    const grades = DATA.ttc_rho.filter(r => r.ttc < 1);
    const scenarios = scenFilter === 'all' ? ['Base','Upturn','Downturn'] : [scenFilter];
    const colors = { Base:'#2b6cb0', Upturn:'#38a169', Downturn:'#e53e3e' };
    const dashes = { Base:[], Upturn:[6,3], Downturn:[6,3] };
    const datasets = [];

    scenarios.forEach(scen => {
        grades.forEach((g, gi) => {
            const data = yrs.map(yr => {
                const zBase = zRaw[yr];
                let z = zBase;
                if (scen === 'Upturn')   z = zBase + Math.abs(zBase) * shock;
                if (scen === 'Downturn') z = zBase - Math.abs(zBase) * shock;
                return vasicekPD(g.ttc, g.rho, z) * 100;
            });
            datasets.push({
                label: `${g.grade} (${scen})`,
                data,
                borderColor: colors[scen],
                borderDash: dashes[scen],
                borderWidth: scenarios.length === 1 ? 2.5 : 1.8,
                pointRadius: scenarios.length === 1 ? 4 : 2,
                fill: false, tension: .3,
                hidden: scenarios.length > 1 && gi > 1,
            });
        });
    });

    makeChart('sensitivity', 'chart-sensitivity', {
        type: 'line',
        data: { labels: yrs, datasets },
        options: {
            responsive: true, animation: { duration: 200 },
            plugins: { legend: { position: 'top', labels: { font: { size: 11 } } },
                       tooltip: { callbacks: { label: c => `${c.dataset.label}: ${c.parsed.y.toFixed(4)}%` } } },
            scales: { y: { title: { display:true, text:'PD (%)' }, beginAtZero:true } }
        }
    });

    // Sensitivity table
    let html = `<table class="data-table"><thead><tr><th>Grade</th><th>TTC</th><th>Rho</th>`;
    scenarios.forEach(s => yrs.forEach(y => html += `<th>${s[0]}${y.slice(-2)}</th>`));
    html += '</tr></thead><tbody>';
    grades.forEach(g => {
        html += `<tr><td>${g.grade}</td><td>${(g.ttc*100).toFixed(2)}%</td><td>${g.rho.toFixed(4)}</td>`;
        scenarios.forEach(scen => {
            yrs.forEach(yr => {
                const zBase = zRaw[yr];
                let z = zBase;
                if (scen === 'Upturn')   z = zBase + Math.abs(zBase) * shock;
                if (scen === 'Downturn') z = zBase - Math.abs(zBase) * shock;
                const pd = vasicekPD(g.ttc, g.rho, z) * 100;
                html += `<td class="heat-cell" style="background:${heatColor(pd, 0, 15)}">${pd.toFixed(2)}%</td>`;
            });
        });
        html += '</tr>';
    });
    document.getElementById('sens-table-wrap').innerHTML = html + '</tbody></table>';
}

/* ── Comparison ────────────────────────────────────────────────────── */
function renderComparison() {
    const sec = document.getElementById('comparison-section');
    if (runHistory.length < 2) { sec.style.display = 'none'; return; }
    sec.style.display = 'block';
    const prev = runHistory[runHistory.length - 2].data;
    const curr = runHistory[runHistory.length - 1].data;
    const prevTime = runHistory[runHistory.length - 2].time;
    const currTime = runHistory[runHistory.length - 1].time;

    let html = `<div class="comp-grid">
        <div class="comp-block"><h4>Previous Run (${prevTime})</h4>`;
    prev.ttc_rho.forEach(r => {
        html += `<div class="comp-val">${r.grade}: TTC=${(r.ttc*100).toFixed(4)}%, rho=${r.rho.toFixed(4)}</div>`;
    });
    html += `</div><div class="comp-block"><h4>Current Run (${currTime})</h4>`;
    curr.ttc_rho.forEach((r, i) => {
        const pttc = prev.ttc_rho[i]?.ttc || 0;
        const delta = (r.ttc - pttc) * 100;
        const cls = delta > 0 ? 'delta-neg' : delta < 0 ? 'delta-pos' : '';
        const sign = delta > 0 ? '+' : '';
        html += `<div class="comp-val">${r.grade}: TTC=${(r.ttc*100).toFixed(4)}% <span class="${cls}">(${sign}${delta.toFixed(4)}pp)</span>, rho=${r.rho.toFixed(4)}</div>`;
    });
    html += '</div></div>';

    // ODR comparison
    const prevOdrs = prev.odr_summary.filter(r => r.odr !== null);
    const currOdrs = curr.odr_summary.filter(r => r.odr !== null);
    if (prevOdrs.length && currOdrs.length) {
        const prevAvg = prevOdrs.reduce((s, r) => s + r.odr, 0) / prevOdrs.length * 100;
        const currAvg = currOdrs.reduce((s, r) => s + r.odr, 0) / currOdrs.length * 100;
        const d = currAvg - prevAvg;
        const cls = d > 0 ? 'delta-neg' : d < 0 ? 'delta-pos' : '';
        html += `<div style="margin-top:12px" class="comp-val"><strong>Avg ODR:</strong> ${prevAvg.toFixed(4)}% -> ${currAvg.toFixed(4)}% <span class="${cls}">(${d>0?'+':''}${d.toFixed(4)}pp)</span></div>`;
    }

    document.getElementById('comparison-content').innerHTML = html;
}

function clearComparison() {
    runHistory = runHistory.slice(-1);
    document.getElementById('comparison-section').style.display = 'none';
}

/* ── Tab Switching ─────────────────────────────────────────────────── */
function switchTab(tabId) {
    document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
    document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));
    document.querySelector(`.tab[data-tab="${tabId}"]`).classList.add('active');
    document.getElementById('tab-' + tabId).classList.add('active');
    // Re-render charts on tab switch (fixes canvas sizing)
    setTimeout(() => Object.values(charts).forEach(c => { if(c) c.resize(); }), 50);
}

/* ── Export / Print ────────────────────────────────────────────────── */
function printReport() { window.print(); }
function downloadExcel() { if (downloadUrl) window.location.href = downloadUrl; }

/* ── PDF Report Modal & Generation ────────────────────────────────── */
function openReportModal() {
    document.getElementById('report-modal').classList.add('active');
    const status = document.getElementById('report-status');
    status.style.display = 'none';
    status.className = 'modal-status';
}

function closeReportModal() {
    document.getElementById('report-modal').classList.remove('active');
}

async function generateReport() {
    if (!DATA) return;
    const btn = document.getElementById('btn-gen-report');
    const status = document.getElementById('report-status');

    btn.disabled = true;
    btn.innerHTML = '<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" class="spin"><circle cx="12" cy="12" r="10"/><path d="M12 6v6l4 2"/></svg> Generating...';
    status.style.display = 'block';
    status.className = 'modal-status loading';
    status.textContent = 'Building charts and tables... This may take a few seconds.';

    const company = document.getElementById('report-company').value.trim();
    const prepared_by = document.getElementById('report-author').value.trim();

    try {
        const resp = await fetch('/api/report', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ data: DATA, company, prepared_by }),
        });
        const json = await resp.json();
        if (json.error) {
            status.className = 'modal-status error';
            status.textContent = 'Error: ' + json.error;
            return;
        }
        status.className = 'modal-status success';
        status.textContent = 'Report generated successfully! Downloading...';
        setTimeout(() => { window.location.href = json.download_url; }, 400);
        setTimeout(closeReportModal, 2000);
    } catch (err) {
        status.className = 'modal-status error';
        status.textContent = 'Request failed: ' + err.message;
    } finally {
        btn.disabled = false;
        btn.innerHTML = '<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/></svg> Generate Report';
    }
}
