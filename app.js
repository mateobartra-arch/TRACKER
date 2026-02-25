// ============================================================
// 🚀 MISAGI DRIVER TRACKING — Web Dashboard Engine
// ============================================================

// =============================================
// ⚙️ CONFIGURACIÓN
// =============================================
// ⚠️ CAMBIA ESTA URL por la de tu nuevo Google Sheet
// Para obtener la URL: publica el sheet (Archivo → Compartir → Publicar en la web → CSV)
// Luego pega la URL aquí. Si usas la hoja ROSTER, agrega gid= con el ID de esa hoja.
const SHEET_URL = 'https://docs.google.com/spreadsheets/d/e/2PACX-1vTXzeKdoXCwEJhCWDE3w83oUZ1Jins--ZwZf70XgaEM1P-YCv-5dD0P1B1absEak77--NXtEO7nTvIt/pubhtml';

// Codificación de operaciones — misma que en el Apps Script
// Para agregar nueva operación: solo agregar aquí + en el Apps Script
const OPERATION_CODES = {
    'H1': { group: 'HUDBAY', color: '#8B4513', label: 'Hudbay' },
    'H2': { group: 'HUDBAY', color: '#A0522D', label: 'Hudbay' },
    'H3': { group: 'HUDBAY', color: '#CD853F', label: 'Hudbay' },
    'T1': { group: 'TOQUEPALA', color: '#228B22', label: 'Toquepala' },
    'T2': { group: 'TOQUEPALA', color: '#32CD32', label: 'Toquepala' },
    'C1': { group: 'CUAJONE', color: '#FFD700', label: 'Cuajone' },
    'C2': { group: 'CUAJONE', color: '#FFA500', label: 'Cuajone' },
    'A1': { group: 'HIERRO', color: '#1E90FF', label: 'Hierro' },
    'A2': { group: 'HIERRO', color: '#00BFFF', label: 'Hierro' },
    'A3': { group: 'HIERRO', color: '#87CEEB', label: 'Hierro' },
    'D': { group: 'DESCANSO', color: '#475569', label: 'Descanso' },
    'd': { group: 'DESCANSO', color: '#475569', label: 'Descanso' },
    'P': { group: 'PARADO', color: '#FCD34D', label: 'Parado' },
    'p': { group: 'PARADO', color: '#FCD34D', label: 'Parado' },
    'M/D': { group: 'MEDIA', color: '#94A3B8', label: 'Media Jornada' },
    'D/M': { group: 'MEDIA', color: '#94A3B8', label: 'Media Jornada' },
    'FERIADO': { group: 'FERIADO', color: '#C0C0C0', label: 'Feriado' }
    // ➕ AGREGA NUEVAS OPERACIONES AQUÍ
    // Ejemplo:
    // 'B1': { group: 'BAMBAS', color: '#059669', label: 'Bambas' },
};

const GROUP_COLORS = {
    'HUDBAY': '#8B4513',
    'TOQUEPALA': '#228B22',
    'CUAJONE': '#FFD700',
    'HIERRO': '#1E90FF',
    'DESCANSO': '#475569',
    'PARADO': '#FCD34D',
    'MEDIA': '#94A3B8',
    'FERIADO': '#C0C0C0'
};

const MONTH_NAMES = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio',
    'Julio', 'Agosto', 'Setiembre', 'Octubre', 'Noviembre', 'Diciembre'];

// =============================================
// 📦 GLOBAL STATE
// =============================================
let globalData = {
    drivers: [],  // [{name, records: [{date, code, comment}]}]
    dates: [],
    months: []    // unique months
};
let chartDonut = null;
let chartBar = null;
let chartDriverDetail = null;

// =============================================
// 🏃 INIT
// =============================================
document.addEventListener('DOMContentLoaded', async () => {
    try {
        await loadData();
        renderDashboard();
        populateSelectors();
        document.getElementById('loading').style.display = 'none';
        document.getElementById('dashboard-content').style.display = 'block';
        document.getElementById('lastUpdate').textContent =
            `Actualizado: ${new Date().toLocaleString('es-PE')}`;
    } catch (err) {
        document.getElementById('loading').innerHTML = `
      <div class="error-state">
        <p>❌ Error al conectar con Google Sheets</p>
        <p style="font-size:12px; margin-top:8px; color:#94a3b8">${err.message}</p>
        <p style="font-size:12px; margin-top:12px; color:#94a3b8">
          Asegúrate de que el sheet esté publicado:<br>
          Archivo → Compartir → Publicar en la web
        </p>
      </div>`;
    }
});

// =============================================
// 📡 FETCH & PARSE DATA
// =============================================

// Mapeo de nombres de mes (como aparecen en la fila 2 del sheet) a índice 0-11
const MONTH_MAP = {
    'ENERO': 0, 'FEBRERO': 1, 'MARZO': 2, 'ABRIL': 3,
    'MAYO': 4, 'JUNIO': 5, 'JULIO': 6, 'AGOSTO': 7,
    'SETIEMBRE': 8, 'SEPTIEMBRE': 8, 'OCTUBRE': 9,
    'NOVIEMBRE': 10, 'DICIEMBRE': 11
};

async function loadData() {
    const response = await fetch(SHEET_URL);
    const text = await response.text();

    const rows = parseCSV(text);
    if (rows.length < 6) throw new Error('No se encontraron datos suficientes');

    // ── Estructura del nuevo Google Sheet ──
    // Row 0: Título ("🚛 MISAGI", "", "ROSTER DE SEGUIMIENTO...")
    // Row 1: Subtítulo ("Tracking", "", "Control de Asistencia...")
    // Row 2: Meses ("" , "", "ENERO",,,,..."FEBRERO",,,...)
    // Row 3: Encabezados ("CONDUCTOR", "CARGO", 1, 2, 3, ... 31, 1, 2, ...)
    // Row 4: Días semana ("", "", "J", "V", "S", "D", ...)
    // Row 5+: Datos de conductores

    const monthRow = rows[2];   // nombres de meses
    const dayNumRow = rows[3];  // números de día (1-31)
    const dates = [];

    // Expandir los meses (las celdas merged aparecen vacías después del nombre)
    let currentMonth = -1;
    const monthPerCol = [];
    for (let i = 2; i < monthRow.length; i++) {
        const mesStr = monthRow[i].trim().toUpperCase();
        if (mesStr && MONTH_MAP[mesStr] !== undefined) {
            currentMonth = MONTH_MAP[mesStr];
        }
        monthPerCol.push(currentMonth);
    }

    // Construir fecha para cada columna de datos
    const DATA_YEAR = 2026;  // Año definido en el sheet
    for (let i = 2; i < dayNumRow.length; i++) {
        const dayStr = dayNumRow[i].trim();
        const dayNum = parseInt(dayStr);
        const mesIdx = monthPerCol[i - 2];

        if (!isNaN(dayNum) && dayNum >= 1 && dayNum <= 31 && mesIdx >= 0) {
            const date = new Date(DATA_YEAR, mesIdx, dayNum);
            dates.push({ col: i, date: date, str: `${dayNum}/${mesIdx + 1}/${DATA_YEAR}` });
        } else {
            dates.push({ col: i, date: null, str: '' });
        }
    }

    globalData.dates = dates;

    // Parse drivers — empiezan en fila 5 (índice 5)
    const drivers = [];
    for (let r = 5; r < rows.length; r++) {
        const name = rows[r][0].trim();
        if (!name) continue;

        // Filtrar filas de leyenda o vacías
        let codeCount = 0;
        let totalNonEmpty = 0;
        for (let c = 2; c < Math.min(rows[r].length, 60); c++) {
            const v = rows[r][c].trim();
            if (v) {
                totalNonEmpty++;
                if (lookupCode(v)) codeCount++;
            }
        }

        // Si menos del 30% son códigos válidos, es leyenda/helper
        if (totalNonEmpty > 0 && codeCount / totalNonEmpty < 0.3) continue;
        if (totalNonEmpty < 5) continue;

        const records = [];
        for (let c = 2; c < rows[r].length; c++) {
            const dateInfo = dates[c - 2];
            const val = rows[r][c].trim();
            const code = lookupCode(val);
            records.push({
                date: dateInfo ? dateInfo.date : null,
                dateStr: dateInfo ? dateInfo.str : '',
                code: code,
                rawValue: val,
                isComment: val && !code
            });
        }

        drivers.push({ name, records });
    }

    globalData.drivers = drivers;

    // Extraer meses únicos
    const monthSet = new Set();
    for (const d of dates) {
        if (d.date) {
            const key = `${d.date.getFullYear()}-${String(d.date.getMonth()).padStart(2, '0')}`;
            monthSet.add(key);
        }
    }
    globalData.months = [...monthSet].sort();
}

function parseCSV(text) {
    const rows = [];
    let current = '';
    let inQuotes = false;
    const lines = text.split('\n');

    for (const line of lines) {
        if (inQuotes) {
            current += '\n' + line;
        } else {
            current = line;
        }

        // Count quotes
        const quoteCount = (current.match(/"/g) || []).length;
        inQuotes = quoteCount % 2 !== 0;

        if (!inQuotes) {
            rows.push(parseCSVLine(current));
            current = '';
        }
    }

    return rows;
}

function parseCSVLine(line) {
    const result = [];
    let current = '';
    let inQuotes = false;

    for (let i = 0; i < line.length; i++) {
        const ch = line[i];
        if (ch === '"') {
            if (inQuotes && i + 1 < line.length && line[i + 1] === '"') {
                current += '"';
                i++;
            } else {
                inQuotes = !inQuotes;
            }
        } else if (ch === ',' && !inQuotes) {
            result.push(current);
            current = '';
        } else {
            current += ch;
        }
    }
    result.push(current);
    return result;
}

function lookupCode(val) {
    if (!val) return null;
    const v = val.trim();
    if (OPERATION_CODES[v]) return OPERATION_CODES[v];
    const upper = v.toUpperCase();
    for (const [key, info] of Object.entries(OPERATION_CODES)) {
        if (key.toUpperCase() === upper) return info;
    }
    return null;
}

// =============================================
// 🏠 DASHBOARD RENDERING
// =============================================
function renderDashboard() {
    renderKPIs();
    renderLegend();
    renderDriverCards();
    renderCharts();
}

function renderKPIs() {
    const drivers = globalData.drivers;
    const today = new Date();
    const currentMonth = today.getMonth();
    const currentYear = today.getFullYear();

    // Count total work days this month
    let totalWorkDays = 0;
    let totalRestDays = 0;
    let totalStoppedDays = 0;
    const opCounts = {};

    for (const driver of drivers) {
        for (const rec of driver.records) {
            if (!rec.date || !rec.code) continue;
            const group = rec.code.group;
            if (!['DESCANSO', 'PARADO', 'MEDIA', 'FERIADO'].includes(group)) {
                totalWorkDays++;
                opCounts[group] = (opCounts[group] || 0) + 1;
            } else if (group === 'DESCANSO') {
                totalRestDays++;
            } else if (group === 'PARADO') {
                totalStoppedDays++;
            }
        }
    }

    // Most frequent operation
    const topOp = Object.entries(opCounts).sort((a, b) => b[1] - a[1])[0];

    const kpiData = [
        { icon: '👥', value: drivers.length, label: 'Conductores' },
        { icon: '⛏️', value: Object.keys(getOperationGroups()).length, label: 'Operaciones' },
        { icon: '📅', value: totalWorkDays, label: 'Días trabajados (total)' },
        { icon: '🏆', value: topOp ? topOp[0] : 'N/A', label: 'Operación más frecuente' },
        { icon: '😴', value: totalRestDays, label: 'Días de descanso (total)' },
        { icon: '📊', value: globalData.months.length, label: 'Meses registrados' }
    ];

    const container = document.getElementById('kpiRow');
    container.innerHTML = kpiData.map(kpi => `
    <div class="kpi-card">
      <div class="kpi-icon">${kpi.icon}</div>
      <div class="kpi-value">${kpi.value}</div>
      <div class="kpi-label">${kpi.label}</div>
    </div>
  `).join('');
}

function getOperationGroups() {
    const groups = {};
    for (const [code, info] of Object.entries(OPERATION_CODES)) {
        if (!['DESCANSO', 'PARADO', 'MEDIA', 'FERIADO'].includes(info.group)) {
            groups[info.group] = info;
        }
    }
    return groups;
}

function renderLegend() {
    const legendItems = [
        { color: '#8B4513', label: 'Hudbay (H1-H3)' },
        { color: '#228B22', label: 'Toquepala (T1-T2)' },
        { color: '#FFD700', label: 'Cuajone (C1-C2)' },
        { color: '#1E90FF', label: 'Hierro (A1-A3)' },
        { color: '#475569', label: 'Descanso (D)' },
        { color: '#FCD34D', label: 'Parado (P)' },
        { color: '#94A3B8', label: 'Media Jornada' },
        { color: '#C0C0C0', label: 'Feriado' }
    ];

    const html = legendItems.map(item => `
    <div class="legend-item">
      <div class="legend-color" style="background:${item.color}"></div>
      ${item.label}
    </div>
  `).join('');

    document.getElementById('legend').innerHTML = html;
    const calLegend = document.getElementById('calendarLegend');
    if (calLegend) calLegend.innerHTML = html;
}

function renderDriverCards() {
    const container = document.getElementById('driverGrid');
    const avatarColors = ['#7c3aed', '#2563eb', '#ea580c', '#10b981', '#ef4444', '#fbbf24', '#06b6d4', '#8b5cf6'];

    container.innerHTML = globalData.drivers.map((driver, idx) => {
        const stats = getDriverStats(driver);
        const lastOp = getLastOperation(driver);
        const color = avatarColors[idx % avatarColors.length];
        const initials = driver.name.split(' ').map(w => w[0]).join('').substring(0, 2);

        // Build mini bar
        const total = stats.worked + stats.rest + stats.stopped;
        const pctWork = total > 0 ? (stats.worked / total * 100) : 0;
        const pctRest = total > 0 ? (stats.rest / total * 100) : 0;
        const pctStop = total > 0 ? (stats.stopped / total * 100) : 0;

        return `
      <div class="driver-card" onclick="showPage('conductores'); 
        document.getElementById('selectDriver').value='${idx}'; showDriverDetail();">
        <div class="driver-header">
          <div class="driver-avatar" style="background:${color}">${initials}</div>
          <div>
            <div class="driver-name">${driver.name}</div>
            <div class="driver-status" style="color:${lastOp.color}">
              <span class="dot" style="background:${lastOp.color}"></span>
              ${lastOp.label}
            </div>
          </div>
        </div>
        <div class="driver-stats">
          <div class="stat-item">
            <div class="stat-value" style="color:${lastOp.color}">${stats.worked}</div>
            <div class="stat-label">Trabajados</div>
          </div>
          <div class="stat-item">
            <div class="stat-value" style="color:#374151">${stats.rest}</div>
            <div class="stat-label">Descanso</div>
          </div>
          <div class="stat-item">
            <div class="stat-value" style="color:#fbbf24">${stats.stopped}</div>
            <div class="stat-label">Parado</div>
          </div>
        </div>
        <div class="mini-bar">
          <div class="segment" style="width:${pctWork}%; background:${lastOp.color}"></div>
          <div class="segment" style="width:${pctRest}%; background:#374151"></div>
          <div class="segment" style="width:${pctStop}%; background:#fbbf24"></div>
        </div>
      </div>
    `;
    }).join('');
}

function getDriverStats(driver) {
    let worked = 0, rest = 0, stopped = 0;
    const opCounts = {};

    for (const rec of driver.records) {
        if (!rec.code) continue;
        const g = rec.code.group;
        if (['DESCANSO'].includes(g)) rest++;
        else if (['PARADO'].includes(g)) stopped++;
        else if (!['MEDIA', 'FERIADO'].includes(g)) {
            worked++;
            opCounts[g] = (opCounts[g] || 0) + 1;
        }
    }

    return { worked, rest, stopped, opCounts };
}

function getLastOperation(driver) {
    // Find the last non-empty code
    for (let i = driver.records.length - 1; i >= 0; i--) {
        const rec = driver.records[i];
        if (rec.code && !['DESCANSO', 'PARADO', 'MEDIA', 'FERIADO'].includes(rec.code.group)) {
            return {
                label: rec.code.label + ' (' + rec.rawValue + ')',
                color: GROUP_COLORS[rec.code.group] || '#6b7280',
                group: rec.code.group
            };
        }
    }
    // Check for parado/descanso
    for (let i = driver.records.length - 1; i >= 0; i--) {
        const rec = driver.records[i];
        if (rec.code) {
            return {
                label: rec.code.label,
                color: GROUP_COLORS[rec.code.group] || '#6b7280',
                group: rec.code.group
            };
        }
    }
    return { label: 'Sin datos', color: '#6b7280', group: 'NONE' };
}

// =============================================
// 📊 CHARTS
// =============================================
function renderCharts() {
    renderDonutChart();
    renderBarChart();
}

function renderDonutChart() {
    const opGroups = {};
    for (const driver of globalData.drivers) {
        for (const rec of driver.records) {
            if (!rec.code) continue;
            const g = rec.code.group;
            if (!['DESCANSO', 'PARADO', 'MEDIA', 'FERIADO'].includes(g)) {
                opGroups[g] = (opGroups[g] || 0) + 1;
            }
        }
    }

    const labels = Object.keys(opGroups);
    const data = Object.values(opGroups);
    const colors = labels.map(l => GROUP_COLORS[l] || '#6b7280');

    if (chartDonut) chartDonut.destroy();
    chartDonut = new Chart(document.getElementById('chartDonut'), {
        type: 'doughnut',
        data: {
            labels,
            datasets: [{ data, backgroundColor: colors, borderWidth: 0 }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: true,
            plugins: {
                legend: {
                    position: 'bottom',
                    labels: { color: '#94a3b8', font: { family: 'Inter', size: 12 }, padding: 16 }
                }
            }
        }
    });
}

function renderBarChart() {
    const labels = globalData.drivers.map(d => {
        const parts = d.name.split(' ');
        return parts.length > 1 ? parts[parts.length - 1] : parts[0]; // Last name
    });

    const opGroupNames = Object.keys(getOperationGroups());
    const datasets = opGroupNames.map(group => {
        const data = globalData.drivers.map(driver => {
            const stats = getDriverStats(driver);
            return stats.opCounts[group] || 0;
        });
        return {
            label: group,
            data,
            backgroundColor: GROUP_COLORS[group] || '#6b7280'
        };
    });

    if (chartBar) chartBar.destroy();
    chartBar = new Chart(document.getElementById('chartBar'), {
        type: 'bar',
        data: { labels, datasets },
        options: {
            responsive: true,
            maintainAspectRatio: true,
            plugins: {
                legend: {
                    position: 'bottom',
                    labels: { color: '#94a3b8', font: { family: 'Inter', size: 11 }, padding: 12 }
                }
            },
            scales: {
                x: {
                    stacked: true,
                    ticks: { color: '#64748b', font: { family: 'Inter', size: 11 } },
                    grid: { color: 'rgba(255,255,255,0.03)' }
                },
                y: {
                    stacked: true,
                    ticks: { color: '#64748b', font: { family: 'Inter', size: 11 } },
                    grid: { color: 'rgba(255,255,255,0.05)' }
                }
            }
        }
    });
}

// =============================================
// 👤 DRIVER DETAIL
// =============================================
function showDriverDetail() {
    const idx = document.getElementById('selectDriver').value;
    if (idx === '') {
        document.getElementById('driverDetail').classList.remove('active');
        document.getElementById('driverChartCard').style.display = 'none';
        return;
    }

    const driver = globalData.drivers[parseInt(idx)];
    const stats = getDriverStats(driver);
    const lastOp = getLastOperation(driver);
    const total = stats.worked + stats.rest + stats.stopped;
    const actPct = total > 0 ? Math.round(stats.worked / total * 100) : 0;

    // Build operation summary
    const opSummary = Object.entries(stats.opCounts)
        .sort((a, b) => b[1] - a[1])
        .map(([op, count]) => `<strong>${op}</strong>: ${count} días`)
        .join(' · ');

    // Find comments
    const comments = driver.records.filter(r => r.isComment && r.rawValue);
    const commentHtml = comments.length > 0
        ? `<p style="margin-top:12px;">📝 <strong>Comentarios encontrados:</strong> ${comments.map(c =>
            `"${c.rawValue}"${c.date ? ' (' + c.date.toLocaleDateString('es-PE') + ')' : ''}`
        ).join(', ')}</p>`
        : '';

    const detail = document.getElementById('driverDetail');
    detail.innerHTML = `
    <button class="close-btn" onclick="this.parentElement.classList.remove('active');
      document.getElementById('driverChartCard').style.display='none';">✕</button>
    <h3>👤 ${driver.name}</h3>
    
    <div class="detail-stats-grid">
      <div class="detail-stat">
        <div class="val" style="color:${lastOp.color}">${stats.worked}</div>
        <div class="lbl">Días Trabajados</div>
      </div>
      <div class="detail-stat">
        <div class="val" style="color:#374151">${stats.rest}</div>
        <div class="lbl">Descanso</div>
      </div>
      <div class="detail-stat">
        <div class="val" style="color:#fbbf24">${stats.stopped}</div>
        <div class="lbl">Parado</div>
      </div>
      <div class="detail-stat">
        <div class="val" style="color:#00d4ff">${actPct}%</div>
        <div class="lbl">Actividad</div>
      </div>
      <div class="detail-stat">
        <div class="val" style="color:${lastOp.color}">${lastOp.group}</div>
        <div class="lbl">Última Operación</div>
      </div>
      <div class="detail-stat">
        <div class="val">${Object.keys(stats.opCounts).length}</div>
        <div class="lbl">Operaciones</div>
      </div>
    </div>
    
    <div class="summary-text">
      <p>📊 <strong>Resumen de operaciones:</strong> ${opSummary || 'Sin datos'}</p>
      <p style="margin-top:8px;">📈 <strong>Porcentaje de actividad:</strong> ${actPct}% del tiempo registrado fue en operación.</p>
      ${commentHtml}
    </div>
  `;
    detail.classList.add('active');

    // Render driver-specific chart
    renderDriverChart(driver);
}

function renderDriverChart(driver) {
    const stats = getDriverStats(driver);
    const labels = Object.keys(stats.opCounts);
    const data = Object.values(stats.opCounts);
    const colors = labels.map(l => GROUP_COLORS[l] || '#6b7280');

    document.getElementById('driverChartCard').style.display = 'block';

    if (chartDriverDetail) chartDriverDetail.destroy();
    chartDriverDetail = new Chart(document.getElementById('chartDriverDetail'), {
        type: 'doughnut',
        data: {
            labels,
            datasets: [{ data, backgroundColor: colors, borderWidth: 0 }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: true,
            plugins: {
                legend: {
                    position: 'bottom',
                    labels: { color: '#94a3b8', font: { family: 'Inter', size: 12 }, padding: 16 }
                }
            }
        }
    });
}

// =============================================
// 📋 MONTHLY REPORTS
// =============================================
function generateReport() {
    const monthKey = document.getElementById('selectMonth').value;
    if (!monthKey) {
        document.getElementById('reportTable').innerHTML = '';
        document.getElementById('reportSummary').style.display = 'none';
        return;
    }

    const [year, month] = monthKey.split('-').map(Number);
    const monthName = MONTH_NAMES[month];

    // Find dates in this month
    const monthDates = globalData.dates.filter(d =>
        d.date && d.date.getMonth() === month && d.date.getFullYear() === year
    );

    if (monthDates.length === 0) {
        document.getElementById('reportTable').innerHTML = '<p style="color:#94a3b8;padding:20px;">No hay datos para este mes.</p>';
        return;
    }

    // Build table
    let html = '<table class="report-table"><thead><tr>';
    html += '<th class="driver-col">Conductor</th>';

    for (const d of monthDates) {
        const day = d.date.getDate();
        const dow = d.date.getDay();
        const isWeekend = dow === 0 || dow === 6;
        html += `<th class="${isWeekend ? 'weekend' : ''}">${day}</th>`;
    }
    html += '<th>Trabajados</th><th>Descanso</th><th>Parado</th><th>%</th>';
    html += '</tr></thead><tbody>';

    // Summary data
    const summaryParts = [];

    for (const driver of globalData.drivers) {
        html += '<tr>';
        html += `<td class="driver-name-cell">${driver.name}</td>`;

        let worked = 0, rest = 0, stopped = 0;
        const opsThisMonth = {};

        for (const d of monthDates) {
            const colIdx = d.col - 2;
            const rec = driver.records[colIdx];
            const dow = d.date.getDay();
            const isWeekend = dow === 0 || dow === 6;

            if (rec && rec.rawValue) {
                const code = rec.code;
                let cellClass = isWeekend ? 'weekend' : '';
                let cellStyle = '';

                if (code) {
                    const g = code.group;
                    const prefix = rec.rawValue.toUpperCase().replace(/[0-9/]/g, '');
                    cellClass += ` code-cell code-${prefix || g[0]}`;
                    cellStyle = `background:${code.color}; color:white;`;

                    if (['DESCANSO'].includes(g)) { rest++; cellStyle = 'background:rgba(55,65,81,0.3); color:#9ca3af;'; }
                    else if (['PARADO'].includes(g)) { stopped++; cellStyle = 'background:rgba(251,191,36,0.15); color:#fbbf24;'; }
                    else if (!['MEDIA', 'FERIADO'].includes(g)) {
                        worked++;
                        opsThisMonth[g] = (opsThisMonth[g] || 0) + 1;
                    }
                } else {
                    cellClass += ' code-comment';
                }

                html += `<td class="${cellClass}" style="${cellStyle}" title="${rec.rawValue}">${rec.rawValue}</td>`;
            } else {
                html += `<td class="${isWeekend ? 'weekend' : ''}"></td>`;
            }
        }

        const total = worked + rest + stopped;
        const pct = total > 0 ? Math.round(worked / total * 100) : 0;

        html += `<td style="font-weight:800;color:#10b981">${worked}</td>`;
        html += `<td style="color:#6b7280">${rest}</td>`;
        html += `<td style="color:#fbbf24">${stopped}</td>`;
        html += `<td style="font-weight:800;color:#00d4ff">${pct}%</td>`;
        html += '</tr>';

        // Build summary text
        const opsText = Object.entries(opsThisMonth)
            .map(([op, count]) => `${count} días en ${op}`)
            .join(', ');
        if (opsText) {
            summaryParts.push(`<strong>${driver.name}</strong>: ${opsText} (${pct}% actividad)`);
        }
    }

    html += '</tbody></table>';
    document.getElementById('reportTable').innerHTML = html;

    // Summary text
    const summaryEl = document.getElementById('reportSummary');
    summaryEl.innerHTML = `
    <p>📊 <strong>Resumen de ${monthName} ${year}:</strong></p>
    <p style="margin-top:8px;">${summaryParts.join('<br>')}</p>
  `;
    summaryEl.style.display = 'block';
}

// =============================================
// 📅 CALENDAR
// =============================================
function renderCalendar() {
    const driverIdx = document.getElementById('calendarDriver').value;
    const monthKey = document.getElementById('calendarMonth').value;

    if (driverIdx === '' || monthKey === '') {
        document.getElementById('calendarGrid').innerHTML = '';
        return;
    }

    const driver = globalData.drivers[parseInt(driverIdx)];
    const [year, month] = monthKey.split('-').map(Number);

    const firstDay = new Date(year, month, 1);
    const lastDay = new Date(year, month + 1, 0);
    const startDow = firstDay.getDay(); // 0=Sunday

    const today = new Date();
    today.setHours(0, 0, 0, 0);

    let html = '';
    // Day headers
    const dayNames = ['Dom', 'Lun', 'Mar', 'Mié', 'Jue', 'Vie', 'Sáb'];
    for (const dn of dayNames) {
        html += `<div class="day-header">${dn}</div>`;
    }

    // Empty cells before first day
    for (let i = 0; i < startDow; i++) {
        html += '<div class="day-cell empty"></div>';
    }

    // Day cells
    for (let d = 1; d <= lastDay.getDate(); d++) {
        const cellDate = new Date(year, month, d);
        const isToday = cellDate.getTime() === today.getTime();

        // Find record for this date
        let dayContent = '';
        for (const rec of driver.records) {
            if (rec.date && rec.date.getDate() === d && rec.date.getMonth() === month && rec.date.getFullYear() === year) {
                if (rec.code) {
                    dayContent += `<div class="day-code" style="background:${rec.code.color};color:white;">${rec.rawValue}</div>`;
                } else if (rec.rawValue) {
                    dayContent += `<div class="day-code" style="background:rgba(255,255,255,0.1);color:#94a3b8;font-size:9px;">${rec.rawValue}</div>`;
                }
                break;
            }
        }

        html += `<div class="day-cell ${isToday ? 'today' : ''}">
      <div class="day-number">${d}</div>
      ${dayContent}
    </div>`;
    }

    document.getElementById('calendarGrid').innerHTML = html;
}

// =============================================
// 🧩 NAVIGATION & SELECTORS
// =============================================
function showPage(pageId) {
    // Hide all pages
    document.querySelectorAll('.section-page').forEach(p => p.classList.remove('active'));
    // Show target
    document.getElementById('page-' + pageId).classList.add('active');
    // Update nav
    document.querySelectorAll('.sidebar-nav a').forEach(a => {
        a.classList.toggle('active', a.dataset.page === pageId);
    });
}

function populateSelectors() {
    // Driver selectors
    const driverOptions = globalData.drivers.map((d, i) =>
        `<option value="${i}">${d.name}</option>`
    ).join('');

    const selDriver = document.getElementById('selectDriver');
    selDriver.innerHTML = '<option value="">— Seleccionar conductor —</option>' + driverOptions;

    const calDriver = document.getElementById('calendarDriver');
    calDriver.innerHTML = '<option value="">— Seleccionar conductor —</option>' + driverOptions;

    // Month selectors
    const monthOptions = globalData.months.map(m => {
        const [y, mo] = m.split('-').map(Number);
        return `<option value="${m}">${MONTH_NAMES[mo]} ${y}</option>`;
    }).join('');

    const selMonth = document.getElementById('selectMonth');
    selMonth.innerHTML = '<option value="">— Seleccionar mes —</option>' + monthOptions;

    const calMonth = document.getElementById('calendarMonth');
    calMonth.innerHTML = '<option value="">— Seleccionar mes —</option>' + monthOptions;
}

