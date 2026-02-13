const fileInput = document.getElementById('fileInput');
const reportDiv = document.getElementById('report');
const controlsDiv = document.getElementById('controls');
const startDateInput = document.getElementById('startDate');
const endDateInput = document.getElementById('endDate');
const generateBtn = document.getElementById('generateBtn');
const exportBtn = document.getElementById('exportBtn');

const orgFileInput = document.getElementById('orgFileInput');
let orgData = {}; // OfficeLocation_ID -> MarketSubSector_Name

const statusColors = {
    'Scheduled': 'green',
    'Awaiting approval': 'yellow',
    'Awaiting withdrawl approval': 'orange',
    'Completed': 'blue',
    'In progress': 'purple',
    'Saved': 'gray'
};

let processedData = {}; // dept -> person -> week -> {total: hours, statuses: Set}
let allWeeks = new Set();
let rawData = []; // store the parsed CSV data
let weeks = []; // for export

function getClosestFriday(date, before = false) {
    const d = new Date(date);
    const day = d.getDay(); // 0=Sun, 1=Mon, ..., 5=Fri, 6=Sat
    // For a week ending on Friday, the week runs Saturday-Friday.
    // This calculates the days to add to get to the upcoming Friday.
    const daysToAdd = (5 - day + 7) % 7;
    d.setDate(d.getDate() + daysToAdd);
    return d;
}

fileInput.addEventListener('change', handleFile);
orgFileInput.addEventListener('change', handleOrgFile);
generateBtn.addEventListener('click', generateReport);
exportBtn.addEventListener('click', exportToExcel);

function handleOrgFile(e) {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = function(e) {
        const csv = e.target.result;
        orgData = parseOrgCSV(csv);
        console.log('Org data loaded:', orgData);
    };
    reader.readAsText(file);
}

function parseOrgCSV(csv) {
    const lines = csv.split('\n');
    const headers = parseCSVLine(lines[0]).map(h => h.trim().toUpperCase());
    const mapping = {};
    const idHeader = 'OFFICELOCATION_ID';
    const nameHeader = 'MARKETSUBSECTOR_NAME';
    const idIndex = headers.indexOf(idHeader);
    const nameIndex = headers.indexOf(nameHeader);

    if (idIndex === -1 || nameIndex === -1) {
        alert(`Org CSV must contain '${idHeader}' and '${nameHeader}' columns.`);
        return {};
    }

    for (let i = 1; i < lines.length; i++) {
        if (!lines[i].trim()) continue;
        const values = parseCSVLine(lines[i]);
        if (values.length === headers.length) {
            mapping[values[idIndex]] = values[nameIndex];
        }
    }
    return mapping;
}


function handleFile(e) {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = function(e) {
        const csv = e.target.result;
        rawData = parseCSV(csv);
        processData(rawData);
        controlsDiv.style.display = 'block';
    };
    reader.readAsText(file);
}

function parseCSV(csv) {
    const lines = csv.split('\n');
    const headers = parseCSVLine(lines[0]).map(h => h.trim().toUpperCase());
    console.log('Headers:', headers);
    const data = [];
    for (let i = 1; i < lines.length; i++) {
        if (!lines[i].trim()) continue;
        const values = parseCSVLine(lines[i]);
        if (values.length !== headers.length) {
            console.log('Field count mismatch: headers', headers.length, 'values', values.length, 'row:', values);
            continue; // skip misaligned rows
        }
        const row = {};
        headers.forEach((h, idx) => {
            let val = values[idx] ? values[idx].trim() : '';
            // Try to parse dates for date columns
            if (h === 'ABSENCE START DATE' || h === 'ABSENCE END DATE') {
                // Dates from CSV are ambiguous (no timezone). Parsing them as YYYY-MM-DD
                // can cause them to be interpreted as UTC midnight, which shifts them to the
                // previous day in timezones west of UTC.
                // To fix this, we'll append a time to treat them as local noon.
                const parsedDate = new Date(val + 'T12:00:00');
                if (!isNaN(parsedDate)) {
                    val = parsedDate;
                } else {
                    console.log('Failed to parse date for', h, ':', val);
                }
            }
            row[h] = val;
        });
        data.push(row);
    }
    console.log('Sample parsed row:', data[0]);
    return data;
}

function parseCSVLine(line) {
    const result = [];
    let current = '';
    let inQuotes = false;
    for (let i = 0; i < line.length; i++) {
        const char = line[i];
        if (char === '"') {
            inQuotes = !inQuotes;
        } else if (char === ',' && !inQuotes) {
            result.push(current);
            current = '';
        } else {
            current += char;
        }
    }
    result.push(current);
    return result.map(s => s.trim());
}

function processData(data) {
    // data is array of objects
    // create departments map: MarketSubSector -> dept -> person -> weeks: {week: hours}
    processedData = {};
    allWeeks = new Set();
    let processedCount = 0;
    data.forEach(row => {
        if (row['ABSENCE STATUS'] === 'Withdrawn' || row['ABSENCE STATUS'] === 'Denied') {
            console.log('Skipping row due to status:', row['ABSENCE STATUS']);
            return;
        }
        if (!row['DEPARTMENT'] || !row['DEPARTMENT'].includes('WSP-ENV')) {
            console.log('Skipping row due to department:', row['DEPARTMENT']);
            return;
        }
        const dept = row['DEPARTMENT'];
        const last = row['LAST NAME'] || '';
        const first = row['FIRST NAME'] || '';
        const name = `${last}, ${first}`.trim();
        const manager = row['LINE MANAGER'] || '';
        if (!name) {
            console.log('Skipping row due to empty name');
            return;
        }
        const deptId = dept.substring(0, 8);
        const marketSubSector = orgData[deptId] || 'Unknown';

        if (!processedData[marketSubSector]) {
            processedData[marketSubSector] = {};
        }
        if (!processedData[marketSubSector][dept]) {
            processedData[marketSubSector][dept] = {
                persons: {}
            };
        }
        if (!processedData[marketSubSector][dept].persons[name]) {
            processedData[marketSubSector][dept].persons[name] = {
                weeks: {},
                manager: manager
            };
        }
        const start = row['ABSENCE START DATE'];
        const end = row['ABSENCE END DATE'];
        if (!start || !end) {
            console.log('Skipping row due to missing dates');
            return;
        }
        const startDate = new Date(start);
        const endDate = new Date(end);
        if (isNaN(startDate) || isNaN(endDate)) {
            console.log('Skipping row due to invalid dates:', start, end);
            return;
        }
        const normal = parseFloat(row['NORMAL WORKING HOURS']) || 40; // assume 40 if not
        const normalPerDay = normal / 5;
        const uom = row['UOM'] || 'H';

        // get all days from start to end
        const days = [];
        const current = new Date(start);
        while (current <= endDate) {
            days.push(new Date(current));
            current.setDate(current.getDate() + 1);
        }

        const workingDaysInAbsence = days.filter(d => d.getDay() >= 1 && d.getDay() <= 5);

        if (workingDaysInAbsence.length === 0) {
            console.log('Skipping absence with no working days:', name, start, end);
            return;
        }

        const absenceDur = parseFloat(row['ABSENCE DURATION']) || 0;
        let adjustedAbsenceDur = absenceDur;
        if (uom === 'D') {
            adjustedAbsenceDur *= normalPerDay;
        }

        // Evenly distribute the total absence duration across the working days.
        // This is a simplifying assumption because the source data's START/END_DATE_DURATION
        // might be associated with non-working days. This approach is safer.
        const hoursPerWorkDay = adjustedAbsenceDur / workingDaysInAbsence.length;

        const absenceType = row['ABSENCE TYPE'] || '';
        const absenceComment = row['ABSENCE COMMENT'] || '';

        workingDaysInAbsence.forEach(day => {
            const week = getWeekNumber(day);
            allWeeks.add(week);
            if (!processedData[marketSubSector][dept].persons[name].weeks[week]) {
                processedData[marketSubSector][dept].persons[name].weeks[week] = { total: 0, statuses: new Set(), dates: {} };
            }
            processedData[marketSubSector][dept].persons[name].weeks[week].total += hoursPerWorkDay;
            processedData[marketSubSector][dept].persons[name].weeks[week].statuses.add(row['ABSENCE STATUS']);
            const dateKey = day.toISOString().split('T')[0];
            if (!processedData[marketSubSector][dept].persons[name].weeks[week].dates[dateKey]) {
                processedData[marketSubSector][dept].persons[name].weeks[week].dates[dateKey] = { total: 0, details: [] };
            }
            processedData[marketSubSector][dept].persons[name].weeks[week].dates[dateKey].total += hoursPerWorkDay;
            processedData[marketSubSector][dept].persons[name].weeks[week].dates[dateKey].details.push({
                type: absenceType,
                comment: absenceComment,
                hours: hoursPerWorkDay
            });
        });

        processedCount++;
    });
    console.log('Processed ' + processedCount + ' absences');
    console.log('Market Sub Sectors:', Object.keys(processedData));
    console.log('All weeks:', Array.from(allWeeks).sort());
}

function generateReport() {
    const startDate = new Date(startDateInput.value);
    if (!startDate) {
        alert('Please select start date');
        return;
    }
    processData(rawData);
    console.log('Selected start: ' + startDateInput.value);
    // Get the closest Friday on or after start date
    const startFriday = getClosestFriday(startDate, false);
    // Generate 52 weeks starting from startFriday
    weeks = [];
    for (let i = 0; i < 52; i++) {
        const endDate = new Date(startFriday);
        endDate.setDate(startFriday.getDate() + i * 7);
        const weekNum = getWeekNumber(endDate);
        weeks.push({ endDate, weekNum });
    }
    console.log('weeks:', weeks);
    // Group weeks by month
    const monthGroups = {};
    weeks.forEach((item, idx) => {
        const month = item.endDate.getMonth();
        const year = item.endDate.getFullYear();
        let monthName = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'][month];
        if (month === 4 && item.endDate.getDate() === 1) monthName = 'Apr';
        const m = `${monthName} ${year}`;
        if (!monthGroups[m]) monthGroups[m] = [];
        monthGroups[m].push({ week: item.weekNum, endDate: item.endDate, idx, absolute: item.weekNum, relative: idx + 1 });
    });
    // Sort weeks within each month by idx
    Object.keys(monthGroups).forEach(m => {
        monthGroups[m].sort((a, b) => a.idx - b.idx);
    });
    console.log('Weeks:', weeks);
    console.log('Month groups:', monthGroups);
    if (weeks.length === 0) {
        reportDiv.innerHTML = '<p>No data for the selected date range.</p>';
        return;
    }
    // Build legend
    let legendHtml = '<div id="legend"><strong>Status Legend:</strong> ';
    Object.keys(statusColors).forEach(s => {
        legendHtml += `<span style="margin-right:10px;">${s}: <span style="display:inline-block; width:8px; height:8px; border-radius:50%; background-color:${statusColors[s]}; vertical-align:middle;"></span></span>`;
    });
    legendHtml += '</div>';
    // Build table
    let html = legendHtml + '<table><thead>';
<<<<<<< HEAD
    html += '<tr><th rowspan="2">Market Sector</th><th rowspan="2">Department</th><th rowspan="2">Person</th>';
=======
    html += '<tr><th rowspan="2">Market Sub Sector / Department / Person</th>';
>>>>>>> 80f132e3bc1f0d55bd9e1c0c59b0854cf7fae7b5
    Object.keys(monthGroups).forEach(m => {
        html += `<th colspan="${monthGroups[m].length}">${m}</th>`;
    });
    html += '<th rowspan="2">Total</th></tr>';
    html += '<tr>';
    Object.keys(monthGroups).forEach(m => {
        monthGroups[m].forEach(item => {
            console.log('item.relative:', item.relative, 'item.week:', item.week);
            const suffix = item.endDate.getFullYear() === 2025 ? ' (-1)' : '';
            const mm = (item.endDate.getMonth() + 1).toString().padStart(2, '0');
            const dd = item.endDate.getDate().toString().padStart(2, '0');
            html += `<th>Week ${item.absolute}${suffix}<br>(${mm}/${dd})</th>`;
        });
    });
    html += '</tr></thead><tbody>';
    // Group by Market Sub Sector
    const overallTotals = {};
    weeks.forEach(item => overallTotals[item.weekNum] = { total: 0, statuses: new Set(), dates: {} });
    let overallTotal = 0;
    const marketSubSectors = Object.keys(processedData).sort();
    marketSubSectors.forEach(marketSubSector => {
        const marketId = marketSubSector.replace(/\s+/g, '_').replace(/[^a-zA-Z0-9_]/g, '');
        // Market Sub Sector total row
        const marketTotals = {};
        weeks.forEach(item => marketTotals[item.weekNum] = { total: 0, statuses: new Set(), dates: {} });
        let marketTotal = 0;

        const depts = Object.keys(processedData[marketSubSector]).sort();
        depts.forEach(dept => {
            const persons = Object.keys(processedData[marketSubSector][dept].persons).sort();
            persons.forEach(person => {
                weeks.forEach(item => {
                    const data = processedData[marketSubSector][dept].persons[person].weeks[item.weekNum] || { total: 0, statuses: new Set(), dates: {} };
                    marketTotals[item.weekNum].total += data.total;
                    marketTotals[item.weekNum].statuses = new Set([...marketTotals[item.weekNum].statuses, ...data.statuses]);
                    for (const [date, dataObj] of Object.entries(data.dates)) {
                        if (!marketTotals[item.weekNum].dates[date]) marketTotals[item.weekNum].dates[date] = { total: 0, details: [] };
                        marketTotals[item.weekNum].dates[date].total += dataObj.total;
                        marketTotals[item.weekNum].dates[date].details.push(...dataObj.details);
                    }
                    marketTotal += data.total;
                });
            });
        });

        html += `<tr class="market-header" onclick="toggleMarket('${marketId}')"><td><strong>${marketSubSector}</strong> <span id="arrow-market-${marketId}">▼</span></td>`;
        weeks.forEach(item => {
            const total = marketTotals[item.weekNum].total;
            html += `<td class="${total === 0 ? 'zero' : ''}"><strong>${total.toFixed(2)}</strong></td>`;
        });
        html += `<td class="${marketTotal === 0 ? 'zero' : ''}"><strong>${marketTotal.toFixed(2)}</strong></td></tr>`;

        depts.forEach(dept => {
            const deptId = dept.replace(/\s+/g, '_').replace(/[^a-zA-Z0-9_]/g, '');
            // Department total row
            const deptTotals = {};
            weeks.forEach(item => deptTotals[item.weekNum] = { total: 0, statuses: new Set(), dates: {} });
            let deptTotal = 0;
            const persons = Object.keys(processedData[marketSubSector][dept].persons).sort();
            persons.forEach(person => {
                weeks.forEach(item => {
                    const data = processedData[marketSubSector][dept].persons[person].weeks[item.weekNum] || { total: 0, statuses: new Set(), dates: {} };
                    deptTotals[item.weekNum].total += data.total;
                    deptTotals[item.weekNum].statuses = new Set([...deptTotals[item.weekNum].statuses, ...data.statuses]);
                    for (const [date, dataObj] of Object.entries(data.dates)) {
                        if (!deptTotals[item.weekNum].dates[date]) deptTotals[item.weekNum].dates[date] = { total: 0, details: [] };
                        deptTotals[item.weekNum].dates[date].total += dataObj.total;
                        deptTotals[item.weekNum].dates[date].details.push(...dataObj.details);
                    }
                    deptTotal += data.total;
                });
            });
            weeks.forEach(item => {
                overallTotals[item.weekNum].total += deptTotals[item.weekNum].total;
                overallTotals[item.weekNum].statuses = new Set([...overallTotals[item.weekNum].statuses, ...deptTotals[item.weekNum].statuses]);
                for (const [date, dataObj] of Object.entries(deptTotals[item.weekNum].dates)) {
                    if (!overallTotals[item.weekNum].dates[date]) overallTotals[item.weekNum].dates[date] = { total: 0, details: [] };
                    overallTotals[item.weekNum].dates[date].total += dataObj.total;
                    overallTotals[item.weekNum].dates[date].details.push(...dataObj.details);
                }
                overallTotal += deptTotals[item.weekNum].total;
            });
<<<<<<< HEAD
            html += `<tr class="dept-header" onclick="toggleDept('${deptId}')"><td>${sector}</td><td><strong>${subdept}</strong> <span id="arrow-${deptId}">▼</span></td><td></td>`;
=======
            html += `<tr class="dept-header ${marketId}" onclick="toggleDept('${deptId}')"><td><strong>${dept}</strong> <span id="arrow-dept-${deptId}">▼</span></td>`;
>>>>>>> 80f132e3bc1f0d55bd9e1c0c59b0854cf7fae7b5
            weeks.forEach(item => {
                const total = deptTotals[item.weekNum].total;
                const circles = Array.from(deptTotals[item.weekNum].statuses).map(s => `<span style="display:inline-block; width:8px; height:8px; border-radius:50%; background-color:${statusColors[s] || 'black'}; margin-left:2px;"></span>`).join('');
                const tooltip = Object.entries(deptTotals[item.weekNum].dates).map(([date, data]) => {
                    let str = `${date}: ${data.total.toFixed(2)}h`;
                    if (data.details.length > 0) {
                        str += ' ' + data.details.map(d => `${d.type}: ${d.comment} (${d.hours.toFixed(2)}h)`).join(', ');
                    }
                    return str;
                }).join('\n');
                html += `<td class="${total === 0 ? 'zero' : ''}" data-tooltip="${tooltip}"><strong>${total.toFixed(2)}${circles}</strong></td>`;
            });
            html += `<td class="${deptTotal === 0 ? 'zero' : ''}"><strong>${deptTotal.toFixed(2)}</strong></td></tr>`;
            // Individual persons
            persons.forEach(person => {
                let personTotal = 0;
                weeks.forEach(item => {
                    const data = processedData[marketSubSector][dept].persons[person].weeks[item.weekNum] || { total: 0, statuses: new Set(), dates: {} };
                    personTotal += data.total;
                });
                if (personTotal > 0) {
<<<<<<< HEAD
                    const manager = processedData[sector][subdept][person].manager;
                    html += `<tr class="person-row ${deptId}" style="display: table-row;"><td>${sector}</td><td>${subdept}</td><td data-tooltip="Manager: ${manager}">${person}</td>`;
=======
                    const manager = processedData[marketSubSector][dept].persons[person].manager;
                    html += `<tr class="person-row ${marketId} ${deptId}" style="display: table-row;"><td data-tooltip="Manager: ${manager}">${person}</td>`;
>>>>>>> 80f132e3bc1f0d55bd9e1c0c59b0854cf7fae7b5
                    weeks.forEach(item => {
                        const data = processedData[marketSubSector][dept].persons[person].weeks[item.weekNum] || { total: 0, statuses: new Set(), dates: {} };
                        const hours = data.total;
                        const circles = Array.from(data.statuses).map(s => `<span style="display:inline-block; width:8px; height:8px; border-radius:50%; background-color:${statusColors[s] || 'black'}; margin-left:2px;"></span>`).join('');
                        const tooltip = Object.entries(data.dates).map(([date, dataObj]) => {
                            let str = `${date}: ${dataObj.total.toFixed(2)}h`;
                            if (dataObj.details.length > 0) {
                                str += ' ' + dataObj.details.map(d => `${d.type}: ${d.comment} (${d.hours.toFixed(2)}h)`).join(', ');
                            }
                            return str;
                        }).join('\n');
                        html += `<td class="${hours === 0 ? 'zero' : ''}" data-tooltip="${tooltip}">${hours.toFixed(2)}${circles}</td>`;
                    });
                    html += `<td class="${personTotal === 0 ? 'zero' : ''}">${personTotal.toFixed(2)}</td></tr>`;
                }
            });
        });
    });
<<<<<<< HEAD
    html += `<tr class="overall-total"><td colspan="3"><strong>Overall Total</strong></td>`;
=======
    html += `<tr class="overall-total"><td><strong>Overall Total</strong></td>`;
>>>>>>> 80f132e3bc1f0d55bd9e1c0c59b0854cf7fae7b5
    weeks.forEach(item => {
        const total = overallTotals[item.weekNum].total;
        const circles = Array.from(overallTotals[item.weekNum].statuses).map(s => `<span style="display:inline-block; width:8px; height:8px; border-radius:50%; background-color:${statusColors[s] || 'black'}; margin-left:2px;"></span>`).join('');
        const tooltip = Object.entries(overallTotals[item.weekNum].dates).map(([date, data]) => {
            let str = `${date}: ${data.total.toFixed(2)}h`;
            if (data.details.length > 0) {
                str += ' ' + data.details.map(d => `${d.type}: ${d.comment} (${d.hours.toFixed(2)}h)`).join(', ');
            }
            return str;
        }).join('\n');
        html += `<td class="${total === 0 ? 'zero' : ''}" data-tooltip="${tooltip}"><strong>${total.toFixed(2)}${circles}</strong></td>`;
    });
    html += `<td class="${overallTotal === 0 ? 'zero' : ''}"><strong>${overallTotal.toFixed(2)}</strong></td></tr>`;
    html += '</tbody></table>';
    html += '<br><button id="longWeekendBtn">Show Long Weekend Impact (Feb 13,15-17)</button>';
    reportDiv.innerHTML = html;
    exportBtn.style.display = 'inline-block';
    document.getElementById('longWeekendBtn').addEventListener('click', generateLongWeekendReport);

    // Setup custom tooltips
    if (!document.getElementById('tooltip')) {
        const tooltipDiv = document.createElement('div');
        tooltipDiv.id = 'tooltip';
        tooltipDiv.className = 'custom-tooltip';
        document.body.appendChild(tooltipDiv);
    }
    const cells = reportDiv.querySelectorAll('td[data-tooltip]');
    cells.forEach(cell => {
        cell.addEventListener('mouseenter', showTooltip);
        cell.addEventListener('mouseleave', hideTooltip);
    });
}

function generateLongWeekendReport() {
    const longWeekendDates = [
        new Date(2026, 1, 13), // Feb 13
        new Date(2026, 1, 15), // Feb 15
        new Date(2026, 1, 16), // Feb 16
        new Date(2026, 1, 17)  // Feb 17
    ];
    
    // Filter rawData to only include absences that overlap with these dates
    const filteredData = rawData.filter(row => {
        const startDate = new Date(row['ABSENCE START DATE']);
        const endDate = new Date(row['ABSENCE END DATE']);
        
        console.log('Checking row:', row['EMPLOYEE NAME'], 'start:', startDate, 'end:', endDate);
        
        // Check if any of the long weekend dates fall within the absence period
        const overlaps = longWeekendDates.some(date => {
            const result = date >= startDate && date <= endDate;
            console.log('  Checking date:', date, 'result:', result);
            return result;
        });
        
        console.log('  Overlaps:', overlaps);
        return overlaps;
    });
    
    console.log('Filtered data length:', filteredData.length);
    
    if (filteredData.length === 0) {
        reportDiv.innerHTML = '<p>No absences found for the long weekend dates.</p>';
        return;
    }
    
    // Process the filtered data
    processData(filteredData);
    
    // Generate report with a fixed start date for the long weekend
    const startDate = new Date(2026, 1, 10); // Feb 10, 2026 - a Monday before the dates
    startDateInput.value = startDate.toISOString().split('T')[0];
    
    generateReport();
    
    // Update the title or add a note
    const titleDiv = document.createElement('div');
    titleDiv.innerHTML = '<h2>Long Weekend Impact Report (Feb 13, 15-17, 2026)</h2>';
    reportDiv.insertBefore(titleDiv, reportDiv.firstChild);
}

const roundToHalf = (value) => Math.round(value * 2) / 2;

function exportToExcel() {
    if (weeks.length === 0) {
        alert('Please generate a report first.');
        return;
    }

    const wb = XLSX.utils.book_new();
    const ws_data = [];
    const row_levels = [];

    // 1. HEADER ROWS
    const monthGroups = {};
    weeks.forEach((item, idx) => {
        const month = item.endDate.getMonth();
        const year = item.endDate.getFullYear();
        let monthName = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'][month];
        if (month === 4 && item.endDate.getDate() === 1) monthName = 'Apr';
        const m = `${monthName} ${year}`;
        if (!monthGroups[m]) monthGroups[m] = 0;
        monthGroups[m]++;
    });

    const header_row_1 = ['', '', '', '']; // Gaps for Market Sub Sector, Dept/Person/Manager
    const merges = [];
    let col_idx = 4; // Start after Market Sub Sector, Dept/Person/Manager
    for (const month in monthGroups) {
        header_row_1.push(month);
        const span = monthGroups[month];
        if (span > 1) {
            merges.push({ s: { r: 0, c: col_idx }, e: { r: 0, c: col_idx + span - 1 } });
        }
        for (let i = 1; i < span; i++) {
            header_row_1.push('');
        }
        col_idx += span;
    }
    header_row_1.push('Total');
    ws_data.push(header_row_1);
    row_levels.push(null);

    const header_row_2 = ['Market Sub Sector', 'Department', 'Person', 'Line Manager'];
    weeks.forEach(item => {
        const suffix = item.endDate.getFullYear() === 2025 ? ' (-1)' : '';
        const mm = (item.endDate.getMonth() + 1).toString().padStart(2, '0');
        const dd = item.endDate.getDate().toString().padStart(2, '0');
        header_row_2.push(`Week ${item.weekNum}${suffix}\n(${mm}/${dd})`);
    });
    header_row_2.push('');
    ws_data.push(header_row_2);
    row_levels.push(null);

    // 2. DATA ROWS
    const marketSubSectors = Object.keys(processedData).sort();
    marketSubSectors.forEach(marketSubSector => {
        const depts = Object.keys(processedData[marketSubSector]).sort();
        depts.forEach(dept => {
            const dept_row = [marketSubSector, dept, '', '']; // Market Sub Sector, Dept name, empty person & manager
            let dept_total_hours = 0;
            weeks.forEach(item => {
                let week_total = 0;
                const persons = Object.keys(processedData[marketSubSector][dept].persons);
                persons.forEach(person => {
                    const data = processedData[marketSubSector][dept].persons[person].weeks[item.weekNum] || { total: 0 };
                    week_total += data.total;
                });
                dept_row.push(roundToHalf(week_total));
                dept_total_hours += week_total;
            });
            dept_row.push(roundToHalf(dept_total_hours));
            ws_data.push(dept_row);
            row_levels.push({ level: 1 });

            const persons = Object.keys(processedData[marketSubSector][dept].persons).sort();
            persons.forEach(person => {
                let person_total = 0;
                weeks.forEach(item => {
                    const data = processedData[marketSubSector][dept].persons[person].weeks[item.weekNum] || { total: 0 };
                    person_total += data.total;
                });

                if (person_total > 0) {
                    const manager = processedData[marketSubSector][dept].persons[person].manager;
                    const person_row = [marketSubSector, dept, person, manager]; // Fill down info
                    weeks.forEach(item => {
                        const data = processedData[marketSubSector][dept].persons[person].weeks[item.weekNum] || { total: 0 };
                        person_row.push(roundToHalf(data.total));
                    });
                    person_row.push(roundToHalf(person_total));
                    ws_data.push(person_row);
                    row_levels.push({ level: 2 });
                }
            });
        });
    });

    // 3. OVERALL TOTAL ROW
    const overall_total_row = ['Overall Total', '', '', ''];
    let overall_grand_total = 0;
    weeks.forEach(item => {
        let week_grand_total = 0;
        marketSubSectors.forEach(marketSubSector => {
            const depts = Object.keys(processedData[marketSubSector]).sort();
            depts.forEach(dept => {
                const persons = Object.keys(processedData[marketSubSector][dept].persons);
                persons.forEach(person => {
                    const data = processedData[marketSubSector][dept].persons[person].weeks[item.weekNum] || { total: 0 };
                    week_grand_total += data.total;
                });
            });
        });
        overall_total_row.push(roundToHalf(week_grand_total));
        overall_grand_total += week_grand_total;
    });
    overall_total_row.push(roundToHalf(overall_grand_total));
    ws_data.push(overall_total_row);
    row_levels.push(null);

    // 4. CREATE SHEET AND ADD TO BOOK
    const ws = XLSX.utils.aoa_to_sheet(ws_data);
    
    if (!ws['!merges']) ws['!merges'] = [];
    ws['!merges'].push(...merges);

    ws['!rows'] = row_levels;

    const cols = Object.keys(header_row_1).length;
    ws['!cols'] = [{ wch: 30 }, { wch: 30 }, { wch: 20 }, { wch: 20 }]; // Market Sub Sector, Dept, Person, and Manager columns
    for(let i=4; i<cols; i++) {
        ws['!cols'].push({ wch: 15 });
    }

    XLSX.utils.book_append_sheet(wb, ws, 'Absence Report');
    XLSX.writeFile(wb, 'absence_report.xlsx');
}

function getWeekNumber(date) {
    // Weeks end on Friday. The first week of 2026 ends on Friday, Jan 2.
    const week1End = new Date(2026, 0, 2);
    const friday = getClosestFriday(date, true);
    const diffMs = friday - week1End;
    // Use Math.round for more robust week calculation against timezone/DST shifts.
    const diffWeeks = Math.round(diffMs / (7 * 24 * 60 * 60 * 1000));
    let week = 1 + diffWeeks;
    if (week <= 0) week += 52;
    return week;
}

function toggleMarket(marketId) {
    const rows = document.querySelectorAll(`.dept-header.${marketId}, .person-row.${marketId}`);
    const arrow = document.getElementById(`arrow-market-${marketId}`);
    let collapsed = false;
    rows.forEach(row => {
        if (row.style.display === 'none') {
            row.style.display = 'table-row';
        } else {
            row.style.display = 'none';
            collapsed = true;
        }
    });
    arrow.textContent = collapsed ? '▶' : '▼';
}

function toggleDept(deptId) {
    const rows = document.querySelectorAll(`.person-row.${deptId}`);
    const arrow = document.getElementById(`arrow-dept-${deptId}`);
    let collapsed = false;
    rows.forEach(row => {
        if (row.style.display === 'none') {
            row.style.display = 'table-row';
        } else {
            row.style.display = 'none';
            collapsed = true;
        }
    });
    arrow.textContent = collapsed ? '▶' : '▼';
}

function getWeekEndDate(weekNum) {
    const week1End = new Date(2026, 0, 3); // Jan 3, 2026 local time
    const endDate = new Date(week1End);
    endDate.setDate(week1End.getDate() + (weekNum - 1) * 7 - (weekNum > 26 ? 364 : 0));
    const mm = (endDate.getMonth() + 1).toString().padStart(2, '0');
    const dd = endDate.getDate().toString().padStart(2, '0');
    const days = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
    const day = days[endDate.getDay()];
    return `${mm}/${dd}`;
}

function showTooltip(e) {
    const tooltip = document.getElementById('tooltip');
    tooltip.textContent = e.target.dataset.tooltip;
    tooltip.style.display = 'block';
    const rect = e.target.getBoundingClientRect();
    const tooltipRect = tooltip.getBoundingClientRect();
    let left = rect.left;
    let top = rect.top - tooltipRect.height - 5;
    if (top < 0) top = rect.bottom + 5;
    if (left + tooltipRect.width > window.innerWidth) left = window.innerWidth - tooltipRect.width - 5;
    if (left < 0) left = 5;
    tooltip.style.left = left + 'px';
    tooltip.style.top = top + 'px';
}

function hideTooltip() {
    document.getElementById('tooltip').style.display = 'none';
}