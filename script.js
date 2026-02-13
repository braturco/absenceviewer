const fileInput = document.getElementById('fileInput');
const reportDiv = document.getElementById('report');
const controlsDiv = document.getElementById('controls');
const startDateInput = document.getElementById('startDate');
const endDateInput = document.getElementById('endDate');
const generateBtn = document.getElementById('generateBtn');
const exportBtn = document.getElementById('exportBtn');

const orgFileInput = document.getElementById('orgFileInput');
const holidayFileInput = document.getElementById('holidayFileInput');
let orgData = {}; // OfficeLocation_ID -> MarketSubSector_Name
let holidayData = {}; // weekNum -> [{date, holiday, appliesTo}, ...]

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
holidayFileInput.addEventListener('change', handleHolidayFile);
generateBtn.addEventListener('click', generateReport);
exportBtn.addEventListener('click', exportToExcel);

function handleOrgFile(e) {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = function(e) {
        const csv = e.target.result;
        orgData = parseOrgCSV(csv);
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

function handleHolidayFile(e) {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = function(e) {
        const csv = e.target.result;
        holidayData = parseHolidayCSV(csv);

        // Show loaded holidays with detailed week calculations
        console.log('\n=== HOLIDAY DATA LOADED ===');
        console.log('Total weeks with holidays:', Object.keys(holidayData).length);
        console.log('\nDetailed holiday data:');
        Object.keys(holidayData).sort((a, b) => parseInt(a) - parseInt(b)).forEach(weekNum => {
            console.log(`\nWeek ${weekNum}:`);
            holidayData[weekNum].forEach(h => {
                console.log(`  - Date: ${h.date}, Holiday: ${h.holiday}, Applies To: ${h.appliesTo}`);
            });
        });
        console.log('===========================\n');
    };
    reader.readAsText(file);
}

function parseHolidayCSV(csv) {
    const lines = csv.split('\n');
    const headers = parseCSVLine(lines[0]).map(h => h.trim().toUpperCase());
    const mapping = {};
    const dateHeader = 'DATE';
    const holidayHeader = 'HOLIDAY';
    const appliesToHeader = 'APPLIES-TO';
    const dateIndex = headers.indexOf(dateHeader);
    const holidayIndex = headers.indexOf(holidayHeader);
    const appliesToIndex = headers.indexOf(appliesToHeader);

    console.log('Parsing holiday CSV...');
    console.log('Headers found:', headers);

    if (dateIndex === -1 || holidayIndex === -1 || appliesToIndex === -1) {
        alert(`Holiday CSV must contain 'Date', 'Holiday', and 'Applies-To' columns.`);
        return {};
    }

    for (let i = 1; i < lines.length; i++) {
        if (!lines[i].trim()) continue;
        const values = parseCSVLine(lines[i]);
        if (values.length >= headers.length) {
            const dateStr = values[dateIndex].trim();
            const holiday = values[holidayIndex].trim();
            const appliesTo = values[appliesToIndex].trim();

            // Parse date and get week number
            const date = new Date(dateStr + 'T12:00:00');
            if (!isNaN(date)) {
                const weekNum = getWeekNumber(date);
                console.log(`  Holiday "${holiday}" on ${dateStr} -> Week ${weekNum}`);
                if (!mapping[weekNum]) {
                    mapping[weekNum] = [];
                }
                mapping[weekNum].push({
                    date: dateStr,
                    holiday: holiday,
                    appliesTo: appliesTo
                });
            } else {
                console.log(`  Failed to parse date: ${dateStr}`);
            }
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
    const data = [];
    for (let i = 1; i < lines.length; i++) {
        if (!lines[i].trim()) continue;
        const values = parseCSVLine(lines[i]);
        if (values.length !== headers.length) {
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
                }
            }
            row[h] = val;
        });
        data.push(row);
    }
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

function escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}

function processData(data) {
    // data is array of objects
    // create departments map: MarketSubSector -> dept -> person -> weeks: {week: hours}
    processedData = {};
    allWeeks = new Set();
    let processedCount = 0;
    data.forEach(row => {
        if (row['ABSENCE STATUS'] === 'Withdrawn' || row['ABSENCE STATUS'] === 'Denied') {
            return;
        }
        if (!row['DEPARTMENT'] || !row['DEPARTMENT'].includes('-ENV-')) {
            return;
        }
        const dept = row['DEPARTMENT'];
        const last = row['LAST NAME'] || '';
        const first = row['FIRST NAME'] || '';
        const name = `${last}, ${first}`.trim();
        const manager = row['LINE MANAGER'] || '';
        if (!name) {
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
            return;
        }
        const startDate = new Date(start);
        const endDate = new Date(end);
        if (isNaN(startDate) || isNaN(endDate)) {
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
}

function generateReport() {
    const startDate = new Date(startDateInput.value);
    if (!startDate) {
        alert('Please select start date');
        return;
    }
    processData(rawData);
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

    console.log('\n=== REPORT GENERATION ===');
    console.log('Holiday data available:', Object.keys(holidayData).length > 0);
    console.log('Week numbers in report:', weeks.map(w => w.weekNum).slice(0, 10).join(', '), '...');
    console.log('=========================\n');

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
    let html = '<div style="margin-bottom: 10px;"><button id="expandAllBtn">Expand All</button> <button id="collapseAllBtn">Collapse All</button></div>';
    html += legendHtml + '<table><thead>';
    html += '<tr><th rowspan="2">Market Sub Sector</th><th rowspan="2">Department</th><th rowspan="2">Person</th>';
    Object.keys(monthGroups).forEach(m => {
        html += `<th colspan="${monthGroups[m].length}">${m}</th>`;
    });
    html += '<th rowspan="2">Total</th></tr>';
    html += '<tr>';
    Object.keys(monthGroups).forEach(m => {
        monthGroups[m].forEach(item => {
            const suffix = item.endDate.getFullYear() === 2025 ? ' (-1)' : '';
            const mm = (item.endDate.getMonth() + 1).toString().padStart(2, '0');
            const dd = item.endDate.getDate().toString().padStart(2, '0');

            // Check if this week has holidays
            const hasHoliday = holidayData[item.weekNum] && holidayData[item.weekNum].length > 0;
            const holidayClass = hasHoliday ? ' holiday-week' : '';
            let holidayTooltip = '';
            if (hasHoliday) {
                holidayTooltip = holidayData[item.weekNum].map(h =>
                    `${h.date}: ${h.holiday} (${h.appliesTo})`
                ).join('|||');
                console.log(`Week ${item.weekNum} HAS HOLIDAY - applying class and tooltip:`, holidayTooltip);
            }

            // Escape the tooltip text for HTML attribute
            const escapedTooltip = escapeHtml(holidayTooltip);
            html += `<th class="week-header${holidayClass}" data-tooltip="${escapedTooltip}">Week ${item.absolute}${suffix}<br>(${mm}/${dd})</th>`;
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

        html += `<tr class="market-header" onclick="toggleMarket('${marketId}')"><td colspan="2"><strong>${marketSubSector}</strong> <span id="arrow-market-${marketId}">▼</span></td><td></td>`;
        weeks.forEach(item => {
            const total = marketTotals[item.weekNum].total;
            const hasHoliday = holidayData[item.weekNum] && holidayData[item.weekNum].length > 0;
            const holidayClass = hasHoliday ? ' holiday-week-cell' : '';
            html += `<td class="${total === 0 ? 'zero' : ''}${holidayClass}"><strong>${total.toFixed(2)}</strong></td>`;
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
            html += `<tr class="dept-header ${marketId}" onclick="toggleDept('${deptId}')"><td></td><td colspan="2"><strong>${dept}</strong> <span id="arrow-dept-${deptId}">▼</span></td>`;
            weeks.forEach(item => {
                const total = deptTotals[item.weekNum].total;
                const hasHoliday = holidayData[item.weekNum] && holidayData[item.weekNum].length > 0;
                const holidayClass = hasHoliday ? ' holiday-week-cell' : '';
                const circles = Array.from(deptTotals[item.weekNum].statuses).map(s => `<span style="display:inline-block; width:8px; height:8px; border-radius:50%; background-color:${statusColors[s] || 'black'}; margin-left:2px;"></span>`).join('');
                const tooltip = Object.entries(deptTotals[item.weekNum].dates).map(([date, data]) => {
                    let str = `${date}: ${data.total.toFixed(2)}h`;
                    if (data.details.length > 0) {
                        str += ' ' + data.details.map(d => `${d.type}: ${d.comment} (${d.hours.toFixed(2)}h)`).join(', ');
                    }
                    return str;
                }).join('\n');
                html += `<td class="${total === 0 ? 'zero' : ''}${holidayClass}" data-tooltip="${tooltip}"><strong>${total.toFixed(2)}${circles}</strong></td>`;
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
                    const manager = processedData[marketSubSector][dept].persons[person].manager;
                    html += `<tr class="person-row ${marketId} ${deptId}" style="display: table-row;"><td></td><td></td><td data-tooltip="Manager: ${manager}">${person}</td>`;
                    weeks.forEach(item => {
                        const data = processedData[marketSubSector][dept].persons[person].weeks[item.weekNum] || { total: 0, statuses: new Set(), dates: {} };
                        const hours = data.total;
                        const hasHoliday = holidayData[item.weekNum] && holidayData[item.weekNum].length > 0;
                        const holidayClass = hasHoliday ? ' holiday-week-cell' : '';
                        const circles = Array.from(data.statuses).map(s => `<span style="display:inline-block; width:8px; height:8px; border-radius:50%; background-color:${statusColors[s] || 'black'}; margin-left:2px;"></span>`).join('');
                        const tooltip = Object.entries(data.dates).map(([date, dataObj]) => {
                            let str = `${date}: ${dataObj.total.toFixed(2)}h`;
                            if (dataObj.details.length > 0) {
                                str += ' ' + dataObj.details.map(d => `${d.type}: ${d.comment} (${d.hours.toFixed(2)}h)`).join(', ');
                            }
                            return str;
                        }).join('\n');
                        html += `<td class="${hours === 0 ? 'zero' : ''}${holidayClass}" data-tooltip="${tooltip}">${hours.toFixed(2)}${circles}</td>`;
                    });
                    html += `<td class="${personTotal === 0 ? 'zero' : ''}">${personTotal.toFixed(2)}</td></tr>`;
                }
            });
        });
    });
    html += `<tr class="overall-total"><td colspan="3"><strong>Overall Total</strong></td>`;
    weeks.forEach(item => {
        const total = overallTotals[item.weekNum].total;
        const hasHoliday = holidayData[item.weekNum] && holidayData[item.weekNum].length > 0;
        const holidayClass = hasHoliday ? ' holiday-week-cell' : '';
        const circles = Array.from(overallTotals[item.weekNum].statuses).map(s => `<span style="display:inline-block; width:8px; height:8px; border-radius:50%; background-color:${statusColors[s] || 'black'}; margin-left:2px;"></span>`).join('');
        const tooltip = Object.entries(overallTotals[item.weekNum].dates).map(([date, data]) => {
            let str = `${date}: ${data.total.toFixed(2)}h`;
            if (data.details.length > 0) {
                str += ' ' + data.details.map(d => `${d.type}: ${d.comment} (${d.hours.toFixed(2)}h)`).join(', ');
            }
            return str;
        }).join('\n');
        html += `<td class="${total === 0 ? 'zero' : ''}${holidayClass}" data-tooltip="${tooltip}"><strong>${total.toFixed(2)}${circles}</strong></td>`;
    });
    html += `<td class="${overallTotal === 0 ? 'zero' : ''}"><strong>${overallTotal.toFixed(2)}</strong></td></tr>`;
    html += '</tbody></table>';
    reportDiv.innerHTML = html;
    exportBtn.style.display = 'inline-block';

    // Setup custom tooltips
    if (!document.getElementById('tooltip')) {
        const tooltipDiv = document.createElement('div');
        tooltipDiv.id = 'tooltip';
        tooltipDiv.className = 'custom-tooltip';
        document.body.appendChild(tooltipDiv);
    }
    const cells = reportDiv.querySelectorAll('td[data-tooltip], th[data-tooltip]');
    const thCells = reportDiv.querySelectorAll('th[data-tooltip]');
    const tdCells = reportDiv.querySelectorAll('td[data-tooltip]');

    console.log('Setting up tooltips:');
    console.log('  Total cells with data-tooltip:', cells.length);
    console.log('  TH elements with data-tooltip:', thCells.length);
    console.log('  TD elements with data-tooltip:', tdCells.length);

    // Log a sample TH element to see its data-tooltip value
    if (thCells.length > 0) {
        console.log('  Sample TH data-tooltip:', thCells[0].dataset.tooltip);
        console.log('  Sample TH text:', thCells[0].textContent.substring(0, 20));
    }

    cells.forEach(cell => {
        cell.addEventListener('mouseenter', showTooltip);
        cell.addEventListener('mouseleave', hideTooltip);
    });

    // Setup expand/collapse all buttons
    document.getElementById('expandAllBtn').addEventListener('click', expandAll);
    document.getElementById('collapseAllBtn').addEventListener('click', collapseAll);
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

        // Check if any of the long weekend dates fall within the absence period
        const overlaps = longWeekendDates.some(date => {
            const result = date >= startDate && date <= endDate;
            return result;
        });

        return overlaps;
    });

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

async function exportToExcel() {
    if (weeks.length === 0) {
        alert('Please generate a report first.');
        return;
    }

    const workbook = new ExcelJS.Workbook();

    // Define styles
    const headerStyle = {
        fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF2F2F2' } },
        font: { bold: true },
        alignment: { horizontal: 'center', vertical: 'middle', wrapText: true }
    };
    const deptStyle = {
        fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD0D0D0' } },
        font: { bold: true }
    };
    const personEvenStyle = {
        fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF0F0F0' } }
    };
    const personOddStyle = {
        fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFFFF' } }
    };
    const totalStyle = {
        fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF888888' } },
        font: { bold: true, color: { argb: 'FFFFFFFF' } }
    };

    // Helper function to create a sheet for a market sector
    function createMarketSheet(marketSubSector, depts) {
        const sheetName = marketSubSector.substring(0, 31).replace(/[:\\\/\?\*\[\]]/g, '_');
        const worksheet = workbook.addWorksheet(sheetName);

        // 1. BUILD MONTH GROUPS
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

        // 2. ADD HEADER ROW 1 (Month headers)
        const header_row_1 = ['', '', '', ''];
        let col_idx = 5; // Start at column 5 (E) after Dept/Manager/Person/empty
        for (const month in monthGroups) {
            header_row_1.push(month);
            const span = monthGroups[month];
            for (let i = 1; i < span; i++) {
                header_row_1.push('');
            }
        }
        header_row_1.push('Total');
        worksheet.addRow(header_row_1);

        // Merge month header cells
        col_idx = 5;
        for (const month in monthGroups) {
            const span = monthGroups[month];
            if (span > 1) {
                worksheet.mergeCells(1, col_idx, 1, col_idx + span - 1);
            }
            col_idx += span;
        }

        // 3. ADD HEADER ROW 2 (Column names)
        const header_row_2 = ['Department', 'Line Manager', 'Person', ''];
        weeks.forEach(item => {
            const suffix = item.endDate.getFullYear() === 2025 ? ' (-1)' : '';
            const mm = (item.endDate.getMonth() + 1).toString().padStart(2, '0');
            const dd = item.endDate.getDate().toString().padStart(2, '0');
            header_row_2.push(`Week ${item.weekNum}${suffix}\n(${mm}/${dd})`);
        });
        header_row_2.push('');
        worksheet.addRow(header_row_2);

        // 4. ADD DATA ROWS
        depts.forEach(dept => {
            // Department row
            const dept_row = [dept, '', '', ''];
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
            worksheet.addRow(dept_row);

            // Person rows
            const persons = Object.keys(processedData[marketSubSector][dept].persons).sort();
            persons.forEach(person => {
                let person_total = 0;
                weeks.forEach(item => {
                    const data = processedData[marketSubSector][dept].persons[person].weeks[item.weekNum] || { total: 0 };
                    person_total += data.total;
                });

                if (person_total > 0) {
                    const manager = processedData[marketSubSector][dept].persons[person].manager;
                    const person_row = [dept, manager, person, ''];
                    weeks.forEach(item => {
                        const data = processedData[marketSubSector][dept].persons[person].weeks[item.weekNum] || { total: 0 };
                        person_row.push(roundToHalf(data.total));
                    });
                    person_row.push(roundToHalf(person_total));
                    worksheet.addRow(person_row);
                }
            });
        });

        // 5. ADD TOTAL ROW
        const total_row = ['Total', '', '', ''];
        let grand_total = 0;
        weeks.forEach(item => {
            let week_total = 0;
            depts.forEach(dept => {
                const persons = Object.keys(processedData[marketSubSector][dept].persons);
                persons.forEach(person => {
                    const data = processedData[marketSubSector][dept].persons[person].weeks[item.weekNum] || { total: 0 };
                    week_total += data.total;
                });
            });
            total_row.push(roundToHalf(week_total));
            grand_total += week_total;
        });
        total_row.push(roundToHalf(grand_total));
        worksheet.addRow(total_row);

        // 6. SET COLUMN WIDTHS
        worksheet.getColumn(1).width = 35; // Department
        worksheet.getColumn(2).width = 25; // Line Manager
        worksheet.getColumn(3).width = 25; // Person
        worksheet.getColumn(4).width = 5;  // Empty
        for (let i = 5; i <= 4 + weeks.length + 1; i++) {
            worksheet.getColumn(i).width = 12; // Week columns
        }

        // 7. APPLY STYLES
        worksheet.eachRow((row, rowNumber) => {
            row.eachCell((cell, colNumber) => {
                // Header rows
                if (rowNumber === 1 || rowNumber === 2) {
                    cell.style = headerStyle;
                }
                // Dept rows (check if manager and person are empty)
                else if (rowNumber > 2 && row.getCell(2).value === '' && row.getCell(3).value === '') {
                    cell.style = deptStyle;
                }
                // Total row (last row)
                else if (rowNumber === worksheet.rowCount) {
                    cell.style = totalStyle;
                }
                // Person rows (alternating)
                else if (rowNumber > 2) {
                    cell.style = (rowNumber % 2 === 0) ? personEvenStyle : personOddStyle;
                }
            });
        });

        // 8. ADD AUTO-FILTER TO ROW 2
        const lastCol = worksheet.columnCount;
        worksheet.autoFilter = {
            from: { row: 2, column: 1 },
            to: { row: 2, column: lastCol }
        };

        return worksheet;
    }

    // Create a sheet for each Market Sub Sector
    const marketSubSectors = Object.keys(processedData).sort();
    marketSubSectors.forEach(marketSubSector => {
        const depts = Object.keys(processedData[marketSubSector]).sort();
        createMarketSheet(marketSubSector, depts);
    });

    // Write file
    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'absence_report.xlsx';
    a.click();
    URL.revokeObjectURL(url);
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
    const deptRows = document.querySelectorAll(`.dept-header.${marketId}`);
    const personRows = document.querySelectorAll(`.person-row.${marketId}`);
    const arrow = document.getElementById(`arrow-market-${marketId}`);
    let collapsed = false;

    // Check current state
    if (deptRows.length > 0 && deptRows[0].style.display === 'none') {
        // Expanding - show dept headers
        deptRows.forEach(row => {
            row.style.display = 'table-row';
        });
        // For person rows, only show if their department is expanded
        personRows.forEach(row => {
            const deptClasses = Array.from(row.classList).filter(c => c !== 'person-row' && c !== marketId);
            if (deptClasses.length > 0) {
                const deptId = deptClasses[0];
                const deptArrow = document.getElementById(`arrow-dept-${deptId}`);
                if (deptArrow && deptArrow.textContent === '▼') {
                    row.style.display = 'table-row';
                }
            }
        });
    } else {
        // Collapsing - hide everything
        deptRows.forEach(row => {
            row.style.display = 'none';
        });
        personRows.forEach(row => {
            row.style.display = 'none';
        });
        collapsed = true;
    }

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
    const tooltipText = e.target.dataset.tooltip;

    // Only log debug info for TH elements (week headers), not TD (employee/dept cells)
    if (e.target.tagName === 'TH') {
        console.log('Showing tooltip for TH element:', e.target.textContent.substring(0, 20));
        console.log('  Tooltip data-tooltip value:', tooltipText);
        console.log('  Tooltip text length:', tooltipText ? tooltipText.length : 0);
    }

    // Convert ||| delimiter back to newlines for display
    const displayText = tooltipText ? tooltipText.replace(/\|\|\|/g, '\n') : '';
    tooltip.textContent = displayText;
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

function expandAll() {
    // Show all department headers and market headers
    const deptHeaders = document.querySelectorAll('.dept-header');
    deptHeaders.forEach(row => {
        row.style.display = 'table-row';
    });

    // Show all person rows
    const personRows = document.querySelectorAll('.person-row');
    personRows.forEach(row => {
        row.style.display = 'table-row';
    });

    // Set all arrows to expanded
    const allArrows = document.querySelectorAll('[id^="arrow-"]');
    allArrows.forEach(arrow => {
        arrow.textContent = '▼';
    });
}

function collapseAll() {
    const allRows = document.querySelectorAll('.dept-header, .person-row');
    allRows.forEach(row => {
        row.style.display = 'none';
    });
    const allArrows = document.querySelectorAll('[id^="arrow-market-"]');
    allArrows.forEach(arrow => {
        arrow.textContent = '▶';
    });
}