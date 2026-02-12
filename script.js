const fileInput = document.getElementById('fileInput');
const reportDiv = document.getElementById('report');
const controlsDiv = document.getElementById('controls');
const startDateInput = document.getElementById('startDate');
const endDateInput = document.getElementById('endDate');
const generateBtn = document.getElementById('generateBtn');
const exportBtn = document.getElementById('exportBtn');

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
    let daysToAdd;
    if (before) {
        // closest Friday on or before
        daysToAdd = (5 - day) % 7;
        if (daysToAdd === 0 && day !== 5) daysToAdd = -7; // if not Friday, go back
    } else {
        // on or after
        daysToAdd = (5 - day + 7) % 7;
    }
    d.setDate(d.getDate() + daysToAdd);
    return d;
}

fileInput.addEventListener('change', handleFile);
generateBtn.addEventListener('click', generateReport);
exportBtn.addEventListener('click', exportToExcel);

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
                const parsedDate = new Date(val);
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
    // create departments map: dept -> person -> weeks: {week: hours}
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
        if (!name) {
            console.log('Skipping row due to empty name');
            return;
        }
        if (!processedData[dept]) processedData[dept] = {};
        if (!processedData[dept][name]) processedData[dept][name] = {};
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
        const startDur = parseFloat(row['START_DATE_DURATION']) || 0;
        const endDur = parseFloat(row['END_DATE_DURATION']) || 0;
        const normal = parseFloat(row['NORMAL WORKING HOURS']) || 40; // assume 40 if not
        const normalPerDay = normal / 5;
        const uom = row['UOM'] || 'H';
        let startDurHours = startDur;
        if (startDur === 0) {
            startDurHours = normalPerDay;
        } else if (uom === 'D') {
            startDurHours *= normalPerDay;
        }
        let endDurHours = endDur;
        if (endDur === 0) {
            // Will calculate later
        } else if (uom === 'D') {
            endDurHours *= normalPerDay;
        }
        // get all days from start to end
        const days = [];
        const current = new Date(start);
        while (current <= endDate) {
            days.push(new Date(current));
            current.setDate(current.getDate() + 1);
        }
        if (days.length === 0) {
            console.log('No days for absence:', name, start, end);
            return;
        }
        const totalDays = days.length;
        const middleDays = days.slice(1, -1); // days between start and end
        const middleWorkingDays = middleDays.filter(d => {
            const dow = d.getDay();
            return dow >= 1 && dow <= 5; // Mon-Fri
        });
        const numMiddleWorking = middleWorkingDays.length;
        const absenceDur = parseFloat(row['ABSENCE DURATION']) || 0;
        let adjustedAbsenceDur = absenceDur;
        if (uom === 'D') {
            adjustedAbsenceDur *= normalPerDay;
        }
        // For simplicity, assume daily rate is normalPerDay, and total matches Absence Duration
        const daily = normalPerDay;
        // Calculate total so far without end
        const totalSoFar = startDurHours + numMiddleWorking * daily;
        if (endDur === 0) {
            endDurHours = adjustedAbsenceDur - totalSoFar;
        }
        // assign hours
        const firstDay = days[0];
        const lastDay = days[totalDays - 1];

        // Hours for first day
        let hoursFirst = startDurHours;
        // Hours for last day
        let hoursLast = endDur === 0 ? 0 : endDurHours;
        // Total hours to distribute to middle
        let totalMiddleHours = adjustedAbsenceDur - hoursFirst - hoursLast;
        let hoursPerMiddleDay = 0;
        if (numMiddleWorking > 0) {
            hoursPerMiddleDay = totalMiddleHours / numMiddleWorking;
        } else {
            // If no middle days, adjust last day
            if (endDur === 0) {
                hoursLast = adjustedAbsenceDur - hoursFirst;
            }
        }

        // Get absence details
        const absenceType = row['ABSENCE TYPE'] || '';
        const absenceComment = row['ABSENCE COMMENT'] || '';

        // Assign to first day
        const weekFirst = getWeekNumber(firstDay);
        allWeeks.add(weekFirst);
        if (!processedData[dept][name][weekFirst]) processedData[dept][name][weekFirst] = { total: 0, statuses: new Set(), dates: {} };
        processedData[dept][name][weekFirst].total += hoursFirst;
        processedData[dept][name][weekFirst].statuses.add(row['ABSENCE STATUS']);
        const dateKeyFirst = firstDay.toISOString().split('T')[0];
        if (!processedData[dept][name][weekFirst].dates[dateKeyFirst]) processedData[dept][name][weekFirst].dates[dateKeyFirst] = { total: 0, details: [] };
        processedData[dept][name][weekFirst].dates[dateKeyFirst].total += hoursFirst;
        processedData[dept][name][weekFirst].dates[dateKeyFirst].details.push({ type: absenceType, comment: absenceComment, hours: hoursFirst });

        // Assign to middle days
        middleWorkingDays.forEach(day => {
            const week = getWeekNumber(day);
            allWeeks.add(week);
            if (!processedData[dept][name][week]) processedData[dept][name][week] = { total: 0, statuses: new Set(), dates: {} };
            processedData[dept][name][week].total += hoursPerMiddleDay;
            processedData[dept][name][week].statuses.add(row['ABSENCE STATUS']);
            const dateKey = day.toISOString().split('T')[0];
            if (!processedData[dept][name][week].dates[dateKey]) processedData[dept][name][week].dates[dateKey] = { total: 0, details: [] };
            processedData[dept][name][week].dates[dateKey].total += hoursPerMiddleDay;
            processedData[dept][name][week].dates[dateKey].details.push({ type: absenceType, comment: absenceComment, hours: hoursPerMiddleDay });
        });

        // Assign to last day
        if (totalDays > 1) {
            if (endDur === 0) {
                hoursLast = adjustedAbsenceDur - hoursFirst - numMiddleWorking * hoursPerMiddleDay;
            }
            const weekLast = getWeekNumber(lastDay);
            allWeeks.add(weekLast);
            if (!processedData[dept][name][weekLast]) processedData[dept][name][weekLast] = { total: 0, statuses: new Set(), dates: {} };
            processedData[dept][name][weekLast].total += hoursLast;
            processedData[dept][name][weekLast].statuses.add(row['ABSENCE STATUS']);
            const dateKeyLast = lastDay.toISOString().split('T')[0];
            if (!processedData[dept][name][weekLast].dates[dateKeyLast]) processedData[dept][name][weekLast].dates[dateKeyLast] = { total: 0, details: [] };
            processedData[dept][name][weekLast].dates[dateKeyLast].total += hoursLast;
            processedData[dept][name][weekLast].dates[dateKeyLast].details.push({ type: absenceType, comment: absenceComment, hours: hoursLast });
        }
        processedCount++;
    });
    console.log('Processed ' + processedCount + ' absences');
    console.log('Departments:', Object.keys(processedData));
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
    const weeks = [];
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
    html += '<tr><th rowspan="2">Department / Person</th>';
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
    // Group by department
    const overallTotals = {};
    weeks.forEach(item => overallTotals[item.weekNum] = { total: 0, statuses: new Set(), dates: {} });
    let overallTotal = 0;
    const depts = Object.keys(processedData).sort();
    depts.forEach(dept => {
        const deptId = dept.replace(/\s+/g, '_').replace(/[^a-zA-Z0-9_]/g, '');
        // Department total row
        const deptTotals = {};
        weeks.forEach(item => deptTotals[item.weekNum] = { total: 0, statuses: new Set(), dates: {} });
        let deptTotal = 0;
        const persons = Object.keys(processedData[dept]).sort();
        persons.forEach(person => {
            weeks.forEach(item => {
                const data = processedData[dept][person][item.weekNum] || { total: 0, statuses: new Set(), dates: {} };
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
        html += `<tr class="dept-header" onclick="toggleDept('${deptId}')"><td><strong>${dept}</strong> <span id="arrow-${deptId}">▼</span></td>`;
        weeks.forEach(item => {
            const total = deptTotals[item.weekNum].total;
            const circles = Array.from(deptTotals[item.weekNum].statuses).map(s => `<span style="display:inline-block; width:8px; height:8px; border-radius:50%; background-color:${statusColors[s] || 'black'}; margin-left:2px;"></span>`).join('');
            const tooltip = Object.entries(deptTotals[item.weekNum].dates).map(([date, data]) => {
                let str = `${date}: ${data.total.toFixed(2)}h`;
                if (data.details.length > 0) {
                    str += '\n' + data.details.map(d => `  - ${d.type}: ${d.comment} (${d.hours.toFixed(2)}h)`).join('\n');
                }
                return str;
            }).join('\n');
            html += `<td class="${total === 0 ? 'zero' : ''}" title="${tooltip}"><strong>${total.toFixed(2)}${circles}</strong></td>`;
        });
        html += `<td class="${deptTotal === 0 ? 'zero' : ''}"><strong>${deptTotal.toFixed(2)}</strong></td></tr>`;
        // Individual persons
        persons.forEach(person => {
            let personTotal = 0;
            weeks.forEach(item => {
                const data = processedData[dept][person][item.weekNum] || { total: 0, statuses: new Set(), dates: {} };
                personTotal += data.total;
            });
            if (personTotal > 0) {
                html += `<tr class="person-row ${deptId}" style="display: table-row;"><td>${person}</td>`;
                weeks.forEach(item => {
                    const data = processedData[dept][person][item.weekNum] || { total: 0, statuses: new Set(), dates: {} };
                    const hours = data.total;
                    const circles = Array.from(data.statuses).map(s => `<span style="display:inline-block; width:8px; height:8px; border-radius:50%; background-color:${statusColors[s] || 'black'}; margin-left:2px;"></span>`).join('');
                    const tooltip = Object.entries(data.dates).map(([date, dataObj]) => {
                        let str = `${date}: ${dataObj.total.toFixed(2)}h`;
                        if (dataObj.details.length > 0) {
                            str += '\n' + dataObj.details.map(d => `  - ${d.type}: ${d.comment} (${d.hours.toFixed(2)}h)`).join('\n');
                        }
                        return str;
                    }).join('\n');
                    html += `<td class="${hours === 0 ? 'zero' : ''}" title="${tooltip}">${hours.toFixed(2)}${circles}</td>`;
                });
                html += `<td class="${personTotal === 0 ? 'zero' : ''}">${personTotal.toFixed(2)}</td></tr>`;
            }
        });
    });
    html += `<tr class="overall-total"><td><strong>Overall Total</strong></td>`;
    weeks.forEach(item => {
        const total = overallTotals[item.weekNum].total;
        const circles = Array.from(overallTotals[item.weekNum].statuses).map(s => `<span style="display:inline-block; width:8px; height:8px; border-radius:50%; background-color:${statusColors[s] || 'black'}; margin-left:2px;"></span>`).join('');
        const tooltip = Object.entries(overallTotals[item.weekNum].dates).map(([date, data]) => {
            let str = `${date}: ${data.total.toFixed(2)}h`;
            if (data.details.length > 0) {
                str += '\n' + data.details.map(d => `  - ${d.type}: ${d.comment} (${d.hours.toFixed(2)}h)`).join('\n');
            }
            return str;
        }).join('\n');
        html += `<td class="${total === 0 ? 'zero' : ''}" title="${tooltip}"><strong>${total.toFixed(2)}${circles}</strong></td>`;
    });
    html += `<td class="${overallTotal === 0 ? 'zero' : ''}"><strong>${overallTotal.toFixed(2)}</strong></td></tr>`;
    html += '</tbody></table>';
    html += '<br><button id="longWeekendBtn">Show Long Weekend Impact (Feb 13,15-17)</button>';
    reportDiv.innerHTML = html;
    exportBtn.style.display = 'inline-block';
    document.getElementById('longWeekendBtn').addEventListener('click', generateLongWeekendReport);
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

function exportToExcel() {
    const table = document.querySelector('#report table');
    if (!table) return;
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.table_to_sheet(table);
    ws['!rows'] = [];
    let rowIndex = 1; // skip header
    // Overall total - no level
    rowIndex++;
    const depts = Object.keys(processedData).sort();
    depts.forEach(dept => {
        // dept header
        ws['!rows'][rowIndex] = { level: 0 };
        rowIndex++;
        // persons
        const persons = Object.keys(processedData[dept]).filter(person => {
            let total = 0;
            weeks.forEach(w => total += (processedData[dept][person][w] || {total:0}).total);
            return total > 0;
        });
        persons.forEach(() => {
            ws['!rows'][rowIndex] = { level: 1 };
            rowIndex++;
        });
    });
    XLSX.utils.book_append_sheet(wb, ws, 'Absence Report');
    XLSX.writeFile(wb, 'absence_report.xlsx');
}

function getClosestFriday(date, before = false) {
    const d = new Date(date);
    const day = d.getDay(); // 0=Sun, 1=Mon, ..., 5=Fri, 6=Sat
    let daysToAdd;
    if (before) {
        // closest Friday on or before
        daysToAdd = (5 - day) % 7;
        if (daysToAdd === 0 && day !== 5) daysToAdd = -7; // if not Friday, go back
    } else {
        // on or after
        daysToAdd = (5 - day + 7) % 7;
    }
    d.setDate(d.getDate() + daysToAdd);
    return d;
}

function getWeekNumber(date) {
    // Weeks end on the date
    const week1End = new Date(2026, 0, 3); // Jan 3, 2026 local time
    const friday = getClosestFriday(date, false);
    const diffMs = friday - week1End;
    const diffWeeks = Math.ceil(diffMs / (7 * 24 * 60 * 60 * 1000));
    let week = 1 + diffWeeks;
    if (week <= 0) week += 52;
    return week;
}

function toggleDept(deptId) {
    const rows = document.querySelectorAll(`.person-row.${deptId}`);
    const arrow = document.getElementById(`arrow-${deptId}`);
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