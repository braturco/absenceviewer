const fileInput = document.getElementById('fileInput');
const reportDiv = document.getElementById('report');
const controlsDiv = document.getElementById('controls');
const startDateInput = document.getElementById('startDate');
const endDateInput = document.getElementById('endDate');
const generateBtn = document.getElementById('generateBtn');

let processedData = {}; // dept -> person -> week -> hours
let allWeeks = new Set();

fileInput.addEventListener('change', handleFile);
generateBtn.addEventListener('click', generateReport);

function handleFile(e) {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = function(e) {
        const csv = e.target.result;
        const json = parseCSV(csv);
        processData(json);
        controlsDiv.style.display = 'block';
    };
    reader.readAsText(file);
}

function parseCSV(csv) {
    const lines = csv.split('\n');
    const headers = lines[0].split(',').map(h => h.trim().toUpperCase());
    console.log('Headers:', headers);
    const data = [];
    for (let i = 1; i < lines.length; i++) {
        if (!lines[i].trim()) continue;
        const values = lines[i].split(',');
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
    return data;
}

function processData(data) {
    // data is array of objects
    // create departments map: dept -> person -> weeks: {week: hours}
    processedData = {};
    allWeeks = new Set();
    let processedCount = 0;
    data.forEach(row => {
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
        days.forEach((day, i) => {
            let hours = 0;
            if (i === 0) hours = startDurHours;
            else if (i === totalDays - 1) hours = endDurHours;
            else if (middleWorkingDays.some(md => md.getTime() === day.getTime())) hours = daily;
            const week = getWeekNumber(day);
            allWeeks.add(week);
            if (!processedData[dept][name][week]) processedData[dept][name][week] = 0;
            processedData[dept][name][week] += hours;
        });
        processedCount++;
    });
    console.log('Processed ' + processedCount + ' absences');
    console.log('Departments:', Object.keys(processedData));
    console.log('All weeks:', Array.from(allWeeks).sort());
}

function generateReport() {
    const startDate = new Date(startDateInput.value);
    const endDate = new Date(endDateInput.value);
    if (!startDate || !endDate) {
        alert('Please select start and end dates');
        return;
    }
    console.log('Selected range: ' + startDateInput.value + ' to ' + endDateInput.value);
    // Find closest Friday on or after startDate
    const startFriday = getClosestFriday(startDate, false); // on or after
    // Find closest Friday on or after endDate
    const endFriday = getClosestFriday(endDate, false);
    const startWeek = getWeekNumber(startFriday);
    const endWeek = getWeekNumber(endFriday);
    console.log('Week range: ' + startWeek + ' to ' + endWeek);
    // Generate all weeks in range
    const weeks = [];
    for (let w = startWeek; w <= endWeek; w++) {
        weeks.push(w);
    }
    console.log('Filtered weeks:', weeks);
    if (weeks.length === 0) {
        reportDiv.innerHTML = '<p>No data for the selected date range.</p>';
        return;
    }
    // Build table
    let html = '<table><thead><tr><th>Department / Person</th>';
    weeks.forEach(w => html += `<th>Week ${w}</th>`);
    html += '<th>Total</th></tr></thead><tbody>';
    // Group by department
    const depts = Object.keys(processedData).sort();
    depts.forEach(dept => {
        // Department total row
        const deptTotals = {};
        weeks.forEach(w => deptTotals[w] = 0);
        let deptTotal = 0;
        const persons = Object.keys(processedData[dept]).sort();
        persons.forEach(person => {
            weeks.forEach(w => {
                const hours = processedData[dept][person][w] || 0;
                deptTotals[w] += hours;
                deptTotal += hours;
            });
        });
        html += `<tr class="dept-header"><td><strong>${dept}</strong></td>`;
        weeks.forEach(w => html += `<td><strong>${deptTotals[w].toFixed(2)}</strong></td>`);
        html += `<td><strong>${deptTotal.toFixed(2)}</strong></td></tr>`;
        // Individual persons
        persons.forEach(person => {
            let personTotal = 0;
            weeks.forEach(w => {
                personTotal += processedData[dept][person][w] || 0;
            });
            if (personTotal > 0) {
                html += `<tr><td>${person}</td>`;
                weeks.forEach(w => {
                    const hours = processedData[dept][person][w] || 0;
                    html += `<td>${hours.toFixed(2)}</td>`;
                });
                html += `<td>${personTotal.toFixed(2)}</td></tr>`;
            }
        });
    });
    html += '</tbody></table>';
    reportDiv.innerHTML = html;
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
    // Weeks end on Friday
    // Week 1 ends on Dec 26, 2025
    const week1End = new Date('2025-12-26');
    // Get the Friday of the week for this date
    const d = new Date(date);
    const day = d.getDay(); // 0=Sun, 1=Mon, ..., 5=Fri, 6=Sat
    const daysToFriday = (5 - day + 7) % 7;
    const friday = new Date(d);
    friday.setDate(d.getDate() + daysToFriday);
    // Calculate weeks from week1End
    const diffMs = friday - week1End;
    const diffWeeks = Math.floor(diffMs / (7 * 24 * 60 * 60 * 1000));
    return 1 + diffWeeks;
}