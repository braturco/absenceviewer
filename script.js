const fileInput = document.getElementById('fileInput');
const reportDiv = document.getElementById('report');

fileInput.addEventListener('change', handleFile);

function handleFile(e) {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = function(e) {
        const csv = e.target.result;
        const json = parseCSV(csv);
        processData(json);
    };
    reader.readAsText(file);
}

function parseCSV(csv) {
    const lines = csv.split('\n');
    const headers = lines[0].split(',').map(h => h.trim());
    const data = [];
    for (let i = 1; i < lines.length; i++) {
        if (!lines[i].trim()) continue;
        const values = lines[i].split(',');
        const row = {};
        headers.forEach((h, idx) => {
            let val = values[idx] ? values[idx].trim() : '';
            // Handle dates if they look like dates
            if (val.match(/^\d{1,2}\/\d{1,2}\/\d{4}$/)) {
                val = new Date(val);
            }
            row[h] = val;
        });
        data.push(row);
    }
    return data;
}

function processData(data) {
    // data is array of objects
    // create persons map: name -> weeks: {week: hours}
    const persons = {};
    data.forEach(row => {
        if (!row['Department'] || !row['Department'].includes('WSP-ENV')) return;
        const last = row['Last Name'] || '';
        const first = row['First Name'] || '';
        const name = `${last}, ${first}`.trim();
        if (!name) return;
        const start = row['Absence start date'];
        const end = row['Absence end date'];
        const startDur = parseFloat(row['Start_Date_Duration']) || 0;
        const endDur = parseFloat(row['End_Date_Duration']) || 0;
        const normal = parseFloat(row['Normal Working Hours']) || 40; // assume 40 if not
        const normalPerDay = normal / 5;
        if (!start || !end) return;
        // get days
        const days = [];
        const current = new Date(start);
        const endDate = new Date(end);
        while (current <= endDate) {
            days.push(new Date(current));
            current.setDate(current.getDate() + 1);
        }
        // assign hours
        days.forEach((day, i) => {
            let hours = normalPerDay;
            if (i === 0) hours = startDur;
            else if (i === days.length - 1) hours = endDur;
            const week = getWeekNumber(day);
            if (!persons[name]) persons[name] = {};
            if (!persons[name][week]) persons[name][week] = 0;
            persons[name][week] += hours;
        });
    });
    // now, get all weeks
    const allWeeks = new Set();
    Object.values(persons).forEach(p => Object.keys(p).forEach(w => allWeeks.add(w)));
    const weeks = Array.from(allWeeks).sort((a,b)=>a-b);
    // persons sorted
    const names = Object.keys(persons).sort();
    // build table
    let html = '<table><thead><tr><th>Person</th>';
    weeks.forEach(w => html += `<th>Week ${w}</th>`);
    html += '</tr></thead><tbody>';
    names.forEach(name => {
        html += `<tr><td>${name}</td>`;
        weeks.forEach(w => {
            const hours = persons[name][w] || 0;
            html += `<td>${hours.toFixed(2)}</td>`;
        });
        html += '</tr>';
    });
    html += '</tbody></table>';
    reportDiv.innerHTML = html;
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