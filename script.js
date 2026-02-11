const fileInput = document.getElementById('fileInput');
const reportDiv = document.getElementById('report');

fileInput.addEventListener('change', handleFile);

function handleFile(e) {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = function(e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, {type: 'array', cellDates: true});
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const json = XLSX.utils.sheet_to_json(worksheet);
        processData(json);
    };
    reader.readAsArrayBuffer(file);
}

function processData(data) {
    // data is array of objects
    // create persons map: name -> weeks: {week: hours}
    const persons = {};
    data.forEach(row => {
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
    const d = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
    const dayNum = d.getUTCDay() || 7;
    d.setUTCDate(d.getUTCDate() + 4 - dayNum);
    const yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
    return Math.ceil((((d - yearStart) / 86400000) + 1) / 7);
}