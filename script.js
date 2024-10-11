document.getElementById('selectFileButton').addEventListener('click', () => {
    document.getElementById('fileInput').click();
});

document.getElementById('fileInput').addEventListener('change', (event) => {
    const file = event.target.files[0];
    if (!file) return;

    const reader = new FileReader();

    reader.onload = (e) => {
        const arrayBuffer = e.target.result;
        const data = new Uint8Array(arrayBuffer);
    
        // Use xlsx-style to read the workbook
        const workbook = XLSX.read(data, { type: 'array' });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
    
        const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
        const output = processSchedule(rows);
        
        // Add new title row with font size 13, center-aligned, and merged columns
        const titleRow = ['REI Daily Schedule'];
        output.unshift(titleRow);
    
        const newWorkbook = XLSX.utils.book_new();
        const newWorksheet = XLSX.utils.aoa_to_sheet(output);
        XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, 'Break Schedule');
    
        // Set column widths
        newWorksheet['!cols'] = [
            { wpx: 40 }, { wpx: 120 }, { wpx: 150 }, { wpx: 100 }, { wpx: 80 }, { wpx: 80 }, { wpx: 80 }, { wpx: 80 }
        ];
    
        // Define the default font style (Arial, 10pt)
        const defaultFont = { font: { name: 'Arial', sz: 10 } };

        // Define the header style
        const headerStyle = {
            font: { bold: true, name: 'Arial', sz: 10 },
            fill: { fgColor: { rgb: 'CCFFCC' } },
            border: {
                top: { style: 'thin', color: { rgb: 'AAAAAA' } },
                bottom: { style: 'thin', color: { rgb: 'AAAAAA' } },
                left: { style: 'thin', color: { rgb: 'AAAAAA' } },
                right: { style: 'thin', color: { rgb: 'AAAAAA' } }
            }
        };
    
        // Define title row style (font size 12, center alignment)
        const titleStyle = {
            font: { bold: true, name: 'Arial', sz: 13 },
            alignment: { horizontal: 'center', vertical: 'center' },
            fill: { fgColor: { rgb: 'CCFFCC' } },
            border: {
                top: { style: 'thin', color: { rgb: 'AAAAAA' } },
                bottom: { style: 'thin', color: { rgb: 'AAAAAA' } },
                left: { style: 'thin', color: { rgb: 'AAAAAA' } },
                right: { style: 'thin', color: { rgb: 'AAAAAA' } }
            }
        };
    
        // Apply default font style to all cells
        for (const cell in newWorksheet) {
            if (cell[0] === '!') continue; // Skip special properties like '!ref'
            if (!newWorksheet[cell].s) {
                newWorksheet[cell].s = {};
            }
            newWorksheet[cell].s.font = defaultFont.font;
        }

        // Apply titleStyle to the first row
        const titleCellAddress = XLSX.utils.encode_cell({ r: 0, c: 0 });
        newWorksheet[titleCellAddress] = { t: 's', v: titleRow[0], s: titleStyle };
    
        // Merge the first row (title)
        newWorksheet['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: output[1].length - 1 } }];
    
        // Apply headerStyle to the second row (now the header)
        output[1].forEach((cell, colIndex) => {
            const cellAddress = XLSX.utils.encode_cell({ r: 1, c: colIndex });
            newWorksheet[cellAddress].s = headerStyle;
        });
    
        // Apply headerStyle and merge department rows
        output.forEach((row, index) => {
            if (index > 1 && row[0]) {
                for (let col = 0; col < row.length; col++) {
                    const cellAddress = XLSX.utils.encode_cell({ r: index, c: col });
                    newWorksheet[cellAddress].s = headerStyle;
                }
                const mergeRange = {
                    s: { r: index, c: 0 },
                    e: { r: index, c: row.length - 1 }
                };
                newWorksheet['!merges'].push(mergeRange);
            }
        });
    
        const today = new Date();
        const formattedDate = today.toLocaleDateString('en-US').replace(/\//g, '-');
        const workbookBlob = XLSX.write(newWorkbook, { bookType: 'xlsx', type: 'binary' });
        const blob = new Blob([s2ab(workbookBlob)], { type: 'application/octet-stream' });
        const url = URL.createObjectURL(blob);
    
        const link = document.getElementById('downloadLink');
        const footer = document.getElementById('downloadLinkFooter');
        link.href = url;
        link.download = `Break Schedule ${formattedDate}.xlsx`;
        link.classList.remove('d-none');
        footer.classList.remove('d-none');
        link.style.display = 'inline-block';
        document.getElementById('footer').style.display = 'block';
    };
    
    reader.readAsArrayBuffer(file);
});

// Function to process the schedule and return output data
function processSchedule(rows) {
    let dept, schedule = [];

    // Loop through each row starting from index 7
    for (let i = 7; i < rows.length; i++) {
        // If the first column contains a department name, assign it to 'dept' and skip further processing for that row
        if (rows[i][0]) {
            dept = rows[i][0];
            continue;
        }

        // Split the time interval (column 4) into start and end times, and convert them to minutes
        let interval = rows[i][4]?.split('-') || [];
        interval = interval.map(bound => timeToMinutes(bound));

        // Add each row to the schedule with the department, job, name, and interval
        schedule.push({
            dept: dept,
            job: rows[i][1],
            name: formatName(rows[i][2]),
            interval: interval
        });
    }

    // Calculate the earliest and latest times for each person's shift
    let shifts = {};
    schedule.forEach(row => {
        if (!(row.name in shifts)) shifts[row.name] = [1440, 0]; // Default shift interval (1440 = 24:00)
        shifts[row.name][0] = Math.min(shifts[row.name][0], row.interval[0]); // Earliest start time
        shifts[row.name][1] = Math.max(shifts[row.name][1], row.interval[1]); // Latest end time
    });

    // Generate breaks based on shift duration
    let breaks = {};
    for (let name in shifts) {
        breaks[name] = [shifts[name][0] + 120]; // First break at 2 hours
        let shiftDuration = shifts[name][1] - shifts[name][0];

        // Add breaks at 4 hours and 6 hours if the shift is long enough
        if (shiftDuration >= 300) breaks[name].push(shifts[name][0] + 240);
        if (shiftDuration >= 390) breaks[name].push(shifts[name][0] + 360);

        // Adjust breaks if they overlap with scheduled intervals
        for (let j = 0; j < breaks[name].length; j++) {
            let breakDuration = j % 2 ? 30 : 15; // 30 minutes for second break, 15 minutes for the first
            let timesDelayed = 0;

            // Resolve conflicts with other schedule intervals
            resolveLoop: while (true) {
                let currentDept = null;

                // Check if the break overlaps with the employee's scheduled intervals
                for (let k = 0; k < schedule.length; k++) {
                    if (schedule[k].name !== name) continue;
                    if (breaks[name][j] < schedule[k].interval[0] || breaks[name][j] + breakDuration > schedule[k].interval[1]) continue;
                    currentDept = schedule[k].dept;
                    break;
                }

                // Skip break adjustment if the employee is not in certain departments
                if (!["Hardgoods", "Frontline", "Softgoods"].includes(currentDept)) break;

                // Adjust breaks to avoid overlap with others in the same department
                for (let k = 0; k < schedule.length; k++) {
                    if (schedule[k].name === name || schedule[k].dept !== currentDept) continue;
                    if (!(schedule[k].name in breaks)) continue;

                    for (let l = 0; l < breaks[schedule[k].name].length; l++) {
                        let breakDuration2 = l % 2 ? 30 : 15;

                        // Check if breaks overlap
                        if (breaks[name][j] + breakDuration > breaks[schedule[k].name][l] &&
                            breaks[name][j] < breaks[schedule[k].name][l] + breakDuration2) {

                            let originalBreak = shifts[name][0] + 120 * (j + 1); // Original break time
                            timesDelayed++;

                            // Adjust break by delaying or advancing it by 15-minute intervals
                            breaks[name][j] = originalBreak - (15 * Math.ceil(timesDelayed / 2) * Math.pow(-1, timesDelayed));

                            // If break has been delayed too many times, log an error and stop adjusting
                            if (timesDelayed > 6) {
                                console.error(`ERROR: ${name} in ${currentDept} at ${minutesToTime(breaks[name][j])}`);
                                break resolveLoop;
                            }

                            continue resolveLoop; // Continue adjusting breaks
                        }
                    }
                }
                break resolveLoop; // Exit loop if no more adjustments are needed
            }
        }
    }

    // Extract special breaks for 'Misc Events'
    let news_breaks = {};
    schedule.forEach(row => {
        if (row.job === 'Misc Events') news_breaks[row.name] = row.interval[0];
    });

    // Remove duplicate rows for 'Mgmt Retail' department, keeping only unique entries by name
    let uniqueRows = [];
    let seen = new Set();
    schedule.forEach(row => {
        if (row.dept === 'Mgmt Retail' && row.job === 'Management') {
            const key = row.name;
            if (!seen.has(key)) {
                seen.add(key);
                uniqueRows.push(row);
            }
        } else {
            uniqueRows.push(row);
        }
    });
    schedule = uniqueRows;

    // Update shift intervals for management jobs to cover the entire shift
    schedule.forEach(row => {
        if (row.job === 'Management') {
            row.interval = shifts[row.name];
        }
    });

    // Remove rows related to 'Misc Events' or 'Meeting'
    schedule = schedule.filter(row => row.job !== 'Misc Events' && row.job !== 'Meeting');

    // Create the output array for Excel
    const output = [];
    output.push(['Dept.', 'Job', 'Name', 'Shift', 'News Break', 'Rest Break', 'Meal Break', 'Rest Break']);

    // Track printed breaks to avoid duplicate entries
    let breaksPrinted = new Set();
    schedule.forEach(row => {
        if (row.dept !== dept) {
            dept = row.dept;
            output.push([dept, '', '', '', '', '', '', '']); // Add department header row
        }
        let intervalStr = row.interval.map(bound => minutesToTime(bound)).join('-'); // Format shift interval
        let rowOutput = ['', row.job, row.name, intervalStr];

        const key = row.name;
        if (!breaksPrinted.has(key)) {
            breaksPrinted.add(key);
            // Add special breaks or empty cells for the row
            if (news_breaks[row.name]) rowOutput.push(minutesToTime(news_breaks[row.name]));
            else rowOutput.push('');

            // Add regular breaks (rest, meal, etc.)
            for (let i = 0; i < 3; i++) {
                if (breaks[row.name][i]) rowOutput.push(minutesToTime(breaks[row.name][i]));
                else rowOutput.push('');
            }
        }

        output.push(rowOutput);
    });

    return output; // Return the processed schedule for further use
}

// Helper function to convert string to ArrayBuffer
function s2ab(s) {
    const buf = new ArrayBuffer(s.length);
    const view = new Uint8Array(buf);
    for (let i = 0; i < s.length; i++) {
        view[i] = s.charCodeAt(i) & 0xFF;
    }
    return buf;
}

// Helper function to convert time strings to minutes since midnight (e.g. "9:00AM" -> 540)
function timeToMinutes(time) {
    return ((+time.split(':')[0] % 12) + (time.includes('PM') ? 12 : 0)) * 60 + +time.split(':')[1].slice(0, 2);
}

// Helper function to convert minutes since midnight to time strings (e.g. 540 -> "9:00AM")
function minutesToTime(minutes) {
    return `${(minutes / 60 + 11) % 12 + 1 | 0}:${(minutes % 60).toString().padStart(2, '0')}${minutes < 720 ? 'AM' : 'PM'}`;
}

// Helper function to properly capitalize a word in a name (e.g., "o'neil" -> "O'Neil", "mcdonald" -> "McDonald")
function formatName(name) {
    return name.split(' ').map(word => {
        let result = '';
        let capitalizeNext = true;
        
        for (let i = 0; i < word.length; i++) {
            let char = word.charAt(i);
            
            // Handle common prefixes like 'Mc', 'Mac', 'O', 'D'
            if (i > 0) {
                if (word.startsWith('Mc') && i === 2 && /[a-zA-Z]/.test(char)) {
                    result += char.toUpperCase();
                    capitalizeNext = false;
                    continue;
                }
                if (word.startsWith('Mac') && i === 3 && /[a-zA-Z]/.test(char)) {
                    result += char.toUpperCase();
                    capitalizeNext = false;
                    continue;
                }
                if (word.startsWith("O'") && i === 2 && /[a-zA-Z]/.test(char)) {
                    result += char.toUpperCase();
                    capitalizeNext = false;
                    continue;
                }
                if (word.startsWith("D'") && i === 2 && /[a-zA-Z]/.test(char)) {
                    result += char.toUpperCase();
                    capitalizeNext = false;
                    continue;
                }
                if (word.startsWith("Al-") && i === 3 && /[a-zA-Z]/.test(char)) {
                    result += char.toUpperCase();
                    capitalizeNext = false;
                    continue;
                }
                if (word.startsWith("El-") && i === 3 && /[a-zA-Z]/.test(char)) {
                    result += char.toUpperCase();
                    capitalizeNext = false;
                    continue;
                }
                if (word.startsWith("St.") && i === 3 && /[a-zA-Z]/.test(char)) {
                    result += char.toUpperCase();
                    capitalizeNext = false;
                    continue;
                }
            }

            // Capitalize the current character if needed and reset capitalizeNext
            if (capitalizeNext && /[a-zA-Z]/.test(char)) {
                result += char.toUpperCase();
                capitalizeNext = false;
            } else {
                result += char.toLowerCase();
            }

            // After spaces, hyphens, or apostrophes, we need to capitalize the next character
            if (char === '-' || char === "'" || char === " ") {
                capitalizeNext = true;
            }
        }
        return result;
    }).join(' ');
}