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
        const workbook = XLSX.read(data, { type: 'array' });

        // Get the first sheet (assuming there is only one sheet)
        const sheetName = workbook.SheetNames[0];

        // Convert the sheet to a JSON array
        const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1 });

        // Process the schedule and get output
        const output = processSchedule(rows);

        // Create a new workbook and worksheet
        const newWorkbook = XLSX.utils.book_new();
        const worksheet = XLSX.utils.aoa_to_sheet(output);
        XLSX.utils.book_append_sheet(newWorkbook, worksheet, 'Break Schedule');

        // Hardcoded styles
        // Set column widths
        worksheet['!cols'] = [
            { wpx: 150 }, // Department
            { wpx: 120 }, // Job
            { wpx: 120 }, // Name
            { wpx: 100 }, // Shift
            { wpx: 100 }, // Rest Break
            { wpx: 100 }, // Meal Break
            { wpx: 100 }, // Rest Break
        ];

        // Apply font and other styles
        for (let col = 0; col < output[0].length; col++) {
            const cellAddress = XLSX.utils.encode_cell({ c: col, r: 0 });
            if (!worksheet[cellAddress]) worksheet[cellAddress] = {}; // Initialize cell if it doesn't exist
            worksheet[cellAddress].s = {
                font: {
                    name: 'Arial',
                    bold: true,
                    sz: 12,
                    color: { rgb: 'FFFFFF' } // White font color
                },
                fill: {
                    fgColor: { rgb: "000000" } // Black background
                }
            };
        }

        // Apply styles for department rows and set bold font
        for (let row = 1; row < output.length; row++) {
            if (output[row][0]) { // Check if there is a department in the first column
                for (let col = 0; col < output[row].length; col++) {
                    const cellAddress = XLSX.utils.encode_cell({ c: col, r: row });
                    if (!worksheet[cellAddress]) worksheet[cellAddress] = {}; // Initialize cell if it doesn't exist
                    worksheet[cellAddress].s = {
                        font: {
                            name: 'Arial',
                            bold: true,
                            sz: 12
                        },
                        fill: {
                            patternType: 'solid',
                            fgColor: { rgb: '00FF00' } // Green background for department rows
                        }
                    };
                }
            }
        }

        // Generate a Blob for the workbook
        const today = new Date();
        const options = { year: 'numeric', month: '2-digit', day: '2-digit' };
        const formattedDate = today.toLocaleDateString('en-US', options).replace(/\//g, '-'); // Format as MM-DD-YYYY
        
        // Create a Blob for the file
        const workbookBlob = XLSX.write(newWorkbook, { bookType: 'xlsx', type: 'binary' });
        const blob = new Blob([s2ab(workbookBlob)], { type: 'application/octet-stream' });
        const url = URL.createObjectURL(blob);
        
        const link = document.getElementById('downloadLink');
        const footer = document.getElementById('downloadLinkFooter');
        link.href = url;
        link.download = `Break Schedule ${formattedDate}.xlsx`; // Set the download filename
        link.classList.remove('d-none'); // Ensure `d-none` is removed
        footer.classList.remove('d-none'); // Ensure `d-none` is removed
        link.style.display = 'inline-block'; // Make sure it has a visible display style

        // Show the footer after the download link is available
        document.getElementById('footer').style.display = 'block';
    };

    reader.readAsArrayBuffer(file);
});

// Function to process the schedule and return output data
function processSchedule(rows) {
    let dept, schedule = [];
    for (let i = 7; i < rows.length; i++) {
        if (rows[i][0]) {
            dept = rows[i][0];
            continue;
        }

        let interval = rows[i][4]?.split('-') || [];
        interval = interval.map(bound => timeToMinutes(bound));
        schedule.push({ dept: dept, job: rows[i][1], name: rows[i][2], interval: interval });
    }

    let shifts = {};
    schedule.forEach(row => {
        if (!(row.name in shifts)) shifts[row.name] = [1440, 0];
        shifts[row.name][0] = Math.min(shifts[row.name][0], row.interval[0]);
        shifts[row.name][1] = Math.max(shifts[row.name][1], row.interval[1]);
    });

    let breaks = {};
    for (let i in shifts) {
        breaks[i] = [shifts[i][0] + 120];
        let shiftDuration = shifts[i][1] - shifts[i][0];

        if (shiftDuration >= 300) breaks[i].push(shifts[i][0] + 240);
        if (shiftDuration >= 390) breaks[i].push(shifts[i][0] + 360);

        for (let j = 0; j < breaks[i].length; j++) {
            let breakDuration = j % 2 ? 30 : 15;
            let timesDelayed = 0;

            resolveLoop: while (true) {
                let dept = null;
                for (let k = 0; k < schedule.length; k++) {
                    if (schedule[k].name !== i) continue;
                    if (breaks[i][j] < schedule[k].interval[0] || breaks[i][j] + breakDuration > schedule[k].interval[1]) continue;
                    dept = schedule[k].dept;
                    break;
                }

                if (!["Hardgoods", "Frontline", "Softgoods"].includes(dept)) break;

                for (let k = 0; k < schedule.length; k++) {
                    if (schedule[k].name === i || schedule[k].dept !== dept) continue;
                    if (!(schedule[k].name in breaks)) continue;

                    for (let l = 0; l < breaks[schedule[k].name].length; l++) {
                        let breakDuration2 = l % 2 ? 30 : 15;

                        if (breaks[i][j] + breakDuration > breaks[schedule[k].name][l] && breaks[i][j] < breaks[schedule[k].name][l] + breakDuration2) {
                            let originalBreak = shifts[i][0] + 120 * (j + 1);
                            timesDelayed++;
                            breaks[i][j] = originalBreak - (15 * Math.ceil(timesDelayed / 2) * Math.pow(-1, timesDelayed));

                            if (timesDelayed > 6) {
                                console.error(`ERROR: ${i} in ${dept} at ${minutesToTime(breaks[i][j])}`);
                                break resolveLoop;
                            }

                            continue resolveLoop;
                        }
                    }
                }
                break resolveLoop;
            }
        }
    }

    schedule = schedule.filter(row => row.job !== 'Misc Events');

    // Create the output array for Excel
    const output = [];
    output.push(['Department', 'Job', 'Name', 'Shift', 'Rest Break', 'Meal Break', 'Rest Break']);
    
    schedule.forEach(row => {
        if (row.dept !== dept) {
            dept = row.dept;
            output.push([dept, '', '', '', '', '', '']);
        }
        let intervalStr = row.interval.map(bound => minutesToTime(bound)).join('-');
        let rowOutput = [``, row.job, row.name, intervalStr];

        for (let i = 0; i < 3; i++) {
            if (breaks[row.name][i]) rowOutput.push(minutesToTime(breaks[row.name][i]));
            else rowOutput.push('');
        }
        output.push(rowOutput);
    });

    return output;
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

function timeToMinutes(time) {
    return ((+time.split(':')[0] % 12) + (time.includes('PM') ? 12 : 0)) * 60 + +time.split(':')[1].slice(0, 2);
}

function minutesToTime(minutes) {
    return `${(minutes / 60 + 11) % 12 + 1 | 0}:${(minutes % 60).toString().padStart(2, '0')}${minutes < 720 ? 'AM' : 'PM'}`;
}
