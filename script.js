// Trigger file input when the Select File button is clicked
document.getElementById("selectFileButton").addEventListener("click", () => {
    document.getElementById("fileInput").click();
});

// Store the file for later use
let selectedFile;

document.getElementById("fileInput").addEventListener("change", (event) => {
    selectedFile = event.target.files[0]; // Store the selected file
    const link = document.getElementById("downloadLink");
    link.classList.remove("d-none"); // Show the download link
    // Show the download button footer after file selection
    const footer = document.getElementById("downloadLinkFooter");
    footer.classList.remove("d-none");
});

// Trigger processing and downloading when the Download button is clicked
document.getElementById("downloadLink").addEventListener("click", () => {
    if (!selectedFile) {
        alert("Please select a file first.");
        return;
    }

    const reader = new FileReader();

    reader.onload = (event) => {
        // Read data from the selected file into an XLSX workbook
        const workbook = XLSX.read(new Uint8Array(event.target.result), { type: "array" });

        // Convert the first worksheet from the XLSX workbook into JSON row data
        const rowData = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]], { header: 1 });

        // Process and style the data
        const processedRowData = processRowData(rowData);
        const styledRowData = styleRowData(processedRowData);

        // Derive the schedule date from the spreadsheet contents for the file name
        const scheduleDate = extractScheduleDate(rowData);

        const workbookBlob = XLSX.write(styledRowData, { bookType: "xlsx", type: "binary" });
        const blob = new Blob([stringToArrayBuffer(workbookBlob)], { type: "application/octet-stream" });
        const url = URL.createObjectURL(blob);

        // Update the hidden download link with the generated file URL and filename
        const hiddenDownloadLink = document.getElementById("hiddenDownloadLink");
        hiddenDownloadLink.href = url;
        hiddenDownloadLink.download = `Break Schedule ${scheduleDate}.xlsx`;

        // Simulate a click on the hidden download link
        hiddenDownloadLink.click();
    };

    reader.readAsArrayBuffer(selectedFile);
});

// Function to style the schedule (restores your original formatting)
function styleRowData(schedule) {
    const newWorkbook = XLSX.utils.book_new();
    const newWorksheet = XLSX.utils.aoa_to_sheet(schedule);
    XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, "Break Schedule");

    // Set column widths
    newWorksheet["!cols"] = [
        { wpx: 40 },  // Dept
        { wpx: 150 }, // Job
        { wpx: 120 }, // Name
        { wpx: 100 }, // Shift
        { wpx: 80 },  // 15
        { wpx: 80 },  // 30
        { wpx: 80 },  // 15
        { wpx: 80 }   // News
    ];

    // Merge the first row (title)
    let merges = [
        { s: { r: 0, c: 0 }, e: { r: 0, c: 7 } }
    ];

    schedule.forEach((rowContents, rowIndex) => {
        if (rowContents[0] && rowIndex > 0) {
            // Add merge range for department rows (merge columns 1-4)
            merges.push({
                s: { r: rowIndex, c: 0 },
                e: { r: rowIndex, c: 3 }
            });
        }

        rowContents.forEach((colContents, colIndex) => {
            const cellAddress = XLSX.utils.encode_cell({ r: rowIndex, c: colIndex });

            if (!newWorksheet[cellAddress]) return;

            // Base font
            newWorksheet[cellAddress].s = { font: { name: "Aptos Narrow (Body)" } };

            // Center align cells with break information
            if (colIndex > 3) {
                newWorksheet[cellAddress].s.alignment = { horizontal: "center" };
            }

            // Title row styling
            if (rowIndex === 0) {
                newWorksheet[cellAddress].s.font.bold = true;
                newWorksheet[cellAddress].s.font.sz = 13;
                newWorksheet[cellAddress].s.alignment = { horizontal: "center" };
            }

            // Department header row styling
            if (rowContents[0] && rowIndex > 0) {
                newWorksheet[cellAddress].s.font.bold = true;
                newWorksheet[cellAddress].s.fill = { fgColor: { rgb: "CCEEFF" } };
                newWorksheet[cellAddress].s.border = {
                    top: { style: "thin", color: { rgb: "AAAAAA" } },
                    bottom: { style: "thin", color: { rgb: "AAAAAA" } },
                    left: { style: "thin", color: { rgb: "AAAAAA" } },
                    right: { style: "thin", color: { rgb: "AAAAAA" } }
                };
            }
        });
    });

    newWorksheet["!merges"] = merges;

    return newWorkbook;
}

// Function to process the schedule and return output data
function processRowData(schedule) {
    let dept, newSchedule = [], shiftSegments = [];

    // Loop through each row starting from index 7
    for (let i = 7; i < schedule.length; i++) {
        // If the first column contains a department name, assign it to "dept" and skip further processing for that row
        if (schedule[i][0]) {
            dept = schedule[i][0];

            // Rename "Mgmt Retail" to "Management"
            if (dept === "Mgmt Retail") {
                dept = "Management";
            }

            // Remove "Training-" from department names that start with it
            if (dept.startsWith("Training-")) {
                dept = dept.replace("Training-", "");
            }

            continue;
        }

        // If there is no name, skip that row
        if (!schedule[i][2]) {
            continue;
        }

        let name = formatName(schedule[i][2]);

        // Split the time interval (column 4) into start and end times, and convert them to minutes
        let interval = schedule[i][4]?.split("-") || [];
        interval = interval.map(bound => timeToMinutes(bound));

        // Add each row to the schedule with the department, job, name, and interval
        newSchedule.push({
            dept: dept,
            job: schedule[i][1],
            name: name,
            interval: interval
        });

        // Accumulate shift segments by employee name
        if (!shiftSegments[name]) {
            shiftSegments[name] = [];
        }

        shiftSegments[name].push({
            dept: dept,
            job: schedule[i][1],
            interval: interval
        });
    }

    // Predefined order of departments
    const departmentOrder = ["Frontline", "Hardgoods", "Softgoods", "Order Fulfillment", "Product Movement", "Shop", "Management"];

    // Function to get the department index for sorting
    function getDeptOrder(d) {
        let index = departmentOrder.indexOf(d);
        return index === -1 ? departmentOrder.length : index;
    }

    // Sort by predefined dept order; keep relative order otherwise
    newSchedule.sort((a, b) => {
        let da = getDeptOrder(a.dept);
        let db = getDeptOrder(b.dept);
        if (da !== db) return da - db;
        return 0;
    });

    // Calculate the earliest and latest times for each person's shift
    let shifts = {};
    newSchedule.forEach(row => {
        if (!(row.name in shifts)) shifts[row.name] = [1440, 0];
        shifts[row.name][0] = Math.min(shifts[row.name][0], row.interval[0]);
        shifts[row.name][1] = Math.max(shifts[row.name][1], row.interval[1]);
    });

    let breaks = {};

    // Detect lunch breaks from 30-minute gaps later in the shift
    for (let name in shiftSegments) {
        let segments = shiftSegments[name];

        // Sort segments by start time
        segments.sort((a, b) => a.interval[0] - b.interval[0]);

        for (let i = 0; i < segments.length - 1; i++) {
            if (segments[i].interval[1] >= shifts[name][0] + 240 &&
                segments[i + 1].interval[0] === segments[i].interval[1] + 30) {
                if (!breaks[name]) {
                    breaks[name] = [];
                }
                // lunch break start at end of segment i
                breaks[name][1] = segments[i].interval[1];
            }
        }
    }

    // First pass: employees with lunch already in UKG
    for (let name in shifts) {
        if (!breaks[name]) continue;
        if (!breaks[name][1]) continue;

        breaks[name][0] = shifts[name][0] + 120;

        let shiftDuration = shifts[name][1] - shifts[name][0];
        let maxBreakStartTime = timeToMinutes(document.getElementById("maxBreakStartTimeInput").value);

        if (shiftDuration >= 420) {
            breaks[name][2] = Math.min(shifts[name][0] + 360, maxBreakStartTime);
        }

        // Adjust breaks to avoid overlap with others in same main+subdepartment
        for (let j = 0; j < breaks[name].length; j++) {
            if (j === 1) continue; // skip lunch (we accept overlaps for meals)

            let breakDuration = j % 2 ? 30 : 15;
            let timesDelayed = 0;

            resolveLoop: while (true) {
                let currentDept = null;
                let currentSubDept = null;

                // Which dept/subdept is this break in?
                for (let k = 0; k < newSchedule.length; k++) {
                    if (newSchedule[k].name !== name) continue;
                    if (breaks[name][j] < newSchedule[k].interval[0] ||
                        breaks[name][j] + breakDuration > newSchedule[k].interval[1]) continue;

                    currentDept = newSchedule[k].dept;
                    currentSubDept = newSchedule[k].job;
                    break;
                }

                // Skip adjustment if this main/sub pair is not enabled
                if (!isBreakAdjustmentEnabled(currentDept, currentSubDept)) break;

                // Adjust breaks to avoid overlap with others in the same main + subdepartment
                for (let k = 0; k < newSchedule.length; k++) {
                    if (newSchedule[k].name === name) continue;
                    if (newSchedule[k].dept !== currentDept || newSchedule[k].job !== currentSubDept) continue;
                    if (!(newSchedule[k].name in breaks)) continue;

                    for (let l = 0; l < breaks[newSchedule[k].name].length; l++) {
                        let breakDuration2 = l % 2 ? 30 : 15;

                        if (breaks[name][j] + breakDuration > breaks[newSchedule[k].name][l] &&
                            breaks[name][j] < breaks[newSchedule[k].name][l] + breakDuration2) {

                            let originalBreak = shifts[name][0] + 120 * (j + 1);
                            timesDelayed++;

                            breaks[name][j] = Math.min(
                                maxBreakStartTime,
                                originalBreak - (15 * Math.ceil(timesDelayed / 2) * Math.pow(-1, timesDelayed))
                            );

                            if (timesDelayed > 6) {
                                console.error(`ERROR: ${name} in ${currentDept} / ${currentSubDept} at ${minutesToTime(breaks[name][j])}`);
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

    // Second pass: employees with no lunch in UKG
    for (let name in shifts) {
        if (breaks[name]) continue;

        breaks[name] = [];
        let shiftDuration = shifts[name][1] - shifts[name][0];

        breaks[name][0] = shifts[name][0] + 120;

        if (shiftDuration >= 300) breaks[name][1] = shifts[name][0] + 240;

        let maxBreakStartTime = timeToMinutes(document.getElementById("maxBreakStartTimeInput").value);
        if (shiftDuration >= 420) breaks[name][2] = Math.min(shifts[name][0] + 360, maxBreakStartTime);

        for (let j = 0; j < breaks[name].length; j++) {
            let breakDuration = j % 2 ? 30 : 15;
            let timesDelayed = 0;

            resolveLoop: while (true) {
                let currentDept = null;
                let currentSubDept = null;

                for (let k = 0; k < newSchedule.length; k++) {
                    if (newSchedule[k].name !== name) continue;
                    if (breaks[name][j] < newSchedule[k].interval[0] ||
                        breaks[name][j] + breakDuration > newSchedule[k].interval[1]) continue;

                    currentDept = newSchedule[k].dept;
                    currentSubDept = newSchedule[k].job;
                    break;
                }

                if (!isBreakAdjustmentEnabled(currentDept, currentSubDept)) break;

                for (let k = 0; k < newSchedule.length; k++) {
                    if (newSchedule[k].name === name) continue;
                    if (newSchedule[k].dept !== currentDept || newSchedule[k].job !== currentSubDept) continue;
                    if (!(newSchedule[k].name in breaks)) continue;

                    for (let l = 0; l < breaks[newSchedule[k].name].length; l++) {
                        let breakDuration2 = l % 2 ? 30 : 15;

                        if (breaks[name][j] + breakDuration > breaks[newSchedule[k].name][l] &&
                            breaks[name][j] < breaks[newSchedule[k].name][l] + breakDuration2) {

                            let originalBreak = shifts[name][0] + 120 * (j + 1);
                            timesDelayed++;

                            breaks[name][j] = Math.min(
                                maxBreakStartTime,
                                originalBreak - (15 * Math.ceil(timesDelayed / 2) * Math.pow(-1, timesDelayed))
                            );

                            if (timesDelayed > 6) {
                                console.error(`ERROR: ${name} in ${currentDept} / ${currentSubDept} at ${minutesToTime(breaks[name][j])}`);
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

    // Extract special news breaks (Misc Events)
    let news_breaks = {};
    newSchedule.forEach(row => {
        if (row.job === "Misc Events") news_breaks[row.name] = row.interval[0];
    });

    // Remove duplicate management rows
    let uniqueRows = [];
    let seen = new Set();
    newSchedule.forEach(row => {
        if (row.dept === "Management" && row.job === "Management") {
            const key = row.name;
            if (!seen.has(key)) {
                seen.add(key);
                uniqueRows.push(row);
            }
        } else {
            uniqueRows.push(row);
        }
    });
    newSchedule = uniqueRows;

    // Update management shift intervals to cover full shift
    newSchedule.forEach(row => {
        if (row.job === "Management") {
            row.interval = shifts[row.name];
        }
    });

    // Remove Misc Events rows
    newSchedule = newSchedule.filter(row => row.job !== "Misc Events");

    // Build final output
    const output = [["REI Daily Schedule"]];
    let currentDept = null;
    let breaksPrinted = new Set();

    newSchedule.forEach(row => {
        if (row.dept !== currentDept) {
            currentDept = row.dept;
            output.push([currentDept, "", "", "", "15", "30", "15", "News"]);
        }

        let intervalStr = row.interval.map(bound => minutesToTime(bound)).join("-");
        let rowOutput = ["", row.job, row.name, intervalStr];

        const key = row.name;
        if (!breaksPrinted.has(key)) {
            breaksPrinted.add(key);

            for (let i = 0; i < 3; i++) {
                if (breaks[row.name][i]) rowOutput.push(minutesToTime(breaks[row.name][i]));
                else rowOutput.push("x");
            }

            if (news_breaks[row.name]) rowOutput.push(minutesToTime(news_breaks[row.name]));
            else rowOutput.push("x");
        } else {
            rowOutput.push("x", "x", "x", "x");
        }

        output.push(rowOutput);
    });

    return output;
}

// Helper: convert string to ArrayBuffer
function stringToArrayBuffer(string) {
    const buffer = new ArrayBuffer(string.length);
    const view = new Uint8Array(buffer);
    for (let i = 0; i < string.length; i++) {
        view[i] = string.charCodeAt(i) & 0xFF;
    }
    return buffer;
}

// Helper: time strings → minutes
function timeToMinutes(time) {
    if (!time) return 0;
    return (time.includes("AM") || time.includes("PM")
        ? ((+time.split(":")[0] % 12) + (time.includes("PM") ? 12 : 0)) * 60 + +time.split(":")[1].slice(0, 2)
        : +time.split(":")[0] * 60 + +time.split(":")[1]);
}

// Helper: minutes → time strings
function minutesToTime(minutes) {
    return `${(minutes / 60 + 11) % 12 + 1 | 0}:${(minutes % 60).toString().padStart(2, "0")}${minutes < 720 ? "AM" : "PM"}`;
}

// Helper: format name
function formatName(name) {
    const words = name.split(" ").map(word => {
        word = word[0].toUpperCase() + word.slice(1).toLowerCase();

        if (word.startsWith("Mc") || word.startsWith("O'")) {
            word = word.slice(0, 2) + word[2].toUpperCase() + word.slice(3);
        }

        return word;
    });

    const formattedWords = words.slice(0, 2);
    return formattedWords.join(" ");
}

// Extract the schedule date from the raw worksheet rows
function extractScheduleDate(schedule) {
    // Look for "Date: YYYY-MM-DD"
    for (let i = 0; i < schedule.length; i++) {
        const cell = schedule[i][0];
        if (typeof cell === "string") {
            const isoMatch = cell.match(/Date:\s*(\d{4}-\d{2}-\d{2})/);
            if (isoMatch) {
                return isoMatch[1];
            }
        }
    }

    // Fallback on today's date
    return new Date().toLocaleDateString("en-US").replace(/\//g, "-");
}

// Read main/subdepartment options from UI
function getSelectedDepartmentPairs() {
    const pairs = [];
    const departmentCheckboxes = document.querySelectorAll("#departmentOverlap .form-check-input");

    departmentCheckboxes.forEach(checkbox => {
        if (!checkbox.checked) return;
        const main = checkbox.dataset.main || "";
        const sub = checkbox.dataset.sub || "*";
        if (main) {
            pairs.push({ main, sub });
        }
    });

    return pairs;
}

// Whether we should stagger for this main+sub combo
function isBreakAdjustmentEnabled(mainDept, subDept) {
    if (!mainDept) return false;

    const selectedPairs = getSelectedDepartmentPairs();
    return selectedPairs.some(pair => {
        if (pair.main !== mainDept) return false;
        if (pair.sub === "*" || !pair.sub) return true;
        return pair.sub === subDept;
    });
}
