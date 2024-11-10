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

        const date = new Date().toLocaleDateString("en-US").replace(/\//g, "-");
        const workbookBlob = XLSX.write(styledRowData, { bookType: "xlsx", type: "binary" });
        const blob = new Blob([stringToArrayBuffer(workbookBlob)], { type: "application/octet-stream" });
        const url = URL.createObjectURL(blob);

        // Update the hidden download link with the generated file URL and filename
        const hiddenDownloadLink = document.getElementById("hiddenDownloadLink");
        hiddenDownloadLink.href = url;
        hiddenDownloadLink.download = `Break Schedule ${date}.xlsx`;

        // Simulate a click on the hidden download link
        hiddenDownloadLink.click();
    };

    reader.readAsArrayBuffer(selectedFile);
});

// Function to style the schedule
function styleRowData(schedule) {
    const newWorkbook = XLSX.utils.book_new();
    const newWorksheet = XLSX.utils.aoa_to_sheet(schedule);
    XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, "Break Schedule");

    // Set column widths
    newWorksheet["!cols"] = [
        { wpx: 40 }, { wpx: 150 }, { wpx: 120 }, { wpx: 100 }, { wpx: 80 }, { wpx: 80 }, { wpx: 80 }, { wpx: 80 }
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
            // Convert row and column indices to cell address, e.g. row 2, col 3 -> "B3"
            const cellAddress = XLSX.utils.encode_cell({ r: rowIndex, c: colIndex });

            // Skip undefined cells
            if (!newWorksheet[cellAddress]) return;

            // Set all cells to Arial
            newWorksheet[cellAddress].s = { font: { name: "Aptos Narrow (Body)" } };

            // Center align cells with break information
            if (colIndex > 3) {
                newWorksheet[cellAddress].s.alignment = { horizontal: "center" };
            }

            // Set title row styling to center align, bold, 13 pt
            if (rowIndex === 0) {
                newWorksheet[cellAddress].s.font.bold = true;
                newWorksheet[cellAddress].s.font.sz = 13;
                newWorksheet[cellAddress].s.alignment = { horizontal: "center" };
            }

            // Style department rows with blue background
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

        // If there is no name, skip that row, otherwise format the name
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
            shiftSegments[name] = []; // Initialize the array if it doesn't exist
        }

        // Push the shift segment (department, job, interval) into the array for this employee
        shiftSegments[name].push({
            dept: dept,
            job: schedule[i][1],
            interval: interval
        });
    }

    // Predefined order of departments
    const departmentOrder = ["Frontline", "Hardgoods", "Softgoods", "Order Fulfillment", "Product Movement", "Shop", "Management"];

    // Function to get the department index for sorting
    function getDeptOrder(dept) {
        let index = departmentOrder.indexOf(dept);
        return index === -1 ? departmentOrder.length : index; // Departments not in the list come last, but in original order
    }

    // Sort the schedule based on the predefined department order, leaving the rest unsorted
    newSchedule.sort((a, b) => {
        let deptOrderA = getDeptOrder(a.dept);
        let deptOrderB = getDeptOrder(b.dept);

        if (deptOrderA !== deptOrderB) {
            return deptOrderA - deptOrderB;  // Sort by department order if they differ
        }

        // Maintain the original order within departments for unsorted departments
        if (deptOrderA === departmentOrder.length) {
            return 0;  // Leave unsorted departments in original order
        }

        return 0;  // Don"t sort within predefined departments by name, keep original order
    });

    // Calculate the earliest and latest times for each person"s shift
    let shifts = {};
    newSchedule.forEach(row => {
        if (!(row.name in shifts)) shifts[row.name] = [1440, 0]; // Default shift interval (1440 = 24:00)
        shifts[row.name][0] = Math.min(shifts[row.name][0], row.interval[0]); // Earliest start time
        shifts[row.name][1] = Math.max(shifts[row.name][1], row.interval[1]); // Latest end time
    });

    let breaks = {};

    // Add logic to detect a lunch break based on a 30-minute gap between segments (between 3-5 hours into the shift)
    for (let name in shiftSegments) {
        let segments = shiftSegments[name];

        // Sort the segments for each employee by start time
        segments.sort((a, b) => a.interval[0] - b.interval[0]);

        for (let i = 0; i < segments.length - 1; i++) {
            if (segments[i].interval[1] >= shifts[name][0] + 240 && segments[i + 1].interval[0] == segments[i].interval[1] + 30) {
                if (!breaks[name]) {
                    breaks[name] = [];
                }
                breaks[name][1] = segments[i].interval[1];
            }
        }
    }

    // First generate breaks for employees with lunch breaks already assigned in UKG
    for (let name in shifts) {
        // If the employee does not have any breaks already, skip that employee
        if (!breaks[name]) {
            continue;
        }

        // If the employee does not have a lunch break already, skip that employee
        if (!breaks[name][1]) {
            continue;
        }

        // Add a first rest break at 2 hours after the start time
        breaks[name][0] = shifts[name][0] + 120;

        let shiftDuration = shifts[name][1] - shifts[name][0];

        let maxBreakStartTime = timeToMinutes(document.getElementById('maxBreakStartTimeInput').value);

        // If the shift duration is equal to or greater than 7 hours, add a second rest break at 6 hours after the start time
        if (shiftDuration >= 420) breaks[name][2] = Math.min(shifts[name][0] + 360, maxBreakStartTime);

        // Adjust breaks if they overlap with other breaks in the same department
        for (let j = 0; j < breaks[name].length; j++) {
            // If the break is a meal break, skip that break
            if (j == 1) {
                continue;
            }

            let breakDuration = j % 2 ? 30 : 15; // 30 minutes for second break, 15 minutes for the first
            let timesDelayed = 0;

            // Resolve conflicts with other schedule intervals
            resolveLoop: while (true) {
                let currentDept = null;

                // Check if the break overlaps with the employee"s scheduled intervals
                for (let k = 0; k < newSchedule.length; k++) {
                    if (newSchedule[k].name !== name) continue;
                    if (breaks[name][j] < newSchedule[k].interval[0] || breaks[name][j] + breakDuration > newSchedule[k].interval[1]) continue;
                    currentDept = newSchedule[k].dept;
                    break;
                }

                // Retrieve selected departments for break adjustment
                const selectedDepartments = getSelectedDepartments();

                // Skip break adjustment if the employee is not in one of the selected departments
                if (!selectedDepartments.includes(currentDept)) break;

                // Adjust breaks to avoid overlap with others in the same department
                for (let k = 0; k < newSchedule.length; k++) {
                    if (newSchedule[k].name === name || newSchedule[k].dept !== currentDept) continue;
                    if (!(newSchedule[k].name in breaks)) continue;

                    for (let l = 0; l < breaks[newSchedule[k].name].length; l++) {
                        let breakDuration2 = l % 2 ? 30 : 15;

                        // Check if breaks overlap
                        if (breaks[name][j] + breakDuration > breaks[newSchedule[k].name][l] &&
                            breaks[name][j] < breaks[newSchedule[k].name][l] + breakDuration2) {

                            let originalBreak = shifts[name][0] + 120 * (j + 1); // Original break time
                            timesDelayed++;

                            // Adjust break by delaying or advancing it by 15-minute intervals
                            breaks[name][j] = Math.min(maxBreakStartTime, originalBreak - (15 * Math.ceil(timesDelayed / 2) * Math.pow(-1, timesDelayed)));

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

    // Then, generate breaks for the rest of the employees (employees not assigned a lunch in UKG)
    for (let name in shifts) {
        // If the employee 
        if (breaks[name]) {
            continue;
        }

        // Initialize breaks
        breaks[name] = [];

        let shiftDuration = shifts[name][1] - shifts[name][0];

        // Add a first rest break at 2 hours after the start time
        breaks[name][0] = shifts[name][0] + 120;

        // If the shift duration is equal to or greater than 5 hours, add a meal break at 4 hours after the start time
        if (shiftDuration >= 300) breaks[name][1] = shifts[name][0] + 240;

        let maxBreakStartTime = timeToMinutes(document.getElementById('maxBreakStartTimeInput').value);

        // If the shift duration is equal to or greater than 7 hours, add a second rest break at 6 hours after the start time
        if (shiftDuration >= 420) breaks[name][2] = Math.min(shifts[name][0] + 360, maxBreakStartTime);

        // Adjust breaks if they overlap with other breaks in the same department
        for (let j = 0; j < breaks[name].length; j++) {
            let breakDuration = j % 2 ? 30 : 15; // 30 minutes for second break, 15 minutes for the first
            let timesDelayed = 0;

            // Resolve conflicts with other schedule intervals
            resolveLoop: while (true) {
                let currentDept = null;

                // Check if the break overlaps with the employee's scheduled intervals
                for (let k = 0; k < newSchedule.length; k++) {
                    if (newSchedule[k].name !== name) continue;
                    if (breaks[name][j] < newSchedule[k].interval[0] || breaks[name][j] + breakDuration > newSchedule[k].interval[1]) continue;
                    currentDept = newSchedule[k].dept;
                    break;
                }

                // Retrieve selected departments for break adjustment
                const selectedDepartments = getSelectedDepartments();

                // Skip break adjustment if the employee is not in one of the selected departments
                if (!selectedDepartments.includes(currentDept)) break;

                // Adjust breaks to avoid overlap with others in the same department
                for (let k = 0; k < newSchedule.length; k++) {
                    if (newSchedule[k].name === name || newSchedule[k].dept !== currentDept) continue;
                    if (!(newSchedule[k].name in breaks)) continue;

                    for (let l = 0; l < breaks[newSchedule[k].name].length; l++) {
                        let breakDuration2 = l % 2 ? 30 : 15;

                        // Check if breaks overlap
                        if (breaks[name][j] + breakDuration > breaks[newSchedule[k].name][l] &&
                            breaks[name][j] < breaks[newSchedule[k].name][l] + breakDuration2) {

                            let originalBreak = shifts[name][0] + 120 * (j + 1); // Original break time
                            timesDelayed++;

                            // Adjust break by delaying or advancing it by 15-minute intervals
                            breaks[name][j] = Math.min(maxBreakStartTime, originalBreak - (15 * Math.ceil(timesDelayed / 2) * Math.pow(-1, timesDelayed)));

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

    // Extract special breaks for "Misc Events"
    let news_breaks = {};
    newSchedule.forEach(row => {
        if (row.job === "Misc Events") news_breaks[row.name] = row.interval[0];
    });

    // Remove duplicate rows for "Management" department, keeping only unique entries by name
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

    // Update shift intervals for management jobs to cover the entire shift
    newSchedule.forEach(row => {
        if (row.job === "Management") {
            row.interval = shifts[row.name];
        }
    });

    // Remove rows related to "Misc Events" or "Meeting"
    newSchedule = newSchedule.filter(row => row.job !== "Misc Events");

    // Create the output array for Excel
    const output = [["REI Daily Schedule"]];
    // output.push(["Dept.", "Job", "Name", "Shift"]);

    // Track printed breaks to avoid duplicate entries
    let breaksPrinted = new Set();
    newSchedule.forEach(row => {
        if (row.dept !== dept) {
            dept = row.dept;
            output.push([dept, "", "", "", "15", "30", "15", "News"]); // Add department header row
        }
        let intervalStr = row.interval.map(bound => minutesToTime(bound)).join("-"); // Format shift interval
        let rowOutput = ["", row.job, row.name, intervalStr];

        const key = row.name;
        if (!breaksPrinted.has(key)) {
            breaksPrinted.add(key);
            // Add regular breaks to each row (15, 30, 15)
            for (let i = 0; i < 3; i++) {
                if (breaks[row.name][i]) rowOutput.push(minutesToTime(breaks[row.name][i]));
                else rowOutput.push("x");
            }

            // Add news break to each row
            if (news_breaks[row.name]) rowOutput.push(minutesToTime(news_breaks[row.name]));
            else rowOutput.push("x");
        }

        output.push(rowOutput);
    });

    return output; // Return the processed schedule for further use
}

// Helper function to convert string to ArrayBuffer
function stringToArrayBuffer(string) {
    const buffer = new ArrayBuffer(string.length);
    const view = new Uint8Array(buffer);
    for (let i = 0; i < string.length; i++) {
        view[i] = string.charCodeAt(i) & 0xFF;
    }
    return buffer;
}

// Helper function to convert time strings to minutes since midnight (e.g. "9:00AM" -> 540)
function timeToMinutes(time) {
    return (time.includes("AM") || time.includes("PM")
        ? ((+time.split(":")[0] % 12) + (time.includes("PM") ? 12 : 0)) * 60 + +time.split(":")[1].slice(0, 2)
        : +time.split(":")[0] * 60 + +time.split(":")[1]);
}

// Helper function to convert minutes since midnight to time strings (e.g. 540 -> "9:00AM")
function minutesToTime(minutes) {
    return `${(minutes / 60 + 11) % 12 + 1 | 0}:${(minutes % 60).toString().padStart(2, "0")}${minutes < 720 ? "AM" : "PM"}`;
}

// Helper function to properly format names
function formatName(name) {
    const words = name.split(" ").map(word => {
        // Set the first letter of the word to upper case and the rest lower case
        word = word[0].toUpperCase() + word.slice(1).toLowerCase();

        // If word starts with "Mc" or "O'", capitalize the third letter of the word
        if (word.startsWith("Mc") || word.startsWith("O'")) {
            // Create a new string with the third letter capitalized
            word = word.slice(0, 2) + word[2].toUpperCase() + word.slice(3);
        }

        return word; // Ensure to return the modified word
    });

    // Remove middle initial if present
    const formattedWords = words.slice(0, 2); // renamed for clarity
    const formattedName = formattedWords.join(" ");
    return formattedName;
}

function getSelectedDepartments() {
    const selectedDepartments = [];
    const departmentCheckboxes = document.querySelectorAll('#departmentOverlap .form-check-input');

    departmentCheckboxes.forEach(checkbox => {
        if (checkbox.checked) {
            selectedDepartments.push(checkbox.nextElementSibling.textContent.trim());
        }
    });

    return selectedDepartments;
}