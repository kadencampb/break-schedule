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
        // Read data from the selected file into an XLSX workbook, preserving cell styles
        const workbook = XLSX.read(new Uint8Array(event.target.result), {
            type: "array",
            cellStyles: true
        });

        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];

        // Convert the first worksheet from the XLSX workbook into JSON row data
        const rowData = XLSX.utils.sheet_to_json(sheet, { header: 1 });

        // Derive the schedule date from the spreadsheet contents
        const scheduleDate = extractScheduleDate(rowData);

        // Get operating hours for the schedule date's day of week
        const operatingHours = getOperatingHoursForDate(scheduleDate);

        // Process the data and insert breaks directly into the existing sheet
        processRowDataInPlace(sheet, rowData, operatingHours);

        // Apply styling to the sheet
        applyReportStyling(sheet, rowData);

        const workbookBlob = XLSX.write(workbook, {
            bookType: "xlsx",
            type: "binary",
            cellStyles: true
        });
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
function detectDataStart(rows) {
    // Look for a header row where column 3 says "Name"
    for (let i = 0; i < rows.length; i++) {
        const row = rows[i];
        if (!row) continue;
        if (typeof row[2] === "string" && row[2].trim().toLowerCase() === "name") {
            return i + 1; // data starts on the next row
        }
    }
    // Fallback if header not found
    return 8;
}

function applyReportStyling(sheet, rows) {
    if (!sheet["!ref"]) return;

    const range = XLSX.utils.decode_range(sheet["!ref"]);
    const lastRow = range.e.r;

    function cellRef(r, c) {
        return XLSX.utils.encode_cell({ r, c });
    }

    // Create / touch a cell and make sure it exists with proper initial structure
    // IMPORTANT: For merged cells, we need to create the cell object even if it's not the top-left
    function touchCell(r, c, forceCreate = true) {
        const ref = cellRef(r, c);
        let cell = sheet[ref];

        if (!cell && forceCreate) {
            // Create a new cell object
            cell = { t: "s", v: " " };
            sheet[ref] = cell;
        }

        if (!cell) return null; // Cell doesn't exist and we're not forcing creation

        // ALWAYS ensure cell has a value for borders to render properly
        if (cell.v === undefined || cell.v === null || cell.v === "") {
            cell.t = "s";
            cell.v = " ";
        }

        // Ensure cell has a style object structure (but don't share references)
        if (!cell.s) {
            cell.s = {
                font: {},
                alignment: {}
            };
        }
        if (!cell.s.font) cell.s.font = {};
        if (!cell.s.alignment) cell.s.alignment = {};

        return cell;
    }

    // 1) Base: ensure EVERY cell in A:G, rows 1..lastRow, exists and has default style
    //    This guarantees blank cells like B7 actually exist and can get borders.
    for (let r = 0; r <= lastRow; r++) {
        for (let c = 0; c <= 6; c++) {  // only columns A–G
            const cell = touchCell(r, c);
            // Each cell gets its own unique style object
            cell.s = {
                font: { name: "Arial", sz: 6.75, bold: false },
                alignment: { horizontal: "left", vertical: "top", wrapText: true }
            };
            // no border by default (will be added later for specific rows)
        }
    }

    // 2) A1: Arial, 12 pt, bold, center, center, no border
    (function () {
        const cell = touchCell(0, 0); // A1
        cell.s = {
            font: { name: "Arial", sz: 12, bold: true },
            alignment: { horizontal: "center", vertical: "center" }
        };
    })();

    // 3) A2:A3 Arial, 7.5 pt, not bold, left, center
    for (let r = 1; r <= 2; r++) {
        const cell = touchCell(r, 0); // column A
        cell.s = {
            font: { name: "Arial", sz: 7.5, bold: false },
            alignment: { horizontal: "left", vertical: "center" }
        };
    }

    // 4) C2:C3 Arial, 7.5 pt, not bold, left, top
    for (let r = 1; r <= 2; r++) {
        const cell = touchCell(r, 2); // column C
        cell.s = {
            font: { name: "Arial", sz: 7.5, bold: false },
            alignment: { horizontal: "left", vertical: "top" }
        };
    }

    // 5) E2:E3 Arial, 7.5 pt, not bold, right, center
    for (let r = 1; r <= 2; r++) {
        const cell = touchCell(r, 4); // column E
        cell.s = {
            font: { name: "Arial", sz: 7.5, bold: false },
            alignment: { horizontal: "right", vertical: "center" }
        };
    }

    // 6) F2:F3 Arial, 7.5 pt, not bold, left, top
    for (let r = 1; r <= 2; r++) {
        const cell = touchCell(r, 5); // column F
        cell.s = {
            font: { name: "Arial", sz: 7.5, bold: false },
            alignment: { horizontal: "left", vertical: "top" }
        };
    }

    // 7) A5:A6 Arial, 7.5 pt, bold, left, center
    for (let r = 4; r <= 5; r++) {
        const cell = touchCell(r, 0); // column A
        cell.s = {
            font: { name: "Arial", sz: 7.5, bold: true },
            alignment: { horizontal: "left", vertical: "center" }
        };
    }

    // 8) A7:G7 Arial, 7.5 pt, bold, left, center, thin 000000 border
    (function () {
        const r = 6; // row 7 (0-based)
        for (let c = 0; c <= 6; c++) {
            const cell = touchCell(r, c);
            cell.s = {
                font: { name: "Arial", sz: 7.5, bold: true },
                alignment: { horizontal: "left", vertical: "center" },
                border: {
                    top:    { style: "thin", color: { rgb: "000000" } },
                    bottom: { style: "thin", color: { rgb: "000000" } },
                    left:   { style: "thin", color: { rgb: "000000" } },
                    right:  { style: "thin", color: { rgb: "000000" } }
                }
            };
        }
    })();

    // 9) Header / non-header rows A–G from dataStart downward
    const dataStart = detectDataStart(rows); // first data row index (0-based)

    for (let r = dataStart; r <= lastRow; r++) {
        const row = rows[r] || [];
        const colA = row[0];
        const colC = row[2];

        const hasDept = colA != null && String(colA).trim() !== "";
        const hasName = colC != null && String(colC).trim() !== "";

        if (hasDept && !hasName) {
            // Header row: columns A-G (bold, with borders)
            for (let c = 0; c <= 6; c++) {
                const cell = touchCell(r, c);
                cell.s = {
                    font: { name: "Arial", sz: 7.5, bold: true },
                    alignment: { horizontal: "left", vertical: "top" },
                    border: {
                        top:    { style: "thin", color: { rgb: "000000" } },
                        bottom: { style: "thin", color: { rgb: "000000" } },
                        left:   { style: "thin", color: { rgb: "000000" } },
                        right:  { style: "thin", color: { rgb: "000000" } }
                    }
                };
            }
        } else if (hasName) {
            // Non-header row: columns A-G (normal, with borders)
            for (let c = 0; c <= 6; c++) {
                const cell = touchCell(r, c);
                cell.s = {
                    font: { name: "Arial", sz: 6.75, bold: false },
                    alignment: { horizontal: "left", vertical: "top" },
                    border: {
                        top:    { style: "thin", color: { rgb: "000000" } },
                        bottom: { style: "thin", color: { rgb: "000000" } },
                        left:   { style: "thin", color: { rgb: "000000" } },
                        right:  { style: "thin", color: { rgb: "000000" } }
                    }
                };
            }
        } else if (!hasDept && !hasName) {
            // Empty row or continuation row: still apply borders to all cells A-G
            for (let c = 0; c <= 6; c++) {
                const cell = touchCell(r, c);
                cell.s = {
                    font: { name: "Arial", sz: 6.75, bold: false },
                    alignment: { horizontal: "left", vertical: "top" },
                    border: {
                        top:    { style: "thin", color: { rgb: "000000" } },
                        bottom: { style: "thin", color: { rgb: "000000" } },
                        left:   { style: "thin", color: { rgb: "000000" } },
                        right:  { style: "thin", color: { rgb: "000000" } }
                    }
                };
            }
        }
    }

    // CRITICAL FIX: For cells in merged ranges, Excel might not display borders properly
    // unless the cells are explicitly created and styled. Let's ensure all cells in
    // merged ranges have the same border styling as their top-left cell.
    if (sheet["!merges"]) {
        sheet["!merges"].forEach(merge => {
            // Only process merges that are in our styling range (rows >= 6, cols A-G)
            if (merge.s.r >= 6 && merge.s.c <= 6 && merge.e.c <= 6) {
                // Get the top-left cell of the merge to copy its style
                const topLeftCell = sheet[cellRef(merge.s.r, merge.s.c)];
                if (!topLeftCell || !topLeftCell.s || !topLeftCell.s.border) return;

                // Apply the same border style to all cells in the merge range
                for (let r = merge.s.r; r <= merge.e.r; r++) {
                    for (let c = merge.s.c; c <= merge.e.c; c++) {
                        // Skip the top-left cell (already styled)
                        if (r === merge.s.r && c === merge.s.c) continue;

                        const cell = touchCell(r, c);
                        if (cell && topLeftCell.s) {
                            // Copy the complete style from top-left cell
                            cell.s = {
                                font: { ...topLeftCell.s.font },
                                alignment: { ...topLeftCell.s.alignment },
                                border: {
                                    top: topLeftCell.s.border.top ? { ...topLeftCell.s.border.top } : undefined,
                                    bottom: topLeftCell.s.border.bottom ? { ...topLeftCell.s.border.bottom } : undefined,
                                    left: topLeftCell.s.border.left ? { ...topLeftCell.s.border.left } : undefined,
                                    right: topLeftCell.s.border.right ? { ...topLeftCell.s.border.right } : undefined
                                }
                            };
                        }
                    }
                }
            }
        });
    }
}

// Function to process the schedule and modify the sheet in place
function processRowDataInPlace(sheet, schedule, operatingHours) {
    // Helper function to set cell value
    function setCell(rZero, cZero, value) {
        const ref = XLSX.utils.encode_cell({ r: rZero, c: cZero });
        if (value == null) {
            // Don't delete - set to space to preserve borders
            const cell = sheet[ref] || {};
            cell.t = "s";
            cell.v = " ";
            sheet[ref] = cell;
            return;
        }
        const cell = sheet[ref] || {};
        cell.t = "s";
        cell.v = value;
        sheet[ref] = cell;
    }

    // Detect where data starts
    const dataStart = detectDataStart(schedule);
    const headerRowIndex = dataStart - 1;

    // Update column headers (D–G): Shift / 15 / 30 / 15
    if (headerRowIndex >= 0) {
        const headerLabels = ["Shift", "15", "30", "15"];
        for (let offset = 0; offset < headerLabels.length; offset++) {
            setCell(headerRowIndex, 3 + offset, headerLabels[offset]);
        }
    }

    // Process the schedule data with operating hours
    const result = processScheduleData(schedule, operatingHours);
    const { breaks, segments } = result;

    // Track which employees have had their breaks written
    const printed = new Set();

    // Write breaks into the sheet for each segment row
    segments.forEach(seg => {
        const empName = seg.name;
        const rZero = seg.rowIndex;
        const empBreaks = breaks[empName] || [];
        const firstRowForEmp = !printed.has(empName);

        // Copy Shift from column E (index 4) to D (index 3)
        if (seg.intervalStr) {
            setCell(rZero, 3, seg.intervalStr);
        }

        if (firstRowForEmp) {
            // First segment row for this person: write their breaks
            setCell(rZero, 4, empBreaks[0] ? minutesToTime(empBreaks[0]) : "x"); // E: first rest
            setCell(rZero, 5, empBreaks[1] ? minutesToTime(empBreaks[1]) : "x"); // F: meal
            setCell(rZero, 6, empBreaks[2] ? minutesToTime(empBreaks[2]) : "x"); // G: second rest

            printed.add(empName);
        } else {
            // Subsequent segments for the same person: mark as "x"
            setCell(rZero, 4, "x");
            setCell(rZero, 5, "x");
            setCell(rZero, 6, "x");
        }
    });

    // Ensure !ref includes columns up through G
    const ref = sheet["!ref"] || "A1";
    const range = XLSX.utils.decode_range(ref);
    if (range.e.c < 6) { // col G is index 6
        range.e.c = 6;
        sheet["!ref"] = XLSX.utils.encode_range(range);
    }
}

// Function to process the schedule and return breaks + segments
function processScheduleData(schedule, operatingHours) {
    let dept, newSchedule = [], shiftSegments = [];
    const segments = []; // Track segments with row indices for writing back

    // Extract operating hours (defaults to 9 AM - 9 PM if not provided)
    const startOfDay = operatingHours ? operatingHours.startTime : 9 * 60;
    const endOfDay = operatingHours ? operatingHours.endTime : 21 * 60;

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
        const intervalStr = schedule[i][4]; // Keep original string for display

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

        // Track segments with their row indices for writing back to the sheet
        segments.push({
            name: name,
            dept: dept,
            job: schedule[i][1],
            start: interval[0],
            end: interval[1],
            intervalStr: intervalStr,
            rowIndex: i
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

    // ====================================================================================
    // NEW CALIFORNIA LAW COMPLIANT BREAK SCHEDULING WITH COVERAGE OPTIMIZATION
    // ====================================================================================

    /**
     * Calculate coverage map for all 15-minute intervals during operating hours
     * Returns a map where keys are time in minutes, values are arrays of employee names working
     */
    function calculateCoverageMap(schedule, shifts, breaks = {}) {
        const coverage = {};

        // Initialize all 15-minute intervals for the operating hours
        for (let time = startOfDay; time < endOfDay; time += 15) {
            coverage[time] = [];
        }

        // For each employee, mark which intervals they're working (not on break)
        schedule.forEach(row => {
            const name = row.name;
            const [shiftStart, shiftEnd] = row.interval;

            // Get this employee's scheduled breaks
            const employeeBreaks = breaks[name] || [];

            // For each 15-minute interval during operating hours
            for (let time = startOfDay; time < endOfDay; time += 15) {
                // Check if employee is working during this interval
                if (time >= shiftStart && time < shiftEnd) {
                    // Check if employee is on a break during this interval
                    let onBreak = false;

                    for (let i = 0; i < employeeBreaks.length; i++) {
                        if (employeeBreaks[i] === undefined) continue;

                        const breakStart = employeeBreaks[i];
                        const breakDuration = (i === 1) ? 30 : 15; // index 1 is meal (30 min), others are rest (15 min)
                        const breakEnd = breakStart + breakDuration;

                        // Check if this 15-min interval overlaps with the break
                        if (time < breakEnd && time + 15 > breakStart) {
                            onBreak = true;
                            break;
                        }
                    }

                    if (!onBreak) {
                        coverage[time].push({
                            name: name,
                            dept: row.dept,
                            subdept: row.job
                        });
                    }
                }
            }
        });

        return coverage;
    }

    /**
     * Get all employees working in the same dept/subdept at a given time
     */
    function getCoworkersAtTime(coverage, time, dept, subdept) {
        if (!coverage[time]) return [];
        return coverage[time].filter(emp => emp.dept === dept && emp.subdept === subdept);
    }

    /**
     * Count how many employees from the same dept/subdept are on break at a specific time
     */
    function countBreaksAtTime(schedule, breaks, targetTime, breakDuration, dept, subdept) {
        let count = 0;

        for (let name in breaks) {
            // Find this employee's dept/subdept
            const empRow = schedule.find(row => row.name === name);
            if (!empRow || empRow.dept !== dept || empRow.job !== subdept) continue;

            // Check all their breaks
            const empBreaks = breaks[name];
            for (let i = 0; i < empBreaks.length; i++) {
                if (empBreaks[i] === undefined) continue;

                const breakStart = empBreaks[i];
                const duration = (i === 1) ? 30 : 15;
                const breakEnd = breakStart + duration;

                // Check if this break overlaps with our target time
                if (targetTime < breakEnd && targetTime + breakDuration > breakStart) {
                    count++;
                    break; // Only count each employee once
                }
            }
        }

        return count;
    }

    // Initialize breaks object
    let breaks = {};

    // Detect lunch breaks already scheduled in UKG (from 30-minute gaps in the schedule)
    for (let name in shiftSegments) {
        let segments = shiftSegments[name];
        segments.sort((a, b) => a.interval[0] - b.interval[0]);

        for (let i = 0; i < segments.length - 1; i++) {
            if (segments[i].interval[1] >= shifts[name][0] + 240 &&
                segments[i + 1].interval[0] === segments[i].interval[1] + 30) {
                if (!breaks[name]) {
                    breaks[name] = [];
                }
                breaks[name][1] = segments[i].interval[1];
            }
        }
    }

    // ====================================================================================
    // STEP 1: SCHEDULE MEAL PERIODS (CALIFORNIA LAW COMPLIANT)
    // ====================================================================================

    // Determine who needs meal periods based on shift duration
    // Process employees in schedule order for consistent break assignment
    const employeesNeedingMeals = [];
    const processedForMeals = new Set();

    newSchedule.forEach(row => {
        const name = row.name;

        // Skip if we've already processed this employee
        if (processedForMeals.has(name)) return;
        processedForMeals.add(name);

        const shiftDuration = shifts[name][1] - shifts[name][0];

        // California law (with 7-minute early clock-in allowance):
        // 0:00-4:45: no meal period
        // 4:46-9:45: one meal period
        // 9:46+: two meal periods

        let mealsNeeded = 0;
        if (shiftDuration > 285) mealsNeeded = 1;      // > 4:45
        if (shiftDuration > 585) mealsNeeded = 2;      // > 9:45

        if (mealsNeeded > 0) {
            employeesNeedingMeals.push({
                name: name,
                shiftStart: shifts[name][0],
                shiftEnd: shifts[name][1],
                mealsNeeded: mealsNeeded,
                hasUKGLunch: breaks[name] && breaks[name][1] !== undefined
            });
        }
    });

    // Schedule first meal period at ideal time (+4 hours), handling conflicts greedily
    for (let emp of employeesNeedingMeals) {
        if (!breaks[emp.name]) breaks[emp.name] = [];

        // Skip if already has lunch from UKG
        if (emp.hasUKGLunch) continue;

        // Find the employee's dept/subdept
        const empRow = newSchedule.find(row => row.name === emp.name);
        if (!empRow) continue;

        const dept = empRow.dept;
        const subdept = empRow.job;

        // Skip if break staggering is disabled for this dept/subdept
        if (!isBreakAdjustmentEnabled(dept, subdept)) {
            // Just schedule at ideal time without conflict resolution
            breaks[emp.name][1] = emp.shiftStart + 240; // +4:00 hours
            continue;
        }

        // Ideal meal time: +4 hours into shift
        const idealMealTime = emp.shiftStart + 240;

        // Count how many coworkers already have lunch at the ideal time
        const conflictsAtIdeal = countBreaksAtTime(newSchedule, breaks, idealMealTime, 30, dept, subdept);

        console.log(`[MEAL DEBUG] ${emp.name} (${dept}/${subdept}): ideal time ${minutesToTime(idealMealTime)}, conflicts: ${conflictsAtIdeal}`);

        // Conflict resolution:
        // 1 employee: okay (0 conflicts)
        // 2+ employees: stagger to delayed time (+4:30)
        // Only allow 1 employee per lunch slot

        if (conflictsAtIdeal === 0) {
            // First employee at ideal time - no conflicts
            breaks[emp.name][1] = idealMealTime;
            console.log(`[MEAL SCHEDULE] ${emp.name} (${dept}/${subdept}): lunch scheduled at ${minutesToTime(idealMealTime)}`);
        } else {
            // Conflict detected - delay to +4:30 hours
            breaks[emp.name][1] = emp.shiftStart + 270;
            console.log(`[MEAL STAGGER] ${emp.name} (${dept}/${subdept}): lunch delayed from ${minutesToTime(idealMealTime)} to ${minutesToTime(emp.shiftStart + 270)} due to ${conflictsAtIdeal} conflict(s)`);
        }
    }

    // Handle second meal period (for shifts > 9:30)
    for (let emp of employeesNeedingMeals) {
        if (emp.mealsNeeded < 2) continue;

        // Second meal should be scheduled later in the shift
        // We'll use a simple approach: +8 hours into shift
        breaks[emp.name].push(emp.shiftStart + 480);
    }

    // ====================================================================================
    // STEP 2: SCHEDULE REST PERIODS (15-MINUTE BREAKS)
    // ====================================================================================

    // Get unique employee names in schedule order (not shifts object order)
    const employeesInOrder = [];
    const seenEmployees = new Set();
    newSchedule.forEach(row => {
        if (!seenEmployees.has(row.name)) {
            seenEmployees.add(row.name);
            employeesInOrder.push(row.name);
        }
    });

    for (let name of employeesInOrder) {
        if (!breaks[name]) breaks[name] = [];

        const shiftStart = shifts[name][0];
        const shiftEnd = shifts[name][1];
        const shiftDuration = shiftEnd - shiftStart;

        // California law for rest periods:
        // 0:00-3:29: no rest break
        // 3:30-5:59: one rest break
        // 6:00-9:59: two rest breaks
        // 10:00-13:59: three rest breaks

        let restBreaksNeeded = 0;
        if (shiftDuration >= 210) restBreaksNeeded = 1;   // >= 3:30
        if (shiftDuration >= 360) restBreaksNeeded = 2;   // >= 6:00
        if (shiftDuration >= 600) restBreaksNeeded = 3;   // >= 10:00

        // Find employee's dept/subdept
        const empRow = newSchedule.find(row => row.name === name);
        if (!empRow) continue;

        const dept = empRow.dept;
        const subdept = empRow.job;

        // Schedule first rest period (+2 hours into shift)
        if (restBreaksNeeded >= 1) {
            const idealFirstBreak = shiftStart + 120; // +2:00 hours

            if (!isBreakAdjustmentEnabled(dept, subdept)) {
                // No conflict resolution
                breaks[name][0] = idealFirstBreak;
            } else {
                // Coverage-based optimization: find the time that maximizes minimum coverage
                // Preference order: ideal (+2:00), then +15, then +30, then -15
                const possibleTimes = [
                    idealFirstBreak,           // +2:00 (ideal)
                    idealFirstBreak + 15,      // +2:15
                    idealFirstBreak + 30,      // +2:30
                    idealFirstBreak - 15       // +1:45
                ];

                let bestTime = idealFirstBreak;
                let bestMinCoverage = -1;
                let bestIndex = 0;

                for (let i = 0; i < possibleTimes.length; i++) {
                    const candidateTime = possibleTimes[i];
                    if (candidateTime < shiftStart || candidateTime + 15 > shiftEnd) continue;

                    // Simulate taking this break and calculate minimum coverage during the break
                    const tempBreaks = { ...breaks };
                    if (!tempBreaks[name]) tempBreaks[name] = [];
                    tempBreaks[name] = [...(breaks[name] || [])];
                    tempBreaks[name][0] = candidateTime;

                    // Calculate coverage map with this break scheduled
                    const tempCoverage = calculateCoverageMap(newSchedule, shifts, tempBreaks);

                    // Find minimum coverage in the broader time window around the break
                    // This accounts for employees starting/ending shifts who could provide coverage
                    // Look from break start through break end and surrounding times
                    let minCoverage = Infinity;
                    for (let t = candidateTime; t < candidateTime + 15; t += 15) {
                        const coworkers = getCoworkersAtTime(tempCoverage, t, dept, subdept);
                        minCoverage = Math.min(minCoverage, coworkers.length);
                    }

                    console.log(`  [EVAL] ${name}: ${minutesToTime(candidateTime)} → coverage ${minCoverage}`);

                    // Choose the time that maximizes the minimum coverage
                    // Ties go to earlier times in the preference order (ideal=0 > +15=1 > +30=2 > -15=3)
                    if (minCoverage > bestMinCoverage || (minCoverage === bestMinCoverage && i < bestIndex)) {
                        bestMinCoverage = minCoverage;
                        bestTime = candidateTime;
                        bestIndex = i;
                    }
                }

                breaks[name][0] = bestTime;

                if (bestTime !== idealFirstBreak) {
                    const offset = bestTime - idealFirstBreak;
                    console.log(`[REST STAGGER] ${name} (${dept}/${subdept}): first break adjusted from ${minutesToTime(idealFirstBreak)} to ${minutesToTime(bestTime)} (offset: ${offset > 0 ? '+' : ''}${offset}min, maintains min coverage of ${bestMinCoverage})`);
                }
            }
        }

        // Schedule second rest period (+2 hours after returning from meal)
        if (restBreaksNeeded >= 2 && breaks[name][1] !== undefined) {
            const idealSecondBreak = breaks[name][1] + 30 + 120; // meal end + 2 hours

            if (!isBreakAdjustmentEnabled(dept, subdept)) {
                breaks[name][2] = idealSecondBreak;
            } else {
                // Coverage-based optimization: find the time that maximizes minimum coverage
                const possibleTimes = [
                    idealSecondBreak,           // +2:00 after meal (ideal)
                    idealSecondBreak + 15,      // +2:15 after meal
                    idealSecondBreak + 30,      // +2:30 after meal
                    idealSecondBreak - 15       // +1:45 after meal
                ];

                let bestTime = idealSecondBreak;
                let bestMinCoverage = -1;
                let bestIndex = 0;

                for (let i = 0; i < possibleTimes.length; i++) {
                    const candidateTime = possibleTimes[i];
                    if (candidateTime < shiftStart || candidateTime + 15 > shiftEnd) continue;

                    // Simulate taking this break and calculate minimum coverage during the break
                    const tempBreaks = { ...breaks };
                    if (!tempBreaks[name]) tempBreaks[name] = [];
                    tempBreaks[name] = [...(breaks[name] || [])];
                    tempBreaks[name][2] = candidateTime;

                    // Calculate coverage map with this break scheduled
                    const tempCoverage = calculateCoverageMap(newSchedule, shifts, tempBreaks);

                    // Find minimum coverage during this break period for this dept/subdept
                    let minCoverage = Infinity;
                    for (let t = candidateTime; t < candidateTime + 15; t += 15) {
                        const coworkers = getCoworkersAtTime(tempCoverage, t, dept, subdept);
                        minCoverage = Math.min(minCoverage, coworkers.length);
                    }

                    // Choose the time that maximizes the minimum coverage
                    // Ties go to earlier times in the preference order (ideal=0 > +15=1 > +30=2 > -15=3)
                    if (minCoverage > bestMinCoverage || (minCoverage === bestMinCoverage && i < bestIndex)) {
                        bestMinCoverage = minCoverage;
                        bestTime = candidateTime;
                        bestIndex = i;
                    }
                }

                breaks[name][2] = bestTime;

                if (bestTime !== idealSecondBreak) {
                    const offset = bestTime - idealSecondBreak;
                    console.log(`[REST STAGGER] ${name} (${dept}/${subdept}): second break adjusted from ${minutesToTime(idealSecondBreak)} to ${minutesToTime(bestTime)} (offset: ${offset > 0 ? '+' : ''}${offset}min, maintains min coverage of ${bestMinCoverage})`);
                }
            }
        }

        // Third rest period (for very long shifts)
        if (restBreaksNeeded >= 3 && breaks[name][2] !== undefined) {
            // Schedule +2 hours after second break
            const idealThirdBreak = breaks[name][2] + 15 + 120;
            breaks[name][3] = idealThirdBreak;
        }
    }

    // Post-processing: swap breaks to preserve schedule order where coverage is identical
    // This handles cases where employee A appears before employee B in the schedule,
    // but B gets an earlier break time due to coverage optimization

    // Group employees by dept/subdept in schedule order
    const deptEmployees = {};
    employeesInOrder.forEach(name => {
        const empRow = newSchedule.find(row => row.name === name);
        if (!empRow) return;

        const deptKey = `${empRow.dept}|${empRow.job}`;
        if (!deptEmployees[deptKey]) {
            deptEmployees[deptKey] = [];
        }
        deptEmployees[deptKey].push(name);
    });

    // Check each dept/subdept for potential swaps
    for (let deptKey in deptEmployees) {
        const [dept, subdept] = deptKey.split('|');
        if (!isBreakAdjustmentEnabled(dept, subdept)) continue;

        const employees = deptEmployees[deptKey];

        // Check each break type (first rest = 0, second rest = 2)
        for (let breakIndex of [0, 2]) {
            // Compare each pair of employees in schedule order
            for (let i = 0; i < employees.length; i++) {
                for (let j = i + 1; j < employees.length; j++) {
                    const empA = employees[i];
                    const empB = employees[j];

                    // Skip if either doesn't have this break
                    if (!breaks[empA] || !breaks[empB]) continue;
                    if (breaks[empA][breakIndex] === undefined || breaks[empB][breakIndex] === undefined) continue;

                    const timeA = breaks[empA][breakIndex];
                    const timeB = breaks[empB][breakIndex];

                    // If A appears before B in schedule but has a later break time, consider swapping
                    if (timeA > timeB) {
                        // Swap and check if coverage remains identical
                        const originalBreaks = JSON.parse(JSON.stringify(breaks));
                        breaks[empA][breakIndex] = timeB;
                        breaks[empB][breakIndex] = timeA;

                        const coverageOriginal = calculateCoverageMap(newSchedule, shifts, originalBreaks);
                        const coverageSwapped = calculateCoverageMap(newSchedule, shifts, breaks);

                        // Check if coverage is identical at all times during operating hours
                        let coverageIdentical = true;
                        for (let t = startOfDay; t <= endOfDay; t += 15) {
                            const origCoworkers = getCoworkersAtTime(coverageOriginal, t, dept, subdept);
                            const swapCoworkers = getCoworkersAtTime(coverageSwapped, t, dept, subdept);
                            if (origCoworkers.length !== swapCoworkers.length) {
                                coverageIdentical = false;
                                break;
                            }
                        }

                        if (coverageIdentical) {
                            const breakName = breakIndex === 0 ? 'first rest' : 'second rest';
                            console.log(`[BREAK SWAP] ${empA} and ${empB} (${dept}/${subdept}): swapped ${breakName} breaks (${minutesToTime(timeA)} ↔ ${minutesToTime(timeB)}) to preserve schedule order with identical coverage`);
                        } else {
                            // Revert the swap
                            breaks[empA][breakIndex] = timeA;
                            breaks[empB][breakIndex] = timeB;
                        }
                    }
                }
            }
        }
    }

    // Calculate and log coverage maps for debugging
    const coverageBefore = calculateCoverageMap(newSchedule, shifts, {});
    const coverageAfter = calculateCoverageMap(newSchedule, shifts, breaks);

    console.log("Coverage optimization complete");
    console.log(`Operating hours: ${minutesToTime(startOfDay)} - ${minutesToTime(endOfDay)}`);
    console.log("Sample coverage (4:00 PM before breaks):", coverageBefore[16 * 60]);
    console.log("Sample coverage (4:00 PM after breaks):", coverageAfter[16 * 60]);

    // Return breaks and segments for writing back to the sheet
    return { breaks, segments };
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

    // Fallback on today's date in ISO format
    const today = new Date();
    return today.toISOString().split('T')[0]; // Returns YYYY-MM-DD
}

// Get operating hours for a specific date based on day of week
function getOperatingHoursForDate(dateString) {
    // Parse the date string (YYYY-MM-DD)
    const date = new Date(dateString + 'T00:00:00'); // Add time to avoid timezone issues
    const dayOfWeek = date.getDay(); // 0 = Sunday, 1 = Monday, ..., 6 = Saturday

    const dayNames = ['sunday', 'monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday'];
    const dayName = dayNames[dayOfWeek];

    // Get the start and end times from the input fields
    const startInput = document.getElementById(`${dayName}Start`);
    const endInput = document.getElementById(`${dayName}End`);

    const startTime = startInput ? timeToMinutes(startInput.value) : 9 * 60; // Default 9 AM
    const endTime = endInput ? timeToMinutes(endInput.value) : 21 * 60; // Default 9 PM

    return { startTime, endTime };
}

// Read main/subdepartment options from UI
function getSelectedDepartmentPairs() {
    const pairs = [];
    const departmentCheckboxes = document.querySelectorAll(".dept-checkbox");

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
