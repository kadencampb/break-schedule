// ====================================================================================
// BREAK SCHEDULER CORE MODULE
// Version: 1.0.0
// Can be used in both web browsers and Excel Office Scripts
// ====================================================================================

// ====================================================================================
// CONFIGURATION AND CONSTANTS
// ====================================================================================

const DEPARTMENT_REGISTRY = {
    "Frontline": [
        "Cashier",
        "Cashier Bldg 2",
        "Customer Service",
        "Customer Service Bldg 2",
        "Greeter",
        "Greeter Bldg 2",
        "Order Pick Up",
        "Order Pick Up Bldg 2"
    ],
    "Hardgoods": [
        "Action Sports",
        "Camping",
        "Climbing",
        "Cycling",
        "Hardgoods",
        "Nordic",
        "Optics",
        "Outfitter",
        "Packs",
        "Paddling",
        "Racks",
        "Rentals",
        "Ski",
        "Snow Clothing",
        "Snow Sports"
    ],
    "Softgoods": [
        "Childrenswear",
        "Clothing",
        "Fitting Room",
        "Footwear",
        "Mens Clothing",
        "Outfitter",
        "Softgoods",
        "Womens Clothing"
    ],
    "Office": [
        "Banker",
        "Office"
    ],
    "Order Fulfillment": [
        "Order Fulfillment",
        "Order Fulfillment Bldg 2"
    ],
    "Product Movement": [
        "Action Sports Stock",
        "Camping Stock",
        "Clothing Stock",
        "Cycling Stock",
        "Footwear Stock",
        "Hardgoods Stock",
        "Ops Stock",
        "Ops Stock Bldg 2",
        "Ship Recv",
        "Ship Recv Bldg 2",
        "Snow Sports Stock",
        "Softgoods Stock",
        "Stocking"
    ],
    "Shop": [
        "Assembler",
        "Service Advisor",
        "Ski Shop"
    ],
    "Mgmt Retail": [
        "Key Holder",
        "Key Holder Bldg 2",
        "Leader on Duty",
        "Management",
        "Management Bldg 2"
    ]
};

// Default coverage optimization groups
const DEFAULT_GROUPS = [
    {
        id: 1,
        name: "Building 2 Cross-trained",
        departments: [
            {main: "Hardgoods", sub: "Action Sports"},
            {main: "Hardgoods", sub: "Rentals"},
            {main: "Frontline", sub: "Cashier Bldg 2"},
            {main: "Frontline", sub: "Customer Service Bldg 2"}
        ]
    },
    {
        id: 2,
        name: "Camping",
        departments: [
            {main: "Hardgoods", sub: "Camping"},
            {main: "Hardgoods", sub: "Hardgoods"}
        ]
    },
    {
        id: 3,
        name: "Clothing",
        departments: [
            {main: "Softgoods", sub: "Clothing"},
            {main: "Softgoods", sub: "Softgoods"},
            {main: "Softgoods", sub: "Fitting Room"},
            {main: "Softgoods", sub: "Mens Clothing"},
            {main: "Softgoods", sub: "Womens Clothing"},
            {main: "Softgoods", sub: "Outfitter"},
            {main: "Softgoods", sub: "Childrenswear"}
        ]
    },
    {
        id: 4,
        name: "Footwear",
        departments: [{main: "Softgoods", sub: "Footwear"}]
    },
    {
        id: 5,
        name: "Cashier",
        departments: [{main: "Frontline", sub: "Cashier"}]
    },
    {
        id: 6,
        name: "Customer Service",
        departments: [{main: "Frontline", sub: "Customer Service"}]
    },
    {
        id: 7,
        name: "Order Fulfillment",
        departments: [
            {main: "Order Fulfillment", sub: "Order Fulfillment"},
            {main: "Order Fulfillment", sub: "Order Fulfillment Bldg 2"}
        ]
    },
    {
        id: 8,
        name: "Stocking",
        departments: [
            {main: "Product Movement", sub: "Softgoods Stock"},
            {main: "Product Movement", sub: "Stocking"},
            {main: "Product Movement", sub: "Snow Sports Stock"},
            {main: "Product Movement", sub: "Ops Stock Bldg 2"},
            {main: "Product Movement", sub: "Ops Stock"},
            {main: "Product Movement", sub: "Hardgoods Stock"},
            {main: "Product Movement", sub: "Footwear Stock"},
            {main: "Product Movement", sub: "Cycling Stock"},
            {main: "Product Movement", sub: "Clothing Stock"},
            {main: "Product Movement", sub: "Camping Stock"},
            {main: "Product Movement", sub: "Action Sports Stock"}
        ]
    },
    {
        id: 9,
        name: "Service Advisor",
        departments: [{main: "Shop", sub: "Service Advisor"}]
    },
    {
        id: 10,
        name: "Management",
        departments: [
            {main: "Mgmt Retail", sub: "Management"},
            {main: "Mgmt Retail", sub: "Management Bldg 2"}
        ]
    }
];

// Default advanced settings
const DEFAULT_ADVANCED_SETTINGS = {
    maxEarly: 15,            // Max minutes to schedule any break early
    maxDelay: 30,            // Max minutes to delay any break
    deptWeightMultiplier: 4, // How much to prioritize same-department coverage
    proximityWeight: 1       // Weight for proximity to ideal time
};

// Default operating hours (9 AM - 9 PM)
const DEFAULT_OPERATING_HOURS = {
    startTime: 9 * 60,   // 9:00 AM in minutes
    endTime: 21 * 60     // 9:00 PM in minutes
};

// ====================================================================================
// HELPER FUNCTIONS
// ====================================================================================

/**
 * Convert time string (e.g., "2:30PM" or "14:30") to minutes since midnight
 */
function timeToMinutes(time) {
    if (!time) return 0;
    return (time.includes("AM") || time.includes("PM")
        ? ((+time.split(":")[0] % 12) + (time.includes("PM") ? 12 : 0)) * 60 + +time.split(":")[1].slice(0, 2)
        : +time.split(":")[0] * 60 + +time.split(":")[1]);
}

/**
 * Convert minutes since midnight to time string (e.g., "2:30PM")
 */
function minutesToTime(minutes) {
    const h = Math.floor(minutes / 60) % 12 || 12;
    const m = (minutes % 60).toString().padStart(2, "0");
    return `${h}:${m}${minutes >= 720 ? "PM" : "AM"}`;
}

/**
 * Format employee name (Last, First → First Last)
 */
function formatName(name) {
    if (!name) return "";
    const parts = name.split(",").map(s => s.trim());
    if (parts.length === 2) {
        return `${parts[1]} ${parts[0]}`;
    }
    return name;
}

/**
 * Find which group (if any) contains the given department/subdepartment
 */
function findGroupContaining(mainDept, subDept, groups = DEFAULT_GROUPS) {
    return groups.find(group =>
        group.departments.some(d => d.main === mainDept && d.sub === subDept)
    );
}

// ====================================================================================
// CORE SCHEDULING ALGORITHM
// ====================================================================================

/**
 * Calculate coverage map for all 15-minute intervals during operating hours
 * Returns a map where keys are time in minutes, values are arrays of employee objects working
 */
function calculateCoverageMap(schedule, shifts, breaks, startOfDay, endOfDay) {
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
 * If the dept/subdept belongs to a group, returns all employees in ALL departments within that group
 */
function getCoworkersAtTime(coverage, time, dept, subdept, groups = DEFAULT_GROUPS) {
    if (!coverage[time]) return [];

    // Check if this department belongs to a group
    const group = findGroupContaining(dept, subdept, groups);

    if (group) {
        // Return all employees in ANY department within this group
        return coverage[time].filter(emp =>
            group.departments.some(d =>
                d.main === emp.dept && d.sub === emp.subdept
            )
        );
    } else {
        // Original behavior - only employees in the same dept/subdept
        return coverage[time].filter(emp => emp.dept === dept && emp.subdept === subdept);
    }
}

/**
 * Find the optimal time for a break by evaluating coverage at different candidate times
 * Consolidates the logic used for meal breaks, first rest breaks, and second rest breaks
 */
function findOptimalBreakTime(params) {
    const {
        empName,
        idealTime,
        breakDuration, // 15 for rest, 30 for meal
        breakIndex, // 0, 1, or 2
        shiftStart,
        shiftEnd,
        dept,
        subdept,
        group,
        breaks,
        newSchedule,
        shifts,
        startOfDay,
        endOfDay,
        advSettings,
        log
    } = params;

    // Generate possible times based on advanced settings
    const possibleTimes = [idealTime];
    for (let delay = 15; delay <= advSettings.maxDelay; delay += 15) {
        possibleTimes.push(idealTime + delay);
    }
    for (let early = 15; early <= advSettings.maxEarly; early += 15) {
        possibleTimes.push(idealTime - early);
    }

    // Filter to only valid times within shift
    const validTimes = possibleTimes.filter(t =>
        t >= shiftStart && t + breakDuration <= shiftEnd
    );

    let bestTime = idealTime;
    let bestScore = -Infinity;

    for (const testTime of validTimes) {
        // Create temporary breaks with this test time
        const tempBreaks = JSON.parse(JSON.stringify(breaks));
        tempBreaks[empName] = tempBreaks[empName] || [];
        tempBreaks[empName][breakIndex] = testTime;

        // Calculate coverage map with this break time
        const coverageMap = calculateCoverageMap(newSchedule, shifts, tempBreaks, startOfDay, endOfDay);

        // Calculate minimum coverage during the break window
        let minDeptCoverage = Infinity;
        let minGroupCoverage = Infinity;

        for (let time = testTime; time < testTime + breakDuration; time += 15) {
            const coworkers = coverageMap[time] || [];

            // Count same-department coverage
            const deptCoverage = coworkers.filter(c => c.dept === dept && c.subdept === subdept).length;
            minDeptCoverage = Math.min(minDeptCoverage, deptCoverage);

            // Count group coverage (all departments in the group)
            const groupCoverage = group
                ? coworkers.filter(c =>
                    group.departments.some(d => d.main === c.dept && d.sub === c.subdept)
                ).length
                : 0;
            minGroupCoverage = Math.min(minGroupCoverage, groupCoverage);
        }

        // Calculate score: prioritize department coverage, then group coverage
        const score = (minDeptCoverage * advSettings.deptWeightMultiplier) + minGroupCoverage;

        // Add proximity bonus for being closer to ideal time
        const maxDistance = Math.max(advSettings.maxEarly, advSettings.maxDelay);
        const maxIntervals = maxDistance / 15;
        const distanceFromIdeal = Math.abs(testTime - idealTime);
        const intervalsAway = distanceFromIdeal / 15;
        const proximityBonus = Math.max(0, advSettings.proximityWeight * (maxIntervals - intervalsAway));

        const finalScore = score + proximityBonus;

        log(`  [EVAL] ${empName}: ${minutesToTime(testTime)} → dept coverage ${minDeptCoverage}, group coverage ${minGroupCoverage}, score ${score}, final ${finalScore.toFixed(2)}`);

        if (finalScore > bestScore) {
            bestScore = finalScore;
            bestTime = testTime;
        }
    }

    return { bestTime, bestScore };
}

/**
 * Main scheduling function - processes schedule data and returns optimized breaks
 *
 * @param {Array} schedule - Schedule data (rows from Excel)
 * @param {Object} options - Configuration options
 * @param {Object} options.operatingHours - {startTime: minutes, endTime: minutes}
 * @param {Array} options.groups - Coverage optimization groups
 * @param {Object} options.advancedSettings - Coverage optimization settings
 * @param {boolean} options.enableLogging - Whether to log debug info (default: true)
 * @returns {Object} - {breaks, segments, schedule, shifts}
 */
function scheduleBreaks(schedule, options = {}) {
    // Extract options with defaults
    const operatingHours = options.operatingHours || DEFAULT_OPERATING_HOURS;
    const groups = options.groups || DEFAULT_GROUPS;
    const advSettings = options.advancedSettings || DEFAULT_ADVANCED_SETTINGS;
    const enableLogging = options.enableLogging !== false; // Default to true
    const dataStart = options.dataStart !== undefined ? options.dataStart : 7; // Use dataStart if provided, otherwise default to 7
    const shiftCol = options.shiftColumnIndex !== undefined ? options.shiftColumnIndex : 4; // Column index for shift data

    const startOfDay = operatingHours.startTime;
    const endOfDay = operatingHours.endTime;

    let dept, newSchedule = [];
    const segments = []; // Track segments with row indices for writing back

    // Logging helper
    const log = enableLogging ? console.log : () => {};

    // Loop through each row starting from dataStart
    for (let i = dataStart; i < schedule.length; i++) {
        // If the first column contains a department name, assign it to "dept" and skip further processing for that row
        if (schedule[i][0]) {
            dept = schedule[i][0];
            continue;
        }

        // If there is no name, skip that row
        if (!schedule[i][2]) {
            continue;
        }

        let name = formatName(schedule[i][2]);

        // Split the time interval into start and end times, and convert them to minutes
        let interval = schedule[i][shiftCol]?.split("-") || [];
        interval = interval.map(bound => timeToMinutes(bound));
        const intervalStr = schedule[i][shiftCol]; // Keep original string for display

        // Add each row to the schedule with the department, job, name, and interval
        newSchedule.push({
            dept: dept,
            job: schedule[i][1],
            name: name,
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

    // Calculate the earliest and latest times for each person's shift
    let shifts = {};
    newSchedule.forEach(row => {
        if (!(row.name in shifts)) shifts[row.name] = [1440, 0];
        shifts[row.name][0] = Math.min(shifts[row.name][0], row.interval[0]);
        shifts[row.name][1] = Math.max(shifts[row.name][1], row.interval[1]);
    });

    // ====================================================================================
    // STEP 1: SCHEDULE MEAL PERIODS (CALIFORNIA LAW COMPLIANT)
    // ====================================================================================

    // Initialize breaks object
    let breaks = {};

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
                mealsNeeded: mealsNeeded
            });
        }
    });

    // Schedule first meal period at ideal time (+4 hours), handling conflicts greedily
    for (let emp of employeesNeedingMeals) {
        if (!breaks[emp.name]) breaks[emp.name] = [];

        // Find the employee's dept/subdept
        const empRow = newSchedule.find(row => row.name === emp.name);
        if (!empRow) continue;

        const dept = empRow.dept;
        const subdept = empRow.job;

        // Check if this department is in a group
        const group = findGroupContaining(dept, subdept, groups);

        // Skip if coverage optimization is disabled (not in any group)
        if (!group) {
            // Just schedule at ideal time without conflict resolution
            breaks[emp.name][1] = emp.shiftStart + 240; // +4:00 hours
            continue;
        }

        // Ideal meal time: +4 hours into shift
        const idealMealTime = emp.shiftStart + 240;

        log(`[MEAL DEBUG] ${emp.name} (${subdept}): ideal time ${minutesToTime(idealMealTime)}`);

        // Find optimal meal time using coverage optimization
        const { bestTime, bestScore } = findOptimalBreakTime({
            empName: emp.name,
            idealTime: idealMealTime,
            breakDuration: 30,
            breakIndex: 1,
            shiftStart: emp.shiftStart,
            shiftEnd: emp.shiftEnd,
            dept,
            subdept,
            group,
            breaks,
            newSchedule,
            shifts,
            startOfDay,
            endOfDay,
            advSettings,
            log
        });

        breaks[emp.name][1] = bestTime;

        if (bestTime === idealMealTime) {
            log(`[MEAL SCHEDULE] ${emp.name} (${subdept}): lunch scheduled at ${minutesToTime(idealMealTime)}`);
        } else {
            log(`[MEAL OPTIMIZE] ${emp.name} (${subdept}): lunch optimized from ${minutesToTime(idealMealTime)} to ${minutesToTime(bestTime)} for coverage (score: ${bestScore})`);
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

        // Calculate hours WORKED (shift duration minus meal breaks)
        // Meal breaks are unpaid, so they don't count toward hours worked
        let hoursWorked = shiftDuration;
        if (breaks[name] && breaks[name][1] !== undefined) {
            // Subtract 30 minutes for meal break
            hoursWorked = shiftDuration - 30;
        }

        // California law for rest periods based on hours WORKED:
        // 0:00-3:29: no rest break
        // 3:30-6:00: one rest break
        // 6:01-9:59: two rest breaks
        // 10:00-13:59: three rest breaks

        let restBreaksNeeded = 0;
        if (hoursWorked >= 210) restBreaksNeeded = 1;   // >= 3:30
        if (hoursWorked > 360) restBreaksNeeded = 2;    // > 6:00
        if (hoursWorked >= 600) restBreaksNeeded = 3;   // >= 10:00

        // Find employee's dept/subdept
        const empRow = newSchedule.find(row => row.name === name);
        if (!empRow) continue;

        const dept = empRow.dept;
        const subdept = empRow.job;

        // Check if this department is in a group
        const group = findGroupContaining(dept, subdept, groups);

        // Schedule first rest period (+2 hours into shift)
        if (restBreaksNeeded >= 1) {
            const idealFirstBreak = shiftStart + 120; // +2:00 hours

            if (!group) {
                // No coverage optimization (not in any group)
                breaks[name][0] = idealFirstBreak;
            } else {
                // Find optimal time using coverage optimization
                const { bestTime, bestScore } = findOptimalBreakTime({
                    empName: name,
                    idealTime: idealFirstBreak,
                    breakDuration: 15,
                    breakIndex: 0,
                    shiftStart,
                    shiftEnd,
                    dept,
                    subdept,
                    group,
                    breaks,
                    newSchedule,
                    shifts,
                    startOfDay,
                    endOfDay,
                    advSettings,
                    log
                });

                breaks[name][0] = bestTime;

                if (bestTime !== idealFirstBreak) {
                    const offset = bestTime - idealFirstBreak;
                    log(`[REST OPTIMIZE] ${name} (${subdept}): first break optimized from ${minutesToTime(idealFirstBreak)} to ${minutesToTime(bestTime)} for coverage (offset: ${offset > 0 ? '+' : ''}${offset}min, score: ${bestScore})`);
                }
            }
        }

        // Schedule second rest period (6.5 hours into shift - splits second half of workday)
        if (restBreaksNeeded >= 2 && breaks[name][1] !== undefined) {
            const idealSecondBreak = shiftStart + 390; // 6.5 hours into shift

            if (!group) {
                // No coverage optimization (not in any group)
                breaks[name][2] = idealSecondBreak;
            } else {
                // Find optimal time using coverage optimization
                const { bestTime } = findOptimalBreakTime({
                    empName: name,
                    idealTime: idealSecondBreak,
                    breakDuration: 15,
                    breakIndex: 2,
                    shiftStart,
                    shiftEnd,
                    dept,
                    subdept,
                    group,
                    breaks,
                    newSchedule,
                    shifts,
                    startOfDay,
                    endOfDay,
                    advSettings,
                    log
                });

                breaks[name][2] = bestTime;
            }
        }

        // Third rest period (for very long shifts)
        if (restBreaksNeeded >= 3 && breaks[name][2] !== undefined) {
            // Schedule +2 hours after second break
            const idealThirdBreak = breaks[name][2] + 15 + 120;
            breaks[name][3] = idealThirdBreak;
        }
    }

    // ====================================================================================
    // STEP 3: POST-PROCESSING - SWAP BREAKS TO PRESERVE SCHEDULE ORDER
    // ====================================================================================

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
        const group = findGroupContaining(dept, subdept, groups);
        if (!group) continue; // Skip if not in any group

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
                        // First, validate that swapped times would be valid for both employees
                        const empARow = newSchedule.find(row => row.name === empA);
                        const empBRow = newSchedule.find(row => row.name === empB);

                        if (!empARow || !empBRow) continue;

                        const breakDuration = breakIndex === 0 || breakIndex === 2 ? 15 : 30;

                        // Check if swapped times are within both employees' shifts
                        const timeB_validForA = timeB >= shifts[empA][0] && (timeB + breakDuration) <= shifts[empA][1];
                        const timeA_validForB = timeA >= shifts[empB][0] && (timeA + breakDuration) <= shifts[empB][1];

                        if (!timeB_validForA || !timeA_validForB) {
                            // Swap would put a break outside someone's shift, skip it
                            continue;
                        }

                        // Swap and check if coverage remains identical
                        const originalBreaks = JSON.parse(JSON.stringify(breaks));
                        breaks[empA][breakIndex] = timeB;
                        breaks[empB][breakIndex] = timeA;

                        const coverageOriginal = calculateCoverageMap(newSchedule, shifts, originalBreaks, startOfDay, endOfDay);
                        const coverageSwapped = calculateCoverageMap(newSchedule, shifts, breaks, startOfDay, endOfDay);

                        // Check if coverage is identical at all times during operating hours
                        let coverageIdentical = true;
                        for (let t = startOfDay; t <= endOfDay; t += 15) {
                            const origCoworkers = getCoworkersAtTime(coverageOriginal, t, dept, subdept, groups);
                            const swapCoworkers = getCoworkersAtTime(coverageSwapped, t, dept, subdept, groups);
                            if (origCoworkers.length !== swapCoworkers.length) {
                                coverageIdentical = false;
                                break;
                            }
                        }

                        if (coverageIdentical) {
                            const breakName = breakIndex === 0 ? 'first rest' : 'second rest';
                            log(`[BREAK SWAP] ${empA} and ${empB} (${subdept}): swapped ${breakName} breaks (${minutesToTime(timeA)} ↔ ${minutesToTime(timeB)}) to preserve schedule order with identical coverage`);
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
    const coverageBefore = calculateCoverageMap(newSchedule, shifts, {}, startOfDay, endOfDay);
    const coverageAfter = calculateCoverageMap(newSchedule, shifts, breaks, startOfDay, endOfDay);

    log("Coverage optimization complete");
    log(`Operating hours: ${minutesToTime(startOfDay)} - ${minutesToTime(endOfDay)}`);
    log("Sample coverage (4:00 PM before breaks):", coverageBefore[16 * 60]);
    log("Sample coverage (4:00 PM after breaks):", coverageAfter[16 * 60]);

    // Return breaks and segments for writing back to the sheet
    return {
        breaks,
        segments,
        schedule: newSchedule,
        shifts
    };
}

// ====================================================================================
// EXPORTS
// ====================================================================================

// For browser environments (check if window exists)
if (typeof window !== 'undefined') {
    window.scheduleBreaks = scheduleBreaks;
    window.timeToMinutes = timeToMinutes;
    window.minutesToTime = minutesToTime;
    window.formatName = formatName;
    window.DEFAULT_GROUPS = DEFAULT_GROUPS;
    window.DEFAULT_ADVANCED_SETTINGS = DEFAULT_ADVANCED_SETTINGS;
    window.DEFAULT_OPERATING_HOURS = DEFAULT_OPERATING_HOURS;
}

// For Office Scripts or Node.js (if needed in future)
if (typeof module !== 'undefined' && module.exports) {
    module.exports = {
        scheduleBreaks,
        timeToMinutes,
        minutesToTime,
        formatName,
        DEFAULT_GROUPS,
        DEFAULT_ADVANCED_SETTINGS,
        DEFAULT_OPERATING_HOURS
    };
}
