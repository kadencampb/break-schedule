// ====================================================================================
// ADVANCED SETTINGS MANAGEMENT
// ====================================================================================
// Note: DEPARTMENT_REGISTRY, DEFAULT_GROUPS, and DEFAULT_ADVANCED_SETTINGS are defined in scheduler.js

// Load advanced settings from localStorage or use defaults
function loadAdvancedSettings() {
    const stored = localStorage.getItem('advancedSettings');
    if (stored) {
        try {
            return { ...DEFAULT_ADVANCED_SETTINGS, ...JSON.parse(stored) };
        } catch (e) {
            console.error('Failed to parse stored advanced settings:', e);
            return DEFAULT_ADVANCED_SETTINGS;
        }
    }
    return DEFAULT_ADVANCED_SETTINGS;
}

// Save advanced settings to localStorage
function saveAdvancedSettings(settings) {
    localStorage.setItem('advancedSettings', JSON.stringify(settings));
}

// Initialize advanced settings inputs from localStorage
function initializeAdvancedSettings() {
    const settings = loadAdvancedSettings();

    const fields = ['maxRestDelay', 'maxRestEarly', 'maxMealDelay', 'maxMealEarly', 'deptWeightMultiplier', 'proximityWeight'];
    fields.forEach(field => {
        const input = document.getElementById(field);
        if (input) {
            input.value = settings[field];
            // Update display value
            updateAdvancedSettingDisplay(field, settings[field]);
        }
    });
}

// Update the display value for an advanced setting
function updateAdvancedSettingDisplay(field, value) {
    const displayElement = document.getElementById(`${field}Value`);
    if (displayElement) {
        if (field === 'deptWeightMultiplier') {
            displayElement.textContent = value;
        } else if (field === 'proximityWeight') {
            displayElement.textContent = value;
        } else {
            displayElement.textContent = value;
        }
    }
}

// Save current advanced settings to localStorage
function saveCurrentAdvancedSettings() {
    const settings = {
        maxRestDelay: parseInt(document.getElementById('maxRestDelay')?.value || DEFAULT_ADVANCED_SETTINGS.maxRestDelay),
        maxRestEarly: parseInt(document.getElementById('maxRestEarly')?.value || DEFAULT_ADVANCED_SETTINGS.maxRestEarly),
        maxMealDelay: parseInt(document.getElementById('maxMealDelay')?.value || DEFAULT_ADVANCED_SETTINGS.maxMealDelay),
        maxMealEarly: parseInt(document.getElementById('maxMealEarly')?.value || DEFAULT_ADVANCED_SETTINGS.maxMealEarly),
        deptWeightMultiplier: parseInt(document.getElementById('deptWeightMultiplier')?.value || DEFAULT_ADVANCED_SETTINGS.deptWeightMultiplier),
        proximityWeight: parseInt(document.getElementById('proximityWeight')?.value || DEFAULT_ADVANCED_SETTINGS.proximityWeight)
    };

    saveAdvancedSettings(settings);
}

// Reset advanced settings to defaults
function resetAdvancedSettings() {
    if (confirm('Reset all advanced settings to their default values?')) {
        saveAdvancedSettings(DEFAULT_ADVANCED_SETTINGS);
        initializeAdvancedSettings();
        alert('Advanced settings reset to defaults!');
    }
}

// ====================================================================================
// STATE LABOR LAWS MANAGEMENT
// ====================================================================================

// Get selected state from localStorage or use default
function getSelectedState() {
    const stored = localStorage.getItem('selectedState');
    return stored || 'california';
}

// Save selected state to localStorage
function saveSelectedState(state) {
    localStorage.setItem('selectedState', state);
}

// Initialize state selection dropdown
function initializeStateSelection() {
    const state = getSelectedState();
    const stateSelect = document.getElementById('stateSelect');

    if (stateSelect) {
        stateSelect.value = state;
    }
}

// ====================================================================================
// OPERATING HOURS MANAGEMENT
// ====================================================================================
// Note: DEFAULT_OPERATING_HOURS is defined in scheduler.js

// Load operating hours from localStorage or use defaults
function loadOperatingHours() {
    const stored = localStorage.getItem('operatingHours');
    if (stored) {
        try {
            return JSON.parse(stored);
        } catch (e) {
            console.error('Failed to parse stored operating hours:', e);
        }
    }

    // Return default hours for all days (10:00 AM - 9:00 PM in 24-hour format)
    return {
        monday: { start: '10:00', end: '21:00' },
        tuesday: { start: '10:00', end: '21:00' },
        wednesday: { start: '10:00', end: '21:00' },
        thursday: { start: '10:00', end: '21:00' },
        friday: { start: '10:00', end: '21:00' },
        saturday: { start: '10:00', end: '21:00' },
        sunday: { start: '10:00', end: '21:00' }
    };
}

// Save operating hours to localStorage
function saveOperatingHours(hours) {
    localStorage.setItem('operatingHours', JSON.stringify(hours));
}

// Initialize operating hours inputs from localStorage
function initializeOperatingHours() {
    const hours = loadOperatingHours();

    Object.keys(hours).forEach(day => {
        const startInput = document.getElementById(`${day}Start`);
        const endInput = document.getElementById(`${day}End`);

        if (startInput) startInput.value = hours[day].start;
        if (endInput) endInput.value = hours[day].end;
    });
}

// Save current operating hours to localStorage
function saveCurrentOperatingHours() {
    const hours = {};
    const days = ['monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday', 'sunday'];

    days.forEach(day => {
        const startInput = document.getElementById(`${day}Start`);
        const endInput = document.getElementById(`${day}End`);

        if (startInput && endInput) {
            hours[day] = {
                start: startInput.value,
                end: endInput.value
            };
        }
    });

    saveOperatingHours(hours);
}

// ====================================================================================
// GROUP MANAGEMENT FUNCTIONS
// ====================================================================================

// Get groups from localStorage or initialize with defaults
function getGroups() {
    const stored = localStorage.getItem('coverageGroups');
    if (stored) {
        try {
            return JSON.parse(stored);
        } catch (e) {
            console.error('Failed to parse stored groups:', e);
            return DEFAULT_GROUPS;
        }
    }
    return DEFAULT_GROUPS;
}

// Save groups to localStorage
function saveGroups(groups) {
    localStorage.setItem('coverageGroups', JSON.stringify(groups));
}

// Initialize groups on first load
function initializeGroups() {
    const groups = getGroups();
    // If no groups in storage, save defaults
    if (!localStorage.getItem('coverageGroups')) {
        saveGroups(DEFAULT_GROUPS);
    }
    return groups;
}

// Find which group (if any) contains a specific department
function findGroupContaining(mainDept, subDept) {
    const groups = getGroups();
    return groups.find(group =>
        group.departments.some(d =>
            d.main === mainDept && d.sub === subDept
        )
    );
}

// Get all departments that are currently in any group
function getDepartmentsInGroups() {
    const groups = getGroups();
    const inGroups = [];
    groups.forEach(group => {
        group.departments.forEach(dept => {
            inGroups.push({main: dept.main, sub: dept.sub});
        });
    });
    return inGroups;
}

// Get all available departments not in any group
function getAvailableDepartments() {
    const inGroups = getDepartmentsInGroups();
    const available = [];

    for (let main in DEPARTMENT_REGISTRY) {
        DEPARTMENT_REGISTRY[main].forEach(sub => {
            const isInGroup = inGroups.some(d => d.main === main && d.sub === sub);
            if (!isInGroup) {
                available.push({main, sub});
            }
        });
    }

    return available;
}

// Add a new group
function addGroup(name, departments) {
    const groups = getGroups();
    const newId = groups.length > 0 ? Math.max(...groups.map(g => g.id)) + 1 : 1;
    const newGroup = {
        id: newId,
        name: name,
        departments: departments
    };
    groups.push(newGroup);
    saveGroups(groups);
    return newGroup;
}

// Delete a group by ID
function deleteGroup(groupId) {
    const groups = getGroups();
    const filtered = groups.filter(g => g.id !== groupId);
    saveGroups(filtered);
    return filtered;
}

// Update a group
function updateGroup(groupId, name, departments) {
    const groups = getGroups();
    const group = groups.find(g => g.id === groupId);
    if (group) {
        group.name = name;
        group.departments = departments;
        saveGroups(groups);
    }
    return groups;
}

// ====================================================================================
// UI RENDERING FUNCTIONS
// ====================================================================================

// Render all groups to the UI
function renderGroups() {
    const groups = getGroups();
    const container = document.getElementById('groupsContainer');

    if (!container) return;

    // Clear existing content
    container.innerHTML = '';

    // Render each group as a card
    groups.forEach(group => {
        const groupCard = document.createElement('div');
        groupCard.className = 'card mb-2';
        const deptCount = group.departments.length;
        groupCard.innerHTML = `
            <div class="card-body p-3">
                <div class="d-flex justify-content-between align-items-start">
                    <div class="flex-grow-1">
                        <h6 class="mb-2 font-weight-bold">
                            ${escapeHtml(group.name)}
                            <span class="badge badge-warning ml-2">${deptCount} dept${deptCount !== 1 ? 's' : ''}</span>
                        </h6>
                        <div class="small" style="color: #4a5568;">
                            ${group.departments.map(d =>
                                `<div>&bull; ${escapeHtml(d.main)} / ${escapeHtml(d.sub)}</div>`
                            ).join('')}
                        </div>
                    </div>
                    <div class="d-flex ml-3">
                        <button class="btn btn-sm btn-outline-info edit-group-btn mr-2" data-group-id="${group.id}">
                            <i class="fas fa-edit mr-1"></i>Edit
                        </button>
                        <button class="btn btn-sm btn-outline-danger delete-group-btn" data-group-id="${group.id}">
                            <i class="fas fa-trash mr-1"></i>Delete
                        </button>
                    </div>
                </div>
            </div>
        `;
        container.appendChild(groupCard);
    });

    // If no groups, show a message
    if (groups.length === 0) {
        container.innerHTML = '<p class="text-muted">No coverage optimization groups configured.</p>';
    }

    // Attach event listeners to Edit and Delete buttons
    document.querySelectorAll('.edit-group-btn').forEach(btn => {
        btn.addEventListener('click', (e) => {
            const groupId = parseInt(e.target.dataset.groupId);
            openEditGroupModal(groupId);
        });
    });

    document.querySelectorAll('.delete-group-btn').forEach(btn => {
        btn.addEventListener('click', (e) => {
            const groupId = parseInt(e.target.dataset.groupId);
            handleDeleteGroup(groupId);
        });
    });
}

// Helper to escape HTML
function escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}

// ====================================================================================
// MODAL MANAGEMENT
// ====================================================================================

// Track currently editing group (null for new group)
let currentEditingGroupId = null;

// Track selected departments in modal
let selectedDepartmentsInModal = [];

// Open modal for adding a new group
function openAddGroupModal() {
    currentEditingGroupId = null;
    selectedDepartmentsInModal = [];

    document.getElementById('groupModalLabel').textContent = 'Add Coverage Group';
    document.getElementById('groupNameInput').value = '';
    document.getElementById('departmentSearchInput').value = '';

    renderDepartmentPicker();
    updateSelectedDepartmentsDisplay();

    $('#groupModal').modal('show');
}

// Open modal for editing an existing group
function openEditGroupModal(groupId) {
    const groups = getGroups();
    const group = groups.find(g => g.id === groupId);

    if (!group) return;

    currentEditingGroupId = groupId;
    selectedDepartmentsInModal = [...group.departments]; // Clone the array

    document.getElementById('groupModalLabel').textContent = 'Edit Coverage Group';
    document.getElementById('groupNameInput').value = group.name;
    document.getElementById('departmentSearchInput').value = '';

    renderDepartmentPicker();
    updateSelectedDepartmentsDisplay();

    $('#groupModal').modal('show');
}

// Render the department picker accordion
function renderDepartmentPicker() {
    const accordion = document.getElementById('departmentAccordion');
    if (!accordion) return;

    accordion.innerHTML = '';

    // Get departments already in groups
    const departmentsInGroups = getDepartmentsInGroups();

    // Get current group's departments if editing
    let currentGroupDepts = [];
    if (currentEditingGroupId !== null) {
        const groups = getGroups();
        const currentGroup = groups.find(g => g.id === currentEditingGroupId);
        if (currentGroup) {
            currentGroupDepts = currentGroup.departments;
        }
    }

    // Build list of all departments with their availability status
    const allDepartments = [];
    for (let main in DEPARTMENT_REGISTRY) {
        DEPARTMENT_REGISTRY[main].forEach(sub => {
            const isInOtherGroup = departmentsInGroups.some(d =>
                d.main === main && d.sub === sub &&
                !currentGroupDepts.some(cd => cd.main === main && cd.sub === sub)
            );
            const groupContaining = isInOtherGroup ? findGroupContaining(main, sub) : null;

            allDepartments.push({
                main,
                sub,
                disabled: isInOtherGroup,
                groupName: groupContaining ? groupContaining.name : null
            });
        });
    }

    // Group all departments by main category
    const byCategory = {};
    allDepartments.forEach(dept => {
        if (!byCategory[dept.main]) {
            byCategory[dept.main] = [];
        }
        byCategory[dept.main].push(dept);
    });

    // Sort categories alphabetically
    const sortedCategories = Object.keys(byCategory).sort();

    // Build accordion
    sortedCategories.forEach((category, index) => {
        const cardId = `category-${index}`;
        const departments = byCategory[category].sort((a, b) => a.sub.localeCompare(b.sub));

        const card = document.createElement('div');
        card.className = 'card';
        card.innerHTML = `
            <div class="card-header p-2" id="heading-${cardId}">
                <button class="btn btn-link btn-sm btn-block text-left p-0" type="button"
                        data-toggle="collapse" data-target="#collapse-${cardId}"
                        aria-expanded="${index === 0}" aria-controls="collapse-${cardId}">
                    ${escapeHtml(category)} (${departments.length})
                </button>
            </div>
            <div id="collapse-${cardId}" class="collapse ${index === 0 ? 'show' : ''}"
                 aria-labelledby="heading-${cardId}" data-parent="#departmentAccordion">
                <div class="card-body p-2">
                    ${departments.map(dept => {
                        const isSelected = selectedDepartmentsInModal.some(d =>
                            d.main === category && d.sub === dept.sub
                        );
                        const labelText = dept.disabled
                            ? `${escapeHtml(dept.sub)} <em class="text-muted" style="font-size: 0.85em;">(in ${escapeHtml(dept.groupName)})</em>`
                            : escapeHtml(dept.sub);

                        return `
                            <div class="form-check">
                                <input type="checkbox" class="form-check-input dept-picker-checkbox"
                                       id="dept-${escapeHtml(category)}-${escapeHtml(dept.sub)}"
                                       data-main="${escapeHtml(category)}"
                                       data-sub="${escapeHtml(dept.sub)}"
                                       ${isSelected ? 'checked' : ''}
                                       ${dept.disabled ? 'disabled' : ''}>
                                <label class="form-check-label small ${dept.disabled ? 'text-muted' : ''}"
                                       for="dept-${escapeHtml(category)}-${escapeHtml(dept.sub)}"
                                       style="${dept.disabled ? 'cursor: not-allowed;' : ''}">
                                    ${labelText}
                                </label>
                            </div>
                        `;
                    }).join('')}
                </div>
            </div>
        `;
        accordion.appendChild(card);
    });

    // Attach event listeners to checkboxes
    document.querySelectorAll('.dept-picker-checkbox').forEach(checkbox => {
        checkbox.addEventListener('change', handleDepartmentCheckboxChange);
    });
}

// Filter departments based on search input
function filterDepartments(searchTerm) {
    const lowerSearch = searchTerm.toLowerCase().trim();

    // Get all department cards
    document.querySelectorAll('#departmentAccordion .card').forEach(card => {
        const checkboxes = card.querySelectorAll('.dept-picker-checkbox');

        let categoryHasMatch = false;

        checkboxes.forEach(checkbox => {
            const sub = checkbox.dataset.sub;
            const main = checkbox.dataset.main;
            const fullName = `${main} ${sub}`.toLowerCase();

            // Check if search term matches
            const matches = !lowerSearch || fullName.includes(lowerSearch) || sub.toLowerCase().includes(lowerSearch);

            // Show/hide the checkbox container
            const formCheck = checkbox.closest('.form-check');
            if (formCheck) {
                formCheck.style.display = matches ? 'block' : 'none';
                if (matches) categoryHasMatch = true;
            }
        });

        // Show/hide entire category if no matches
        card.style.display = categoryHasMatch ? 'block' : 'none';
    });
}

// Handle checkbox change in department picker
function handleDepartmentCheckboxChange(e) {
    const main = e.target.dataset.main;
    const sub = e.target.dataset.sub;

    if (e.target.checked) {
        // Add to selected
        if (!selectedDepartmentsInModal.some(d => d.main === main && d.sub === sub)) {
            selectedDepartmentsInModal.push({ main, sub });
        }
    } else {
        // Remove from selected
        selectedDepartmentsInModal = selectedDepartmentsInModal.filter(d =>
            !(d.main === main && d.sub === sub)
        );
    }

    updateSelectedDepartmentsDisplay();
}

// Update the selected departments display area
function updateSelectedDepartmentsDisplay() {
    const display = document.getElementById('selectedDepartments');
    if (!display) return;

    if (selectedDepartmentsInModal.length === 0) {
        display.innerHTML = '<span class="text-muted">No departments selected</span>';
    } else {
        display.innerHTML = selectedDepartmentsInModal.map(d => `
            <span class="badge badge-primary mr-1 mb-1">
                ${escapeHtml(d.main)} / ${escapeHtml(d.sub)}
                <button type="button" class="close ml-1" style="font-size: 1rem;"
                        data-main="${escapeHtml(d.main)}" data-sub="${escapeHtml(d.sub)}"
                        onclick="handleRemoveDepartmentBadge(this)">
                    &times;
                </button>
            </span>
        `).join('');
    }
}

// Handle removing a department from the badge display
function handleRemoveDepartmentBadge(button) {
    const main = button.dataset.main;
    const sub = button.dataset.sub;

    selectedDepartmentsInModal = selectedDepartmentsInModal.filter(d =>
        !(d.main === main && d.sub === sub)
    );

    // Uncheck the corresponding checkbox
    const checkbox = document.querySelector(
        `.dept-picker-checkbox[data-main="${main}"][data-sub="${sub}"]`
    );
    if (checkbox) {
        checkbox.checked = false;
    }

    updateSelectedDepartmentsDisplay();
}

// Handle saving the group (add or update)
function handleSaveGroup() {
    const name = document.getElementById('groupNameInput').value.trim();

    // Validation
    if (!name) {
        alert('Please enter a group name.');
        return;
    }

    if (selectedDepartmentsInModal.length === 0) {
        alert('Please select at least one department.');
        return;
    }

    if (currentEditingGroupId === null) {
        // Add new group
        addGroup(name, selectedDepartmentsInModal);
    } else {
        // Update existing group
        updateGroup(currentEditingGroupId, name, selectedDepartmentsInModal);
    }

    // Close modal and refresh display
    $('#groupModal').modal('hide');
    renderGroups();
}

// Handle deleting a group
function handleDeleteGroup(groupId) {
    const groups = getGroups();
    const group = groups.find(g => g.id === groupId);

    if (!group) return;

    if (confirm(`Are you sure you want to delete the group "${group.name}"?`)) {
        deleteGroup(groupId);
        renderGroups();
    }
}

// ====================================================================================
// EXPORT/IMPORT/RESET FUNCTIONS
// ====================================================================================

// Export groups to JSON file
function handleExportGroups() {
    const groups = getGroups();
    const dataStr = JSON.stringify(groups, null, 2);
    const blob = new Blob([dataStr], { type: 'application/json' });
    const url = URL.createObjectURL(blob);

    const link = document.createElement('a');
    link.href = url;
    link.download = 'coverage-groups.json';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
}

// Import groups from JSON file
function handleImportGroups() {
    const input = document.getElementById('importGroupsInput');
    if (input) {
        input.click();
    }
}

// Process imported file
function processImportedFile(file) {
    const reader = new FileReader();

    reader.onload = (e) => {
        try {
            const imported = JSON.parse(e.target.result);

            // Validate structure
            if (!Array.isArray(imported)) {
                alert('Invalid file format: Expected an array of groups.');
                return;
            }

            // Basic validation
            const isValid = imported.every(group =>
                group.id !== undefined &&
                group.name &&
                Array.isArray(group.departments) &&
                group.departments.every(d => d.main && d.sub)
            );

            if (!isValid) {
                alert('Invalid file format: Groups must have id, name, and departments array.');
                return;
            }

            // Confirm replacement
            if (confirm('This will replace all existing groups. Continue?')) {
                saveGroups(imported);
                renderGroups();
                alert('Groups imported successfully!');
            }
        } catch (err) {
            alert('Error parsing file: ' + err.message);
        }
    };

    reader.readAsText(file);
}

// Reset groups to defaults
function handleResetGroups() {
    if (confirm('This will reset all groups to the default configuration. All custom groups will be lost. Continue?')) {
        saveGroups(DEFAULT_GROUPS);
        renderGroups();
        alert('Groups reset to defaults!');
    }
}

// Initialize groups and operating hours on page load
document.addEventListener('DOMContentLoaded', () => {
    initializeGroups();
    renderGroups();
    initializeStateSelection();
    initializeOperatingHours();
    initializeAdvancedSettings();

    // Wire up Add Group button
    const addGroupBtn = document.getElementById('addGroupBtn');
    if (addGroupBtn) {
        addGroupBtn.addEventListener('click', openAddGroupModal);
    }

    // Wire up Save Group button in modal
    const saveGroupBtn = document.getElementById('saveGroupBtn');
    if (saveGroupBtn) {
        saveGroupBtn.addEventListener('click', handleSaveGroup);
    }

    // Wire up department search input
    const searchInput = document.getElementById('departmentSearchInput');
    if (searchInput) {
        searchInput.addEventListener('input', (e) => {
            filterDepartments(e.target.value);
        });
    }

    // Wire up Export button
    const exportBtn = document.getElementById('exportGroupsBtn');
    if (exportBtn) {
        exportBtn.addEventListener('click', handleExportGroups);
    }

    // Wire up Import button
    const importBtn = document.getElementById('importGroupsBtn');
    if (importBtn) {
        importBtn.addEventListener('click', handleImportGroups);
    }

    // Wire up Import file input
    const importInput = document.getElementById('importGroupsInput');
    if (importInput) {
        importInput.addEventListener('change', (e) => {
            if (e.target.files.length > 0) {
                processImportedFile(e.target.files[0]);
                e.target.value = ''; // Reset input
            }
        });
    }

    // Wire up Reset button
    const resetBtn = document.getElementById('resetGroupsBtn');
    if (resetBtn) {
        resetBtn.addEventListener('click', handleResetGroups);
    }

    // Wire up operating hours inputs to save on change
    document.querySelectorAll('.operating-hours-input').forEach(input => {
        input.addEventListener('change', saveCurrentOperatingHours);
    });

    // Wire up state selection to save on change
    const stateSelect = document.getElementById('stateSelect');
    if (stateSelect) {
        stateSelect.addEventListener('change', (e) => {
            saveSelectedState(e.target.value);
        });
    }

    // Wire up advanced settings inputs to save on change and update display
    const advancedFields = ['maxRestDelay', 'maxRestEarly', 'maxMealDelay', 'maxMealEarly', 'deptWeightMultiplier', 'proximityWeight'];
    advancedFields.forEach(field => {
        const input = document.getElementById(field);
        if (input) {
            // Update display on input (real-time as slider moves)
            input.addEventListener('input', (e) => {
                updateAdvancedSettingDisplay(field, e.target.value);
            });
            // Save to localStorage on change (when user releases slider)
            input.addEventListener('change', saveCurrentAdvancedSettings);
        }
    });

    // Wire up Reset Advanced Settings button
    const resetAdvancedBtn = document.getElementById('resetAdvancedBtn');
    if (resetAdvancedBtn) {
        resetAdvancedBtn.addEventListener('click', resetAdvancedSettings);
    }
});

// Trigger file input when the Select File button is clicked
document.getElementById("selectFileButton").addEventListener("click", () => {
    document.getElementById("fileInput").click();
});

// Store the file for later use
let selectedFile;

document.getElementById("fileInput").addEventListener("change", (event) => {
    selectedFile = event.target.files[0]; // Store the selected file

    // Update file name display
    const fileNameDisplay = document.getElementById("selectedFileName");
    if (selectedFile) {
        fileNameDisplay.textContent = selectedFile.name;
        fileNameDisplay.classList.remove("badge-secondary");
        fileNameDisplay.classList.add("badge-success");
    } else {
        fileNameDisplay.textContent = "No file selected";
        fileNameDisplay.classList.remove("badge-success");
        fileNameDisplay.classList.add("badge-secondary");
    }

    // Show the download button container after file selection
    const container = document.getElementById("downloadLinkContainer");
    container.classList.remove("d-none");
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

        // Check if this is a multi-day schedule file
        const dailySchedules = splitIntoDailySchedules(rowData);

        if (dailySchedules.length > 1) {
            // Process multiple daily schedules
            processMultipleDailySchedules(dailySchedules);
        } else {
            // Single schedule - process as before
            const scheduleDate = extractScheduleDate(rowData);
            const operatingHours = getOperatingHoursForDate(scheduleDate);

            processRowDataInPlace(sheet, rowData, operatingHours);
            applyReportStyling(sheet, rowData);

            const workbookBlob = XLSX.write(workbook, {
                bookType: "xlsx",
                type: "binary",
                cellStyles: true
            });
            const blob = new Blob([stringToArrayBuffer(workbookBlob)], { type: "application/octet-stream" });
            const url = URL.createObjectURL(blob);

            const hiddenDownloadLink = document.getElementById("hiddenDownloadLink");
            hiddenDownloadLink.href = url;
            hiddenDownloadLink.download = `Break Schedule ${scheduleDate}.xlsx`;
            hiddenDownloadLink.click();
        }
    };

    reader.readAsArrayBuffer(selectedFile);
});

// Split a multi-day schedule file into individual daily schedules
function splitIntoDailySchedules(rowData) {
    const dailySchedules = [];
    let currentSchedule = [];
    let currentDate = null;

    for (let i = 0; i < rowData.length; i++) {
        const row = rowData[i];
        const cell = row[0];

        // Check if this row contains a date marker
        if (typeof cell === "string") {
            const isoMatch = cell.match(/Date:\s*(\d{4}-\d{2}-\d{2})/);
            if (isoMatch) {
                // If we have a current schedule being built, save it
                if (currentSchedule.length > 0 && currentDate) {
                    dailySchedules.push({
                        date: currentDate,
                        rows: currentSchedule
                    });
                }
                // Start a new schedule
                currentDate = isoMatch[1];
                currentSchedule = [];
            }
        }

        // Add this row to the current schedule
        currentSchedule.push([...row]);
    }

    // Don't forget the last schedule
    if (currentSchedule.length > 0 && currentDate) {
        dailySchedules.push({
            date: currentDate,
            rows: currentSchedule
        });
    }

    return dailySchedules;
}

// Process multiple daily schedules and generate separate files
function processMultipleDailySchedules(dailySchedules) {
    console.log(`Processing ${dailySchedules.length} daily schedules...`);

    dailySchedules.forEach((schedule, index) => {
        // Create a new workbook for this daily schedule
        const wb = XLSX.utils.book_new();

        // Convert rows to sheet
        const ws = XLSX.utils.aoa_to_sheet(schedule.rows);

        // Get operating hours for this date
        const operatingHours = getOperatingHoursForDate(schedule.date);

        // Process breaks for this schedule
        processRowDataInPlace(ws, schedule.rows, operatingHours);

        // Apply styling
        applyReportStyling(ws, schedule.rows);

        // Add worksheet to workbook
        XLSX.utils.book_append_sheet(wb, ws, "Schedule");

        // Generate file
        const workbookBlob = XLSX.write(wb, {
            bookType: "xlsx",
            type: "binary",
            cellStyles: true
        });
        const blob = new Blob([stringToArrayBuffer(workbookBlob)], { type: "application/octet-stream" });
        const url = URL.createObjectURL(blob);

        // Download the file
        const link = document.createElement('a');
        link.href = url;
        link.download = `Break Schedule ${schedule.date}.xlsx`;
        link.style.display = 'none';
        document.body.appendChild(link);

        // Use setTimeout to allow multiple downloads
        setTimeout(() => {
            link.click();
            document.body.removeChild(link);
            URL.revokeObjectURL(url);
        }, index * 500); // Stagger downloads by 500ms
    });
}

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
    const dataStart = detectDataStart(rows);

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
    //    This guarantees blank cells actually exist and can get borders.
    for (let r = 0; r <= lastRow; r++) {
        for (let c = 0; c <= 6; c++) {  // only columns A–G (column D deleted)
            const cell = touchCell(r, c);
            // Each cell gets its own unique style object
            cell.s = {
                font: { name: "Arial", sz: 6.75, bold: false },
                alignment: { horizontal: "left", vertical: "top", wrapText: true }
            };
            // no border by default (will be added later for specific rows)
        }
    }

    // 2) Row 1 (Date row): A1-G1 merged, Arial 9pt bold left aligned
    for (let c = 0; c <= 6; c++) {
        const cell = touchCell(0, c); // A1-G1
        cell.s = {
            font: { name: "Arial", sz: 9, bold: true },
            alignment: { horizontal: "left", vertical: "center" }
        };
    }

    // 3) Row 2 (Location row): A2-G2 merged, Arial 9pt bold left aligned
    for (let c = 0; c <= 6; c++) {
        const cell = touchCell(1, c); // A2-G2
        cell.s = {
            font: { name: "Arial", sz: 9, bold: true },
            alignment: { horizontal: "left", vertical: "center" }
        };
    }

    // 4) Row 3 (Dept/Job header): Bold 7.5pt left aligned
    for (let c = 0; c <= 6; c++) {
        const cell = touchCell(2, c); // A3-G3
        cell.s = {
            font: { name: "Arial", sz: 7.5, bold: true },
            alignment: { horizontal: "left", vertical: "center" }
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

    // 8) Header row: Arial, 7.5 pt, bold, left, center, thin 000000 border
    (function () {
        const r = dataStart - 1; // The header row (row before data starts)
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

    // Set column widths (after deleting column D)
    // A: 6.14, B: 16, C-G: 13
    sheet["!cols"] = [
        { wch: 6.14 },  // A
        { wch: 16 },    // B
        { wch: 13 },    // C
        { wch: 13 },    // D (was E - Shift)
        { wch: 13 },    // E (was F - first break)
        { wch: 13 },    // F (was G - meal break)
        { wch: 13 }     // G (was H - second break)
    ];

    // Create merges array if it doesn't exist
    if (!sheet["!merges"]) {
        sheet["!merges"] = [];
    }

    // Merge A1:G1 for the title row (Date: YYYY-MM-DD)
    sheet["!merges"].push({ s: { r: 0, c: 0 }, e: { r: 0, c: 6 } });

    // Merge A2:G2 for Location row
    sheet["!merges"].push({ s: { r: 1, c: 0 }, e: { r: 1, c: 6 } });

    // Find the header row (row with "Name" in column C) and merge A:B for that row
    const headerRowIndex = dataStart - 1; // The row before data starts
    if (headerRowIndex >= 0) {
        sheet["!merges"].push({ s: { r: headerRowIndex, c: 0 }, e: { r: headerRowIndex, c: 1 } });
    }

    // Merge columns A and B for each department header row (rows with dept but no name)
    for (let r = dataStart; r <= lastRow; r++) {
        const row = rows[r] || [];
        const colA = row[0];
        const colC = row[2];

        const hasDept = colA != null && String(colA).trim() !== "";
        const hasName = colC != null && String(colC).trim() !== "";

        if (hasDept && !hasName) {
            // This is a department header row - merge A and B
            sheet["!merges"].push({ s: { r: r, c: 0 }, e: { r: r, c: 1 } });
        }
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

    // DELETE COLUMN D (shift label column) from BOTH the sheet AND the schedule array
    // First, delete from the schedule array
    const modifiedSchedule = schedule.map(row => {
        if (!row || row.length <= 4) return row; // Keep rows with 4 or fewer columns as-is
        // Remove column D (index 3) by creating a new array without it
        return [
            row[0],  // A: Dept
            row[1],  // B: Job
            row[2],  // C: Name
            // Skip row[3] - this is the shift label column D
            ...row.slice(4)  // D onwards: Shift (was E) and any other columns
        ];
    });

    // Now delete column D from the sheet
    const sheetRange = XLSX.utils.decode_range(sheet["!ref"] || "A1");
    const lastRow = sheetRange.e.r;
    const lastCol = sheetRange.e.c;

    // For each row, shift columns E onward to column D
    for (let r = 0; r <= lastRow; r++) {
        // Shift columns E through end to D through end-1
        for (let c = 4; c <= lastCol; c++) { // c=4 is column E
            const sourceRef = XLSX.utils.encode_cell({ r, c });
            const targetRef = XLSX.utils.encode_cell({ r, c: c - 1 }); // shift left by 1

            if (sheet[sourceRef]) {
                sheet[targetRef] = sheet[sourceRef];
            } else {
                delete sheet[targetRef]; // Remove if source doesn't exist
            }
        }
        // Delete the last column (which is now a duplicate)
        const lastColRef = XLSX.utils.encode_cell({ r, c: lastCol });
        delete sheet[lastColRef];
    }

    // Update the sheet range to reflect the removed column
    sheetRange.e.c = lastCol - 1;
    sheet["!ref"] = XLSX.utils.encode_range(sheetRange);

    // After deleting column D, the structure is now:
    // A: Dept, B: Job, C: Name, D: Shift (was E)
    // We add breaks in E, F, G (indices 4, 5, 6)
    if (headerRowIndex >= 0) {
        setCell(headerRowIndex, 4, "15");  // First rest break
        setCell(headerRowIndex, 5, "30");  // Meal break
        setCell(headerRowIndex, 6, "15");  // Second rest break
    }

    // Process the schedule data with operating hours
    // Get settings from DOM (will use defaults if not available)
    const groups = getGroups();
    const advSettings = loadAdvancedSettings();

    const result = scheduleBreaks(modifiedSchedule, {
        operatingHours: operatingHours,
        groups: groups,
        advancedSettings: advSettings,
        enableLogging: true,
        dataStart: dataStart,  // Pass the detected data start row
        shiftColumnIndex: 3  // After deleting column D, shift is now in column D (index 3)
    });
    const { breaks, segments } = result;

    // Track which employees have had their breaks written
    const printed = new Set();

    // Write breaks into the sheet for each segment row
    // After deleting column D, breaks go in columns E, F, G (indices 4, 5, 6)
    segments.forEach(seg => {
        const empName = seg.name;
        const rZero = seg.rowIndex;
        const empBreaks = breaks[empName] || [];
        const firstRowForEmp = !printed.has(empName);

        if (firstRowForEmp) {
            // First segment row for this person: write their breaks
            setCell(rZero, 4, empBreaks[0] !== undefined ? minutesToTime(empBreaks[0]) : "");
            setCell(rZero, 5, empBreaks[1] !== undefined ? minutesToTime(empBreaks[1]) : "");
            setCell(rZero, 6, empBreaks[2] !== undefined ? minutesToTime(empBreaks[2]) : "");

            printed.add(empName);
        } else {
            // Subsequent segments for the same person: leave empty
            setCell(rZero, 4, "");
            setCell(rZero, 5, "");
            setCell(rZero, 6, "");
        }
    });

    // Ensure !ref includes columns up through G (after deleting column D)
    const ref = sheet["!ref"] || "A1";
    const refRange = XLSX.utils.decode_range(ref);
    if (refRange.e.c < 6) { // col G is index 6
        refRange.e.c = 6;
        sheet["!ref"] = XLSX.utils.encode_range(refRange);
    }
}

// NOTE: The scheduleBreaks() function has been moved to scheduler.js
// This allows the same core scheduling logic to be used in both the web app and Office Scripts
// Helper functions timeToMinutes, minutesToTime, and formatName are also in scheduler.js

// Helper: convert string to ArrayBuffer
function stringToArrayBuffer(string) {
    const buffer = new ArrayBuffer(string.length);
    const view = new Uint8Array(buffer);
    for (let i = 0; i < string.length; i++) {
        view[i] = string.charCodeAt(i) & 0xFF;
    }
    return buffer;
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

// Whether we should enable break optimization for this main+sub combo
// Break optimization is enabled if the department is part of any group
function isBreakAdjustmentEnabled(mainDept, subDept) {
    if (!mainDept) return false;

    // Check if this department is in any group
    const group = findGroupContaining(mainDept, subDept);
    return group !== undefined;
}
