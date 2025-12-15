// ====================================================================================
// DEPARTMENT REGISTRY - All departments organized hierarchically
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
        departments: [{main: "Hardgoods", sub: "Camping"}]
    },
    {
        id: 3,
        name: "Clothing",
        departments: [{main: "Softgoods", sub: "Clothing"}]
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
    }
];

// ====================================================================================
// ADVANCED SETTINGS MANAGEMENT
// ====================================================================================

// Default advanced settings
const DEFAULT_ADVANCED_SETTINGS = {
    maxRestDelay: 30,          // Maximum delay for rest breaks (minutes)
    maxRestEarly: 15,          // Maximum early start for rest breaks (minutes)
    maxMealDelay: 30,          // Maximum delay for meal breaks (minutes)
    maxMealEarly: 0,           // Maximum early start for meal breaks (minutes)
    deptWeightMultiplier: 4,   // Weight for same-department coverage vs group coverage (1-10 scale)
    proximityWeight: 1         // Weight for proximity to ideal break time (1-10 scale)
};

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

// Default operating hours
const DEFAULT_OPERATING_HOURS = {
    monday: { start: "10:00", end: "21:00" },
    tuesday: { start: "10:00", end: "21:00" },
    wednesday: { start: "10:00", end: "21:00" },
    thursday: { start: "10:00", end: "21:00" },
    friday: { start: "09:00", end: "21:00" },
    saturday: { start: "09:00", end: "21:00" },
    sunday: { start: "10:00", end: "21:00" }
};

// Load operating hours from localStorage or use defaults
function loadOperatingHours() {
    const stored = localStorage.getItem('operatingHours');
    if (stored) {
        try {
            return JSON.parse(stored);
        } catch (e) {
            console.error('Failed to parse stored operating hours:', e);
            return DEFAULT_OPERATING_HOURS;
        }
    }
    return DEFAULT_OPERATING_HOURS;
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

    // Get available departments (not in other groups)
    const available = getAvailableDepartments();

    // If editing, add current group's departments to available list
    if (currentEditingGroupId !== null) {
        const groups = getGroups();
        const currentGroup = groups.find(g => g.id === currentEditingGroupId);
        if (currentGroup) {
            currentGroup.departments.forEach(dept => {
                if (!available.some(d => d.main === dept.main && d.sub === dept.sub)) {
                    available.push(dept);
                }
            });
        }
    }

    // Group available departments by main category
    const byCategory = {};
    available.forEach(dept => {
        if (!byCategory[dept.main]) {
            byCategory[dept.main] = [];
        }
        byCategory[dept.main].push(dept.sub);
    });

    // Sort categories alphabetically
    const sortedCategories = Object.keys(byCategory).sort();

    // Build accordion
    sortedCategories.forEach((category, index) => {
        const cardId = `category-${index}`;
        const departments = byCategory[category].sort();

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
                    ${departments.map(sub => {
                        const isSelected = selectedDepartmentsInModal.some(d =>
                            d.main === category && d.sub === sub
                        );
                        return `
                            <div class="form-check">
                                <input type="checkbox" class="form-check-input dept-picker-checkbox"
                                       id="dept-${escapeHtml(category)}-${escapeHtml(sub)}"
                                       data-main="${escapeHtml(category)}"
                                       data-sub="${escapeHtml(sub)}"
                                       ${isSelected ? 'checked' : ''}>
                                <label class="form-check-label small" for="dept-${escapeHtml(category)}-${escapeHtml(sub)}">
                                    ${escapeHtml(sub)}
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
     * If the dept/subdept belongs to a group, returns all employees in ALL departments within that group
     */
    function getCoworkersAtTime(coverage, time, dept, subdept) {
        if (!coverage[time]) return [];

        // Check if this department belongs to a group
        const group = findGroupContaining(dept, subdept);

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
     * Count how many employees from the same dept/subdept are on break at a specific time
     * If the dept/subdept belongs to a group, counts breaks across ALL departments within that group
     */
    function countBreaksAtTime(schedule, breaks, targetTime, breakDuration, dept, subdept) {
        let count = 0;

        // Check if this department belongs to a group
        const group = findGroupContaining(dept, subdept);

        for (let name in breaks) {
            // Find this employee's dept/subdept
            const empRow = schedule.find(row => row.name === name);
            if (!empRow) continue;

            // Check if employee belongs to the same group or department
            let isInScope = false;
            if (group) {
                // Check if employee is in ANY department within this group
                isInScope = group.departments.some(d =>
                    d.main === empRow.dept && d.sub === empRow.job
                );
            } else {
                // Original behavior - only same dept/subdept
                isInScope = (empRow.dept === dept && empRow.job === subdept);
            }

            if (!isInScope) continue;

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

        // Check if this department is in a group
        const group = findGroupContaining(dept, subdept);

        // Skip if break staggering is disabled (not in any group)
        if (!group) {
            // Just schedule at ideal time without conflict resolution
            breaks[emp.name][1] = emp.shiftStart + 240; // +4:00 hours
            continue;
        }

        // Ideal meal time: +4 hours into shift
        const idealMealTime = emp.shiftStart + 240;

        // Load advanced settings
        const advSettings = loadAdvancedSettings();

        // Count how many coworkers already have lunch at the ideal time
        const conflictsAtIdeal = countBreaksAtTime(newSchedule, breaks, idealMealTime, 30, dept, subdept);

        console.log(`[MEAL DEBUG] ${emp.name} (${group.name}): ideal time ${minutesToTime(idealMealTime)}, conflicts: ${conflictsAtIdeal}`);

        // Coverage-based optimization:
        // Try multiple meal times and pick the one that maximizes coverage
        // Priority: 1) Same department coverage, 2) Group coverage

        // Generate possible times based on advanced settings
        const possibleTimes = [idealMealTime];

        // Add delayed times (after ideal) in 15-minute increments
        for (let delay = 15; delay <= advSettings.maxMealDelay; delay += 15) {
            possibleTimes.push(idealMealTime + delay);
        }

        // Add early times (before ideal) in 15-minute increments
        for (let early = 15; early <= advSettings.maxMealEarly; early += 15) {
            possibleTimes.push(idealMealTime - early);
        }

        // Filter to only valid times within shift
        const validTimes = possibleTimes.filter(t => t >= emp.shiftStart && t + 30 <= emp.shiftEnd);

        let bestTime = idealMealTime;
        let bestScore = -Infinity;

        for (const testTime of validTimes) {
            // Create temporary breaks with this test time
            const tempBreaks = JSON.parse(JSON.stringify(breaks));
            tempBreaks[emp.name] = tempBreaks[emp.name] || [];
            tempBreaks[emp.name][1] = testTime;

            // Calculate coverage map with this meal time
            const coverageMap = calculateCoverageMap(newSchedule, shifts, tempBreaks);

            // Calculate minimum coverage during the actual meal break window (30 minutes)
            // Priority 1: Same department coverage
            // Priority 2: Group coverage
            let minDeptCoverage = Infinity;
            let minGroupCoverage = Infinity;

            for (let time = testTime; time < testTime + 30; time += 15) {
                const coworkers = coverageMap[time] || [];

                // Count same-department coverage
                const deptCoverage = coworkers.filter(c => c.dept === dept && c.subdept === subdept).length;
                minDeptCoverage = Math.min(minDeptCoverage, deptCoverage);

                // Count group coverage (all departments in the group)
                const groupCoverage = coworkers.filter(c =>
                    group.departments.some(d => d.main === c.dept && d.sub === c.subdept)
                ).length;
                minGroupCoverage = Math.min(minGroupCoverage, groupCoverage);
            }

            // Score: prioritize same-department coverage, then group coverage
            // Weight department coverage based on advanced settings
            const score = (minDeptCoverage * advSettings.deptWeightMultiplier) + minGroupCoverage;

            // Add proximity bonus for being closer to ideal time
            // Bonus strength based on advanced settings
            const distanceFromIdeal = Math.abs(testTime - idealMealTime);
            const proximityBonus = Math.max(0, advSettings.proximityWeight - (distanceFromIdeal / 20));

            const finalScore = score + proximityBonus;

            console.log(`  [EVAL] ${emp.name}: ${minutesToTime(testTime)} → dept coverage ${minDeptCoverage}, group coverage ${minGroupCoverage}, score ${score}, final ${finalScore.toFixed(2)}`);

            if (finalScore > bestScore) {
                bestScore = finalScore;
                bestTime = testTime;
            }
        }

        breaks[emp.name][1] = bestTime;

        if (bestTime === idealMealTime) {
            console.log(`[MEAL SCHEDULE] ${emp.name} (${group.name}): lunch scheduled at ${minutesToTime(idealMealTime)}`);
        } else {
            console.log(`[MEAL STAGGER] ${emp.name} (${group.name}): lunch optimized from ${minutesToTime(idealMealTime)} to ${minutesToTime(bestTime)} for coverage (score: ${bestScore})`);
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

        // Check if this department is in a group
        const group = findGroupContaining(dept, subdept);

        // Schedule first rest period (+2 hours into shift)
        if (restBreaksNeeded >= 1) {
            const idealFirstBreak = shiftStart + 120; // +2:00 hours

            if (!group) {
                // No conflict resolution (not in any group)
                breaks[name][0] = idealFirstBreak;
            } else {
                // Load advanced settings
                const advSettings = loadAdvancedSettings();

                // Coverage-based optimization: find the time that maximizes minimum coverage
                // Generate possible times based on advanced settings
                const possibleTimes = [idealFirstBreak];

                // Add delayed times (after ideal)
                for (let delay = 15; delay <= advSettings.maxRestDelay; delay += 15) {
                    possibleTimes.push(idealFirstBreak + delay);
                }

                // Add early times (before ideal)
                for (let early = 15; early <= advSettings.maxRestEarly; early += 15) {
                    possibleTimes.push(idealFirstBreak - early);
                }

                let bestTime = idealFirstBreak;
                let bestMinCoverage = -1;

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

                    // Calculate minimum coverage during the break
                    // Priority 1: Same department coverage
                    // Priority 2: Group coverage
                    let minDeptCoverage = Infinity;
                    let minGroupCoverage = Infinity;

                    for (let t = candidateTime; t < candidateTime + 15; t += 15) {
                        const coworkers = tempCoverage[t] || [];

                        // Count same-department coverage
                        const deptCoverage = coworkers.filter(c => c.dept === dept && c.subdept === subdept).length;
                        minDeptCoverage = Math.min(minDeptCoverage, deptCoverage);

                        // Count group coverage (all departments in the group)
                        const groupCoverage = coworkers.filter(c =>
                            group.departments.some(d => d.main === c.dept && d.sub === c.subdept)
                        ).length;
                        minGroupCoverage = Math.min(minGroupCoverage, groupCoverage);
                    }

                    // Score: prioritize same-department coverage, then group coverage
                    // Weight department coverage based on advanced settings
                    const score = (minDeptCoverage * advSettings.deptWeightMultiplier) + minGroupCoverage;

                    // Add proximity bonus for being closer to ideal time
                    // Bonus strength based on advanced settings
                    const distanceFromIdeal = Math.abs(candidateTime - idealFirstBreak);
                    const proximityBonus = Math.max(0, advSettings.proximityWeight - (distanceFromIdeal / 5));

                    const finalScore = score + proximityBonus;

                    console.log(`  [EVAL] ${name}: ${minutesToTime(candidateTime)} → dept coverage ${minDeptCoverage}, group coverage ${minGroupCoverage}, score ${score}, final ${finalScore.toFixed(2)}`);

                    // Choose the time that maximizes the final score
                    if (finalScore > bestMinCoverage) {
                        bestMinCoverage = finalScore;
                        bestTime = candidateTime;
                        bestIndex = i;
                    }
                }

                breaks[name][0] = bestTime;

                if (bestTime !== idealFirstBreak) {
                    const offset = bestTime - idealFirstBreak;
                    console.log(`[REST STAGGER] ${name} (${group.name}): first break adjusted from ${minutesToTime(idealFirstBreak)} to ${minutesToTime(bestTime)} (offset: ${offset > 0 ? '+' : ''}${offset}min, maintains min coverage of ${bestMinCoverage})`);
                }
            }
        }

        // Schedule second rest period (+2 hours after returning from meal)
        if (restBreaksNeeded >= 2 && breaks[name][1] !== undefined) {
            const idealSecondBreak = breaks[name][1] + 30 + 120; // meal end + 2 hours

            if (!group) {
                // No conflict resolution (not in any group)
                breaks[name][2] = idealSecondBreak;
            } else {
                // Load advanced settings
                const advSettings = loadAdvancedSettings();

                // Coverage-based optimization: find the time that maximizes minimum coverage
                // Generate possible times based on advanced settings
                const possibleTimes = [idealSecondBreak];

                // Add delayed times (after ideal)
                for (let delay = 15; delay <= advSettings.maxRestDelay; delay += 15) {
                    possibleTimes.push(idealSecondBreak + delay);
                }

                // Add early times (before ideal)
                for (let early = 15; early <= advSettings.maxRestEarly; early += 15) {
                    possibleTimes.push(idealSecondBreak - early);
                }

                let bestTime = idealSecondBreak;
                let bestMinCoverage = -1;

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

                    // Calculate minimum coverage during the break
                    // Priority 1: Same department coverage
                    // Priority 2: Group coverage
                    let minDeptCoverage = Infinity;
                    let minGroupCoverage = Infinity;

                    for (let t = candidateTime; t < candidateTime + 15; t += 15) {
                        const coworkers = tempCoverage[t] || [];

                        // Count same-department coverage
                        const deptCoverage = coworkers.filter(c => c.dept === dept && c.subdept === subdept).length;
                        minDeptCoverage = Math.min(minDeptCoverage, deptCoverage);

                        // Count group coverage (all departments in the group)
                        const groupCoverage = coworkers.filter(c =>
                            group.departments.some(d => d.main === c.dept && d.sub === c.subdept)
                        ).length;
                        minGroupCoverage = Math.min(minGroupCoverage, groupCoverage);
                    }

                    // Score: prioritize same-department coverage, then group coverage
                    // Weight department coverage based on advanced settings
                    const score = (minDeptCoverage * advSettings.deptWeightMultiplier) + minGroupCoverage;

                    // Add proximity bonus for being closer to ideal time
                    // Bonus strength based on advanced settings
                    const distanceFromIdeal = Math.abs(candidateTime - idealSecondBreak);
                    const proximityBonus = Math.max(0, advSettings.proximityWeight - (distanceFromIdeal / 5));

                    const finalScore = score + proximityBonus;

                    // Choose the time that maximizes the final score
                    if (finalScore > bestMinCoverage) {
                        bestMinCoverage = finalScore;
                        bestTime = candidateTime;
                    }
                }

                breaks[name][2] = bestTime;

                if (bestTime !== idealSecondBreak) {
                    const offset = bestTime - idealSecondBreak;
                    console.log(`[REST STAGGER] ${name} (${group.name}): second break adjusted from ${minutesToTime(idealSecondBreak)} to ${minutesToTime(bestTime)} (offset: ${offset > 0 ? '+' : ''}${offset}min, maintains min coverage of ${bestMinCoverage})`);
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
        const group = findGroupContaining(dept, subdept);
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
                            console.log(`[BREAK SWAP] ${empA} and ${empB} (${group.name}): swapped ${breakName} breaks (${minutesToTime(timeA)} ↔ ${minutesToTime(timeB)}) to preserve schedule order with identical coverage`);
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

// Whether we should enable break optimization for this main+sub combo
// Break optimization is enabled if the department is part of any group
function isBreakAdjustmentEnabled(mainDept, subDept) {
    if (!mainDept) return false;

    // Check if this department is in any group
    const group = findGroupContaining(mainDept, subDept);
    return group !== undefined;
}
