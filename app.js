// App State & Multi-Hospital Configurations
let activeHospital = 'naeem'; // 'naeem' or 'cavalry'

const HOSPITAL_CONFIGS = {
    naeem: {
        id: 'naeem',
        name: 'Naeem Surgical',
        doctorRate: 0.30,
        hospitalRate: 0.70,
        doctorPercentLabel: '30%',
        hospitalPercentLabel: '70%',
        dailyKey: 'doctor_naeem_daily_patients',
        ledgerKey: 'doctor_naeem_all_patients',
        operationsKey: 'doctor_naeem_operations',
        activeDayKey: 'doctor_naeem_active_day',
        sheetsUrlKey: 'doctor_naeem_google_sheet_url'
    },
    cavalry: {
        id: 'cavalry',
        name: 'Cavalry Hospital',
        doctorRate: 0.70,
        hospitalRate: 0.30,
        doctorPercentLabel: '70%',
        hospitalPercentLabel: '30%',
        dailyKey: 'doctor_cavalry_daily_patients',
        ledgerKey: 'doctor_cavalry_all_patients',
        operationsKey: 'doctor_cavalry_operations',
        activeDayKey: 'doctor_cavalry_active_day',
        sheetsUrlKey: 'doctor_cavalry_google_sheet_url'
    }
};

let dailyPatients = [];
let allPatients = [];
let operationsList = [];
let activeDay = '';
let googleSheetUrl = '';

// Modal confirmation state
let pendingAction = null;

// Apps Script Code Snippet
const APPS_SCRIPT_CODE = `function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    
    // Check for connection test
    if (data.type === 'test') {
      return ContentService.createTextOutput(JSON.stringify({ status: "success", message: "Connected!" }))
        .setMimeType(ContentService.MimeType.JSON)
        .addHeader("Access-Control-Allow-Origin", "*");
    }
    
    // Append headers if sheet is brand new
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(["Daily No.", "Patient Name", "Date", "Gross Charges (Rs.)", "Doctor Share 30% (Rs.)", "Hospital Share 70% (Rs.)", "Unique ID"]);
    }
    
    sheet.appendRow([
      data.dailyIndex,
      data.name,
      data.date,
      data.charges,
      data.split30,
      data.split70,
      data.id
    ]);
    
    return ContentService.createTextOutput(JSON.stringify({ status: "success" }))
      .setMimeType(ContentService.MimeType.JSON)
      .addHeader("Access-Control-Allow-Origin", "*");
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON)
      .addHeader("Access-Control-Allow-Origin", "*");
  }
}`;

// Initialize App
document.addEventListener('DOMContentLoaded', () => {
    try {
        // 1. Initialize Lucide Icons
        if (window.lucide) {
            window.lucide.createIcons();
        }

        // 2. Set default date to today
        const todayStr = getTodayLocalDateString();
        const dateInput = document.getElementById('entry-date');
        if (dateInput) {
            dateInput.value = todayStr;
            dateInput.max = todayStr;
        }

        const opDateInput = document.getElementById('op-date');
        if (opDateInput) {
            opDateInput.value = todayStr;
            opDateInput.max = todayStr;
        }

        // 3. Load data
        loadState();

        // 4. Check for daily reset
        checkDailyReset();

        // 5. Setup event listeners
        setupEventListeners();

        // 6. Start Clock
        startClock();

        // 7. Initial Render
        renderAll();
    } catch (err) {
        console.error("App Initialization Warning:", err);
    }
});

// Helper: Get today's date in YYYY-MM-DD local format
function getTodayLocalDateString() {
    const today = new Date();
    const yyyy = today.getFullYear();
    const mm = String(today.getMonth() + 1).padStart(2, '0');
    const dd = String(today.getDate()).padStart(2, '0');
    return `${yyyy}-${mm}-${dd}`;
}

// Helper: Format Currency (Rs.)
function formatCurrency(value) {
    return new Intl.NumberFormat('en-PK', {
        style: 'currency',
        currency: 'PKR'
    }).format(value);
}

// Helper: Format Date for Display (e.g. Jul 15, 2026)
function formatDateDisplay(dateStr) {
    if (!dateStr) return '';
    const date = new Date(dateStr + 'T00:00:00'); // Prevent timezone offset shift
    return date.toLocaleDateString('en-US', {
        month: 'short',
        day: 'numeric',
        year: 'numeric'
    });
}

// Load state from LocalStorage based on active hospital
function loadState() {
    const savedHospital = localStorage.getItem('doctor_active_hospital');
    if (savedHospital && HOSPITAL_CONFIGS[savedHospital]) {
        activeHospital = savedHospital;
    }

    // Legacy data migration -> Naeem Surgical
    if (localStorage.getItem('doctor_all_patients') && !localStorage.getItem('doctor_naeem_all_patients')) {
        localStorage.setItem('doctor_naeem_daily_patients', localStorage.getItem('doctor_daily_patients') || '[]');
        localStorage.setItem('doctor_naeem_all_patients', localStorage.getItem('doctor_all_patients') || '[]');
        localStorage.setItem('doctor_naeem_active_day', localStorage.getItem('doctor_active_day') || '');
        if (localStorage.getItem('doctor_google_sheet_url')) {
            localStorage.setItem('doctor_naeem_google_sheet_url', localStorage.getItem('doctor_google_sheet_url'));
        }
    }

    const cfg = HOSPITAL_CONFIGS[activeHospital] || HOSPITAL_CONFIGS.naeem;
    const storedDaily = localStorage.getItem(cfg.dailyKey);
    const storedLedger = localStorage.getItem(cfg.ledgerKey);
    const storedOps = localStorage.getItem(cfg.operationsKey);
    const storedActiveDay = localStorage.getItem(cfg.activeDayKey);
    const storedSheetUrl = localStorage.getItem(cfg.sheetsUrlKey);

    dailyPatients = storedDaily ? JSON.parse(storedDaily) : [];
    allPatients = storedLedger ? JSON.parse(storedLedger) : [];
    operationsList = storedOps ? JSON.parse(storedOps) : [];
    activeDay = storedActiveDay || getTodayLocalDateString();
    googleSheetUrl = storedSheetUrl || '';

    if (!storedActiveDay) {
        localStorage.setItem(cfg.activeDayKey, activeDay);
    }
}

// Save state to LocalStorage for active hospital
function saveState() {
    const cfg = HOSPITAL_CONFIGS[activeHospital] || HOSPITAL_CONFIGS.naeem;
    localStorage.setItem('doctor_active_hospital', activeHospital);
    localStorage.setItem(cfg.dailyKey, JSON.stringify(dailyPatients));
    localStorage.setItem(cfg.ledgerKey, JSON.stringify(allPatients));
    localStorage.setItem(cfg.operationsKey, JSON.stringify(operationsList));
    localStorage.setItem(cfg.activeDayKey, activeDay);
    if (googleSheetUrl) {
        localStorage.setItem(cfg.sheetsUrlKey, googleSheetUrl);
    } else {
        localStorage.removeItem(cfg.sheetsUrlKey);
    }
}

// Set active hospital and re-render
function setHospital(hospitalId) {
    if (!HOSPITAL_CONFIGS[hospitalId]) return;
    activeHospital = hospitalId;
    localStorage.setItem('doctor_active_hospital', hospitalId);
    loadState();
    checkDailyReset();
    renderAll();
    showToast(`Switched to ${HOSPITAL_CONFIGS[activeHospital].name}`, 'info');
}

// Update Hospital Switcher Toolbar UI & Labels
function updateHospitalSwitcherUI() {
    const naeemBtn = document.getElementById('hospital-naeem-btn');
    const cavalryBtn = document.getElementById('hospital-cavalry-btn');
    if (naeemBtn && cavalryBtn) {
        naeemBtn.classList.toggle('active', activeHospital === 'naeem');
        cavalryBtn.classList.toggle('active', activeHospital === 'cavalry');
    }
    updateDynamicLabels();
}

// Update Dynamic Split Labels across all sheets
function updateDynamicLabels() {
    const cfg = HOSPITAL_CONFIGS[activeHospital] || HOSPITAL_CONFIGS.naeem;
    
    // Sheet 1 stats
    const todayDocLabel = document.getElementById('today-doctor-label');
    if (todayDocLabel) todayDocLabel.textContent = `Doctor Share (${cfg.doctorPercentLabel})`;

    // Sheet 2 stats & table headers
    const ledgerDocLabel = document.getElementById('ledger-doctor-label');
    if (ledgerDocLabel) ledgerDocLabel.textContent = `Doctor Share (${cfg.doctorPercentLabel})`;

    const ledgerHospLabel = document.getElementById('ledger-hospital-label');
    if (ledgerHospLabel) ledgerHospLabel.textContent = `Hospital Share (${cfg.hospitalPercentLabel})`;

    const ledgerThDoc = document.getElementById('ledger-th-doctor');
    if (ledgerThDoc) ledgerThDoc.textContent = `Doctor Share ${cfg.doctorPercentLabel} (Rs.)`;

    const ledgerThHosp = document.getElementById('ledger-th-hospital');
    if (ledgerThHosp) ledgerThHosp.textContent = `Hospital Share ${cfg.hospitalPercentLabel} (Rs.)`;

    // Sheet 3 analytics labels
    const analyticsDocLabel = document.getElementById('analytics-doctor-label');
    if (analyticsDocLabel) analyticsDocLabel.textContent = `Doctor Earnings (${cfg.doctorPercentLabel})`;

    const analyticsHospLabel = document.getElementById('analytics-hospital-label');
    if (analyticsHospLabel) analyticsHospLabel.textContent = `Hospital Share (${cfg.hospitalPercentLabel})`;
}

// Check if day rolled over and daily log needs clearing
function checkDailyReset() {
    const todayStr = getTodayLocalDateString();
    if (activeDay !== todayStr) {
        // Clear daily sheet for the new day
        if (dailyPatients.length > 0) {
            dailyPatients = [];
            showToast('Daily Log (Sheet 1) has been reset for the new day.', 'info');
        }
        activeDay = todayStr;
        saveState();
    }
}

// Clock updates date & time badge
function startClock() {
    const clockDisplay = document.getElementById('current-datetime-display');
    if (!clockDisplay) return;
    const update = () => {
        const now = new Date();
        const formatted = now.toLocaleDateString('en-US', {
            weekday: 'short',
            month: 'short',
            day: 'numeric',
            year: 'numeric'
        }) + ' | ' + now.toLocaleTimeString('en-US', {
            hour: '2-digit',
            minute: '2-digit',
            second: '2-digit'
        });
        clockDisplay.textContent = formatted;
    };
    update();
    setInterval(update, 1000);
}

// Setup Event Listeners
function setupEventListeners() {
    // Tab switching
    const tab1Btn = document.getElementById('tab1-btn');
    const tab2Btn = document.getElementById('tab2-btn');
    const tab3Btn = document.getElementById('tab3-btn');
    const tab4Btn = document.getElementById('tab4-btn');
    const sheet1 = document.getElementById('sheet1-content');
    const sheet2 = document.getElementById('sheet2-content');
    const sheet3 = document.getElementById('sheet3-content');
    const sheet4 = document.getElementById('sheet4-content');

    const switchTab = (activeBtn, activeSheet, renderFn) => {
        [tab1Btn, tab2Btn, tab3Btn, tab4Btn].forEach(b => {
            if (b) {
                const isActive = (b === activeBtn);
                b.classList.toggle('active', isActive);
                b.setAttribute('aria-selected', isActive ? 'true' : 'false');
            }
        });
        [sheet1, sheet2, sheet3, sheet4].forEach(s => {
            if (s) s.classList.toggle('active', s === activeSheet);
        });
        if (renderFn) renderFn();
    };

    if (tab1Btn) tab1Btn.addEventListener('click', () => switchTab(tab1Btn, sheet1, renderDailyTable));
    if (tab2Btn) tab2Btn.addEventListener('click', () => switchTab(tab2Btn, sheet2, renderLedgerTable));
    if (tab3Btn) tab3Btn.addEventListener('click', () => switchTab(tab3Btn, sheet3, renderAnalytics));
    if (tab4Btn) tab4Btn.addEventListener('click', () => switchTab(tab4Btn, sheet4, renderOperationsPage));

    // Hospital Switcher Buttons
    const naeemBtn = document.getElementById('hospital-naeem-btn');
    const cavalryBtn = document.getElementById('hospital-cavalry-btn');
    if (naeemBtn) naeemBtn.addEventListener('click', () => setHospital('naeem'));
    if (cavalryBtn) cavalryBtn.addEventListener('click', () => setHospital('cavalry'));

    // Patient Name Input: Auto-Count Badge preview
    const nameInput = document.getElementById('patient-name');
    const countPreview = document.getElementById('count-preview');

    if (nameInput && countPreview) {
        nameInput.addEventListener('input', () => {
            const nameVal = nameInput.value.trim();
            if (nameVal.length > 0) {
                const nextCount = dailyPatients.length + 1;
                countPreview.textContent = `#${nextCount}`;
                countPreview.classList.add('visible');
            } else {
                countPreview.classList.remove('visible');
            }
        });
    }

    // Form Submissions
    const form = document.getElementById('patient-form');
    if (form) {
        form.addEventListener('submit', (e) => {
            e.preventDefault();
            handleAddPatient();
        });
    }

    const opForm = document.getElementById('operation-form');
    if (opForm) {
        opForm.addEventListener('submit', (e) => {
            e.preventDefault();
            handleAddOperation();
        });
    }

    // Patient Category Pill Toggle (Private vs Hospital)
    const pillPrivate = document.getElementById('pill-private');
    const pillHospital = document.getElementById('pill-hospital');
    if (pillPrivate && pillHospital) {
        pillPrivate.addEventListener('click', () => {
            pillPrivate.classList.add('active');
            pillHospital.classList.remove('active');
        });
        pillHospital.addEventListener('click', () => {
            pillHospital.classList.add('active');
            pillPrivate.classList.remove('active');
        });
    }

    // Ledger Search / Filters
    const searchInput = document.getElementById('search-input');
    if (searchInput) searchInput.addEventListener('input', renderLedgerTable);

    const startDateInput = document.getElementById('filter-start-date');
    if (startDateInput) startDateInput.addEventListener('change', renderLedgerTable);

    const endDateInput = document.getElementById('filter-end-date');
    if (endDateInput) endDateInput.addEventListener('change', renderLedgerTable);

    const ledgerTypeSelect = document.getElementById('ledger-type-select');
    if (ledgerTypeSelect) {
        ledgerTypeSelect.addEventListener('change', renderLedgerTable);
    }

    // Month Selector on Sheet 2
    const ledgerMonthSelect = document.getElementById('ledger-month-select');
    if (ledgerMonthSelect) {
        ledgerMonthSelect.addEventListener('change', () => {
            const ym = ledgerMonthSelect.value;
            if (ym) {
                const parts = ym.split('-');
                const year = parseInt(parts[0]);
                const monthIndex = parseInt(parts[1]) - 1;
                const daysInMonth = new Date(year, monthIndex + 1, 0).getDate();
                
                if (startDateInput) startDateInput.value = `${ym}-01`;
                if (endDateInput) endDateInput.value = `${ym}-${String(daysInMonth).padStart(2, '0')}`;
            } else {
                if (startDateInput) startDateInput.value = '';
                if (endDateInput) endDateInput.value = '';
            }
            renderLedgerTable();
        });
    }

    // Clear Filters
    const clearFiltersBtn = document.getElementById('clear-filters-btn');
    if (clearFiltersBtn) {
        clearFiltersBtn.addEventListener('click', () => {
            if (searchInput) searchInput.value = '';
            if (startDateInput) startDateInput.value = '';
            if (endDateInput) endDateInput.value = '';
            if (ledgerMonthSelect) ledgerMonthSelect.value = '';
            if (ledgerTypeSelect) ledgerTypeSelect.value = 'all';
            renderLedgerTable();
            showToast('Filters cleared', 'info');
        });
    }

    // Export CSV
    const exportCsvBtn = document.getElementById('export-csv-btn');
    if (exportCsvBtn) exportCsvBtn.addEventListener('click', exportLedgerToCSV);

    // Delete All Ledger Entries Button
    const clearLedgerBtn = document.getElementById('clear-ledger-btn');
    if (clearLedgerBtn) {
        clearLedgerBtn.addEventListener('click', () => {
            if (allPatients.length === 0) {
                showToast('No records available to delete.', 'info');
                return;
            }
            openModal(
                'Delete All Ledger Records?',
                'Are you sure you want to permanently delete all records from the Main Ledger (Sheet 2)? This action cannot be undone.',
                () => {
                    allPatients = [];
                    dailyPatients = [];
                    saveState();
                    renderAll();
                    showToast('All ledger records have been deleted permanently.', 'info');
                }
            );
        });
    }

    // Sync All Button
    const syncAllBtn = document.getElementById('sync-all-btn');
    if (syncAllBtn) syncAllBtn.addEventListener('click', syncAllPending);

    // Settings Modal Triggers
    const settingsModal = document.getElementById('settings-modal');
    const settingsBtn = document.getElementById('settings-btn');
    const closeSettings = document.getElementById('close-settings-btn');
    const cancelSettings = document.getElementById('cancel-settings-btn');
    const saveSettings = document.getElementById('save-settings-btn');
    const disconnectBtn = document.getElementById('disconnect-sheet-btn');
    const copyScriptBtn = document.getElementById('copy-script-btn');

    settingsBtn.addEventListener('click', () => {
        // Load settings values
        document.getElementById('apps-script-code').value = APPS_SCRIPT_CODE;
        document.getElementById('google-sheet-url').value = googleSheetUrl;
        
        if (googleSheetUrl) {
            disconnectBtn.style.display = 'block';
        } else {
            disconnectBtn.style.display = 'none';
        }
        
        settingsModal.classList.add('active');
    });

    const closeSettingsModal = () => {
        settingsModal.classList.remove('active');
    };

    closeSettings.addEventListener('click', closeSettingsModal);
    cancelSettings.addEventListener('click', closeSettingsModal);
    
    // Copy Code snippet
    copyScriptBtn.addEventListener('click', () => {
        const textarea = document.getElementById('apps-script-code');
        textarea.select();
        document.execCommand('copy');
        showToast('Apps Script code copied to clipboard!', 'success');
    });

    // Save and Test URL
    saveSettings.addEventListener('click', async () => {
        const urlInput = document.getElementById('google-sheet-url').value.trim();
        if (!urlInput) {
            googleSheetUrl = '';
            saveState();
            closeSettingsModal();
            renderAll();
            showToast('Google Sheet connection removed.', 'info');
            return;
        }

        // Validate basic URL structure
        if (!urlInput.startsWith('https://script.google.com/')) {
            showToast('Please enter a valid Google Apps Script Web App URL.', 'danger');
            return;
        }

        saveSettings.disabled = true;
        saveSettings.querySelector('span').textContent = 'Testing...';

        try {
            // Attempt to send connection test
            const response = await fetch(urlInput, {
                method: 'POST',
                mode: 'no-cors', // Avoids CORS blocker on redirect
                body: JSON.stringify({ type: 'test' }),
                headers: { 'Content-Type': 'text/plain' }
            });

            googleSheetUrl = urlInput;
            saveState();
            closeSettingsModal();
            renderAll();
            showToast('Google Sheet settings saved and test ping sent!', 'success');
            
            // Sync any existing pending unsynced records
            syncAllPending();
        } catch (err) {
            console.error(err);
            showToast('Connection failed. Please check the URL and try again.', 'danger');
        } finally {
            saveSettings.disabled = false;
            saveSettings.querySelector('span').textContent = 'Test & Save';
        }
    });

    // Disconnect Button
    disconnectBtn.addEventListener('click', () => {
        googleSheetUrl = '';
        
        // Reset all pending statuses back to local
        allPatients.forEach(p => {
            if (p.syncStatus === 'pending') {
                p.syncStatus = 'local';
            }
        });
        
        saveState();
        closeSettingsModal();
        renderAll();
        showToast('Disconnected from Google Sheet.', 'info');
    });

    // Confirmation Modal event hooks
    document.getElementById('modal-cancel-btn').addEventListener('click', closeModal);
    document.getElementById('confirm-modal').addEventListener('click', (e) => {
        if (e.target === document.getElementById('confirm-modal')) {
            closeModal();
        }
    });
    document.getElementById('modal-confirm-btn').addEventListener('click', () => {
        if (pendingAction) {
            pendingAction();
            closeModal();
        }
    });
}

// Add Patient Action
function handleAddPatient() {
    const nameInput = document.getElementById('patient-name');
    const dateInput = document.getElementById('entry-date');
    const chargesInput = document.getElementById('entry-charges');

    const name = nameInput.value.trim();
    const date = dateInput.value;
    const chargesVal = chargesInput.value;

    if (!name || !date || !chargesVal) {
        showToast('Please fill out all fields.', 'danger');
        return;
    }

    const charges = parseFloat(chargesVal);
    if (isNaN(charges) || charges < 0) {
        showToast('Please enter a valid charge amount.', 'danger');
        return;
    }

    const patientId = Date.now().toString();
    const nextDailyIndex = dailyPatients.length + 1;
    const initialSyncStatus = googleSheetUrl ? 'pending' : 'local';
    const cfg = HOSPITAL_CONFIGS[activeHospital] || HOSPITAL_CONFIGS.naeem;

    // Create records
    const dailyRecord = {
        id: patientId,
        dailyIndex: nextDailyIndex,
        name: name,
        date: date,
        charges: charges
    };

    const ledgerRecord = {
        id: patientId,
        dailyIndex: nextDailyIndex,
        name: name,
        date: date,
        charges: charges,
        split30: charges * cfg.doctorRate,      // Doctor Share
        split70: charges * cfg.hospitalRate,    // Hospital Share
        hospitalId: activeHospital,
        syncStatus: initialSyncStatus
    };

    // Add to daily list
    dailyPatients.push(dailyRecord);

    // Add to all-time database
    allPatients.push(ledgerRecord);

    // Save
    saveState();

    // Reset Form Input
    nameInput.value = '';
    chargesInput.value = '';
    dateInput.value = getTodayLocalDateString();
    document.getElementById('count-preview').classList.remove('visible');

    // Trigger google sheets sync in background if URL is active
    if (googleSheetUrl) {
        syncRecordToGoogleSheet(patientId);
    }

    // Re-render
    renderAll();

    showToast(`Patient "${name}" added successfully!`, 'success');
}

// Add Operation Action
function handleAddOperation() {
    const nameInput = document.getElementById('op-patient-name');
    const procInput = document.getElementById('op-procedure-type');
    const dateInput = document.getElementById('op-date');
    const chargesInput = document.getElementById('op-charges');
    const categoryPill = document.querySelector('.patient-category-toggle .category-pill.active');

    const name = nameInput ? nameInput.value.trim() : '';
    const procedure = procInput ? procInput.value.trim() : '';
    const date = dateInput ? dateInput.value : '';
    const chargesVal = chargesInput ? chargesInput.value : '';
    const category = categoryPill && categoryPill.id === 'pill-hospital' ? 'hospital' : 'private';

    if (!name || !procedure || !date || !chargesVal) {
        showToast('Please fill out all operation fields.', 'danger');
        return;
    }

    const charges = parseFloat(chargesVal);
    if (isNaN(charges) || charges < 0) {
        showToast('Please enter a valid amount.', 'danger');
        return;
    }

    const opRecord = {
        id: Date.now().toString(),
        name: name,
        procedure: procedure,
        date: date,
        category: category,
        charges: charges,
        hospitalId: activeHospital
    };

    operationsList.push(opRecord);
    saveState();

    // Reset Form Input
    nameInput.value = '';
    procInput.value = '';
    chargesInput.value = '';
    dateInput.value = getTodayLocalDateString();

    renderAll();
    showToast(`Operation for "${name}" saved!`, 'success');
}

// Delete Operation Action
function handleDeleteOperation(opId) {
    const op = operationsList.find(o => o.id === opId);
    if (!op) return;

    openModal(
        'Delete Operation Record?',
        `Are you sure you want to delete the operation record for "${op.name}" (${op.procedure})?`,
        () => {
            operationsList = operationsList.filter(o => o.id !== opId);
            saveState();
            renderAll();
            showToast('Operation record deleted.', 'info');
        }
    );
}

// Render Sheet 4 Operations Page
function renderOperationsPage() {
    const tbody = document.getElementById('operations-table-body');
    const emptyState = document.getElementById('op-empty-state');
    if (!tbody) return;

    tbody.innerHTML = '';

    const totalCount = operationsList.length;
    const totalCharges = operationsList.reduce((sum, op) => sum + op.charges, 0);
    const privateOps = operationsList.filter(op => op.category === 'private');
    const hospitalOps = operationsList.filter(op => op.category === 'hospital');

    const privateGross = privateOps.reduce((sum, op) => sum + op.charges, 0);
    const hospitalGross = hospitalOps.reduce((sum, op) => sum + op.charges, 0);

    const cEl = document.getElementById('op-total-count');
    const rEl = document.getElementById('op-total-charges');
    const pEl = document.getElementById('op-private-count');
    const hEl = document.getElementById('op-hospital-count');

    if (cEl) cEl.textContent = totalCount;
    if (rEl) rEl.textContent = formatCurrency(totalCharges);
    if (pEl) pEl.textContent = `${privateOps.length} (${formatCurrency(privateGross)})`;
    if (hEl) hEl.textContent = `${hospitalOps.length} (${formatCurrency(hospitalGross)})`;

    if (operationsList.length === 0) {
        if (emptyState) emptyState.style.display = 'block';
        return;
    }

    if (emptyState) emptyState.style.display = 'none';

    // Sort operations by date descending
    const sortedOps = [...operationsList].sort((a, b) => b.date.localeCompare(a.date));

    sortedOps.forEach((op, index) => {
        const tr = document.createElement('tr');
        
        const catBadge = op.category === 'private'
            ? '<span class="type-badge op-private">Private</span>'
            : '<span class="type-badge op-hospital">Hospital</span>';

        tr.innerHTML = `
            <td><span class="patient-badge">#${index + 1}</span></td>
            <td style="font-weight: 500;">${escapeHtml(op.name)}</td>
            <td style="color: var(--text-secondary);">${escapeHtml(op.procedure)}</td>
            <td>${catBadge}</td>
            <td>${formatDateDisplay(op.date)}</td>
            <td class="charge-text" style="text-align: right; color: var(--success); font-weight: 600;">${formatCurrency(op.charges)}</td>
            <td style="text-align: right;">
                <button class="delete-op-row-btn" data-id="${op.id}" style="background: none; border: none; color: var(--danger); cursor: pointer; padding: 0.25rem;">
                    <i data-lucide="trash"></i>
                </button>
            </td>
        `;
        tbody.appendChild(tr);
    });

    tbody.querySelectorAll('.delete-op-row-btn').forEach(btn => {
        btn.addEventListener('click', () => {
            const id = btn.getAttribute('data-id');
            handleDeleteOperation(id);
        });
    });

    if (window.lucide) {
        window.lucide.createIcons();
    }
}

// Push a single record to Google Sheets
function syncRecordToGoogleSheet(patientId) {
    const record = allPatients.find(p => p.id === patientId);
    if (!record || !googleSheetUrl) return;

    fetch(googleSheetUrl, {
        method: 'POST',
        mode: 'no-cors', // Omit CORS checks on redirects
        body: JSON.stringify(record),
        headers: { 'Content-Type': 'text/plain' }
    })
    .then(() => {
        // Mark as synced on successful post response
        updateRecordSyncStatus(patientId, 'synced');
    })
    .catch(err => {
        console.error("Sync error:", err);
        updateRecordSyncStatus(patientId, 'pending');
    });
}

// Update Sync Status for a specific record
function updateRecordSyncStatus(id, newStatus) {
    const record = allPatients.find(p => p.id === id);
    if (record) {
        record.syncStatus = newStatus;
        saveState();
        renderAll();
    }
}

// Sync all pending records
function syncAllPending() {
    if (!googleSheetUrl) {
        showToast('Please configure a Google Sheet connection first.', 'danger');
        return;
    }

    const pending = allPatients.filter(p => p.syncStatus === 'pending');
    if (pending.length === 0) return;

    showToast(`Syncing ${pending.length} pending record(s) to Google Sheets...`, 'info');

    let syncCompletedCount = 0;

    pending.forEach(p => {
        fetch(googleSheetUrl, {
            method: 'POST',
            mode: 'no-cors',
            body: JSON.stringify(p),
            headers: { 'Content-Type': 'text/plain' }
        })
        .then(() => {
            p.syncStatus = 'synced';
            syncCompletedCount++;
            if (syncCompletedCount === pending.length) {
                saveState();
                renderAll();
                showToast('All pending records synced successfully!', 'success');
            }
        })
        .catch(err => {
            console.error("Batch sync error for ID " + p.id + ":", err);
        });
    });
}

// Delete Patient Action
function handleDeletePatient(patientId, isFromLedger = false) {
    const listToSearch = isFromLedger ? allPatients : dailyPatients;
    const recordObj = listToSearch.find(p => p.id === patientId);

    if (!recordObj) return;

    openModal(
        'Delete Patient Record?',
        `Are you sure you want to delete the record for "${recordObj.name}"? This will remove it from both the Daily Log and All-Time Ledger permanently.`,
        () => {
            const dateToReindex = recordObj.date;

            // Remove from daily list
            dailyPatients = dailyPatients.filter(p => p.id !== patientId);
            // Remove from all-time ledger
            allPatients = allPatients.filter(p => p.id !== patientId);

            // Reindex daily lists for that date to keep indices sequential
            reindexDailyLog();
            reindexLedgerForDate(dateToReindex);

            saveState();
            renderAll();
            showToast('Patient record deleted and indices updated.', 'info');
        }
    );
}

// Reindex the current day's log
function reindexDailyLog() {
    dailyPatients.forEach((p, index) => {
        p.dailyIndex = index + 1;
    });
}

// Reindex archive entries for a specific date so they stay correct (1, 2, 3...)
function reindexLedgerForDate(dateStr) {
    const dayEntries = allPatients.filter(p => p.date === dateStr);
    dayEntries.sort((a, b) => a.id.localeCompare(b.id));
    dayEntries.forEach((p, index) => {
        p.dailyIndex = index + 1;
    });
}

// Render everything
function renderAll() {
    updateHospitalSwitcherUI();
    renderDailyTable();
    renderDailyStats();
    renderLedgerTable();
    renderLedgerStats();
    renderOperationsPage();
    updateAutocompleteSource();
    updateSyncIndicator();
    populateLedgerMonthSelect();
    populateAnalyticsMonthSelect();
    const sheet3 = document.getElementById('sheet3-content');
    if (sheet3 && sheet3.classList.contains('active')) {
        renderAnalytics();
    }
}

// Populate Autocomplete past-patients-list dynamically
function updateAutocompleteSource() {
    const datalist = document.getElementById('past-patients-list');
    datalist.innerHTML = '';
    
    // Find unique names from all patients
    const names = allPatients.map(p => p.name);
    const uniqueNames = [...new Set(names)].sort();

    uniqueNames.forEach(name => {
        const option = document.createElement('option');
        option.value = name;
        datalist.appendChild(option);
    });
}

// Update settings-related buttons and counts on Sheet 2
function updateSyncIndicator() {
    const pending = allPatients.filter(p => p.syncStatus === 'pending');
    const syncAllBtn = document.getElementById('sync-all-btn');
    
    if (googleSheetUrl && pending.length > 0) {
        syncAllBtn.style.display = 'inline-flex';
        document.getElementById('pending-sync-count').textContent = pending.length;
    } else {
        syncAllBtn.style.display = 'none';
    }
}

// Render Sheet 1 Stats
function renderDailyStats() {
    const cfg = HOSPITAL_CONFIGS[activeHospital] || HOSPITAL_CONFIGS.naeem;
    const todayCount = dailyPatients.length;
    const todayTotal = dailyPatients.reduce((sum, p) => sum + p.charges, 0);
    const todayDoctorShare = todayTotal * cfg.doctorRate;

    const countEl = document.getElementById('today-count');
    const chargesEl = document.getElementById('today-charges');
    const split30El = document.getElementById('today-split-30');

    if (countEl) countEl.textContent = todayCount;
    if (chargesEl) chargesEl.textContent = formatCurrency(todayTotal);
    if (split30El) split30El.textContent = formatCurrency(todayDoctorShare);
}

// Render Sheet 1 Table
function renderDailyTable() {
    const tbody = document.getElementById('daily-table-body');
    const emptyState = document.getElementById('daily-empty-state');
    tbody.innerHTML = '';

    if (dailyPatients.length === 0) {
        emptyState.style.display = 'block';
        return;
    }

    emptyState.style.display = 'none';

    dailyPatients.forEach(p => {
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td><span class="patient-badge">#${p.dailyIndex}</span></td>
            <td style="font-weight: 500;">${escapeHtml(p.name)}</td>
            <td>${formatDateDisplay(p.date)}</td>
            <td class="charge-text" style="text-align: right;">${formatCurrency(p.charges)}</td>
            <td style="text-align: right;">
                <button class="delete-row-btn" data-id="${p.id}" style="background: none; border: none; color: var(--danger); cursor: pointer; padding: 0.25rem;">
                    <i data-lucide="trash"></i>
                </button>
            </td>
        `;
        tbody.appendChild(tr);
    });

    // Initialize delete icon buttons
    tbody.querySelectorAll('.delete-row-btn').forEach(btn => {
        btn.addEventListener('click', () => {
            const id = btn.getAttribute('data-id');
            handleDeletePatient(id, false);
        });
    });

    if (window.lucide) {
        window.lucide.createIcons();
    }
}

// Render Sheet 2 Stats & Operations Summary Block
function renderLedgerStats() {
    const cfg = HOSPITAL_CONFIGS[activeHospital] || HOSPITAL_CONFIGS.naeem;
    const totalCount = allPatients.length;
    const grossTotal = allPatients.reduce((sum, p) => sum + p.charges, 0);
    const totalDoctor = allPatients.reduce((sum, p) => sum + (p.split30 !== undefined ? p.split30 : p.charges * cfg.doctorRate), 0);
    const totalHospital = allPatients.reduce((sum, p) => sum + (p.split70 !== undefined ? p.split70 : p.charges * cfg.hospitalRate), 0);

    document.getElementById('ledger-total-count').textContent = totalCount;
    document.getElementById('ledger-gross').textContent = formatCurrency(grossTotal);
    document.getElementById('ledger-share-30').textContent = formatCurrency(totalDoctor);
    document.getElementById('ledger-share-70').textContent = formatCurrency(totalHospital);

    // Operations Block on Sheet 2
    const opCount = operationsList.length;
    const opGross = operationsList.reduce((sum, op) => sum + op.charges, 0);
    const privateOps = operationsList.filter(op => op.category === 'private');
    const hospitalOps = operationsList.filter(op => op.category === 'hospital');

    const privateGross = privateOps.reduce((sum, op) => sum + op.charges, 0);
    const hospitalGross = hospitalOps.reduce((sum, op) => sum + op.charges, 0);

    const opCountEl = document.getElementById('ledger-op-count');
    const opGrossEl = document.getElementById('ledger-op-gross');
    const opPrivateEl = document.getElementById('ledger-op-private-summary');
    const opHospEl = document.getElementById('ledger-op-hospital-summary');

    if (opCountEl) opCountEl.textContent = opCount;
    if (opGrossEl) opGrossEl.textContent = formatCurrency(opGross);
    if (opPrivateEl) opPrivateEl.textContent = `${privateOps.length} (${formatCurrency(privateGross)})`;
    if (opHospEl) opHospEl.textContent = `${hospitalOps.length} (${formatCurrency(hospitalGross)})`;
}

// Render Sheet 2 Table with Filters (Includes Operations & Entry Type Filter)
function renderLedgerTable() {
    const tbody = document.getElementById('ledger-table-body');
    const emptyState = document.getElementById('ledger-empty-state');
    tbody.innerHTML = '';

    const searchQuery = document.getElementById('search-input').value.toLowerCase().trim();
    const startDate = document.getElementById('filter-start-date').value;
    const endDate = document.getElementById('filter-end-date').value;
    const ledgerTypeSelect = document.getElementById('ledger-type-select');
    const selectedType = ledgerTypeSelect ? ledgerTypeSelect.value : 'all';

    // Build unified records list
    let unifiedRecords = [];

    if (selectedType === 'all' || selectedType === 'opd') {
        allPatients.forEach(p => {
            unifiedRecords.push({ ...p, recordKind: 'opd' });
        });
    }

    if (selectedType === 'all' || selectedType === 'operation') {
        operationsList.forEach(op => {
            unifiedRecords.push({
                id: op.id,
                dailyIndex: '-',
                name: op.name,
                procedure: op.procedure,
                date: op.date,
                charges: op.charges,
                split30: op.charges,
                split70: 0,
                category: op.category,
                recordKind: 'operation',
                syncStatus: 'local'
            });
        });
    }

    // Filter unified records
    const filteredRecords = unifiedRecords.filter(p => {
        const matchesSearch = p.name.toLowerCase().includes(searchQuery) || (p.procedure && p.procedure.toLowerCase().includes(searchQuery));
        let matchesStartDate = true;
        if (startDate) matchesStartDate = p.date >= startDate;
        let matchesEndDate = true;
        if (endDate) matchesEndDate = p.date <= endDate;
        return matchesSearch && matchesStartDate && matchesEndDate;
    });

    // Sort by date (descending, newer first)
    filteredRecords.sort((a, b) => b.date.localeCompare(a.date));

    // Update Filtered Summary Banner
    const cfg = HOSPITAL_CONFIGS[activeHospital] || HOSPITAL_CONFIGS.naeem;
    const banner = document.getElementById('filtered-summary-banner');
    const bannerLabel = document.getElementById('summary-banner-label');
    const ledgerMonthSelect = document.getElementById('ledger-month-select');
    const hasActiveFilters = searchQuery.length > 0 || startDate !== '' || endDate !== '' || (ledgerMonthSelect && ledgerMonthSelect.value !== '') || selectedType !== 'all';

    if (banner) {
        if (hasActiveFilters) {
            banner.classList.add('active-filter');
            if (bannerLabel) bannerLabel.textContent = `Filtered View Summary (${filteredRecords.length} matching records)`;
        } else {
            banner.classList.remove('active-filter');
            if (bannerLabel) bannerLabel.textContent = `All-Time Ledger Summary`;
        }
    }

    const filteredCount = filteredRecords.length;
    const filteredGross = filteredRecords.reduce((sum, p) => sum + p.charges, 0);
    const filteredDoctor = filteredRecords.reduce((sum, p) => sum + (p.recordKind === 'operation' ? p.charges : (p.split30 !== undefined ? p.split30 : p.charges * cfg.doctorRate)), 0);
    const filteredHospital = filteredRecords.reduce((sum, p) => sum + (p.recordKind === 'operation' ? 0 : (p.split70 !== undefined ? p.split70 : p.charges * cfg.hospitalRate)), 0);

    const fCountEl = document.getElementById('filtered-count');
    const fGrossEl = document.getElementById('filtered-gross');
    const fDoctorEl = document.getElementById('filtered-doctor-share');
    const fHospitalEl = document.getElementById('filtered-hospital-share');
    const fDoctorLabel = document.getElementById('filtered-doctor-label');
    const fHospitalLabel = document.getElementById('filtered-hospital-label');

    if (fCountEl) fCountEl.textContent = filteredCount;
    if (fGrossEl) fGrossEl.textContent = formatCurrency(filteredGross);
    if (fDoctorEl) fDoctorEl.textContent = formatCurrency(filteredDoctor);
    if (fHospitalEl) fHospitalEl.textContent = formatCurrency(filteredHospital);

    if (fDoctorLabel) fDoctorLabel.textContent = `Doctor Share (${cfg.doctorPercentLabel})`;
    if (fHospitalLabel) fHospitalLabel.textContent = `Hospital Share (${cfg.hospitalPercentLabel})`;

    if (filteredRecords.length === 0) {
        emptyState.style.display = 'block';
        return;
    }

    emptyState.style.display = 'none';

    filteredRecords.forEach(p => {
        const tr = document.createElement('tr');
        
        let typeBadgeHtml = '<span class="type-badge opd">OPD</span>';
        if (p.recordKind === 'operation') {
            typeBadgeHtml = p.category === 'private'
                ? '<span class="type-badge op-private">OP (Private)</span>'
                : '<span class="type-badge op-hospital">OP (Hospital)</span>';
        }

        let syncBadgeHtml = '<span class="sync-badge local">Local</span>';
        if (p.syncStatus === 'synced') {
            syncBadgeHtml = '<span class="sync-badge synced"><i data-lucide="cloud-check" style="width:14px;height:14px;"></i> Synced</span>';
        } else if (p.syncStatus === 'pending') {
            syncBadgeHtml = '<span class="sync-badge pending"><i data-lucide="alert-circle" style="width:14px;height:14px;"></i> Pending</span>';
        }

        const nameDisplay = p.recordKind === 'operation'
            ? `${escapeHtml(p.name)} <span style="font-size: 0.8rem; color: var(--text-muted);">(${escapeHtml(p.procedure)})</span>`
            : escapeHtml(p.name);

        const noBadgeHtml = (p.dailyIndex && p.dailyIndex !== '-')
            ? `<span class="patient-badge">#${p.dailyIndex}</span>`
            : '<span style="color: var(--text-muted); font-size: 0.8rem;">-</span>';

        tr.innerHTML = `
            <td>${typeBadgeHtml}</td>
            <td>${noBadgeHtml}</td>
            <td style="font-weight: 500;">${nameDisplay}</td>
            <td>${formatDateDisplay(p.date)}</td>
            <td class="charge-text" style="text-align: right;">${formatCurrency(p.charges)}</td>
            <td class="charge-text" style="text-align: right; color: var(--success); font-weight: 500;">${formatCurrency(p.recordKind === 'operation' ? p.charges : (p.split30 !== undefined ? p.split30 : p.charges * cfg.doctorRate))}</td>
            <td class="charge-text" style="text-align: right; color: var(--text-muted);">${formatCurrency(p.recordKind === 'operation' ? 0 : (p.split70 !== undefined ? p.split70 : p.charges * cfg.hospitalRate))}</td>
            <td style="text-align: center;">${syncBadgeHtml}</td>
            <td style="text-align: right;">
                <button class="${p.recordKind === 'operation' ? 'delete-op-row-btn' : 'delete-ledger-row-btn'}" data-id="${p.id}" style="background: none; border: none; color: var(--danger); cursor: pointer; padding: 0.25rem;">
                    <i data-lucide="trash"></i>
                </button>
            </td>
        `;
        tbody.appendChild(tr);
    });

    // Initialize delete icon buttons
    tbody.querySelectorAll('.delete-ledger-row-btn').forEach(btn => {
        btn.addEventListener('click', () => {
            const id = btn.getAttribute('data-id');
            handleDeletePatient(id, true);
        });
    });

    tbody.querySelectorAll('.delete-op-row-btn').forEach(btn => {
        btn.addEventListener('click', () => {
            const id = btn.getAttribute('data-id');
            handleDeleteOperation(id);
        });
    });

    if (window.lucide) {
        window.lucide.createIcons();
    }
}

    // Initialize individual pending sync badge retries
    tbody.querySelectorAll('.sync-badge.pending').forEach(badge => {
        badge.addEventListener('click', (e) => {
            const rowId = badge.closest('tr').querySelector('.delete-ledger-row-btn').getAttribute('data-id');
            syncRecordToGoogleSheet(rowId);
        });
    });

    if (window.lucide) {
        window.lucide.createIcons();
    }
}

// Export ledger to CSV format
function exportLedgerToCSV() {
    if (allPatients.length === 0) {
        showToast('No records available to export.', 'danger');
        return;
    }

    const cfg = HOSPITAL_CONFIGS[activeHospital] || HOSPITAL_CONFIGS.naeem;

    // CSV Headers
    let csvContent = "data:text/csv;charset=utf-8,";
    csvContent += `Serial No.,Patient Name,Date,Gross Charges (Rs.),Doctor Share ${cfg.doctorPercentLabel} (Rs.),Hospital Share ${cfg.hospitalPercentLabel} (Rs.),Hospital Name,Sync Status\r\n`;

    // Sort all records chronologically before export
    const sorted = [...allPatients].sort((a, b) => {
        if (a.date !== b.date) return a.date.localeCompare(b.date);
        return a.dailyIndex - b.dailyIndex;
    });

    // CSV Rows
    sorted.forEach(p => {
        const row = [
            `#${p.dailyIndex}`,
            `"${p.name.replace(/"/g, '""')}"`, // escape quotes
            p.date,
            p.charges.toFixed(2),
            (p.split30 !== undefined ? p.split30 : p.charges * cfg.doctorRate).toFixed(2),
            (p.split70 !== undefined ? p.split70 : p.charges * cfg.hospitalRate).toFixed(2),
            `"${cfg.name}"`,
            p.syncStatus || 'local'
        ].join(",");
        csvContent += row + "\r\n";
    });

    // Trigger Download
    const encodedUri = encodeURI(csvContent);
    const link = document.createElement("a");
    const todayStr = getTodayLocalDateString();
    link.setAttribute("href", encodedUri);
    link.setAttribute("download", `${cfg.id.toUpperCase()}_Ledger_Export_${todayStr}.csv`);
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);

    showToast('CSV file downloaded successfully!', 'success');
}

// Modal Utilities
function openModal(title, desc, confirmCallback) {
    document.getElementById('modal-title-text').textContent = title;
    document.getElementById('modal-desc-text').textContent = desc;
    document.getElementById('confirm-modal').classList.add('active');
    pendingAction = confirmCallback;
}

// Close Modal helper
function closeModal() {
    document.getElementById('confirm-modal').classList.remove('active');
    pendingAction = null;
}

// Toast System
function showToast(message, type = 'success') {
    const container = document.getElementById('toast-container');
    const toast = document.createElement('div');
    toast.className = `toast ${type}`;

    let iconName = 'check-circle';
    if (type === 'danger') iconName = 'alert-triangle';
    if (type === 'info') iconName = 'info';

    toast.innerHTML = `
        <i data-lucide="${iconName}"></i>
        <div class="toast-message">${escapeHtml(message)}</div>
    `;

    container.appendChild(toast);

    if (window.lucide) {
        window.lucide.createIcons();
    }

    // Auto remove after 4 seconds
    setTimeout(() => {
        toast.style.animation = 'slideIn 0.3s reverse';
        setTimeout(() => {
            toast.remove();
        }, 300);
    }, 4000);
}

// HTML escaping helper to prevent XSS
function escapeHtml(text) {
    const map = {
        '&': '&amp;',
        '<': '&lt;',
        '>': '&gt;',
        '"': '&quot;',
        "'": '&#039;'
    };
    return text.toString().replace(/[&<>"']/g, function(m) { return map[m]; });
}

// Chart Instances
let flowChartInstance = null;
let revenueChartInstance = null;
let historyChartInstance = null;

// Populate Sheet 2 Ledger Quick Month Select Dropdown
function populateLedgerMonthSelect() {
    const select = document.getElementById('ledger-month-select');
    if (!select) return;

    const currentSelected = select.value;
    select.innerHTML = '<option value="">All Months</option>';

    const monthsSet = new Set();
    allPatients.forEach(p => {
        if (p.date && p.date.length >= 7) {
            monthsSet.add(p.date.substring(0, 7));
        }
    });

    const sortedMonths = Array.from(monthsSet).sort().reverse();

    sortedMonths.forEach(ym => {
        const option = document.createElement('option');
        option.value = ym;
        
        const parts = ym.split('-');
        const dateObj = new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, 1);
        option.textContent = dateObj.toLocaleDateString('en-US', { month: 'long', year: 'numeric' });
        
        select.appendChild(option);
    });

    if (currentSelected && sortedMonths.includes(currentSelected)) {
        select.value = currentSelected;
    }
}

// Populate Analytics Month Select Dropdown
function populateAnalyticsMonthSelect() {
    const select = document.getElementById('analytics-month-select');
    if (!select) return;

    const currentSelected = select.value;
    select.innerHTML = '';

    const monthsSet = new Set();
    const todayYYYYMM = getTodayLocalDateString().substring(0, 7);
    monthsSet.add(todayYYYYMM);

    allPatients.forEach(p => {
        if (p.date && p.date.length >= 7) {
            monthsSet.add(p.date.substring(0, 7));
        }
    });

    const sortedMonths = Array.from(monthsSet).sort().reverse();

    sortedMonths.forEach(ym => {
        const option = document.createElement('option');
        option.value = ym;
        
        const parts = ym.split('-');
        const dateObj = new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, 1);
        option.textContent = dateObj.toLocaleDateString('en-US', { month: 'long', year: 'numeric' });
        
        select.appendChild(option);
    });

    if (currentSelected && sortedMonths.includes(currentSelected)) {
        select.value = currentSelected;
    } else {
        select.value = todayYYYYMM;
    }
}

// Render Sheet 3 Analytics & Charts
function renderAnalytics() {
    populateAnalyticsMonthSelect();

    const cfg = HOSPITAL_CONFIGS[activeHospital] || HOSPITAL_CONFIGS.naeem;
    const select = document.getElementById('analytics-month-select');
    if (!select) return;
    const selectedMonth = select.value || getTodayLocalDateString().substring(0, 7);

    // Filter patients for selected month
    const monthPatients = allPatients.filter(p => p.date && p.date.startsWith(selectedMonth));

    const totalCount = monthPatients.length;
    const grossTotal = monthPatients.reduce((sum, p) => sum + p.charges, 0);
    const doctorShare = monthPatients.reduce((sum, p) => sum + (p.split30 !== undefined ? p.split30 : p.charges * cfg.doctorRate), 0);
    const hospitalShare = monthPatients.reduce((sum, p) => sum + (p.split70 !== undefined ? p.split70 : p.charges * cfg.hospitalRate), 0);

    const cEl = document.getElementById('analytics-month-count');
    const gEl = document.getElementById('analytics-month-gross');
    const dEl = document.getElementById('analytics-month-doctor');
    const hEl = document.getElementById('analytics-month-hospital');

    if (cEl) cEl.textContent = totalCount;
    if (gEl) gEl.textContent = formatCurrency(grossTotal);
    if (dEl) dEl.textContent = formatCurrency(doctorShare);
    if (hEl) hEl.textContent = formatCurrency(hospitalShare);

    // Days in selected month
    const parts = selectedMonth.split('-');
    const year = parseInt(parts[0]);
    const monthIndex = parseInt(parts[1]) - 1;
    const daysInMonth = new Date(year, monthIndex + 1, 0).getDate();

    const dayLabels = [];
    const dailyPatientCounts = new Array(daysInMonth).fill(0);
    const dailyDoctorShares = new Array(daysInMonth).fill(0);
    const dailyHospitalShares = new Array(daysInMonth).fill(0);

    for (let i = 1; i <= daysInMonth; i++) {
        dayLabels.push(`Day ${i}`);
    }

    monthPatients.forEach(p => {
        const dayNum = parseInt(p.date.split('-')[2]);
        if (dayNum >= 1 && dayNum <= daysInMonth) {
            dailyPatientCounts[dayNum - 1] += 1;
            dailyDoctorShares[dayNum - 1] += (p.split30 !== undefined ? p.split30 : p.charges * cfg.doctorRate);
            dailyHospitalShares[dayNum - 1] += (p.split70 !== undefined ? p.split70 : p.charges * cfg.hospitalRate);
        }
    });

    // Chart 1: Daily Patient Flow (Line Chart)
    if (flowChartInstance) {
        flowChartInstance.destroy();
    }

    const flowCtx = document.getElementById('patient-flow-chart');
    if (flowCtx && window.Chart) {
        flowChartInstance = new Chart(flowCtx, {
            type: 'line',
            data: {
                labels: dayLabels,
                datasets: [{
                    label: `Daily Patients (${cfg.name})`,
                    data: dailyPatientCounts,
                    borderColor: '#0284c7',
                    backgroundColor: 'rgba(2, 132, 199, 0.12)',
                    fill: true,
                    tension: 0.3,
                    borderWidth: 2,
                    pointRadius: 4,
                    pointBackgroundColor: '#0284c7'
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    legend: { display: false }
                },
                scales: {
                    y: {
                        beginAtZero: true,
                        ticks: { stepSize: 1, precision: 0 }
                    }
                }
            }
        });
    }

    // Chart 2: Daily Revenue & Split Breakdown (Bar Chart)
    if (revenueChartInstance) {
        revenueChartInstance.destroy();
    }

    const revenueCtx = document.getElementById('revenue-breakdown-chart');
    if (revenueCtx && window.Chart) {
        revenueChartInstance = new Chart(revenueCtx, {
            type: 'bar',
            data: {
                labels: dayLabels,
                datasets: [
                    {
                        label: `Doctor Share ${cfg.doctorPercentLabel} (Rs.)`,
                        data: dailyDoctorShares,
                        backgroundColor: '#10b981',
                        borderRadius: 4
                    },
                    {
                        label: `Hospital Share ${cfg.hospitalPercentLabel} (Rs.)`,
                        data: dailyHospitalShares,
                        backgroundColor: '#0284c7',
                        borderRadius: 4
                    }
                ]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    legend: { position: 'top' }
                },
                scales: {
                    x: { stacked: true },
                    y: { stacked: true, beginAtZero: true }
                }
            }
        });
    }

    // Chart 3: Past 6 Months Volume Comparison
    const last6MonthsLabels = [];
    const last6MonthsCounts = [];
    const now = new Date();

    for (let i = 5; i >= 0; i--) {
        const d = new Date(now.getFullYear(), now.getMonth() - i, 1);
        const ymStr = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
        const label = d.toLocaleDateString('en-US', { month: 'short', year: 'numeric' });
        last6MonthsLabels.push(label);

        const count = allPatients.filter(p => p.date && p.date.startsWith(ymStr)).length;
        last6MonthsCounts.push(count);
    }

    if (historyChartInstance) {
        historyChartInstance.destroy();
    }

    const historyCtx = document.getElementById('monthly-history-chart');
    if (historyCtx && window.Chart) {
        historyChartInstance = new Chart(historyCtx, {
            type: 'bar',
            data: {
                labels: last6MonthsLabels,
                datasets: [{
                    label: 'Total Patients',
                    data: last6MonthsCounts,
                    backgroundColor: 'rgba(2, 132, 199, 0.85)',
                    borderColor: '#0284c7',
                    borderWidth: 1,
                    borderRadius: 6
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    legend: { display: false }
                },
                scales: {
                    y: {
                        beginAtZero: true,
                        ticks: { precision: 0 }
                    }
                }
            }
        });
    }
}
