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
        activeDayKey: 'doctor_cavalry_active_day',
        sheetsUrlKey: 'doctor_cavalry_google_sheet_url'
    }
};

let dailyPatients = [];
let allPatients = [];
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
function initApp() {
    try {
        // 1. Load state & start clock
        loadState();
        startClock();

        // 2. Initial Render
        renderAll();

        // 3. Set default date to today
        const todayStr = getTodayLocalDateString();
        const dateInput = document.getElementById('entry-date');
        if (dateInput) {
            dateInput.value = todayStr;
            dateInput.max = todayStr;
        }

        // 4. Check for daily reset
        checkDailyReset();

        // 5. Setup event listeners
        setupEventListeners();

        // 6. Initialize Lucide Icons
        if (window.lucide && typeof window.lucide.createIcons === 'function') {
            try {
                window.lucide.createIcons();
            } catch (e) {
                console.warn("Lucide icon render note:", e);
            }
        }
    } catch (err) {
        console.error("App Initialization Warning:", err);
    }
}

if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', initApp);
} else {
    initApp();
}

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
    const storedActiveDay = localStorage.getItem(cfg.activeDayKey);
    const storedSheetUrl = localStorage.getItem(cfg.sheetsUrlKey);

    try { dailyPatients = storedDaily ? JSON.parse(storedDaily) : []; } catch (e) { dailyPatients = []; }
    try { allPatients = storedLedger ? JSON.parse(storedLedger) : []; } catch (e) { allPatients = []; }
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
    const sheet1 = document.getElementById('sheet1-content');
    const sheet2 = document.getElementById('sheet2-content');
    const sheet3 = document.getElementById('sheet3-content');

    const switchTab = (activeBtn, activeSheet, renderFn) => {
        [tab1Btn, tab2Btn, tab3Btn].forEach(b => {
            if (b) {
                const isActive = (b === activeBtn);
                b.classList.toggle('active', isActive);
                b.setAttribute('aria-selected', isActive ? 'true' : 'false');
            }
        });
        [sheet1, sheet2, sheet3].forEach(s => {
            if (s) s.classList.toggle('active', s === activeSheet);
        });
        if (renderFn) renderFn();
    };

    if (tab1Btn) tab1Btn.addEventListener('click', () => switchTab(tab1Btn, sheet1, renderDailyTable));
    if (tab2Btn) tab2Btn.addEventListener('click', () => switchTab(tab2Btn, sheet2, renderLedgerTable));
    if (tab3Btn) tab3Btn.addEventListener('click', () => switchTab(tab3Btn, sheet3, renderAnalytics));

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

    // Form Submission
    const form = document.getElementById('patient-form');
    if (form) {
        form.addEventListener('submit', (e) => {
            e.preventDefault();
            handleAddPatient();
        });
    }

    // End Day Button
    const endDayBtn = document.getElementById('end-day-btn');
    if (endDayBtn) {
        endDayBtn.addEventListener('click', () => {
            openModal(
                'End Current Day & Reset?',
                'This will clear today\'s patient log (Sheet 1). The All-Time Ledger (Sheet 2) is safe and will not be altered. Would you like to proceed?',
                () => {
                    dailyPatients = [];
                    saveState();
                    renderAll();
                    showToast('Daily Log cleared. Starting fresh.', 'info');
                }
            );
        });
    }

    // Ledger Search / Filters
    const searchInput = document.getElementById('search-input');
    if (searchInput) searchInput.addEventListener('input', renderLedgerTable);

    const startDateInput = document.getElementById('filter-start-date');
    if (startDateInput) startDateInput.addEventListener('change', renderLedgerTable);

    const endDateInput = document.getElementById('filter-end-date');
    if (endDateInput) endDateInput.addEventListener('change', renderLedgerTable);

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
            renderLedgerTable();
            showToast('Filters cleared', 'info');
        });
    }

    // Export PDF Statement
    const exportPdfBtn = document.getElementById('export-pdf-btn');
    if (exportPdfBtn) exportPdfBtn.addEventListener('click', exportMonthlyPDF);

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

    if (settingsBtn && settingsModal) {
        settingsBtn.addEventListener('click', () => {
            const codeBox = document.getElementById('apps-script-code');
            const urlBox = document.getElementById('google-sheet-url');
            if (codeBox) codeBox.value = APPS_SCRIPT_CODE;
            if (urlBox) urlBox.value = googleSheetUrl;
            
            if (disconnectBtn) {
                disconnectBtn.style.display = googleSheetUrl ? 'block' : 'none';
            }
            
            settingsModal.classList.add('active');
        });
    }

    const closeSettingsModal = () => {
        if (settingsModal) settingsModal.classList.remove('active');
    };

    if (closeSettings) closeSettings.addEventListener('click', closeSettingsModal);
    if (cancelSettings) cancelSettings.addEventListener('click', closeSettingsModal);
    
    // Copy Code snippet
    if (copyScriptBtn) {
        copyScriptBtn.addEventListener('click', () => {
            const textarea = document.getElementById('apps-script-code');
            if (textarea) {
                textarea.select();
                document.execCommand('copy');
                showToast('Apps Script code copied to clipboard!', 'success');
            }
        });
    }

    // Save and Test URL
    if (saveSettings) {
        saveSettings.addEventListener('click', async () => {
            const urlEl = document.getElementById('google-sheet-url');
            const urlInput = urlEl ? urlEl.value.trim() : '';
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
            const btnSpan = saveSettings.querySelector('span');
            if (btnSpan) btnSpan.textContent = 'Testing...';

            try {
                // Attempt to send connection test
                await fetch(urlInput, {
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
                if (btnSpan) btnSpan.textContent = 'Test & Save';
            }
        });
    }

    // Disconnect Button
    if (disconnectBtn) {
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
    }

    // Confirmation Modal event hooks
    const cancelModalBtn = document.getElementById('modal-cancel-btn');
    const confirmModalEl = document.getElementById('confirm-modal');
    const confirmModalBtn = document.getElementById('modal-confirm-btn');

    if (cancelModalBtn) cancelModalBtn.addEventListener('click', closeModal);
    if (confirmModalEl) {
        confirmModalEl.addEventListener('click', (e) => {
            if (e.target === confirmModalEl) {
                closeModal();
            }
        });
    }
    if (confirmModalBtn) {
        confirmModalBtn.addEventListener('click', () => {
            if (pendingAction) {
                pendingAction();
                closeModal();
            }
        });
    }
}

// Add Patient Action
function handleAddPatient() {
    const nameInput = document.getElementById('patient-name');
    const dateInput = document.getElementById('entry-date');
    const chargesInput = document.getElementById('entry-charges');

    const name = nameInput ? nameInput.value.trim() : '';
    const date = dateInput ? dateInput.value : '';
    const chargesVal = chargesInput ? chargesInput.value : '';

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
        split30: charges * cfg.doctorRate,
        split70: charges * cfg.hospitalRate,
        syncStatus: initialSyncStatus
    };

    // Push to lists
    dailyPatients.push(dailyRecord);
    allPatients.push(ledgerRecord);

    // Save State
    saveState();

    // Reset Form Input
    if (nameInput) nameInput.value = '';
    if (chargesInput) chargesInput.value = '';
    if (dateInput) dateInput.value = getTodayLocalDateString();
    const countPreview = document.getElementById('count-preview');
    if (countPreview) countPreview.classList.remove('visible');

    // Trigger google sheets sync in background if URL is active
    if (googleSheetUrl) {
        syncRecordToGoogleSheet(patientId);
    }

    // Re-render
    renderAll();

    showToast(`Patient "${name}" added successfully!`, 'success');
}

// Push a single record to Google Sheets
function syncRecordToGoogleSheet(patientId) {
    const record = allPatients.find(p => p.id === patientId);
    if (!record || !googleSheetUrl) return;

    fetch(googleSheetUrl, {
        method: 'POST',
        mode: 'no-cors',
        body: JSON.stringify(record),
        headers: { 'Content-Type': 'text/plain' }
    })
    .then(() => {
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
    const pending = allPatients.filter(p => p.syncStatus === 'pending');
    if (pending.length === 0 || !googleSheetUrl) return;

    showToast(`Syncing ${pending.length} record(s) to Google Sheet...`, 'info');

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

            dailyPatients = dailyPatients.filter(p => p.id !== patientId);
            allPatients = allPatients.filter(p => p.id !== patientId);

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
    if (!datalist) return;
    datalist.innerHTML = '';
    
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
    const pendingCountEl = document.getElementById('pending-sync-count');
    
    if (syncAllBtn) {
        if (googleSheetUrl && pending.length > 0) {
            syncAllBtn.style.display = 'inline-flex';
            if (pendingCountEl) pendingCountEl.textContent = pending.length;
        } else {
            syncAllBtn.style.display = 'none';
        }
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
    if (!tbody) return;
    tbody.innerHTML = '';

    if (dailyPatients.length === 0) {
        if (emptyState) emptyState.style.display = 'block';
        return;
    }

    if (emptyState) emptyState.style.display = 'none';

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

    tbody.querySelectorAll('.delete-row-btn').forEach(btn => {
        btn.addEventListener('click', () => {
            const id = btn.getAttribute('data-id');
            handleDeletePatient(id, false);
        });
    });

    if (window.lucide && typeof window.lucide.createIcons === 'function') {
        try { window.lucide.createIcons(); } catch (e) {}
    }
}

// Render Sheet 2 Stats
function renderLedgerStats() {
    const cfg = HOSPITAL_CONFIGS[activeHospital] || HOSPITAL_CONFIGS.naeem;
    const totalCount = allPatients.length;
    const grossTotal = allPatients.reduce((sum, p) => sum + p.charges, 0);
    const totalDoctor = allPatients.reduce((sum, p) => sum + (p.split30 !== undefined ? p.split30 : p.charges * cfg.doctorRate), 0);
    const totalHospital = allPatients.reduce((sum, p) => sum + (p.split70 !== undefined ? p.split70 : p.charges * cfg.hospitalRate), 0);

    const countEl = document.getElementById('ledger-total-count');
    const grossEl = document.getElementById('ledger-gross');
    const s30El = document.getElementById('ledger-share-30');
    const s70El = document.getElementById('ledger-share-70');

    if (countEl) countEl.textContent = totalCount;
    if (grossEl) grossEl.textContent = formatCurrency(grossTotal);
    if (s30El) s30El.textContent = formatCurrency(totalDoctor);
    if (s70El) s70El.textContent = formatCurrency(totalHospital);
}

// Render Sheet 2 Table with Filters
function renderLedgerTable() {
    const tbody = document.getElementById('ledger-table-body');
    const emptyState = document.getElementById('ledger-empty-state');
    if (!tbody) return;
    tbody.innerHTML = '';

    const searchInput = document.getElementById('search-input');
    const startDateInput = document.getElementById('filter-start-date');
    const endDateInput = document.getElementById('filter-end-date');

    const searchQuery = searchInput ? searchInput.value.toLowerCase().trim() : '';
    const startDate = startDateInput ? startDateInput.value : '';
    const endDate = endDateInput ? endDateInput.value : '';

    const filteredPatients = allPatients.filter(p => {
        const matchesSearch = p.name.toLowerCase().includes(searchQuery);
        let matchesStartDate = true;
        if (startDate) matchesStartDate = p.date >= startDate;
        let matchesEndDate = true;
        if (endDate) matchesEndDate = p.date <= endDate;
        return matchesSearch && matchesStartDate && matchesEndDate;
    });

    filteredPatients.sort((a, b) => {
        if (a.date !== b.date) {
            return b.date.localeCompare(a.date);
        }
        return b.dailyIndex - a.dailyIndex;
    });

    // Filtered Summary Banner
    const cfg = HOSPITAL_CONFIGS[activeHospital] || HOSPITAL_CONFIGS.naeem;
    const banner = document.getElementById('filtered-summary-banner');
    const bannerLabel = document.getElementById('summary-banner-label');
    const ledgerMonthSelect = document.getElementById('ledger-month-select');
    const hasActiveFilters = searchQuery.length > 0 || startDate !== '' || endDate !== '' || (ledgerMonthSelect && ledgerMonthSelect.value !== '');

    if (banner) {
        if (hasActiveFilters) {
            banner.classList.add('active-filter');
            if (bannerLabel) bannerLabel.textContent = `Filtered View Summary (${filteredPatients.length} matching records)`;
        } else {
            banner.classList.remove('active-filter');
            if (bannerLabel) bannerLabel.textContent = `All-Time Ledger Summary`;
        }
    }

    const filteredCount = filteredPatients.length;
    const filteredGross = filteredPatients.reduce((sum, p) => sum + p.charges, 0);
    const filteredDoctor = filteredPatients.reduce((sum, p) => sum + (p.split30 !== undefined ? p.split30 : p.charges * cfg.doctorRate), 0);
    const filteredHospital = filteredPatients.reduce((sum, p) => sum + (p.split70 !== undefined ? p.split70 : p.charges * cfg.hospitalRate), 0);

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

    if (filteredPatients.length === 0) {
        if (emptyState) emptyState.style.display = 'block';
        return;
    }

    if (emptyState) emptyState.style.display = 'none';

    filteredPatients.forEach(p => {
        const tr = document.createElement('tr');
        
        let syncBadgeHtml = '<span class="sync-badge local">Local</span>';
        if (p.syncStatus === 'synced') {
            syncBadgeHtml = '<span class="sync-badge synced"><i data-lucide="cloud-check" style="width:14px;height:14px;"></i> Synced</span>';
        } else if (p.syncStatus === 'pending') {
            syncBadgeHtml = '<span class="sync-badge pending" title="Click to retry syncing"><i data-lucide="alert-circle" style="width:14px;height:14px;"></i> Pending</span>';
        }

        const doctorVal = (p.split30 !== undefined) ? p.split30 : (p.charges * cfg.doctorRate);
        const hospitalVal = (p.split70 !== undefined) ? p.split70 : (p.charges * cfg.hospitalRate);

        tr.innerHTML = `
            <td><span class="patient-badge">#${p.dailyIndex}</span></td>
            <td style="font-weight: 500;">${escapeHtml(p.name)}</td>
            <td>${formatDateDisplay(p.date)}</td>
            <td class="charge-text" style="text-align: right;">${formatCurrency(p.charges)}</td>
            <td class="charge-text" style="text-align: right; color: var(--success); font-weight: 500;">${formatCurrency(doctorVal)}</td>
            <td class="charge-text" style="text-align: right; color: var(--text-muted);">${formatCurrency(hospitalVal)}</td>
            <td style="text-align: center;">${syncBadgeHtml}</td>
            <td style="text-align: right;">
                <button class="delete-ledger-row-btn" data-id="${p.id}" style="background: none; border: none; color: var(--danger); cursor: pointer; padding: 0.25rem;">
                    <i data-lucide="trash"></i>
                </button>
            </td>
        `;
        tbody.appendChild(tr);
    });

    tbody.querySelectorAll('.delete-ledger-row-btn').forEach(btn => {
        btn.addEventListener('click', () => {
            const id = btn.getAttribute('data-id');
            handleDeletePatient(id, true);
        });
    });

    tbody.querySelectorAll('.sync-badge.pending').forEach(badge => {
        badge.addEventListener('click', () => {
            const row = badge.closest('tr');
            if (row) {
                const btn = row.querySelector('.delete-ledger-row-btn');
                if (btn) {
                    const id = btn.getAttribute('data-id');
                    syncRecordToGoogleSheet(id);
                }
            }
        });
    });

    if (window.lucide && typeof window.lucide.createIcons === 'function') {
        try { window.lucide.createIcons(); } catch (e) {}
    }
}

// Export Ledger to CSV File
function exportLedgerToCSV() {
    const cfg = HOSPITAL_CONFIGS[activeHospital] || HOSPITAL_CONFIGS.naeem;
    if (allPatients.length === 0) {
        showToast('No records in ledger to export.', 'info');
        return;
    }

    let csvContent = "data:text/csv;charset=utf-8,";
    csvContent += `Daily No.,Patient Name,Date,Gross Charges (PKR),Doctor Share (${cfg.doctorPercentLabel}) (PKR),Hospital Share (${cfg.hospitalPercentLabel}) (PKR),Sync Status\n`;

    allPatients.forEach(p => {
        const dVal = (p.split30 !== undefined ? p.split30 : p.charges * cfg.doctorRate);
        const hVal = (p.split70 !== undefined ? p.split70 : p.charges * cfg.hospitalRate);
        const row = [
            p.dailyIndex,
            `"${p.name.replace(/"/g, '""')}"`,
            p.date,
            p.charges.toFixed(2),
            dVal.toFixed(2),
            hVal.toFixed(2),
            p.syncStatus || 'local'
        ].join(",");
        csvContent += row + "\n";
    });

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

// Export 1-Page Monthly Settlement Statement to PDF (Privacy Compliant)
function exportMonthlyPDF() {
    const cfg = HOSPITAL_CONFIGS[activeHospital] || HOSPITAL_CONFIGS.naeem;
    const ledgerMonthSelect = document.getElementById('ledger-month-select');
    const startDateInput = document.getElementById('filter-start-date');
    const endDateInput = document.getElementById('filter-end-date');

    let targetMonth = ledgerMonthSelect ? ledgerMonthSelect.value : '';
    let monthTitle = '';

    if (targetMonth) {
        const parts = targetMonth.split('-');
        const dObj = new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, 1);
        monthTitle = dObj.toLocaleDateString('en-US', { month: 'long', year: 'numeric' });
    } else {
        const today = new Date();
        monthTitle = today.toLocaleDateString('en-US', { month: 'long', year: 'numeric' });
    }

    const startDate = startDateInput ? startDateInput.value : '';
    const endDate = endDateInput ? endDateInput.value : '';
    const searchInput = document.getElementById('search-input');
    const searchQuery = searchInput ? searchInput.value.toLowerCase().trim() : '';

    const recordsToInclude = allPatients.filter(p => {
        const matchesSearch = p.name.toLowerCase().includes(searchQuery);
        let matchesMonth = true;
        if (targetMonth) matchesMonth = p.date && p.date.startsWith(targetMonth);
        let matchesStartDate = true;
        if (startDate) matchesStartDate = p.date >= startDate;
        let matchesEndDate = true;
        if (endDate) matchesEndDate = p.date <= endDate;
        return matchesSearch && matchesMonth && matchesStartDate && matchesEndDate;
    });

    if (recordsToInclude.length === 0) {
        showToast('No records match the selected month/filter to export PDF.', 'info');
        return;
    }

    const totalCount = recordsToInclude.length;
    const grossTotal = recordsToInclude.reduce((sum, p) => sum + p.charges, 0);
    const doctorShare = recordsToInclude.reduce((sum, p) => sum + (p.split30 !== undefined ? p.split30 : p.charges * cfg.doctorRate), 0);
    const hospitalShare = recordsToInclude.reduce((sum, p) => sum + (p.split70 !== undefined ? p.split70 : p.charges * cfg.hospitalRate), 0);

    // Populate PDF Elements
    const nameEl = document.getElementById('pdf-hospital-name');
    const monthEl = document.getElementById('pdf-month-name');
    const dateEl = document.getElementById('pdf-issued-date');
    const countEl = document.getElementById('pdf-patient-count');
    const grossEl = document.getElementById('pdf-gross-revenue');
    const docShareEl = document.getElementById('pdf-doctor-share');
    const hospShareEl = document.getElementById('pdf-hospital-share');
    const docLabelEl = document.getElementById('pdf-doctor-label');
    const hospLabelEl = document.getElementById('pdf-hospital-label');

    const rowGrossEl = document.getElementById('pdf-row-gross');
    const rowDocRateEl = document.getElementById('pdf-row-doc-rate');
    const rowDocShareEl = document.getElementById('pdf-row-doc-share');
    const rowHospRateEl = document.getElementById('pdf-row-hosp-rate');
    const rowHospShareEl = document.getElementById('pdf-row-hosp-share');

    if (nameEl) nameEl.textContent = cfg.name;
    if (monthEl) monthEl.textContent = monthTitle;
    if (dateEl) dateEl.textContent = formatDateDisplay(getTodayLocalDateString());
    if (countEl) countEl.textContent = totalCount;
    if (grossEl) grossEl.textContent = formatCurrency(grossTotal);
    if (docShareEl) docShareEl.textContent = formatCurrency(doctorShare);
    if (hospShareEl) hospShareEl.textContent = formatCurrency(hospitalShare);

    if (docLabelEl) docLabelEl.textContent = `Doctor Share (${cfg.doctorPercentLabel})`;
    if (hospLabelEl) hospLabelEl.textContent = `Hospital Share (${cfg.hospitalPercentLabel})`;

    if (rowGrossEl) rowGrossEl.textContent = formatCurrency(grossTotal);
    if (rowDocRateEl) rowDocRateEl.textContent = cfg.doctorPercentLabel;
    if (rowDocShareEl) rowDocShareEl.textContent = formatCurrency(doctorShare);
    if (rowHospRateEl) rowHospRateEl.textContent = cfg.hospitalPercentLabel;
    if (rowHospShareEl) rowHospShareEl.textContent = formatCurrency(hospitalShare);

    // Trigger Print Dialog
    window.print();
}

// Modal Utilities
function openModal(title, desc, confirmCallback) {
    const titleEl = document.getElementById('modal-title-text');
    const descEl = document.getElementById('modal-desc-text');
    const modalEl = document.getElementById('confirm-modal');
    if (titleEl) titleEl.textContent = title;
    if (descEl) descEl.textContent = desc;
    if (modalEl) modalEl.classList.add('active');
    pendingAction = confirmCallback;
}

function closeModal() {
    const modalEl = document.getElementById('confirm-modal');
    if (modalEl) modalEl.classList.remove('active');
    pendingAction = null;
}

// Toast System
function showToast(message, type = 'success') {
    const container = document.getElementById('toast-container');
    if (!container) return;

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

    if (window.lucide && typeof window.lucide.createIcons === 'function') {
        try { window.lucide.createIcons(); } catch (e) {}
    }

    setTimeout(() => {
        toast.style.animation = 'slideIn 0.3s reverse';
        setTimeout(() => {
            toast.remove();
        }, 300);
    }, 4000);
}

// HTML escaping helper to prevent XSS
function escapeHtml(text) {
    if (text === undefined || text === null) return '';
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

    if (flowChartInstance) {
        flowChartInstance.destroy();
    }

    const flowCtx = document.getElementById('flow-chart');
    if (flowCtx && window.Chart) {
        flowChartInstance = new Chart(flowCtx, {
            type: 'line',
            data: {
                labels: dayLabels,
                datasets: [{
                    label: 'Daily Patient Volume',
                    data: dailyPatientCounts,
                    borderColor: '#0284c7',
                    backgroundColor: 'rgba(2, 132, 199, 0.15)',
                    fill: true,
                    tension: 0.3,
                    pointRadius: 4,
                    pointBackgroundColor: '#0284c7'
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: { legend: { display: false } },
                scales: {
                    y: { beginAtZero: true, ticks: { precision: 0 } }
                }
            }
        });
    }

    if (revenueChartInstance) {
        revenueChartInstance.destroy();
    }

    const revenueCtx = document.getElementById('revenue-chart');
    if (revenueCtx && window.Chart) {
        revenueChartInstance = new Chart(revenueCtx, {
            type: 'bar',
            data: {
                labels: dayLabels,
                datasets: [
                    {
                        label: `Doctor Share (${cfg.doctorPercentLabel})`,
                        data: dailyDoctorShares,
                        backgroundColor: '#10b981'
                    },
                    {
                        label: `Hospital Share (${cfg.hospitalPercentLabel})`,
                        data: dailyHospitalShares,
                        backgroundColor: '#0284c7'
                    }
                ]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: { legend: { position: 'top' } },
                scales: {
                    x: { stacked: true },
                    y: { stacked: true, beginAtZero: true }
                }
            }
        });
    }

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
                plugins: { legend: { display: false } },
                scales: {
                    y: { beginAtZero: true, ticks: { precision: 0 } }
                }
            }
        });
    }
}
