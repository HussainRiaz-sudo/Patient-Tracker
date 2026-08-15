// Theme State
let currentTheme = localStorage.getItem('doctor_theme') || 'dark';

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
let cavalryProcedures = [];
let naeemAdmitProcedures = [];
let activeDay = '';
let googleSheetUrl = '';

// Modal confirmation state
let pendingAction = null;
let pendingProcToCompleteId = null;
let pendingNaeemAdmitId = null;

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

// Initialize App Theme
function applyTheme(theme) {
    currentTheme = theme;
    document.documentElement.setAttribute('data-theme', theme);
    localStorage.setItem('doctor_theme', theme);

    const btn = document.getElementById('theme-toggle-btn');
    if (btn) {
        btn.setAttribute('title', theme === 'dark' ? 'Switch to Light Theme' : 'Switch to Dark Theme');
        btn.innerHTML = theme === 'dark' ? '<i data-lucide="sun"></i>' : '<i data-lucide="moon"></i>';
        if (window.lucide && typeof window.lucide.createIcons === 'function') {
            try { window.lucide.createIcons(); } catch (e) {}
        }
    }

    const sheet3 = document.getElementById('sheet3-content');
    if (sheet3 && sheet3.classList.contains('active')) {
        renderAnalytics();
    }
}

function toggleTheme() {
    const nextTheme = currentTheme === 'dark' ? 'light' : 'dark';
    applyTheme(nextTheme);
    showToast(`Switched to ${nextTheme === 'dark' ? 'Dark' : 'Light'} Mode`, 'info');
}

// Immediately apply theme on script load before DOM render to prevent flash
applyTheme(currentTheme);

// Initialize App
function initApp() {
    try {
        // 0. Ensure theme icon is set
        applyTheme(currentTheme);

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

        const procDateInput = document.getElementById('proc-due-date');
        if (procDateInput) {
            procDateInput.value = todayStr;
        }

        const naeemAdmitDateInput = document.getElementById('naeem-admit-date');
        if (naeemAdmitDateInput) {
            naeemAdmitDateInput.value = todayStr;
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
    const storedProcedures = localStorage.getItem('doctor_cavalry_procedures');
    const storedNaeemAdmit = localStorage.getItem('doctor_naeem_admit_procedures');
    const storedActiveDay = localStorage.getItem(cfg.activeDayKey);
    const storedSheetUrl = localStorage.getItem(cfg.sheetsUrlKey);

    try { dailyPatients = storedDaily ? JSON.parse(storedDaily) : []; } catch (e) { dailyPatients = []; }
    try { allPatients = storedLedger ? JSON.parse(storedLedger) : []; } catch (e) { allPatients = []; }
    try { cavalryProcedures = storedProcedures ? JSON.parse(storedProcedures) : []; } catch (e) { cavalryProcedures = []; }
    try { naeemAdmitProcedures = storedNaeemAdmit ? JSON.parse(storedNaeemAdmit) : []; } catch (e) { naeemAdmitProcedures = []; }
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
    localStorage.setItem('doctor_cavalry_procedures', JSON.stringify(cavalryProcedures));
    localStorage.setItem('doctor_naeem_admit_procedures', JSON.stringify(naeemAdmitProcedures));
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
    const tab4Btn = document.getElementById('tab4-btn');
    const tab5Btn = document.getElementById('tab5-btn');
    const sheet4 = document.getElementById('sheet4-content');
    const sheet5 = document.getElementById('sheet5-content');

    if (naeemBtn && cavalryBtn) {
        naeemBtn.classList.toggle('active', activeHospital === 'naeem');
        cavalryBtn.classList.toggle('active', activeHospital === 'cavalry');
    }

    if (tab4Btn) {
        if (activeHospital === 'cavalry') {
            tab4Btn.style.display = 'inline-flex';
        } else {
            tab4Btn.style.display = 'none';
            if (tab4Btn.classList.contains('active')) {
                const tab1Btn = document.getElementById('tab1-btn');
                const sheet1 = document.getElementById('sheet1-content');
                if (tab1Btn && sheet1) {
                    [tab1Btn, document.getElementById('tab2-btn'), document.getElementById('tab3-btn'), tab4Btn, tab5Btn].forEach(b => {
                        if (b) b.classList.toggle('active', b === tab1Btn);
                    });
                    [sheet1, document.getElementById('sheet2-content'), document.getElementById('sheet3-content'), sheet4, sheet5].forEach(s => {
                        if (s) s.classList.toggle('active', s === sheet1);
                    });
                }
            }
        }
    }

    if (tab5Btn) {
        if (activeHospital === 'naeem') {
            tab5Btn.style.display = 'inline-flex';
        } else {
            tab5Btn.style.display = 'none';
            if (tab5Btn.classList.contains('active')) {
                const tab1Btn = document.getElementById('tab1-btn');
                const sheet1 = document.getElementById('sheet1-content');
                if (tab1Btn && sheet1) {
                    [tab1Btn, document.getElementById('tab2-btn'), document.getElementById('tab3-btn'), tab4Btn, tab5Btn].forEach(b => {
                        if (b) b.classList.toggle('active', b === tab1Btn);
                    });
                    [sheet1, document.getElementById('sheet2-content'), document.getElementById('sheet3-content'), sheet4, sheet5].forEach(s => {
                        if (s) s.classList.toggle('active', s === sheet1);
                    });
                }
            }
        }
    }

    updateDynamicLabels();
}

// Update Dynamic Split Labels across all sheets (ratio numbers hidden for privacy)
function updateDynamicLabels() {
    // Sheet 1 stats
    const todayDocLabel = document.getElementById('today-doctor-label');
    if (todayDocLabel) todayDocLabel.textContent = `Doctor Share`;

    // Sheet 2 stats & table headers
    const ledgerDocLabel = document.getElementById('ledger-doctor-label');
    if (ledgerDocLabel) ledgerDocLabel.textContent = `Doctor Share`;

    const ledgerHospLabel = document.getElementById('ledger-hospital-label');
    if (ledgerHospLabel) ledgerHospLabel.textContent = `Hospital Share`;

    const ledgerThDoc = document.getElementById('ledger-th-doctor');
    if (ledgerThDoc) ledgerThDoc.textContent = `Doctor Share (Rs.)`;

    const ledgerThHosp = document.getElementById('ledger-th-hospital');
    if (ledgerThHosp) ledgerThHosp.textContent = `Hospital Share (Rs.)`;

    // Sheet 3 analytics labels
    const analyticsDocLabel = document.getElementById('analytics-doctor-label');
    if (analyticsDocLabel) analyticsDocLabel.textContent = `Doctor Earnings`;

    const analyticsHospLabel = document.getElementById('analytics-hospital-label');
    if (analyticsHospLabel) analyticsHospLabel.textContent = `Hospital Share`;
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
    const tab5Btn = document.getElementById('tab5-btn');
    const sheet1 = document.getElementById('sheet1-content');
    const sheet2 = document.getElementById('sheet2-content');
    const sheet3 = document.getElementById('sheet3-content');
    const sheet4 = document.getElementById('sheet4-content');
    const sheet5 = document.getElementById('sheet5-content');

    const switchTab = (activeBtn, activeSheet, renderFn) => {
        [tab1Btn, tab2Btn, tab3Btn, tab4Btn, tab5Btn].forEach(b => {
            if (b) {
                const isActive = (b === activeBtn);
                b.classList.toggle('active', isActive);
                b.setAttribute('aria-selected', isActive ? 'true' : 'false');
            }
        });
        [sheet1, sheet2, sheet3, sheet4, sheet5].forEach(s => {
            if (s) s.classList.toggle('active', s === activeSheet);
        });
        if (renderFn) renderFn();
    };

    if (tab1Btn) tab1Btn.addEventListener('click', () => switchTab(tab1Btn, sheet1, renderDailyTable));
    if (tab2Btn) tab2Btn.addEventListener('click', () => switchTab(tab2Btn, sheet2, renderLedgerTable));
    if (tab3Btn) tab3Btn.addEventListener('click', () => switchTab(tab3Btn, sheet3, renderAnalytics));
    if (tab4Btn) tab4Btn.addEventListener('click', () => switchTab(tab4Btn, sheet4, renderCavalryProcedures));
    if (tab5Btn) tab5Btn.addEventListener('click', () => switchTab(tab5Btn, sheet5, renderNaeemAdmitProcedures));

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

    // Patient Form Submission
    const form = document.getElementById('patient-form');
    if (form) {
        form.addEventListener('submit', (e) => {
            e.preventDefault();
            handleAddPatient();
        });
    }

    // Cavalry Procedure Form Submission
    const procForm = document.getElementById('cavalry-procedure-form');
    if (procForm) {
        procForm.addEventListener('submit', (e) => {
            e.preventDefault();
            handleAddProcedure();
        });
    }

    // Naeem Admit Form Submission
    const naeemAdmitForm = document.getElementById('naeem-admit-form');
    if (naeemAdmitForm) {
        naeemAdmitForm.addEventListener('submit', (e) => {
            e.preventDefault();
            handleAddNaeemAdmit();
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

    // Theme Toggle Button
    const themeBtn = document.getElementById('theme-toggle-btn');
    if (themeBtn) themeBtn.addEventListener('click', toggleTheme);

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

    // Patient History Modal Close Event Hooks
    const closeHistModalBtn = document.getElementById('close-history-modal-btn');
    const closeHistBtn = document.getElementById('close-history-btn');
    const histModalEl = document.getElementById('patient-history-modal');

    const closeHistModal = () => {
        if (histModalEl) histModalEl.classList.remove('active');
    };

    if (closeHistModalBtn) closeHistModalBtn.addEventListener('click', closeHistModal);
    if (closeHistBtn) closeHistBtn.addEventListener('click', closeHistModal);
    if (histModalEl) {
        histModalEl.addEventListener('click', (e) => {
            if (e.target === histModalEl) closeHistModal();
        });
    }

    // Procedure Fee Modal Event Hooks
    const procFeeForm = document.getElementById('proc-fee-form');
    const cancelProcFeeBtn = document.getElementById('cancel-proc-fee-btn');
    const procFeeModalEl = document.getElementById('procedure-fee-modal');

    const closeProcFeeModal = () => {
        if (procFeeModalEl) procFeeModalEl.classList.remove('active');
        pendingProcToCompleteId = null;
    };

    if (cancelProcFeeBtn) cancelProcFeeBtn.addEventListener('click', closeProcFeeModal);
    if (procFeeModalEl) {
        procFeeModalEl.addEventListener('click', (e) => {
            if (e.target === procFeeModalEl) closeProcFeeModal();
        });
    }

    if (procFeeForm) {
        procFeeForm.addEventListener('submit', (e) => {
            e.preventDefault();
            const feeInput = document.getElementById('proc-fee-amount');
            const feeVal = feeInput ? parseFloat(feeInput.value) : 0;

            if (isNaN(feeVal) || feeVal < 0) {
                showToast('Please enter a valid procedure amount.', 'danger');
                return;
            }

            if (pendingProcToCompleteId) {
                confirmProcedureCompletion(pendingProcToCompleteId, feeVal);
                closeProcFeeModal();
            }
        });
    }

    // Naeem Admit Decision Modal Hooks
    const naeemModalEl = document.getElementById('naeem-admit-modal');
    const closeNaeemModalBtn = document.getElementById('close-naeem-modal-btn');
    const naeemYesBtn = document.getElementById('naeem-choice-yes-btn');
    const naeemNoBtn = document.getElementById('naeem-choice-no-btn');
    const naeemProcBackBtn = document.getElementById('naeem-proc-back-btn');
    const naeemAdmitBackBtn = document.getElementById('naeem-admit-back-btn');
    const naeemProcForm = document.getElementById('naeem-proc-form');
    const naeemAdmitFeeForm = document.getElementById('naeem-admit-fee-form');

    const closeNaeemModal = () => {
        if (naeemModalEl) naeemModalEl.classList.remove('active');
        pendingNaeemAdmitId = null;
    };

    if (closeNaeemModalBtn) closeNaeemModalBtn.addEventListener('click', closeNaeemModal);
    if (naeemModalEl) {
        naeemModalEl.addEventListener('click', (e) => {
            if (e.target === naeemModalEl) closeNaeemModal();
        });
    }

    if (naeemYesBtn) {
        naeemYesBtn.addEventListener('click', () => {
            const step1 = document.getElementById('naeem-step-1');
            const procForm = document.getElementById('naeem-proc-form');
            if (step1) step1.style.display = 'none';
            if (procForm) procForm.style.display = 'block';
        });
    }

    if (naeemNoBtn) {
        naeemNoBtn.addEventListener('click', () => {
            const step1 = document.getElementById('naeem-step-1');
            const admitFeeForm = document.getElementById('naeem-admit-fee-form');
            if (step1) step1.style.display = 'none';
            if (admitFeeForm) admitFeeForm.style.display = 'block';
        });
    }

    if (naeemProcBackBtn) {
        naeemProcBackBtn.addEventListener('click', () => {
            const step1 = document.getElementById('naeem-step-1');
            const procForm = document.getElementById('naeem-proc-form');
            if (procForm) procForm.style.display = 'none';
            if (step1) step1.style.display = 'block';
        });
    }

    if (naeemAdmitBackBtn) {
        naeemAdmitBackBtn.addEventListener('click', () => {
            const step1 = document.getElementById('naeem-step-1');
            const admitFeeForm = document.getElementById('naeem-admit-fee-form');
            if (admitFeeForm) admitFeeForm.style.display = 'none';
            if (step1) step1.style.display = 'block';
        });
    }

    if (naeemProcForm) {
        naeemProcForm.addEventListener('submit', (e) => {
            e.preventDefault();
            const typeInput = document.getElementById('naeem-proc-type-input');
            const feeInput = document.getElementById('naeem-proc-fee-input');

            const procType = typeInput ? typeInput.value.trim() : '';
            const feeVal = feeInput ? parseFloat(feeInput.value) : 0;

            if (!procType || isNaN(feeVal) || feeVal < 0) {
                showToast('Please enter a valid procedure name and fee.', 'danger');
                return;
            }

            if (pendingNaeemAdmitId) {
                confirmNaeemProcedureCompletion(pendingNaeemAdmitId, procType, feeVal);
                closeNaeemModal();
            }
        });
    }

    if (naeemAdmitFeeForm) {
        naeemAdmitFeeForm.addEventListener('submit', (e) => {
            e.preventDefault();
            const feeInput = document.getElementById('naeem-admit-fee-input');
            const feeVal = feeInput ? parseFloat(feeInput.value) : 0;

            if (isNaN(feeVal) || feeVal < 0) {
                showToast('Please enter a valid admission fee.', 'danger');
                return;
            }

            if (pendingNaeemAdmitId) {
                confirmNaeemAdmissionFeeCompletion(pendingNaeemAdmitId, feeVal);
                closeNaeemModal();
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

    const now = new Date();
    const createdTimeStr = now.toLocaleTimeString('en-US', { hour: '2-digit', minute: '2-digit' });

    // Create records
    const dailyRecord = {
        id: patientId,
        dailyIndex: nextDailyIndex,
        name: name,
        date: date,
        charges: charges,
        createdTime: createdTimeStr
    };

    const ledgerRecord = {
        id: patientId,
        dailyIndex: nextDailyIndex,
        name: name,
        date: date,
        charges: charges,
        split30: charges * cfg.doctorRate,
        split70: charges * cfg.hospitalRate,
        syncStatus: initialSyncStatus,
        createdTime: createdTimeStr
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
    if (activeHospital === 'cavalry') {
        renderCavalryProcedures();
    }
    if (activeHospital === 'naeem') {
        renderNaeemAdmitProcedures();
    }
    const sheet3 = document.getElementById('sheet3-content');
    if (sheet3 && sheet3.classList.contains('active')) {
        renderAnalytics();
    }
}

// Helper: Extract clean base patient name by stripping (Procedure: ...) or (Admission Fee) tags
function extractBasePatientName(rawName) {
    if (!rawName) return '';
    let name = rawName.trim();
    name = name.replace(/\s*\((Procedure|Admission Fee).*?\)$/i, '');
    return name.trim();
}

// Populate Autocomplete past-patients-list dynamically with clean base patient names only
function updateAutocompleteSource() {
    const datalist = document.getElementById('past-patients-list');
    if (!datalist) return;
    datalist.innerHTML = '';
    
    // Extract base patient names across all patients and procedures
    const rawNames = [];
    allPatients.forEach(p => { if (p.name) rawNames.push(p.name); });
    cavalryProcedures.forEach(p => { if (p.patientName) rawNames.push(p.patientName); });
    naeemAdmitProcedures.forEach(p => { if (p.patientName) rawNames.push(p.patientName); });

    const cleanNames = rawNames
        .map(n => extractBasePatientName(n))
        .filter(n => n.length > 0);

    const uniqueNames = [...new Set(cleanNames)].sort();

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
            <td>
                <span class="patient-name-link" data-name="${escapeHtml(p.name)}" title="Click to view full patient history">
                    <span>${escapeHtml(p.name)}</span>
                    <i data-lucide="external-link" style="width:13px;height:13px;"></i>
                </span>
            </td>
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

    tbody.querySelectorAll('.patient-name-link').forEach(link => {
        link.addEventListener('click', () => {
            const name = link.getAttribute('data-name');
            openPatientHistory(name);
        });
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

    if (fDoctorLabel) fDoctorLabel.textContent = `Doctor Share`;
    if (fHospitalLabel) fHospitalLabel.textContent = `Hospital Share`;

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

        let badgeTag = `<span class="patient-badge">#${p.dailyIndex}</span>`;
        if (p.isProcedure) {
            badgeTag = '<span class="procedure-badge" title="100% Doctor Payout Procedure"><i data-lucide="stethoscope" style="width:12px;height:12px;"></i> Procedure</span>';
        } else if (p.isAdmissionFee) {
            badgeTag = '<span class="admission-badge" title="100% Doctor Payout Admission Fee"><i data-lucide="bed" style="width:12px;height:12px;"></i> Admission Fee</span>';
        }

        tr.innerHTML = `
            <td>${badgeTag}</td>
            <td>
                <span class="patient-name-link" data-name="${escapeHtml(p.name)}" title="Click to view full patient history">
                    <span>${escapeHtml(p.name)}</span>
                    <i data-lucide="external-link" style="width:13px;height:13px;"></i>
                </span>
            </td>
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

    tbody.querySelectorAll('.patient-name-link').forEach(link => {
        link.addEventListener('click', () => {
            const name = link.getAttribute('data-name');
            openPatientHistory(name);
        });
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

// Export Monthly Detailed Patient Log & Financial Settlement Report to PDF
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

    recordsToInclude.sort((a, b) => {
        if (a.date !== b.date) {
            return a.date.localeCompare(b.date);
        }
        return a.dailyIndex - b.dailyIndex;
    });

    const totalCount = recordsToInclude.length;
    const grossTotal = recordsToInclude.reduce((sum, p) => sum + p.charges, 0);
    const doctorShare = recordsToInclude.reduce((sum, p) => sum + (p.split30 !== undefined ? p.split30 : p.charges * cfg.doctorRate), 0);
    const hospitalShare = recordsToInclude.reduce((sum, p) => sum + (p.split70 !== undefined ? p.split70 : p.charges * cfg.hospitalRate), 0);

    // Populate PDF Header & Summary
    const nameEl = document.getElementById('pdf-hospital-name');
    const monthEl = document.getElementById('pdf-month-name');
    const dateEl = document.getElementById('pdf-issued-date');
    const countEl = document.getElementById('pdf-patient-count');
    const grossEl = document.getElementById('pdf-gross-revenue');
    const docShareEl = document.getElementById('pdf-doctor-share');
    const hospShareEl = document.getElementById('pdf-hospital-share');
    const docLabelEl = document.getElementById('pdf-doctor-label');
    const hospLabelEl = document.getElementById('pdf-hospital-label');
    const thDocEl = document.getElementById('pdf-th-doc');
    const thHospEl = document.getElementById('pdf-th-hosp');

    if (nameEl) nameEl.textContent = cfg.name;
    if (monthEl) monthEl.textContent = monthTitle;
    if (dateEl) dateEl.textContent = formatDateDisplay(getTodayLocalDateString());
    if (countEl) countEl.textContent = totalCount;
    if (grossEl) grossEl.textContent = formatCurrency(grossTotal);
    if (docShareEl) docShareEl.textContent = formatCurrency(doctorShare);
    if (hospShareEl) hospShareEl.textContent = formatCurrency(hospitalShare);

    if (docLabelEl) docLabelEl.textContent = `Doctor Share`;
    if (hospLabelEl) hospLabelEl.textContent = `Hospital Share`;
    if (thDocEl) thDocEl.textContent = `Doctor Share (PKR)`;
    if (thHospEl) thHospEl.textContent = `Hospital Share (PKR)`;

    // Populate Detailed Patient Table
    const tbody = document.getElementById('pdf-patient-table-body');
    if (tbody) {
        tbody.innerHTML = '';
        recordsToInclude.forEach((p, idx) => {
            const dVal = (p.split30 !== undefined ? p.split30 : p.charges * cfg.doctorRate);
            const hVal = (p.split70 !== undefined ? p.split70 : p.charges * cfg.hospitalRate);

            // Robust timestamp extraction: use createdTime, or derive from ID timestamp
            let timeVal = p.createdTime;
            if (!timeVal && p.id && !isNaN(p.id) && p.id.length >= 10) {
                try {
                    const parsedDate = new Date(parseInt(p.id));
                    if (!isNaN(parsedDate.getTime())) {
                        timeVal = parsedDate.toLocaleTimeString('en-US', { hour: '2-digit', minute: '2-digit' });
                    }
                } catch (e) {}
            }
            if (!timeVal) timeVal = '12:00 PM';
            const timeDisplay = ` | ${timeVal}`;

            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td>#${idx + 1}</td>
                <td style="font-weight: 600;">${escapeHtml(p.name)}</td>
                <td>${formatDateDisplay(p.date)}${timeDisplay}</td>
                <td style="text-align: right; font-weight: 500;">${formatCurrency(p.charges)}</td>
                <td style="text-align: right; color: #15803d; font-weight: 500;">${formatCurrency(dVal)}</td>
                <td style="text-align: right; color: #0284c7;">${formatCurrency(hVal)}</td>
            `;
            tbody.appendChild(tr);
        });
    }

    // Populate Table Footer Summary
    const footGross = document.getElementById('pdf-foot-gross');
    const footDoc = document.getElementById('pdf-foot-doc');
    const footHosp = document.getElementById('pdf-foot-hosp');

    if (footGross) footGross.textContent = formatCurrency(grossTotal);
    if (footDoc) footDoc.textContent = formatCurrency(doctorShare);
    if (footHosp) footHosp.textContent = formatCurrency(hospitalShare);

    // Trigger Print Dialog
    window.print();
}

// Open Patient History & Clinical Record Modal (Retroactive for all past & current patients)
function openPatientHistory(patientName) {
    if (!patientName) return;

    // Extract clean base patient name (e.g. "Yahya" from "Yahya (Admission Fee)")
    const basePatientName = extractBasePatientName(patientName);
    const targetNameClean = basePatientName.toLowerCase();

    // 1. Gather all records across both Naeem Surgical and Cavalry Hospital from LocalStorage
    const naeemLedgerStr = localStorage.getItem('doctor_naeem_all_patients');
    const cavalryLedgerStr = localStorage.getItem('doctor_cavalry_all_patients');

    let naeemRecords = [];
    let cavalryRecords = [];

    try { naeemRecords = naeemLedgerStr ? JSON.parse(naeemLedgerStr) : []; } catch (e) { naeemRecords = []; }
    try { cavalryRecords = cavalryLedgerStr ? JSON.parse(cavalryLedgerStr) : []; } catch (e) { cavalryRecords = []; }

    // Add hospital tag to each record
    naeemRecords.forEach(r => { r.hospitalName = 'Naeem Surgical'; r.hospId = 'naeem'; });
    cavalryRecords.forEach(r => { r.hospitalName = 'Cavalry Hospital'; r.hospId = 'cavalry'; });

    // Combine all records
    const combinedAll = [...naeemRecords, ...cavalryRecords];

    // Filter matching base patient name across all records (regular visits, procedures, admission fees)
    const patientVisits = combinedAll.filter(r => {
        if (!r.name) return false;
        const rBaseName = extractBasePatientName(r.name).toLowerCase();
        return rBaseName === targetNameClean;
    });

    if (patientVisits.length === 0) {
        showToast(`No past records found for "${basePatientName}".`, 'info');
        return;
    }

    // Sort chronologically by date and ID
    patientVisits.sort((a, b) => {
        if (a.date !== b.date) {
            return a.date.localeCompare(b.date);
        }
        return (a.dailyIndex || 0) - (b.dailyIndex || 0);
    });

    const totalVisits = patientVisits.length;
    const grossTotal = patientVisits.reduce((sum, p) => sum + (p.charges || 0), 0);
    const doctorTotal = patientVisits.reduce((sum, p) => {
        const rate = (p.hospId === 'cavalry') ? 0.70 : 0.30;
        return sum + (p.split30 !== undefined ? p.split30 : (p.charges || 0) * rate);
    }, 0);

    const firstVisitDate = formatDateDisplay(patientVisits[0].date);
    const lastVisitDate = formatDateDisplay(patientVisits[patientVisits.length - 1].date);

    // Populate Modal Elements
    const titleEl = document.getElementById('history-patient-name');
    const visitsEl = document.getElementById('history-total-visits');
    const chargesEl = document.getElementById('history-total-charges');
    const doctorEl = document.getElementById('history-total-doctor');
    const firstEl = document.getElementById('history-first-visit');
    const lastEl = document.getElementById('history-last-visit');
    const tbody = document.getElementById('history-timeline-body');

    if (titleEl) titleEl.textContent = `Patient Record: ${basePatientName}`;
    if (visitsEl) visitsEl.textContent = totalVisits;
    if (chargesEl) chargesEl.textContent = formatCurrency(grossTotal);
    if (doctorEl) doctorEl.textContent = formatCurrency(doctorTotal);
    if (firstEl) firstEl.textContent = `First: ${firstVisitDate}`;
    if (lastEl) lastEl.textContent = `Last: ${lastVisitDate}`;

    if (tbody) {
        tbody.innerHTML = '';
        patientVisits.forEach((v, index) => {
            const rate = (v.hospId === 'cavalry') ? 0.70 : 0.30;
            const docVal = (v.split30 !== undefined ? v.split30 : (v.charges || 0) * rate);
            const timeDisplay = v.createdTime ? ` | ${v.createdTime}` : '';

            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td><span class="patient-badge">#${index + 1}</span></td>
                <td>${formatDateDisplay(v.date)}${timeDisplay}</td>
                <td><span style="font-weight: 500;">${escapeHtml(v.hospitalName)}</span></td>
                <td style="text-align: right; font-weight: 500;">${formatCurrency(v.charges)}</td>
                <td style="text-align: right; color: var(--success); font-weight: 600;">${formatCurrency(docVal)}</td>
            `;
            tbody.appendChild(tr);
        });
    }

    const histModalEl = document.getElementById('patient-history-modal');
    if (histModalEl) {
        histModalEl.classList.add('active');
        if (window.lucide && typeof window.lucide.createIcons === 'function') {
            try { window.lucide.createIcons(); } catch (e) {}
        }
    }
}

// Cavalry Hospital Procedures Feature (Add, Toggle, Delete, Render Priority Schedule)
function handleAddProcedure() {
    const patientInput = document.getElementById('proc-patient-name');
    const titleInput = document.getElementById('proc-title');
    const dateInput = document.getElementById('proc-due-date');

    const patientName = patientInput ? patientInput.value.trim() : '';
    const procedureName = titleInput ? titleInput.value.trim() : '';
    const dueDate = dateInput ? dateInput.value : '';

    if (!patientName || !procedureName || !dueDate) {
        showToast('Please fill out patient name, procedure name, and due date.', 'danger');
        return;
    }

    const procRecord = {
        id: Date.now().toString(),
        patientName: patientName,
        procedureName: procedureName,
        dueDate: dueDate,
        status: 'pending',
        createdDate: getTodayLocalDateString()
    };

    cavalryProcedures.push(procRecord);
    saveState();

    if (patientInput) patientInput.value = '';
    if (titleInput) titleInput.value = '';
    if (dateInput) dateInput.value = getTodayLocalDateString();

    renderCavalryProcedures();
    showToast(`Procedure for "${patientName}" added to Cavalry priority schedule!`, 'success');
}

function handleToggleProcedureStatus(procId) {
    const proc = cavalryProcedures.find(p => p.id === procId);
    if (!proc) return;

    if (proc.status === 'completed') {
        // Toggle back to pending: prompt confirmation
        openModal(
            'Mark Procedure as Pending?',
            `Revert procedure "${proc.procedureName}" for ${proc.patientName} back to pending? (The associated ledger record will be removed).`,
            () => {
                proc.status = 'pending';
                delete proc.fee;
                // Remove corresponding procedure entry from allPatients ledger
                allPatients = allPatients.filter(p => p.procedureId !== procId);
                saveState();
                renderAll();
                showToast(`Procedure "${proc.procedureName}" marked as pending.`, 'info');
            }
        );
    } else {
        // Mark completed: prompt for procedure fee amount
        pendingProcToCompleteId = procId;
        const titleEl = document.getElementById('proc-fee-title');
        const subtitleEl = document.getElementById('proc-fee-subtitle');
        const amountInput = document.getElementById('proc-fee-amount');

        if (titleEl) titleEl.textContent = `Procedure Fee: ${proc.procedureName}`;
        if (subtitleEl) subtitleEl.textContent = `Patient: ${proc.patientName} | Due: ${formatDateDisplay(proc.dueDate)}`;
        if (amountInput) amountInput.value = proc.fee || '';

        const modalEl = document.getElementById('procedure-fee-modal');
        if (modalEl) modalEl.classList.add('active');
    }
}

function confirmProcedureCompletion(procId, feeAmount) {
    const proc = cavalryProcedures.find(p => p.id === procId);
    if (!proc) return;

    proc.status = 'completed';
    proc.fee = feeAmount;

    // Create a specialized procedure entry in All-Time Ledger with 100% Doctor Payout
    const todayStr = getTodayLocalDateString();
    const timeStr = new Date().toLocaleTimeString('en-US', { hour: '2-digit', minute: '2-digit' });
    const dayEntries = allPatients.filter(p => p.date === todayStr);
    const dailyIdx = dayEntries.length + 1;

    // Remove any existing entry for this procedure first to avoid duplicates
    allPatients = allPatients.filter(p => p.procedureId !== procId);

    const procLedgerRecord = {
        id: `proc-ledger-${Date.now()}`,
        dailyIndex: dailyIdx,
        name: `${proc.patientName} (Procedure: ${proc.procedureName})`,
        date: todayStr,
        createdTime: timeStr,
        charges: feeAmount,
        split30: feeAmount, // 100% Doctor Payout (bypasses 70/30 split)
        split70: 0,         // 0% Hospital Split
        isProcedure: true,
        procedureId: proc.id,
        syncStatus: 'local'
    };

    allPatients.unshift(procLedgerRecord);
    saveState();
    renderAll();

    showToast(`Completed procedure for "${proc.patientName}"! Added ${formatCurrency(feeAmount)} (100% Doctor Share) to All-Time Ledger.`, 'success');
}

function handleDeleteProcedure(procId) {
    const proc = cavalryProcedures.find(p => p.id === procId);
    if (!proc) return;

    openModal(
        'Delete Scheduled Procedure?',
        `Are you sure you want to delete the scheduled procedure "${proc.procedureName}" for ${proc.patientName}?`,
        () => {
            cavalryProcedures = cavalryProcedures.filter(p => p.id !== procId);
            saveState();
            renderCavalryProcedures();
            showToast('Scheduled procedure deleted.', 'info');
        }
    );
}

function renderCavalryProcedures() {
    const tbody = document.getElementById('proc-table-body');
    const emptyState = document.getElementById('proc-empty-state');
    if (!tbody) return;
    tbody.innerHTML = '';

    const todayStr = getTodayLocalDateString();
    const todayMs = new Date(todayStr + 'T00:00:00').getTime();

    // Priority Sort Engine:
    // 1. Pending procedures sorted by dueDate ASC (Earliest due date at top!)
    // 2. Completed procedures placed at bottom
    const pendingItems = cavalryProcedures.filter(p => p.status !== 'completed');
    const completedItems = cavalryProcedures.filter(p => p.status === 'completed');

    pendingItems.sort((a, b) => a.dueDate.localeCompare(b.dueDate));
    completedItems.sort((a, b) => b.dueDate.localeCompare(a.dueDate));

    const sortedProcedures = [...pendingItems, ...completedItems];

    // Priority Counts
    let urgentCount = 0;
    let soonCount = 0;

    pendingItems.forEach(p => {
        const dueMs = new Date(p.dueDate + 'T00:00:00').getTime();
        const diffDays = Math.ceil((dueMs - todayMs) / (1000 * 60 * 60 * 24));
        if (diffDays <= 0) {
            urgentCount++;
        } else if (diffDays <= 3) {
            soonCount++;
        }
    });

    const totalEl = document.getElementById('proc-total-count');
    const urgentEl = document.getElementById('proc-urgent-count');
    const soonEl = document.getElementById('proc-soon-count');

    if (totalEl) totalEl.textContent = cavalryProcedures.length;
    if (urgentEl) urgentEl.textContent = urgentCount;
    if (soonEl) soonEl.textContent = soonCount;

    if (sortedProcedures.length === 0) {
        if (emptyState) emptyState.style.display = 'block';
        return;
    }

    if (emptyState) emptyState.style.display = 'none';

    sortedProcedures.forEach(p => {
        const tr = document.createElement('tr');
        
        const dueMs = new Date(p.dueDate + 'T00:00:00').getTime();
        const diffDays = Math.ceil((dueMs - todayMs) / (1000 * 60 * 60 * 24));

        let badgeHtml = '';
        if (p.status === 'completed') {
            badgeHtml = '<span class="priority-badge completed"><i data-lucide="check-circle" style="width:13px;height:13px;"></i> Done</span>';
        } else if (diffDays < 0) {
            badgeHtml = `<span class="priority-badge urgent"><i data-lucide="alert-triangle" style="width:13px;height:13px;"></i> Overdue (${Math.abs(diffDays)}d)</span>`;
        } else if (diffDays === 0) {
            badgeHtml = '<span class="priority-badge urgent"><i data-lucide="clock" style="width:13px;height:13px;"></i> Due Today</span>';
        } else if (diffDays <= 3) {
            badgeHtml = `<span class="priority-badge soon"><i data-lucide="bell" style="width:13px;height:13px;"></i> Soon (${diffDays}d)</span>`;
        } else {
            badgeHtml = `<span class="priority-badge scheduled"><i data-lucide="calendar" style="width:13px;height:13px;"></i> ${diffDays} Days Out</span>`;
        }

        const isCompletedStyle = (p.status === 'completed') ? 'opacity: 0.6; text-decoration: line-through;' : '';

        tr.innerHTML = `
            <td>${badgeHtml}</td>
            <td style="${isCompletedStyle}">
                <span class="patient-name-link" data-name="${escapeHtml(p.patientName)}" title="Click to view full patient history">
                    <span>${escapeHtml(p.patientName)}</span>
                    <i data-lucide="external-link" style="width:13px;height:13px;"></i>
                </span>
            </td>
            <td style="font-weight: 500; ${isCompletedStyle}">${escapeHtml(p.procedureName)}</td>
            <td style="${isCompletedStyle}">${formatDateDisplay(p.dueDate)}</td>
            <td style="text-align: center;">
                <div style="display: flex; gap: 0.35rem; justify-content: center;">
                    <button class="toggle-proc-btn btn" data-id="${p.id}" title="${p.status === 'completed' ? 'Mark Pending' : 'Mark Completed'}" style="padding: 0.25rem 0.5rem; font-size: 0.75rem; background: var(--bg-primary); border: 1px solid var(--border); color: var(--text-primary);">
                        <i data-lucide="${p.status === 'completed' ? 'rotate-ccw' : 'check'}" style="width:14px;height:14px;"></i>
                    </button>
                    <button class="delete-proc-btn" data-id="${p.id}" title="Delete Procedure" style="background: none; border: none; color: var(--danger); cursor: pointer; padding: 0.25rem;">
                        <i data-lucide="trash" style="width:14px;height:14px;"></i>
                    </button>
                </div>
            </td>
        `;
        tbody.appendChild(tr);
    });

    tbody.querySelectorAll('.patient-name-link').forEach(link => {
        link.addEventListener('click', () => {
            const name = link.getAttribute('data-name');
            openPatientHistory(name);
        });
    });

    tbody.querySelectorAll('.toggle-proc-btn').forEach(btn => {
        btn.addEventListener('click', () => {
            const id = btn.getAttribute('data-id');
            handleToggleProcedureStatus(id);
        });
    });

    tbody.querySelectorAll('.delete-proc-btn').forEach(btn => {
        btn.addEventListener('click', () => {
            const id = btn.getAttribute('data-id');
            handleDeleteProcedure(id);
        });
    });

    if (window.lucide && typeof window.lucide.createIcons === 'function') {
        try { window.lucide.createIcons(); } catch (e) {}
    }
}

// Naeem Surgical Admit & Procedures Feature Functions
function handleAddNaeemAdmit() {
    const patientInput = document.getElementById('naeem-admit-patient-name');
    const dateInput = document.getElementById('naeem-admit-date');

    const patientName = patientInput ? patientInput.value.trim() : '';
    const admitDate = dateInput ? dateInput.value : '';

    if (!patientName || !admitDate) {
        showToast('Please fill out patient name and admission date.', 'danger');
        return;
    }

    const admitRecord = {
        id: Date.now().toString(),
        patientName: patientName,
        admitDate: admitDate,
        status: 'admitted',
        createdDate: getTodayLocalDateString()
    };

    naeemAdmitProcedures.unshift(admitRecord);
    saveState();

    if (patientInput) patientInput.value = '';
    if (dateInput) dateInput.value = getTodayLocalDateString();

    renderNaeemAdmitProcedures();
    showToast(`Admitted patient "${patientName}" added!`, 'success');
}

function handleToggleNaeemAdmitStatus(admitId) {
    const record = naeemAdmitProcedures.find(r => r.id === admitId);
    if (!record) return;

    if (record.status === 'completed') {
        // Toggle back to admitted: confirm prompt
        openModal(
            'Revert Patient Status to Admitted?',
            `Revert patient "${record.patientName}" back to admitted status? (Associated ledger record will be removed).`,
            () => {
                record.status = 'admitted';
                delete record.fee;
                delete record.entryType;
                delete record.procedureName;
                // Remove corresponding entry from Naeem allPatients ledger
                allPatients = allPatients.filter(p => p.admitId !== admitId);
                saveState();
                renderAll();
                showToast(`Status for "${record.patientName}" reverted to Admitted.`, 'info');
            }
        );
    } else {
        // Open decision modal
        pendingNaeemAdmitId = admitId;
        const titleEl = document.getElementById('naeem-modal-patient-title');
        const step1 = document.getElementById('naeem-step-1');
        const procForm = document.getElementById('naeem-proc-form');
        const admitForm = document.getElementById('naeem-admit-fee-form');

        if (titleEl) titleEl.innerHTML = `<i data-lucide="user-check" style="color: var(--primary);"></i> <span>Patient Admission: ${record.patientName}</span>`;
        if (step1) step1.style.display = 'block';
        if (procForm) procForm.style.display = 'none';
        if (admitForm) admitForm.style.display = 'none';

        const procTypeInput = document.getElementById('naeem-proc-type-input');
        const procFeeInput = document.getElementById('naeem-proc-fee-input');
        const admitFeeInput = document.getElementById('naeem-admit-fee-input');

        if (procTypeInput) procTypeInput.value = '';
        if (procFeeInput) procFeeInput.value = '';
        if (admitFeeInput) admitFeeInput.value = '';

        const modalEl = document.getElementById('naeem-admit-modal');
        if (modalEl) {
            modalEl.classList.add('active');
            if (window.lucide && typeof window.lucide.createIcons === 'function') {
                try { window.lucide.createIcons(); } catch (e) {}
            }
        }
    }
}

function confirmNaeemProcedureCompletion(admitId, procType, feeAmount) {
    const record = naeemAdmitProcedures.find(r => r.id === admitId);
    if (!record) return;

    record.status = 'completed';
    record.entryType = 'procedure';
    record.procedureName = procType;
    record.fee = feeAmount;

    // Create 100% Doctor Payout Procedure Ledger Entry
    const todayStr = getTodayLocalDateString();
    const timeStr = new Date().toLocaleTimeString('en-US', { hour: '2-digit', minute: '2-digit' });
    const dayEntries = allPatients.filter(p => p.date === todayStr);
    const dailyIdx = dayEntries.length + 1;

    allPatients = allPatients.filter(p => p.admitId !== admitId);

    const procLedgerRecord = {
        id: `naeem-proc-${Date.now()}`,
        dailyIndex: dailyIdx,
        name: `${record.patientName} (Procedure: ${procType})`,
        date: todayStr,
        createdTime: timeStr,
        charges: feeAmount,
        split30: feeAmount, // 100% Doctor Payout
        split70: 0,         // 0% Hospital Split
        isProcedure: true,
        admitId: record.id,
        syncStatus: 'local'
    };

    allPatients.unshift(procLedgerRecord);
    saveState();
    renderAll();

    showToast(`Procedure "${procType}" completed for "${record.patientName}"! ${formatCurrency(feeAmount)} (100% Doctor Share) posted to Ledger.`, 'success');
}

function confirmNaeemAdmissionFeeCompletion(admitId, feeAmount) {
    const record = naeemAdmitProcedures.find(r => r.id === admitId);
    if (!record) return;

    record.status = 'completed';
    record.entryType = 'admission';
    record.fee = feeAmount;

    // Create 100% Doctor Payout Admission Fee Ledger Entry
    const todayStr = getTodayLocalDateString();
    const timeStr = new Date().toLocaleTimeString('en-US', { hour: '2-digit', minute: '2-digit' });
    const dayEntries = allPatients.filter(p => p.date === todayStr);
    const dailyIdx = dayEntries.length + 1;

    allPatients = allPatients.filter(p => p.admitId !== admitId);

    const admitLedgerRecord = {
        id: `naeem-admit-${Date.now()}`,
        dailyIndex: dailyIdx,
        name: `${record.patientName} (Admission Fee)`,
        date: todayStr,
        createdTime: timeStr,
        charges: feeAmount,
        split30: feeAmount, // 100% Doctor Payout
        split70: 0,         // 0% Hospital Split
        isAdmissionFee: true,
        admitId: record.id,
        syncStatus: 'local'
    };

    allPatients.unshift(admitLedgerRecord);
    saveState();
    renderAll();

    showToast(`Admission fee for "${record.patientName}" completed! ${formatCurrency(feeAmount)} (100% Doctor Share) posted to Ledger.`, 'success');
}

function handleDeleteNaeemAdmit(admitId) {
    const record = naeemAdmitProcedures.find(r => r.id === admitId);
    if (!record) return;

    openModal(
        'Delete Admitted Patient Record?',
        `Are you sure you want to delete the record for ${record.patientName}?`,
        () => {
            naeemAdmitProcedures = naeemAdmitProcedures.filter(r => r.id !== admitId);
            allPatients = allPatients.filter(p => p.admitId !== admitId);
            saveState();
            renderAll();
            showToast('Admitted patient record deleted.', 'info');
        }
    );
}

function renderNaeemAdmitProcedures() {
    const tbody = document.getElementById('naeem-admit-table-body');
    const emptyState = document.getElementById('naeem-admit-empty-state');
    if (!tbody) return;
    tbody.innerHTML = '';

    const currentAdmitted = naeemAdmitProcedures.filter(r => r.status === 'admitted');
    const completedAdmitted = naeemAdmitProcedures.filter(r => r.status === 'completed');

    const totalEl = document.getElementById('naeem-admit-total');
    const currentEl = document.getElementById('naeem-admit-current');
    const completedEl = document.getElementById('naeem-admit-completed');

    if (totalEl) totalEl.textContent = naeemAdmitProcedures.length;
    if (currentEl) currentEl.textContent = currentAdmitted.length;
    if (completedEl) completedEl.textContent = completedAdmitted.length;

    if (naeemAdmitProcedures.length === 0) {
        if (emptyState) emptyState.style.display = 'block';
        return;
    }

    if (emptyState) emptyState.style.display = 'none';

    naeemAdmitProcedures.forEach(r => {
        const tr = document.createElement('tr');

        let statusBadgeHtml = '';
        if (r.status === 'completed') {
            if (r.entryType === 'procedure') {
                statusBadgeHtml = `<span class="priority-badge completed" title="Procedure: ${escapeHtml(r.procedureName || '')}"><i data-lucide="check-circle" style="width:13px;height:13px;"></i> Procedure Done</span>`;
            } else {
                statusBadgeHtml = '<span class="priority-badge completed"><i data-lucide="check-circle" style="width:13px;height:13px;"></i> Admission Fee</span>';
            }
        } else {
            statusBadgeHtml = '<span class="priority-badge scheduled" style="background:#e0f2fe; color:#0284c7; border-color:#7dd3fc;"><i data-lucide="bed" style="width:13px;height:13px;"></i> Admitted</span>';
        }

        const isCompletedStyle = (r.status === 'completed') ? 'opacity: 0.6; text-decoration: line-through;' : '';

        tr.innerHTML = `
            <td>${statusBadgeHtml}</td>
            <td style="${isCompletedStyle}">
                <span class="patient-name-link" data-name="${escapeHtml(r.patientName)}" title="Click to view full patient history">
                    <span>${escapeHtml(r.patientName)}</span>
                    <i data-lucide="external-link" style="width:13px;height:13px;"></i>
                </span>
            </td>
            <td style="${isCompletedStyle}">${formatDateDisplay(r.admitDate)}</td>
            <td style="text-align: center;">
                <div style="display: flex; gap: 0.35rem; justify-content: center;">
                    <button class="toggle-naeem-admit-btn btn" data-id="${r.id}" title="${r.status === 'completed' ? 'Revert to Admitted' : 'Mark Done'}" style="padding: 0.25rem 0.5rem; font-size: 0.75rem; background: var(--bg-primary); border: 1px solid var(--border); color: var(--text-primary);">
                        <i data-lucide="${r.status === 'completed' ? 'rotate-ccw' : 'check'}" style="width:14px;height:14px;"></i>
                    </button>
                    <button class="delete-naeem-admit-btn" data-id="${r.id}" title="Delete Record" style="background: none; border: none; color: var(--danger); cursor: pointer; padding: 0.25rem;">
                        <i data-lucide="trash" style="width:14px;height:14px;"></i>
                    </button>
                </div>
            </td>
        `;
        tbody.appendChild(tr);
    });

    tbody.querySelectorAll('.patient-name-link').forEach(link => {
        link.addEventListener('click', () => {
            const name = link.getAttribute('data-name');
            openPatientHistory(name);
        });
    });

    tbody.querySelectorAll('.toggle-naeem-admit-btn').forEach(btn => {
        btn.addEventListener('click', () => {
            const id = btn.getAttribute('data-id');
            handleToggleNaeemAdmitStatus(id);
        });
    });

    tbody.querySelectorAll('.delete-naeem-admit-btn').forEach(btn => {
        btn.addEventListener('click', () => {
            const id = btn.getAttribute('data-id');
            handleDeleteNaeemAdmit(id);
        });
    });

    if (window.lucide && typeof window.lucide.createIcons === 'function') {
        try { window.lucide.createIcons(); } catch (e) {}
    }
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

    const isDark = currentTheme === 'dark';
    const gridColor = isDark ? 'rgba(255, 255, 255, 0.1)' : 'rgba(0, 0, 0, 0.08)';
    const textColor = isDark ? '#94a3b8' : '#64748b';

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
                    borderColor: isDark ? '#38bdf8' : '#0284c7',
                    backgroundColor: isDark ? 'rgba(56, 189, 248, 0.15)' : 'rgba(2, 132, 199, 0.15)',
                    fill: true,
                    tension: 0.3,
                    pointRadius: 4,
                    pointBackgroundColor: isDark ? '#38bdf8' : '#0284c7'
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: { legend: { display: false } },
                scales: {
                    x: { ticks: { color: textColor }, grid: { color: gridColor } },
                    y: { beginAtZero: true, ticks: { precision: 0, color: textColor }, grid: { color: gridColor } }
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
                        backgroundColor: isDark ? '#34d399' : '#10b981'
                    },
                    {
                        label: `Hospital Share (${cfg.hospitalPercentLabel})`,
                        data: dailyHospitalShares,
                        backgroundColor: isDark ? '#38bdf8' : '#0284c7'
                    }
                ]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: { legend: { position: 'top', labels: { color: textColor } } },
                scales: {
                    x: { stacked: true, ticks: { color: textColor }, grid: { color: gridColor } },
                    y: { stacked: true, beginAtZero: true, ticks: { color: textColor }, grid: { color: gridColor } }
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
