// Global Variables
let currentUser = null;
let users = [];
let productionData = [];
let feedbackMessages = [];

// --- Chart handles (safe init) ---
window.utilisationChart = null;
window.stackRankingChart = null;
window.userPerformanceChart = null;

// Manager charts
window.managerProjectChart = null;
window.managerTeamChart = null;

function safeDestroyChart(refName) {
  const ch = window[refName];
  if (ch && typeof ch.destroy === 'function') {
    try { 
      ch.destroy(); 
    } catch (e) { 
      console.warn(refName, 'destroy failed', e); 
    }
  }
  window[refName] = null;
}

// Sample data for demonstration
const sampleUsers = [
    { email: 'admin@dm2.com', password: 'admin123', role: 'admin', status: 'active' },
    { email: 'manager@dm2.com', password: 'manager123', role: 'manager', status: 'active' },
    { email: 'praveen.kumar@invensis.net', password: 'praveen123', role: 'manager', status: 'active' },
    { email: 'tl@dm2.com', password: 'tl123', role: 'tl', status: 'active' }
];

// Sample feedback messages for demonstration
const sampleFeedbackMessages = [
    {
        id: 1,
        from: 'manager@dm2.com',
        to: 'tl@dm2.com',
        message: 'Great work on the project performance this month! Keep it up.',
        timestamp: new Date().toISOString(),
        parentId: null
    },
    {
        id: 2,
        from: 'tl@dm2.com',
        to: 'manager@dm2.com',
        message: 'Thank you for the feedback! The team is motivated to do even better.',
        timestamp: new Date(Date.now() + 60000).toISOString(),
        parentId: 1
    },
    {
        id: 3,
        from: 'praveen.kumar@invensis.net',
        to: 'tl@dm2.com',
        message: 'Please review the productivity metrics and provide suggestions for improvement.',
        timestamp: new Date(Date.now() + 120000).toISOString(),
        parentId: null
    }
];

// Real data based on your exact format: Date, Employee ID, Name, Client Name, Process Name, Productivity, Target, Client Errors, Internal Errors, Hours Worked, Actual Hours
const sampleProductionData = [
    { date: '4/21/2025', employeeId: 'ITPL6647', name: 'Ramya', clientName: 'Widespread Electric', processName: 'Billing Invoice', productivity: 119, target: 119, clientErrors: 0, internalErrors: 0, hoursWorked: 8, actualHours: 8 },
    { date: '4/21/2025', employeeId: 'ITPL9648', name: 'Geetha', clientName: 'Corr-Jensen', processName: 'Account Payables', productivity: 18, target: 18, clientErrors: 0, internalErrors: 0, hoursWorked: 8, actualHours: 8 },
    { date: '4/21/2025', employeeId: 'ITPL8058', name: 'Giri Prasad R', clientName: 'Soho Studio', processName: 'ST - Order Entry', productivity: 65, target: 65, clientErrors: 0, internalErrors: 0, hoursWorked: 8, actualHours: 8 },
    { date: '4/21/2025', employeeId: 'ITPL10359', name: 'A Dharani', clientName: 'Soho Studio', processName: 'ST - Order Entry', productivity: 92, target: 92, clientErrors: 0, internalErrors: 0, hoursWorked: 8, actualHours: 8 },
    { date: '4/21/2025', employeeId: 'ITPL8360', name: 'Meddabalmi Vishnuchakram', clientName: 'Soho Studio', processName: 'ST - Claims', productivity: 28, target: 28, clientErrors: 0, internalErrors: 0, hoursWorked: 8, actualHours: 8 },
    { date: '4/21/2025', employeeId: 'ITPL9668', name: 'K Kiruba Karan', clientName: 'Medscan Lab', processName: 'Blood Samples', productivity: 3, target: 3, clientErrors: 0, internalErrors: 0, hoursWorked: 8, actualHours: 8 },
    { date: '4/21/2025', employeeId: 'ITPL11680', name: 'M Deepika', clientName: 'Medscan Lab', processName: 'Blood Samples', productivity: 19, target: 19, clientErrors: 0, internalErrors: 0, hoursWorked: 8, actualHours: 8 },
    { date: '4/21/2025', employeeId: 'ITPL10376', name: 'Yavanika', clientName: 'Curexa Pharmacy', processName: 'Curexa - OE', productivity: 196, target: 220, clientErrors: 0, internalErrors: 0, hoursWorked: 8, actualHours: 8 },
    { date: '4/21/2025', employeeId: 'RITPL5165', name: 'Surya Narayan', clientName: 'Curexa Pharmacy', processName: 'Curexa - OE', productivity: 225, target: 220, clientErrors: 0, internalErrors: 0, hoursWorked: 8, actualHours: 8 },
    { date: '4/21/2025', employeeId: 'ITPL10723', name: 'Meghana', clientName: 'Curexa Pharmacy', processName: 'Curexa - OE', productivity: 159, target: 220, clientErrors: 0, internalErrors: 0, hoursWorked: 8, actualHours: 8 },
    { date: '4/21/2025', employeeId: 'RITPL5298', name: 'Shyam Kumar', clientName: 'Curexa Pharmacy', processName: 'Curexa - OE', productivity: 166, target: 220, clientErrors: 0, internalErrors: 0, hoursWorked: 8, actualHours: 8 },
    { date: '4/21/2025', employeeId: 'ITPL9727', name: 'Mounika', clientName: 'Curexa Pharmacy', processName: 'Curexa - OE', productivity: 209, target: 220, clientErrors: 0, internalErrors: 0, hoursWorked: 8, actualHours: 8 },
    { date: '4/21/2025', employeeId: 'ITPL9146', name: 'Nagendra', clientName: 'Curexa Pharmacy', processName: 'Curexa - OE', productivity: 119, target: 220, clientErrors: 0, internalErrors: 0, hoursWorked: 8, actualHours: 8 },
    { date: '4/21/2025', employeeId: 'RITPL5544', name: 'Pavan Kumar', clientName: 'Curexa Pharmacy', processName: 'Curexa - OE', productivity: 0, target: 220, clientErrors: 0, internalErrors: 0, hoursWorked: 0, actualHours: 0 },
    { date: '4/21/2025', employeeId: 'ITPL9646', name: 'Kusuma', clientName: 'Curexa Pharmacy', processName: 'Curexa - OE', productivity: 0, target: 220, clientErrors: 0, internalErrors: 0, hoursWorked: 0, actualHours: 0 },
    { date: '4/21/2025', employeeId: 'ITPL8128', name: 'Rajesh T', clientName: 'Hit Promo', processName: 'Hit - QC', productivity: 60, target: 60, clientErrors: 0, internalErrors: 0, hoursWorked: 8, actualHours: 8 },
    { date: '4/21/2025', employeeId: 'ITPL8326', name: 'Uday Kiran Y', clientName: 'Hit Promo', processName: 'Hit - QC', productivity: 60, target: 60, clientErrors: 0, internalErrors: 0, hoursWorked: 8, actualHours: 8 },
    
    // Additional comprehensive data for consistency across devices
    { date: '4/22/2025', employeeId: 'ITPL9621', name: 'Mahesh N', clientName: 'Negotiatus', processName: 'Data Processing', productivity: 150, target: 120, clientErrors: 0, internalErrors: 0, hoursWorked: 8, actualHours: 8 },
    { date: '4/22/2025', employeeId: 'ITPL10376', name: 'Yavanika', clientName: 'Curexa Pharmacy', processName: 'Curexa - OE', productivity: 180, target: 160, clientErrors: 0, internalErrors: 0, hoursWorked: 8, actualHours: 8 },
    { date: '4/22/2025', employeeId: 'RITPL5165', name: 'Surya Narayan', clientName: 'Curexa Pharmacy', processName: 'Curexa - OE', productivity: 225, target: 200, clientErrors: 0, internalErrors: 0, hoursWorked: 8, actualHours: 8 },
    { date: '4/22/2025', employeeId: 'ITPL9146', name: 'Nagendra', clientName: 'Curexa Pharmacy', processName: 'Curexa - OE', productivity: 190, target: 180, clientErrors: 0, internalErrors: 0, hoursWorked: 8, actualHours: 8 },
    { date: '4/22/2025', employeeId: 'RITPL5298', name: 'Shyam Kumar', clientName: 'Curexa Pharmacy', processName: 'Curexa - OE', productivity: 170, target: 160, clientErrors: 0, internalErrors: 0, hoursWorked: 8, actualHours: 8 },
    { date: '4/22/2025', employeeId: 'ITPL9448', name: 'A Hemanth', clientName: 'Cali Bamboo', processName: 'Order Processing', productivity: 140, target: 120, clientErrors: 0, internalErrors: 0, hoursWorked: 8, actualHours: 8 },
    { date: '4/22/2025', employeeId: 'ITPL10904', name: 'Sujatha R D', clientName: 'Hit Promo', processName: 'Hit - QC', productivity: 120, target: 100, clientErrors: 1, internalErrors: 0, hoursWorked: 8, actualHours: 8 },
    
    // May 2025 data for month filtering
    { date: '5/15/2025', employeeId: 'ITPL9621', name: 'Mahesh N', clientName: 'Negotiatus', processName: 'Data Processing', productivity: 200, target: 180, clientErrors: 0, internalErrors: 0, hoursWorked: 8, actualHours: 8 },
    { date: '5/15/2025', employeeId: 'ITPL10376', name: 'Yavanika', clientName: 'Curexa Pharmacy', processName: 'Curexa - OE', productivity: 220, target: 200, clientErrors: 0, internalErrors: 0, hoursWorked: 8, actualHours: 8 },
    { date: '5/15/2025', employeeId: 'RITPL5165', name: 'Surya Narayan', clientName: 'Curexa Pharmacy', processName: 'Curexa - OE', productivity: 250, target: 220, clientErrors: 0, internalErrors: 0, hoursWorked: 8, actualHours: 8 }
];

// Enhanced initialization with better error handling
document.addEventListener('DOMContentLoaded', function() {
    // Check for required libraries
    checkRequiredLibraries();
    
    // Clear old cached sample data only if schema changed
    if (!localStorage.getItem('dm2_production_data')) {
        localStorage.removeItem('dm2_production_data');
    }
    
    // Initialize with error handling
    try {
    initializeApp();
    } catch (error) {
        console.error('Initialization failed:', error);
        showInitializationError(error);
    }
    
    // Set default Google Sheets link for cross-browser compatibility
    setDefaultGoogleSheetsLink();
    
    // Auto-load data if user is already logged in
    setTimeout(() => {
        if (currentUser) {
            console.log('Auto-loading data for logged in user:', currentUser.role);
            switch(currentUser.role) {
                case 'tl':
                    loadProductionData();
                    break;
                case 'manager':
                    loadManagerData();
                    break;
            }
        }
    }, 1000);
});

// Set default Google Sheets link that works across all browsers
function setDefaultGoogleSheetsLink() {
    const defaultSheetLink = 'https://docs.google.com/spreadsheets/d/1SUYakZKT_m4sw2xkFbuRnqX9r8_3qP-v/edit?usp=sharing&ouid=110973781089127610610&rtpof=true&sd=true';
    
    // Check if a Google Sheets link is already saved (check both possible keys)
    let savedLink = localStorage.getItem('dm2_google_sheets_link') || localStorage.getItem('dm2_sheet_url');
    
    if (!savedLink) {
        // Set default link if none exists (save to both keys for compatibility)
        localStorage.setItem('dm2_google_sheets_link', defaultSheetLink);
        localStorage.setItem('dm2_sheet_url', defaultSheetLink);
        savedLink = defaultSheetLink;
        console.log('Set default Google Sheets link for cross-browser compatibility');
    }
    
    // Update the UI to show the saved link (try multiple input field IDs)
    const sheetLinkInput = document.getElementById('sheetLink') || document.getElementById('googleSheetUrl');
    if (sheetLinkInput) {
        sheetLinkInput.value = savedLink;
        console.log('Updated sheet link input field with:', savedLink);
    }
    
    // Update the saved link display
    updateSavedLinkDisplay();
    
    // Also try to load data immediately if we have a valid link
    if (savedLink && savedLink.includes('docs.google.com')) {
        console.log('Auto-loading data from saved Google Sheets link');
        setTimeout(() => {
            loadDataFromGoogleSheetsAsync(savedLink);
        }, 1000);
    }
}

// Update the saved link display in the UI
function updateSavedLinkDisplay() {
    // Check both possible storage keys
    const savedLink = localStorage.getItem('dm2_google_sheets_link') || localStorage.getItem('dm2_sheet_url');
    const savedLinkDiv = document.getElementById('savedLink') || document.getElementById('savedSheetUrlHint');
    
    if (savedLinkDiv) {
        if (savedLink) {
            savedLinkDiv.innerHTML = `Saved: <a href="${savedLink}" target="_blank" style="color: #007bff; text-decoration: underline;">${savedLink}</a>`;
            savedLinkDiv.style.display = 'block';
            console.log('Updated saved link display with:', savedLink);
        } else {
            savedLinkDiv.innerHTML = 'No sheet link saved yet.';
            savedLinkDiv.style.display = 'block';
            console.log('No saved link found, showing default message');
        }
    } else {
        console.warn('Could not find saved link display element');
    }
}

function checkRequiredLibraries() {
    const missingLibraries = [];
    
    if (typeof Chart === 'undefined') {
        missingLibraries.push('Chart.js');
    }
    
    if (typeof XLSX === 'undefined') {
        missingLibraries.push('SheetJS (XLSX)');
    }
    
    if (missingLibraries.length > 0) {
        console.warn('Missing libraries:', missingLibraries);
        // Don't block initialization, but show warning
        setTimeout(() => {
            if (missingLibraries.includes('Chart.js')) {
                console.error('Chart.js not loaded - charts will not work');
            }
            if (missingLibraries.includes('SheetJS (XLSX)')) {
                console.error('SheetJS not loaded - Excel upload will not work');
            }
        }, 1000);
    }
}

function showInitializationError(error) {
    const errorDiv = document.createElement('div');
    errorDiv.style.cssText = `
        position: fixed;
        top: 20px;
        right: 20px;
        background: #dc3545;
        color: white;
        padding: 15px;
        border-radius: 5px;
        z-index: 10000;
        max-width: 300px;
    `;
    errorDiv.innerHTML = `
        <strong>Initialization Error</strong><br>
        ${error.message}<br>
        <small>Please refresh the page or check your internet connection.</small>
    `;
    document.body.appendChild(errorDiv);
    
    // Auto-remove after 10 seconds
    setTimeout(() => {
        if (errorDiv.parentNode) {
            errorDiv.parentNode.removeChild(errorDiv);
        }
    }, 10000);
}

function initializeApp() {
    console.log('Initializing application...');
    
    // Always ensure users data is available (for consistency across devices)
    const savedUsers = localStorage.getItem('dm2_users');
    if (savedUsers) {
        try {
        users = JSON.parse(savedUsers);
            console.log('Loaded users from localStorage:', users.length);
        } catch (error) {
            console.warn('Failed to parse users data, using sample data:', error);
            users = [...sampleUsers];
            localStorage.setItem('dm2_users', JSON.stringify(users));
        }
    } else {
        console.log('No users in localStorage, using sample data');
        users = [...sampleUsers];
        localStorage.setItem('dm2_users', JSON.stringify(users));
    }
    
    // Always ensure production data is available (for consistency across devices)
    const existing = localStorage.getItem('dm2_production_data');
    if (existing) {
        try {
            productionData = JSON.parse(existing) || [];
            console.log('Loaded production data from localStorage:', productionData.length, 'records');
        } catch (error) {
            console.warn('Failed to parse production data, using sample data:', error);
            productionData = [...sampleProductionData];
            localStorage.setItem('dm2_production_data', JSON.stringify(productionData));
        }
    } else {
        console.log('No production data in localStorage, using sample data');
        productionData = [...sampleProductionData];
        localStorage.setItem('dm2_production_data', JSON.stringify(productionData));
    }
    
    // Always ensure feedback messages are available
    const savedFeedback = localStorage.getItem('dm2_feedback');
    if (savedFeedback) {
        try {
        feedbackMessages = JSON.parse(savedFeedback);
            console.log('Loaded feedback messages from localStorage:', feedbackMessages.length);
        } catch (error) {
            console.warn('Failed to parse feedback data:', error);
            feedbackMessages = [...sampleFeedbackMessages];
            localStorage.setItem('dm2_feedback', JSON.stringify(feedbackMessages));
        }
    } else {
        console.log('No feedback messages in localStorage, using sample data');
        feedbackMessages = [...sampleFeedbackMessages];
        localStorage.setItem('dm2_feedback', JSON.stringify(feedbackMessages));
    }
    
    // Set up event listeners
    setupEventListeners();
    
    // Check for existing session
    checkExistingSession();
    
    console.log('Application initialization complete');
}

function checkExistingSession() {
    try {
        const sessionData = sessionStorage.getItem('dm2_current_user');
        if (sessionData) {
            const session = JSON.parse(sessionData);
            const user = users.find(u => u.email === session.email);
            
            if (user && user.status === 'active') {
                currentUser = user;
                console.log('Restored session for:', user.email);
                
                // Redirect to appropriate page
                switch(user.role) {
                    case 'admin':
                        showPage('adminPage');
                        return;
                    case 'manager':
                        showPage('managerPage');
                        return;
                    case 'tl':
                        showPage('tlPage');
                        return;
                }
            }
        }
    } catch (error) {
        console.warn('Failed to restore session:', error);
        sessionStorage.removeItem('dm2_current_user');
    }
    
    // If no valid session, show login page
    showPage('loginPage');
}

function setupEventListeners() {
    // Login form
    const loginForm = document.getElementById('loginForm');
    if (loginForm) {
        loginForm.addEventListener('submit', handleLogin);
        console.log('Login form event listener attached');
    } else {
        console.error('Login form element not found');
    }
    
    // Admin form
    const createUserForm = document.getElementById('createUserForm');
    if (createUserForm) {
        createUserForm.addEventListener('submit', handleCreateUser);
        console.log('Create user form event listener attached');
    } else {
        console.error('Create user form element not found');
    }
    
    // Manager feedback form
    const managerFeedbackForm = document.getElementById('managerFeedbackForm');
    if (managerFeedbackForm) {
        managerFeedbackForm.addEventListener('submit', handleManagerFeedback);
        console.log('Manager feedback form event listener attached');
    } else {
        console.error('Manager feedback form element not found');
    }
    
    // Excel upload - wait for DOM to be ready
    setTimeout(() => {
        const excelUpload = document.getElementById('excelUpload');
        if (excelUpload) {
            excelUpload.addEventListener('change', handleExcelUpload);
            console.log('Excel upload event listener attached');
        } else {
            console.error('Excel upload element not found');
        }
    }, 100);
}

// Page Navigation
function showPage(pageId) {
    // Hide all pages
    document.querySelectorAll('.page').forEach(page => {
        page.classList.remove('active');
    });
    
    // Show selected page
    document.getElementById(pageId).classList.add('active');
    
    // Load page-specific data
    switch(pageId) {
        case 'adminPage':
            loadAdminData();
            break;
        case 'managerPage':
            loadManagerPage();
            break;
        case 'tlPage':
            loadTLPage();
            break;
    }
}

// Login Functionality with enhanced error handling
function handleLogin(e) {
    e.preventDefault();
    console.log('Login form submitted');
    
    const email = document.getElementById('email').value.trim();
    const password = document.getElementById('password').value;
    const errorDiv = document.getElementById('loginError');
    
    console.log('Login attempt:', email, password);
    
    // Clear previous errors
    errorDiv.classList.remove('show');
    errorDiv.textContent = '';
    
    // Basic validation
    if (!email || !password) {
        errorDiv.textContent = 'Please enter both email and password';
        errorDiv.classList.add('show');
        return;
    }
    
    console.log('Available users:', users);
    
    // Find user with case-insensitive email comparison
    const user = users.find(u => 
        u.email.toLowerCase() === email.toLowerCase() && 
        u.password === password
    );
    
    console.log('Found user:', user);
    
    if (user && user.status === 'active') {
        currentUser = user;
        
        // Store login state for persistence across page refreshes
        try {
            sessionStorage.setItem('dm2_current_user', JSON.stringify({
                email: user.email,
                role: user.role,
                loginTime: new Date().toISOString()
            }));
        } catch (error) {
            console.warn('Could not save login state:', error);
        }
        
        // Redirect based on role
        switch(user.role) {
            case 'admin':
                showPage('adminPage');
                break;
            case 'manager':
                showPage('managerPage');
                break;
            case 'tl':
                showPage('tlPage');
                break;
            default:
                errorDiv.textContent = 'Invalid user role';
                errorDiv.classList.add('show');
                return;
        }
        
        console.log('Login successful for:', user.email, 'Role:', user.role);
    } else {
        errorDiv.textContent = 'Invalid email or password';
        errorDiv.classList.add('show');
        console.log('Login failed for:', email);
    }
}

// Logout Functionality
function logout() {
    currentUser = null;
    
    // Clear session data
    try {
        sessionStorage.removeItem('dm2_current_user');
    } catch (error) {
        console.warn('Could not clear session data:', error);
    }
    
    showPage('loginPage');
    document.getElementById('loginForm').reset();
    
    // Clear any error messages
    const errorDiv = document.getElementById('loginError');
    if (errorDiv) {
        errorDiv.classList.remove('show');
        errorDiv.textContent = '';
    }
    
    console.log('User logged out successfully');
}

// Admin Portal Functions
function loadAdminData() {
    if (currentUser && currentUser.role === 'admin') {
        document.getElementById('adminUserEmail').textContent = currentUser.email;
        loadUsersTable();
    }
}

function loadUsersTable() {
    const tbody = document.getElementById('usersTableBody');
    tbody.innerHTML = '';
    
    users.forEach(user => {
        if (user.role !== 'admin') { // Don't show admin users in the table
            const row = document.createElement('tr');
            row.innerHTML = `
                <td>${user.email}</td>
                <td>${user.role === 'manager' ? 'Manager' : 'TL & Coordinator'}</td>
                <td><span class="status-${user.status}">${user.status}</span></td>
                <td>
                    <button onclick="toggleUserStatus('${user.email}')" class="btn btn-secondary">
                        ${user.status === 'active' ? 'Deactivate' : 'Activate'}
                    </button>
                </td>
            `;
            tbody.appendChild(row);
        }
    });
}

function handleCreateUser(e) {
    e.preventDefault();
    
    const email = document.getElementById('newUserEmail').value;
    const role = document.getElementById('newUserRole').value;
    const password = document.getElementById('newUserPassword').value;
    
    // Check if user already exists
    if (users.find(u => u.email === email)) {
        alert('User with this email already exists');
        return;
    }
    
    // Create new user
    const newUser = {
        email: email,
        password: password,
        role: role,
        status: 'active'
    };
    
    users.push(newUser);
    localStorage.setItem('dm2_users', JSON.stringify(users));
    
    // Reset form
    document.getElementById('createUserForm').reset();
    
    // Reload table
    loadUsersTable();
    
    alert('User created successfully!');
}

function toggleUserStatus(email) {
    const user = users.find(u => u.email === email);
    if (user) {
        user.status = user.status === 'active' ? 'inactive' : 'active';
        localStorage.setItem('dm2_users', JSON.stringify(users));
        loadUsersTable();
    }
}

// Manager Portal Functions
function loadManagerPage() {
    if (currentUser && currentUser.role === 'manager') {
        document.getElementById('managerUserEmail').textContent = currentUser.email;
        
        // Load production data first
        const savedData = localStorage.getItem('dm2_production_data');
        if (savedData) {
            productionData = JSON.parse(savedData);
        }
        
        // Set up Google Sheets link for this page
        setDefaultGoogleSheetsLink();
        
        // Load filters and data
        loadManagerFilters();
        loadManagerData();
        loadManagerFeedbackMessages();
    }
}

function loadManagerFilters() {
    // Load projects - use clientName instead of project
    const projectFilter = document.getElementById('managerProjectFilter');
    if (!projectFilter) return;
    
    const projects = [...new Set(productionData.map(d => d.clientName).filter(Boolean))];
    projectFilter.innerHTML = '<option value="">All Projects</option>';
    projects.forEach(project => {
        const option = document.createElement('option');
        option.value = project;
        option.textContent = project;
        projectFilter.appendChild(option);
    });
    
    // Load TLs (for feedback)
    const feedbackTL = document.getElementById('feedbackTL');
    if (feedbackTL) {
    const tls = users.filter(u => u.role === 'tl' && u.status === 'active');
    feedbackTL.innerHTML = '<option value="">Select TL/Coordinator</option>';
    tls.forEach(tl => {
        const option = document.createElement('option');
        option.value = tl.email;
        option.textContent = tl.email;
        feedbackTL.appendChild(option);
    });
    }
    
    console.log('Loaded Manager project filters:', projects.length, 'projects');
}

function onManagerProjectChange() {
    const projectFilter = document.getElementById('managerProjectFilter');
    const userFilterGroup = document.getElementById('managerUserFilterGroup');
    const userFilter = document.getElementById('managerUserFilter');
    
    if (!projectFilter || !userFilterGroup || !userFilter) return;
    
    if (projectFilter.value) {
        // Show user dropdown and populate with users from selected project
        userFilterGroup.style.display = 'block';
        
        // Filter production data by selected project
        const projectUsers = [...new Set(
            productionData
                .filter(d => d.clientName === projectFilter.value)
                .map(d => ({
                    employeeId: d.employeeId,
                    name: d.name
                }))
                .filter(u => u.employeeId && u.name)
        )];
        
        // Remove duplicates by employeeId
        const uniqueUsers = [];
        const seenIds = new Set();
        projectUsers.forEach(u => {
            if (!seenIds.has(u.employeeId)) {
                seenIds.add(u.employeeId);
                uniqueUsers.push(u);
            }
        });
        
        // Sort by name
        uniqueUsers.sort((a, b) => a.name.localeCompare(b.name));
        
        // Store current selection
        const currentUserValue = userFilter.value;
        
        userFilter.innerHTML = '<option value="">All Users</option>';
        uniqueUsers.forEach(user => {
            const option = document.createElement('option');
            option.value = user.employeeId;
            option.textContent = `${user.name} (${user.employeeId})`;
            userFilter.appendChild(option);
        });
        
        // Restore selection if it still exists in the new list
        if (currentUserValue && uniqueUsers.some(u => u.employeeId === currentUserValue)) {
            userFilter.value = currentUserValue;
        } else {
            userFilter.value = '';
        }
        
        console.log('Loaded users for project:', projectFilter.value, uniqueUsers.length, 'users');
    } else {
        // Hide user dropdown when no project is selected
        userFilterGroup.style.display = 'none';
        userFilter.value = '';
    }
    
    // Reload data with new filters
    loadManagerData();
}

function loadManagerData() {
    console.log('Loading manager data...');
    
    // Check if we're on the manager page
    const managerPage = document.getElementById('managerPage');
    if (!managerPage || managerPage.style.display === 'none') {
        console.log('Manager page not active, skipping chart creation');
        return;
    }
    
    // Load production data from localStorage (same as TL portal)
    const savedData = localStorage.getItem('dm2_production_data');
    if (savedData) {
        productionData = JSON.parse(savedData);
        console.log('Manager loaded production data:', productionData.length, 'records');
    }
    
    // Fallback to sample data if none exists
    if (!Array.isArray(productionData) || productionData.length === 0) {
        productionData = [...sampleProductionData];
        console.log('Manager using sample data:', productionData.length, 'records');
    }
    
    // Debug: Log sample data structure
    if (productionData.length > 0) {
        console.log('Manager sample data record:', productionData[0]);
        console.log('Manager date field sample:', productionData[0].date);
        console.log('Manager date field type:', typeof productionData[0].date);
        
        // Check a few more records to see the pattern
        for (let i = 0; i < Math.min(5, productionData.length); i++) {
            console.log(`Manager record ${i} date:`, productionData[i].date, 'type:', typeof productionData[i].date);
        }
    }
    
    // Apply project filter if selected
    let filteredData = [...productionData];
    const projectFilter = document.getElementById('managerProjectFilter');
    if (projectFilter && projectFilter.value) {
        filteredData = filteredData.filter(d => d.clientName === projectFilter.value);
        console.log('Manager filtered by project:', projectFilter.value, 'Records:', filteredData.length);
    }
    
    // Apply user filter if selected
    const userFilter = document.getElementById('managerUserFilter');
    if (userFilter && userFilter.value) {
        filteredData = filteredData.filter(d => d.employeeId === userFilter.value);
        console.log('Manager filtered by user:', userFilter.value, 'Records:', filteredData.length);
    }
    
    // Apply date filter if selected (for Manager portal)
    filteredData = applyManagerDateFilter(filteredData);
    console.log('Manager after date filtering:', filteredData.length, 'records');
    
    // Create charts with filtered data
    createManagerProjectChart(filteredData);
    createManagerTeamChart(filteredData);
    createManagerQualityChart(filteredData);
    
    // Calculate metrics for utilisation and stack ranking charts
    const metrics = calculateMetrics(filteredData);
    createManagerUtilisationChart(metrics);
    createManagerStackRankingChart(metrics);
    
    // Create manager ID cards with top performers
    createManagerIDCards(filteredData);
    
    console.log('Manager data loading complete');
}

function createManagerProjectChart(dataToProcess = null) {
  const container = document.getElementById('managerProjectChart');
  if (!container) {
    console.error('Manager project chart container not found');
    return;
  }

  // Clear any existing messages
  const existingMessage = container.querySelector('div[style*="position:absolute"]');
  if (existingMessage) {
    existingMessage.remove();
  }

  // Self-heal: (Re)create canvas if missing
  let canvas = container.querySelector('#managerProjectCanvas');
  if (!canvas) {
    console.log('Canvas missing, recreating...');
    container.innerHTML = '<canvas id="managerProjectCanvas"></canvas>';
    canvas = container.querySelector('#managerProjectCanvas');
  }
  
  console.log('Manager project chart - container:', container, 'canvas:', canvas);

  if (typeof Chart === 'undefined') {
    container.innerHTML = '<p style="color:red;padding:20px;">Chart.js not loaded.</p>';
    return;
  }

  // Destroy previous chart safely without removing canvas
  if (window.managerProjectChart?.destroy) {
    try { 
      window.managerProjectChart.destroy(); 
    } catch(e) { 
      console.warn('project destroy failed', e); 
    }
  }

  // Use provided data (already filtered by date and project)
  const filteredData = dataToProcess || productionData;
  
  console.log('Manager project chart - using filtered data:', filteredData.length, 'records');

  // Build project map with normalized keys (project only, no process)
  const projectMap = new Map();
  filteredData.forEach(d => {
    const pKey = (d.clientName || '').trim().toLowerCase();
    const displayName = (d.clientName || '').trim();
    
    if (!pKey) return; // Skip if no client name
    
    if (!projectMap.has(pKey)) {
      projectMap.set(pKey, {
        displayName: displayName,
        rows: [],
        members: new Set()
      });
    }
    
    const project = projectMap.get(pKey);
    project.rows.push(d);
    if (d.employeeId) {
      project.members.add(String(d.employeeId).trim());
    }
  });

  if (projectMap.size === 0) {
    // Show message without removing canvas
    const messageDiv = document.createElement('div');
    messageDiv.className = 'no-data-msg';
    messageDiv.style.cssText = 'color:#666;padding:20px;text-align:center;position:absolute;top:50%;left:50%;transform:translate(-50%,-50%);z-index:10;';
    messageDiv.textContent = 'No project data available';
    container.appendChild(messageDiv);
    return;
  }

  // Remove any existing "no data" message
  container.querySelector('.no-data-msg')?.remove();

  // Calculate stats for all projects with proper percentage calculation
  const projectStats = Array.from(projectMap.entries()).map(([pKey, { displayName, rows, members }]) => {
    const memberCount = members.size;
    const totalActual = rows.reduce((s, r) => s + (Number(r.productivity ?? r.actual) || 0), 0);
    const totalTarget = rows.reduce((s, r) => s + (Number(r.target) || 0), 0);
    
    // Ensure proper percentage calculation with distinct values
    let performance = 0;
    if (totalTarget > 0) {
      performance = Math.round((totalActual / totalTarget) * 100);
      // Ensure minimum 1% difference between projects to avoid duplicates
      performance = Math.max(1, performance);
    }
    
    return { 
      name: displayName, 
      count: memberCount, 
      totalActual, 
      totalTarget, 
      performance,
      rawPerformance: totalTarget > 0 ? (totalActual / totalTarget) * 100 : 0
    };
  });
  
  // Sort by performance descending, then by actual count, then by name
  projectStats.sort((a, b) => {
    if (b.performance !== a.performance) {
      return b.performance - a.performance;
    }
    if (b.totalActual !== a.totalActual) {
      return b.totalActual - a.totalActual;
    }
    return a.name.localeCompare(b.name);
  });
  
  // Ensure distinct percentages starting from 100% down
  if (projectStats.length > 1) {
    const maxPerformance = Math.max(...projectStats.map(p => p.rawPerformance));
    if (maxPerformance > 0) {
      projectStats.forEach((project, index) => {
        if (index === 0) {
          project.performance = 100; // Best project gets 100%
        } else {
          // Calculate relative performance compared to the best
          const relativePerformance = (project.rawPerformance / maxPerformance) * 100;
          project.performance = Math.max(1, Math.round(relativePerformance));
        }
      });
    }
  }

  // Different color schemes for different contexts
  const projectFilter = document.getElementById('managerProjectFilter');
  const isProjectSelected = projectFilter && projectFilter.value;
  let colors;
  if (isProjectSelected) {
    // When specific project selected, use distinct colors
    colors = ['#e53e3e', '#38a169', '#3182ce', '#d69e2e', '#805ad5', '#dd6b20', '#319795', '#e91e63'];
  } else {
    // When showing all projects, use original colors
    colors = ['#667eea','#764ba2','#f093fb','#4ecdc4','#45b7d1','#96ceb4','#feca57','#ff9ff3'];
  }

  const ctx2d = canvas.getContext('2d');
  window.managerProjectChart = new Chart(ctx2d, {
    type: 'doughnut',
    data: {
      labels: projectStats.map(p => `${p.name} (${p.count} users)`),
      datasets: [{
        data: projectStats.map(p => p.performance),
        backgroundColor: colors.slice(0, projectStats.length),
        borderWidth: 2,
        borderColor: '#fff'
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        title: { 
          display: true, 
          text: `All Projects Performance Overview${dataToProcess ? ' (Filtered)' : ' (All Data)'}`, 
          font: { size: 16, weight: 'bold' } 
        },
        legend: { position: 'bottom', labels: { padding: 20, usePointStyle: true } },
        tooltip: {
          callbacks: {
            label: (context) => {
              const p = projectStats[context.dataIndex];
              // Determine current period label for tooltip context
              const dateFilter = document.getElementById('managerDateFilter');
              let periodLabel = 'Overall';
              if (dateFilter && dateFilter.value === 'month') {
                const m = document.getElementById('managerSelectedMonth');
                if (m && m.value) periodLabel = m.value;
              } else if (dateFilter && dateFilter.value === 'year') {
                const y = document.getElementById('managerSelectedYear');
                if (y && y.value) periodLabel = y.value;
              } else if (dateFilter && dateFilter.value === 'date') {
                const s = document.getElementById('managerStartDate');
                const e = document.getElementById('managerEndDate');
                if (s && s.value && e && e.value) periodLabel = `${s.value} → ${e.value}`;
              }
              return [
                `1. Project name: ${p.name} (${p.count} users)`,
                `2. Performance percentage: ${p.performance}% (${periodLabel})`,
                `3. Actual counts: ${p.totalActual.toLocaleString()}`
              ];
            }
          }
        }
      }
    }
  });
}

function createManagerTeamChart(dataToProcess = null) {
  const container = document.getElementById('managerTeamChart');
  const canvas = document.getElementById('managerTeamCanvas');
  if (!container || !canvas) return;

  if (typeof Chart === 'undefined') {
    container.innerHTML = '<p style="color:red;padding:20px;">Chart.js not loaded.</p>';
    return;
  }

  if (window.managerTeamChart && typeof window.managerTeamChart.destroy === 'function') {
    try { window.managerTeamChart.destroy(); } catch (e) {}
  }

  // Use provided data or fallback to global productionData
  const data = dataToProcess || productionData;

  // Check if a specific project is selected
  const projectFilter = document.getElementById('managerProjectFilter');
  const isProjectSelected = projectFilter && projectFilter.value;
  
  // Check if a specific user is selected
  const userFilter = document.getElementById('managerUserFilter');
  const isUserSelected = userFilter && userFilter.value;

  // 1) Group rows by project key (project only, not client-process)
  const projectMap = new Map();
  (data || []).forEach(r => {
    const k = (r.clientName || '').trim().toLowerCase();
    if (!k) return; // Skip if no client name
    
    if (!projectMap.has(k)) {
      projectMap.set(k, { 
        displayName: (r.clientName || '').trim(), 
        rows: [] 
      });
    }
    projectMap.get(k).rows.push(r);
  });

  // 2) Compute stats with new formula: (actual count percentage * 50%) + (errors percentage * 50%)
  const teamStats = Array.from(projectMap.values()).map(g => {
    const totalActual = g.rows.reduce((s, r) => s + (Number(r.productivity ?? r.actual) || 0), 0);
    const totalTarget = g.rows.reduce((s, r) => s + (Number(r.target) || 0), 0);
    const totalClientErrors = g.rows.reduce((s, r) => s + (Number(r.clientErrors) || 0), 0);
    const totalInternalErrors = g.rows.reduce((s, r) => s + (Number(r.internalErrors) || 0), 0);
    const totalErrors = totalClientErrors + totalInternalErrors;
    const members = new Set(g.rows.map(r => String(r.employeeId || '').trim()).filter(Boolean));
    const memberCount = members.size;
    
    // Calculate actual count percentage
    const actualCountPercentage = totalTarget > 0 ? (totalActual / totalTarget) * 100 : 0;
    
    // Calculate errors percentage (inverse - lower errors = higher percentage)
    const maxPossibleErrors = totalActual * 0.1; // Assume max 10% error rate
    const errorsPercentage = maxPossibleErrors > 0 ? Math.max(0, 100 - (totalErrors / maxPossibleErrors) * 100) : 100;
    
    // New formula: (actual count percentage * 50%) + (errors percentage * 50%)
    const performance = Math.round((actualCountPercentage * 0.5) + (errorsPercentage * 0.5));
    
    return { 
      name: g.displayName, 
      memberCount, 
      totalActual, 
      totalTarget, 
      totalErrors,
      performance 
    };
  });

  // 3) Filter logic:
  // - If a project is selected and has < 3 members, still show it (fix the bug)
  // - If no project selected, only show projects with >= 3 members
  // - If a user is selected, show only that user's project (if it exists)
  let filteredTeamStats = teamStats;
  
  if (isUserSelected) {
    // When user is selected, show only that user's project
    filteredTeamStats = teamStats.filter(t => t.memberCount > 0);
  } else if (isProjectSelected) {
    // When project is selected, show it even if it has < 3 members
    filteredTeamStats = teamStats;
  } else {
    // When no project selected, only show projects with >= 3 members
    filteredTeamStats = teamStats.filter(t => t.memberCount >= 3);
  }
  
  filteredTeamStats.sort((a, b) => b.performance - a.performance || b.totalActual - a.totalActual || a.name.localeCompare(b.name));

  if (filteredTeamStats.length === 0) {
    container.innerHTML = '<p style="color:#666;padding:20px;text-align:center;">No data available</p>';
    return;
  }

  const ctx2d = canvas.getContext('2d');
  window.managerTeamChart = new Chart(ctx2d, {
    type: 'bar',
    data: {
      labels: filteredTeamStats.map(t => `${t.name} (${t.memberCount} users)`),
      datasets: [{
        label: 'Team Performance %',
        data: filteredTeamStats.map(t => t.performance),
        backgroundColor: filteredTeamStats.map((_, i) => {
          // Different colors when specific project is selected
          const projectFilter = document.getElementById('managerProjectFilter');
          const isProjectSelected = projectFilter && projectFilter.value;
          
          if (isProjectSelected) {
            const distinctColors = ['#e53e3e', '#38a169', '#3182ce', '#d69e2e', '#805ad5', '#dd6b20', '#319795', '#e91e63'];
            return distinctColors[i % distinctColors.length];
          } else {
            const originalColors = ['#667eea','#764ba2','#f093fb','#4ecdc4','#45b7d1'];
            return originalColors[i % originalColors.length];
          }
        }),
        borderWidth: 2,
        borderColor: '#fff'
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      scales: {
        y: { beginAtZero: true, max: 150, title: { display: true, text: 'Performance %' }, grid: { color: 'rgba(0,0,0,0.1)' } },
        x: { grid: { display: false } }
      },
      plugins: {
        title: { 
          display: true, 
          text: isProjectSelected ? `Team Performance - ${projectFilter.value}` : 'Team Performance Rankings (≥3 members)', 
          font: { size: 16, weight: 'bold' } 
        },
        legend: { display: false },
        tooltip: {
          callbacks: {
            label: (ctx) => {
              const t = filteredTeamStats[ctx.dataIndex];
              return [
                `1. Project name: ${t.name} (${t.memberCount} users)`,
                `2. Performance percentage: ${t.performance}%`,
                `3. Actual counts: ${t.totalActual.toLocaleString()}`,
                `4. Target counts: ${t.totalTarget.toLocaleString()}`
              ];
            }
          }
        }
      }
    }
  });
}

function createManagerQualityChart(dataToProcess = null) {
    const container = document.getElementById('managerQualityChart');
    const canvas = document.getElementById('managerQualityCanvas');
    if (!container || !canvas) return;

    if (typeof Chart === 'undefined') {
        container.innerHTML = '<p style="color:red;padding:20px;">Chart.js not loaded.</p>';
        return;
    }

    if (window.managerQualityChart && typeof window.managerQualityChart.destroy === 'function') {
        try { window.managerQualityChart.destroy(); } catch (e) {}
    }

    const data = dataToProcess || productionData;
    const projectFilter = document.getElementById('managerProjectFilter');
    const userFilter = document.getElementById('managerUserFilter');
    const isProjectSelected = projectFilter && projectFilter.value;
    const isUserSelected = userFilter && userFilter.value;

    // Group data by project or user depending on filter
    let qualityData = [];
    
    if (isUserSelected) {
        // Group by user
        const userMap = new Map();
        data.forEach(r => {
            const key = String(r.employeeId || '').trim();
            if (!key) return;
            
            if (!userMap.has(key)) {
                userMap.set(key, {
                    name: r.name || 'Unknown',
                    employeeId: r.employeeId,
                    clientErrors: 0,
                    internalErrors: 0,
                    totalRecords: 0
                });
            }
            const user = userMap.get(key);
            user.clientErrors += Number(r.clientErrors || 0);
            user.internalErrors += Number(r.internalErrors || 0);
            user.totalRecords += 1;
        });
        
        qualityData = Array.from(userMap.values()).map(u => ({
            label: `${u.name} (${u.employeeId})`,
            clientErrors: u.clientErrors,
            internalErrors: u.internalErrors,
            errorPercentage: u.totalRecords > 0 ? ((u.clientErrors + u.internalErrors) / u.totalRecords * 100).toFixed(2) : 0
        }));
    } else if (isProjectSelected) {
        // Group by project (for selected project)
        const projectMap = new Map();
        data.forEach(r => {
            const key = (r.clientName || '').trim();
            if (!key) return;
            
            if (!projectMap.has(key)) {
                projectMap.set(key, {
                    name: key,
                    clientErrors: 0,
                    internalErrors: 0,
                    totalRecords: 0
                });
            }
            const project = projectMap.get(key);
            project.clientErrors += Number(r.clientErrors || 0);
            project.internalErrors += Number(r.internalErrors || 0);
            project.totalRecords += 1;
        });
        
        qualityData = Array.from(projectMap.values()).map(p => ({
            label: p.name,
            clientErrors: p.clientErrors,
            internalErrors: p.internalErrors,
            errorPercentage: p.totalRecords > 0 ? ((p.clientErrors + p.internalErrors) / p.totalRecords * 100).toFixed(2) : 0
        }));
    } else {
        // Group by all projects
        const projectMap = new Map();
        data.forEach(r => {
            const key = (r.clientName || '').trim();
            if (!key) return;
            
            if (!projectMap.has(key)) {
                projectMap.set(key, {
                    name: key,
                    clientErrors: 0,
                    internalErrors: 0,
                    totalRecords: 0
                });
            }
            const project = projectMap.get(key);
            project.clientErrors += Number(r.clientErrors || 0);
            project.internalErrors += Number(r.internalErrors || 0);
            project.totalRecords += 1;
        });
        
        qualityData = Array.from(projectMap.values())
            .map(p => ({
                label: p.name,
                clientErrors: p.clientErrors,
                internalErrors: p.internalErrors,
                errorPercentage: p.totalRecords > 0 ? ((p.clientErrors + p.internalErrors) / p.totalRecords * 100).toFixed(2) : 0
            }))
            .sort((a, b) => parseFloat(b.errorPercentage) - parseFloat(a.errorPercentage))
            .slice(0, 10); // Show top 10
    }

    if (qualityData.length === 0) {
        container.innerHTML = '<p style="color:#666;padding:20px;text-align:center;">No quality data available</p>';
        return;
    }

    // Sort by error percentage descending
    qualityData.sort((a, b) => parseFloat(b.errorPercentage) - parseFloat(a.errorPercentage));

    const ctx2d = canvas.getContext('2d');
    window.managerQualityChart = new Chart(ctx2d, {
        type: 'bar',
        data: {
            labels: qualityData.map(d => d.label),
            datasets: [
                {
                    label: 'Client Errors',
                    data: qualityData.map(d => d.clientErrors),
                    backgroundColor: 'rgba(220, 53, 69, 0.7)',
                    borderColor: 'rgba(220, 53, 69, 1)',
                    borderWidth: 2
                },
                {
                    label: 'Internal Errors',
                    data: qualityData.map(d => d.internalErrors),
                    backgroundColor: 'rgba(255, 193, 7, 0.7)',
                    borderColor: 'rgba(255, 193, 7, 1)',
                    borderWidth: 2
                },
                {
                    label: 'Error Percentage',
                    data: qualityData.map(d => parseFloat(d.errorPercentage)),
                    type: 'line',
                    yAxisID: 'y1',
                    borderColor: 'rgba(102, 126, 234, 1)',
                    backgroundColor: 'rgba(102, 126, 234, 0.1)',
                    borderWidth: 3,
                    fill: false,
                    pointRadius: 5,
                    pointHoverRadius: 7
                }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            scales: {
                y: {
                    beginAtZero: true,
                    title: {
                        display: true,
                        text: 'Number of Errors'
                    }
                },
                y1: {
                    type: 'linear',
                    position: 'right',
                    beginAtZero: true,
                    title: {
                        display: true,
                        text: 'Error Percentage (%)'
                    },
                    grid: {
                        drawOnChartArea: false
                    }
                }
            },
            plugins: {
                title: {
                    display: true,
                    text: isUserSelected ? 'Quality Metrics - User View' : 
                          isProjectSelected ? `Quality Metrics - ${projectFilter.value}` : 
                          'Quality Metrics by Project',
                    font: { size: 16, weight: 'bold' }
                },
                legend: {
                    display: true,
                    position: 'top'
                },
                tooltip: {
                    callbacks: {
                        label: (context) => {
                            const index = context.dataIndex;
                            const data = qualityData[index];
                            if (context.datasetIndex === 2) {
                                return `Error Percentage: ${data.errorPercentage}%`;
                            }
                            return context.dataset.label + ': ' + context.parsed.y;
                        }
                    }
                }
            }
        }
    });
}

function createManagerUtilisationChart(metrics) {
    const ctx = document.getElementById('managerUtilisationChart');
    if (!ctx) {
        console.error('Manager utilisation chart canvas not found');
        return;
    }
    
    if (typeof Chart === 'undefined') {
        ctx.innerHTML = '<p style="color: red; padding: 20px;">Chart.js library not loaded.</p>';
        return;
    }
    
    safeDestroyChart('managerUtilisationChart');
    
    if (!metrics || metrics.length === 0) {
        ctx.innerHTML = '<p style="color: #666; padding: 20px; text-align: center;">No data available for utilisation chart</p>';
        return;
    }
    
    // Remove duplicates and sort by utilisation descending
    const uniqueEmployees = [];
    const seenIds = new Set();
    
    metrics.forEach(employee => {
        if (!seenIds.has(employee.employeeId)) {
            seenIds.add(employee.employeeId);
            uniqueEmployees.push(employee);
        }
    });
    
    // Sort by utilisation descending (best performers first)
    uniqueEmployees.sort((a, b) => b.utilisation - a.utilisation);
    
    try {
        window.managerUtilisationChart = new Chart(ctx, {
            type: 'bar',
            data: {
                labels: uniqueEmployees.map(m => `${m.name} (${m.clientName})`),
                datasets: [{
                    label: 'Utilisation %',
                    data: uniqueEmployees.map(m => m.utilisation),
                    backgroundColor: uniqueEmployees.map(m => 
                        m.utilisation >= 100 ? '#28a745' : 
                        m.utilisation >= 80 ? '#ffc107' : '#dc3545'
                    ),
                    borderColor: uniqueEmployees.map(m => 
                        m.utilisation >= 100 ? '#1e7e34' : 
                        m.utilisation >= 80 ? '#e0a800' : '#c82333'
                    ),
                    borderWidth: 2
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                scales: {
                    y: {
                        beginAtZero: true,
                        title: {
                            display: true,
                            text: 'Utilisation %'
                        }
                    },
                    x: {
                        title: {
                            display: true,
                            text: 'Employees (Ranked)'
                        }
                    }
                },
                plugins: {
                    legend: {
                        display: false
                    },
                    title: {
                        display: true,
                        text: 'Employee Utilisation Analysis (Productivity/Per Hour Count/Hours Worked * 100)'
                    },
                    tooltip: {
                        callbacks: {
                            title: ctx => `${uniqueEmployees[ctx[0].dataIndex].name} (${uniqueEmployees[ctx[0].dataIndex].clientName})`
                        }
                    }
                }
            }
        });
    } catch (error) {
        console.error('Error creating manager utilisation chart:', error);
        ctx.innerHTML = '<p style="color: red; padding: 20px;">Error creating chart. Please try again.</p>';
    }
}

function createManagerStackRankingChart(metrics) {
    const ctx = document.getElementById('managerStackRankingChart');
    if (!ctx) {
        console.error('Manager stack ranking chart canvas not found');
        return;
    }
    
    if (typeof Chart === 'undefined') {
        ctx.innerHTML = '<p style="color: red; padding: 20px;">Chart.js library not loaded.</p>';
        return;
    }
    
    safeDestroyChart('managerStackRankingChart');
    
    if (!metrics || metrics.length === 0) {
        ctx.innerHTML = '<p style="color: #666; padding: 20px; text-align: center;">No data available for stack ranking chart</p>';
        return;
    }
    
    // Remove duplicates and sort by stack ranking points descending
    const uniqueEmployees = [];
    const seenIds = new Set();
    
    metrics.forEach(employee => {
        if (!seenIds.has(employee.employeeId)) {
            seenIds.add(employee.employeeId);
            uniqueEmployees.push(employee);
        }
    });
    
    // Sort by stack ranking points descending (best performers first)
    uniqueEmployees.sort((a, b) => b.stackRankingPoints - a.stackRankingPoints);
    
    try {
        window.managerStackRankingChart = new Chart(ctx, {
            type: 'bar',
            data: {
                labels: uniqueEmployees.map(m => `${m.name} (${m.clientName})`),
                datasets: [{
                    label: 'Stack Ranking Points',
                    data: uniqueEmployees.map(m => m.stackRankingPoints),
                    backgroundColor: uniqueEmployees.map((m, index) => {
                        if (index === 0) return '#ffd700'; // Gold for 1st
                        if (index === 1) return '#c0c0c0'; // Silver for 2nd
                        if (index === 2) return '#cd7f32'; // Bronze for 3rd
                        return '#667eea';
                    }),
                    borderColor: uniqueEmployees.map((m, index) => {
                        if (index === 0) return '#ffb300'; // Gold border
                        if (index === 1) return '#a0a0a0'; // Silver border
                        if (index === 2) return '#b8860b'; // Bronze border
                        return '#5a6fd8';
                    }),
                    borderWidth: 2
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                scales: {
                    y: {
                        beginAtZero: true,
                        max: 12,
                        title: {
                            display: true,
                            text: 'Stack Ranking Points'
                        }
                    },
                    x: {
                        title: {
                            display: true,
                            text: 'Employees (Ranked)'
                        }
                    }
                },
                plugins: {
                    legend: {
                        display: false
                    },
                    title: {
                        display: true,
                        text: 'Stack Ranking (Target Achievement + Errors + Working Days Points)'
                    },
                    tooltip: {
                        callbacks: {
                            title: ctx => `${uniqueEmployees[ctx[0].dataIndex].name} (${uniqueEmployees[ctx[0].dataIndex].clientName})`
                        }
                    }
                }
            }
        });
    } catch (error) {
        console.error('Error creating manager stack ranking chart:', error);
        ctx.innerHTML = '<p style="color: red; padding: 20px;">Error creating chart. Please try again.</p>';
    }
}

function showManagerTab(tabName, ev) {
    // Hide all manager tabs
    document.querySelectorAll('#managerUtilisationTab, #managerStackRankingTab').forEach(tab => {
        tab.classList.remove('active');
    });
    document.querySelectorAll('.manager-visualization-section .tab-btn').forEach(btn => {
        btn.classList.remove('active');
    });
    
    // Show selected tab
    const tabId = 'manager' + tabName.charAt(0).toUpperCase() + tabName.slice(1) + 'Tab';
    const tabElement = document.getElementById(tabId);
    if (tabElement) {
        tabElement.classList.add('active');
    }
    if (ev && ev.target) ev.target.classList.add('active');
}

function createManagerIDCards(dataToProcess = null) {
    // Use provided data or fallback to global productionData
    const data = dataToProcess || productionData;
    const metrics = calculateMetrics(data);
    const container = document.getElementById('managerTopPerformers');
    if (!container) {
        console.error('Manager top performers container not found');
        return;
    }
    container.innerHTML = '';
    
    if (!metrics || metrics.length === 0) {
        container.innerHTML = '<p style="text-align: center; color: #666; padding: 20px;">No performance data available.</p>';
        return;
    }
    
    // Check current filters
    const projectFilter = document.getElementById('managerProjectFilter');
    const userFilter = document.getElementById('managerUserFilter');
    const isProjectSelected = projectFilter && projectFilter.value;
    const isUserSelected = userFilter && userFilter.value;
    
    // Remove duplicates by employeeId and aggregate metrics
    const employeeMap = new Map();
    metrics.forEach(employee => {
        const empId = employee.employeeId;
        if (!empId) return;
        
        if (!employeeMap.has(empId)) {
            employeeMap.set(empId, {
                employeeId: empId,
                name: employee.name,
                clientName: employee.clientName,
                target: 0,
                productivity: 0,
                clientErrors: 0,
                internalErrors: 0,
                stackRankingPoints: employee.stackRankingPoints || 0,
                utilisation: employee.utilisation || 0,
                teamPerformance: employee.teamPerformance || 0,
                hoursWorked: employee.hoursWorked || 8
            });
        }
        
        const emp = employeeMap.get(empId);
        // Aggregate values
        emp.target += (employee.target || 0);
        emp.productivity += (employee.productivity || 0);
        emp.clientErrors += (employee.clientErrors || 0);
        emp.internalErrors += (employee.internalErrors || 0);
        
        // Use the highest values for ranking metrics
        if (employee.stackRankingPoints > emp.stackRankingPoints) {
            emp.stackRankingPoints = employee.stackRankingPoints;
        }
        if (employee.utilisation > emp.utilisation) {
            emp.utilisation = employee.utilisation;
        }
        if (employee.teamPerformance > emp.teamPerformance) {
            emp.teamPerformance = employee.teamPerformance;
        }
    });
    
    let topPerformers = [];
    
    if (!isProjectSelected && !isUserSelected) {
        // OVERALL VIEW: pick one top performer per project, then take top 6
        const projectToEmployees = new Map();
        Array.from(employeeMap.values()).forEach(emp => {
            const key = (emp.clientName || '').trim().toLowerCase();
            if (!key) return;
            if (!projectToEmployees.has(key)) projectToEmployees.set(key, []);
            projectToEmployees.get(key).push(emp);
        });
        
        // Sort inside each project using strict comparator
        const comparator = (a, b) => (
            b.stackRankingPoints - a.stackRankingPoints ||
            b.utilisation - a.utilisation ||
            b.teamPerformance - a.teamPerformance ||
            b.productivity - a.productivity
        );
        
        const topPerProject = [];
        projectToEmployees.forEach(arr => {
            arr.sort(comparator);
            if (arr.length) topPerProject.push(arr[0]);
        });
        
        // Now take global top 6 by the same comparator
        topPerProject.sort(comparator);
        topPerformers = topPerProject.slice(0, 6);
    } else {
        // PROJECT VIEW (or user selected): show users from that project only
        let projectUsers = Array.from(employeeMap.values());
        if (isUserSelected) {
            projectUsers = projectUsers.filter(e => e.employeeId === userFilter.value);
        } else if (isProjectSelected) {
            projectUsers = projectUsers.filter(e => e.clientName === projectFilter.value);
        }
        
        projectUsers.sort((a, b) => (
            b.stackRankingPoints - a.stackRankingPoints ||
            b.utilisation - a.utilisation ||
            b.teamPerformance - a.teamPerformance ||
            b.productivity - a.productivity
        ));
        
        const maxCards = projectUsers.length > 6 ? 6 : projectUsers.length;
        topPerformers = projectUsers.slice(0, maxCards);
    }
    
    console.log('Manager Final top performers:', topPerformers.length);
    
    if (topPerformers.length === 0) {
        container.innerHTML = '<p style="text-align: center; color: #666; padding: 20px;">No performance data available for the selected filters.</p>';
        return;
    }
    
    topPerformers.forEach((user, index) => {
        const card = document.createElement('div');
        card.className = 'user-card';
        card.innerHTML = `
            <h4>${index === 0 ? '🥇' : index === 1 ? '🥈' : index === 2 ? '🥉' : '🏆'} ${user.name}</h4>
            <div class="user-info">
                <div class="info-row">
                    <span class="info-label">Employee ID:</span>
                    <span class="info-value">${user.employeeId || 'N/A'}</span>
                </div>
                <div class="info-row">
                    <span class="info-label">Client/Project Name:</span>
                    <span class="info-value">${user.clientName || 'N/A'}</span>
                </div>
                <div class="info-row">
                    <span class="info-label">Total Productivity:</span>
                    <span class="info-value">${user.productivity.toLocaleString()}</span>
                </div>
                <div class="info-row">
                    <span class="info-label">Client Errors:</span>
                    <span class="info-value">${user.clientErrors}</span>
                </div>
                <div class="info-row">
                    <span class="info-label">Internal Errors:</span>
                    <span class="info-value">${user.internalErrors}</span>
                </div>
                <div class="info-row">
                    <span class="info-label">Stack Ranking:</span>
                    <span class="info-value">#${index + 1} (${user.stackRankingPoints} pts)</span>
                </div>
                <div class="info-row">
                    <span class="info-label">Utilisation:</span>
                    <span class="info-value">${user.utilisation.toFixed(1)}%</span>
                </div>
                <div class="info-row">
                    <span class="info-label">Team Performance:</span>
                    <span class="info-value">${user.teamPerformance.toFixed(1)}%</span>
                </div>
            </div>
        `;
        container.appendChild(card);
    });
}

function loadManagerFeedbackMessages() {
    // Create feedback messages section if it doesn't exist
    let feedbackSection = document.getElementById('managerFeedbackMessages');
    if (!feedbackSection) {
        const managerContent = document.querySelector('.manager-content');
        feedbackSection = document.createElement('div');
        feedbackSection.id = 'managerFeedbackMessages';
        feedbackSection.className = 'feedback-section';
        feedbackSection.innerHTML = `
            <h3>Feedback Conversations</h3>
            <div id="managerFeedbackList" class="feedback-messages">
            </div>
        `;
        managerContent.appendChild(feedbackSection);
    }
    
    const container = document.getElementById('managerFeedbackList');
    container.innerHTML = '';
    
    // Get all feedback messages involving this manager
    const managerFeedback = feedbackMessages.filter(f => 
        f.from === currentUser.email || f.to === currentUser.email
    );
    
    if (managerFeedback.length === 0) {
        container.innerHTML = '<p>No feedback conversations yet.</p>';
        return;
    }
    
    // Group by conversation thread
    const conversations = {};
    managerFeedback.forEach(feedback => {
        const threadId = feedback.parentId || feedback.id;
        if (!conversations[threadId]) {
            conversations[threadId] = [];
        }
        conversations[threadId].push(feedback);
    });
    
    // Display full conversations with all messages
    Object.values(conversations).forEach(conversation => {
        conversation.sort((a, b) => new Date(a.timestamp) - new Date(b.timestamp));
        
        const conversationDiv = document.createElement('div');
        conversationDiv.className = 'conversation-thread';
        
        const otherUser = conversation[0].from === currentUser.email ? conversation[0].to : conversation[0].from;
        
        let conversationHTML = '<div class="conversation-header">';
        conversationHTML += `<h5>Conversation with: ${otherUser}</h5>`;
        conversationHTML += '</div>';
        
        // Show all messages in the conversation
        conversation.forEach(msg => {
            const isFromCurrentUser = msg.from === currentUser.email;
            conversationHTML += `
                <div class="message ${isFromCurrentUser ? 'message-sent' : 'message-received'}">
                    <div class="message-header">
                        <strong>${isFromCurrentUser ? 'You' : msg.from}</strong>
                        <small>${new Date(msg.timestamp).toLocaleString()}</small>
                    </div>
                    <div class="message-content">${msg.message}</div>
                </div>
            `;
        });
        
        conversationDiv.innerHTML = conversationHTML;
        container.appendChild(conversationDiv);
    });
}

function handleManagerFeedback(e) {
    e.preventDefault();
    
    const toTL = document.getElementById('feedbackTL').value;
    const message = document.getElementById('feedbackMessage').value;
    
    const feedback = {
        id: Date.now(),
        from: currentUser.email,
        to: toTL,
        message: message,
        timestamp: new Date().toISOString(),
        read: false
    };
    
    feedbackMessages.push(feedback);
    localStorage.setItem('dm2_feedback', JSON.stringify(feedbackMessages));
    
    document.getElementById('managerFeedbackForm').reset();
    alert('Feedback sent successfully!');
    loadManagerFeedbackMessages();
}

// TL & Coordinator Portal Functions
function loadTLPage() {
    if (currentUser && currentUser.role === 'tl') {
        document.getElementById('tlUserEmail').textContent = currentUser.email;
        
        // Always load the latest data from localStorage
        const savedData = localStorage.getItem('dm2_production_data');
        if (savedData) {
            productionData = JSON.parse(savedData);
        }
        
        // Create sample feedback if none exists
        createSampleFeedback();
        
        // Set up Google Sheets link for this page
        setDefaultGoogleSheetsLink();
        
        // Load project filters
        loadTLProjectFilters();
        
        // Load data and show it
        loadProductionData();
        loadFeedbackMessages();
        
        // Populate saved sheet URL hint
        const savedUrl = localStorage.getItem('dm2_sheet_url');
        const hint = document.getElementById('savedSheetUrlHint');
        const input = document.getElementById('googleSheetUrl');
        if (hint) {
            hint.textContent = savedUrl ? `Saved: ${savedUrl}` : 'No sheet link saved yet.';
        }
        if (input && savedUrl) input.value = savedUrl;
        
        // Show data status
        showDataStatus();
    }
}

function loadTLProjectFilters() {
    // Load projects for TL portal
    const projectFilter = document.getElementById('projectFilter');
    if (!projectFilter) return;
    
    // Get unique projects from production data
    const projects = [...new Set(productionData.map(d => d.clientName).filter(Boolean))];
    projectFilter.innerHTML = '<option value="">All Projects</option>';
    projects.forEach(project => {
        const option = document.createElement('option');
        option.value = project;
        option.textContent = project;
        projectFilter.appendChild(option);
    });
    
    console.log('Loaded TL project filters:', projects.length, 'projects');
}

function refreshProjectFilters() {
    // Refresh project filters for both TL and Manager portals
    loadTLProjectFilters();
    loadManagerFilters();
}

function createSampleFeedback() {
    // Check if feedback already exists
    if (feedbackMessages.length > 0) return;
    
    // Create sample feedback messages
    const sampleFeedback = [
        {
            id: 1,
            from: 'manager@dm2.com',
            to: 'tl@dm2.com',
            message: 'Great work on the team performance this month! Please review the utilization metrics and provide feedback on areas for improvement.',
            timestamp: new Date(Date.now() - 2 * 24 * 60 * 60 * 1000).toISOString(), // 2 days ago
            read: false
        },
        {
            id: 2,
            from: 'manager@dm2.com',
            to: 'tl@dm2.com',
            message: 'The stack ranking analysis looks good. Can you share insights on the top performers and any training needs?',
            timestamp: new Date(Date.now() - 1 * 24 * 60 * 60 * 1000).toISOString(), // 1 day ago
            read: false
        }
    ];
    
    feedbackMessages = sampleFeedback;
    localStorage.setItem('dm2_feedback', JSON.stringify(feedbackMessages));
}

function saveSheetUrl() {
    const input = document.getElementById('googleSheetUrl') || document.getElementById('sheetLink');
    const url = (input && input.value || '').trim();
    if (!url) {
        alert('Please paste a valid Google Sheet link.');
        return;
    }
    
    // Save to both storage keys for cross-browser compatibility
    localStorage.setItem('dm2_sheet_url', url);
    localStorage.setItem('dm2_google_sheets_link', url);
    
    // Update the saved link display
    updateSavedLinkDisplay();
    
    alert('✅ Google Sheet link saved. Click Connect Google Sheets to load data.');
}

// Debug function to test Google Sheet columns
async function debugGoogleSheetColumns() {
    const userUrl = localStorage.getItem('dm2_sheet_url');
    if (!userUrl) {
        alert('Please save a Google Sheet link first.');
        return;
    }
    
    const { sheetId, gid } = parseSheetUrl(userUrl);
    if (!sheetId) {
        alert('Invalid Google Sheet URL.');
        return;
    }
    
    try {
        const { rows, cols } = await fetchGoogleSheetsDataWithFallback(sheetId, gid);
        console.log('=== GOOGLE SHEET DEBUG INFO ===');
        console.log('Sheet ID:', sheetId);
        console.log('GID:', gid);
        console.log('Total rows:', rows.length);
        console.log('Columns found:', cols);
        console.log('First few rows:', rows.slice(0, 3));
        
        alert(`Debug info logged to console. Found ${cols.length} columns: ${cols.join(', ')}`);
    } catch (error) {
        console.error('Debug failed:', error);
        alert('Debug failed: ' + error.message);
    }
}

function openGoogleSheets() {
    // Try to get URL from both possible storage keys for cross-browser compatibility
    let userUrl = localStorage.getItem('dm2_sheet_url') || localStorage.getItem('dm2_google_sheets_link');
    
    if (!userUrl) {
        // Use default URL if none is saved
        userUrl = 'https://docs.google.com/spreadsheets/d/1SUYakZKT_m4sw2xkFbuRnqX9r8_3qP-v/edit?usp=sharing&ouid=110973781089127610610&rtpof=true&sd=true';
        localStorage.setItem('dm2_google_sheets_link', userUrl);
        console.log('Using default Google Sheets URL for cross-browser compatibility');
    }
    
    console.log('Connecting to Google Sheets with URL:', userUrl);
    
    const { sheetId, gid } = parseSheetUrl(userUrl);
    
    if (!sheetId) {
        alert('❌ Invalid Google Sheet URL. Please ensure it contains the sheet ID.');
        return;
    }
    
    // Show loading
    const connectBtn = document.querySelector('button[onclick="openGoogleSheets()"]');
    if (connectBtn) {
        connectBtn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Connecting...';
        connectBtn.disabled = true;
    }
    
    // Try multiple approaches to fetch Google Sheets data
    fetchGoogleSheetsDataWithFallback(sheetId, gid)
        .then(({ rows, cols }) => {
            console.log('Columns:', cols);
            console.log('Rows:', rows.length);
            
            if (rows.length === 0) {
                throw new Error('No data found in the sheet');
            }
            
            // Map to our format
            const mapped = mapGoogleSheetRows(rows, cols);
            console.log('Mapped data:', mapped.length, 'records');
            
            if (mapped.length === 0) {
                throw new Error('Could not map any data from the sheet. Please check column headers match expected format.');
            }
            
            productionData = mapped;
            localStorage.setItem('dm2_production_data', JSON.stringify(productionData));
            loadProductionData();
            alert('✅ Successfully loaded ' + mapped.length + ' records from Google Sheets!');
            
            // Refresh project filters after data update
            refreshProjectFilters();
        })
            .catch(err => {
                console.error('Google Sheets fetch failed:', err);
                alert('❌ Failed to load Google Sheets data:\n\n' + err.message + '\n\n🔧 CRITICAL FIX REQUIRED:\n1. Open your Google Sheet\n2. Click "Share" button (top-right)\n3. Change to "Anyone with the link can view"\n4. Copy the new link and try again\n\n📋 Expected columns: Date, Employee ID, Name, Client Name, Process Name, Productivity, Target, Client Errors, Internal Errors, Hours Worked, Actual Hours\n\nUsing sample data for now.');
                loadProductionData();
            })
        .finally(() => {
            if (connectBtn) {
                connectBtn.innerHTML = '<i class="fab fa-google"></i> Connect Google Sheets';
                connectBtn.disabled = false;
            }
        });
}


// Direct function to handle Excel upload button click
function uploadExcelFile() {
    const fileInput = document.getElementById('excelUpload');
    if (fileInput) {
        fileInput.click();
    } else {
        alert('Excel upload not available. Please refresh the page.');
    }
}

function handleExcelUpload(e) {
    const file = e.target.files[0];
    if (!file) {
        console.log('No file selected');
        return;
    }
    
    console.log('Excel file selected:', file.name, file.size, 'bytes');
    
    // Check if XLSX library is loaded
    if (typeof XLSX === 'undefined') {
        alert('❌ Excel parser not loaded. Please check internet connection for SheetJS library.');
        return;
    }
    
    // Show loading indicator
    const uploadBtn = document.querySelector('button[onclick*="excelUpload"]');
    let originalText = '';
    if (uploadBtn) {
        originalText = uploadBtn.innerHTML;
        uploadBtn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Processing...';
        uploadBtn.disabled = true;
    }

    console.log('Starting Excel file processing...');
    
    parseExcelFile(file)
        .then(({ rows }) => {
            console.log('Excel parsed successfully, rows:', rows.length);
            if (!rows || !rows.length) {
                alert('❌ No rows detected in the Excel file. Please check the sheet has data.');
                return;
            }

            const mapped = mapExcelRows(rows);
            console.log('Mapped rows:', mapped.length);
            if (!mapped.length) {
                alert('❌ Could not map any rows. Please verify header names match expected format.\n\nExpected headers: ITPL#, Name, Client Name, Process Name, Productivity, Target, Client Errors, Internal Errors, Leaves, Working Days, Date');
                return;
            }
            productionData = mapped;
            localStorage.setItem('dm2_production_data', JSON.stringify(productionData));
            loadProductionData();
            alert('✅ Excel uploaded successfully! Loaded ' + mapped.length + ' employee records.');
            
            // Refresh project filters after data update
            refreshProjectFilters();
        })
        .catch((err) => {
            console.error('Excel parse failed', err);
            alert('❌ Failed to load/parse the Excel file. Please ensure:\n1. File is a valid .xlsx or .xls format\n2. File is not corrupted\n3. File is not password protected\n\nTry again with a different file.');
        })
        .finally(() => {
            // Restore button and clear file input
            if (uploadBtn && originalText) {
                uploadBtn.innerHTML = originalText;
                uploadBtn.disabled = false;
            }
            e.target.value = ''; // Clear the file input
        });
}

function parseExcelFile(file) {
    // Try ArrayBuffer first, then fall back to BinaryString
    return new Promise((resolve, reject) => {
        const attemptArray = () => {
            const reader = new FileReader();
            reader.onload = (evt) => {
                try {
                    const buffer = evt.target.result; // ArrayBuffer
                    const workbook = XLSX.read(buffer, { type: 'array' });
                    resolve(extractFirstSheetRows(workbook));
                } catch (e) {
                    console.warn('ArrayBuffer parse failed, retrying as binary', e);
                    attemptBinary();
                }
            };
            reader.onerror = reject;
            reader.readAsArrayBuffer(file);
        };

        const attemptBinary = () => {
            const reader = new FileReader();
            reader.onload = (evt) => {
                try {
                    const bstr = evt.target.result;
                    const workbook = XLSX.read(bstr, { type: 'binary' });
                    resolve(extractFirstSheetRows(workbook));
                } catch (e) {
                    reject(e);
                }
            };
            reader.onerror = reject;
            reader.readAsBinaryString(file);
        };

        attemptArray();
    });
}

function extractFirstSheetRows(workbook) {
    const sheetName = workbook.SheetNames[0];
    const ws = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(ws, { defval: '', raw: false });
    return { rows };
}

function mapExcelRows(jsonRows) {
    // Build a case-insensitive map for each row
    const normalized = jsonRows.map((row) => {
        const norm = {};
        Object.keys(row).forEach((k) => {
            norm[k] = row[k];
            norm[k.toLowerCase()] = row[k];
            norm[k.replace(/\s+/g, '').toLowerCase()] = row[k];
        });
        return norm;
    });

    const pick = (r, keys) => {
        for (const k of keys) {
            if (r[k] !== undefined && r[k] !== null && String(r[k]).trim() !== '') return r[k];
        }
        return '';
    };

    return normalized.map((r) => {
        const userId = pick(r, ['itpl#','itpl','employee id','emp id','id','employeeid','itpl#'.replace(/\s+/g,'')]);
        const name = pick(r, ['name','employee name','employeename']);
        const project = pick(r, ['client name','project','project name','clientname','projectname']);
        const processName = pick(r, ['process name','process','processname']);
        const toNum = (v) => {
            const s = (v === undefined || v === null) ? '' : String(v).toString().trim();
            if (s === '' || s === '-' || s.toLowerCase() === 'na') return 0;
            const n = Number(s);
            return isNaN(n) ? 0 : n;
        };
        const productivity = toNum(pick(r, ['productivity','actual','count','volume']));
        const target = toNum(pick(r, ['target','monthly target','monthlytarget']));
        const clientErr = toNum(pick(r, ['client errors','client error','clienterrors']));
        const internalErr = toNum(pick(r, ['internal errors','internal error','internalerrors']));
        const leaves = Number(pick(r, ['leaves','leave days','leavedays'])) || 0;
        const workingDays = Number(pick(r, ['working days','workingdays'])) || 22;
        const month = pick(r, ['date','month']);

        return {
            userId: String(userId || '').trim() || 'EMP-' + Math.random().toString(36).slice(2,7).toUpperCase(),
            name: String(name || 'Unknown'),
            project: String(project || 'NA'),
            processName: String(processName || ''),
            actual: productivity,
            target: target,
            errors: clientErr + internalErr,
            leaves: leaves,
            workingDays: workingDays,
            month: month ? String(month) : '',
            team: 'Team A'
        };
    }).filter(r => r.name && r.userId);
}

function updateDateInputs() {
    const dateFilter = document.getElementById('dateFilter').value;
    const dateInputs = document.getElementById('dateInputs');
    
    dateInputs.innerHTML = '';
    
    switch(dateFilter) {
        case 'month':
            dateInputs.innerHTML = `
                <div class="form-group">
                    <label for="selectedMonth">Month:</label>
                    <input type="month" id="selectedMonth" value="" onchange="loadProductionData()">
                </div>
            `;
            break;
        case 'date':
            dateInputs.innerHTML = `
                <div class="form-group">
                    <label for="startDate">Start Date:</label>
                    <input type="date" id="startDate" value="" onchange="loadProductionData()">
                </div>
                <div class="form-group">
                    <label for="endDate">End Date:</label>
                    <input type="date" id="endDate" value="" onchange="loadProductionData()">
                </div>
            `;
            break;
        case 'year':
            dateInputs.innerHTML = `
                <div class="form-group">
                    <label for="selectedYear">Year:</label>
                    <input type="number" id="selectedYear" value="" min="2020" max="2030" placeholder="Enter year" onchange="loadProductionData()">
                </div>
            `;
            break;
        default:
            dateInputs.innerHTML = '<p style="color: #666; font-style: italic;">Please select a period type above</p>';
            break;
    }
}

// Helper function to parse dates robustly
function parseDate(dateString) {
    if (!dateString) return null;
    
    // Handle different date formats
    let date;
    
    // Handle Date object format like "Date(2025,5,16)"
    if (typeof dateString === 'string' && dateString.startsWith('Date(')) {
        try {
            // Extract the date parts from "Date(2025,5,16)" format
            const match = dateString.match(/Date\((\d+),(\d+),(\d+)\)/);
            if (match) {
                const year = parseInt(match[1]);
                const month = parseInt(match[2]); // Already 0-based in this format
                const day = parseInt(match[3]);
                date = new Date(year, month, day);
                if (!isNaN(date.getTime())) return date;
            }
        } catch (e) {
            console.warn('Error parsing Date() format:', dateString, e);
        }
    }
    
    // Handle Date objects directly
    if (dateString instanceof Date) {
        return dateString;
    }
    
    // Try parsing as-is first
    date = new Date(dateString);
    if (!isNaN(date.getTime())) return date;
    
    // Try parsing M/D/YYYY format (like "4/21/2025")
    if (typeof dateString === 'string' && dateString.includes('/')) {
        const parts = dateString.split('/');
        if (parts.length === 3) {
            const month = parseInt(parts[0]) - 1; // JavaScript months are 0-based
            const day = parseInt(parts[1]);
            const year = parseInt(parts[2]);
            date = new Date(year, month, day);
            if (!isNaN(date.getTime())) return date;
        }
    }
    
    // Try parsing YYYY-MM-DD format
    if (typeof dateString === 'string' && dateString.includes('-')) {
        date = new Date(dateString);
        if (!isNaN(date.getTime())) return date;
    }
    
    console.warn('Could not parse date:', dateString);
    return null;
}

function applyDateFilter(data) {
    const dateFilter = document.getElementById('dateFilter');
    if (!dateFilter || !dateFilter.value) {
        return data; // No date filter selected, return all data
    }
    
    let filteredData = [...data];
    
    switch(dateFilter.value) {
        case 'month':
            const selectedMonth = document.getElementById('selectedMonth');
            if (selectedMonth && selectedMonth.value) {
                const [year, month] = selectedMonth.value.split('-');
                const targetYear = parseInt(year);
                const targetMonth = parseInt(month) - 1; // JavaScript months are 0-based
                
                filteredData = data.filter(d => {
                    if (!d.date) return false;
                    const recordDate = parseDate(d.date);
                    if (!recordDate) return false;
                    return recordDate.getFullYear() === targetYear && recordDate.getMonth() === targetMonth;
                });
                console.log('TL Filtered by month:', selectedMonth.value, 'Records:', filteredData.length, 'from', data.length);
            }
            break;
            
        case 'date':
            const startDate = document.getElementById('startDate');
            const endDate = document.getElementById('endDate');
            if (startDate && startDate.value && endDate && endDate.value) {
                const start = new Date(startDate.value);
                const end = new Date(endDate.value);
                end.setHours(23, 59, 59, 999); // Include the entire end date
                
                filteredData = data.filter(d => {
                    if (!d.date) return false;
                    const recordDate = parseDate(d.date);
                    if (!recordDate) return false;
                    return recordDate >= start && recordDate <= end;
                });
                console.log('TL Filtered by date range:', startDate.value, 'to', endDate.value, 'Records:', filteredData.length);
            }
            break;
            
        case 'year':
            const selectedYear = document.getElementById('selectedYear');
            if (selectedYear && selectedYear.value) {
                const year = parseInt(selectedYear.value);
                filteredData = data.filter(d => {
                    if (!d.date) return false;
                    const recordDate = parseDate(d.date);
                    if (!recordDate) return false;
                    return recordDate.getFullYear() === year;
                });
                console.log('TL Filtered by year:', year, 'Records:', filteredData.length);
            }
            break;
    }
    
    return filteredData;
}

function updateManagerDateInputs() {
    const dateFilter = document.getElementById('managerDateFilter').value;
    const dateInputs = document.getElementById('managerDateInputs');
    
    dateInputs.innerHTML = '';
    
    switch(dateFilter) {
        case 'month':
            dateInputs.innerHTML = `
                <div class="form-group">
                    <label for="managerSelectedMonth">Month:</label>
                    <input type="month" id="managerSelectedMonth" value="" onchange="loadManagerData()">
                </div>
            `;
            break;
        case 'date':
            dateInputs.innerHTML = `
                <div class="form-group">
                    <label for="managerStartDate">Start Date:</label>
                    <input type="date" id="managerStartDate" value="" onchange="loadManagerData()">
                </div>
                <div class="form-group">
                    <label for="managerEndDate">End Date:</label>
                    <input type="date" id="managerEndDate" value="" onchange="loadManagerData()">
                </div>
            `;
            break;
        case 'year':
            dateInputs.innerHTML = `
                <div class="form-group">
                    <label for="managerSelectedYear">Year:</label>
                    <input type="number" id="managerSelectedYear" value="" min="2020" max="2030" placeholder="Enter year" onchange="loadManagerData()">
                </div>
            `;
            break;
        default:
            dateInputs.innerHTML = '<p style="color: #666; font-style: italic;">Please select a period type above</p>';
            break;
    }
}

function applyManagerDateFilter(data) {
    const dateFilter = document.getElementById('managerDateFilter');
    if (!dateFilter || !dateFilter.value) {
        return data; // No date filter selected, return all data
    }
    
    let filteredData = [...data];
    
    switch(dateFilter.value) {
        case 'month':
            const selectedMonth = document.getElementById('managerSelectedMonth');
            if (selectedMonth && selectedMonth.value) {
                const [year, month] = selectedMonth.value.split('-');
                const targetYear = parseInt(year);
                const targetMonth = parseInt(month) - 1; // JavaScript months are 0-based
                
                filteredData = data.filter(d => {
                    if (!d.date) return false;
                    const recordDate = parseDate(d.date);
                    if (!recordDate) return false;
                    return recordDate.getFullYear() === targetYear && recordDate.getMonth() === targetMonth;
                });
                console.log('Manager Filtered by month:', selectedMonth.value, 'Records:', filteredData.length, 'from', data.length);
            }
            break;
            
        case 'date':
            const startDate = document.getElementById('managerStartDate');
            const endDate = document.getElementById('managerEndDate');
            if (startDate && startDate.value && endDate && endDate.value) {
                const start = new Date(startDate.value);
                const end = new Date(endDate.value);
                end.setHours(23, 59, 59, 999); // Include the entire end date
                
                filteredData = data.filter(d => {
                    if (!d.date) return false;
                    const recordDate = parseDate(d.date);
                    if (!recordDate) return false;
                    return recordDate >= start && recordDate <= end;
                });
                console.log('Manager Filtered by date range:', startDate.value, 'to', endDate.value, 'Records:', filteredData.length);
            }
            break;
            
        case 'year':
            const selectedYear = document.getElementById('managerSelectedYear');
            if (selectedYear && selectedYear.value) {
                const year = parseInt(selectedYear.value);
                filteredData = data.filter(d => {
                    if (!d.date) return false;
                    const recordDate = parseDate(d.date);
                    if (!recordDate) return false;
                    return recordDate.getFullYear() === year;
                });
                console.log('Manager Filtered by year:', year, 'Records:', filteredData.length);
            }
            break;
    }
    
    return filteredData;
}

function loadProductionData() {
    console.log('Loading production data...');
    
    // Check if Chart.js is loaded
    if (typeof Chart === 'undefined') {
        console.error('Chart.js library not loaded. Please check internet connection.');
        alert('Chart.js library not loaded. Please check your internet connection and refresh the page.');
        return;
    }
    
    // Try to load from Google Sheets first (if link is saved and not already loaded)
    const savedSheetLink = localStorage.getItem('dm2_google_sheets_link');
    if (savedSheetLink && savedSheetLink.includes('docs.google.com') && (!productionData || productionData.length === 0)) {
        console.log('Attempting to load data from Google Sheets:', savedSheetLink);
        // Load data from Google Sheets using the existing function
        loadDataFromGoogleSheetsAsync(savedSheetLink);
        return; // This will call loadProductionData again after loading
    }
    
    // Update the global productionData with the latest data
    const savedData = localStorage.getItem('dm2_production_data');
    if (savedData) {
        try {
        productionData = JSON.parse(savedData);
        console.log('Loaded data from localStorage:', productionData.length, 'records');
        } catch (error) {
            console.warn('Failed to parse saved data, using sample data:', error);
            productionData = [...sampleProductionData];
            localStorage.setItem('dm2_production_data', JSON.stringify(productionData));
        }
    }
    
    // Fallback if still empty
    if (!Array.isArray(productionData) || productionData.length === 0) {
        productionData = [...sampleProductionData];
        localStorage.setItem('dm2_production_data', JSON.stringify(productionData));
        console.log('Using sample data:', productionData.length, 'records');
    }
    
    // Debug: Log sample data structure
    if (productionData.length > 0) {
        console.log('Sample data record:', productionData[0]);
        console.log('Date field sample:', productionData[0].date);
        console.log('Date field type:', typeof productionData[0].date);
        
        // Check a few more records to see the pattern
        for (let i = 0; i < Math.min(5, productionData.length); i++) {
            console.log(`Record ${i} date:`, productionData[i].date, 'type:', typeof productionData[i].date);
        }
    }
    
    // Apply project filter if selected
    let filteredData = [...productionData];
    const projectFilter = document.getElementById('projectFilter');
    if (projectFilter && projectFilter.value) {
        filteredData = filteredData.filter(d => d.clientName === projectFilter.value);
        console.log('Filtered by project:', projectFilter.value, 'Records:', filteredData.length);
    }
    
    // Apply date filter if selected
    filteredData = applyDateFilter(filteredData);
    console.log('After date filtering:', filteredData.length, 'records');
    
    // Calculate metrics with the filtered data
    const metrics = calculateMetrics(filteredData);
    console.log('Calculated metrics:', metrics.length, 'records');
    
    if (!metrics.length) {
        console.warn('No metrics to display');
        return;
    }

    // Create visualizations
    console.log('Creating visualizations...');
    createUtilisationChart(metrics);
    createStackRankingChart(metrics);
    createUserPerformanceChart(metrics);
    
    // Create user cards
    createUserCards(metrics);
    
    // Update data status
    showDataStatus();
    
    console.log('Data loading complete!');
}

function calculateMetrics(dataToProcess = null) {
    // Use provided data or fallback to global productionData
    const data = dataToProcess || productionData;
    
    // Group users by client/process to form teams
    const teamGroups = {};
    (data || []).forEach(user => {
        const teamKey = `${user.clientName} - ${user.processName}`;
        if (!teamGroups[teamKey]) {
            teamGroups[teamKey] = [];
        }
        teamGroups[teamKey].push(user);
    });
    
    // Process each team
    let allMetrics = [];
    
    Object.keys(teamGroups).forEach(teamName => {
        const teamUsers = teamGroups[teamName];
        
        // Only process teams with 4+ members as per your requirement
        if (teamUsers.length < 4) {
            return; // Skip teams with less than 4 members
        }
        
        const teamMetrics = teamUsers.map(user => {
            const productivity = Number(user.productivity) || 0;
            const target = Number(user.target) || 0;
            const clientErrors = Number(user.clientErrors) || 0;
            const internalErrors = Number(user.internalErrors) || 0;
            const hoursWorked = Number(user.hoursWorked) || 8;
            const actualHours = Number(user.actualHours) || 8;
            
            // 1. New Utilisation Formula: (Actual count / per hour count) / total hours worked * 100
            // Per hour count = target / 8 (assuming 8 hours per day)
            const perHourCount = target / 8;
            const totalHoursWorked = hoursWorked || 8; // Use hoursWorked or default to 8
            const utilisation = (perHourCount > 0 && totalHoursWorked > 0) ? 
                (productivity / perHourCount) / totalHoursWorked * 100 : 0;
            
            // 2. New Stack Ranking Points Calculation
            let stackRankingPoints = 0;
            
            // A. Actual count points (based on target achievement)
            if (target > 0) {
                const achievementRatio = (productivity / target) * 100;
                if (achievementRatio >= 100) stackRankingPoints += 5;      // 100%+ = 5 points
                else if (achievementRatio >= 90) stackRankingPoints += 4;  // 90-99% = 4 points
                else if (achievementRatio >= 80) stackRankingPoints += 3;  // 80-89% = 3 points
                else if (achievementRatio >= 70) stackRankingPoints += 3;  // 70-79% = 3 points
                else if (achievementRatio >= 60) stackRankingPoints += 2;  // 60-69% = 2 points
                else stackRankingPoints += 1;                              // Below 60% = 1 point
            }
            
            // B. Error points (inverse - lower errors = higher points)
            const totalErrors = clientErrors + internalErrors;
            if (totalErrors === 0) stackRankingPoints += 5;                    // 0 errors = 5 points
            else if (totalErrors >= 1 && totalErrors <= 3) stackRankingPoints += 4;  // 1-3 errors = 4 points
            else if (totalErrors >= 4 && totalErrors <= 6) stackRankingPoints += 3;  // 4-6 errors = 3 points
            else if (totalErrors >= 7 && totalErrors <= 9) stackRankingPoints += 2;  // 7-9 errors = 2 points
            else if (totalErrors >= 10) stackRankingPoints += 1;                    // 10+ errors = 1 point
            
            // C. Working days points (based on leaves/attendance)
            const expectedHours = 8 * 22; // 8 hours per day * 22 working days per month
            const missingHours = expectedHours - actualHours;
            const missingDays = Math.max(0, missingHours / 8); // Convert hours to days, minimum 0
            
            if (missingDays <= 0) stackRankingPoints += 2;        // No missing days = 2 points
            else if (missingDays <= 2) stackRankingPoints += 1;   // 1-2 missing days = 1 point
            else if (missingDays <= 4) stackRankingPoints += 0;   // 3-4 missing days = 0 points
            else if (missingDays <= 6) stackRankingPoints -= 1;   // 5-6 missing days = -1 point
            else stackRankingPoints -= 2;                         // 7+ missing days = -2 points
            
            // 3. Team Performance (based on productivity vs target)
            const teamPerformance = target > 0 ? (productivity / target) * 100 : 0;
            
            return {
                ...user,
                utilisation: Math.round(utilisation * 100) / 100,
                stackRankingPoints: stackRankingPoints,
                teamPerformance: Math.round(teamPerformance * 100) / 100,
                performanceStatus: teamPerformance >= 100 ? 'excellent' : 
                                 teamPerformance >= 80 ? 'good' : 'poor',
                team: teamName
            };
        });
        
        // Sort team by stack ranking points for ranking
        teamMetrics.sort((a, b) => b.stackRankingPoints - a.stackRankingPoints);
        
        // Add ranking positions
        teamMetrics.forEach((user, index) => {
            user.stackRankingPosition = index + 1;
            user.utilisationPosition = teamMetrics
                .sort((a, b) => b.utilisation - a.utilisation)
                .findIndex(u => u.employeeId === user.employeeId) + 1;
            user.teamPerformancePosition = teamMetrics
                .sort((a, b) => b.teamPerformance - a.teamPerformance)
                .findIndex(u => u.employeeId === user.employeeId) + 1;
        });
        
        allMetrics = allMetrics.concat(teamMetrics);
    });
    
    return allMetrics;
}

// Load data from Google Sheets asynchronously
async function loadDataFromGoogleSheetsAsync(sheetUrl) {
    try {
        console.log('Loading data from Google Sheets:', sheetUrl);
        
        const { sheetId, gid } = parseSheetUrl(sheetUrl);
        
        if (!sheetId) {
            throw new Error('Invalid Google Sheet URL');
        }
        
        // Fetch data using the existing function
        const { rows, cols } = await fetchGoogleSheetsDataWithFallback(sheetId, gid);
        
        if (rows.length === 0) {
            throw new Error('No data found in the sheet');
        }
        
        // Map to our format
        const mapped = mapGoogleSheetRows(rows, cols);
        console.log('Mapped data:', mapped.length, 'records');
        
        if (mapped.length === 0) {
            throw new Error('Could not map any data from the sheet');
        }
        
        productionData = mapped;
        localStorage.setItem('dm2_production_data', JSON.stringify(productionData));
        console.log('Data loaded from Google Sheets:', productionData.length, 'records');
        
        // Update the saved link display
        updateSavedLinkDisplay();
        
        // Continue with chart creation directly (avoid recursive call)
        if (currentUser) {
            switch(currentUser.role) {
                case 'tl':
                    createChartsAndCards();
                    break;
                case 'manager':
                    createManagerChartsAndCards();
                    break;
            }
        }
    } catch (error) {
        console.error('Error loading data from Google Sheets:', error);
        console.log('Falling back to sample data due to Google Sheets error');
        
        // Fallback to sample data
        productionData = [...sampleProductionData];
        localStorage.setItem('dm2_production_data', JSON.stringify(productionData));
        
        // Update the saved link display
        updateSavedLinkDisplay();
        
        // Continue with chart creation directly (avoid recursive call)
        if (currentUser) {
            switch(currentUser.role) {
                case 'tl':
                    createChartsAndCards();
                    break;
                case 'manager':
                    createManagerChartsAndCards();
                    break;
            }
        }
    }
}

// Create charts and cards for TL portal
function createChartsAndCards() {
    console.log('Creating charts and cards for TL portal...');
    
    // Apply project filter if selected
    let filteredData = [...productionData];
    const projectFilter = document.getElementById('projectFilter');
    if (projectFilter && projectFilter.value) {
        filteredData = filteredData.filter(d => d.clientName === projectFilter.value);
        console.log('Filtered by project:', projectFilter.value, 'Records:', filteredData.length);
    }
    
    // Apply date filter if selected
    filteredData = applyDateFilter(filteredData);
    console.log('After date filtering:', filteredData.length, 'records');
    
    // Calculate metrics with the filtered data
    const metrics = calculateMetrics(filteredData);
    console.log('Calculated metrics:', metrics.length, 'records');
    
    if (!metrics.length) {
        console.warn('No metrics to display');
        return;
    }

    // Create visualizations
    console.log('Creating visualizations...');
    createUtilisationChart(metrics);
    createStackRankingChart(metrics);
    createUserPerformanceChart(metrics);
    
    // Create user cards
    createUserCards(metrics);
    
    // Update data status
    showDataStatus();
    
    console.log('TL charts and cards creation complete!');
}

// Create charts and cards for Manager portal
function createManagerChartsAndCards() {
    console.log('Creating charts and cards for Manager portal...');
    
    // This function is kept for backward compatibility
    // The main function loadManagerData() now handles all chart creation
    loadManagerData();
}

// Enhanced Google Sheets data fetching with multiple fallback approaches
async function fetchGoogleSheetsDataWithFallback(sheetId, gid) {
    const approaches = [
        // Approach 1: Direct gviz API
        `https://docs.google.com/spreadsheets/d/${sheetId}/gviz/tq?tqx=out:json&gid=${gid}`,
        // Approach 2: Alternative gviz format
        `https://docs.google.com/spreadsheets/d/${sheetId}/gviz/tq?tqx=out:json&tq=select%20*&gid=${gid}`,
        // Approach 3: CSV export (fallback)
        `https://docs.google.com/spreadsheets/d/${sheetId}/export?format=csv&gid=${gid}`,
        // Approach 4: CORS proxy for gviz
        `https://cors-anywhere.herokuapp.com/https://docs.google.com/spreadsheets/d/${sheetId}/gviz/tq?tqx=out:json&gid=${gid}`,
        // Approach 5: Alternative CORS proxy (CSV export through proxy)
        `https://api.allorigins.win/raw?url=${encodeURIComponent(`https://docs.google.com/spreadsheets/d/${sheetId}/export?format=csv&gid=${gid}`)}`
    ];

    for (let i = 0; i < approaches.length; i++) {
        const url = approaches[i];
        console.log(`Trying approach ${i + 1}:`, url);
        
        try {
            const response = await fetch(url, { 
                cache: 'no-store', 
                credentials: 'omit',
                headers: {
                    'Accept': 'application/json, text/plain, */*'
                }
            });
            
            if (!response.ok) {
                throw new Error(`HTTP ${response.status}: ${response.statusText}`);
            }
            
            const text = await response.text();
            console.log(`Approach ${i + 1} response length:`, text.length);
            
            // Check if response looks like HTML (error page)
            if (text.includes('<!DOCTYPE html>') || text.includes('<html')) {
                throw new Error('Received HTML instead of data - sheet may not be publicly accessible');
            }
            
            // Content-aware parser: detect format by content, not approach index
            const looksLikeGviz = text.startsWith("/*O_o*/") || text.includes("setResponse(") || text.includes("google.visualization");
            
            if (looksLikeGviz) {
                // Parse as GViz JSON
                const cleaned = text
                    .replace(/^\/\*O_o\*\/\s*/, '')         // CRITICAL FIX: strip /*O_o*/ comment
                    .replace(/^\)]}'\n?/, '')               // strip XSSI guard
                    .replace(/^.*setResponse\(/, '')        // remove wrapper
                    .replace(/\);?\s*$/, '');               // trailing );
                
                const json = JSON.parse(cleaned);
                const cols = (json.table.cols || []).map(c => (c && c.label) ? c.label : '');
                const rows = (json.table.rows || []).map(r => (r.c || []).map(c => c ? (c.v ?? '') : ''));
                
                console.log(`Approach ${i + 1} successful - parsed ${rows.length} rows from GViz JSON`);
                return { rows, cols };
            } else {
                // Parse as CSV
                const lines = text.split('\n').filter(line => line.trim());
                if (lines.length < 2) {
                    throw new Error('No data rows found in CSV');
                }
                
                const cols = lines[0].split(',').map(col => col.replace(/"/g, '').trim());
                const rows = lines.slice(1).map(line => 
                    line.split(',').map(cell => cell.replace(/"/g, '').trim())
                );
                
                console.log(`Approach ${i + 1} successful - parsed ${rows.length} rows from CSV`);
                return { rows, cols };
            }
            
        } catch (error) {
            console.warn(`Approach ${i + 1} failed:`, error.message);
            if (i === approaches.length - 1) {
                throw new Error(`All approaches failed. Last error: ${error.message}`);
            }
        }
    }
}

// Legacy function for backward compatibility
async function fetchGoogleSheetData(gvizUrl) {
    const response = await fetch(gvizUrl, { cache: 'no-store', credentials: 'omit' });
    const text = await response.text();
    const cleaned = text
        .replace(/^\)]}'\n?/, '')
        .replace(/^.*setResponse\(/, '')
        .replace(/\);?\s*$/, '');
    const json = JSON.parse(cleaned);
    const cols = (json.table.cols || []).map(c => (c && c.label) ? c.label : '');
    const rows = (json.table.rows || []).map(r => (r.c || []).map(c => c ? (c.v ?? '') : ''));
    return { rows, cols };
}

function parseSheetUrl(url) {
    try {
        console.log('Parsing URL:', url);
        
        // Clean the URL first
        let cleanUrl = url.trim();
        
        // Handle different URL formats
        if (cleanUrl.includes('/edit')) {
            cleanUrl = cleanUrl.replace('/edit', '');
        }
        
        // Extract sheet ID using multiple regex patterns
        let sheetId = '';
        const patterns = [
            /\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/,
            /\/d\/([a-zA-Z0-9-_]+)/,
            /spreadsheets\/d\/([a-zA-Z0-9-_]+)/
        ];
        
        for (const pattern of patterns) {
            const match = cleanUrl.match(pattern);
            if (match && match[1]) {
                sheetId = match[1];
                break;
            }
        }
        
        // Extract gid using multiple patterns
        let gid = '0';
        const gidPatterns = [
            /[#&]gid=(\d+)/,
            /gid=(\d+)/,
            /#gid=(\d+)/
        ];
        
        for (const pattern of gidPatterns) {
            const match = cleanUrl.match(pattern);
            if (match && match[1]) {
                gid = match[1];
                break;
            }
        }
        
        console.log('Extracted sheetId:', sheetId, 'gid:', gid);
        
        if (!sheetId) {
            throw new Error('Could not extract sheet ID from URL');
        }
        
        return { sheetId, gid };
    } catch (error) {
        console.error('Error parsing sheet URL:', error);
        return { sheetId: '', gid: '0' };
    }
}

function mapGoogleSheetRows(rows, cols) {
    console.log('Mapping Google Sheets data...');
    console.log('Available columns:', cols);
    
    // More flexible column finding with multiple patterns - prioritize exact matches
    const findColIndex = (patterns) => {
        for (const pattern of patterns) {
            const index = cols.findIndex(col => {
                if (!col) return false;
                const colTrimmed = col.trim();
                const patternTrimmed = pattern.trim();
                
                // PRIORITY 1: Exact match (case-sensitive)
                if (colTrimmed === patternTrimmed) return true;
                
                // PRIORITY 2: Exact match (case-insensitive)
                if (colTrimmed.toLowerCase() === patternTrimmed.toLowerCase()) return true;
                
                // PRIORITY 3: Contains match (case-insensitive)
                if (colTrimmed.toLowerCase().includes(patternTrimmed.toLowerCase())) return true;
                
                // PRIORITY 4: Handle common variations
                const colLower = colTrimmed.toLowerCase();
                const patternLower = patternTrimmed.toLowerCase();
                
                const variations = [
                    colLower.replace(/[^a-z0-9]/g, ''), // Remove special chars
                    colLower.replace(/\s+/g, ''), // Remove spaces
                    colLower.replace(/[#]/g, ''), // Remove # symbol
                ];
                
                for (const variation of variations) {
                    if (variation.includes(patternLower.replace(/[^a-z0-9]/g, ''))) {
                        return true;
                    }
                }
                
                return false;
            });
            if (index !== -1) return index;
        }
        return -1;
    };
    
    // Map to your exact format: Date, Employee ID, Name, Client Name, Process Name, Productivity, Target, Client Errors, Internal Errors, Hours Worked, Actual Hours
    // CRITICAL FIX: Handle actual API response codes and map by position if headers are corrupted
    
    // First, try to find exact column names
    const dateIdx = findColIndex(['Date', 'date', 'month', 'period', 'time', 'timestamp']);
    const employeeIdIdx = findColIndex(['Employee ID', 'ITPL', 'itpl', 'employee id', 'emp id', 'id', 'employeeid']);
    const nameIdx = findColIndex(['Name', 'name', 'employee name', 'employeename']);
    const clientNameIdx = findColIndex(['Client Name', 'client name', 'clientname', 'client', 'company', 'customer']);
    const processNameIdx = findColIndex(['Process Name', 'process name', 'processname', 'process']);
    const productivityIdx = findColIndex(['Productivity', 'productivity', 'actual', 'count', 'volume', 'output']);
    const targetIdx = findColIndex(['Target', 'target', 'monthly target', 'monthlytarget', 'goal']);
    const clientErrorsIdx = findColIndex(['Client Errors', 'client errors', 'clienterror', 'client error']);
    const internalErrorsIdx = findColIndex(['Internal Errors', 'internal errors', 'internalerror', 'internal error']);
    const hoursWorkedIdx = findColIndex(['Hours Worked', 'hours worked', 'working days', 'workdays']);
    const actualHoursIdx = findColIndex(['Actual Hours', 'actual hours', 'actualhours', 'hours']);
    
    // If essential columns not found, try to map by position (assuming standard order)
    let finalDateIdx = dateIdx;
    let finalEmployeeIdIdx = employeeIdIdx;
    let finalNameIdx = nameIdx;
    let finalClientNameIdx = clientNameIdx;
    let finalProcessNameIdx = processNameIdx;
    let finalProductivityIdx = productivityIdx;
    let finalTargetIdx = targetIdx;
    let finalClientErrorsIdx = clientErrorsIdx;
    let finalInternalErrorsIdx = internalErrorsIdx;
    let finalHoursWorkedIdx = hoursWorkedIdx;
    let finalActualHoursIdx = actualHoursIdx;
    
    if (employeeIdIdx === -1 || nameIdx === -1) {
        console.log('Essential columns not found by name, trying position-based mapping...');
        console.log('Available columns:', cols);
        
        // Map by position based on your sheet structure: Date, ITPL, Name, Process Name, Productivity, Target, Client Errors, Internal Errors, Hours Worked, Actual Hours
        if (cols.length >= 10) {
            console.log('Using position-based mapping for columns');
            finalDateIdx = 0;        // Column A: Date
            finalEmployeeIdIdx = 1;  // Column B: ITPL (Employee ID)
            finalNameIdx = 2;        // Column C: Name
            finalClientNameIdx = 3;  // Column D: Process Name (use as Client Name)
            finalProcessNameIdx = 3; // Column D: Process Name
            finalProductivityIdx = 4;// Column E: Productivity
            finalTargetIdx = 5;      // Column F: Target
            finalClientErrorsIdx = 6;// Column G: Client Errors
            finalInternalErrorsIdx = 7; // Column H: Internal Errors
            finalHoursWorkedIdx = 8; // Column I: Hours Worked
            finalActualHoursIdx = 9; // Column J: Actual Hours
        }
    }
    
    console.log('Column mapping results:', {
        date: finalDateIdx, employeeId: finalEmployeeIdIdx, name: finalNameIdx, clientName: finalClientNameIdx,
        processName: finalProcessNameIdx, productivity: finalProductivityIdx, target: finalTargetIdx,
        clientErrors: finalClientErrorsIdx, internalErrors: finalInternalErrorsIdx,
        hoursWorked: finalHoursWorkedIdx, actualHours: finalActualHoursIdx
    });
    
    // If we can't find essential columns, show helpful error with detailed debugging
    if (finalEmployeeIdIdx === -1 || finalNameIdx === -1) {
        const foundCols = cols.filter(col => col && col.trim()).join(', ');
        console.error('Column mapping failed. Details:');
        console.error('Available columns:', cols);
        console.error('Column indices found:', { employeeIdIdx, nameIdx });
        console.error('All column indices:', {
            date: dateIdx, employeeId: employeeIdIdx, name: nameIdx, clientName: clientNameIdx,
            processName: processNameIdx, productivity: productivityIdx, target: targetIdx,
            clientErrors: clientErrorsIdx, internalErrors: internalErrorsIdx,
            hoursWorked: hoursWorkedIdx, actualHours: actualHoursIdx
        });
        
        // Try to find any column that might be ID or Name
        const possibleIdCols = cols.filter((col, idx) => 
            col && (col.toLowerCase().includes('id') || col.toLowerCase().includes('itpl'))
        );
        const possibleNameCols = cols.filter((col, idx) => 
            col && col.toLowerCase().includes('name')
        );
        
        let errorMsg = `Essential columns not found.\n\nFound columns: ${foundCols}\n\n`;
        if (possibleIdCols.length > 0) {
            errorMsg += `Possible ID columns: ${possibleIdCols.join(', ')}\n`;
        }
        if (possibleNameCols.length > 0) {
            errorMsg += `Possible Name columns: ${possibleNameCols.join(', ')}\n`;
        }
        errorMsg += `\nPlease ensure your sheet has columns for Employee ID and Name.`;
        
        throw new Error(errorMsg);
    }
    
    const mapped = rows
        .filter((row, index) => {
            // Skip empty rows or header row
            if (index === 0) return false;
            // At minimum, we need Employee ID and Name
            return row[finalEmployeeIdIdx] && row[finalNameIdx] && 
                   String(row[finalEmployeeIdIdx]).trim() && String(row[finalNameIdx]).trim();
        })
        .map(row => {
            const toNum = (val) => {
                if (!val) return 0;
                const str = String(val).trim();
                if (str === '' || str === '-' || str.toLowerCase() === 'na' || str.toLowerCase() === 'leave') return 0;
                const num = Number(str);
                return isNaN(num) ? 0 : num;
            };
            
            return {
                date: row[finalDateIdx] ? String(row[finalDateIdx]) : '',
                employeeId: String(row[finalEmployeeIdIdx] || '').trim(),
                name: String(row[finalNameIdx] || '').trim(),
                clientName: String(row[finalClientNameIdx] || '').trim(),
                processName: String(row[finalProcessNameIdx] || '').trim(),
                productivity: toNum(row[finalProductivityIdx]),
                target: toNum(row[finalTargetIdx]),
                clientErrors: toNum(row[finalClientErrorsIdx]),
                internalErrors: toNum(row[finalInternalErrorsIdx]),
                hoursWorked: toNum(row[finalHoursWorkedIdx]) || 8,
                actualHours: toNum(row[finalActualHoursIdx]) || 8
            };
        })
        .filter(record => record.employeeId && record.name);
    
    console.log('Successfully mapped records:', mapped.length);
    return mapped;
}

function createUtilisationChart(metrics) {
    const ctx = document.getElementById('utilisationChart');
    if (!ctx) {
        console.error('Utilisation chart canvas not found');
        return;
    }
    
    // Check if Chart.js is loaded
    if (typeof Chart === 'undefined') {
        console.error('Chart.js library not loaded');
        ctx.innerHTML = '<p style="color: red; padding: 20px;">Chart.js library not loaded. Please check internet connection and refresh the page.</p>';
        return;
    }
    
    // before creating new chart
    safeDestroyChart('utilisationChart');
    
    console.log('Creating utilisation chart with metrics:', metrics.length, 'records');
    
    if (!metrics || metrics.length === 0) {
        console.warn('No metrics data for utilisation chart');
        ctx.innerHTML = '<p style="color: #666; padding: 20px; text-align: center;">No data available for utilisation chart</p>';
        return;
    }
    
    // Validate metrics data
    const validMetrics = metrics.filter(m => m && typeof m.utilisation === 'number' && !isNaN(m.utilisation));
    if (validMetrics.length === 0) {
        console.warn('No valid metrics data for utilisation chart');
        ctx.innerHTML = '<p style="color: #666; padding: 20px; text-align: center;">No valid data available for utilisation chart</p>';
        return;
    }
    
    // Remove duplicates and sort by utilisation descending
    const uniqueEmployees = [];
    const seenIds = new Set();
    
    validMetrics.forEach(employee => {
        if (!seenIds.has(employee.employeeId)) {
            seenIds.add(employee.employeeId);
            uniqueEmployees.push(employee);
        }
    });
    
    // Sort by utilisation descending (best performers first)
    uniqueEmployees.sort((a, b) => b.utilisation - a.utilisation);
    
    try {
        window.utilisationChart = new Chart(ctx, {
            type: 'bar',
            data: {
                labels: uniqueEmployees.map(m => `${m.name} (${m.clientName})`),
                datasets: [{
                    label: 'Utilisation %',
                    data: uniqueEmployees.map(m => m.utilisation),
                    backgroundColor: uniqueEmployees.map(m => 
                        m.utilisation >= 100 ? '#28a745' : 
                        m.utilisation >= 80 ? '#ffc107' : '#dc3545'
                    ),
                    borderColor: uniqueEmployees.map(m => 
                        m.utilisation >= 100 ? '#1e7e34' : 
                        m.utilisation >= 80 ? '#e0a800' : '#bd2130'
                    ),
                    borderWidth: 2
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                scales: {
                    y: {
                        beginAtZero: true,
                        max: 400,
                        title: {
                            display: true,
                            text: 'Utilisation %'
                        }
                    },
                    x: {
                        title: {
                            display: true,
                            text: 'Employees'
                        }
                    }
                },
                plugins: {
                    legend: {
                        display: false
                    },
                    title: {
                        display: true,
                        text: 'Employee Utilisation Analysis (Productivity/Per Hour Count/Hours Worked * 100)'
                    },
                    tooltip: {
                        callbacks: {
                            title: ctx => `${uniqueEmployees[ctx[0].dataIndex].name} (${uniqueEmployees[ctx[0].dataIndex].clientName})`
                        }
                    }
                }
            }
        });
    } catch (error) {
        console.error('Error creating utilisation chart:', error);
        ctx.innerHTML = '<p style="color: red; padding: 20px;">Error creating chart. Please try again.</p>';
    }
}

function createStackRankingChart(metrics) {
    const ctx = document.getElementById('stackRankingChart');
    if (!ctx) {
        console.error('Stack ranking chart canvas not found');
        return;
    }
    
    // Check if Chart.js is loaded
    if (typeof Chart === 'undefined') {
        console.error('Chart.js library not loaded');
        ctx.innerHTML = '<p style="color: red; padding: 20px;">Chart.js library not loaded. Please check internet connection.</p>';
        return;
    }
    
    // before creating new chart
    safeDestroyChart('stackRankingChart');
    
    if (!metrics || metrics.length === 0) {
        console.warn('No metrics data for stack ranking chart');
        ctx.innerHTML = '<p style="color: #666; padding: 20px; text-align: center;">No data available for stack ranking chart</p>';
        return;
    }
    
    // Remove duplicates and sort by stack ranking points descending
    const uniqueEmployees = [];
    const seenIds = new Set();
    
    metrics.forEach(employee => {
        if (!seenIds.has(employee.employeeId)) {
            seenIds.add(employee.employeeId);
            uniqueEmployees.push(employee);
        }
    });
    
    // Sort by stack ranking points descending (best performers first)
    uniqueEmployees.sort((a, b) => b.stackRankingPoints - a.stackRankingPoints);
    
    try {
        window.stackRankingChart = new Chart(ctx, {
            type: 'bar',
            data: {
                labels: uniqueEmployees.map(m => `${m.name} (${m.clientName})`),
                datasets: [{
                    label: 'Stack Ranking Points',
                    data: uniqueEmployees.map(m => m.stackRankingPoints),
                    backgroundColor: uniqueEmployees.map((m, index) => {
                        if (index === 0) return '#ffd700'; // Gold for 1st
                        if (index === 1) return '#c0c0c0'; // Silver for 2nd
                        if (index === 2) return '#cd7f32'; // Bronze for 3rd
                        return '#667eea';
                    }),
                    borderColor: uniqueEmployees.map((m, index) => {
                        if (index === 0) return '#ffb300'; // Gold border
                        if (index === 1) return '#a0a0a0'; // Silver border
                        if (index === 2) return '#b8860b'; // Bronze border
                        return '#5a6fd8';
                    }),
                    borderWidth: 2
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                scales: {
                    y: {
                        beginAtZero: true,
                        max: 12,
                        title: {
                            display: true,
                            text: 'Stack Ranking Points'
                        }
                    },
                    x: {
                        title: {
                            display: true,
                            text: 'Employees (Ranked)'
                        }
                    }
                },
                plugins: {
                    legend: {
                        display: false
                    },
                    title: {
                        display: true,
                        text: 'Stack Ranking (Target Achievement + Errors + Working Days Points)'
                    },
                    tooltip: {
                        callbacks: {
                            title: ctx => `${uniqueEmployees[ctx[0].dataIndex].name} (${uniqueEmployees[ctx[0].dataIndex].clientName})`
                        }
                    }
                }
            }
        });
    } catch (error) {
        console.error('Error creating stack ranking chart:', error);
        ctx.innerHTML = '<p style="color: red; padding: 20px;">Error creating chart. Please try again.</p>';
    }
}

function createUserPerformanceChart(metrics) {
    const ctx = document.getElementById('userPerformanceChart');
    if (!ctx) {
        console.error('User performance chart canvas not found');
        return;
    }
    
    // Check if Chart.js is loaded
    if (typeof Chart === 'undefined') {
        console.error('Chart.js library not loaded');
        ctx.innerHTML = '<p style="color: red; padding: 20px;">Chart.js library not loaded. Please check internet connection.</p>';
        return;
    }
    
    // before creating new chart
    safeDestroyChart('userPerformanceChart');
    
    if (!metrics || metrics.length === 0) {
        console.warn('No metrics data for user performance chart');
        ctx.innerHTML = '<p style="color: #666; padding: 20px; text-align: center;">No data available for user performance chart</p>';
        return;
    }
    
    // Remove duplicates and sort by team performance descending
    const uniqueEmployees = [];
    const seenIds = new Set();
    
    metrics.forEach(employee => {
        if (!seenIds.has(employee.employeeId)) {
            seenIds.add(employee.employeeId);
            uniqueEmployees.push(employee);
        }
    });
    
    // Sort by team performance descending (best performers first)
    uniqueEmployees.sort((a, b) => b.teamPerformance - a.teamPerformance);
    
    try {
        window.userPerformanceChart = new Chart(ctx, {
            type: 'line',
            data: {
                labels: uniqueEmployees.map(m => `${m.name} (${m.clientName})`),
                datasets: [{
                    label: 'Team Performance %',
                    data: uniqueEmployees.map(m => m.teamPerformance),
                    borderColor: '#667eea',
                    backgroundColor: 'rgba(102, 126, 234, 0.1)',
                    tension: 0.4,
                    fill: true,
                    borderWidth: 3,
                    pointBackgroundColor: uniqueEmployees.map(m => 
                        m.teamPerformance >= 100 ? '#28a745' : 
                        m.teamPerformance >= 80 ? '#ffc107' : '#dc3545'
                    ),
                    pointBorderColor: '#fff',
                    pointBorderWidth: 2,
                    pointRadius: 6
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                scales: {
                    y: {
                        beginAtZero: true,
                        max: 150,
                        title: {
                            display: true,
                            text: 'Team Performance %'
                        }
                    },
                    x: {
                        title: {
                            display: true,
                            text: 'Employees'
                        }
                    }
                },
                plugins: {
                    title: {
                        display: true,
                        text: 'Team Performance Analysis (Volume vs Target)'
                    },
                    tooltip: {
                        callbacks: {
                            title: ctx => `${uniqueEmployees[ctx[0].dataIndex].name} (${uniqueEmployees[ctx[0].dataIndex].clientName})`
                        }
                    }
                }
            }
        });
    } catch (error) {
        console.error('Error creating user performance chart:', error);
        ctx.innerHTML = '<p style="color: red; padding: 20px;">Error creating chart. Please try again.</p>';
    }
}

function createUserCards(metrics) {
    const container = document.getElementById('userCards');
    container.innerHTML = '';
    
    if (!metrics || metrics.length === 0) {
        container.innerHTML = '<p style="text-align: center; color: #666; padding: 20px;">No performance data available.</p>';
        return;
    }
    
    // Group all records by employeeId to calculate total productivity across all projects
    const employeeGroups = {};
    metrics.forEach(employee => {
        const empId = employee.employeeId;
        if (!employeeGroups[empId]) {
            employeeGroups[empId] = {
                employeeId: empId,
                name: employee.name,
                clientName: employee.clientName,
                target: employee.target || 0,
                clientErrors: employee.clientErrors || 0,
                stackRankingPoints: employee.stackRankingPoints || 0,
                totalProductivity: 0,
                utilisation: employee.utilisation || 0,
                teamPerformance: employee.teamPerformance || 0
            };
        }
        // Sum productivity across all projects/tasks for this employee
        employeeGroups[empId].totalProductivity += (employee.productivity || 0);
        // Use the highest values for other metrics
        if (employee.utilisation > employeeGroups[empId].utilisation) {
            employeeGroups[empId].utilisation = employee.utilisation;
        }
        if (employee.teamPerformance > employeeGroups[empId].teamPerformance) {
            employeeGroups[empId].teamPerformance = employee.teamPerformance;
        }
    });
    
    // Group by project and find top performer from each project
    const projectGroups = {};
    Object.values(employeeGroups).forEach(employee => {
        const projectKey = (employee.clientName || '').toLowerCase();
        if (!projectGroups[projectKey]) {
            projectGroups[projectKey] = [];
        }
        projectGroups[projectKey].push(employee);
    });
    
    console.log('TL Project groups:', Object.keys(projectGroups));
    console.log('TL Project groups details:', projectGroups);
    
    // Get top performer from each project (highest total productivity)
    const topPerformersByProject = [];
    Object.values(projectGroups).forEach(projectEmployees => {
        // Sort by total productivity descending to get the best performer
        projectEmployees.sort((a, b) => b.totalProductivity - a.totalProductivity);
        if (projectEmployees.length > 0) {
            topPerformersByProject.push(projectEmployees[0]);
        }
    });
    
    console.log('TL Top performers by project:', topPerformersByProject.length);
    
    // Sort all top performers by total productivity descending
    topPerformersByProject.sort((a, b) => b.totalProductivity - a.totalProductivity);
    
    // Show top 3 performers (one from each project) - TL Portal
    const topPerformers = topPerformersByProject.slice(0, 3);
    
    console.log('TL Final top performers:', topPerformers.length);
    
    topPerformers.forEach((user, index) => {
        const card = document.createElement('div');
        card.className = 'user-card';
        card.innerHTML = `
            <h4>${index === 0 ? '🥇' : index === 1 ? '🥈' : '🥉'} ${user.name}</h4>
            <div class="user-info">
                <div class="info-row">
                    <span class="info-label">Employee ID:</span>
                    <span class="info-value">${user.employeeId || 'N/A'}</span>
                </div>
                <div class="info-row">
                    <span class="info-label">Client/Project Name:</span>
                    <span class="info-value">${user.clientName || 'N/A'}</span>
                </div>
                <div class="info-row">
                    <span class="info-label">Total Productivity:</span>
                    <span class="info-value">${user.totalProductivity.toLocaleString()}</span>
                </div>
                <div class="info-row">
                    <span class="info-label">Client Errors:</span>
                    <span class="info-value">${user.clientErrors}</span>
                </div>
                <div class="info-row">
                    <span class="info-label">Stack Ranking:</span>
                    <span class="info-value">#${index + 1}</span>
                </div>
                <div class="info-row">
                    <span class="info-label">Utilisation:</span>
                    <span class="info-value">${user.utilisation.toFixed(1)}%</span>
                </div>
                <div class="info-row">
                    <span class="info-label">Team Performance:</span>
                    <span class="info-value">${user.teamPerformance.toFixed(1)}%</span>
                </div>
            </div>
        `;
        container.appendChild(card);
    });
}

function showTab(tabName, ev) {
    // Hide all tabs
    document.querySelectorAll('.tab-content').forEach(tab => {
        tab.classList.remove('active');
    });
    document.querySelectorAll('.tab-btn').forEach(btn => {
        btn.classList.remove('active');
    });
    
    // Show selected tab
    document.getElementById(tabName + 'Tab').classList.add('active');
    if (ev && ev.target) ev.target.classList.add('active');
}

function loadFeedbackMessages() {
    const container = document.getElementById('feedbackMessages');
    container.innerHTML = '';
    
    // Get all feedback messages involving this user (both sent and received)
    const userFeedback = feedbackMessages.filter(f => 
        f.to === currentUser.email || f.from === currentUser.email
    );
    
    if (userFeedback.length === 0) {
        container.innerHTML = '<p style="text-align: center; color: #666; padding: 20px;">No feedback messages yet.</p>';
        return;
    }
    
    // Group messages by conversation thread
    const conversations = {};
    userFeedback.forEach(feedback => {
        const threadId = feedback.parentId || feedback.id;
        if (!conversations[threadId]) {
            conversations[threadId] = [];
        }
        conversations[threadId].push(feedback);
    });
    
    // Display each conversation
    Object.values(conversations).forEach(conversation => {
        conversation.sort((a, b) => new Date(a.timestamp) - new Date(b.timestamp));
        
        const conversationDiv = document.createElement('div');
        conversationDiv.className = 'conversation-thread';
        
        let conversationHTML = '<div class="conversation-header">';
        const otherUser = conversation[0].from === currentUser.email ? conversation[0].to : conversation[0].from;
        conversationHTML += `<h5>Conversation with: ${otherUser}</h5>`;
        conversationHTML += '</div>';
        
        conversation.forEach(msg => {
            const isFromCurrentUser = msg.from === currentUser.email;
            conversationHTML += `
                <div class="message ${isFromCurrentUser ? 'message-sent' : 'message-received'}">
                    <div class="message-header">
                        <strong>${isFromCurrentUser ? 'You' : msg.from}</strong>
                        <small>${new Date(msg.timestamp).toLocaleString()}</small>
                    </div>
                    <div class="message-content">${msg.message}</div>
                </div>
            `;
        });
        
        // Add reply section only for messages received by current user
        const lastMessage = conversation[conversation.length - 1];
        if (lastMessage.to === currentUser.email) {
            conversationHTML += `
                <div class="reply-section">
                    <button onclick="showReplySection(${lastMessage.id})" class="btn btn-primary btn-sm">
                        <i class="fas fa-reply"></i> Reply to this feedback
                    </button>
                </div>
            `;
        }
        
        conversationDiv.innerHTML = conversationHTML;
        container.appendChild(conversationDiv);
    });
}

// Global variable to store current feedback ID for reply
let currentReplyFeedbackId = null;

function showReplySection(feedbackId) {
    currentReplyFeedbackId = feedbackId;
    const replySection = document.getElementById('replySection');
    const replyMessage = document.getElementById('replyMessage');
    
    if (replySection && replyMessage) {
        replySection.style.display = 'block';
        replyMessage.value = '';
        replyMessage.focus();
        
        // Scroll to reply section
        replySection.scrollIntoView({ behavior: 'smooth' });
    }
}

function cancelReply() {
    const replySection = document.getElementById('replySection');
    const replyMessage = document.getElementById('replyMessage');
    
    if (replySection && replyMessage) {
        replySection.style.display = 'none';
        replyMessage.value = '';
        currentReplyFeedbackId = null;
    }
}

function sendReplyToManager() {
    const replyMessage = document.getElementById('replyMessage');
    const replyText = replyMessage.value.trim();
    
    if (!replyText) {
        alert('Please enter a reply message.');
        return;
    }
    
    if (!currentReplyFeedbackId) {
        alert('No feedback selected for reply.');
        return;
    }
    
    // Find the original message to get the correct recipient
    const originalMessage = feedbackMessages.find(f => f.id === currentReplyFeedbackId);
    if (!originalMessage) {
        alert('Original message not found.');
        return;
    }
    
    const reply = {
        id: Date.now(),
        from: currentUser.email,
        to: originalMessage.from,
        message: replyText,
        timestamp: new Date().toISOString(),
        parentId: currentReplyFeedbackId,
        read: false
    };
    
    feedbackMessages.push(reply);
    localStorage.setItem('dm2_feedback', JSON.stringify(feedbackMessages));
    
    // Clear and hide reply section
    replyMessage.value = '';
    cancelReply();
    
    alert('✅ Reply sent successfully!');
    loadFeedbackMessages();
}

function sendReply(feedbackId) {
    // Legacy function for backward compatibility
    showReplySection(feedbackId);
}

function showDataStatus() {
    let statusDiv = document.getElementById('dataStatus');
    if (!statusDiv) {
        statusDiv = document.createElement('div');
        statusDiv.id = 'dataStatus';
        statusDiv.className = 'data-status';
        
        // Insert after the data upload section
        const uploadSection = document.querySelector('.data-upload-section');
        if (uploadSection) {
            uploadSection.insertAdjacentElement('afterend', statusDiv);
        }
    }
    
    if (productionData && productionData.length > 0) {
        statusDiv.innerHTML = `
            <div class="status-success">
                <i class="fas fa-check-circle"></i>
                <strong>Data Loaded:</strong> ${productionData.length} employees from ${new Set(productionData.map(d => d.project)).size} projects
                <br><small>Last updated: ${new Date().toLocaleString()}</small>
                <br><button onclick="clearDataAndReload()" class="btn btn-secondary btn-sm" style="margin-top: 10px;">
                    <i class="fas fa-refresh"></i> Clear & Reload Data
                </button>
            </div>
        `;
    } else {
        statusDiv.innerHTML = `
            <div class="status-warning">
                <i class="fas fa-exclamation-triangle"></i>
                <strong>No Data:</strong> Please upload Google Sheets or Excel file to see data
            </div>
        `;
    }
}

function clearDataAndReload() {
    if (confirm('Are you sure you want to clear all data and reload? This will remove uploaded data and show sample data.')) {
        localStorage.removeItem('dm2_production_data');
        productionData = [...sampleProductionData];
        localStorage.setItem('dm2_production_data', JSON.stringify(productionData));
        loadProductionData();
        alert('✅ Data cleared and reloaded with sample data!');
    }
}

// Initialize date inputs on page load
document.addEventListener('DOMContentLoaded', function() {
    updateDateInputs();
});

