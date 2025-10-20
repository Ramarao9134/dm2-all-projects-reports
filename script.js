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
    { email: 'tl@dm2.com', password: 'tl123', role: 'tl', status: 'active' }
];

// Real data based on Google Sheets structure
const sampleProductionData = [
    { userId: 'ITPL6647', name: 'Ramya', project: 'Widespread Electric', target: 119, actual: 119, leaves: 0, errors: 0, month: '2024-04', workingDays: 22, team: 'Team A', processName: 'Billing Invoice' },
    { userId: 'ITPL9648', name: 'Geetha', project: 'Corr-Jensen', target: 18, actual: 18, leaves: 0, errors: 0, month: '2024-04', workingDays: 22, team: 'Team B', processName: 'Account Payables' },
    { userId: 'ITPL8058', name: 'Giri Prasad R', project: 'Soho Studio', target: 71, actual: 71, leaves: 0, errors: 0, month: '2024-04', workingDays: 22, team: 'Team A', processName: 'ST - Order Entry & Claims' },
    { userId: 'ITPL10359', name: 'A Dharani', project: 'Soho Studio', target: 92, actual: 92, leaves: 0, errors: 0, month: '2024-04', workingDays: 22, team: 'Team A', processName: 'ST - Order Entry' },
    { userId: 'ITPL8360', name: 'Meddabalmi Vishnuchakram', project: 'Soho Studio', target: 28, actual: 28, leaves: 0, errors: 0, month: '2024-04', workingDays: 22, team: 'Team A', processName: 'ST - Claims' },
    { userId: 'ITPL9668', name: 'K Kiruba Karan', project: 'Medscan Lab', target: 3, actual: 3, leaves: 0, errors: 0, month: '2024-04', workingDays: 22, team: 'Team C', processName: 'Blood Samples' },
    { userId: 'ITPL11680', name: 'M Deepika', project: 'Medscan Lab', target: 19, actual: 19, leaves: 0, errors: 0, month: '2024-04', workingDays: 22, team: 'Team C', processName: 'Blood Samples' },
    { userId: 'ITPL10376', name: 'Yavanika', project: 'Curexa Pharmacy', target: 220, actual: 196, leaves: 0, errors: 0, month: '2024-04', workingDays: 22, team: 'Team B', processName: 'Curexa - OE' },
    { userId: 'RITPL5165', name: 'Surya Narayan', project: 'Curexa Pharmacy', target: 220, actual: 225, leaves: 0, errors: 0, month: '2024-04', workingDays: 22, team: 'Team B', processName: 'Curexa - OE' },
    { userId: 'ITPL10723', name: 'Meghana', project: 'Curexa Pharmacy', target: 220, actual: 159, leaves: 0, errors: 0, month: '2024-04', workingDays: 22, team: 'Team B', processName: 'Curexa - OE' },
    { userId: 'RITPL5298', name: 'Shyam Kumar', project: 'Curexa Pharmacy', target: 220, actual: 166, leaves: 0, errors: 0, month: '2024-04', workingDays: 22, team: 'Team B', processName: 'Curexa - OE' },
    { userId: 'ITPL9727', name: 'Mounika', project: 'Curexa Pharmacy', target: 220, actual: 209, leaves: 0, errors: 0, month: '2024-04', workingDays: 22, team: 'Team B', processName: 'Curexa - OE' },
    { userId: 'ITPL9146', name: 'Nagendra', project: 'Curexa Pharmacy', target: 220, actual: 119, leaves: 0, errors: 0, month: '2024-04', workingDays: 22, team: 'Team B', processName: 'Curexa - OE' },
    { userId: 'RITPL5544', name: 'Pavan Kumar', project: 'Curexa Pharmacy', target: 220, actual: 0, leaves: 1, errors: 0, month: '2024-04', workingDays: 22, team: 'Team B', processName: 'Curexa - OE' },
    { userId: 'ITPL9646', name: 'Kusuma', project: 'Curexa Pharmacy', target: 220, actual: 0, leaves: 1, errors: 0, month: '2024-04', workingDays: 22, team: 'Team B', processName: 'Curexa - OE' },
    { userId: 'ITPL8128', name: 'Rajesh T', project: 'Hit Promo', target: 60, actual: 60, leaves: 0, errors: 0, month: '2024-04', workingDays: 22, team: 'Team C', processName: 'Hit - QC' },
    { userId: 'ITPL8326', name: 'Uday Kiran Y', project: 'Hit Promo', target: 60, actual: 60, leaves: 0, errors: 0, month: '2024-04', workingDays: 22, team: 'Team C', processName: 'Hit - QC' }
];

// Initialize the application
document.addEventListener('DOMContentLoaded', function() {
    // Clear old cached sample data only if schema changed
    if (!localStorage.getItem('dm2_production_data')) {
        localStorage.removeItem('dm2_production_data');
    }
    initializeApp();
});

function initializeApp() {
    // Load users from localStorage or use sample data
    const savedUsers = localStorage.getItem('dm2_users');
    if (savedUsers) {
        users = JSON.parse(savedUsers);
    } else {
        users = [...sampleUsers];
        localStorage.setItem('dm2_users', JSON.stringify(users));
    }
    
    // Load production data from localStorage, else seed with sample for first run
    const existing = localStorage.getItem('dm2_production_data');
    if (existing) {
        try {
            productionData = JSON.parse(existing) || [];
        } catch (_) {
            productionData = [...sampleProductionData];
        }
    } else {
        productionData = [...sampleProductionData];
        localStorage.setItem('dm2_production_data', JSON.stringify(productionData));
    }
    
    // Load feedback messages
    const savedFeedback = localStorage.getItem('dm2_feedback');
    if (savedFeedback) {
        feedbackMessages = JSON.parse(savedFeedback);
    }
    
    // Set up event listeners
    setupEventListeners();
    
    // Show login page
    showPage('loginPage');
}

function setupEventListeners() {
    // Login form
    document.getElementById('loginForm').addEventListener('submit', handleLogin);
    
    // Admin form
    document.getElementById('createUserForm').addEventListener('submit', handleCreateUser);
    
    // Manager feedback form
    document.getElementById('managerFeedbackForm').addEventListener('submit', handleManagerFeedback);
    
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

// Login Functionality
function handleLogin(e) {
    e.preventDefault();
    
    const email = document.getElementById('email').value;
    const password = document.getElementById('password').value;
    const errorDiv = document.getElementById('loginError');
    
    // Find user
    const user = users.find(u => u.email === email && u.password === password);
    
    if (user && user.status === 'active') {
        currentUser = user;
        errorDiv.classList.remove('show');
        
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
        }
    } else {
        errorDiv.textContent = 'Invalid email or password';
        errorDiv.classList.add('show');
    }
}

// Logout Functionality
function logout() {
    currentUser = null;
    showPage('loginPage');
    document.getElementById('loginForm').reset();
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
        loadManagerFilters();
        loadManagerData();
        loadManagerFeedbackMessages();
    }
}

function loadManagerFilters() {
    // Load projects
    const projectFilter = document.getElementById('managerProjectFilter');
    const projects = [...new Set(productionData.map(d => d.project))];
    projectFilter.innerHTML = '<option value="">All Projects</option>';
    projects.forEach(project => {
        const option = document.createElement('option');
        option.value = project;
        option.textContent = project;
        projectFilter.appendChild(option);
    });
    
    // Load TLs (for feedback)
    const feedbackTL = document.getElementById('feedbackTL');
    const tls = users.filter(u => u.role === 'tl' && u.status === 'active');
    feedbackTL.innerHTML = '<option value="">Select TL/Coordinator</option>';
    tls.forEach(tl => {
        const option = document.createElement('option');
        option.value = tl.email;
        option.textContent = tl.email;
        feedbackTL.appendChild(option);
    });
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
    
    // Create charts with real data
    createManagerProjectChart();
    createManagerTeamChart();
    
    // Create manager ID cards with top performers
    createManagerIDCards();
    
    console.log('Manager data loading complete');
}

function createManagerProjectChart() {
  const container = document.getElementById('managerProjectChart');
  const canvas = document.getElementById('managerProjectCanvas');
  console.log('Manager project chart - container:', container, 'canvas:', canvas);
  if (!container || !canvas) { 
    console.error('Manager project chart elements not found - container:', !!container, 'canvas:', !!canvas); 
    return; 
  }

  if (typeof Chart === 'undefined') {
    container.innerHTML = '<p style="color:red;padding:20px;">Chart.js not loaded.</p>';
    return;
  }

  // Destroy previous
  if (window.managerProjectChart && typeof window.managerProjectChart.destroy === 'function') {
    try { window.managerProjectChart.destroy(); } catch(e) { console.warn('project destroy failed', e); }
  }

  // Data prep (unchanged)
  const projects = [...new Set(productionData.map(d => d.project))];
  if (projects.length === 0) {
    container.innerHTML = '<p style="color:#666;padding:20px;text-align:center;">No project data available</p>';
    return;
  }

  const projectStats = projects.map(project => {
    const projectData = productionData.filter(d => d.project === project);
    const totalActual = projectData.reduce((s, d) => s + (d.actual || 0), 0);
    const totalTarget = projectData.reduce((s, d) => s + (d.target || 0), 0);
    const performance = totalTarget > 0 ? Math.round((totalActual / totalTarget) * 100) : 0;
    return { name: project, count: projectData.length, totalActual, totalTarget, performance };
  });

  const colors = ['#667eea','#764ba2','#f093fb','#4ecdc4','#45b7d1','#96ceb4','#feca57','#ff9ff3'];

  const ctx2d = canvas.getContext('2d');
  window.managerProjectChart = new Chart(ctx2d, {
    type: 'doughnut',
    data: {
      labels: projectStats.map(p => `${p.name} (${p.count} employees)`),
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
        title: { display: true, text: 'All Projects Performance Overview', font: { size: 16, weight: 'bold' } },
        legend: { position: 'bottom', labels: { padding: 20, usePointStyle: true } },
        tooltip: {
          callbacks: {
            label: (context) => {
              const p = projectStats[context.dataIndex];
              return [
                `${p.name}`,
                `Employees: ${p.count}`,
                `Performance: ${p.performance}%`,
                `Actual: ${p.totalActual.toLocaleString()}`,
                `Target: ${p.totalTarget.toLocaleString()}`
              ];
            }
          }
        }
      }
    }
  });
}

function createManagerTeamChart() {
  const container = document.getElementById('managerTeamChart');
  const canvas = document.getElementById('managerTeamCanvas');
  console.log('Manager team chart - container:', container, 'canvas:', canvas);
  if (!container || !canvas) { 
    console.error('Manager team chart elements not found - container:', !!container, 'canvas:', !!canvas); 
    return; 
  }

  if (typeof Chart === 'undefined') {
    container.innerHTML = '<p style="color:red;padding:20px;">Chart.js not loaded.</p>';
    return;
  }

  if (window.managerTeamChart && typeof window.managerTeamChart.destroy === 'function') {
    try { window.managerTeamChart.destroy(); } catch(e) { console.warn('team destroy failed', e); }
  }

  const teams = [...new Set(productionData.map(d => d.team))];
  if (teams.length === 0) {
    container.innerHTML = '<p style="color:#666;padding:20px;text-align:center;">No team data available</p>';
    return;
  }

  const teamStats = teams.map(team => {
    const teamData = productionData.filter(d => d.team === team);
    const totalTarget = teamData.reduce((s,d)=>s+(d.target||0),0);
    const totalActual = teamData.reduce((s,d)=>s+(d.actual||0),0);
    const totalErrors = teamData.reduce((s,d)=>s+(d.errors||0),0);
    const totalLeaves = teamData.reduce((s,d)=>s+(d.leaves||0),0);
    const performance = totalTarget>0 ? Math.round((totalActual/totalTarget)*100) : 0;
    return { name: team, count: teamData.length, performance, totalActual, totalTarget, totalErrors, totalLeaves };
  });

  const ctx2d = canvas.getContext('2d');
  window.managerTeamChart = new Chart(ctx2d, {
    type: 'bar',
    data: {
      labels: teamStats.map(t => `${t.name} (${t.count} members)`),
      datasets: [{
        label: 'Team Performance %',
        data: teamStats.map(t => t.performance),
        backgroundColor: teamStats.map((_, i) => {
          const colors = ['#667eea','#764ba2','#f093fb','#4ecdc4','#45b7d1'];
          return colors[i % colors.length];
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
        title: { display: true, text: 'All Teams Performance Overview', font: { size: 16, weight: 'bold' } },
        legend: { display: false },
        tooltip: {
          callbacks: {
            label: (c) => {
              const t = teamStats[c.dataIndex];
              return [
                `${t.name}`,
                `Members: ${t.count}`,
                `Performance: ${t.performance}%`,
                `Actual: ${t.totalActual.toLocaleString()}`,
                `Target: ${t.totalTarget.toLocaleString()}`,
                `Errors: ${t.totalErrors}`,
                `Leaves: ${t.totalLeaves}`
              ];
            }
          }
        }
      }
    }
  });
}

function createManagerIDCards() {
    const metrics = calculateMetrics();
    const container = document.getElementById('managerTopPerformers');
    container.innerHTML = '';
    
    // Group by project
    const projectGroups = {};
    metrics.forEach(user => {
        if (!projectGroups[user.project]) {
            projectGroups[user.project] = [];
        }
        projectGroups[user.project].push(user);
    });
    
    // Create ID cards for each project's top performer (max 6 cards)
    const projectNames = Object.keys(projectGroups).slice(0, 6);
    projectNames.forEach(projectName => {
        const projectUsers = projectGroups[projectName];
        const topPerformer = projectUsers[0]; // Already sorted by stack ranking points
        
        // before appending, build a stable id
        const cardId = `tp-${projectName.replace(/\W+/g,'-').toLowerCase()}-${topPerformer.userId}`;
        if (document.getElementById(cardId)) {
          // already exists (defensive)
          return;
        }
        
        const card = document.createElement('div');
        card.id = cardId;                        // NEW
        card.className = 'manager-id-card';
        card.innerHTML = `
            <div class="card-header">
                <h4>${topPerformer.name}</h4>
                <div class="medal-icon">🥇</div>
            </div>
            <div class="card-body">
                <div class="employee-info">
                    <div class="info-row">
                        <span class="info-label">Employee ID:</span>
                        <span class="info-value">${topPerformer.userId}</span>
                    </div>
                    <div class="info-row">
                        <span class="info-label">Project Name:</span>
                        <span class="info-value">${topPerformer.project}</span>
                    </div>
                    <div class="info-row">
                        <span class="info-label">Process:</span>
                        <span class="info-value">${topPerformer.processName || 'N/A'}</span>
                    </div>
                    <div class="info-row">
                        <span class="info-label">Current Month Count:</span>
                        <span class="info-value">${topPerformer.actual.toLocaleString()}</span>
                    </div>
                    <div class="info-row">
                        <span class="info-label">Utilisation Position:</span>
                        <span class="info-value rank-${topPerformer.utilisationPosition}">#${topPerformer.utilisationPosition}</span>
                    </div>
                    <div class="info-row">
                        <span class="info-label">Stack Ranking Position:</span>
                        <span class="info-value rank-${topPerformer.stackRankingPosition}">#${topPerformer.stackRankingPosition}</span>
                    </div>
                    <div class="info-row">
                        <span class="info-label">Top Performance Status:</span>
                        <span class="info-value performance-${topPerformer.performanceStatus}">${topPerformer.performanceStatus.toUpperCase()}</span>
                    </div>
                </div>
                <div class="performance-stats">
                    <div class="stat-item">
                        <span class="stat-label">Utilisation:</span>
                        <span class="stat-value">${topPerformer.utilisation}%</span>
                    </div>
                    <div class="stat-item">
                        <span class="stat-label">Stack Points:</span>
                        <span class="stat-value">${topPerformer.stackRankingPoints}</span>
                    </div>
                    <div class="stat-item">
                        <span class="stat-label">Team Performance:</span>
                        <span class="stat-value">${topPerformer.teamPerformance}%</span>
                    </div>
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
    
    // Display conversations (simplified and shorter)
    Object.values(conversations).forEach(conversation => {
        conversation.sort((a, b) => new Date(a.timestamp) - new Date(b.timestamp));
        
        const conversationDiv = document.createElement('div');
        conversationDiv.className = 'conversation-thread-simple';
        
        const otherUser = conversation[0].from === currentUser.email ? conversation[0].to : conversation[0].from;
        const lastMessage = conversation[conversation.length - 1];
        const messageCount = conversation.length;
        
        // Show only the last message and count
        conversationDiv.innerHTML = `
            <div class="conversation-summary">
                <div class="conversation-header-simple">
                    <strong>${otherUser}</strong>
                    <span class="message-count">${messageCount} messages</span>
                </div>
                <div class="last-message">
                    <span class="sender">${lastMessage.from === currentUser.email ? 'You' : 'Them'}:</span>
                    <span class="message-text">${lastMessage.message.length > 50 ? lastMessage.message.substring(0, 50) + '...' : lastMessage.message}</span>
                </div>
                <div class="conversation-time">
                    <small>${new Date(lastMessage.timestamp).toLocaleDateString()}</small>
                </div>
            </div>
        `;
        
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
    const input = document.getElementById('googleSheetUrl');
    const url = (input && input.value || '').trim();
    if (!url) {
        alert('Please paste a valid Google Sheet link.');
        return;
    }
    localStorage.setItem('dm2_sheet_url', url);
    const hint = document.getElementById('savedSheetUrlHint');
    if (hint) hint.textContent = `Saved: ${url}`;
    alert('✅ Google Sheet link saved. Click Connect Google Sheets to load data.');
}

function openGoogleSheets() {
    const userUrl = localStorage.getItem('dm2_sheet_url');
    if (!userUrl) {
        alert('Please paste and save a Google Sheet link first.');
        return;
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
    
    const gvizUrl = `https://docs.google.com/spreadsheets/d/${sheetId}/gviz/tq?tqx=out:json&gid=${gid}`;
    console.log('Fetching from:', gvizUrl);
    
    // Try to fetch the data
    fetch(gvizUrl)
        .then(response => {
            if (!response.ok) {
                throw new Error(`HTTP ${response.status}: ${response.statusText}`);
            }
            return response.text();
        })
        .then(text => {
            console.log('Raw response received, length:', text.length);
            
            // Parse the gviz response
            const cleaned = text
                .replace(/^\)]}'\n?/, '')
                .replace(/^.*setResponse\(/, '')
                .replace(/\);?\s*$/, '');
            
            const json = JSON.parse(cleaned);
            const cols = (json.table.cols || []).map(c => (c && c.label) ? c.label : '');
            const rows = (json.table.rows || []).map(r => (r.c || []).map(c => c ? (c.v ?? '') : ''));
            
            console.log('Columns:', cols);
            console.log('Rows:', rows.length);
            
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
            loadProductionData();
            alert('✅ Successfully loaded ' + mapped.length + ' records from Google Sheets!');
        })
        .catch(err => {
            console.error('Google Sheets fetch failed:', err);
            alert('❌ Failed to load Google Sheets data:\n\n' + err.message + '\n\nPlease ensure:\n1. Sheet is publicly accessible\n2. URL is correct\n3. Sheet has data\n\nUsing sample data instead.');
            loadProductionData();
        })
        .finally(() => {
            if (connectBtn) {
                connectBtn.innerHTML = '<i class="fab fa-google"></i> Connect Google Sheets';
                connectBtn.disabled = false;
            }
        });
}

function openOtherSources() {
    alert('Other data sources integration would be implemented here. For demo purposes, sample data is loaded.');
    loadProductionData();
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
                    <input type="month" id="selectedMonth" value="">
                </div>
            `;
            break;
        case 'date':
            dateInputs.innerHTML = `
                <div class="form-group">
                    <label for="startDate">Start Date:</label>
                    <input type="date" id="startDate" value="">
                </div>
                <div class="form-group">
                    <label for="endDate">End Date:</label>
                    <input type="date" id="endDate" value="">
                </div>
            `;
            break;
        case 'year':
            dateInputs.innerHTML = `
                <div class="form-group">
                    <label for="selectedYear">Year:</label>
                    <input type="number" id="selectedYear" value="" min="2020" max="2030" placeholder="Enter year">
                </div>
            `;
            break;
        default:
            dateInputs.innerHTML = '<p style="color: #666; font-style: italic;">Please select a period type above</p>';
            break;
    }
}

function loadProductionData() {
    console.log('Loading production data...');
    
    // Check if Chart.js is loaded
    if (typeof Chart === 'undefined') {
        console.error('Chart.js library not loaded. Please check internet connection.');
        alert('Chart.js library not loaded. Please check your internet connection and refresh the page.');
        return;
    }
    
    // Update the global productionData with the latest data
    const savedData = localStorage.getItem('dm2_production_data');
    if (savedData) {
        productionData = JSON.parse(savedData);
        console.log('Loaded data from localStorage:', productionData.length, 'records');
    }
    
    // Fallback if still empty
    if (!Array.isArray(productionData) || productionData.length === 0) {
        productionData = [...sampleProductionData];
        localStorage.setItem('dm2_production_data', JSON.stringify(productionData));
        console.log('Using sample data:', productionData.length, 'records');
    }
    
    // Calculate metrics with the updated data
    const metrics = calculateMetrics();
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

function calculateMetrics() {
    // Group users by team to check team size
    const teamGroups = {};
    (productionData || []).forEach(user => {
        if (!teamGroups[user.team]) {
            teamGroups[user.team] = [];
        }
        teamGroups[user.team].push(user);
    });
    
    // Process each team
    let allMetrics = [];
    
    Object.keys(teamGroups).forEach(teamName => {
        const teamUsers = teamGroups[teamName];
        
        // If team has less than 5 members, combine with other small teams
        if (teamUsers.length < 5) {
            // For demo, we'll treat small teams as individual teams
            // In real implementation, you'd combine small teams here
        }
        
        const teamMetrics = teamUsers.map(user => {
            // 1. Utilisation Formula: user total count / working days * 100
            const workingDays = Number(user.workingDays) || 22;
            const actual = Number(user.actual) || 0;
            const target = Number(user.target) || 0;
            const errors = Number(user.errors) || 0;
            const leaves = Number(user.leaves) || 0;
            const utilisation = workingDays > 0 ? (actual / workingDays) * 100 : 0;
            
            // 2. Stack Ranking Points Calculation
            let stackRankingPoints = 0;
            
            // Target Achievement Points (5, 4, 3, 2, 1)
            if (target > 0 && actual >= target) {
                const excessRatio = (actual - target) / target;
                if (excessRatio >= 0.3) stackRankingPoints += 5; // 30%+ above target
                else if (excessRatio >= 0.2) stackRankingPoints += 4; // 20-29% above target
                else if (excessRatio >= 0.1) stackRankingPoints += 3; // 10-19% above target
                else stackRankingPoints += 2; // 0-9% above target
            } else {
                stackRankingPoints += 1; // Below target
            }
            
            // Error Points (5, 4, 3, 2, 1)
            if (errors === 0) stackRankingPoints += 5;
            else if (errors >= 1 && errors <= 3) stackRankingPoints += 4;
            else if (errors >= 4 && errors <= 6) stackRankingPoints += 3;
            else if (errors >= 7 && errors <= 9) stackRankingPoints += 2;
            else if (errors >= 10) stackRankingPoints += 1;
            
            // Leave Points (2, 1, -1, -2)
            if (leaves === 0) stackRankingPoints += 2;
            else if (leaves >= 1 && leaves <= 2) stackRankingPoints += 1;
            else if (leaves === 3) stackRankingPoints += -1;
            else if (leaves >= 4) stackRankingPoints += -2;
            
            // 3. Team Performance (based on volume vs target)
            const teamPerformance = target > 0 ? (actual / target) * 100 : 0;
            
            return {
                ...user,
                utilisation: Math.round(utilisation * 100) / 100,
                stackRankingPoints: stackRankingPoints,
                teamPerformance: Math.round(teamPerformance * 100) / 100,
                performanceStatus: teamPerformance >= 100 ? 'excellent' : 
                                 teamPerformance >= 80 ? 'good' : 'poor'
            };
        });
        
        // Sort team by stack ranking points for ranking
        teamMetrics.sort((a, b) => b.stackRankingPoints - a.stackRankingPoints);
        
        // Add ranking positions
        teamMetrics.forEach((user, index) => {
            user.stackRankingPosition = index + 1;
            user.utilisationPosition = teamMetrics
                .sort((a, b) => b.utilisation - a.utilisation)
                .findIndex(u => u.userId === user.userId) + 1;
            user.teamPerformancePosition = teamMetrics
                .sort((a, b) => b.teamPerformance - a.teamPerformance)
                .findIndex(u => u.userId === user.userId) + 1;
        });
        
        allMetrics = allMetrics.concat(teamMetrics);
    });
    
    return allMetrics;
}

// Fetch Google Sheet via gviz JSON (public or view-only)
async function fetchGoogleSheetData(gvizUrl) {
    // Try direct, then CORS proxy fallback
    let txt;
    try {
        const res = await fetch(gvizUrl, { cache: 'no-store', credentials: 'omit' });
        txt = await res.text();
        if (!res.ok) throw new Error('HTTP ' + res.status);
    } catch (e) {
        const proxied = 'https://cors.isomorphic-git.org/' + gvizUrl;
        const res2 = await fetch(proxied, { cache: 'no-store', credentials: 'omit' });
        txt = await res2.text();
        if (!res2.ok) throw new Error('HTTP ' + res2.status);
    }
    // gviz returns something like: google.visualization.Query.setResponse({...})
    const cleaned = txt
        .replace(/^\)]}'\n?/, '') // strip XSSI guard if present
        .replace(/^.*setResponse\(/, '')
        .replace(/\);?\s*$/, '');
    const json = JSON.parse(cleaned);
    const cols = (json.table.cols || []).map(c => (c && c.label) ? c.label : '');
    const rows = (json.table.rows || []).map(r => (r.c || []).map(c => c ? (c.v ?? '') : ''));
    // Map to our structure by header names
    const dateIdx = cols.findIndex(h => /Date/i.test(h));
    const idIdx = cols.findIndex(h => /(ITPL#|ITPL)/i.test(h));
    const nameIdx = cols.findIndex(h => /Name/i.test(h));
    const clientIdx = cols.findIndex(h => /(Client Name)/i.test(h));
    const processIdx = cols.findIndex(h => /Process Name/i.test(h));
    const prodIdx = cols.findIndex(h => /Productivity/i.test(h));
    const targetIdx = cols.findIndex(h => /Target/i.test(h));
    const clientErrIdx = cols.findIndex(h => /Client Errors/i.test(h));
    const internalErrIdx = cols.findIndex(h => /Internal Errors/i.test(h));

    const mapped = rows
        .filter(r => r[idIdx] && r[nameIdx])
        .map(r => {
            const errors = (Number(r[clientErrIdx]) || 0) + (Number(r[internalErrIdx]) || 0);
            return {
                userId: String(r[idIdx]).trim(),
                name: String(r[nameIdx]).trim(),
                project: String(r[clientIdx] || 'NA').trim(),
                processName: String(r[processIdx] || '').trim(),
                actual: Number(r[prodIdx]) || 0,
                target: Number(r[targetIdx]) || 0,
                errors: errors,
                leaves: 0,
                month: (r[dateIdx] ? String(r[dateIdx]) : ''),
                workingDays: 22,
                team: 'Team A'
            };
        });
    return mapped;
}

function parseSheetUrl(url) {
    try {
        console.log('Parsing URL:', url);
        
        // Extract sheet ID using regex - more reliable
        const sheetIdMatch = url.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
        const sheetId = sheetIdMatch ? sheetIdMatch[1] : '';
        
        // Extract gid using regex
        const gidMatch = url.match(/[#&]gid=(\d+)/);
        const gid = gidMatch ? gidMatch[1] : '0';
        
        console.log('Extracted sheetId:', sheetId, 'gid:', gid);
        
        return { sheetId, gid };
    } catch (error) {
        console.error('Error parsing sheet URL:', error);
        return { sheetId: '', gid: '0' };
    }
}

function mapGoogleSheetRows(rows, cols) {
    console.log('Mapping Google Sheets data...');
    console.log('Columns:', cols);
    
    // Find column indices
    const findColIndex = (patterns) => {
        for (const pattern of patterns) {
            const index = cols.findIndex(col => 
                col && col.toLowerCase().includes(pattern.toLowerCase())
            );
            if (index !== -1) return index;
        }
        return -1;
    };
    
    const idIdx = findColIndex(['itpl', 'employee id', 'emp id', 'id']);
    const nameIdx = findColIndex(['name', 'employee name']);
    const clientIdx = findColIndex(['client name', 'project', 'client']);
    const processIdx = findColIndex(['process name', 'process']);
    const prodIdx = findColIndex(['productivity', 'actual', 'count', 'volume']);
    const targetIdx = findColIndex(['target', 'monthly target']);
    const clientErrIdx = findColIndex(['client errors', 'client error']);
    const internalErrIdx = findColIndex(['internal errors', 'internal error']);
    const leavesIdx = findColIndex(['leaves', 'leave days']);
    const workingDaysIdx = findColIndex(['working days']);
    const dateIdx = findColIndex(['date', 'month']);
    
    console.log('Column indices:', {
        id: idIdx, name: nameIdx, client: clientIdx, process: processIdx,
        prod: prodIdx, target: targetIdx, clientErr: clientErrIdx,
        internalErr: internalErrIdx, leaves: leavesIdx, workingDays: workingDaysIdx, date: dateIdx
    });
    
    const mapped = rows
        .filter((row, index) => {
            // Skip empty rows or header row
            if (index === 0) return false;
            return row[idIdx] && row[nameIdx];
        })
        .map(row => {
            const toNum = (val) => {
                if (!val) return 0;
                const num = Number(val);
                return isNaN(num) ? 0 : num;
            };
            
            const errors = toNum(row[clientErrIdx]) + toNum(row[internalErrIdx]);
            
            return {
                userId: String(row[idIdx] || '').trim(),
                name: String(row[nameIdx] || '').trim(),
                project: String(row[clientIdx] || 'NA').trim(),
                processName: String(row[processIdx] || '').trim(),
                actual: toNum(row[prodIdx]),
                target: toNum(row[targetIdx]),
                errors: errors,
                leaves: toNum(row[leavesIdx]),
                workingDays: toNum(row[workingDaysIdx]) || 22,
                month: row[dateIdx] ? String(row[dateIdx]) : '',
                team: 'Team A'
            };
        })
        .filter(record => record.userId && record.name);
    
    console.log('Mapped records:', mapped.length);
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
        ctx.innerHTML = '<p style="color: red; padding: 20px;">Chart.js library not loaded. Please check internet connection.</p>';
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
    
    try {
        window.utilisationChart = new Chart(ctx, {
            type: 'bar',
            data: {
                labels: metrics.map(m => `${m.name} (${m.userId})`),
                datasets: [{
                    label: 'Utilisation %',
                    data: metrics.map(m => m.utilisation),
                    backgroundColor: metrics.map(m => 
                        m.utilisation >= 300 ? '#28a745' : 
                        m.utilisation >= 250 ? '#ffc107' : '#dc3545'
                    ),
                    borderColor: metrics.map(m => 
                        m.utilisation >= 300 ? '#1e7e34' : 
                        m.utilisation >= 250 ? '#e0a800' : '#bd2130'
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
                        text: 'Employee Utilisation Analysis (Count/Working Days * 100)'
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
    
    try {
        window.stackRankingChart = new Chart(ctx, {
            type: 'bar',
            data: {
                labels: metrics.map(m => `${m.name} (${m.userId})`),
                datasets: [{
                    label: 'Stack Ranking Points',
                    data: metrics.map(m => m.stackRankingPoints),
                    backgroundColor: metrics.map((m, index) => {
                        if (index === 0) return '#ffd700'; // Gold for 1st
                        if (index === 1) return '#c0c0c0'; // Silver for 2nd
                        if (index === 2) return '#cd7f32'; // Bronze for 3rd
                        return '#667eea';
                    }),
                    borderColor: metrics.map((m, index) => {
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
                        text: 'Stack Ranking (Target + Error + Leave Points)'
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
    
    try {
        window.userPerformanceChart = new Chart(ctx, {
            type: 'line',
            data: {
                labels: metrics.map(m => `${m.name} (${m.userId})`),
                datasets: [{
                    label: 'Team Performance %',
                    data: metrics.map(m => m.teamPerformance),
                    borderColor: '#667eea',
                    backgroundColor: 'rgba(102, 126, 234, 0.1)',
                    tension: 0.4,
                    fill: true,
                    borderWidth: 3,
                    pointBackgroundColor: metrics.map(m => 
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
    
    // Show top 3 performers
    const topPerformers = metrics.slice(0, 3);
    
    topPerformers.forEach((user, index) => {
        const card = document.createElement('div');
        card.className = 'user-card';
        card.innerHTML = `
            <h4>${index === 0 ? '🥇' : index === 1 ? '🥈' : '🥉'} ${user.name}</h4>
            <div class="user-info">
                <div class="info-row">
                    <span class="info-label">Employee ID:</span>
                    <span class="info-value">${user.userId}</span>
                </div>
                <div class="info-row">
                    <span class="info-label">Project:</span>
                    <span class="info-value">${user.project}</span>
                </div>
                <div class="info-row">
                    <span class="info-label">Process:</span>
                    <span class="info-value">${user.processName || 'N/A'}</span>
                </div>
                <div class="info-row">
                    <span class="info-label">Current Count:</span>
                    <span class="info-value">${user.actual.toLocaleString()}</span>
                </div>
                <div class="info-row">
                    <span class="info-label">Target:</span>
                    <span class="info-value">${user.target.toLocaleString()}</span>
                </div>
                <div class="info-row">
                    <span class="info-label">Utilisation:</span>
                    <span class="info-value">${user.utilisation}%</span>
                </div>
                <div class="info-row">
                    <span class="info-label">Stack Ranking:</span>
                    <span class="info-value">#${user.stackRankingPosition} (${user.stackRankingPoints} pts)</span>
                </div>
                <div class="info-row">
                    <span class="info-label">Team Performance:</span>
                    <span class="info-value performance-${user.performanceStatus}">${user.teamPerformance}%</span>
                </div>
                <div class="info-row">
                    <span class="info-label">Errors:</span>
                    <span class="info-value">${user.errors}</span>
                </div>
                <div class="info-row">
                    <span class="info-label">Leaves:</span>
                    <span class="info-value">${user.leaves} days</span>
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

