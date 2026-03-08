<script>
  let currentData = [];
  let currentHeaders = [];

  // Initialize the app
  window.onload = function() {
    loadDashboard();
  };

  // Load dashboard view
  function loadDashboard() {
    setActiveNav('dashboard');
    document.querySelector('.page-title h1').textContent = 'Dashboard';
    
    google.script.run
      .withSuccessHandler(displayDashboard)
      .withFailureHandler(showError)
      .getSheetData();
  }

  // Load data view
  function loadDataView() {
    setActiveNav('data');
    document.querySelector('.page-title h1').textContent = 'Data View';
    
    google.script.run
      .withSuccessHandler(displayDataTable)
      .withFailureHandler(showError)
      .getSheetData();
  }

  // Load analytics view
  function loadAnalytics() {
    setActiveNav('analytics');
    document.querySelector('.page-title h1').textContent = 'Analytics';
    
    google.script.run
      .withSuccessHandler(displayAnalytics)
      .withFailureHandler(showError)
      .getSheetData();
  }

  // Load settings view
  function loadSettings() {
    setActiveNav('settings');
    document.querySelector('.page-title h1').textContent = 'Settings';
    document.getElementById('content-area').innerHTML = `
      <div class="data-table-container">
        <h2 style="margin-bottom: 20px;">Settings</h2>
        <div class="form-group">
          <label>Refresh Interval (seconds)</label>
          <input type="number" value="30" id="refreshInterval">
        </div>
        <button class="btn btn-primary" onclick="saveSettings()">Save Settings</button>
      </div>
    `;
  }

  // Display dashboard with statistics
  function displayDashboard(result) {
    if (!result.success) {
      showError(result.error);
      return;
    }

    currentData = result.data;
    currentHeaders = result.headers;

    const stats = calculateStats(result.data);
    
    let html = `
      <div class="dashboard-grid">
        <div class="stat-card">
          <div class="stat-header">
            <i class="fas fa-database"></i>
            <span class="badge badge-success">Total Records</span>
          </div>
          <div class="stat-value">${result.data.length}</div>
          <div class="stat-label">Records in database</div>
        </div>
        
        <div class="stat-card">
          <div class="stat-header">
            <i class="fas fa-calendar"></i>
            <span class="badge badge-warning">Last Updated</span>
          </div>
          <div class="stat-value">${new Date().toLocaleDateString()}</div>
          <div class="stat-label">Data freshness</div>
        </div>
        
        <div class="stat-card">
          <div class="stat-header">
            <i class="fas fa-chart-line"></i>
            <span class="badge badge-success">Growth</span>
          </div>
          <div class="stat-value">+12%</div>
          <div class="stat-label">vs last month</div>
        </div>
      </div>
      
      <div class="data-table-container">
        <h3 style="margin-bottom: 16px;">Recent Data</h3>
        <table class="data-table">
          <thead>
            <tr>
              ${result.headers.map(header => `<th>${header}</th>`).join('')}
            </tr>
          </thead>
          <tbody>
            ${result.data.slice(0, 5).map(row => `
              <tr>
                ${result.headers.map(header => `<td>${row[header] || ''}</td>`).join('')}
              </tr>
            `).join('')}
          </tbody>
        </table>
      </div>
    `;

    document.getElementById('content-area').innerHTML = html;
  }

  // Display data table
  function displayDataTable(result) {
    if (!result.success) {
      showError(result.error);
      return;
    }

    currentData = result.data;
    currentHeaders = result.headers;

    let html = `
      <div class="data-table-container">
        <div style="margin-bottom: 20px; display: flex; gap: 10px;">
          <input type="text" 
                 placeholder="Search..." 
                 class="search-input" 
                 onkeyup="filterData(this.value)"
                 style="flex: 1; padding: 8px; border: 1px solid #dee2e6; border-radius: 6px;">
          <button class="btn btn-icon" onclick="exportToCSV()">
            <i class="fas fa-file-csv"></i> Export CSV
          </button>
        </div>
        
        <table class="data-table" id="data-table">
          <thead>
            <tr>
              ${result.headers.map(header => `<th onclick="sortData('${header}')" style="cursor: pointer;">
                ${header} <i class="fas fa-sort"></i>
              </th>`).join('')}
            </tr>
          </thead>
          <tbody id="table-body">
            ${renderTableRows(result.data, result.headers)}
          </tbody>
        </table>
        
        <div style="margin-top: 20px; display: flex; justify-content: space-between; align-items: center;">
          <span>Showing ${result.data.length} records</span>
          <div>
            <button class="btn btn-icon" onclick="previousPage()"><i class="fas fa-chevron-left"></i></button>
            <span id="page-info">Page 1</span>
            <button class="btn btn-icon" onclick="nextPage()"><i class="fas fa-chevron-right"></i></button>
          </div>
        </div>
      </div>
    `;

    document.getElementById('content-area').innerHTML = html;
  }

  // Display analytics
  function displayAnalytics(result) {
    if (!result.success) {
      showError(result.error);
      return;
    }

    let html = `
      <div class="dashboard-grid">
        <div class="stat-card">
          <div class="stat-header">
            <i class="fas fa-chart-pie"></i>
          </div>
          <div class="stat-value">${result.data.length}</div>
          <div class="stat-label">Total Records</div>
        </div>
        
        <div class="stat-card">
          <div class="stat-header">
            <i class="fas fa-chart-bar"></i>
          </div>
          <div class="stat-value">${result.headers.length}</div>
          <div class="stat-label">Total Fields</div>
        </div>
      </div>
      
      <div class="data-table-container">
        <h3>Data Summary</h3>
        <table class="data-table">
          <thead>
            <tr>
              <th>Field Name</th>
              <th>Data Type</th>
              <th>Unique Values</th>
              <th>Completion Rate</th>
            </tr>
          </thead>
          <tbody>
            ${result.headers.map(header => `
              <tr>
                <td><strong>${header}</strong></td>
                <td>${inferDataType(result.data, header)}</td>
                <td>${countUniqueValues(result.data, header)}</td>
                <td>${calculateCompletion(result.data, header)}%</td>
              </tr>
            `).join('')}
          </tbody>
        </table>
      </div>
    `;

    document.getElementById('content-area').innerHTML = html;
  }

  // Helper Functions
  function renderTableRows(data, headers) {
    return data.map(row => `
      <tr>
        ${headers.map(header => `<td>${row[header] || ''}</td>`).join('')}
      </tr>
    `).join('');
  }

  function calculateStats(data) {
    // Add your custom statistics calculation here
    return {};
  }

  function inferDataType(data, field) {
    const sample = data.find(row => row[field])?.[field];
    if (!sample) return 'Unknown';
    if (!isNaN(sample) && sample.toString().trim() !== '') return 'Number';
    if (Date.parse(sample)) return 'Date';
    return 'Text';
  }

  function countUniqueValues(data, field) {
    const values = new Set(data.map(row => row[field]).filter(v => v));
    return values.size;
  }

  function calculateCompletion(data, field) {
    const filled = data.filter(row => row[field] && row[field].toString().trim() !== '').length;
    return Math.round((filled / data.length) * 100);
  }

  function filterData(searchTerm) {
    const filtered = currentData.filter(row => {
      return Object.values(row).some(value => 
        value && value.toString().toLowerCase().includes(searchTerm.toLowerCase())
      );
    });
    
    document.getElementById('table-body').innerHTML = 
      renderTableRows(filtered, currentHeaders);
  }

  function sortData(header) {
    const sorted = [...currentData].sort((a, b) => {
      if (a[header] < b[header]) return -1;
      if (a[header] > b[header]) return 1;
      return 0;
    });
    
    document.getElementById('table-body').innerHTML = 
      renderTableRows(sorted, currentHeaders);
  }

  function exportToCSV() {
    let csv = currentHeaders.join(',') + '\n';
    csv += currentData.map(row => 
      currentHeaders.map(header => 
        `"${(row[header] || '').toString().replace(/"/g, '""')}"`
      ).join(',')
    ).join('\n');
    
    const blob = new Blob([csv], { type: 'text/csv' });
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'data_export.csv';
    a.click();
  }

  function refreshData() {
    document.getElementById('content-area').innerHTML = 
      '<div class="loading-spinner"><i class="fas fa-spinner fa-spin"></i> Refreshing...</div>';
    
    google.script.run
      .withSuccessHandler(displayDataTable)
      .withFailureHandler(showError)
      .getSheetData();
  }

  function showAddDataModal() {
    const modal = document.getElementById('dataModal');
    const modalBody = document.getElementById('modal-body');
    
    let formHtml = '<form onsubmit="submitData(event)">';
    currentHeaders.forEach(header => {
      formHtml += `
        <div class="form-group">
          <label>${header}</label>
          <input type="text" name="${header}" required>
        </div>
      `;
    });
    formHtml += `
      <button type="submit" class="btn btn-primary">Submit</button>
    </form>`;
    
    modalBody.innerHTML = formHtml;
    modal.style.display = 'block';
  }

  function submitData(event) {
    event.preventDefault();
    // Implement your data submission logic here
    alert('Data submitted successfully!');
    closeModal();
    refreshData();
  }

  function closeModal() {
    document.getElementById('dataModal').style.display = 'none';
  }

  function showError(error) {
    document.getElementById('content-area').innerHTML = `
      <div style="background: #ffebee; color: #f44336; padding: 20px; border-radius: 6px;">
        <i class="fas fa-exclamation-circle"></i> Error: ${error}
      </div>
    `;
  }

  function setActiveNav(page) {
    document.querySelectorAll('.nav-item').forEach(item => {
      item.classList.remove('active');
    });
    
    const navMap = {
      'dashboard': 0,
      'data': 1,
      'analytics': 2,
      'settings': 3
    };
    
    const navItems = document.querySelectorAll('.nav-item');
    if (navItems[navMap[page]]) {
      navItems[navMap[page]].classList.add('active');
    }
  }

  // Close modal when clicking outside
  window.onclick = function(event) {
    const modal = document.getElementById('dataModal');
    if (event.target === modal) {
      modal.style.display = 'none';
    }
  };
</script>
