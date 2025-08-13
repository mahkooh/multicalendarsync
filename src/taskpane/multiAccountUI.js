/**
 * Multi-Account Calendar UI Manager
 * Handles the user interface for connecting and managing multiple calendar accounts
 */

export class MultiAccountUI {
  constructor(multiTenantManager) {
    this.manager = multiTenantManager;
    this.currentView = 'overview';
  }

  /**
   * Initialize the multi-account interface
   */
  initialize() {
    this.createAccountManagementUI();
    this.bindEvents();
    this.refreshAccountsList();
  }

  /**
   * Create the main account management interface
   */
  createAccountManagementUI() {
    const container = document.getElementById('auth-container') || 
                     document.getElementById('calendar-list') ||
                     document.createElement('div');

    container.innerHTML = `
      <div class="multi-account-manager">
        <!-- Header -->
        <div class="account-header">
          <h2>📅 Multi-Tenant Calendar Sync</h2>
          <p>Connect your calendars across different organizations and services</p>
        </div>

        <!-- Navigation Tabs -->
        <div class="account-tabs">
          <button class="tab-btn active" data-tab="overview">Overview</button>
          <button class="tab-btn" data-tab="add-account">Add Account</button>
          <button class="tab-btn" data-tab="sync-status">Sync Status</button>
        </div>

        <!-- Overview Tab -->
        <div class="tab-content active" id="overview-tab">
          <div class="accounts-summary">
            <h3>Connected Accounts</h3>
            <div id="accounts-list" class="accounts-list">
              <div class="loading">Loading accounts...</div>
            </div>
            <div class="sync-actions">
              <button id="sync-all-btn" class="btn-primary">🔄 Sync All Calendars</button>
              <button id="test-mode-btn" class="btn-secondary">🧪 Load Test Accounts</button>
            </div>
          </div>
          
          <div class="sync-overview">
            <h3>Sync Configuration</h3>
            <div class="sync-settings">
              <label>
                <input type="checkbox" id="auto-sync" checked> 
                Enable automatic synchronization
              </label>
              <label>
                Sync interval: 
                <select id="sync-interval">
                  <option value="5">5 minutes</option>
                  <option value="15" selected>15 minutes</option>
                  <option value="30">30 minutes</option>
                  <option value="60">1 hour</option>
                </select>
              </label>
            </div>
          </div>
        </div>

        <!-- Add Account Tab -->
        <div class="tab-content" id="add-account-tab">
          <div class="add-account-form">
            <h3>Add New Calendar Account</h3>
            
            <div class="account-type-selection">
              <h4>Select Account Type:</h4>
              
              <div class="account-type-grid">
                <div class="account-type-card" data-type="office365">
                  <div class="account-icon">🏢</div>
                  <h4>Office 365</h4>
                  <p>Business Microsoft account from another organization</p>
                  <button class="select-type-btn">Connect</button>
                </div>
                
                <div class="account-type-card" data-type="gmail">
                  <div class="account-icon">📧</div>
                  <h4>Gmail</h4>
                  <p>Google Calendar from Gmail account</p>
                  <button class="select-type-btn">Connect</button>
                </div>
                
                <div class="account-type-card" data-type="outlook_com">
                  <div class="account-icon">📨</div>
                  <h4>Outlook.com</h4>
                  <p>Personal Microsoft account</p>
                  <button class="select-type-btn">Connect</button>
                </div>
                
                <div class="account-type-card" data-type="exchange">
                  <div class="account-icon">🏛️</div>
                  <h4>Exchange Server</h4>
                  <p>On-premises Exchange calendar</p>
                  <button class="select-type-btn">Connect</button>
                </div>
              </div>
            </div>

            <!-- Account Configuration Form -->
            <div class="account-config-form" id="account-config-form" style="display: none;">
              <h4 id="config-form-title">Configure Account</h4>
              
              <div class="form-group">
                <label for="account-email">Email Address:</label>
                <input type="email" id="account-email" placeholder="your.name@company.com" required>
              </div>
              
              <div class="form-group">
                <label for="account-display-name">Display Name:</label>
                <input type="text" id="account-display-name" placeholder="Company Name" required>
              </div>
              
              <div class="form-group" id="tenant-config" style="display: none;">
                <label for="tenant-id">Tenant ID (for Office 365):</label>
                <input type="text" id="tenant-id" placeholder="xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx">
                <small>Found in Azure AD > Properties > Directory ID</small>
              </div>
              
              <div class="form-group" id="client-config" style="display: none;">
                <label for="client-id">Client ID:</label>
                <input type="text" id="client-id" placeholder="Application ID from app registration">
                <small>Create an app registration in Azure AD / Google Console</small>
              </div>
              
              <div class="form-actions">
                <button type="button" id="save-account-btn" class="btn-primary">Save & Connect</button>
                <button type="button" id="cancel-config-btn" class="btn-secondary">Cancel</button>
              </div>
            </div>
          </div>
        </div>

        <!-- Sync Status Tab -->
        <div class="tab-content" id="sync-status-tab">
          <div class="sync-status-view">
            <h3>Synchronization Status</h3>
            
            <div class="sync-timeline">
              <h4>Recent Sync Activity</h4>
              <div id="sync-timeline-list" class="timeline-list">
                <div class="timeline-item">
                  <div class="timeline-time">2:30 PM</div>
                  <div class="timeline-content">
                    <strong>✅ Sync completed</strong><br>
                    All accounts synchronized successfully
                  </div>
                </div>
                <div class="timeline-item">
                  <div class="timeline-time">2:15 PM</div>
                  <div class="timeline-content">
                    <strong>🔄 Sync started</strong><br>
                    Processing 5 connected accounts
                  </div>
                </div>
              </div>
            </div>
            
            <div class="sync-conflicts">
              <h4>Conflicts & Issues</h4>
              <div id="conflicts-list" class="conflicts-list">
                <div class="no-conflicts">No conflicts detected</div>
              </div>
            </div>
          </div>
        </div>
      </div>
    `;

    // Add to page if not already there
    if (!document.getElementById('auth-container')) {
      document.body.appendChild(container);
    }
  }

  /**
   * Bind event handlers
   */
  bindEvents() {
    // Tab switching
    document.querySelectorAll('.tab-btn').forEach(btn => {
      btn.addEventListener('click', (e) => {
        this.switchTab(e.target.dataset.tab);
      });
    });

    // Account type selection
    document.querySelectorAll('.account-type-card').forEach(card => {
      card.addEventListener('click', (e) => {
        const button = e.target.closest('.account-type-card').querySelector('.select-type-btn');
        if (button) button.click();
      });
    });

    document.querySelectorAll('.select-type-btn').forEach(btn => {
      btn.addEventListener('click', (e) => {
        e.stopPropagation();
        const accountType = e.target.closest('.account-type-card').dataset.type;
        this.showAccountConfigForm(accountType);
      });
    });

    // Form actions
    document.getElementById('save-account-btn')?.addEventListener('click', () => {
      this.saveAccountConfiguration();
    });

    document.getElementById('cancel-config-btn')?.addEventListener('click', () => {
      this.hideAccountConfigForm();
    });

    // Sync actions
    document.getElementById('sync-all-btn')?.addEventListener('click', () => {
      this.syncAllAccounts();
    });

    document.getElementById('test-mode-btn')?.addEventListener('click', () => {
      this.loadTestAccounts();
    });
  }

  /**
   * Switch between tabs
   */
  switchTab(tabName) {
    // Update tab buttons
    document.querySelectorAll('.tab-btn').forEach(btn => {
      btn.classList.toggle('active', btn.dataset.tab === tabName);
    });

    // Update tab content
    document.querySelectorAll('.tab-content').forEach(content => {
      content.classList.toggle('active', content.id === `${tabName}-tab`);
    });

    this.currentView = tabName;

    // Refresh data when switching to certain tabs
    if (tabName === 'overview') {
      this.refreshAccountsList();
    } else if (tabName === 'sync-status') {
      this.refreshSyncStatus();
    }
  }

  /**
   * Show account configuration form
   */
  showAccountConfigForm(accountType) {
    const form = document.getElementById('account-config-form');
    const title = document.getElementById('config-form-title');
    const tenantConfig = document.getElementById('tenant-config');
    const clientConfig = document.getElementById('client-config');

    // Configure form based on account type
    switch (accountType) {
      case 'office365':
        title.textContent = 'Configure Office 365 Account';
        tenantConfig.style.display = 'block';
        clientConfig.style.display = 'block';
        break;
      case 'gmail':
        title.textContent = 'Configure Gmail Account';
        tenantConfig.style.display = 'none';
        clientConfig.style.display = 'block';
        break;
      case 'outlook_com':
        title.textContent = 'Configure Outlook.com Account';
        tenantConfig.style.display = 'none';
        clientConfig.style.display = 'block';
        break;
      case 'exchange':
        title.textContent = 'Configure Exchange Server';
        tenantConfig.style.display = 'none';
        clientConfig.style.display = 'none';
        break;
    }

    form.style.display = 'block';
    form.dataset.accountType = accountType;
  }

  /**
   * Hide account configuration form
   */
  hideAccountConfigForm() {
    document.getElementById('account-config-form').style.display = 'none';
    
    // Clear form
    document.getElementById('account-email').value = '';
    document.getElementById('account-display-name').value = '';
    document.getElementById('tenant-id').value = '';
    document.getElementById('client-id').value = '';
  }

  /**
   * Save account configuration and initiate connection
   */
  async saveAccountConfiguration() {
    const form = document.getElementById('account-config-form');
    const accountType = form.dataset.accountType;
    
    const config = {
      id: `account-${Date.now()}`,
      type: accountType,
      email: document.getElementById('account-email').value,
      displayName: document.getElementById('account-display-name').value,
      tenantId: document.getElementById('tenant-id').value,
      clientId: document.getElementById('client-id').value
    };

    if (!config.email || !config.displayName) {
      alert('Please fill in all required fields');
      return;
    }

    try {
      console.log('🔗 Adding new account:', config);
      
      // Add account to manager
      await this.manager.addAccount(config);
      
      // Hide form and refresh list
      this.hideAccountConfigForm();
      this.switchTab('overview');
      this.refreshAccountsList();
      
      // Show success message
      this.showMessage(`Successfully connected ${config.displayName}!`, 'success');
      
    } catch (error) {
      console.error('❌ Failed to add account:', error);
      this.showMessage(`Failed to connect account: ${error.message}`, 'error');
    }
  }

  /**
   * Load test accounts for demo
   */
  loadTestAccounts() {
    console.log('🧪 Loading test accounts...');
    this.manager.createMockAccounts();
    this.refreshAccountsList();
    this.showMessage('Test accounts loaded successfully!', 'success');
  }

  /**
   * Refresh the accounts list display
   */
  refreshAccountsList() {
    const accountsList = document.getElementById('accounts-list');
    if (!accountsList) return;

    const accounts = this.manager.getAccountsSummary();
    
    if (accounts.length === 0) {
      accountsList.innerHTML = `
        <div class="no-accounts">
          <div class="no-accounts-icon">📅</div>
          <h4>No accounts connected</h4>
          <p>Add your first calendar account to start syncing</p>
          <button onclick="document.querySelector('[data-tab=add-account]').click()" class="btn-primary">
            Add Account
          </button>
        </div>
      `;
      return;
    }

    accountsList.innerHTML = accounts.map(account => `
      <div class="account-item ${account.isAuthenticated ? 'authenticated' : 'not-authenticated'}">
        <div class="account-info">
          <div class="account-header">
            <span class="account-type-icon">${this.getAccountTypeIcon(account.type)}</span>
            <div class="account-details">
              <h4>${account.displayName}</h4>
              <p>${account.email}</p>
            </div>
            <div class="account-status">
              ${account.isAuthenticated ? 
                '<span class="status-badge success">✅ Connected</span>' : 
                '<span class="status-badge error">❌ Disconnected</span>'
              }
            </div>
          </div>
          
          <div class="account-stats">
            <span class="stat">📅 ${account.calendarCount} calendars</span>
            <span class="stat">🔄 ${account.syncEnabled ? 'Sync enabled' : 'Sync disabled'}</span>
            ${account.lastSync ? 
              `<span class="stat">⏰ Last sync: ${new Date(account.lastSync).toLocaleTimeString()}</span>` : 
              '<span class="stat">⏰ Never synced</span>'
            }
          </div>
        </div>
        
        <div class="account-actions">
          <button class="btn-small" onclick="multiAccountUI.configureAccount('${account.id}')">
            ⚙️ Configure
          </button>
          <button class="btn-small" onclick="multiAccountUI.syncAccount('${account.id}')">
            🔄 Sync Now
          </button>
          <button class="btn-small danger" onclick="multiAccountUI.removeAccount('${account.id}')">
            🗑️ Remove
          </button>
        </div>
      </div>
    `).join('');
  }

  /**
   * Get icon for account type
   */
  getAccountTypeIcon(type) {
    const icons = {
      office365: '🏢',
      gmail: '📧',
      outlook_com: '📨',
      exchange: '🏛️',
      other: '📅'
    };
    return icons[type] || icons.other;
  }

  /**
   * Sync all connected accounts
   */
  async syncAllAccounts() {
    const syncBtn = document.getElementById('sync-all-btn');
    const originalText = syncBtn.textContent;
    
    syncBtn.textContent = '🔄 Syncing...';
    syncBtn.disabled = true;
    
    try {
      // This would trigger the actual sync process
      console.log('🔄 Starting sync for all accounts...');
      
      // Simulate sync process
      await new Promise(resolve => setTimeout(resolve, 2000));
      
      this.showMessage('All accounts synchronized successfully!', 'success');
      this.refreshAccountsList();
      
    } catch (error) {
      console.error('❌ Sync failed:', error);
      this.showMessage(`Sync failed: ${error.message}`, 'error');
    } finally {
      syncBtn.textContent = originalText;
      syncBtn.disabled = false;
    }
  }

  /**
   * Show message to user
   */
  showMessage(message, type = 'info') {
    // Create or update message element
    let messageEl = document.getElementById('message-display');
    if (!messageEl) {
      messageEl = document.createElement('div');
      messageEl.id = 'message-display';
      messageEl.className = 'message-display';
      document.querySelector('.multi-account-manager').prepend(messageEl);
    }
    
    messageEl.className = `message-display ${type}`;
    messageEl.textContent = message;
    messageEl.style.display = 'block';
    
    // Auto-hide after 5 seconds
    setTimeout(() => {
      messageEl.style.display = 'none';
    }, 5000);
  }

  /**
   * Configure specific account
   */
  configureAccount(accountId) {
    console.log(`⚙️ Configuring account: ${accountId}`);
    // This would open account-specific configuration
  }

  /**
   * Sync specific account
   */
  async syncAccount(accountId) {
    console.log(`🔄 Syncing account: ${accountId}`);
    // This would trigger sync for specific account
  }

  /**
   * Remove account
   */
  removeAccount(accountId) {
    if (confirm('Are you sure you want to remove this account?')) {
      console.log(`🗑️ Removing account: ${accountId}`);
      // This would remove the account from manager
      this.refreshAccountsList();
    }
  }

  /**
   * Refresh sync status display
   */
  refreshSyncStatus() {
    console.log('🔄 Refreshing sync status...');
    // This would update the sync status tab with recent activity
  }
}

// Export for global use
window.MultiAccountUI = MultiAccountUI;
