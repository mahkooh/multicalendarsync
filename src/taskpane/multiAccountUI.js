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
                <div class="account-type-card demo-card">
                  <div class="account-icon">🧪</div>
                  <h4>Demo Mode</h4>
                  <p>Try the interface with mock data to see how it works</p>
                  <button class="select-type-btn demo-btn">Try Demo</button>
                </div>
                
                <div class="account-type-card disabled-card" data-type="office365">
                  <div class="account-icon">🏢</div>
                  <h4>Office 365</h4>
                  <p>Business Microsoft account from another organization</p>
                  <div class="setup-required">
                    <small>⚠️ Requires Azure App Registration</small>
                  </div>
                  <button class="select-type-btn disabled" disabled>Setup Required</button>
                </div>
                
                <div class="account-type-card disabled-card" data-type="gmail">
                  <div class="account-icon">📧</div>
                  <h4>Gmail</h4>
                  <p>Google Calendar from Gmail account</p>
                  <div class="setup-required">
                    <small>⚠️ Requires Google Cloud Project</small>
                  </div>
                  <button class="select-type-btn disabled" disabled>Setup Required</button>
                </div>
                
                <div class="account-type-card disabled-card" data-type="outlook_com">
                  <div class="account-icon">📨</div>
                  <h4>Outlook.com</h4>
                  <p>Personal Microsoft account</p>
                  <div class="setup-required">
                    <small>⚠️ Requires Azure App Registration</small>
                  </div>
                  <button class="select-type-btn disabled" disabled>Setup Required</button>
                </div>
                
                <div class="account-type-card disabled-card" data-type="exchange">
                  <div class="account-icon">🏛️</div>
                  <h4>Exchange Server</h4>
                  <p>On-premises Exchange calendar</p>
                  <div class="setup-required">
                    <small>⚠️ Requires EWS Configuration</small>
                  </div>
                  <button class="select-type-btn disabled" disabled>Setup Required</button>
                </div>
              </div>
              
              <!-- Setup Instructions -->
              <div class="setup-instructions">
                <h4>🔧 Production Setup Required</h4>
                <p>To connect real accounts, you need to set up authentication for each service:</p>
                
                <div class="setup-steps">
                  <div class="setup-step">
                    <h5>1. Office 365 / Outlook.com</h5>
                    <ul>
                      <li>Create Azure App Registration in each tenant</li>
                      <li>Configure redirect URIs and permissions</li>
                      <li>Grant admin consent for calendar access</li>
                    </ul>
                  </div>
                  
                  <div class="setup-step">
                    <h5>2. Gmail</h5>
                    <ul>
                      <li>Create Google Cloud Project</li>
                      <li>Enable Google Calendar API</li>
                      <li>Create OAuth 2.0 credentials</li>
                    </ul>
                  </div>
                  
                  <div class="setup-step">
                    <h5>3. Exchange Server</h5>
                    <ul>
                      <li>Configure Exchange Web Services (EWS)</li>
                      <li>Set up authentication credentials</li>
                      <li>Configure network access</li>
                    </ul>
                  </div>
                </div>
                
                <div class="demo-notice">
                  <strong>💡 For now, use Demo Mode to see the interface in action!</strong>
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
        
        if (e.target.classList.contains('demo-btn')) {
          this.handleDemoAccount();
        } else if (!e.target.disabled) {
          const accountType = e.target.closest('.account-type-card').dataset.type;
          this.showAccountConfigForm(accountType);
        }
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

    // Load accounts from localStorage for demo mode
    const storedAccounts = JSON.parse(localStorage.getItem('calendar_accounts') || '[]');
    
    if (storedAccounts.length === 0) {
      accountsList.innerHTML = `
        <div class="no-accounts">
          <div class="no-accounts-icon">📅</div>
          <h4>No accounts connected</h4>
          <p>Add your first calendar account to start syncing across multiple tenants</p>
          <button onclick="document.querySelector('[data-tab=add-account]').click()" class="btn-primary">
            Add Account
          </button>
        </div>
      `;
      return;
    }

    accountsList.innerHTML = storedAccounts.map(account => `
      <div class="account-item ${account.isAuthenticated ? 'authenticated' : 'not-authenticated'}">
        <div class="account-info">
          <div class="account-header">
            <span class="account-type-icon">${this.getAccountTypeIcon(account.type)}</span>
            <div class="account-details">
              <h4>${account.displayName}</h4>
              <p>${account.email}</p>
              <small style="color: #666;">Tenant: ${account.tenantInfo?.name || 'Unknown'}</small>
            </div>
            <div class="account-status">
              ${account.isAuthenticated ? 
                '<span class="status-badge success">✅ Connected</span>' : 
                '<span class="status-badge error">❌ Disconnected</span>'
              }
            </div>
          </div>
          
          <div class="account-stats">
            <span class="stat">📅 ${account.calendars?.length || 0} calendars</span>
            <span class="stat">� ${account.stats?.totalEvents || 0} events</span>
            <span class="stat">⚠️ ${account.stats?.conflicts || 0} conflicts</span>
            ${account.lastSync ? 
              `<span class="stat">⏰ Last sync: ${new Date(account.lastSync).toLocaleTimeString()}</span>` : 
              '<span class="stat">⏰ Never synced</span>'
            }
          </div>
          
          <div class="account-actions">
            <button class="btn-secondary sync-account-btn" data-account-id="${account.id}">
              🔄 Sync
            </button>
            <button class="btn-outline remove-account-btn" data-account-id="${account.id}">
              🗑️ Remove
            </button>
          </div>
        </div>
      </div>
    `).join('');

    // Update summary stats
    this.updateAccountsSummary(storedAccounts);
  }

  /**
   * Update accounts summary display
   */
  updateAccountsSummary(accounts) {
    const totalAccounts = accounts.length;
    const authenticatedAccounts = accounts.filter(acc => acc.isAuthenticated).length;
    const totalCalendars = accounts.reduce((sum, acc) => sum + (acc.calendars?.length || 0), 0);
    const totalEvents = accounts.reduce((sum, acc) => sum + (acc.stats?.totalEvents || 0), 0);
    const conflicts = accounts.reduce((sum, acc) => sum + (acc.stats?.conflicts || 0), 0);

    // Update the summary display if it exists
    const summaryEl = document.querySelector('.accounts-summary');
    if (summaryEl) {
      summaryEl.innerHTML = `
        <h3>📊 Multi-Tenant Calendar Overview</h3>
        <div class="summary-stats">
          <div class="stat-item">
            <span class="stat-number">${totalAccounts}</span>
            <span class="stat-label">Connected Accounts</span>
          </div>
          <div class="stat-item">
            <span class="stat-number">${totalCalendars}</span>
            <span class="stat-label">Total Calendars</span>
          </div>
          <div class="stat-item">
            <span class="stat-number">${totalEvents}</span>
            <span class="stat-label">Total Events</span>
          </div>
          <div class="stat-item">
            <span class="stat-number">${conflicts}</span>
            <span class="stat-label">Conflicts</span>
          </div>
        </div>
        ${totalAccounts > 0 ? `
          <div class="sync-all-container">
            <button class="btn-primary sync-all-btn">🔄 Sync All Accounts</button>
          </div>
        ` : ''}
      `;
    }
  }
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

  /**
   * Handle demo account creation
   */
  async handleDemoAccount() {
    try {
      console.log('🧪 Creating demo account...');
      
      // Create a demo account with realistic data
      const demoAccount = {
        id: 'demo_' + Date.now(),
        type: 'demo',
        email: 'demo@example.com',
        displayName: 'Demo Multi-Tenant Calendar',
        isAuthenticated: true,
        lastSync: new Date().toISOString(),
        calendars: [
          {
            id: 'demo_cal_1',
            name: 'Work Calendar',
            color: '#0078d4',
            events: 15
          },
          {
            id: 'demo_cal_2', 
            name: 'Personal Calendar',
            color: '#8a2be2',
            events: 8
          }
        ],
        stats: {
          totalEvents: 23,
          upcomingEvents: 5,
          conflicts: 2,
          lastSyncStatus: 'success'
        },
        tenantInfo: {
          name: 'Demo Organization',
          domain: 'demo.example.com'
        }
      };

      // Store demo account
      const existingAccounts = JSON.parse(localStorage.getItem('calendar_accounts') || '[]');
      existingAccounts.push(demoAccount);
      localStorage.setItem('calendar_accounts', JSON.stringify(existingAccounts));
      
      // Show success message
      this.showNotification('Demo account created successfully! Check the Overview tab to see your demo calendar data.', 'success');
      
      // Refresh the accounts list
      this.refreshAccountsList();
      
      // Switch to overview tab to show the results
      setTimeout(() => {
        this.switchTab('overview');
      }, 1000);
      
    } catch (error) {
      console.error('Error creating demo account:', error);
      this.showNotification('Failed to create demo account. Please try again.', 'error');
    }
  }

  /**
   * Show notification message
   */
  showNotification(message, type = 'info') {
    // Create notification element
    const notification = document.createElement('div');
    notification.className = `notification notification-${type}`;
    notification.innerHTML = `
      <div class="notification-content">
        <span class="notification-icon">${type === 'success' ? '✅' : type === 'error' ? '❌' : 'ℹ️'}</span>
        <span class="notification-message">${message}</span>
      </div>
    `;
    
    // Add to page
    document.body.appendChild(notification);
    
    // Animate in
    setTimeout(() => notification.classList.add('show'), 100);
    
    // Remove after delay
    setTimeout(() => {
      notification.classList.remove('show');
      setTimeout(() => document.body.removeChild(notification), 300);
    }, 4000);
  }
}

// Export for global use
window.MultiAccountUI = MultiAccountUI;
