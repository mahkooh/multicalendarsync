/**
 * Simplified Authentication for Multi-Tenant Calendar Access
 * This approach uses manual login URLs instead of complex Azure App Registrations
 */

class SimpleCalendarAuth {
  constructor() {
    this.accounts = [];
    this.currentAccountIndex = 0;
  }

  /**
   * Simplified approach: Generate manual login URLs for each tenant
   */
  generateAccountLoginUrls() {
    const baseAuthUrl = 'https://login.microsoftonline.com/common/oauth2/v2.0/authorize';
    const clientId = 'your-app-id'; // This would be a single app registration in your main tenant
    const redirectUri = encodeURIComponent('https://mahkooh.github.io/multicalendarsync/auth-callback.html');
    const scope = encodeURIComponent('https://graph.microsoft.com/Calendars.ReadWrite https://graph.microsoft.com/User.Read');
    
    return [
      {
        name: 'Account 1 (Tenant 1)',
        loginUrl: `${baseAuthUrl}?client_id=${clientId}&response_type=code&redirect_uri=${redirectUri}&scope=${scope}&prompt=login`,
        description: 'Click to authenticate with your first Microsoft account'
      },
      {
        name: 'Account 2 (Tenant 2)', 
        loginUrl: `${baseAuthUrl}?client_id=${clientId}&response_type=code&redirect_uri=${redirectUri}&scope=${scope}&prompt=login`,
        description: 'Click to authenticate with your second Microsoft account'
      },
      {
        name: 'Account 3 (Tenant 3)',
        loginUrl: `${baseAuthUrl}?client_id=${clientId}&response_type=code&redirect_uri=${redirectUri}&scope=${scope}&prompt=login`,
        description: 'Click to authenticate with your third Microsoft account'
      },
      {
        name: 'Account 4 (Tenant 4)',
        loginUrl: `${baseAuthUrl}?client_id=${clientId}&response_type=code&redirect_uri=${redirectUri}&scope=${scope}&prompt=login`,
        description: 'Click to authenticate with your fourth Microsoft account'
      },
      {
        name: 'Account 5 (Tenant 5)',
        loginUrl: `${baseAuthUrl}?client_id=${clientId}&response_type=code&redirect_uri=${redirectUri}&scope=${scope}&prompt=login`,
        description: 'Click to authenticate with your fifth Microsoft account'
      }
    ];
  }

  /**
   * Alternative: Use device code flow for each account
   */
  async initiateDeviceCodeFlow(accountName) {
    const deviceCodeUrl = 'https://login.microsoftonline.com/common/oauth2/v2.0/devicecode';
    const clientId = 'your-app-id';
    const scope = 'https://graph.microsoft.com/Calendars.ReadWrite https://graph.microsoft.com/User.Read';

    try {
      const response = await fetch(deviceCodeUrl, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/x-www-form-urlencoded'
        },
        body: `client_id=${clientId}&scope=${encodeURIComponent(scope)}`
      });

      const data = await response.json();
      
      return {
        accountName,
        deviceCode: data.device_code,
        userCode: data.user_code,
        verificationUri: data.verification_uri,
        verificationUriComplete: data.verification_uri_complete,
        expiresIn: data.expires_in,
        interval: data.interval,
        message: data.message
      };
    } catch (error) {
      console.error(`Failed to initiate device code flow for ${accountName}:`, error);
      throw error;
    }
  }

  /**
   * Simpler approach: Use iframe-based authentication with prompt
   */
  async authenticateAccountWithIframe(accountIndex) {
    return new Promise((resolve, reject) => {
      const authUrl = this.generateAccountLoginUrls()[accountIndex].loginUrl;
      
      // Create a popup window for authentication
      const popup = window.open(
        authUrl,
        `auth_${accountIndex}`,
        'width=500,height=600,scrollbars=yes,resizable=yes'
      );

      // Listen for the popup to close or receive a message
      const checkClosed = setInterval(() => {
        if (popup.closed) {
          clearInterval(checkClosed);
          // Check localStorage for auth result
          const authResult = localStorage.getItem(`auth_result_${accountIndex}`);
          if (authResult) {
            const result = JSON.parse(authResult);
            localStorage.removeItem(`auth_result_${accountIndex}`);
            resolve(result);
          } else {
            reject(new Error('Authentication was cancelled or failed'));
          }
        }
      }, 1000);

      // Set a timeout
      setTimeout(() => {
        if (!popup.closed) {
          popup.close();
          clearInterval(checkClosed);
          reject(new Error('Authentication timeout'));
        }
      }, 300000); // 5 minutes timeout
    });
  }

  /**
   * Mock mode for development and demonstration
   */
  getMockAccounts() {
    return [
      {
        id: 'mock-account-1',
        email: 'user1@tenant1.com',
        displayName: 'User One',
        tenant: 'tenant1.onmicrosoft.com',
        calendars: [
          { id: 'calendar1', name: 'Work Calendar', isDefault: true },
          { id: 'calendar2', name: 'Personal Calendar', isDefault: false }
        ]
      },
      {
        id: 'mock-account-2', 
        email: 'user2@tenant2.com',
        displayName: 'User Two',
        tenant: 'tenant2.onmicrosoft.com',
        calendars: [
          { id: 'calendar3', name: 'Corporate Calendar', isDefault: true }
        ]
      },
      {
        id: 'mock-account-3',
        email: 'user3@tenant3.com', 
        displayName: 'User Three',
        tenant: 'tenant3.onmicrosoft.com',
        calendars: [
          { id: 'calendar4', name: 'Project Calendar', isDefault: true },
          { id: 'calendar5', name: 'Team Calendar', isDefault: false }
        ]
      },
      {
        id: 'mock-account-4',
        email: 'user4@tenant4.com',
        displayName: 'User Four', 
        tenant: 'tenant4.onmicrosoft.com',
        calendars: [
          { id: 'calendar6', name: 'Executive Calendar', isDefault: true }
        ]
      },
      {
        id: 'mock-account-5',
        email: 'user5@tenant5.com',
        displayName: 'User Five',
        tenant: 'tenant5.onmicrosoft.com',
        calendars: [
          { id: 'calendar7', name: 'Consulting Calendar', isDefault: true },
          { id: 'calendar8', name: 'Client Calendar', isDefault: false }
        ]
      }
    ];
  }

  /**
   * Display authentication options to user
   */
  renderAuthenticationOptions() {
    const container = document.getElementById('auth-container');
    if (!container) return;

    const accounts = this.generateAccountLoginUrls();
    
    container.innerHTML = `
      <div class="auth-section">
        <h3>📅 Multi-Calendar Authentication</h3>
        <p>Since you have calendars across multiple tenants, you have two options:</p>
        
        <div class="auth-option">
          <h4>Option 1: Manual Login (Recommended)</h4>
          <p>Click each link below to authenticate with each of your Microsoft accounts:</p>
          <div class="account-list">
            ${accounts.map((account, index) => `
              <div class="account-item">
                <button class="auth-btn" onclick="window.open('${account.loginUrl}', '_blank')">
                  🔐 ${account.name}
                </button>
                <small>${account.description}</small>
              </div>
            `).join('')}
          </div>
        </div>
        
        <div class="auth-option">
          <h4>Option 2: Mock Mode (Testing)</h4>
          <p>Use mock data to test the synchronization functionality:</p>
          <button class="mock-btn" onclick="calendarAuth.useMockMode()">
            🧪 Use Mock Data
          </button>
        </div>
        
        <div class="auth-help">
          <h4>ℹ️ Why Multiple Logins?</h4>
          <p>Each Microsoft tenant requires separate authentication. This approach avoids the complexity of setting up Azure App Registrations in every tenant.</p>
        </div>
      </div>
    `;
  }

  /**
   * Use mock mode for testing
   */
  useMockMode() {
    this.accounts = this.getMockAccounts();
    console.log('📊 Using mock accounts:', this.accounts);
    
    // Trigger the calendar display
    if (window.calendarSync) {
      window.calendarSync.displayAccounts(this.accounts);
    }
  }
}

// Create global instance
window.calendarAuth = new SimpleCalendarAuth();
