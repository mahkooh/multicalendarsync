/**
 * Multi-Tenant Calendar Connection Manager
 * Handles authentication and calendar access across multiple tenants and services
 */

export class MultiTenantCalendarManager {
  constructor() {
    this.connectedAccounts = new Map();
    this.authTokens = new Map();
    this.calendarCache = new Map();
    this.syncInProgress = false;
    
    // Supported account types
    this.accountTypes = {
      OFFICE365: 'office365',
      EXCHANGE: 'exchange', 
      GMAIL: 'gmail',
      OUTLOOK_COM: 'outlook_com',
      OTHER: 'other'
    };
  }

  /**
   * Add a new account connection
   */
  async addAccount(accountConfig) {
    const {
      id,
      type,
      email,
      displayName,
      tenantId,
      clientId,
      authUrl,
      calendarUrl
    } = accountConfig;

    console.log(`🔗 Adding account: ${displayName} (${email})`);

    const account = {
      id,
      type,
      email,
      displayName,
      tenantId,
      clientId,
      authUrl,
      calendarUrl,
      isAuthenticated: false,
      calendars: [],
      lastSync: null,
      syncEnabled: true
    };

    this.connectedAccounts.set(id, account);
    
    // Trigger authentication flow
    await this.authenticateAccount(id);
    
    return account;
  }

  /**
   * Authenticate with different account types
   */
  async authenticateAccount(accountId) {
    const account = this.connectedAccounts.get(accountId);
    if (!account) throw new Error(`Account ${accountId} not found`);

    console.log(`🔐 Authenticating ${account.displayName}...`);

    switch (account.type) {
      case this.accountTypes.OFFICE365:
        return await this.authenticateOffice365(account);
      case this.accountTypes.GMAIL:
        return await this.authenticateGmail(account);
      case this.accountTypes.OUTLOOK_COM:
        return await this.authenticateOutlookCom(account);
      case this.accountTypes.EXCHANGE:
        return await this.authenticateExchange(account);
      default:
        throw new Error(`Unsupported account type: ${account.type}`);
    }
  }

  /**
   * Office 365 / Azure AD Authentication
   */
  async authenticateOffice365(account) {
    try {
      // For Office 365 accounts in different tenants
      const authUrl = `https://login.microsoftonline.com/${account.tenantId}/oauth2/v2.0/authorize`;
      const params = new URLSearchParams({
        client_id: account.clientId,
        response_type: 'code',
        redirect_uri: `${window.location.origin}/auth-callback`,
        scope: 'https://graph.microsoft.com/Calendars.ReadWrite https://graph.microsoft.com/User.Read',
        state: account.id,
        prompt: 'login' // Force re-authentication for different tenant
      });

      const fullAuthUrl = `${authUrl}?${params.toString()}`;
      
      // Open authentication popup
      const authResult = await this.openAuthPopup(fullAuthUrl, account.id);
      
      if (authResult.access_token) {
        this.authTokens.set(account.id, {
          access_token: authResult.access_token,
          refresh_token: authResult.refresh_token,
          expires_at: Date.now() + (authResult.expires_in * 1000)
        });
        
        account.isAuthenticated = true;
        await this.discoverAccountCalendars(account.id);
        return true;
      }
      
      throw new Error('Failed to get access token');
      
    } catch (error) {
      console.error(`❌ Office 365 auth failed for ${account.email}:`, error);
      throw error;
    }
  }

  /**
   * Gmail Authentication
   */
  async authenticateGmail(account) {
    try {
      // Google OAuth 2.0 flow
      const authUrl = 'https://accounts.google.com/o/oauth2/auth';
      const params = new URLSearchParams({
        client_id: account.clientId,
        response_type: 'code',
        redirect_uri: `${window.location.origin}/auth-callback`,
        scope: 'https://www.googleapis.com/auth/calendar',
        access_type: 'offline',
        prompt: 'consent',
        state: account.id
      });

      const fullAuthUrl = `${authUrl}?${params.toString()}`;
      const authResult = await this.openAuthPopup(fullAuthUrl, account.id);
      
      if (authResult.access_token) {
        this.authTokens.set(account.id, authResult);
        account.isAuthenticated = true;
        await this.discoverAccountCalendars(account.id);
        return true;
      }
      
      throw new Error('Gmail authentication failed');
      
    } catch (error) {
      console.error(`❌ Gmail auth failed for ${account.email}:`, error);
      throw error;
    }
  }

  /**
   * Outlook.com Personal Account Authentication
   */
  async authenticateOutlookCom(account) {
    try {
      // Microsoft personal account OAuth
      const authUrl = 'https://login.microsoftonline.com/common/oauth2/v2.0/authorize';
      const params = new URLSearchParams({
        client_id: account.clientId,
        response_type: 'code',
        redirect_uri: `${window.location.origin}/auth-callback`,
        scope: 'https://graph.microsoft.com/Calendars.ReadWrite https://graph.microsoft.com/User.Read',
        state: account.id
      });

      const fullAuthUrl = `${authUrl}?${params.toString()}`;
      const authResult = await this.openAuthPopup(fullAuthUrl, account.id);
      
      if (authResult.access_token) {
        this.authTokens.set(account.id, authResult);
        account.isAuthenticated = true;
        await this.discoverAccountCalendars(account.id);
        return true;
      }
      
      throw new Error('Outlook.com authentication failed');
      
    } catch (error) {
      console.error(`❌ Outlook.com auth failed for ${account.email}:`, error);
      throw error;
    }
  }

  /**
   * Exchange Server Authentication (Basic Auth or NTLM)
   */
  async authenticateExchange(account) {
    try {
      // For on-premises Exchange servers
      // This would typically require username/password or certificate auth
      console.log(`🔐 Exchange authentication for ${account.email} not yet implemented`);
      
      // Placeholder - would implement EWS authentication
      account.isAuthenticated = false;
      return false;
      
    } catch (error) {
      console.error(`❌ Exchange auth failed for ${account.email}:`, error);
      throw error;
    }
  }

  /**
   * Open authentication popup and wait for result
   */
  async openAuthPopup(authUrl, accountId) {
    return new Promise((resolve, reject) => {
      const popup = window.open(
        authUrl,
        `auth_${accountId}`,
        'width=500,height=600,scrollbars=yes,resizable=yes'
      );

      // Listen for messages from the popup
      const messageListener = (event) => {
        if (event.origin !== window.location.origin) return;
        
        if (event.data.type === 'AUTH_SUCCESS' && event.data.accountId === accountId) {
          window.removeEventListener('message', messageListener);
          popup.close();
          resolve(event.data.tokens);
        } else if (event.data.type === 'AUTH_ERROR' && event.data.accountId === accountId) {
          window.removeEventListener('message', messageListener);
          popup.close();
          reject(new Error(event.data.error));
        }
      };

      window.addEventListener('message', messageListener);

      // Check if popup was closed manually
      const checkClosed = setInterval(() => {
        if (popup.closed) {
          clearInterval(checkClosed);
          window.removeEventListener('message', messageListener);
          reject(new Error('Authentication was cancelled'));
        }
      }, 1000);

      // Timeout after 5 minutes
      setTimeout(() => {
        if (!popup.closed) {
          popup.close();
          clearInterval(checkClosed);
          window.removeEventListener('message', messageListener);
          reject(new Error('Authentication timeout'));
        }
      }, 300000);
    });
  }

  /**
   * Discover calendars for an authenticated account
   */
  async discoverAccountCalendars(accountId) {
    const account = this.connectedAccounts.get(accountId);
    if (!account || !account.isAuthenticated) {
      throw new Error(`Account ${accountId} not authenticated`);
    }

    console.log(`📅 Discovering calendars for ${account.email}...`);

    switch (account.type) {
      case this.accountTypes.OFFICE365:
      case this.accountTypes.OUTLOOK_COM:
        return await this.getGraphCalendars(accountId);
      case this.accountTypes.GMAIL:
        return await this.getGmailCalendars(accountId);
      case this.accountTypes.EXCHANGE:
        return await this.getExchangeCalendars(accountId);
      default:
        throw new Error(`Calendar discovery not supported for ${account.type}`);
    }
  }

  /**
   * Get calendars via Microsoft Graph API
   */
  async getGraphCalendars(accountId) {
    const account = this.connectedAccounts.get(accountId);
    const tokens = this.authTokens.get(accountId);
    
    if (!tokens || !tokens.access_token) {
      throw new Error(`No access token for account ${accountId}`);
    }

    try {
      const response = await fetch('https://graph.microsoft.com/v1.0/me/calendars', {
        headers: {
          'Authorization': `Bearer ${tokens.access_token}`,
          'Content-Type': 'application/json'
        }
      });

      if (!response.ok) {
        throw new Error(`Graph API error: ${response.status} ${response.statusText}`);
      }

      const data = await response.json();
      
      account.calendars = data.value.map(cal => ({
        id: cal.id,
        name: cal.name,
        isDefault: cal.isDefaultCalendar,
        canEdit: cal.canEdit,
        owner: cal.owner,
        syncEnabled: cal.isDefaultCalendar // Enable sync for default calendar by default
      }));

      console.log(`📅 Found ${account.calendars.length} calendars for ${account.email}`);
      return account.calendars;
      
    } catch (error) {
      console.error(`❌ Failed to get Graph calendars for ${account.email}:`, error);
      throw error;
    }
  }

  /**
   * Get calendars via Google Calendar API
   */
  async getGmailCalendars(accountId) {
    const account = this.connectedAccounts.get(accountId);
    const tokens = this.authTokens.get(accountId);
    
    if (!tokens || !tokens.access_token) {
      throw new Error(`No access token for account ${accountId}`);
    }

    try {
      const response = await fetch('https://www.googleapis.com/calendar/v3/users/me/calendarList', {
        headers: {
          'Authorization': `Bearer ${tokens.access_token}`,
          'Content-Type': 'application/json'
        }
      });

      if (!response.ok) {
        throw new Error(`Google Calendar API error: ${response.status} ${response.statusText}`);
      }

      const data = await response.json();
      
      account.calendars = data.items.map(cal => ({
        id: cal.id,
        name: cal.summary,
        isDefault: cal.primary,
        canEdit: cal.accessRole === 'owner' || cal.accessRole === 'writer',
        owner: { name: cal.summary, address: cal.id },
        syncEnabled: cal.primary // Enable sync for primary calendar by default
      }));

      console.log(`📅 Found ${account.calendars.length} Gmail calendars for ${account.email}`);
      return account.calendars;
      
    } catch (error) {
      console.error(`❌ Failed to get Gmail calendars for ${account.email}:`, error);
      throw error;
    }
  }

  /**
   * Get calendars via Exchange Web Services
   */
  async getExchangeCalendars(accountId) {
    // Placeholder for EWS implementation
    console.log(`📅 Exchange calendar discovery for ${accountId} not yet implemented`);
    return [];
  }

  /**
   * Get all connected accounts summary
   */
  getAccountsSummary() {
    const summary = [];
    
    for (const [id, account] of this.connectedAccounts) {
      summary.push({
        id,
        email: account.email,
        displayName: account.displayName,
        type: account.type,
        isAuthenticated: account.isAuthenticated,
        calendarCount: account.calendars.length,
        syncEnabled: account.syncEnabled,
        lastSync: account.lastSync
      });
    }
    
    return summary;
  }

  /**
   * Create mock accounts for testing
   */
  createMockAccounts() {
    const mockAccounts = [
      {
        id: 'tenant-a',
        type: this.accountTypes.OFFICE365,
        email: 'dan.hookham@tenanta.com',
        displayName: 'Dan Hookham (Tenant A)',
        tenantId: 'tenant-a-guid',
        clientId: 'client-a-guid'
      },
      {
        id: 'tenant-b',
        type: this.accountTypes.OFFICE365,
        email: 'dan.hookham@tenantb.com', 
        displayName: 'Dan Hookham (Tenant B)',
        tenantId: 'tenant-b-guid',
        clientId: 'client-b-guid'
      },
      {
        id: 'tenant-c',
        type: this.accountTypes.OFFICE365,
        email: 'dan.hookham@tenantc.com',
        displayName: 'Dan Hookham (Tenant C)', 
        tenantId: 'tenant-c-guid',
        clientId: 'client-c-guid'
      },
      {
        id: 'gmail-personal',
        type: this.accountTypes.GMAIL,
        email: 'dan.hookham@gmail.com',
        displayName: 'Dan Hookham (Gmail)',
        clientId: 'gmail-client-id'
      },
      {
        id: 'outlook-personal',
        type: this.accountTypes.OUTLOOK_COM,
        email: 'dan.hookham@outlook.com',
        displayName: 'Dan Hookham (Personal)',
        clientId: 'outlook-client-id'
      }
    ];

    // Add mock authentication status
    mockAccounts.forEach(config => {
      const account = { ...config };
      account.isAuthenticated = true;
      account.calendars = [
        {
          id: `${config.id}-default`,
          name: 'Calendar',
          isDefault: true,
          canEdit: true,
          syncEnabled: true
        }
      ];
      account.lastSync = new Date();
      account.syncEnabled = true;
      
      this.connectedAccounts.set(account.id, account);
    });

    console.log('📊 Created mock accounts for testing');
    return this.getAccountsSummary();
  }
}

// Create global instance
window.multiTenantManager = new MultiTenantCalendarManager();
