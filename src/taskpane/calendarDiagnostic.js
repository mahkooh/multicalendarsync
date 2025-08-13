/**
 * Calendar Discovery Diagnostic Tool
 * This will help identify what calendars are available through different methods
 */

class CalendarDiscoveryDiagnostic {
  constructor() {
    this.results = {
      graphCalendars: [],
      outlookAccounts: [],
      connectedAccounts: [],
      errors: []
    };
  }

  async runFullDiagnostic() {
    console.log('🔍 Starting comprehensive calendar discovery diagnostic...');
    
    const diagnosticResults = {
      timestamp: new Date().toISOString(),
      officeContext: this.getOfficeContext(),
      graphApiResults: await this.testGraphAPI(),
      outlookContext: this.getOutlookContext(),
      localStorageData: this.checkLocalStorage(),
      recommendations: []
    };

    this.generateRecommendations(diagnosticResults);
    this.displayResults(diagnosticResults);
    
    return diagnosticResults;
  }

  getOfficeContext() {
    try {
      if (!Office || !Office.context) {
        return { available: false, error: 'Office context not available' };
      }

      const diagnostics = Office.context.mailbox?.diagnostics;
      
      return {
        available: true,
        hostName: diagnostics?.hostName,
        hostVersion: diagnostics?.hostVersion,
        platform: diagnostics?.platform,
        outlookVersion: diagnostics?.outlookVersion,
        ewsUrl: Office.context.mailbox?.ewsUrl,
        userEmail: Office.context.mailbox?.userProfile?.emailAddress,
        userName: Office.context.mailbox?.userProfile?.displayName,
        timeZone: Office.context.mailbox?.userProfile?.timeZone,
        authenticationCapable: !!Office.context.auth
      };
    } catch (error) {
      return { 
        available: false, 
        error: error.message,
        fallbackData: {
          officeAvailable: !!window.Office,
          location: window.location.href
        }
      };
    }
  }

  async testGraphAPI() {
    const results = {
      accessTokenObtained: false,
      userProfileSuccess: false,
      calendarsSuccess: false,
      calendarsFound: [],
      errors: []
    };

    try {
      // Test 1: Get access token
      console.log('🔐 Testing Graph API access token...');
      const accessToken = await this.getGraphAccessToken();
      results.accessTokenObtained = true;
      console.log('✅ Access token obtained');

      // Test 2: Get user profile
      console.log('👤 Testing user profile API...');
      const profileResponse = await fetch('https://graph.microsoft.com/v1.0/me', {
        headers: {
          'Authorization': `Bearer ${accessToken}`,
          'Content-Type': 'application/json'
        }
      });

      if (profileResponse.ok) {
        const profile = await profileResponse.json();
        results.userProfile = profile;
        results.userProfileSuccess = true;
        console.log(`✅ User profile: ${profile.displayName} (${profile.mail || profile.userPrincipalName})`);
      } else {
        throw new Error(`Profile API failed: ${profileResponse.status}`);
      }

      // Test 3: Get calendars with detailed info
      console.log('📅 Testing calendar discovery...');
      const calendarResponse = await fetch('https://graph.microsoft.com/v1.0/me/calendars?$select=id,name,canEdit,isDefaultCalendar,owner,canShare,canViewPrivateItems,changeKey,color', {
        headers: {
          'Authorization': `Bearer ${accessToken}`,
          'Content-Type': 'application/json'
        }
      });

      if (calendarResponse.ok) {
        const calendarData = await calendarResponse.json();
        results.calendarsSuccess = true;
        results.calendarsFound = calendarData.value || [];
        console.log(`✅ Found ${results.calendarsFound.length} calendars via Graph API`);
        
        // Log each calendar
        results.calendarsFound.forEach((cal, index) => {
          console.log(`  📅 ${index + 1}. ${cal.name} (${cal.id})`);
          console.log(`     - Default: ${cal.isDefaultCalendar}`);
          console.log(`     - Can Edit: ${cal.canEdit}`);
          console.log(`     - Owner: ${cal.owner?.name || 'Unknown'}`);
        });
      } else {
        throw new Error(`Calendar API failed: ${calendarResponse.status}`);
      }

      // Test 4: Check for additional mailboxes/accounts
      console.log('📮 Testing for additional mailboxes...');
      try {
        const mailboxResponse = await fetch('https://graph.microsoft.com/v1.0/me/mailFolders', {
          headers: {
            'Authorization': `Bearer ${accessToken}`,
            'Content-Type': 'application/json'
          }
        });
        
        if (mailboxResponse.ok) {
          const mailboxData = await mailboxResponse.json();
          results.mailboxInfo = {
            foldersFound: mailboxData.value?.length || 0,
            rootFolders: mailboxData.value?.map(f => f.displayName) || []
          };
        }
      } catch (mailboxError) {
        results.errors.push(`Mailbox check failed: ${mailboxError.message}`);
      }

    } catch (error) {
      results.errors.push(error.message);
      console.error('❌ Graph API test failed:', error);
    }

    return results;
  }

  async getGraphAccessToken() {
    return new Promise((resolve, reject) => {
      if (!Office?.context?.auth?.getAccessToken) {
        reject(new Error('Office.js authentication not available'));
        return;
      }

      const tokenOptions = {
        allowConsentPrompt: true,
        forMSGraphAccess: true
      };

      Office.context.auth.getAccessToken(tokenOptions, (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve(result.value);
        } else {
          reject(new Error(`Authentication failed: ${result.error?.message || 'Unknown error'}`));
        }
      });
    });
  }

  getOutlookContext() {
    try {
      if (!Office?.context?.mailbox) {
        return { available: false };
      }

      // Try to get information about connected accounts
      const mailbox = Office.context.mailbox;
      const userProfile = mailbox.userProfile;
      
      return {
        available: true,
        primaryEmail: userProfile?.emailAddress,
        displayName: userProfile?.displayName,
        timeZone: userProfile?.timeZone,
        serverUrl: mailbox.ewsUrl,
        // Note: There's no direct API to get connected accounts in Office.js
        limitationNote: 'Connected accounts (Gmail, etc.) are not accessible via Office.js APIs'
      };
    } catch (error) {
      return { 
        available: false, 
        error: error.message 
      };
    }
  }

  checkLocalStorage() {
    try {
      const keys = Object.keys(localStorage);
      const calendarKeys = keys.filter(key => 
        key.includes('calendar') || 
        key.includes('account') || 
        key.includes('sync')
      );

      return {
        totalKeys: keys.length,
        calendarRelatedKeys: calendarKeys,
        syncConfig: localStorage.getItem('calendarSyncConfig'),
        lastSync: localStorage.getItem('lastSync')
      };
    } catch (error) {
      return { 
        available: false, 
        error: error.message 
      };
    }
  }

  generateRecommendations(diagnosticResults) {
    const recommendations = [];

    // Check Graph API success
    if (!diagnosticResults.graphApiResults.accessTokenObtained) {
      recommendations.push({
        type: 'error',
        title: 'Graph API Authentication Failed',
        description: 'Cannot access Microsoft Graph API. This means we can only see mock calendars.',
        action: 'Ensure the add-in is properly installed and has necessary permissions.'
      });
    } else if (diagnosticResults.graphApiResults.calendarsFound.length === 0) {
      recommendations.push({
        type: 'warning',
        title: 'No Calendars Found via Graph API',
        description: 'Graph API authentication works but no calendars were found.',
        action: 'Check if you have any calendars in your Office 365/Exchange account.'
      });
    } else {
      recommendations.push({
        type: 'success',
        title: `Found ${diagnosticResults.graphApiResults.calendarsFound.length} Calendar(s)`,
        description: 'These are calendars accessible through Microsoft Graph API.',
        action: 'These calendars can be synchronized automatically.'
      });
    }

    // Check for limitations
    recommendations.push({
      type: 'info',
      title: 'Connected Account Limitation',
      description: 'Calendars from connected accounts (Gmail, Yahoo, personal Outlook.com) added to Outlook app are not visible through Graph API.',
      action: 'To sync these calendars, we need to add manual calendar connection features.'
    });

    // Suggest next steps
    if (diagnosticResults.graphApiResults.calendarsFound.length < 5) {
      recommendations.push({
        type: 'suggestion',
        title: 'Add More Calendar Sources',
        description: 'To sync your 5 calendars across multiple accounts, we need to add support for additional calendar services.',
        action: 'Consider implementing Google Calendar API, manual account addition, or calendar import features.'
      });
    }

    diagnosticResults.recommendations = recommendations;
  }

  displayResults(diagnosticResults) {
    // Create a detailed report in the console
    console.log('\n🔍 =================== CALENDAR DISCOVERY DIAGNOSTIC REPORT ===================');
    console.log(`📅 Timestamp: ${diagnosticResults.timestamp}`);
    
    console.log('\n📊 Office Context:');
    console.log(JSON.stringify(diagnosticResults.officeContext, null, 2));
    
    console.log('\n🔗 Graph API Results:');
    console.log(JSON.stringify(diagnosticResults.graphApiResults, null, 2));
    
    console.log('\n💡 Recommendations:');
    diagnosticResults.recommendations.forEach((rec, index) => {
      console.log(`${index + 1}. [${rec.type.toUpperCase()}] ${rec.title}`);
      console.log(`   ${rec.description}`);
      console.log(`   Action: ${rec.action}\n`);
    });
    
    console.log('=================== END DIAGNOSTIC REPORT ===================\n');

    // Also update the UI if possible
    this.updateUI(diagnosticResults);
  }

  updateUI(diagnosticResults) {
    try {
      const container = document.getElementById('calendar-list') || document.getElementById('auth-container');
      if (!container) return;

      const html = `
        <div class="diagnostic-results">
          <h3>🔍 Calendar Discovery Results</h3>
          
          <div class="result-section">
            <h4>📊 Office Context Status</h4>
            <p><strong>Available:</strong> ${diagnosticResults.officeContext.available ? '✅ Yes' : '❌ No'}</p>
            ${diagnosticResults.officeContext.available ? `
              <p><strong>Host:</strong> ${diagnosticResults.officeContext.hostName}</p>
              <p><strong>User:</strong> ${diagnosticResults.officeContext.userName} (${diagnosticResults.officeContext.userEmail})</p>
            ` : `
              <p><strong>Error:</strong> ${diagnosticResults.officeContext.error}</p>
            `}
          </div>

          <div class="result-section">
            <h4>🔗 Microsoft Graph API Status</h4>
            <p><strong>Authentication:</strong> ${diagnosticResults.graphApiResults.accessTokenObtained ? '✅ Success' : '❌ Failed'}</p>
            <p><strong>Calendars Found:</strong> ${diagnosticResults.graphApiResults.calendarsFound.length}</p>
            ${diagnosticResults.graphApiResults.calendarsFound.length > 0 ? `
              <ul>
                ${diagnosticResults.graphApiResults.calendarsFound.map(cal => 
                  `<li>📅 ${cal.name} ${cal.isDefaultCalendar ? '(Default)' : ''}</li>`
                ).join('')}
              </ul>
            ` : ''}
          </div>

          <div class="result-section">
            <h4>💡 Next Steps</h4>
            ${diagnosticResults.recommendations.map(rec => `
              <div class="recommendation ${rec.type}">
                <strong>${rec.title}</strong><br>
                ${rec.description}<br>
                <em>Action: ${rec.action}</em>
              </div>
            `).join('')}
          </div>
        </div>
      `;

      container.innerHTML = html;
    } catch (error) {
      console.warn('Could not update UI with diagnostic results:', error);
    }
  }
}

// Make it globally available
window.CalendarDiscoveryDiagnostic = CalendarDiscoveryDiagnostic;
