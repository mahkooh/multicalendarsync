# MultiCalendar Sync - Outlook Add-in

An Outlook add-in that synchronizes busy time across multiple calendars to ensure consistent availability visibility to colleagues from different companies.

## Problem Solved

When you work with multiple companies and have separate calendars for each, colleagues from Company A can only see your Company A calendar when checking your availability. They don't see that you're busy with meetings from Companies B, C, D, or E, leading to scheduling conflicts.

## Solution

MultiCalendar Sync automatically creates private "busy" blocks in each calendar when you have appointments in your other calendars. This ensures that:

- ✅ Colleagues see you as busy across all calendars
- ✅ Your calendar view stays clean (busy blocks are hidden from your view)
- ✅ Meeting details remain private (only busy/free status is synced)
- ✅ You maintain control over which calendars participate in sync

## Features

### Core Functionality
- **Automatic Synchronization**: Monitors all connected calendars and creates busy blocks when conflicts are detected
- **Privacy Protection**: Only syncs busy/free status, never meeting details or attendees
- **Clean User Experience**: Sync blocks are hidden from your personal calendar view
- **Manual Override**: Start manual sync operations when needed
- **Real-time Monitoring**: Continuous monitoring with configurable sync intervals

### Calendar Management
- **Multi-Account Support**: Works with Exchange, Office 365, Outlook.com, and SharePoint calendars
- **Selective Sync**: Choose which calendars participate in synchronization
- **Status Dashboard**: View sync status and recent activity
- **Configuration Options**: Customize sync behavior, intervals, and busy block appearance

### Security & Compliance
- **Least Privilege**: Requests only necessary calendar permissions
- **Local Configuration**: Settings stored locally for privacy
- **Audit Trail**: Activity logging for troubleshooting and compliance

## How It Works

1. **Discovery**: The add-in discovers all calendars connected to your Outlook
2. **Monitoring**: Continuously monitors enabled calendars for busy time
3. **Analysis**: Identifies conflicts where you're busy in one calendar but free in others
4. **Synchronization**: Creates private busy blocks in calendars that don't show the conflict
5. **Maintenance**: Removes outdated sync blocks and updates as needed

## Technical Architecture

### Frontend
- **Office.js API**: Native Outlook add-in framework
- **Modern JavaScript**: ES6+ with modular architecture
- **Fluent UI**: Microsoft's design system for consistent appearance
- **Responsive Design**: Works across desktop, web, and mobile Outlook

### Backend Integration
- **Microsoft Graph API**: For calendar access and manipulation
- **Exchange Web Services**: Fallback for older Exchange environments
- **MSAL Authentication**: Secure OAuth 2.0 authentication flow

### Privacy Design
- **Private Events**: Sync blocks are marked as private to hide from user view
- **No Content Sync**: Only start/end times and busy status are synchronized
- **Local Storage**: Configuration and preferences stored locally
- **Minimal Permissions**: Requests only calendar read/write permissions

## Installation & Setup

### Prerequisites
- Microsoft Outlook (Desktop, Web, or Mobile)
- Office 365 or Exchange Online account
- Node.js 16+ for development

### Development Setup
1. Clone the repository
2. Install dependencies: `npm install`
3. Start the development server: `npm start`
4. Sideload the add-in in Outlook for testing

### Production Deployment
1. Build the production package: `npm run build`
2. Deploy to your web hosting service
3. Update manifest URLs to production endpoints
4. Distribute through Microsoft AppSource or organization catalog

## Configuration Options

### Sync Settings
- **Sync Interval**: How often to check for calendar changes (default: 15 minutes)
- **Look-ahead Period**: How far in the future to sync events (default: 14 days)
- **Look-behind Period**: How far in the past to maintain sync (default: 1 day)

### Busy Block Customization
- **Subject Line**: Customize the subject of sync blocks (default: "[Auto-Sync] Busy")
- **Category**: Set calendar category for easy identification
- **Visibility**: Configure privacy settings for sync blocks

### Calendar Selection
- **Enable/Disable**: Toggle synchronization for individual calendars
- **Priority Settings**: Set which calendars take precedence during conflicts
- **Exclusion Rules**: Exclude specific types of events from sync

## Privacy & Security

### Data Handling
- **No Cloud Storage**: All configuration stored locally in your browser
- **Minimal Data Access**: Only accesses calendar start/end times and busy/free status
- **No Content Reading**: Never reads meeting subjects, attendees, or content
- **Local Processing**: All sync logic runs locally in your Outlook client

### Permissions
- **MailboxItem.ReadWrite.User**: Access to calendar items in your mailbox
- **Calendars.ReadWrite.Shared**: Access to shared and connected calendars
- **No Network Access**: No external API calls except to Microsoft Graph

### Compliance
- **GDPR Compatible**: No personal data stored or transmitted
- **SOC 2 Aligned**: Follows Microsoft security standards
- **Audit Logging**: Local activity logs for compliance requirements

## Troubleshooting

### Common Issues
1. **Permission Errors**: Ensure all required permissions are granted in Office 365 admin center
2. **Sync Not Working**: Check that at least 2 calendars are enabled for synchronization
3. **Performance Issues**: Reduce look-ahead period or increase sync interval
4. **Missing Calendars**: Verify calendar sharing permissions and connection status

### Debug Mode
Enable debug logging in the console to troubleshoot sync issues:
```javascript
localStorage.setItem('calendarSync_debug', 'true');
```

### Support
- Check the Activity Log in the add-in for recent sync operations
- Review browser console for detailed error messages
- Verify calendar permissions in Outlook settings

## Development Roadmap

### Version 1.1 (Current)
- ✅ Basic multi-calendar sync
- ✅ Private busy block creation
- ✅ Manual sync operations
- ✅ Calendar enable/disable controls

### Version 1.2 (Planned)
- ⏳ Advanced configuration options
- ⏳ Conflict resolution rules
- ⏳ Performance optimizations
- ⏳ Enhanced error handling

### Version 2.0 (Future)
- 📋 Google Calendar integration
- 📋 Mobile app companion
- 📋 Team calendar sharing
- 📋 Analytics and reporting

## Contributing

This project welcomes contributions! Please see CONTRIBUTING.md for guidelines on:
- Code style and standards
- Testing requirements
- Pull request process
- Issue reporting

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Acknowledgments

- Microsoft Office Add-ins team for the excellent development framework
- Microsoft Graph API for robust calendar integration
- The Fluent UI team for beautiful and accessible design components
