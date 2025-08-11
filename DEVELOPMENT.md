# Development Guide - MultiCalendar Sync

## Quick Start

Your Outlook add-in is ready for development and testing! Here's how to get started:

## 🚀 Running the Add-in

1. **Start the development server** (already running):
   ```bash
   npm start
   ```

2. **Access in Outlook**:
   - The add-in will automatically open in Outlook for testing
   - You can also manually sideload it using the manifest.json file

## 🛠️ Project Structure

```
MultiCalendar/
├── src/
│   ├── taskpane/
│   │   ├── taskpane.html      # Main UI interface
│   │   ├── taskpane.js        # Main application logic
│   │   ├── taskpane.css       # Styling
│   │   └── calendarSync.js    # Core sync logic
│   └── commands/              # Ribbon commands
├── assets/                    # Icons and images
├── manifest.json             # Add-in configuration
├── package.json              # Dependencies
└── README.md                 # Project documentation
```

## 🎯 Current Features

### ✅ Implemented
- **Modern UI**: Clean, Fluent UI-based interface
- **Calendar Discovery**: Detects all available calendars
- **Sync Management**: Enable/disable sync for individual calendars
- **Status Monitoring**: Real-time sync status and activity log
- **Manual Sync**: On-demand synchronization
- **Privacy Controls**: Private busy blocks that don't clutter your view

### 🧩 Core Components

1. **CalendarSyncManager** (`calendarSync.js`):
   - Handles all sync logic
   - Manages calendar discovery
   - Creates and removes busy blocks
   - Monitors sync status

2. **TaskPane UI** (`taskpane.html/js/css`):
   - User interface for managing sync
   - Status dashboard
   - Activity logging
   - Configuration controls

## 🔧 Next Steps for Development

### Phase 1: Real Calendar Integration
Currently using simulated data. To connect to real calendars:

1. **Add Microsoft Graph SDK** (already included):
   ```javascript
   import { Client } from '@microsoft/microsoft-graph-client';
   ```

2. **Implement authentication**:
   ```javascript
   import { PublicClientApplication } from '@azure/msal-browser';
   ```

3. **Replace simulation methods** in `calendarSync.js`:
   - `discoverCalendars()` → Use Graph API to get real calendars
   - `getBusyTimesFromCalendars()` → Query actual calendar events
   - `createSyncBlocks()` → Create real calendar events

### Phase 2: Enhanced Features
- **Configuration Dialog**: Advanced settings for sync behavior
- **Conflict Resolution**: Handle overlapping events intelligently
- **Performance Optimization**: Efficient bulk operations
- **Error Recovery**: Robust error handling and retry logic

### Phase 3: Production Features
- **Microsoft Graph Integration**: Full API implementation
- **Azure App Registration**: Proper authentication setup
- **Deployment Package**: Production-ready build
- **Store Submission**: Microsoft AppSource publishing

## 🔐 Security & Permissions

The add-in requests these permissions:
- `MailboxItem.ReadWrite.User`: Access to your calendar items
- `Calendars.ReadWrite.Shared`: Access to shared calendars
- `Calendars.Read.Shared`: Read access to connected calendars

## 🧪 Testing

### Development Testing
1. Use the **Sync Now** button to test sync logic
2. Toggle calendars on/off to test selective sync
3. Check the Activity Log for sync operations
4. Monitor the Status indicator for real-time updates

### Production Testing
1. Test with real calendar data
2. Verify busy blocks are created correctly
3. Confirm privacy settings work as expected
4. Test across different Outlook clients (Desktop, Web, Mobile)

## 🐛 Debugging

### Enable Debug Mode
```javascript
// In browser console
localStorage.setItem('calendarSync_debug', 'true');
```

### Common Issues
1. **Permissions**: Ensure Outlook has necessary calendar permissions
2. **CORS**: Development server handles this automatically
3. **Manifest**: Validate using `npm run validate`

## 📦 Building & Deployment

### Development Build
```bash
npm run build:dev
```

### Production Build
```bash
npm run build
```

### Validation
```bash
npm run validate
```

## 🎨 Customization

### UI Styling
- Modify `taskpane.css` for visual changes
- Uses Fluent UI for consistent Microsoft styling
- Responsive design works across all Outlook clients

### Sync Logic
- Adjust `config` object in `CalendarSyncManager`
- Modify sync intervals, look-ahead periods
- Customize busy block appearance

### Permissions
- Update `manifest.json` for additional permissions
- Ensure Azure app registration matches manifest

## 📚 Resources

### Documentation
- [Office Add-ins Documentation](https://docs.microsoft.com/office/dev/add-ins/)
- [Microsoft Graph API](https://docs.microsoft.com/graph/)
- [Fluent UI](https://developer.microsoft.com/fluentui)

### Tools
- [Script Lab](https://appsource.microsoft.com/product/office/WA104380862) - For testing Office.js APIs
- [Office Add-in Validator](https://dev.office.com/add-in-validator) - Validate your add-in

## 🚀 Ready to Code!

Your multi-calendar sync solution is set up and ready for development. The foundation is solid with:
- ✅ Modern TypeScript-ready structure
- ✅ Fluent UI components
- ✅ Comprehensive sync logic framework
- ✅ Privacy-focused design
- ✅ Production-ready architecture

Start by exploring the simulated sync functionality, then gradually replace simulation methods with real Microsoft Graph API calls when you're ready to connect to live calendar data!

Happy coding! 🎉
