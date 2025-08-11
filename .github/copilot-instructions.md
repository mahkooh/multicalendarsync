# Copilot Instructions for MultiCalendar Outlook Add-in

<!-- Use this file to provide workspace-specific custom instructions to Copilot. For more details, visit https://code.visualstudio.com/docs/copilot/copilot-customization#_use-a-githubcopilotinstructionsmd-file -->

## Project Overview
This is an Outlook add-in that synchronizes busy time across multiple calendars to ensure consistent availability visibility to colleagues. The add-in creates private "busy" blocks in calendars when the user has appointments in other calendars, without cluttering the user's own view.

## Core Functionality
- Monitor all calendars in the user's Outlook
- Detect busy time from each calendar
- Create private/hidden busy blocks in other calendars
- Maintain user privacy (no meeting details shared)
- Provide clean calendar view for the user
- Enable manual control and configuration

## Technical Stack
- **Frontend**: Office.js API, HTML, CSS, JavaScript
- **Authentication**: Microsoft Graph API
- **Calendar Access**: Exchange Web Services (EWS) or Graph API
- **Permissions**: MailboxItem.ReadWrite.User, Calendars.ReadWrite

## Key Design Principles
1. **Privacy First**: Never expose meeting details across calendars
2. **User Experience**: Keep the user's calendar view clean
3. **Reliability**: Robust error handling and conflict resolution
4. **Performance**: Efficient calendar monitoring and synchronization
5. **Configurability**: Allow users to control sync behavior

## Code Guidelines
- Use modern JavaScript (ES6+) features
- Implement proper error handling for Office.js APIs
- Follow Microsoft Graph API best practices
- Use Office UI Fabric for consistent styling
- Implement debouncing for real-time sync operations
- Add comprehensive logging for debugging

## Calendar Sync Logic
- Create busy blocks with `ShowAs: "Busy"` and `IsPrivate: true`
- Use specific subject line pattern for identification (e.g., "Busy (Auto-Sync)")
- Implement conflict detection to avoid duplicate sync blocks
- Handle different calendar types (Exchange, Office 365)
- Support configurable sync intervals

## Security Considerations
- Validate all Graph API responses
- Implement proper token refresh handling
- Use least-privilege permissions
- Sanitize any user input
- Implement rate limiting for API calls
