# SharePoint Internal Deployment Guide

## 🏢 Deploy MultiCalendar Sync to SharePoint (Internal Only)

Perfect for organization administrators who want to keep everything internal and secure.

## 🚀 SharePoint Deployment Steps

### Step 1: Create SharePoint Site/Library
1. Go to your **SharePoint Admin Center**
2. Create a new site or use existing: `/sites/apps` or `/sites/outlook-addins`
3. Create a **Document Library** called "MultiCalendar"

### Step 2: Upload Add-in Files
1. **Build the production files**:
   ```bash
   npm run build
   ```

2. **Upload all files from the `dist` folder** to SharePoint:
   - `taskpane.html`
   - `taskpane.js` 
   - `taskpane.css`
   - `commands.html`
   - `commands.js`
   - `assets/` folder (all icon files)
   - `manifest.json`

### Step 3: Update Manifest for SharePoint URLs

Replace the localhost URLs in your `manifest.json`:

```json
// BEFORE (Development):
"page": "https://localhost:3000/taskpane.html"

// AFTER (SharePoint):
"page": "https://yourcompany.sharepoint.com/sites/apps/MultiCalendar/taskpane.html"
```

### Step 4: Configure SharePoint Permissions
- **Site Permissions**: All users in organization
- **Library Permissions**: Read access for all users
- **External Sharing**: Disabled (internal only)

### Step 5: Deploy via Microsoft 365 Admin Center
1. Go to **Microsoft 365 Admin Center**
2. **Settings** > **Integrated apps**
3. **Upload custom apps**
4. Upload your **SharePoint-updated manifest.json**
5. Deploy to organization

## 🔒 Security Benefits

### Internal Only Access
- ✅ **No public internet exposure**
- ✅ **Only your organization can access**
- ✅ **Uses existing SharePoint security**
- ✅ **Integrated with your Microsoft 365 tenant**

### Compliance
- ✅ **Meets corporate security policies**
- ✅ **No external dependencies**
- ✅ **Audit logs in SharePoint**
- ✅ **Data sovereignty maintained**

## 📋 SharePoint URLs Format

Your final URLs will look like:
```
https://yourcompany.sharepoint.com/sites/apps/MultiCalendar/taskpane.html
https://yourcompany.sharepoint.com/sites/apps/MultiCalendar/commands.html
https://yourcompany.sharepoint.com/sites/apps/MultiCalendar/commands.js
```

## ⚡ Quick Setup Commands

### Update Manifest for SharePoint (PowerShell):
```powershell
# Replace with your actual SharePoint URL
$sharepointUrl = "https://yourcompany.sharepoint.com/sites/apps/MultiCalendar"
$manifestPath = "manifest.json"

# Update all localhost URLs to SharePoint
(Get-Content $manifestPath) -replace "https://localhost:3000", $sharepointUrl | Set-Content $manifestPath

# Update valid domains
(Get-Content $manifestPath) -replace "contoso\.com", "yourcompany.sharepoint.com" | Set-Content $manifestPath
```

### Validate Updated Manifest:
```bash
npm run validate
```

## 🎯 Alternative: OneDrive for Business

If SharePoint sites are restricted, you can also use **OneDrive for Business**:

1. **Upload files to OneDrive for Business**
2. **Share with organization**
3. **Get shareable links**
4. **Update manifest with OneDrive URLs**

Example URL: `https://yourcompany-my.sharepoint.com/personal/admin_yourcompany_com/_layouts/15/download.aspx?share=...`

## ✅ Advantages of Internal Hosting

### Security
- **No external attack surface**
- **Corporate firewall protection**
- **Integrated authentication**

### Compliance
- **Meets internal security policies**
- **Data stays within organization**
- **Audit trail maintained**

### Management
- **IT department controls hosting**
- **Easy to update/maintain**
- **Integrated with existing systems**

## 🚀 Ready to Deploy Internally?

SharePoint internal deployment gives you:
- ✅ **Complete privacy and security**
- ✅ **No external dependencies**
- ✅ **Easy management and updates**
- ✅ **Fast deployment to your organization**

Would you like me to help you set up the SharePoint deployment or create the updated manifest for your internal URLs?
