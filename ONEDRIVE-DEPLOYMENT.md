# OneDrive/SharePoint Quick Deployment Guide

## 🚀 Fastest Internal Deployment (5 Minutes)

Since you want to use OneDrive/SharePoint for internal hosting, here's the quickest approach:

## Step 1: Build Production Files

```bash
# Build the production version
npm run build
```

This creates a `dist` folder with all the files you need.

## Step 2: Upload to SharePoint Online

### Via Web Browser:
1. **Go to SharePoint**: `https://yourcompany.sharepoint.com`
2. **Create or go to a site**: Like "IT Apps" or "Outlook Add-ins"
3. **Create a new Document Library**: Call it "MultiCalendar"
4. **Upload ALL files** from your `dist` folder:
   - `taskpane.html`
   - `taskpane.js`
   - `commands.html`
   - `commands.js`
   - All files in `assets/` folder
   - `manifest.json`

### Via OneDrive Sync:
1. **Sync your SharePoint library** to your local OneDrive
2. **Copy all `dist` folder contents** to the synced folder
3. **Wait for sync** (files appear on SharePoint automatically)

## Step 3: Get the Web URLs

After uploading, your files will be accessible at URLs like:
```
https://yourcompany.sharepoint.com/sites/your-site/Shared%20Documents/MultiCalendar/taskpane.html
https://yourcompany.sharepoint.com/sites/your-site/Shared%20Documents/MultiCalendar/commands.html
https://yourcompany.sharepoint.com/sites/your-site/Shared%20Documents/MultiCalendar/commands.js
```

## Step 4: Update Manifest URLs

### Quick PowerShell Script:
```powershell
# Replace with your actual SharePoint base URL
$baseUrl = "https://yourcompany.sharepoint.com/sites/your-site/Shared%20Documents/MultiCalendar"

# Update manifest.json
$manifest = Get-Content "manifest.json" | ConvertFrom-Json

# Update the runtime URLs (find the specific sections and update)
# You'll need to manually edit these in the JSON file:
# "page": "$baseUrl/taskpane.html"
# "page": "$baseUrl/commands.html" 
# "script": "$baseUrl/commands.js"
```

### Manual Edit:
Open `manifest.json` and find these lines:

```json
// FIND:
"page": "https://localhost:3000/taskpane.html"

// REPLACE WITH:
"page": "https://yourcompany.sharepoint.com/sites/your-site/Shared%20Documents/MultiCalendar/taskpane.html"
```

```json
// FIND:
"page": "https://localhost:3000/commands.html",
"script": "https://localhost:3000/commands.js"

// REPLACE WITH:
"page": "https://yourcompany.sharepoint.com/sites/your-site/Shared%20Documents/MultiCalendar/commands.html",
"script": "https://yourcompany.sharepoint.com/sites/your-site/Shared%20Documents/MultiCalendar/commands.js"
```

## Step 5: Deploy to Organization

1. **Go to Microsoft 365 Admin Center**: `https://admin.microsoft.com`
2. **Navigate**: Settings > Integrated apps
3. **Click**: "Upload custom apps"
4. **Upload**: Your updated `manifest.json`
5. **Configure**: Deploy to "Everyone" with auto-deploy enabled
6. **Click**: "Deploy"

## 🎯 Even Simpler: Test with File URLs First

### For immediate testing, you can even use OneDrive direct file URLs:

1. **Upload files to any OneDrive folder**
2. **Right-click each file** > "Copy link" > "People with existing access"
3. **Use those URLs** in your manifest for testing

Example URLs look like:
```
https://yourcompany-my.sharepoint.com/personal/your_email_com/_layouts/15/download.aspx?share=...
```

**Note**: These are longer URLs but work for testing!

## ✅ What This Achieves

- ✅ **Completely internal** to your organization
- ✅ **No external hosting** required
- ✅ **Uses your existing SharePoint/OneDrive**
- ✅ **Proper HTTPS URLs** that Outlook can access
- ✅ **Easy to manage and update**

## 🚀 Ready to Try?

The key steps are:
1. **Build**: `npm run build`
2. **Upload**: Copy `dist` contents to SharePoint/OneDrive
3. **Update**: Change URLs in `manifest.json`
4. **Deploy**: Upload manifest to Admin Center

Would you like me to help you with the specific URL format for your organization or walk through updating the manifest file?
