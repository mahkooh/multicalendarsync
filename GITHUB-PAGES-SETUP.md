# GitHub Pages Setup for MultiCalendar Sync Add-in

## Quick GitHub Pages Configuration

Your code has been successfully pushed to GitHub! Now follow these steps to enable GitHub Pages hosting:

### 1. Enable GitHub Pages
1. Go to your repository: https://github.com/mahkooh/multicalendarsync
2. Click on **Settings** tab
3. Scroll down to **Pages** section in the left sidebar
4. Under **Source**, select **Deploy from a branch**
5. Choose **main** branch
6. Select **/ (root)** folder
7. Click **Save**

### 2. Verify Deployment
- GitHub will automatically deploy your site to: `https://mahkooh.github.io/multicalendarsync`
- It may take 5-10 minutes for the site to be available
- You'll see a green checkmark in your repository's Actions tab when deployment is complete

### 3. Test the Add-in
Once GitHub Pages is active, test these URLs:
- Main page: `https://mahkooh.github.io/multicalendarsync/dist/taskpane.html`
- Manifest: `https://mahkooh.github.io/multicalendarsync/dist/manifest.xml`

## Deploy to Microsoft 365 Admin Center

### 1. Access Admin Center
1. Go to https://admin.microsoft.com
2. Navigate to **Settings** > **Integrated apps**
3. Click **Upload custom apps**

### 2. Upload Manifest
1. Select **Upload manifest file**
2. Download the manifest from: `https://mahkooh.github.io/multicalendarsync/dist/manifest.xml`
3. Upload this file to the admin center

### 3. Configure Deployment
1. Choose **Deploy to entire organization** 
2. Review permissions (Calendar read/write access)
3. Click **Deploy**

### 4. Verification
- Deployment typically takes 6-24 hours
- Users will see "MultiCalendar Sync" in their Outlook add-ins
- Check deployment status in the admin center

## Alternative: Direct Sideloading for Testing

If you want to test immediately without waiting for organization deployment:

### 1. Download Manifest
- Save manifest from: `https://mahkooh.github.io/multicalendarsync/dist/manifest.xml`

### 2. Sideload in Outlook
1. Open Outlook on the web or desktop
2. Go to **Get Add-ins** > **My add-ins** > **Add a custom add-in**
3. Choose **Add from file**
4. Upload the downloaded manifest.json

## Troubleshooting

### Common Issues
1. **HTTPS Required**: GitHub Pages provides HTTPS automatically
2. **CORS Enabled**: The manifest includes proper CORS configuration
3. **Manifest Validation**: All URLs point to GitHub Pages domain

### Verification Steps
```bash
# Test manifest accessibility
curl -I https://mahkooh.github.io/multicalendarsync/dist/manifest.xml

# Test main taskpane
curl -I https://mahkooh.github.io/multicalendarsync/dist/taskpane.html
```

## Next Steps
1. ✅ Code pushed to GitHub
2. ⏳ Enable GitHub Pages (follow steps above)
3. ⏳ Upload to Microsoft 365 Admin Center
4. ⏳ Deploy to organization
5. ⏳ Test with users

Your add-in is production-ready and configured for GitHub Pages hosting!
