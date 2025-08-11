# Publishing Guide - Microsoft AppSource

## 🚀 Publishing Your MultiCalendar Sync Add-in

### Prerequisites for AppSource Submission

#### 1. **Microsoft Partner Center Account**
- Register at [Microsoft Partner Center](https://partner.microsoft.com/)
- Complete publisher verification (can take 1-2 weeks)
- Pay one-time registration fee ($99 USD for individual developers)

#### 2. **Azure App Registration**
- Register your app in [Azure Portal](https://portal.azure.com/)
- Configure OAuth 2.0 permissions
- Set up redirect URIs for production

#### 3. **Production Hosting**
- Deploy your add-in to a production web server
- Ensure HTTPS is properly configured
- Update manifest.json with production URLs

### Step-by-Step Publishing Process

#### Phase 1: Prepare Production Build

1. **Update Manifest for Production**
   ```json
   {
     "runtimes": [{
       "code": {
         "page": "https://yourdomain.com/taskpane.html"
       }
     }]
   }
   ```

2. **Build Production Package**
   ```bash
   npm run build
   ```

3. **Test with Production Manifest**
   - Update all localhost URLs to production URLs
   - Test sideloading with production manifest

#### Phase 2: Azure App Registration

1. **Register Application**
   - Go to Azure Portal > App Registrations
   - Create new registration
   - Note Application (client) ID

2. **Configure Permissions**
   - Add Microsoft Graph permissions:
     - `Calendars.ReadWrite`
     - `Calendars.ReadWrite.Shared`
     - `User.Read`

3. **Set Redirect URIs**
   - Add your production domain
   - Configure for single-page application

#### Phase 3: Partner Center Submission

1. **Create New Office Add-in**
   - Login to Partner Center
   - Go to Office Store > Overview
   - Click "Create a new add-in"

2. **Upload Manifest**
   - Upload your production manifest.json
   - Validate all URLs are accessible

3. **Complete Store Listing**
   - App name: "MultiCalendar Sync"
   - Short description: "Synchronize busy time across multiple calendars"
   - Long description: Include features, benefits, privacy info
   - Screenshots: App interface, features in action
   - Category: Productivity
   - Supported languages: English (add more as needed)

4. **Privacy & Compliance**
   - Privacy policy URL (required)
   - Terms of use URL (required)
   - Data handling declarations
   - Age rating information

#### Phase 4: Technical Validation

Microsoft will test:
- ✅ Manifest validation
- ✅ All URLs accessible
- ✅ Core functionality works
- ✅ No security vulnerabilities
- ✅ Follows Office Store policies

### Required Documentation

#### 1. **Privacy Policy** (Required)
Create at `https://yourdomain.com/privacy-policy`

#### 2. **Terms of Use** (Required)
Create at `https://yourdomain.com/terms-of-use`

#### 3. **Support Documentation**
- User guide
- Installation instructions
- Troubleshooting guide

### Timeline & Costs

#### Development Phase
- **Duration**: 2-4 weeks (depending on Graph API integration)
- **Cost**: Development time only

#### Registration & Hosting
- **Partner Center**: $99 USD one-time fee
- **Azure hosting**: $10-50/month (depending on usage)
- **Domain/SSL**: $10-20/year

#### Review Process
- **Initial review**: 3-5 business days
- **Certification**: 1-2 weeks after submission
- **Total time**: 2-4 weeks from submission

### Alternative Distribution Options

#### 1. **Organization Catalog** (Faster)
- Deploy only to your organization
- No AppSource review required
- Admin can install for all users
- Timeline: 1-2 days

#### 2. **Direct Sideloading** (Immediate)
- Share manifest.json directly
- Users manually install
- Good for testing/pilot groups
- Timeline: Immediate

#### 3. **Teams App Store** (If applicable)
- Package as Teams app with Outlook support
- Different submission process
- Broader Microsoft 365 integration

### Recommended Approach

For **immediate use** while preparing for AppSource:

1. **Start with Organization Catalog**
   - Deploy to your company first
   - Get user feedback
   - Refine features

2. **Prepare AppSource Submission**
   - Register Partner Center account
   - Set up production hosting
   - Create required documentation

3. **Submit to AppSource**
   - Once stable and tested
   - For broader distribution
   - Professional credibility

### Next Steps for Your Add-in

1. **Complete Graph API Integration**
   - Replace simulation with real calendar access
   - Implement proper authentication

2. **Set Up Production Environment**
   - Choose hosting provider (Azure, AWS, etc.)
   - Configure HTTPS and domain

3. **Create Required Pages**
   - Privacy policy
   - Terms of use
   - Support documentation

4. **Register Partner Center Account**
   - Start verification process early
   - Prepare business information

Would you like me to help you with any specific part of this process?
