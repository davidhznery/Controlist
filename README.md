# Controlist - Google Apps Script Project Management System

A comprehensive Google Apps Script solution for managing projects with automated folder creation, KPI dashboards, and project tracking for SOS and DLE companies.

## 🚀 Features

### Project Management
- **Automated Project Creation**: Create project folders with standardized structure
- **Multi-Company Support**: Handles both SOS and DLE companies with different workflows
- **Template System**: Automatic file copying from predefined templates
- **Project Numbering**: Auto-generates sequential project numbers (e.g., 001-25, 002-25)

### Dashboard & Analytics
- **KPI Dashboard**: Real-time project status tracking and analytics
- **Multi-Year Support**: Tracks projects from 2024 and 2025
- **Status Categorization**: Automatically categorizes projects by status
- **Interactive Reports**: Click on status counts to see detailed project lists

### Data Management
- **Client Validation**: Smart client name normalization and validation
- **Status Tracking**: Comprehensive status management with flexible matching
- **Data Validation**: Prevents invalid data entry with smart suggestions

## 📁 Project Structure

```
Controlist/
├── code.gs.js                 # Main Google Apps Script code
├── simple-dashboard.html      # KPI Dashboard interface
├── sidebar-new-project.html   # New project creation form
├── appscript.json            # Apps Script configuration
└── README.md                 # This file
```

## 🛠️ Setup Instructions

### 1. Google Apps Script Setup
1. Open [Google Apps Script](https://script.google.com)
2. Create a new project
3. Copy the contents of `code.gs.js` into `Code.gs`
4. Create HTML files:
   - `simple-dashboard.html` (copy from the file)
   - `sidebar-new-project.html` (copy from the file)
5. Update the `appscript.json` with your configuration

### 2. Google Sheets Configuration
1. Create a Google Sheet with two tabs: "SOS" and "DLE"
2. Set up the following columns:
   - **SOS**: Proj., No.*, Type, Days for Deadline, *Client*, Client Ref No, Received, Month, *Subject*, Deadline, Month, Status, LEADER PJ
   - **DLE**: Proj., No.*, Division, Days for Deadline, *Client*, Client Ref No, Received, Month, *Subject*, Deadline, Month, Status, LEADER PJ

### 3. Google Drive Setup
1. Create parent folders for SOS and DLE projects
2. Update the folder IDs in the configuration section of `code.gs.js`:
   ```javascript
   const SOS_PARENT_FOLDER_ID = "your-sos-folder-id";
   const DLE_PARENT_FOLDER_ID = "your-dle-folder-id";
   ```

## 🎯 Usage

### Creating New Projects
1. In Google Sheets, go to **Projects** → **New Project…**
2. Fill in the project details:
   - Company (SOS or DLE)
   - Division (for DLE)
   - Project number (auto-generated)
   - Client, Subject, Leader, Status
   - Received and Deadline dates
3. Click **Create Project**

### Viewing Analytics
1. Go to **Reports** → **📊 Simple KPI Dashboard…**
2. View real-time project statistics
3. Click on any status count to see detailed project lists

### Debugging
- **Reports** → **🔍 Debug All Awarded Projects…**: View all awarded projects by year
- **Reports** → **🔍 Debug Dashboard Data…**: Check dashboard data processing
- **Reports** → **🔄 Clear Cache…**: Clear cached data and refresh

## 📊 Dashboard Features

### Project Status Tracking
- **RFQ Process**: Pending creation, sending, waiting for quotations
- **Proposal Process**: Ready to quote, technical proposals, estimations
- **Project Not Quoted**: Technical queries, restrictions, cancellations
- **Awarded Projects**: Successfully completed projects

### KPI Metrics
- **Conversion Rates**: Awarded projects vs total projects
- **Offer Conversion**: Commercial proposals vs total projects
- **Project Distribution**: Status breakdown and trends

## 🔧 Configuration

### Folder Structure Templates
The system creates standardized folder structures:

**SOS Structure:**
- 1- Documents_&_Data_From_Client
- 2- RFQ
- 3- Supplier
- 4- Estimation
- 5- Contracts Agreements
- 6- Native files
- 7- Proposal
- 8- Sent_to_the_Client

**DLE Structure (varies by division):**
- ChampionX, Procurement, Trading divisions
- Similar structure with division-specific templates

### Status Categories
- **RFQ Process**: Waiting RFQ Creation, Pending RFQ Sending
- **Waiting for Quotation**: Various supplier waiting statuses
- **Proposal Process**: Ready to quote, technical proposals, estimations
- **Awarded**: All variations of "Awarded" status
- **Not Quoted**: Technical/commercial reasons, restrictions, cancellations

## 🚨 Troubleshooting

### Common Issues
1. **"Invalid Client" Error**: Use the client normalization feature
2. **Missing Awarded Projects**: Check year filtering and status matching
3. **Dashboard Not Updating**: Clear cache and refresh data

### Debug Tools
- Use the built-in debug functions in the Reports menu
- Check the browser console for JavaScript errors
- Verify Google Apps Script execution logs

## 📈 Recent Updates

### Multi-Year Support
- Dashboard now shows projects from both 2024 and 2025
- Awarded projects popup displays projects from both years
- Improved project sorting by year and project number

### Enhanced Status Detection
- Flexible matching for awarded projects
- Better handling of status variations
- Improved data validation and normalization

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Test thoroughly
5. Submit a pull request

## 📄 License

This project is licensed under the MIT License - see the LICENSE file for details.

## 👨‍💻 Author

**David Hernandez**
- GitHub: [@davidhznery](https://github.com/davidhznery)
- Repository: [Controlist](https://github.com/davidhznery/Controlist)

## 📞 Support

For support and questions:
1. Check the troubleshooting section
2. Use the built-in debug tools
3. Create an issue in the GitHub repository

---

*Last updated: January 2025*
