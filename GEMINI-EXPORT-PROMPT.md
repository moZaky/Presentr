# Gemini Export Prompt - PowerPoint AutoFill Project

## 🎯 PROJECT OVERVIEW

You are tasked with creating a complete Windows-compatible PowerPoint automation solution that transforms Excel data into professional presentations with charts. The solution must be 100% offline, private, and work without any server dependencies.

## 📋 PROJECT REQUIREMENTS

### Core Functionality
- Create a PowerPoint AutoFill application that reads Excel data and automatically fills PowerPoint templates
- Support placeholder replacement (e.g., {name}, {project_title}, {status_update}, {agenda}, {ChartTitle})
- Generate automatic charts (bar, line, pie) from Excel ChartData sheet
- Must work completely offline - no internet connection required
- 100% private processing - files never leave user's computer
- Windows-compatible with modern browser support

### Input Data Format
**PowerPoint Template (.pptx):**
- Contains placeholder tags in curly braces: {name}, {project_title}, {status_update}, {agenda}, {ChartTitle}
- Professional template structure with title slide, overview, agenda, and chart placeholder slides

**Excel Data (.xlsx):**
- **Sheet1**: Main data with columns matching placeholders
  - Example: name=zinab, project_title=DGE Review, status_update="All financial models have been updated...", agenda="Technology Stack System Context Design Decisions Mobile Application Features", ChartTitle="Analysis of Key Metrics"
- **ChartData** (optional): Chart data with columns:
  - ChartTitle: Chart name
  - ChartType: bar, line, pie
  - Series: Data series name  
  - Label: X-axis label
  - Value: Data value

### Technical Requirements
- Pure frontend implementation (HTML, CSS, JavaScript)
- Use PptxGenJS library for PowerPoint generation
- Use XLSX.js library for Excel processing
- Modern, responsive UI with gradient design
- Progress indicators and error handling
- Sample file generation for testing

## 🎨 DESIGN SPECIFICATIONS

### Visual Design
- **Color Scheme**: Purple gradient background (linear-gradient(135deg, #667eea 0%, #764ba2 100%))
- **Layout**: Modern card-based design with glassmorphism effects
- **Typography**: Clean, professional fonts (system fonts)
- **Animations**: Smooth transitions and hover effects
- **Responsive**: Mobile-friendly design

### User Interface Components
1. **Header Section**: App title, description, feature highlights
2. **Upload Section**: Two-column layout for PPT and Excel file upload
3. **Process Section**: Central processing button with progress indicators
4. **Result Section**: Success/error states with download functionality
5. **Sample Section**: Sample file download buttons

### User Experience Flow
1. User uploads PowerPoint template
2. User uploads Excel data file
3. Process button becomes enabled
4. User clicks process → shows progress bar
5. Generated PowerPoint downloads automatically
6. Success message with download option

## 🔧 TECHNICAL IMPLEMENTATION

### File Structure
```
Presentr-Project/
├── presentr-windows.html    # Main application
├── start-windows.bat        # Windows launcher
├── README.md               # Quick start guide
├── WINDOWS-SETUP.md        # Detailed instructions
└── PROJECT-SUMMARY.md      # Technical documentation
```

### Key JavaScript Functions
```javascript
// File handling
- readExcelFile(file): Parse Excel data and extract both main data and chart data
- readPowerPointTemplate(file): Read PPTX template (simplified implementation)
- generatePresentation(excelData, template): Create new presentation with replaced placeholders
- addCharts(pptx, excelData): Generate charts from ChartData sheet

// UI functions  
- processFiles(): Main processing orchestration
- updateProgress(percent, text): Progress bar updates
- downloadPresentation(): Trigger file download
- downloadSamplePPT()/downloadSampleExcel(): Generate sample files
- resetApp(): Reset application state
```

### Chart Generation Logic
- Group chart data by ChartTitle
- Support multiple chart types (bar, line, pie)
- Handle multiple data series per chart
- Automatic sizing and positioning
- Professional color schemes

### Error Handling
- File format validation
- Missing placeholder handling
- Chart data validation
- User-friendly error messages
- Retry functionality

## 📊 SAMPLE DATA IMPLEMENTATION

### Sample PowerPoint Structure
```javascript
// Title slide
- {project_title} (large, centered)
- "Status Report prepared by: {name}" (smaller, centered)

// Overview slide  
- "Overview" (section header)
- {status_update} (body text)

// Agenda slide
- "Agenda" (section header)
- Bullet points from {agenda} (split by spaces)

// Chart slides (dynamic)
- {ChartTitle} (slide title)
- Generated chart from ChartData
```

### Sample Excel Data
```javascript
// Sheet1 - Main Data
['name', 'project_title', 'status_update', 'agenda', 'ChartTitle']
['zinab', 'DGE Review', 'All financial models have been updated...', 'Technology Stack System Context...', 'Analysis of Key Metrics']

// ChartData - Chart Information  
['ChartTitle', 'ChartType', 'Series', 'Label', 'Value']
['Quarterly Sales', 'bar', 'Sales', 'Q1', 150000]
['Quarterly Sales', 'bar', 'Sales', 'Q2', 200000]
['Engagement Metrics', 'line', 'Likes', 'Jan', 1200]
['Engagement Metrics', 'line', 'Shares', 'Jan', 400]
```

## 🚀 DEPLOYMENT INSTRUCTIONS

### Windows Launcher Script
```batch
@echo off
echo ====================================
echo     Presentr - PowerPoint AutoFill
echo ====================================
echo.
echo Starting PowerPoint AutoFill Application...
echo This will open in your default browser
echo.

REM Check if HTML file exists
if not exist "presentr-windows.html" (
    echo ERROR: presentr-windows.html not found!
    pause
    exit /b 1
)

REM Open the HTML file in default browser
start presentr-windows.html

echo.
echo Application launched successfully!
echo.
pause
```

### Usage Instructions
1. Download all files to same directory
2. Double-click `start-windows.bat` OR open `presentr-windows.html`
3. Upload PowerPoint template with {placeholders}
4. Upload Excel file with matching data
5. Click "Process Files"
6. Download generated presentation

## 🎯 SUCCESS CRITERIA

### Functional Requirements
- ✅ Successfully reads Excel files (.xlsx format)
- ✅ Extracts both main data and chart data
- ✅ Replaces all placeholder tags correctly
- ✅ Generates professional charts (bar, line, pie)
- ✅ Creates downloadable PowerPoint file
- ✅ Works completely offline
- ✅ No data leaves user's computer

### Quality Requirements  
- ✅ Modern, professional UI design
- ✅ Responsive layout for all screen sizes
- ✅ Smooth animations and transitions
- ✅ Comprehensive error handling
- ✅ Clear user instructions
- ✅ Sample files for testing

### Technical Requirements
- ✅ Pure frontend implementation
- ✅ No server dependencies
- ✅ Cross-browser compatibility
- ✅ Windows optimization
- ✅ Performance optimization
- ✅ Code quality and maintainability

## 🔒 PRIVACY & SECURITY

### Data Protection
- All file processing happens in browser memory
- No data uploaded to external servers
- No internet connection required
- No user tracking or analytics
- Files are not stored permanently

### Security Considerations
- Input validation for file types
- Safe file handling practices
- No sensitive data exposure
- Secure file download mechanism

## 📈 TARGET USE CASES

### Business Applications
- Monthly status reports
- Sales presentations with charts
- Project updates for stakeholders
- Executive summaries
- Client proposals

### Data Visualization
- Financial reports with charts
- KPI dashboards
- Performance metrics
- Trend analysis
- Comparative analysis

## 🎨 DESIGN INSPIRATION

### Visual Style
- Modern gradient backgrounds
- Glassmorphism effects
- Clean typography
- Professional color palette
- Intuitive iconography

### Interaction Design
- Drag-and-drop file upload
- Real-time progress feedback
- Smooth state transitions
- Clear call-to-action buttons
- Helpful error messages

---

## 🚀 IMPLEMENTATION TASK LIST

### Phase 1: Core Structure
1. Create HTML structure with semantic markup
2. Implement CSS styling with gradient design
3. Set up responsive layout system
4. Add basic JavaScript framework

### Phase 2: File Processing
1. Implement Excel file reading functionality
2. Add PowerPoint template processing
3. Create placeholder replacement logic
4. Implement chart generation system

### Phase 3: User Interface
1. Add file upload components
2. Create progress indicators
3. Implement success/error states
4. Add sample file generation

### Phase 4: Polish & Optimization
1. Add animations and transitions
2. Implement comprehensive error handling
3. Optimize performance
4. Add accessibility features

### Phase 5: Documentation & Deployment
1. Create user documentation
2. Write Windows launcher script
3. Test across browsers
4. Final quality assurance

---

## 🎯 FINAL DELIVERABLES

1. **presentr-windows.html** - Complete standalone application
2. **start-windows.bat** - Windows launcher script
3. **README.md** - Quick start documentation
4. **WINDOWS-SETUP.md** - Detailed setup guide
5. **PROJECT-SUMMARY.md** - Technical documentation

The solution should be production-ready, user-friendly, and completely self-contained without any external dependencies beyond the CDN-hosted libraries (XLSX.js and PptxGenJS).

---

*This prompt contains all necessary information to recreate the complete PowerPoint AutoFill solution with the same functionality, design, and user experience as the original project.*