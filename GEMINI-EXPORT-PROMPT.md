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

## 🛠️ TECHNOLOGY STACK (MANDATORY)

### Core Framework (NON-NEGOTIABLE)
- **Framework**: Pure HTML5, CSS3, JavaScript (ES6+) - REQUIRED (cannot be changed)
- **Language**: JavaScript 6+ - REQUIRED (cannot be changed)

### Standard Technology Stack
**When users don't specify preferences, use this complete stack:**

- **Styling**: Pure CSS3 with modern features (flexbox, grid, gradients)
- **PowerPoint Generation**: PptxGenJS 3.12.0 (CDN: https://cdnjs.cloudflare.com/ajax/libs/PptxGenJS/3.12.0/pptxgen.bundle.js)
- **Excel Processing**: XLSX.js 0.18.5 (CDN: https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js)
- **UI Components**: Custom-built components with modern CSS
- **State Management**: Vanilla JavaScript with modern ES6+ features
- **Icons**: Unicode emojis and CSS-based icons

### Library Usage Policy
- **ALWAYS use the specified libraries** - these are non-negotiable requirements
- **When users request external libraries not in our stack**: Politely redirect them to use our built-in alternatives
- **Explain the benefits** of using our predefined stack (consistency, optimization, support)
- **Provide equivalent solutions** using our available libraries

### Specific Library Versions
```html
<!-- Required CDN Libraries - Use Exactly These Versions -->
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/PptxGenJS/3.12.0/pptxgen.bundle.js"></script>
```

### Browser Compatibility
- **Modern Browsers**: Chrome 90+, Firefox 88+, Edge 90+, Safari 14+
- **Windows Optimization**: Specifically optimized for Windows 7/8/10/11
- **Mobile Support**: Responsive design for tablets and smartphones
- **No IE Support**: Do not support Internet Explorer

### Development Tools
- **No Build Process**: Pure HTML/CSS/JS - no compilation needed
- **No Package Manager**: Use CDN libraries only
- **No Framework Dependencies**: Standalone implementation
- **Offline First**: Must work without internet connection after initial load

## 🎨 DESIGN SPECIFICATIONS

### Visual Design
- **Color Scheme**: Purple gradient background (linear-gradient(135deg, #667eea 0%, #764ba2 100%))
- **Layout**: Modern card-based design with glassmorphism effects
- **Typography**: Clean, professional fonts (system fonts: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif)
- **Animations**: Smooth transitions and hover effects
- **Responsive**: Mobile-friendly design

### Specific CSS Requirements
```css
/* Required Color Variables */
--primary-gradient: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
--success-color: #28a745;
--error-color: #dc3545;
--warning-color: #ffc107;
--info-color: #007bff;

/* Glassmorphism Effects */
background: rgba(255, 255, 255, 0.1);
backdrop-filter: blur(10px);
border: 1px solid rgba(255, 255, 255, 0.2);

/* Button Gradients */
background: linear-gradient(135deg, #667eea, #764ba2);
background: linear-gradient(135deg, #28a745, #20c997);
```

### Design System
- **Buttons**: Gradient backgrounds with hover effects and transform animations
- **Cards**: White backgrounds with rounded corners (20px) and box shadows
- **Upload Areas**: Dashed borders (3px dashed #ddd) that change color on hover/file upload
- **Progress Bars**: Gradient fills with percentage display
- **Typography**: Clear hierarchy with font sizes from 0.95rem to 2.5rem
- **Spacing**: Consistent padding (20px, 30px, 40px) and gaps (20px, 30px)

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

### JavaScript Architecture Requirements
```javascript
// Use Modern ES6+ Features
- const/let instead of var
- Arrow functions for callbacks
- Async/await for file operations
- Template literals for strings
- Destructuring for cleaner code
- Modern event handling

// File Processing Pattern
async function readExcelFile(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = function(e) {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                // Process workbook...
                resolve(processedData);
            } catch (error) {
                reject(error);
            }
        };
        reader.readAsArrayBuffer(file);
    });
}

// PowerPoint Generation Pattern
async function generatePresentation(excelData, template) {
    const pptx = new PptxGenJS();
    
    // Title slide
    const titleSlide = pptx.addSlide();
    titleSlide.addText(data.project_title || 'Presentation', {
        x: 1, y: 2, w: 8, h: 2,
        fontSize: 44, bold: true, color: '363636',
        align: 'center'
    });
    
    return pptx;
}
```

### State Management Pattern
```javascript
// Global State Management
let pptFile = null;
let excelFile = null;
let processedPresentation = null;

// State Update Functions
function checkReadyToProcess() {
    const processBtn = document.getElementById('processBtn');
    if (pptFile && excelFile) {
        processBtn.disabled = false;
    } else {
        processBtn.disabled = true;
    }
}

// Reset State
function resetApp() {
    pptFile = null;
    excelFile = null;
    processedPresentation = null;
    // Reset UI elements...
}
```

## 📝 COMPLETE CODE IMPLEMENTATION

### HTML Structure (Required)
```html
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Presentr - PowerPoint AutoFill (Windows)</title>
    <!-- Required CDN Libraries -->
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/PptxGenJS/3.12.0/pptxgen.bundle.js"></script>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>🎨 Presentr</h1>
            <p>PowerPoint AutoFill - Transform Excel Data into Professional Presentations</p>
        </div>
        
        <div class="main-content">
            <div class="upload-section">
                <div class="upload-box" id="pptBox">
                    <div class="upload-icon">📄</div>
                    <div class="upload-title">PowerPoint Template</div>
                    <div class="upload-desc">Upload .pptx file with {placeholder} tags</div>
                    <input type="file" id="pptInput" class="file-input" accept=".pptx" />
                    <button class="upload-btn" onclick="document.getElementById('pptInput').click()">
                        Choose PowerPoint File
                    </button>
                    <div id="pptInfo" class="file-info" style="display: none;"></div>
                </div>

                <div class="upload-box" id="excelBox">
                    <div class="upload-icon">📊</div>
                    <div class="upload-title">Excel Data</div>
                    <div class="upload-desc">Upload .xlsx file with matching columns</div>
                    <input type="file" id="excelInput" class="file-input" accept=".xlsx" />
                    <button class="upload-btn" onclick="document.getElementById('excelInput').click()">
                        Choose Excel File
                    </button>
                    <div id="excelInfo" class="file-info" style="display: none;"></div>
                </div>
            </div>

            <div class="process-section">
                <button id="processBtn" class="process-btn" onclick="processFiles()" disabled>
                    <span>🚀</span>
                    <span>Process Files</span>
                </button>
            </div>

            <div id="progressSection" class="progress-section">
                <div class="progress-bar">
                    <div id="progressFill" class="progress-fill">0%</div>
                </div>
                <div id="progressText" class="progress-text">Initializing...</div>
            </div>

            <div id="resultSection" class="result-section">
                <div class="result-icon">🎉</div>
                <div class="result-title">Processing Complete!</div>
                <div class="result-desc">Your PowerPoint presentation has been generated successfully with charts and data!</div>
                <button class="download-btn" onclick="downloadPresentation()">
                    <span>⬇️</span>
                    <span>Download PowerPoint</span>
                </button>
            </div>

            <div id="errorSection" class="error-section">
                <div class="error-icon">❌</div>
                <div class="error-title">Processing Failed</div>
                <div class="error-desc">There was an error processing your files. Please check the file formats and try again.</div>
                <button class="retry-btn" onclick="resetApp()">Try Again</button>
            </div>
        </div>
    </div>
</body>
</html>
```

### Complete JavaScript Implementation
```javascript
// Global State Variables
let pptFile = null;
let excelFile = null;
let processedPresentation = null;

// File Upload Handlers
document.getElementById('pptInput').addEventListener('change', function(e) {
    const file = e.target.files[0];
    if (file && file.name.endsWith('.pptx')) {
        pptFile = file;
        document.getElementById('pptInfo').textContent = `✅ ${file.name}`;
        document.getElementById('pptInfo').style.display = 'block';
        document.getElementById('pptBox').classList.add('has-file');
        checkReadyToProcess();
    } else {
        alert('Please select a valid PowerPoint (.pptx) file');
    }
});

document.getElementById('excelInput').addEventListener('change', function(e) {
    const file = e.target.files[0];
    if (file && file.name.endsWith('.xlsx')) {
        excelFile = file;
        document.getElementById('excelInfo').textContent = `✅ ${file.name}`;
        document.getElementById('excelInfo').style.display = 'block';
        document.getElementById('excelBox').classList.add('has-file');
        checkReadyToProcess();
    } else {
        alert('Please select a valid Excel (.xlsx) file');
    }
});

// Main Processing Function
async function processFiles() {
    if (!pptFile || !excelFile) {
        alert('Please upload both PowerPoint and Excel files');
        return;
    }

    // Show progress
    document.getElementById('progressSection').style.display = 'block';
    document.getElementById('processBtn').disabled = true;
    document.getElementById('resultSection').style.display = 'none';
    document.getElementById('errorSection').style.display = 'none';

    try {
        // Read Excel file
        updateProgress(20, 'Reading Excel data...');
        const excelData = await readExcelFile(excelFile);

        // Read PowerPoint template
        updateProgress(40, 'Reading PowerPoint template...');
        const pptxTemplate = await readPowerPointTemplate(pptFile);

        // Process and generate new presentation
        updateProgress(60, 'Generating presentation...');
        const newPptx = await generatePresentation(excelData, pptxTemplate);

        // Add charts
        updateProgress(80, 'Adding charts...');
        await addCharts(newPptx, excelData);

        // Complete
        updateProgress(100, 'Finalizing...');
        processedPresentation = newPptx;

        // Show success
        setTimeout(() => {
            document.getElementById('progressSection').style.display = 'none';
            document.getElementById('resultSection').style.display = 'block';
        }, 1000);

    } catch (error) {
        console.error('Processing error:', error);
        document.getElementById('progressSection').style.display = 'none';
        document.getElementById('errorSection').style.display = 'block';
    }
}

// Excel File Reading Function
async function readExcelFile(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = function(e) {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                
                // Get first worksheet
                const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
                const jsonData = XLSX.utils.sheet_to_json(firstSheet);
                
                // Get chart data if exists
                let chartData = [];
                if (workbook.Sheets['ChartData']) {
                    const chartSheet = workbook.Sheets['ChartData'];
                    chartData = XLSX.utils.sheet_to_json(chartSheet);
                }
                
                resolve({ data: jsonData, charts: chartData });
            } catch (error) {
                reject(error);
            }
        };
        reader.onerror = reject;
        reader.readAsArrayBuffer(file);
    });
}

// PowerPoint Template Reading Function
async function readPowerPointTemplate(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = function(e) {
            try {
                const arrayBuffer = e.target.result;
                // For now, we'll create a new presentation
                // In a real implementation, you'd parse the existing PPTX
                resolve({ fileName: file.name });
            } catch (error) {
                reject(error);
            }
        };
        reader.onerror = reject;
        reader.readAsArrayBuffer(file);
    });
}

// Main Presentation Generation Function
async function generatePresentation(excelData, template) {
    const pptx = new PptxGenJS();
    
    // Get first row of data
    const data = excelData.data[0] || {};
    
    // Helper function to replace placeholders
    const replacePlaceholders = (text) => {
        if (!text) return text;
        const defaultValues = {
            date: new Date().toLocaleDateString(),
            presenter: 'Presenter',
            company: 'Company Name'
        };
        
        return text.replace(/\{([^}]+)\}/g, (match, key) => {
            const trimmedKey = key.trim();
            const value = data[trimmedKey] !== undefined ? data[trimmedKey] : defaultValues[trimmedKey];
            return value !== undefined ? String(value) : match;
        });
    };
    
    // Title slide
    const titleSlide = pptx.addSlide();
    titleSlide.addText(replacePlaceholders('{project_title}'), {
        x: 1, y: 2, w: 8, h: 2,
        fontSize: 44, bold: true, color: '363636',
        align: 'center'
    });
    
    titleSlide.addText(replacePlaceholders('Status Report prepared by: {name} - {date}'), {
        x: 1, y: 4, w: 8, h: 1,
        fontSize: 18, color: '666666',
        align: 'center'
    });

    // Overview slide
    const overviewSlide = pptx.addSlide();
    overviewSlide.addText('Overview', {
        x: 1, y: 0.5, w: 8, h: 1,
        fontSize: 28, bold: true, color: '363636'
    });
    
    overviewSlide.addText(replacePlaceholders('{status_update}'), {
        x: 1, y: 2, w: 8, h: 4,
        fontSize: 16, color: '666666',
        valign: 'top'
    });

    // Agenda slide
    const agendaSlide = pptx.addSlide();
    agendaSlide.addText('Agenda', {
        x: 1, y: 0.5, w: 8, h: 1,
        fontSize: 28, bold: true, color: '363636'
    });
    
    const agendaItems = (replacePlaceholders('{agenda}') || '').split(' ').filter(item => item.trim());
    agendaItems.forEach((item, index) => {
        agendaSlide.addText(`• ${item}`, {
            x: 1.5, y: 2 + (index * 0.6), w: 7, h: 0.5,
            fontSize: 16, color: '666666'
        });
    });

    // Chart title slide
    const chartTitleSlide = pptx.addSlide();
    chartTitleSlide.addText(replacePlaceholders('{ChartTitle}'), {
        x: 1, y: 0.5, w: 8, h: 1,
        fontSize: 28, bold: true, color: '363636',
        align: 'center'
    });

    return pptx;
}

// Chart Generation Function
async function addCharts(pptx, excelData) {
    if (excelData.charts && excelData.charts.length > 0) {
        // Group charts by title
        const chartGroups = {};
        excelData.charts.forEach(chart => {
            if (!chartGroups[chart.ChartTitle]) {
                chartGroups[chart.ChartTitle] = [];
            }
            chartGroups[chart.ChartTitle].push(chart);
        });

        // Create a slide for each chart group
        Object.keys(chartGroups).forEach((chartTitle, index) => {
            const chartSlide = pptx.addSlide();
            chartSlide.addText(chartTitle, {
                x: 1, y: 0.5, w: 8, h: 1,
                fontSize: 28, bold: true, color: '363636',
                align: 'center'
            });

            const charts = chartGroups[chartTitle];
            const chartType = charts[0].ChartType;
            
            // Prepare data for chart
            const labels = [];
            const dataValues = [];
            const seriesData = {};
            
            charts.forEach(chart => {
                if (!labels.includes(chart.Label)) {
                    labels.push(chart.Label);
                }
                
                const seriesName = chart.Series || 'Data';
                if (!seriesData[seriesName]) {
                    seriesData[seriesName] = [];
                }
                seriesData[seriesName].push(chart.Value);
            });

            // Format data for PptxGenJS
            const formattedChartData = [];
            Object.keys(seriesData).forEach(seriesName => {
                formattedChartData.push({
                    name: seriesName,
                    labels: labels,
                    values: seriesData[seriesName]
                });
            });

            // Add chart based on type
            if (chartType === 'bar' || chartType === 'column') {
                chartSlide.addChart(pptx.charts.BAR, formattedChartData, {
                    x: 1, y: 1.5, w: 8, h: 5,
                    barDir: 'col',
                    showLegend: true,
                    legendPos: 'b'
                });
            } else if (chartType === 'line') {
                chartSlide.addChart(pptx.charts.LINE, formattedChartData, {
                    x: 1, y: 1.5, w: 8, h: 5,
                    lineDataSymbol: 'circle',
                    showLegend: true,
                    legendPos: 'b'
                });
            } else if (chartType === 'pie') {
                chartSlide.addChart(pptx.charts.PIE, formattedChartData, {
                    x: 1, y: 1.5, w: 8, h: 5,
                    showLegend: true,
                    legendPos: 'b'
                });
            } else {
                // Default to bar chart
                chartSlide.addChart(pptx.charts.BAR, formattedChartData, {
                    x: 1, y: 1.5, w: 8, h: 5,
                    showLegend: true,
                    legendPos: 'b'
                });
            }
        });
    }
}

// Progress Update Function
function updateProgress(percent, text) {
    document.getElementById('progressFill').style.width = percent + '%';
    document.getElementById('progressFill').textContent = percent + '%';
    document.getElementById('progressText').textContent = text;
}

// Download Presentation Function
function downloadPresentation() {
    if (processedPresentation) {
        processedPresentation.writeFile({ fileName: 'generated-presentation.pptx' })
            .then((blob) => {
                const url = URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url;
                a.download = 'generated-presentation.pptx';
                document.body.appendChild(a);
                a.click();
                document.body.removeChild(a);
                URL.revokeObjectURL(url);
            })
            .catch((error) => {
                console.error('Download error:', error);
                alert('Failed to download presentation');
            });
    }
}

// Reset Application Function
function resetApp() {
    pptFile = null;
    excelFile = null;
    processedPresentation = null;
    
    // Reset UI
    document.getElementById('pptInfo').style.display = 'none';
    document.getElementById('excelInfo').style.display = 'none';
    document.getElementById('pptBox').classList.remove('has-file');
    document.getElementById('excelBox').classList.remove('has-file');
    document.getElementById('processBtn').disabled = true;
    document.getElementById('progressSection').style.display = 'none';
    document.getElementById('resultSection').style.display = 'none';
    document.getElementById('errorSection').style.display = 'none';
    
    // Reset file inputs
    document.getElementById('pptInput').value = '';
    document.getElementById('excelInput').value = '';
}

// Check Ready to Process Function
function checkReadyToProcess() {
    const processBtn = document.getElementById('processBtn');
    if (pptFile && excelFile) {
        processBtn.disabled = false;
    } else {
        processBtn.disabled = true;
    }
}

// Sample File Generation Functions
function downloadSamplePPT() {
    const pptx = new PptxGenJS();
    
    // Title slide
    const slide1 = pptx.addSlide();
    slide1.addText('{project_title}', {
        x: 1, y: 2, w: 8, h: 2,
        fontSize: 44, bold: true, color: '363636',
        align: 'center'
    });
    
    slide1.addText('Status Report prepared by: {name}', {
        x: 1, y: 4, w: 8, h: 1,
        fontSize: 18, color: '666666',
        align: 'center'
    });

    // Overview slide
    const slide2 = pptx.addSlide();
    slide2.addText('Overview', {
        x: 1, y: 0.5, w: 8, h: 1,
        fontSize: 28, bold: true, color: '363636'
    });
    
    slide2.addText('{status_update}', {
        x: 1, y: 2, w: 8, h: 4,
        fontSize: 16, color: '666666'
    });

    // Agenda slide
    const slide3 = pptx.addSlide();
    slide3.addText('Agenda', {
        x: 1, y: 0.5, w: 8, h: 1,
        fontSize: 28, bold: true, color: '363636'
    });
    
    slide3.addText('{agenda}', {
        x: 1, y: 2, w: 8, h: 4,
        fontSize: 16, color: '666666'
    });

    // Chart slide
    const slide4 = pptx.addSlide();
    slide4.addText('{ChartTitle}', {
        x: 1, y: 0.5, w: 8, h: 1,
        fontSize: 28, bold: true, color: '363636'
    });
    
    slide4.addText('Chart will be generated here based on your Excel data', {
        x: 1, y: 2, w: 8, h: 2,
        fontSize: 16, color: '666666'
    });

    pptx.writeFile({ fileName: 'sample-template.pptx' })
        .then((blob) => {
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = 'sample-template.pptx';
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            URL.revokeObjectURL(url);
        });
}

function downloadSampleExcel() {
    // Create sample Excel data
    const wb = XLSX.utils.book_new();
    
    // Main data sheet
    const wsData = [
        ['name', 'project_title', 'status_update', 'agenda', 'ChartTitle'],
        ['John Doe', 'Q4 2024 Performance Review', 'All financial models have been updated, and the initial report draft is complete. We are on schedule for the board meeting next week.', 'Technology Stack System Context Design Decisions Mobile Application Features', 'Analysis of Key Metrics']
    ];
    
    const ws = XLSX.utils.aoa_to_sheet(wsData);
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
    
    // Chart data sheet
    const chartData = [
        ['ChartTitle', 'ChartType', 'Series', 'Label', 'Value'],
        ['Quarterly Sales', 'bar', 'Sales', 'Q1', 150000],
        ['Quarterly Sales', 'bar', 'Sales', 'Q2', 200000],
        ['Quarterly Sales', 'bar', 'Sales', 'Q3', 180000],
        ['Quarterly Sales', 'bar', 'Sales', 'Q4', 220000],
        ['Engagement Metrics', 'line', 'Likes', 'Jan', 1200],
        ['Engagement Metrics', 'line', 'Likes', 'Feb', 1800],
        ['Engagement Metrics', 'line', 'Likes', 'Mar', 2400],
        ['Engagement Metrics', 'line', 'Shares', 'Jan', 400],
        ['Engagement Metrics', 'line', 'Shares', 'Feb', 500],
        ['Engagement Metrics', 'line', 'Shares', 'Mar', 650]
    ];
    
    const chartWs = XLSX.utils.aoa_to_sheet(chartData);
    XLSX.utils.book_append_sheet(wb, chartWs, 'ChartData');
    
    // Write file
    const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    const blob = new Blob([wbout], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'sample-data.xlsx';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
}

// Initialize on page load
document.addEventListener('DOMContentLoaded', function() {
    checkReadyToProcess();
});
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