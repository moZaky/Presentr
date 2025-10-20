# PowerPoint AutoFill - Offline Version Guide

## 📋 Overview

The offline version of PowerPoint AutoFill works completely in your browser without requiring any server or internet connection. It uses client-side JavaScript libraries to process your files and generate PowerPoint presentations.

## 🚀 Quick Start

### Method 1: Direct File Access
1. Open `compiled-offline.html` in your web browser
2. The application will load immediately with all functionality available

### Method 2: Local Server (Recommended)
```bash
# Navigate to the project directory
cd /path/to/project

# Start a local server
python3 -m http.server 8080
# OR
npx serve .

# Open in browser
http://localhost:8080/compiled-offline.html
```

## ✨ Features

### ✅ Working Features
- **File Upload**: Drag & drop or click to upload PPTX and XLSX files
- **Data Processing**: Automatic extraction of data from Excel files
- **Presentation Generation**: Creates professional PowerPoint presentations
- **Chart Generation**: Supports bar charts, line charts, and pie charts
- **Sample Files**: Built-in sample file generation for testing
- **Progress Tracking**: Real-time progress updates during processing
- **Toast Notifications**: User-friendly feedback system
- **Download Functionality**: Direct download of generated presentations

### 📊 Supported Chart Types
- **Bar Charts** (vertical and horizontal)
- **Line Charts** (with data points)
- **Pie Charts** (with percentages)

## 📁 File Requirements

### PowerPoint Template (.pptx)
- Should contain placeholder text in the format `{placeholder_name}`
- Supported placeholders:
  - `{project_title}` - Main project title
  - `{name}` - Presenter name
  - `{date}` - Current date (auto-generated)
  - `{status_update}` - Status information
  - `{agenda}` - Agenda items
  - `{ChartTitle}` - Chart section title

### Excel Data File (.xlsx)
Must contain two sheets:

#### Sheet1: Main Data
| Column | Description | Example |
|--------|-------------|---------|
| name | Presenter name | "John Doe" |
| project_title | Project title | "Q4 2024 Review" |
| status_update | Status information | "All models updated..." |
| agenda | Agenda items | "Tech Stack, Design..." |
| ChartTitle | Chart section title | "Analysis of Metrics" |

#### ChartData: Chart Information
| Column | Description | Example |
|--------|-------------|---------|
| ChartTitle | Chart name | "Quarterly Sales" |
| ChartType | Chart type | "bar", "line", "pie" |
| Series | Data series name | "Q1", "Q2", "Likes" |
| Label | Data label | "Jan", "Feb", "" |
| Value | Numeric value | 150, 200, 1200 |

## 🎯 How to Use

### Step 1: Get Sample Files
1. Click "Sample PowerPoint" to download a template with placeholders
2. Click "Sample Excel" to download sample data with chart information

### Step 2: Upload Your Files
1. Upload your PowerPoint template (.pptx file)
2. Upload your Excel data file (.xlsx file)
3. The "Generate Presentation" button will become active

### Step 3: Generate Presentation
1. Click "Generate Presentation"
2. Watch the progress bar as files are processed
3. View the results when processing is complete

### Step 4: Download Results
1. Click "Download PowerPoint" to get your generated presentation
2. The file will include:
   - Title slide with your data
   - Overview slide with status updates
   - Agenda slide with formatted agenda
   - Chart slides with visual data

## 🔧 Technical Details

### Libraries Used
- **PptxGenJS**: PowerPoint generation library
- **XLSX.js**: Excel file processing
- **FileSaver.js**: File download functionality
- **Tailwind CSS**: Styling framework

### Browser Compatibility
- ✅ Chrome 60+
- ✅ Firefox 55+
- ✅ Safari 12+
- ✅ Edge 79+

### File Size Limits
- PowerPoint files: Up to 50MB
- Excel files: Up to 10MB
- Generated presentations: Typically 1-5MB

## 🐛 Troubleshooting

### Common Issues

#### "Files not uploading"
- Ensure files are in correct format (.pptx and .xlsx)
- Check file size limits
- Try refreshing the page

#### "Processing failed"
- Verify Excel file has both "Sheet1" and "ChartData" sheets
- Check that ChartData sheet has required columns
- Ensure numeric values in Value column

#### "Download not working"
- Check browser download settings
- Try a different browser
- Disable ad blockers temporarily

#### "Charts not displaying"
- Verify ChartType column has valid values ("bar", "line", "pie")
- Check that Value column contains numbers only
- Ensure Label column is not empty for categorical data

### Error Messages
- **"Upload Error"**: Invalid file type or corrupted file
- **"Processing Error"**: Invalid data format or missing sheets
- **"Download Error"**: Browser permission issue or file generation failure

## 🎨 Customization

### Adding New Placeholders
1. Edit the PowerPoint template to add `{new_placeholder}`
2. Add corresponding column in Excel Sheet1
3. The system will automatically replace it

### Modifying Chart Styles
Chart styles can be customized in the JavaScript code:
```javascript
// Example: Change bar chart colors
chartSlide.addChart(pptx.charts.BAR, formattedChartData, {
    x: 0.5, y: 1.5, w: 9, h: 5,
    barDir: 'col',
    showLegend: true,
    legendPos: 'b',
    // Add custom colors
    colors: ['FF6384', '36A2EB', 'FFCE56']
});
```

## 📱 Mobile Support

The offline version is responsive and works on:
- ✅ Mobile phones (iOS Safari, Android Chrome)
- ✅ Tablets (iPad Safari, Android Chrome)
- ✅ Desktop browsers

## 🔒 Security & Privacy

- **100% Offline**: No data sent to external servers
- **Local Processing**: All processing happens in your browser
- **No Data Storage**: Files are processed and discarded
- **Privacy First**: Your data never leaves your device

## 🆚 Online vs Offline

| Feature | Online Version | Offline Version |
|---------|----------------|-----------------|
| Server Required | ✅ Yes | ❌ No |
| Internet Needed | ✅ Yes | ❌ No |
| File Processing | Server-side | Client-side |
| Data Privacy | Server logs | 100% Private |
| Performance | Faster for large files | Limited by browser |
| Deployment | Complex | Simple HTML file |

## 📞 Support

If you encounter issues:
1. Check this troubleshooting guide
2. Try the sample files first
3. Test with different browsers
4. Check browser console for errors

## 🔄 Updates

To update the offline version:
1. Replace `compiled-offline.html` with the new version
2. Clear browser cache
3. Test with sample files

---

**Note**: This offline version provides the same core functionality as the online version while maintaining complete privacy and working without any server requirements.