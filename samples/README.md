# Sample Files for PowerPoint AutoFill

This directory contains sample files to test the PowerPoint AutoFill application.

## Files

### 📄 professional-template.pptx (Recommended)
A **professional PowerPoint template** created with pptxgenjs library, featuring:
- **4 complete slides** with professional styling
- **6 placeholders** for dynamic content replacement
- **65,729 bytes** - Full PowerPoint structure with all XML components
- **Microsoft PowerPoint compatible** - Tested and verified to open correctly

**Placeholders:**
- `{project_title}` - Main title on the first slide
- `{name}` - Presenter name  
- `{date}` - Date of presentation
- `{status_update}` - Status text on the Overview slide
- `{agenda}` - Agenda items on the Agenda slide
- `{ChartTitle}` - Title for the chart slide

**This is the recommended template file as it has the most complete and reliable structure.**

### 📄 valid-sample-template.pptx
A **valid PowerPoint template** with the same placeholders, created using manual ZIP assembly:
- **31,811 bytes** - Smaller but functional
- **Microsoft PowerPoint compatible** - Tested and verified to open correctly

### 📄 sample-template.pptx
Copy of the professional template for backward compatibility.

### 📊 sample-data.xlsx
Excel file with two sheets:

#### Sheet1 (Main Data)
| Column | Value |
|--------|-------|
| name | Sarah Johnson |
| project_title | Q4 2024 Performance Review |
| status_update | All financial models have been updated, and the initial report draft is complete. We are on schedule for the board meeting next week. Key metrics show positive growth across all departments. |
| agenda | Technology Stack System Context Design Decisions Mobile Application Features |
| ChartTitle | Analysis of Key Metrics |
| date | 12/18/2024 |
| department | Finance |
| email | sarah.johnson@company.com |

#### ChartData (Chart Information)
| ChartTitle | ChartType | Series | Label | Value |
|------------|-----------|--------|-------|-------|
| Quarterly Sales | bar | Q1 | Q1 2024 | 150000 |
| Quarterly Sales | bar | Q2 | Q2 2024 | 200000 |
| Quarterly Sales | bar | Q3 | Q3 2024 | 180000 |
| Quarterly Sales | bar | Q4 | Q4 2024 | 220000 |
| Engagement Metrics | line | Likes | January | 1200 |
| Engagement Metrics | line | Likes | February | 1800 |
| Engagement Metrics | line | Likes | March | 2400 |
| Engagement Metrics | line | Shares | January | 400 |
| Engagement Metrics | line | Shares | February | 500 |
| Engagement Metrics | line | Shares | March | 650 |
| User Demographics | pie | Age 18-24 | 18-24 | 25 |
| User Demographics | pie | Age 25-34 | 25-34 | 35 |
| User Demographics | pie | Age 35-44 | 35-44 | 20 |
| User Demographics | pie | Age 45+ | 45+ | 20 |

## How to Use

1. **Open the application** at `http://localhost:3000`
2. **Download the PowerPoint template**: Click "Choose File" under "PowerPoint Template" and select `professional-template.pptx` (recommended)
3. **Upload the Excel data**: Click "Choose File" under "Excel Data" and select `sample-data.xlsx`
4. **Process files**: Click the "Process Files" button
5. **Download result**: After processing, click "Download" to get the filled PowerPoint

## Expected Result

The processed PowerPoint will contain:
- **Slide 1**: "Q4 2024 Performance Review" with "Status Report prepared by: Sarah Johnson - 12/18/2024"
- **Slide 2**: "Overview" with the full status update text
- **Slide 3**: "Agenda" with the agenda items
- **Slide 4**: "Analysis of Key Metrics" with placeholder text
- **Slide 5**: "Quarterly Sales" - **BAR CHART** showing Q1-Q4 sales data (150K, 200K, 180K, 220K)
- **Slide 6**: "Engagement Metrics" - **LINE CHART** showing Likes and Shares trends over 3 months
- **Slide 7**: "User Demographics" - **PIE CHART** showing age distribution (18-24: 25%, 25-34: 35%, 35-44: 20%, 45+: 20%)

**🎯 Chart Features:**
- **Visual Charts**: Actual PowerPoint charts (not just text placeholders)
- **Multiple Types**: Bar charts, line charts, and pie charts supported
- **Automatic Sizing**: Charts are properly sized and positioned
- **Legend Support**: Charts include legends for better readability
- **Data Labels**: Clear labels and values displayed on charts

## Placeholders Supported

The application automatically replaces any text wrapped in curly braces `{}` with the corresponding value from the Excel file. Column names in Excel should match the placeholder names exactly (without the curly braces).

Common placeholders:
- `{name}` - Person's name
- `{project_title}` - Project or presentation title
- `{date}` - Date (auto-filled with current date if not provided)
- `{status_update}` - Status information
- `{agenda}` - Agenda items
- `{ChartTitle}` - Chart title
- `{department}` - Department name
- `{email}` - Email address

## Tips

- Make sure your Excel column names exactly match the placeholder names in your PowerPoint
- The application reads the first row of data from Sheet1
- Chart data should be in a sheet named "ChartData"
- Multiple chart types are supported: bar, line, pie