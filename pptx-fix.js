// Fixed PowerPoint generation functions using PptxGenJS

// Create sample PowerPoint using PptxGenJS
function createSamplePowerPoint() {
    try {
        // Create a new presentation
        const pptx = new PptxGenJS();
        
        // Set presentation properties
        pptx.author = 'Presentr - PowerPoint AutoFill';
        pptx.company = 'AutoFill System';
        pptx.title = 'Sample PowerPoint Template';
        
        // Sample data for placeholders
        const sampleData = {
            project_title: '{project_title}',
            name: '{name}',
            date: '{date}',
            status_update: '{status_update}',
            agenda: '{agenda}',
            ChartTitle: '{ChartTitle}'
        };
        
        // Helper function to show placeholders
        const showPlaceholders = (text) => text;
        
        // Slide 1: Title Slide
        const slide1 = pptx.addSlide();
        slide1.addText(
            showPlaceholders('{project_title}'),
            { x: 0.5, y: 2, w: 9, h: 1.5, fontSize: 44, bold: true, color: '363636', align: 'center' }
        );
        
        slide1.addText(
            'Status Report prepared by: {name}',
            { x: 0.5, y: 3.5, w: 9, h: 1, fontSize: 24, color: '666666', align: 'center' }
        );
        
        slide1.addText(
            '{date}',
            { x: 0.5, y: 4.2, w: 9, h: 0.8, fontSize: 18, color: '666666', align: 'center' }
        );
        
        // Slide 2: Overview
        const slide2 = pptx.addSlide();
        slide2.addText(
            'Overview',
            { x: 0.5, y: 0.5, w: 9, h: 1, fontSize: 36, bold: true, color: '363636' }
        );
        
        slide2.addText(
            showPlaceholders('{status_update}'),
            { x: 0.5, y: 1.5, w: 9, h: 4, fontSize: 18, color: '444444', valign: 'top', wrapText: true }
        );
        
        // Slide 3: Agenda
        const slide3 = pptx.addSlide();
        slide3.addText(
            'Agenda',
            { x: 0.5, y: 0.5, w: 9, h: 1, fontSize: 36, bold: true, color: '363636' }
        );
        
        slide3.addText(
            showPlaceholders('{agenda}'),
            { x: 0.5, y: 1.5, w: 9, h: 4, fontSize: 18, color: '444444', valign: 'top', wrapText: true }
        );
        
        // Slide 4: Chart Title Placeholder
        const slide4 = pptx.addSlide();
        slide4.addText(
            showPlaceholders('{ChartTitle}'),
            { x: 0.5, y: 0.5, w: 9, h: 1, fontSize: 36, bold: true, color: '363636' }
        );
        
        slide4.addText(
            'This is a placeholder slide. The title above will be filled from your Excel data. Charts defined in the \'ChartData\' sheet will be added as new slides following this one.',
            { x: 0.5, y: 1.5, w: 9, h: 3, fontSize: 16, color: '666666', valign: 'top', wrapText: true }
        );
        
        // Generate and download the presentation
        pptx.writeFile({ fileName: 'sample_template.pptx' }).then((fileName) => {
            showToast('Sample Downloaded', 'Sample PowerPoint template downloaded successfully', 'success');
        }).catch((error) => {
            console.error('Error generating PowerPoint:', error);
            showToast('Generation Error', 'Failed to generate PowerPoint template', 'error');
        });
        
    } catch (error) {
        console.error('Error in createSamplePowerPoint:', error);
        showToast('Error', 'Failed to create sample PowerPoint: ' + error.message, 'error');
    }
}

// Create presentation with data using PptxGenJS (replace the existing function)
async function createSimplePresentation(data) {
    try {
        // Create a new presentation
        const pptx = new PptxGenJS();
        
        // Set presentation properties
        pptx.author = 'Presentr - PowerPoint AutoFill';
        pptx.company = 'AutoFill System';
        pptx.title = data.mainData?.project_title || 'AutoFill Presentation';
        
        // Get the main data
        const mainData = data.mainData || {};
        const chartData = data.chartData || [];
        
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
                const value = mainData[trimmedKey] !== undefined ? mainData[trimmedKey] : defaultValues[trimmedKey];
                return value !== undefined ? String(value) : match;
            });
        };
        
        // Slide 1: Title Slide
        const slide1 = pptx.addSlide();
        slide1.addText(
            replacePlaceholders('{project_title}'),
            { x: 0.5, y: 2, w: 9, h: 1.5, fontSize: 44, bold: true, color: '363636', align: 'center' }
        );
        
        slide1.addText(
            replacePlaceholders('Status Report prepared by: {name} - {date}'),
            { x: 0.5, y: 3.5, w: 9, h: 1, fontSize: 24, color: '666666', align: 'center' }
        );
        
        // Slide 2: Overview
        const slide2 = pptx.addSlide();
        slide2.addText(
            'Overview',
            { x: 0.5, y: 0.5, w: 9, h: 1, fontSize: 36, bold: true, color: '363636' }
        );
        
        slide2.addText(
            replacePlaceholders('{status_update}'),
            { x: 0.5, y: 1.5, w: 9, h: 4, fontSize: 18, color: '444444', valign: 'top', wrapText: true }
        );
        
        // Slide 3: Agenda
        const slide3 = pptx.addSlide();
        slide3.addText(
            'Agenda',
            { x: 0.5, y: 0.5, w: 9, h: 1, fontSize: 36, bold: true, color: '363636' }
        );
        
        slide3.addText(
            replacePlaceholders('{agenda}'),
            { x: 0.5, y: 1.5, w: 9, h: 4, fontSize: 18, color: '444444', valign: 'top', wrapText: true }
        );
        
        // Slide 4: Chart Title Placeholder
        const slide4 = pptx.addSlide();
        slide4.addText(
            replacePlaceholders('{ChartTitle}'),
            { x: 0.5, y: 0.5, w: 9, h: 1, fontSize: 36, bold: true, color: '363636' }
        );
        
        slide4.addText(
            'Chart data will be displayed in the following slides based on your Excel data.',
            { x: 0.5, y: 1.5, w: 9, h: 1, fontSize: 16, color: '666666' }
        );
        
        // Process chart data and create chart slides
        if (chartData && Array.isArray(chartData) && chartData.length > 0) {
            const chartGroups = groupChartData(chartData);
            
            Object.entries(chartGroups).forEach(([chartTitle, chartGroup]) => {
                const chartSlide = pptx.addSlide();
                chartSlide.addText(
                    chartTitle,
                    { x: 0.5, y: 0.5, w: 9, h: 1, fontSize: 36, bold: true, color: '363636' }
                );
                
                // Create chart based on chart type
                const chartType = chartGroup[0].ChartType?.toLowerCase();
                const formattedChartData = formatChartData(chartGroup);
                
                if (chartType === 'bar' || chartType === 'column') {
                    chartSlide.addChart(pptx.charts.BAR, formattedChartData, {
                        x: 0.5, y: 1.5, w: 9, h: 5,
                        barDir: 'col',
                        catAxisLabelRotate: 45,
                        showLegend: true,
                        legendPos: 'b'
                    });
                } else if (chartType === 'line') {
                    chartSlide.addChart(pptx.charts.LINE, formattedChartData, {
                        x: 0.5, y: 1.5, w: 9, h: 5,
                        lineDataSymbol: 'circle',
                        showLegend: true,
                        legendPos: 'b'
                    });
                } else if (chartType === 'pie') {
                    chartSlide.addChart(pptx.charts.PIE, formattedChartData, {
                        x: 0.5, y: 1.5, w: 9, h: 5,
                        showLegend: true,
                        legendPos: 'b',
                        dataLabelFormatCode: '#,##0.0%'
                    });
                } else {
                    // Default to bar chart if type is not recognized
                    chartSlide.addChart(pptx.charts.BAR, formattedChartData, {
                        x: 0.5, y: 1.5, w: 9, h: 5,
                        barDir: 'col',
                        showLegend: true,
                        legendPos: 'b'
                    });
                }
            });
        }
        
        return pptx;
        
    } catch (error) {
        console.error('Error creating presentation:', error);
        throw new Error('Failed to create presentation: ' + error.message);
    }
}

// Group chart data by ChartTitle
function groupChartData(chartData) {
    const grouped = {};
    
    chartData.forEach(item => {
        const title = item.ChartTitle || 'Chart';
        if (!grouped[title]) {
            grouped[title] = [];
        }
        grouped[title].push(item);
    });
    
    return grouped;
}

// Format chart data for PptxGenJS
function formatChartData(chartGroup) {
    const chartData = [];
    
    // Group by series if multiple series exist
    const seriesGroups = {};
    
    chartGroup.forEach(item => {
        const seriesName = item.Series || 'Data';
        if (!seriesGroups[seriesName]) {
            seriesGroups[seriesName] = [];
        }
        seriesGroups[seriesName].push(item);
    });
    
    // Format for PptxGenJS
    Object.entries(seriesGroups).forEach(([seriesName, seriesData]) => {
        const labels = seriesData.map(item => item.Label || '');
        const values = seriesData.map(item => Number(item.Value) || 0);
        
        chartData.push({
            name: seriesName,
            labels: labels,
            values: values
        });
    });
    
    return chartData;
}