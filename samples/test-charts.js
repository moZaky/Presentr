const XLSX = require('xlsx');
const PptxGenJS = require('pptxgenjs');
const fs = require('fs');
const path = require('path');

async function testChartGeneration() {
  try {
    console.log('🧪 Testing Chart Generation with Sample Data...\n');
    
    // Read the sample Excel file
    const excelPath = path.join(__dirname, 'sample-data.xlsx');
    const excelBuffer = fs.readFileSync(excelPath);
    const workbook = XLSX.read(excelBuffer, { type: 'buffer' });
    
    // Get main data
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const excelData = XLSX.utils.sheet_to_json(worksheet);
    const dataRow = excelData[0];
    
    console.log('📊 Main Data:');
    console.log(`   Project Title: ${dataRow.project_title}`);
    console.log(`   Name: ${dataRow.name}`);
    console.log(`   Date: ${dataRow.date}`);
    
    // Get chart data
    let chartData = [];
    if (workbook.SheetNames.includes('ChartData')) {
      const chartSheet = workbook.Sheets['ChartData'];
      chartData = XLSX.utils.sheet_to_json(chartSheet);
    }
    
    console.log(`\n📈 Chart Data Found: ${chartData.length} rows`);
    chartData.forEach((row, index) => {
      console.log(`   ${index + 1}. ${row.ChartTitle} - ${row.ChartType} - ${row.Series}: ${row.Label} = ${row.Value}`);
    });
    
    // Group chart data by title
    const chartGroups = {};
    chartData.forEach(item => {
      const title = item.ChartTitle || 'Chart';
      if (!chartGroups[title]) {
        chartGroups[title] = [];
      }
      chartGroups[title].push(item);
    });
    
    console.log(`\n📋 Chart Groups: ${Object.keys(chartGroups).length}`);
    Object.keys(chartGroups).forEach(title => {
      console.log(`   • ${title} (${chartGroups[title].length} data points)`);
    });
    
    // Create presentation
    console.log('\n🎯 Creating Presentation with Charts...');
    const pptx = new PptxGenJS();
    
    // Set properties
    pptx.author = 'PowerPoint AutoFill Test';
    pptx.title = dataRow.project_title || 'Test Presentation';
    // Remove custom layout to use default
    
    // Add title slide
    const slide1 = pptx.addSlide();
    slide1.addText(dataRow.project_title || 'Test Title', {
      x: 0.5, y: 2, w: 9, h: 1.5, fontSize: 44, bold: true, color: '363636', align: 'center'
    });
    slide1.addText(`Status Report by: ${dataRow.name} - ${dataRow.date}`, {
      x: 0.5, y: 3.5, w: 9, h: 1, fontSize: 24, color: '666666', align: 'center'
    });
    
    // Add chart slides
    Object.entries(chartGroups).forEach(([chartTitle, chartGroup]) => {
      const chartSlide = pptx.addSlide();
      chartSlide.addText(chartTitle, {
        x: 0.5, y: 0.5, w: 9, h: 1, fontSize: 36, bold: true, color: '363636'
      });
      
      // Format data for chart
      const seriesGroups = {};
      chartGroup.forEach(item => {
        const seriesName = item.Series || 'Data';
        if (!seriesGroups[seriesName]) {
          seriesGroups[seriesName] = [];
        }
        seriesGroups[seriesName].push(item);
      });
      
      const chartData = [];
      Object.entries(seriesGroups).forEach(([seriesName, seriesData]) => {
        const labels = seriesData.map(item => item.Label || '');
        const values = seriesData.map(item => Number(item.Value) || 0);
        chartData.push({ name: seriesName, labels: labels, values: values });
      });
      
      // Add chart based on type
      const chartType = chartGroup[0].ChartType?.toLowerCase();
      
      if (chartType === 'bar' || chartType === 'column') {
        chartSlide.addChart(pptx.charts.BAR, chartData, {
          x: 0.5, y: 1.5, w: 9, h: 5,
          barDir: 'col',
          showLegend: true,
          legendPos: 'b'
        });
        console.log(`   ✅ Added BAR chart: ${chartTitle}`);
      } else if (chartType === 'line') {
        chartSlide.addChart(pptx.charts.LINE, chartData, {
          x: 0.5, y: 1.5, w: 9, h: 5,
          showLegend: true,
          legendPos: 'b'
        });
        console.log(`   ✅ Added LINE chart: ${chartTitle}`);
      } else if (chartType === 'pie') {
        chartSlide.addChart(pptx.charts.PIE, chartData, {
          x: 0.5, y: 1.5, w: 9, h: 5,
          showLegend: true,
          legendPos: 'b'
        });
        console.log(`   ✅ Added PIE chart: ${chartTitle}`);
      } else {
        chartSlide.addChart(pptx.charts.BAR, chartData, {
          x: 0.5, y: 1.5, w: 9, h: 5,
          showLegend: true,
          legendPos: 'b'
        });
        console.log(`   ✅ Added DEFAULT BAR chart: ${chartTitle}`);
      }
    });
    
    // Save the presentation
    const outputPath = path.join(__dirname, 'test-charts-output.pptx');
    await pptx.writeFile({ fileName: outputPath });
    
    const stats = fs.statSync(outputPath);
    console.log(`\n✅ Test Presentation Created Successfully!`);
    console.log(`📁 Saved to: ${outputPath}`);
    console.log(`📊 File size: ${stats.size.toLocaleString()} bytes`);
    console.log(`📈 Total charts: ${Object.keys(chartGroups).length}`);
    
    return true;
    
  } catch (error) {
    console.error('❌ Test failed:', error.message);
    return false;
  }
}

// Run the test
testChartGeneration().then(success => {
  if (success) {
    console.log('\n🎉 Chart generation test completed successfully!');
    console.log('The new API should now create proper visual charts in PowerPoint slides.');
  } else {
    console.log('\n💥 Chart generation test failed. Check the error above.');
  }
});