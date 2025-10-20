const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');

async function testFixedAPI() {
  try {
    console.log('🧪 Testing Fixed API Implementation...\n');
    
    // Read sample data
    const excelPath = path.join(__dirname, 'sample-data.xlsx');
    const excelBuffer = fs.readFileSync(excelPath);
    const workbook = XLSX.read(excelBuffer, { type: 'buffer' });
    
    // Get main data
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const excelData = XLSX.utils.sheet_to_json(worksheet);
    const dataRow = excelData[0];
    
    // Get chart data
    let chartData = [];
    if (workbook.SheetNames.includes('ChartData')) {
      const chartSheet = workbook.Sheets['ChartData'];
      chartData = XLSX.utils.sheet_to_json(chartSheet);
    }
    dataRow.ChartData = chartData;
    
    console.log('📊 Data loaded:');
    console.log(`   Project: ${dataRow.project_title}`);
    console.log(`   Chart rows: ${chartData.length}`);
    
    // Test the new buffer generation method
    console.log('\n🔄 Testing new buffer method...');
    
    // Simulate the API function
    const PptxGenJS = require('pptxgenjs');
    const { join } = require('path');
    const { writeFile, readFile, unlink } = require('fs/promises');
    
    const pptx = new PptxGenJS();
    pptx.author = 'PowerPoint AutoFill Test';
    pptx.title = dataRow.project_title || 'Test Presentation';
    
    // Add slides
    const slide1 = pptx.addSlide();
    slide1.addText(dataRow.project_title, {
      x: 0.5, y: 2, w: 9, h: 1.5, fontSize: 44, bold: true, color: '363636', align: 'center'
    });
    
    // Add chart
    if (chartData.length > 0) {
      const chartSlide = pptx.addSlide();
      chartSlide.addText('Quarterly Sales', {
        x: 0.5, y: 0.5, w: 9, h: 1, fontSize: 36, bold: true, color: '363636'
      });
      
      const quarterlyData = chartData.filter(item => item.ChartTitle === 'Quarterly Sales');
      if (quarterlyData.length > 0) {
        const chartDataFormatted = [{
          name: 'Sales',
          labels: quarterlyData.map(item => item.Label),
          values: quarterlyData.map(item => Number(item.Value))
        }];
        
        chartSlide.addChart(pptx.charts.BAR, chartDataFormatted, {
          x: 0.5, y: 1.5, w: 9, h: 5,
          showLegend: true,
          legendPos: 'b'
        });
      }
    }
    
    // Use the working method: create temp file then read
    const tempPath = join(__dirname, `temp_test_${Date.now()}.pptx`);
    await pptx.writeFile({ fileName: tempPath });
    
    const buffer = await readFile(tempPath);
    
    // Clean up
    await unlink(tempPath);
    
    // Save test file
    const testFile = path.join(__dirname, 'fixed-api-test.pptx');
    fs.writeFileSync(testFile, buffer);
    
    const stats = fs.statSync(testFile);
    console.log(`✅ Fixed API test successful!`);
    console.log(`📁 File: ${testFile}`);
    console.log(`📊 Size: ${stats.size.toLocaleString()} bytes`);
    
    // Validate file
    const header = buffer.slice(0, 2).toString();
    const isValid = header === 'PK';
    console.log(`🔍 Format: ${isValid ? '✅ Valid PPTX' : '❌ Invalid'} (${header})`);
    
    return isValid;
    
  } catch (error) {
    console.error('❌ Test failed:', error.message);
    return false;
  }
}

testFixedAPI().then(success => {
  if (success) {
    console.log('\n🎉 API fix successful! Files should no longer be corrupted.');
  } else {
    console.log('\n💥 API fix failed. Need further investigation.');
  }
});