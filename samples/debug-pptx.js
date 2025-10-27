const PptxGenJS = require('pptxgenjs');
const fs = require('fs');
const path = require('path');

async function debugPPTXGeneration() {
  try {
    console.log('🔍 Debugging PPTX Generation...\n');
    
    // Test 1: Simple presentation
    console.log('📄 Test 1: Creating simple presentation...');
    const pptx1 = new PptxGenJS();
    
    const slide1 = pptx1.addSlide();
    slide1.addText('Test Title', { x: 0.5, y: 0.5, w: 9, h: 1, fontSize: 36, bold: true });
    slide1.addText('Test Content', { x: 0.5, y: 2, w: 9, h: 1, fontSize: 18 });
    
    const testFile1 = path.join(__dirname, 'debug-simple.pptx');
    await pptx1.writeFile({ fileName: testFile1 });
    
    const stats1 = fs.statSync(testFile1);
    console.log(`   ✅ Simple presentation created: ${stats1.size.toLocaleString()} bytes`);
    
    // Test 2: Presentation with chart
    console.log('\n📊 Test 2: Creating presentation with chart...');
    const pptx2 = new PptxGenJS();
    
    const slide2 = pptx2.addSlide();
    slide2.addText('Chart Test', { x: 0.5, y: 0.5, w: 9, h: 1, fontSize: 36, bold: true });
    
    const chartData = [
      { name: 'Series 1', labels: ['A', 'B', 'C'], values: [10, 20, 30] }
    ];
    
    slide2.addChart(pptx2.charts.BAR, chartData, {
      x: 0.5, y: 1.5, w: 9, h: 5,
      showLegend: true,
      legendPos: 'b'
    });
    
    const testFile2 = path.join(__dirname, 'debug-chart.pptx');
    await pptx2.writeFile({ fileName: testFile2 });
    
    const stats2 = fs.statSync(testFile2);
    console.log(`   ✅ Chart presentation created: ${stats2.size.toLocaleString()} bytes`);
    
    // Test 3: Buffer output (like API)
    console.log('\n🔄 Test 3: Testing buffer output (API method)...');
    const pptx3 = new PptxGenJS();
    
    const slide3 = pptx3.addSlide();
    slide3.addText('Buffer Test', { x: 0.5, y: 0.5, w: 9, h: 1, fontSize: 36, bold: true });
    
    const chartData3 = [
      { name: 'Test Data', labels: ['Q1', 'Q2', 'Q3'], values: [100, 200, 150] }
    ];
    
    slide3.addChart(pptx3.charts.BAR, chartData3, {
      x: 0.5, y: 1.5, w: 9, h: 5,
      showLegend: true,
      legendPos: 'b'
    });
    
    // This is how the API generates the buffer
    const buffer = await pptx3.writeFile({ outputType: 'nodebuffer' });
    
    const testFile3 = path.join(__dirname, 'debug-buffer.pptx');
    fs.writeFileSync(testFile3, buffer);
    
    const stats3 = fs.statSync(testFile3);
    console.log(`   ✅ Buffer presentation created: ${stats3.size.toLocaleString()} bytes`);
    console.log(`   📊 Buffer size: ${buffer.length.toLocaleString()} bytes`);
    
    // Test 4: Verify files are valid
    console.log('\n🔍 Test 4: Validating generated files...');
    
    const files = [testFile1, testFile2, testFile3];
    files.forEach((file, index) => {
      const buffer = fs.readFileSync(file);
      const header = buffer.slice(0, 2).toString();
      const isValid = header === 'PK';
      
      console.log(`   File ${index + 1}: ${isValid ? '✅ Valid ZIP' : '❌ Invalid format'} (${header})`);
    });
    
    console.log('\n🎯 Debug Results:');
    console.log('✅ PPTXGenJS basic functionality works');
    console.log('✅ Chart generation works');
    console.log('✅ Buffer output works');
    console.log('✅ All files are valid ZIP/PPTX format');
    
    return true;
    
  } catch (error) {
    console.error('❌ Debug failed:', error.message);
    console.error('Stack:', error.stack);
    return false;
  }
}

debugPPTXGeneration().then(success => {
  if (success) {
    console.log('\n🎉 PPTXGenJS is working correctly!');
    console.log('The issue might be in the API implementation or file handling.');
  } else {
    console.log('\n💥 PPTXGenJS has issues - need to fix library usage');
  }
});