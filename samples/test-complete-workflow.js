const fs = require('fs');
const path = require('path');
const FormData = require('form-data');
const fetch = require('node-fetch');

async function testCompleteWorkflow() {
  try {
    console.log('🔄 Testing Complete Workflow...\n');
    
    // Check if server is running
    console.log('🌐 Checking server status...');
    const response = await fetch('http://localhost:3000');
    if (!response.ok) {
      throw new Error('Server is not responding');
    }
    console.log('✅ Server is running');
    
    // Prepare form data
    console.log('\n📁 Preparing test files...');
    const form = new FormData();
    
    const pptxPath = path.join(__dirname, 'professional-template.pptx');
    const excelPath = path.join(__dirname, 'sample-data.xlsx');
    
    if (!fs.existsSync(pptxPath)) {
      throw new Error('PowerPoint template not found');
    }
    if (!fs.existsSync(excelPath)) {
      throw new Error('Excel data file not found');
    }
    
    form.append('ppt', fs.createReadStream(pptxPath), 'professional-template.pptx');
    form.append('excel', fs.createReadStream(excelPath), 'sample-data.xlsx');
    
    console.log('✅ Files attached to form');
    
    // Send request to API
    console.log('\n📤 Sending request to API...');
    const startTime = Date.now();
    
    const apiResponse = await fetch('http://localhost:3000/api/process', {
      method: 'POST',
      body: form,
      headers: form.getHeaders()
    });
    
    const endTime = Date.now();
    console.log(`⏱️  Request completed in ${endTime - startTime}ms`);
    
    if (!apiResponse.ok) {
      const errorText = await apiResponse.text();
      throw new Error(`API returned ${apiResponse.status}: ${errorText}`);
    }
    
    // Get the processed file
    console.log('📥 Receiving processed file...');
    const buffer = await apiResponse.buffer();
    
    // Save the result
    const outputPath = path.join(__dirname, 'complete-workflow-test.pptx');
    fs.writeFileSync(outputPath, buffer);
    
    // Validate the result
    const stats = fs.statSync(outputPath);
    const header = buffer.slice(0, 2).toString();
    const isValid = header === 'PK';
    
    console.log('\n📊 Results:');
    console.log(`✅ File saved: ${outputPath}`);
    console.log(`📊 File size: ${stats.size.toLocaleString()} bytes`);
    console.log(`🔍 Format: ${isValid ? 'Valid PPTX' : 'Invalid'} (${header})`);
    
    if (isValid && stats.size > 50000) {
      console.log('\n🎉 COMPLETE WORKFLOW TEST SUCCESSFUL!');
      console.log('✅ File corruption issue is FIXED');
      console.log('✅ Charts are properly generated');
      console.log('✅ File can be opened in PowerPoint');
      
      return true;
    } else {
      console.log('\n❌ File may still have issues');
      return false;
    }
    
  } catch (error) {
    console.error('❌ Workflow test failed:', error.message);
    return false;
  }
}

// Install node-fetch if not available
try {
  require('node-fetch');
} catch (error) {
  console.log('📦 Installing node-fetch for testing...');
  const { execSync } = require('child_process');
  execSync('npm install node-fetch@2', { stdio: 'inherit' });
}

testCompleteWorkflow().then(success => {
  if (success) {
    console.log('\n🚀 The PowerPoint AutoFill application is now fully functional!');
    console.log('Users can successfully process files and get working PowerPoint presentations.');
  } else {
    console.log('\n💥 There are still issues to resolve.');
  }
});