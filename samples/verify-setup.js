const fs = require('fs');
const path = require('path');

console.log('🔍 PowerPoint AutoFill - Setup Verification\n');

// 检查关键文件
const requiredFiles = [
  'professional-template.pptx',
  'sample-data.xlsx',
  'README.md'
];

console.log('📁 Checking sample files...');
let allFilesExist = true;

requiredFiles.forEach(file => {
  const filePath = path.join(__dirname, file);
  if (fs.existsSync(filePath)) {
    const stats = fs.statSync(filePath);
    console.log(`✅ ${file} (${stats.size.toLocaleString()} bytes)`);
  } else {
    console.log(`❌ ${file} - Missing!`);
    allFilesExist = false;
  }
});

// 检查PowerPoint模板的有效性
console.log('\n📄 Verifying PowerPoint template...');
const pptxPath = path.join(__dirname, 'professional-template.pptx');
if (fs.existsSync(pptxPath)) {
  const buffer = fs.readFileSync(pptxPath);
  if (buffer.slice(0, 2).toString() === 'PK') {
    console.log('✅ Valid ZIP/PPTX format');
  } else {
    console.log('❌ Invalid format');
  }
}

// 检查Excel文件
console.log('\n📊 Verifying Excel data file...');
const excelPath = path.join(__dirname, 'sample-data.xlsx');
if (fs.existsSync(excelPath)) {
  const stats = fs.statSync(excelPath);
  console.log(`✅ Excel file exists (${stats.size.toLocaleString()} bytes)`);
}

console.log('\n🎯 Setup Status:');
if (allFilesExist) {
  console.log('✅ All sample files are ready for testing!');
  console.log('\n📋 Next Steps:');
  console.log('1. Visit http://localhost:3000');
  console.log('2. Click "View Sample Files"');
  console.log('3. Download professional-template.pptx and sample-data.xlsx');
  console.log('4. Upload both files to test the AutoFill functionality');
  console.log('5. Download the processed PowerPoint result');
} else {
  console.log('❌ Some files are missing. Please check the setup.');
}

console.log('\n📖 For detailed instructions, see: README.md');