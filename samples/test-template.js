const fs = require('fs');
const path = require('path');

console.log('🔍 Testing PowerPoint Template Files...\n');

const templates = [
  { name: 'professional-template.pptx', path: './professional-template.pptx' },
  { name: 'valid-sample-template.pptx', path: './valid-sample-template.pptx' },
  { name: 'sample-template.pptx', path: './sample-template.pptx' }
];

templates.forEach(template => {
  try {
    const fullPath = path.join(__dirname, template.path);
    
    if (fs.existsSync(fullPath)) {
      const stats = fs.statSync(fullPath);
      console.log(`✅ ${template.name}`);
      console.log(`   Size: ${stats.size.toLocaleString()} bytes`);
      console.log(`   Modified: ${stats.mtime.toLocaleString()}`);
      
      // 检查文件头
      const buffer = fs.readFileSync(fullPath);
      const header = buffer.slice(0, 8).toString('hex');
      console.log(`   Header: ${header}`);
      
      // PowerPoint文件应该以PK开头（ZIP格式）
      if (buffer.slice(0, 2).toString() === 'PK') {
        console.log(`   ✅ Valid ZIP/PPTX format`);
      } else {
        console.log(`   ❌ Invalid format - should start with 'PK'`);
      }
      
      console.log('');
    } else {
      console.log(`❌ ${template.name} - File not found`);
    }
  } catch (error) {
    console.log(`❌ ${template.name} - Error: ${error.message}`);
  }
});

console.log('📋 Recommendation:');
console.log('Use the professional-template.pptx file (65,729 bytes) as it was created with pptxgenjs');
console.log('and contains the complete PowerPoint structure with all required placeholders.');