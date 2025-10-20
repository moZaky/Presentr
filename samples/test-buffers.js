const PptxGenJS = require('pptxgenjs');
const fs = require('fs');
const path = require('path');

async function testBufferMethods() {
  try {
    console.log('🔍 Testing different buffer methods...\n');
    
    const pptx = new PptxGenJS();
    const slide = pptx.addSlide();
    slide.addText('Buffer Test', { x: 0.5, y: 0.5, w: 9, h: 1, fontSize: 36 });
    
    // Method 1: outputType: 'nodebuffer'
    console.log('🔄 Method 1: outputType: "nodebuffer"');
    try {
      const buffer1 = await pptx.writeFile({ outputType: 'nodebuffer' });
      console.log(`   Buffer size: ${buffer1.length} bytes`);
      console.log(`   First 4 bytes: ${buffer1.slice(0, 4).toString()}`);
      
      fs.writeFileSync(path.join(__dirname, 'test-method1.pptx'), buffer1);
      console.log('   ✅ Method 1 saved');
    } catch (error) {
      console.log(`   ❌ Method 1 failed: ${error.message}`);
    }
    
    // Method 2: outputType: 'buffer'
    console.log('\n🔄 Method 2: outputType: "buffer"');
    try {
      const buffer2 = await pptx.writeFile({ outputType: 'buffer' });
      console.log(`   Buffer size: ${buffer2.length} bytes`);
      console.log(`   First 4 bytes: ${buffer2.slice(0, 4).toString()}`);
      
      fs.writeFileSync(path.join(__dirname, 'test-method2.pptx'), buffer2);
      console.log('   ✅ Method 2 saved');
    } catch (error) {
      console.log(`   ❌ Method 2 failed: ${error.message}`);
    }
    
    // Method 3: Create file then read as buffer
    console.log('\n🔄 Method 3: Create file then read');
    try {
      const tempFile = path.join(__dirname, 'temp.pptx');
      await pptx.writeFile({ fileName: tempFile });
      
      const buffer3 = fs.readFileSync(tempFile);
      console.log(`   Buffer size: ${buffer3.length} bytes`);
      console.log(`   First 4 bytes: ${buffer3.slice(0, 4).toString()}`);
      
      fs.writeFileSync(path.join(__dirname, 'test-method3.pptx'), buffer3);
      console.log('   ✅ Method 3 saved');
      
      // Clean up temp file
      fs.unlinkSync(tempFile);
    } catch (error) {
      console.log(`   ❌ Method 3 failed: ${error.message}`);
    }
    
    // Method 4: Using base64 then convert
    console.log('\n🔄 Method 4: Base64 conversion');
    try {
      const base64 = await pptx.writeFile({ outputType: 'base64' });
      const buffer4 = Buffer.from(base64, 'base64');
      console.log(`   Base64 length: ${base64.length}`);
      console.log(`   Buffer size: ${buffer4.length} bytes`);
      console.log(`   First 4 bytes: ${buffer4.slice(0, 4).toString()}`);
      
      fs.writeFileSync(path.join(__dirname, 'test-method4.pptx'), buffer4);
      console.log('   ✅ Method 4 saved');
    } catch (error) {
      console.log(`   ❌ Method 4 failed: ${error.message}`);
    }
    
  } catch (error) {
    console.error('❌ Test failed:', error.message);
  }
}

testBufferMethods();