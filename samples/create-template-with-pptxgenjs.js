const PptxGenJS = require('pptxgenjs');
const fs = require('fs');

// Create a valid PowerPoint presentation using PptxGenJS
async function createValidTemplate() {
  try {
    // Create a new PowerPoint presentation
    const pptx = new PptxGenJS();
    
    // Define slide size (16:9 aspect ratio)
    pptx.defineLayout({ name: 'A4', width: 10, height: 7.5 });
    pptx.layout = 'A4';

    // Slide 1 - Title Slide
    const slide1 = pptx.addSlide();
    slide1.addText('{project_title}', {
      x: 1,
      y: 2,
      w: 8,
      h: 1.5,
      fontSize: 44,
      bold: true,
      color: '363636',
      align: 'center',
      fontFace: 'Arial'
    });
    
    slide1.addText('Status Report prepared by: {name}', {
      x: 1,
      y: 5,
      w: 8,
      h: 0.75,
      fontSize: 18,
      color: '666666',
      align: 'center',
      fontFace: 'Arial'
    });
    
    slide1.addText('{date}', {
      x: 1,
      y: 5.5,
      w: 8,
      h: 0.5,
      fontSize: 16,
      color: '666666',
      align: 'center',
      fontFace: 'Arial'
    });

    // Slide 2 - Overview
    const slide2 = pptx.addSlide();
    slide2.addText('Overview', {
      x: 1,
      y: 0.5,
      w: 8,
      h: 1,
      fontSize: 36,
      bold: true,
      color: '363636',
      fontFace: 'Arial'
    });
    
    slide2.addText('{status_update}', {
      x: 1,
      y: 1.5,
      w: 8,
      h: 4,
      fontSize: 20,
      color: '444444',
      valign: 'top',
      fontFace: 'Arial',
      wrapText: true
    });

    // Slide 3 - Agenda
    const slide3 = pptx.addSlide();
    slide3.addText('Agenda', {
      x: 1,
      y: 0.5,
      w: 8,
      h: 1,
      fontSize: 36,
      bold: true,
      color: '363636',
      fontFace: 'Arial'
    });
    
    slide3.addText('{agenda}', {
      x: 1,
      y: 1.5,
      w: 8,
      h: 4,
      fontSize: 20,
      color: '444444',
      valign: 'top',
      fontFace: 'Arial',
      wrapText: true
    });

    // Slide 4 - Chart
    const slide4 = pptx.addSlide();
    slide4.addText('{ChartTitle}', {
      x: 1,
      y: 0.5,
      w: 8,
      h: 1,
      fontSize: 36,
      bold: true,
      color: '363636',
      fontFace: 'Arial'
    });
    
    slide4.addText('This is a placeholder slide. The title above will be filled from your Excel data. Charts defined in the \'ChartData\' sheet will be added as new slides following this one.', {
      x: 1,
      y: 2,
      w: 8,
      h: 3,
      fontSize: 18,
      color: '666666',
      valign: 'top',
      fontFace: 'Arial',
      wrapText: true
    });

    // Generate the PowerPoint file
    const pptxBuffer = await pptx.writeFile({ 
      outputType: 'nodebuffer',
      fileName: 'valid-sample-template.pptx'
    });
    
    // Write the file
    fs.writeFileSync('/home/z/my-project/samples/valid-sample-template.pptx', pptxBuffer);
    
    console.log('Valid PowerPoint template created: valid-sample-template.pptx');
    console.log('File size:', pptxBuffer.length, 'bytes');
    
  } catch (error) {
    console.error('Error creating PowerPoint template:', error);
  }
}

createValidTemplate();