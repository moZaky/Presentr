const officegen = require('officegen');
const fs = require('fs');

// Create a valid PowerPoint presentation
async function createValidPPTX() {
  // Create a new PowerPoint presentation
  const pptx = officegen('pptx');

  // Slide 1 - Title Slide
  const slide1 = pptx.makeNewSlide();
  slide1.name = 'Title Slide';
  slide1.addText('{project_title}', {
    x: 1,
    y: 2,
    cx: 8,
    cy: 1.5,
    font_size: 44,
    font_face: 'Arial',
    bold: true,
    color: '363636',
    align: 'center'
  });

  slide1.addText('Status Report prepared by: {name}', {
    x: 1,
    y: 5,
    cx: 8,
    cy: 0.75,
    font_size: 18,
    font_face: 'Arial',
    color: '666666',
    align: 'center'
  });

  slide1.addText('{date}', {
    x: 1,
    y: 5.5,
    cx: 8,
    cy: 0.5,
    font_size: 16,
    font_face: 'Arial',
    color: '666666',
    align: 'center'
  });

  // Slide 2 - Overview
  const slide2 = pptx.makeNewSlide();
  slide2.name = 'Overview';
  slide2.addText('Overview', {
    x: 1,
    y: 0.5,
    cx: 8,
    cy: 1,
    font_size: 36,
    font_face: 'Arial',
    bold: true,
    color: '363636'
  });

  slide2.addText('{status_update}', {
    x: 1,
    y: 1.5,
    cx: 8,
    cy: 4,
    font_size: 20,
    font_face: 'Arial',
    color: '444444',
    valign: 'top'
  });

  // Slide 3 - Agenda
  const slide3 = pptx.makeNewSlide();
  slide3.name = 'Agenda';
  slide3.addText('Agenda', {
    x: 1,
    y: 0.5,
    cx: 8,
    cy: 1,
    font_size: 36,
    font_face: 'Arial',
    bold: true,
    color: '363636'
  });

  slide3.addText('{agenda}', {
    x: 1,
    y: 1.5,
    cx: 8,
    cy: 4,
    font_size: 20,
    font_face: 'Arial',
    color: '444444',
    valign: 'top'
  });

  // Slide 4 - Chart
  const slide4 = pptx.makeNewSlide();
  slide4.name = 'Chart';
  slide4.addText('{ChartTitle}', {
    x: 1,
    y: 0.5,
    cx: 8,
    cy: 1,
    font_size: 36,
    font_face: 'Arial',
    bold: true,
    color: '363636'
  });

  slide4.addText('This is a placeholder slide. The title above will be filled from your Excel data. Charts defined in the \'ChartData\' sheet will be added as new slides following this one.', {
    x: 1,
    y: 2,
    cx: 8,
    cy: 3,
    font_size: 18,
    font_face: 'Arial',
    color: '666666',
    valign: 'top'
  });

  // Generate the PowerPoint file
  const out = fs.createWriteStream('/home/z/my-project/samples/valid-sample-template.pptx');
  
  pptx.generate(out, {
    'finalize': function(written) {
      console.log('Finished to create a PowerPoint file: valid-sample-template.pptx');
      console.log('Total bytes created: ' + written + ' bytes');
    },
    'error': function(err) {
      console.log(err);
    }
  });
}

createValidPPTX();