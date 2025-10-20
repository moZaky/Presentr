const PptxGenJS = require("pptxgenjs");
const fs = require("fs");
const path = require("path");

// 创建一个新的PowerPoint演示文稿
let pptx = new PptxGenJS();

// 设置演示文稿属性
pptx.author = "PowerPoint AutoFill";
pptx.company = "AutoFill System";
pptx.title = "Sample Template";
pptx.subject = "Template for AutoFill Demo";

// 定义幻灯片尺寸 (16:9)
pptx.defineLayout({ name: 'A4', width: 10, height: 7.5 });
pptx.layout = 'A4';

// 幻灯片1: 标题页
let slide1 = pptx.addSlide();
slide1.addText(
    [
        { text: '{project_title}', options: { fontSize: 44, bold: true, color: '363636', align: 'center' } },
        { text: '\nPresenter - {date}', options: { fontSize: 24, color: '666666', align: 'center' } },
        { text: '\nStatus Report prepared by: {name}', options: { fontSize: 18, color: '666666', align: 'center' } }
    ],
    { x: 0.5, y: 2, w: 9, h: 3.5, align: 'center', valign: 'middle' }
);

// 添加页脚
slide1.addText(
    'PowerPoint AutoFill System',
    { x: 0.5, y: 6.8, w: 9, h: 0.3, fontSize: 10, color: '999999', align: 'center' }
);

// 幻灯片2: 概述页
let slide2 = pptx.addSlide();
slide2.addText(
    'Overview',
    { x: 0.5, y: 0.5, w: 9, h: 1, fontSize: 36, bold: true, color: '363636' }
);

slide2.addText(
    '{status_update}',
    { x: 0.5, y: 1.5, w: 9, h: 4, fontSize: 18, color: '444444', valign: 'top', wrapText: true }
);

slide2.addText(
    'PowerPoint AutoFill System',
    { x: 0.5, y: 6.8, w: 9, h: 0.3, fontSize: 10, color: '999999', align: 'center' }
);

// 幻灯片3: 议程页
let slide3 = pptx.addSlide();
slide3.addText(
    'Agenda',
    { x: 0.5, y: 0.5, w: 9, h: 1, fontSize: 36, bold: true, color: '363636' }
);

slide3.addText(
    '{agenda}',
    { x: 0.5, y: 1.5, w: 9, h: 4, fontSize: 18, color: '444444', valign: 'top', wrapText: true }
);

slide3.addText(
    'PowerPoint AutoFill System',
    { x: 0.5, y: 6.8, w: 9, h: 0.3, fontSize: 10, color: '999999', align: 'center' }
);

// 幻灯片4: 图表页
let slide4 = pptx.addSlide();
slide4.addText(
    '{ChartTitle}',
    { x: 0.5, y: 0.5, w: 9, h: 1, fontSize: 36, bold: true, color: '363636' }
);

slide4.addText(
    'This is a placeholder slide. The title above will be filled from your Excel data.\n\nCharts defined in the \'ChartData\' sheet will be added as new slides following this one.\n\n📊 Chart data will be automatically inserted here based on your Excel file.',
    { x: 0.5, y: 1.5, w: 9, h: 4, fontSize: 16, color: '444444', valign: 'top', wrapText: true }
);

slide4.addText(
    'PowerPoint AutoFill System',
    { x: 0.5, y: 6.8, w: 9, h: 0.3, fontSize: 10, color: '999999', align: 'center' }
);

// 保存文件
const outputPath = path.join(__dirname, 'professional-template.pptx');
pptx.writeFile({ fileName: outputPath })
    .then(() => {
        console.log('✅ Professional PowerPoint template created successfully!');
        console.log(`📁 Saved to: ${outputPath}`);
        
        // 检查文件大小
        const stats = fs.statSync(outputPath);
        console.log(`📊 File size: ${stats.size} bytes`);
        
        // 验证文件结构
        console.log('🔍 Template contains the following placeholders:');
        console.log('   • {project_title} - Project title');
        console.log('   • {name} - Presenter name');
        console.log('   • {date} - Presentation date');
        console.log('   • {status_update} - Status update text');
        console.log('   • {agenda} - Agenda items');
        console.log('   • {ChartTitle} - Chart title');
    })
    .catch(err => {
        console.error('❌ Error creating PowerPoint template:', err);
    });