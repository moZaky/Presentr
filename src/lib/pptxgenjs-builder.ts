/**
 * PPTXGenJS Builder - Creates fully compatible PPTX files using PPTXGenJS
 */

import PptxGenJS from 'pptxgenjs';

export interface PPTXData {
  [key: string]: any;
}

export class PPTXGenJSBuilder {
  private pptx: PptxGenJS;
  private data: PPTXData;

  constructor() {
    this.pptx = new PptxGenJS();
    this.data = {};
    this.setupDefaults();
  }

  /**
   * Setup default PPTX properties
   */
  private setupDefaults(): void {
    this.pptx.defineLayout({ name: 'A4', width: 10, height: 7.5 });
    this.pptx.layout = 'A4';
    
    // Set default font
    this.pptx.defineSlideMaster({
      title: 'DEFAULT_SLIDE',
      background: { fill: 'FFFFFF' },
      objects: [
        {
          text: {
            text: '',
            options: {
              x: 1,
              y: 0.5,
              w: 8,
              h: 1,
              fontSize: 36,
              fontFace: 'Calibri',
              bold: true,
              color: '363636'
            }
          }
        }
      ],
      slideNumber: { x: 9.3, y: 7.0 }
    });
  }

  /**
   * Set data for replacement
   */
  setData(data: PPTXData): void {
    this.data = data;
  }

  /**
   * Create slides from template data
   */
  async createFromTemplate(): Promise<void> {
    try {
      // Slide 1: Title slide
      await this.createTitleSlide();
      
      // Slide 2: Overview
      await this.createOverviewSlide();
      
      // Slide 3: Agenda
      await this.createAgendaSlide();
      
      // Slide 4: Chart placeholder
      await this.createChartSlide();
      
    } catch (error) {
      console.error('Error creating slides:', error);
      throw error;
    }
  }

  /**
   * Create title slide
   */
  private async createTitleSlide(): Promise<void> {
    const slide = this.pptx.addSlide();
    
    // Title
    slide.addText(this.data.project_title || 'Project Title', {
      x: 1,
      y: 1.5,
      w: 8,
      h: 1.5,
      fontSize: 44,
      fontFace: 'Calibri',
      bold: true,
      color: '363636',
      align: 'center'
    });
    
    // Subtitle with presenter and date
    const subtitle = `Status Report prepared by: ${this.data.name || 'Presenter'}\n${new Date().toLocaleDateString()}`;
    slide.addText(subtitle, {
      x: 1,
      y: 3.5,
      w: 8,
      h: 1,
      fontSize: 18,
      fontFace: 'Calibri',
      color: '666666',
      align: 'center',
      valign: 'middle'
    });
  }

  /**
   * Create overview slide
   */
  private async createOverviewSlide(): Promise<void> {
    const slide = this.pptx.addSlide();
    
    // Title
    slide.addText('Overview', {
      x: 1,
      y: 0.5,
      w: 8,
      h: 1,
      fontSize: 36,
      fontFace: 'Calibri',
      bold: true,
      color: '363636'
    });
    
    // Status update content
    const statusText = this.data.status_update || 'Status update will be displayed here.';
    slide.addText(statusText, {
      x: 1,
      y: 2,
      w: 8,
      h: 4,
      fontSize: 20,
      fontFace: 'Calibri',
      color: '444444',
      valign: 'top',
      wrapText: true
    });
  }

  /**
   * Create agenda slide
   */
  private async createAgendaSlide(): Promise<void> {
    const slide = this.pptx.addSlide();
    
    // Title
    slide.addText('Agenda', {
      x: 1,
      y: 0.5,
      w: 8,
      h: 1,
      fontSize: 36,
      fontFace: 'Calibri',
      bold: true,
      color: '363636'
    });
    
    // Agenda items
    const agendaText = this.data.agenda || 'Technology Stack\nSystem Context\nDesign Decisions\nMobile Application Features';
    const agendaItems = agendaText.split('\n').filter(item => item.trim());
    
    let yPos = 2;
    agendaItems.forEach((item, index) => {
      slide.addText(`${(index + 1).toString().padStart(2, '0')}. ${item}`, {
        x: 1.5,
        y: yPos,
        w: 7,
        h: 0.6,
        fontSize: 20,
        fontFace: 'Calibri',
        color: '444444',
        wrapText: true
      });
      yPos += 0.8;
    });
  }

  /**
   * Create chart slide
   */
  private async createChartSlide(): Promise<void> {
    const slide = this.pptx.addSlide();
    
    // Title
    const chartTitle = this.data.ChartTitle || 'Analysis of Key Metrics';
    slide.addText(chartTitle, {
      x: 1,
      y: 0.5,
      w: 8,
      h: 1,
      fontSize: 36,
      fontFace: 'Calibri',
      bold: true,
      color: '363636'
    });
    
    // Placeholder text for chart
    slide.addText('This is a placeholder slide.\nThe title above will be filled from your Excel data.\n\nCharts defined in the \'ChartData\' sheet will be added as new slides following this one.', {
      x: 1,
      y: 2,
      w: 8,
      h: 3,
      fontSize: 18,
      fontFace: 'Calibri',
      color: '666666',
      align: 'center',
      valign: 'middle',
      wrapText: true
    });
    
    // Add a simple chart placeholder box
    slide.addShape(this.pptx.ShapeType.rect, {
      x: 2,
      y: 4,
      w: 6,
      h: 3,
      fill: { color: 'F0F0F0' },
      line: { color: 'CCCCCC', width: 1 }
    });
    
    slide.addText('📊 Chart data will be automatically inserted here', {
      x: 2,
      y: 5,
      w: 6,
      h: 1,
      fontSize: 16,
      fontFace: 'Calibri',
      color: '888888',
      align: 'center',
      valign: 'middle'
    });
  }

  /**
   * Generate PPTX as buffer
   */
  async generate(): Promise<Buffer> {
    // PPTXGenJS returns a filename when using nodebuffer output
    const fileName = await this.pptx.writeFile({ outputType: 'nodebuffer' }) as string;
    
    // Read the file and return as buffer
    const fs = await import('fs/promises');
    const buffer = await fs.readFile(fileName);
    
    // Clean up the temporary file
    try {
      await fs.unlink(fileName);
    } catch (error) {
      // Ignore cleanup errors
    }
    
    return buffer;
  }

  /**
   * Create additional chart slides from chart data
   */
  async createChartSlides(chartData: any[]): Promise<void> {
    if (!chartData || !Array.isArray(chartData)) {
      return;
    }
    
    // Group charts by title
    const chartGroups: { [key: string]: any[] } = {};
    chartData.forEach(chart => {
      if (!chartGroups[chart.ChartTitle]) {
        chartGroups[chart.ChartTitle] = [];
      }
      chartGroups[chart.ChartTitle].push(chart);
    });
    
    // Create a slide for each chart group
    Object.entries(chartGroups).forEach(([title, charts]) => {
      this.createChartSlideWithData(title, charts);
    });
  }

  /**
   * Create a chart slide with actual data
   */
  private createChartSlideWithData(title: string, charts: any[]): void {
    const slide = this.pptx.addSlide();
    
    // Title
    slide.addText(title, {
      x: 1,
      y: 0.5,
      w: 8,
      h: 1,
      fontSize: 36,
      fontFace: 'Calibri',
      bold: true,
      color: '363636'
    });
    
    // Create a simple table to display chart data
    const tableData = this.formatChartData(charts);
    if (tableData.length > 0) {
      slide.addTable(tableData, {
        x: 1,
        y: 2,
        w: 8,
        h: 4,
        fontSize: 14,
        fontFace: 'Calibri',
        border: { pt: 1, color: 'CCCCCC' },
        fill: { color: 'F9F9F9' },
        align: 'center',
        valign: 'middle'
      });
    }
  }

  /**
   * Format chart data for table display
   */
  private formatChartData(charts: any[]): any[][] {
    if (!charts || charts.length === 0) {
      return [];
    }
    
    const tableData: any[][] = [];
    
    // Header row
    tableData.push(['Series', 'Label', 'Value', 'Type']);
    
    // Data rows
    charts.forEach(chart => {
      tableData.push([
        chart.Series || '',
        chart.Label || '',
        chart.Value?.toString() || '',
        chart.ChartType || ''
      ]);
    });
    
    return tableData;
  }
}