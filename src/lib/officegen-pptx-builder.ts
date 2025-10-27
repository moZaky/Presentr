/**
 * OfficeGen PPTX Builder - Creates fully compatible PPTX files using officegen
 */

// eslint-disable-next-line @typescript-eslint/no-require-imports
const officegen = require('officegen');

export interface PPTXData {
  [key: string]: any;
}

export class OfficeGenPPTXBuilder {
  private pptx: any;
  private data: PPTXData;

  constructor() {
    this.pptx = officegen('pptx');
    this.data = {};
    this.setupDefaults();
  }

  /**
   * Setup default PPTX properties
   */
  private setupDefaults(): void {
    // OfficeGen doesn't have these methods, so we'll skip them
    // The properties will be set with default values automatically
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
    const slide = this.pptx.makeNewSlide();
    slide.name = 'Title Slide';
    slide.back = 'ffffff';
    
    // Title
    slide.addText(this.data.project_title || 'Project Title', {
      x: 1,
      y: 1.5,
      cx: 8,
      cy: 1.5,
      font_size: 44,
      font_face: 'Calibri',
      bold: true,
      color: '363636',
      align: 'center'
    });
    
    // Subtitle with presenter and date
    const subtitle = `Status Report prepared by: ${this.data.name || 'Presenter'}\n${new Date().toLocaleDateString()}`;
    slide.addText(subtitle, {
      x: 1,
      y: 3.5,
      cx: 8,
      cy: 1,
      font_size: 18,
      font_face: 'Calibri',
      color: '666666',
      align: 'center'
    });
  }

  /**
   * Create overview slide
   */
  private async createOverviewSlide(): Promise<void> {
    const slide = this.pptx.makeNewSlide();
    slide.name = 'Overview';
    slide.back = 'ffffff';
    
    // Title
    slide.addText('Overview', {
      x: 1,
      y: 0.5,
      cx: 8,
      cy: 1,
      font_size: 36,
      font_face: 'Calibri',
      bold: true,
      color: '363636'
    });
    
    // Status update content
    const statusText = this.data.status_update || 'Status update will be displayed here.';
    slide.addText(statusText, {
      x: 1,
      y: 2,
      cx: 8,
      cy: 4,
      font_size: 20,
      font_face: 'Calibri',
      color: '444444',
      valign: 'top'
    });
  }

  /**
   * Create agenda slide
   */
  private async createAgendaSlide(): Promise<void> {
    const slide = this.pptx.makeNewSlide();
    slide.name = 'Agenda';
    slide.back = 'ffffff';
    
    // Title
    slide.addText('Agenda', {
      x: 1,
      y: 0.5,
      cx: 8,
      cy: 1,
      font_size: 36,
      font_face: 'Calibri',
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
        cx: 7,
        cy: 0.6,
        font_size: 20,
        font_face: 'Calibri',
        color: '444444'
      });
      yPos += 0.8;
    });
  }

  /**
   * Create chart slide
   */
  private async createChartSlide(): Promise<void> {
    const slide = this.pptx.makeNewSlide();
    slide.name = 'Chart Slide';
    slide.back = 'ffffff';
    
    // Title
    const chartTitle = this.data.ChartTitle || 'Analysis of Key Metrics';
    slide.addText(chartTitle, {
      x: 1,
      y: 0.5,
      cx: 8,
      cy: 1,
      font_size: 36,
      font_face: 'Calibri',
      bold: true,
      color: '363636'
    });
    
    // Placeholder text for chart
    slide.addText('This is a placeholder slide.\nThe title above will be filled from your Excel data.\n\nCharts defined in the \'ChartData\' sheet will be added as new slides following this one.', {
      x: 1,
      y: 2,
      cx: 8,
      cy: 3,
      font_size: 18,
      font_face: 'Calibri',
      color: '666666',
      align: 'center'
    });
    
    // Add a simple chart placeholder box
    slide.addShape('rect', {
      x: 2,
      y: 4,
      cx: 6,
      cy: 3,
      fill: 'f0f0f0',
      line: 'cccccc',
      line_size: 1
    });
    
    slide.addText('📊 Chart data will be automatically inserted here', {
      x: 2,
      y: 5,
      cx: 6,
      cy: 1,
      font_size: 16,
      font_face: 'Calibri',
      color: '888888',
      align: 'center'
    });
  }

  /**
   * Generate PPTX as buffer
   */
  async generate(): Promise<Buffer> {
    return new Promise((resolve, reject) => {
      const chunks: Buffer[] = [];
      
      // OfficeGen generate method with callback
      this.pptx.generate((err: Error | null, stream: any) => {
        if (err) {
          reject(err);
          return;
        }
        
        // The stream should be a readable stream
        if (stream && typeof stream.on === 'function') {
          stream.on('data', (chunk: Buffer) => {
            chunks.push(chunk);
          });
          
          stream.on('end', () => {
            const buffer = Buffer.concat(chunks);
            resolve(buffer);
          });
          
          stream.on('error', (error: Error) => {
            reject(error);
          });
        } else if (stream && Buffer.isBuffer(stream)) {
          // If it's already a buffer
          resolve(stream);
        } else {
          // Fallback
          reject(new Error('Unable to generate PPTX buffer'));
        }
      });
    });
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
    const slide = this.pptx.makeNewSlide();
    slide.name = `Chart: ${title}`;
    slide.back = 'ffffff';
    
    // Title
    slide.addText(title, {
      x: 1,
      y: 0.5,
      cx: 8,
      cy: 1,
      font_size: 36,
      font_face: 'Calibri',
      bold: true,
      color: '363636'
    });
    
    // Create a simple table to display chart data
    const tableData = this.formatChartData(charts);
    if (tableData.length > 0) {
      slide.addTable(tableData, {
        x: 1,
        y: 2,
        cx: 8,
        cy: 4,
        font_size: 14,
        font_face: 'Calibri',
        border: { pt: 1, color: 'cccccc' },
        fill: { color: 'f9f9f9' }
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