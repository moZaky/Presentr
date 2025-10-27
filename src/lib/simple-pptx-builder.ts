/**
 * Simple PPTX Builder - Minimal approach to avoid corruption
 */

import * as JSZip from 'jszip';

export interface PPTXData {
  [key: string]: any;
}

export class SimplePPTXBuilder {
  private zip: JSZip;
  private data: PPTXData;

  constructor() {
    this.zip = new JSZip();
    this.data = {};
  }

  /**
   * Load PPTX template from buffer
   */
  async loadTemplate(templateBuffer: ArrayBuffer): Promise<void> {
    try {
      this.zip = await JSZip.loadAsync(templateBuffer);
    } catch (error) {
      console.error('Error loading template:', error);
      throw new Error('Failed to load PPTX template');
    }
  }

  /**
   * Set data for replacement
   */
  setData(data: PPTXData): void {
    this.data = data;
  }

  /**
   * Simple XML escaping - only escape what's absolutely necessary
   */
  private escapeXmlContent(content: string): string {
    if (typeof content !== 'string') {
      content = String(content || '');
    }
    // Only escape the most critical characters that break XML
    return content
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
  }

  /**
   * Replace placeholders in all slide XML files
   */
  async replacePlaceholders(): Promise<void> {
    const slideFiles = Object.keys(this.zip.files)
      .filter(filename => filename.startsWith('ppt/slides/slide') && filename.endsWith('.xml'));

    for (const slideFile of slideFiles) {
      await this.replaceInSlide(slideFile);
    }
  }

  /**
   * Replace placeholders in a specific slide - minimal approach
   */
  private async replaceInSlide(slideFile: string): Promise<void> {
    const slideXml = await this.zip.file(slideFile)?.async('string');
    if (!slideXml) return;

    let updatedXml = slideXml;

    // Simple placeholder replacement - no complex processing
    Object.entries(this.data).forEach(([key, value]) => {
      const placeholder = `{${key}}`;
      if (updatedXml.includes(placeholder)) {
        const escapedValue = this.escapeXmlContent(String(value || ''));
        // Simple string replacement - no regex complexity
        updatedXml = updatedXml.split(placeholder).join(escapedValue);
      }
    });

    // Handle special placeholders like date
    updatedXml = this.handleSpecialPlaceholders(updatedXml);

    // Update the slide content
    this.zip.file(slideFile, updatedXml);
  }

  /**
   * Handle special placeholders like date
   */
  private handleSpecialPlaceholders(xml: string): string {
    const commonPlaceholders = {
      date: () => new Date().toLocaleDateString(),
      time: () => new Date().toLocaleTimeString(),
      datetime: () => new Date().toLocaleString(),
      year: () => new Date().getFullYear().toString(),
      month: () => new Date().toLocaleDateString('en-US', { month: 'long' }),
      day: () => new Date().getDate().toString()
    };

    Object.entries(commonPlaceholders).forEach(([key, generator]) => {
      if (!this.data[key]) {
        const placeholder = `{${key}}`;
        if (xml.includes(placeholder)) {
          xml = xml.split(placeholder).join(generator());
        }
      }
    });
    
    return xml;
  }

  /**
   * Process chart data - simplified approach
   */
  async processChartData(chartData: any[]): Promise<void> {
    if (!chartData || !Array.isArray(chartData) || chartData.length === 0) {
      return;
    }

    // Group charts by title
    const chartGroups: { [key: string]: any[] } = {};
    chartData.forEach(chart => {
      const titleField = Object.keys(chart).find(key => 
        key.toLowerCase().includes('title') || 
        key.toLowerCase().includes('name') ||
        key.toLowerCase().includes('chart')
      ) || 'ChartTitle';
      
      const title = chart[titleField] || 'Untitled Chart';
      if (!chartGroups[title]) {
        chartGroups[title] = [];
      }
      chartGroups[title].push(chart);
    });

    // Find chart placeholder slides
    const placeholderSlides = await this.findChartPlaceholderSlides();
    
    // Create chart slides
    for (const [chartTitle, charts] of Object.entries(chartGroups)) {
      await this.createChartSlide(chartTitle, charts, placeholderSlides);
    }
  }

  /**
   * Find chart placeholder slides
   */
  private async findChartPlaceholderSlides(): Promise<string[]> {
    const slideFiles = Object.keys(this.zip.files)
      .filter(filename => filename.startsWith('ppt/slides/slide') && filename.endsWith('.xml'));

    const placeholderSlides: string[] = [];
    
    for (const slideFile of slideFiles) {
      const content = await this.zip.file(slideFile)?.async('string');
      if (content) {
        const chartPlaceholders = content.match(/\{[^}]*chart[^}]*\}/gi) || 
                                 content.match(/\{[^}]*title[^}]*\}/gi) ||
                                 content.match(/\{[^}]*data[^}]*\}/gi);
        
        if (chartPlaceholders && chartPlaceholders.length > 0) {
          placeholderSlides.push(slideFile);
        }
      }
    }
    
    return placeholderSlides;
  }

  /**
   * Create chart slide - simplified
   */
  private async createChartSlide(chartTitle: string, charts: any[], placeholderSlides: string[]): Promise<void> {
    if (placeholderSlides.length === 0) return;

    const placeholderSlideFile = placeholderSlides[0];
    const placeholderSlideContent = await this.zip.file(placeholderSlideFile)?.async('string');
    
    if (!placeholderSlideContent) return;

    // Get next slide number
    const slideFiles = Object.keys(this.zip.files)
      .filter(filename => filename.startsWith('ppt/slides/slide') && filename.endsWith('.xml'));
    
    const slideNumbers = slideFiles.map(file => {
      const match = file.match(/slide(\d+)\.xml/);
      return match ? parseInt(match[1]) : 0;
    });
    const nextSlideNumber = Math.max(...slideNumbers) + 1;

    // Create new slide
    const newSlideFile = `ppt/slides/slide${nextSlideNumber}.xml`;
    let newSlideContent = placeholderSlideContent;

    // Replace placeholders
    Object.entries(this.data).forEach(([key, value]) => {
      const placeholder = `{${key}}`;
      if (newSlideContent.includes(placeholder)) {
        const escapedValue = this.escapeXmlContent(String(value || ''));
        newSlideContent = newSlideContent.split(placeholder).join(escapedValue);
      }
    });

    // Replace chart title
    const chartTitlePlaceholder = this.findChartTitlePlaceholder(placeholderSlideContent);
    if (chartTitlePlaceholder) {
      const placeholder = `{${chartTitlePlaceholder}}`;
      if (newSlideContent.includes(placeholder)) {
        const escapedTitle = this.escapeXmlContent(chartTitle);
        newSlideContent = newSlideContent.split(placeholder).join(escapedTitle);
      }
    }

    // Add chart data
    const chartDataText = this.formatChartDataForSlide(charts);
    newSlideContent = newSlideContent.replace(/This is a placeholder slide\.[\s\S]*?Charts defined in[^.]*\.?/g, chartDataText);

    // Remove remaining placeholders
    newSlideContent = newSlideContent.replace(/{[^}]*}/g, '');

    this.zip.file(newSlideFile, newSlideContent);

    // Update presentation files
    await this.updatePresentationXml(nextSlideNumber);
    await this.updatePresentationRels(nextSlideNumber);
  }

  /**
   * Find chart title placeholder
   */
  private findChartTitlePlaceholder(slideContent: string): string {
    const chartPlaceholders = slideContent.match(/\{[^}]*\}/g) || [];
    
    for (const placeholder of chartPlaceholders) {
      const cleanName = placeholder.replace(/[{}]/g, '').toLowerCase();
      if (cleanName.includes('chart') || cleanName.includes('title')) {
        return cleanName;
      }
    }
    
    return chartPlaceholders.length > 0 ? chartPlaceholders[0].replace(/[{}]/g, '') : 'ChartTitle';
  }

  /**
   * Format chart data for slide
   */
  private formatChartDataForSlide(charts: any[]): string {
    if (!charts || charts.length === 0) {
      return 'No chart data available.';
    }

    const lines = ['Chart Data Summary', ''];
    
    const groupedData: { [key: string]: any[] } = {};
    charts.forEach(chart => {
      const series = chart.Series || chart.series || 'Data';
      if (!groupedData[series]) {
        groupedData[series] = [];
      }
      groupedData[series].push(chart);
    });

    Object.entries(groupedData).forEach(([series, data]) => {
      lines.push(this.escapeXmlContent(series));
      data.forEach(item => {
        const label = this.escapeXmlContent(String(item.Label || item.label || 'N/A'));
        const value = this.escapeXmlContent(String(item.Value || item.value || 'N/A'));
        const type = this.escapeXmlContent(String(item.ChartType || item.type || ''));
        lines.push(`  • ${label}: ${value} ${type ? `(${type})` : ''}`);
      });
      lines.push('');
    });

    return lines.join('\n');
  }

  /**
   * Update presentation.xml
   */
  private async updatePresentationXml(newSlideNumber: number): Promise<void> {
    const presFile = 'ppt/presentation.xml';
    const presXml = await this.zip.file(presFile)?.async('string');
    if (!presXml) return;

    const slideIdListMatch = presXml.match(/<p:sldIdLst>([\s\S]*?)<\/p:sldIdLst>/);
    if (!slideIdListMatch) return;

    const idMatches = presXml.match(/<p:sldId id="(\d+)"/g);
    const existingIds = idMatches ? idMatches.map(match => parseInt(match.match(/id="(\d+)"/)![1])) : [];
    const newId = existingIds.length > 0 ? Math.max(...existingIds) + 1 : 256;

    const rIdMatches = presXml.match(/r:id="rId(\d+)"/g);
    const existingRIds = rIdMatches ? rIdMatches.map(match => parseInt(match.match(/r:id="rId(\d+)"/)![1])) : [];
    const newRId = existingRIds.length > 0 ? Math.max(...existingRIds) + 1 : 2;

    const newSlideId = `        <p:sldId id="${newId}" r:id="rId${newRId}"/>`;
    const updatedSlideIdList = slideIdListMatch[1] + '\n' + newSlideId;
    const updatedPresXml = presXml.replace(
      /<p:sldIdLst>[\s\S]*?<\/p:sldIdLst>/,
      `<p:sldIdLst>${updatedSlideIdList}\n    </p:sldIdLst>`
    );

    this.zip.file(presFile, updatedPresXml);
  }

  /**
   * Update presentation.xml.rels
   */
  private async updatePresentationRels(newSlideNumber: number): Promise<void> {
    const relsFile = 'ppt/_rels/presentation.xml.rels';
    const relsXml = await this.zip.file(relsFile)?.async('string');
    if (!relsXml) return;

    const rIdMatches = relsXml.match(/Id="rId(\d+)"/g);
    const existingRIds = rIdMatches ? rIdMatches.map(match => parseInt(match.match(/Id="rId(\d+)"/)![1])) : [];
    const newRId = existingRIds.length > 0 ? Math.max(...existingRIds) + 1 : 2;

    const newRelationship = `    <Relationship Id="rId${newRId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${newSlideNumber}.xml"/>`;
    const updatedRelsXml = relsXml.replace(
      /<\/Relationships>/,
      `${newRelationship}\n<\/Relationships>`
    );

    this.zip.file(relsFile, updatedRelsXml);
  }

  /**
   * Clean up remaining placeholders
   */
  private async cleanupRemainingPlaceholders(): Promise<void> {
    const slideFiles = Object.keys(this.zip.files)
      .filter(filename => filename.startsWith('ppt/slides/slide') && filename.endsWith('.xml'));

    for (const slideFile of slideFiles) {
      const slideXml = await this.zip.file(slideFile)?.async('string');
      if (slideXml) {
        // Remove any remaining template blocks
        let cleanedXml = slideXml
          .replace(/\{#[^}]*\}[\s\S]*?\{\/[^}]*\}/g, '') // Remove template blocks
          .replace(/{[^}]*}/g, ''); // Remove remaining placeholders

        this.zip.file(slideFile, cleanedXml);
      }
    }
  }

  /**
   * Generate the final PPTX as buffer
   */
  async generate(): Promise<Buffer> {
    return await this.zip.generateAsync({ type: 'nodebuffer' });
  }

  /**
   * Build the complete presentation with charts
   */
  async buildWithCharts(chartData: any[]): Promise<Buffer> {
    try {
      await this.replacePlaceholders();
      await this.processChartData(chartData);
      await this.cleanupRemainingPlaceholders();
      return await this.generate();
    } catch (error) {
      console.error('Error during PPTX build:', error);
      throw new Error(`PPTX build failed: ${(error as Error).message}`);
    }
  }
}