/**
 * Safe PPTX Builder - Minimal XML manipulation for maximum PowerPoint compatibility
 * Uses only string replacement to avoid XML parsing issues
 */

import JSZip from 'jszip';

export interface PPTXData {
  [key: string]: any;
}

export class SafePPTXBuilder {
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
    this.zip = await JSZip.loadAsync(templateBuffer);
  }

  /**
   * Set data to apply to template
   */
  setData(data: PPTXData): void {
    this.data = data;
  }

  /**
   * Build the final PPTX with data applied
   */
  async build(): Promise<ArrayBuffer> {
    // Process slide files using string replacement only
    const slideFiles = this.zip.file(/ppt\/slides\/slide\d+\.xml/);
    
    for (const file of slideFiles) {
      const content = await file.async('string');
      const processedContent = this.processSlideContentString(content);
      this.zip.file(file.name, processedContent);
    }

    // Generate the buffer
    return await this.zip.generateAsync({ type: 'arraybuffer' });
  }

  /**
   * Build the final PPTX with data and chart slides
   */
  async buildWithCharts(chartData: any[]): Promise<ArrayBuffer> {
    // Process existing slide files
    const slideFiles = this.zip.file(/ppt\/slides\/slide\d+\.xml/);
    
    for (const file of slideFiles) {
      const content = await file.async('string');
      const processedContent = this.processSlideContentString(content);
      this.zip.file(file.name, processedContent);
    }

    // Add chart slides if chart data exists
    if (chartData && chartData.length > 0) {
      const currentSlideCount = slideFiles.length;
      const chartGroups = this.groupChartData(chartData);
      const newSlidesCount = Object.keys(chartGroups).length;
      
      // Get the current highest slide number
      const maxSlideNum = currentSlideCount;
      
      // Get slide layout to copy
      const layoutSlide = await this.getSlideLayout();
      
      // Add chart slides
      let slideNum = maxSlideNum + 1;
      for (const [chartTitle, chartGroup] of Object.entries(chartGroups)) {
        const slideContent = this.createChartSlideContentString(chartTitle, chartGroup as any[], layoutSlide);
        const slideFileName = `ppt/slides/slide${slideNum}.xml`;
        
        this.zip.file(slideFileName, slideContent);
        slideNum++;
      }

      // Update presentation.xml to include new slides
      await this.updatePresentationXmlString(currentSlideCount, newSlidesCount);
      
      // Update content types
      await this.updateContentTypesString(currentSlideCount, newSlidesCount);
      
      // Update slide relationships
      await this.updateSlideRelationshipsString(currentSlideCount, newSlidesCount);
    }

    // Generate the buffer
    return await this.zip.generateAsync({ type: 'arraybuffer' });
  }

  /**
   * Process slide content using only string replacement
   */
  private processSlideContentString(xmlContent: string): string {
    let processedContent = xmlContent;
    
    // Process each data item
    for (const [key, value] of Object.entries(this.data)) {
      const placeholder = `{${key}}`;
      if (processedContent.includes(placeholder)) {
        const escapedValue = this.escapeXmlText(String(value || ''));
        processedContent = processedContent.split(placeholder).join(escapedValue);
      }
    }

    // Handle special placeholders
    const specialPlaceholders = {
      'date': new Date().toLocaleDateString(),
      'presenter': this.data.name || 'Presenter'
    };

    for (const [placeholder, value] of Object.entries(specialPlaceholders)) {
      const placeholderPattern = `{${placeholder}}`;
      if (processedContent.includes(placeholderPattern)) {
        const escapedValue = this.escapeXmlText(String(value));
        processedContent = processedContent.split(placeholderPattern).join(escapedValue);
      }
    }

    return processedContent;
  }

  /**
   * Escape text for XML - minimal escaping for PowerPoint compatibility
   */
  private escapeXmlText(text: string): string {
    return text
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  }

  /**
   * Group chart data by title
   */
  private groupChartData(chartData: any[]): Record<string, any[]> {
    const grouped: Record<string, any[]> = {};
    
    chartData.forEach(item => {
      const title = item.ChartTitle || 'Chart';
      if (!grouped[title]) {
        grouped[title] = [];
      }
      grouped[title].push(item);
    });
    
    return grouped;
  }

  /**
   * Get a slide layout to copy for new slides
   */
  private async getSlideLayout(): Promise<string> {
    // Try to get the last slide as a template
    const slideFiles = this.zip.file(/ppt\/slides\/slide\d+\.xml/);
    if (slideFiles.length > 0) {
      const lastSlide = slideFiles[slideFiles.length - 1];
      return await lastSlide.async('string');
    }
    
    // Fallback to a basic slide structure
    return this.getBasicSlideStructure();
  }

  /**
   * Get basic slide structure as fallback
   */
  private getBasicSlideStructure(): string {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr>
        <p:cNvPr id="1" name=""/>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr>
        <a:xfrm>
          <a:off x="0" y="0"/>
          <a:ext cx="0" cy="0"/>
          <a:chOff x="0" y="0"/>
          <a:chExt cx="0" cy="0"/>
        </a:xfrm>
      </p:grpSpPr>
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="2" name="Title"/>
          <p:cNvSpPr>
            <a:spLocks noGrp="1"/>
          </p:cNvSpPr>
          <p:nvPr>
            <p:ph type="title"/>
          </p:nvPr>
        </p:nvSpPr>
        <p:spPr>
          <a:xfrm>
            <a:off x="457200" y="274320"/>
            <a:ext cx="8229600" cy="1143000"/>
          </a:xfrm>
        </p:spPr>
        <p:txBody>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p>
            <a:pPr algn="ctr">
              <a:buNone/>
            </a:pPr>
            <a:r>
              <a:rPr lang="en-US" sz="4400" b="1"/>
              <a:t>Chart Title</a:t>
            </a:r>
          </a:p>
        </p:txBody>
      </p:sp>
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="3" name="Content"/>
          <p:cNvSpPr>
            <a:spLocks noGrp="1"/>
          </p:cNvSpPr>
          <p:nvPr>
            <p:ph type="body"/>
          </p:nvPr>
        </p:nvSpPr>
        <p:spPr>
          <a:xfrm>
            <a:off x="457200" y="1600200"/>
            <a:ext cx="8229600" cy="4527520"/>
          </a:xfrm>
        </p:spPr>
        <p:txBody>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p>
            <a:pPr>
              <a:buNone/>
            </a:pPr>
            <a:r>
              <a:rPr lang="en-US" sz="2800"/>
              <a:t>Chart content will be displayed here.</a:t>
            </a:r>
          </a:p>
        </p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr>
    <a:masterClrMapping/>
  </p:clrMapOvr>
</p:sld>`;
  }

  /**
   * Create chart slide content using string replacement
   */
  private createChartSlideContentString(chartTitle: string, chartGroup: any[], layoutSlide: string): string {
    let slideContent = layoutSlide;
    
    // Replace title - look for common title patterns
    const titlePatterns = [
      '>Chart Title<',
      '>Title<',
      '>Slide Title<',
      '<a:t>Chart Title</a:t>',
      '<a:t>Title</a:t>',
      '<a:t>Slide Title</a:t>'
    ];
    
    const escapedTitle = this.escapeXmlText(chartTitle);
    titlePatterns.forEach(pattern => {
      slideContent = slideContent.split(pattern).join('>' + escapedTitle + '<');
    });
    
    // Replace content with chart information
    const chartInfo = chartGroup.map(item => 
      `${item.Series || 'Data'}: ${item.Value || item.Label || 'N/A'}`
    ).join(', ');
    
    const contentPatterns = [
      '>Chart content will be displayed here.<',
      '>Content<',
      '>Slide content<',
      '<a:t>Chart content will be displayed here.</a:t>',
      '<a:t>Content</a:t>',
      '<a:t>Slide content</a:t>'
    ];
    
    const escapedChartInfo = this.escapeXmlText(chartInfo);
    contentPatterns.forEach(pattern => {
      slideContent = slideContent.split(pattern).join('>' + escapedChartInfo + '<');
    });
    
    return slideContent;
  }

  /**
   * Update presentation.xml using string replacement
   */
  private async updatePresentationXmlString(currentSlideCount: number, newSlidesCount: number): Promise<void> {
    const presFile = this.zip.file('ppt/presentation.xml');
    if (!presFile) return;

    const presContent = await presFile.async('string');
    let updatedContent = presContent;
    
    // Find the last slide ID and increment it
    const slideIdMatch = updatedContent.match(/<p:sldId[^>]*id="(\d+)"[^>]*>/);
    if (slideIdMatch) {
      const lastId = parseInt(slideIdMatch[1]);
      
      // Add new slide IDs
      let newSlideIds = '';
      for (let i = 1; i <= newSlidesCount; i++) {
        const newId = lastId + i;
        const rId = currentSlideCount + i + 2; // +2 to account for slide master and theme relationships
        newSlideIds += `    <p:sldId id="${newId}" r:id="rId${rId}"/>\n`;
      }
      
      // Insert new slide IDs before the closing tag
      updatedContent = updatedContent.replace(
        '</p:sldIdLst>',
        newSlideIds + '  </p:sldIdLst>'
      );
    }
    
    this.zip.file('ppt/presentation.xml', updatedContent);
  }

  /**
   * Update [Content_Types].xml using string replacement
   */
  private async updateContentTypesString(currentSlideCount: number, newSlidesCount: number): Promise<void> {
    const contentTypesFile = this.zip.file('[Content_Types].xml');
    if (!contentTypesFile) return;

    const contentTypesContent = await contentTypesFile.async('string');
    let updatedContent = contentTypesContent;
    
    // Add new slide overrides
    let newOverrides = '';
    for (let i = 1; i <= newSlidesCount; i++) {
      const slideNum = currentSlideCount + i;
      newOverrides += `    <Override PartName="/ppt/slides/slide${slideNum}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>\n`;
    }
    
    // Insert before the closing Types tag
    updatedContent = updatedContent.replace(
      '</Types>',
      newOverrides + '</Types>'
    );
    
    this.zip.file('[Content_Types].xml', updatedContent);
  }

  /**
   * Update slide relationships using string replacement
   */
  private async updateSlideRelationshipsString(currentSlideCount: number, newSlidesCount: number): Promise<void> {
    // Create new relationship files for each new slide
    for (let i = 1; i <= newSlidesCount; i++) {
      const slideNum = currentSlideCount + i;
      const relsContent = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
</Relationships>`;
      
      this.zip.file(`ppt/slides/_rels/slide${slideNum}.xml.rels`, relsContent);
    }
  }
}