/**
 * PPTX Builder - Preserves original template design while applying data
 */

import JSZip from 'jszip';
import { XMLParser, XMLBuilder } from 'fast-xml-parser';

export interface PPTXData {
  [key: string]: any;
}

export class PPTXBuilder {
  private zip: JSZip;
  private parser: XMLParser;
  private builder: XMLBuilder;
  private data: PPTXData;

  constructor() {
    this.zip = new JSZip();
    this.parser = new XMLParser({
      ignoreAttributes: false,
      attributeNamePrefix: "@_",
      textNodeName: "#text",
      parseAttributeValue: true,
      parseTagValue: true,
      trimValues: true
    });
    this.builder = new XMLBuilder({
      ignoreAttributes: false,
      attributeNamePrefix: "@_",
      textNodeName: "#text",
      format: true
    });
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
    // Process slide files
    const slideFiles = this.zip.file(/ppt\/slides\/slide\d+\.xml/);
    
    for (const file of slideFiles) {
      const content = await file.async('string');
      const processedContent = this.processSlideContent(content);
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
      const processedContent = this.processSlideContent(content);
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
        const slideContent = this.createChartSlideContent(chartTitle, chartGroup as any[], layoutSlide);
        const slideFileName = `ppt/slides/slide${slideNum}.xml`;
        
        this.zip.file(slideFileName, slideContent);
        slideNum++;
      }

      // Update presentation.xml to include new slides
      await this.updatePresentationXml(currentSlideCount, newSlidesCount);
      
      // Update content types
      await this.updateContentTypes(currentSlideCount, newSlidesCount);
      
      // Update slide relationships
      await this.updateSlideRelationships(currentSlideCount, newSlidesCount);
    }

    // Generate the buffer
    return await this.zip.generateAsync({ type: 'arraybuffer' });
  }

  /**
   * Process slide content and replace placeholders
   */
  private processSlideContent(xmlContent: string): string {
    try {
      const parsed = this.parser.parse(xmlContent);
      this.processNode(parsed);
      return this.builder.build(parsed);
    } catch (error) {
      console.error('Error processing slide content:', error);
      return xmlContent;
    }
  }

  /**
   * Recursively process XML nodes to replace placeholders
   */
  private processNode(node: any): void {
    if (!node) return;

    // Process text nodes
    if (node['#text'] && typeof node['#text'] === 'string') {
      node['#text'] = this.replacePlaceholders(node['#text']);
    }

    // Process a:t (text) elements specifically
    if (node['a:t'] && typeof node['a:t'] === 'string') {
      node['a:t'] = this.replacePlaceholders(node['a:t']);
    }

    // Process arrays of nodes
    if (Array.isArray(node)) {
      node.forEach(child => this.processNode(child));
    }

    // Process object nodes
    if (typeof node === 'object' && node !== null) {
      Object.keys(node).forEach(key => {
        if (key !== '#text') {
          this.processNode(node[key]);
        }
      });
    }
  }

  /**
   * Replace placeholders in text with actual data
   */
  private replacePlaceholders(text: string): string {
    if (!text || typeof text !== 'string') return text;

    return text.replace(/\{([^}]+)\}/g, (match, key) => {
      const trimmedKey = key.trim();
      const value = this.data[trimmedKey];
      
      if (value !== undefined && value !== null) {
        return String(value);
      }

      // Handle special cases
      switch (trimmedKey) {
        case 'date':
          return new Date().toLocaleDateString();
        case 'presenter':
          return this.data.name || 'Presenter';
        default:
          return match; // Keep original placeholder if no data found
      }
    });
  }

  /**
   * Add chart slides based on chart data
   */
  async addChartSlides(chartData: any[]): Promise<void> {
    if (!chartData || chartData.length === 0) return;

    // Group chart data by title
    const chartGroups = this.groupChartData(chartData);
    
    // Get the current highest slide number
    const slideFiles = this.zip.file(/ppt\/slides\/slide\d+\.xml/);
    const maxSlideNum = slideFiles.length;
    
    // Get slide layout to copy
    const layoutSlide = await this.getSlideLayout();
    
    // Add chart slides
    let slideNum = maxSlideNum + 1;
    for (const [chartTitle, chartGroup] of Object.entries(chartGroups)) {
      const slideContent = this.createChartSlideContent(chartTitle, chartGroup as any[], layoutSlide);
      const slideFileName = `ppt/slides/slide${slideNum}.xml`;
      
      this.zip.file(slideFileName, slideContent);
      slideNum++;
    }

    // Update presentation.xml to include new slides
    await this.updatePresentationXml(maxSlideNum, Object.keys(chartGroups).length);
    
    // Update content types
    await this.updateContentTypes(maxSlideNum, Object.keys(chartGroups).length);
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
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr>
    <a:masterClrMapping/>
  </p:clrMapOvr>
</p:sld>`;
  }

  /**
   * Create chart slide content
   */
  private createChartSlideContent(chartTitle: string, chartGroup: any[], layoutSlide: string): string {
    try {
      const parsed = this.parser.parse(layoutSlide);
      
      // Replace title placeholder
      this.replaceTextInNode(parsed, chartTitle);
      
      // Add chart placeholder text
      this.addChartPlaceholder(parsed, chartGroup);
      
      return this.builder.build(parsed);
    } catch (error) {
      console.error('Error creating chart slide:', error);
      return layoutSlide;
    }
  }

  /**
   * Replace text in parsed XML node
   */
  private replaceTextInNode(node: any, newText: string): void {
    if (!node) return;

    // Find and replace title text
    if (node['p:sld'] && node['p:sld']['p:cSld'] && node['p:sld']['p:cSld']['p:spTree']) {
      const shapes = node['p:sld']['p:cSld']['p:spTree']['p:sp'];
      if (Array.isArray(shapes)) {
        shapes.forEach((shape: any) => {
          if (shape['p:nvSpPr'] && shape['p:nvSpPr']['p:nvPr'] && 
              shape['p:nvSpPr']['p:nvPr']['p:ph'] && 
              shape['p:nvSpPr']['p:nvPr']['p:ph']['@_type'] === 'title') {
            
            if (shape['p:txBody'] && shape['p:txBody']['a:p'] && shape['p:txBody']['a:p']['a:r']) {
              const textRuns = shape['p:txBody']['a:p']['a:r'];
              if (Array.isArray(textRuns)) {
                textRuns.forEach((run: any) => {
                  if (run['a:t']) {
                    run['a:t'] = newText;
                  }
                });
              } else if (textRuns && textRuns['a:t']) {
                textRuns['a:t'] = newText;
              }
            }
          }
        });
      }
    }
  }

  /**
   * Add chart placeholder to slide
   */
  private addChartPlaceholder(node: any, chartGroup: any[]): void {
    // For now, just add a text placeholder indicating chart data
    // In a full implementation, you would add actual chart elements
    const chartInfo = `Chart data: ${chartGroup.length} data points`;
    
    // Find content placeholder and add chart info
    if (node['p:sld'] && node['p:sld']['p:cSld'] && node['p:sld']['p:cSld']['p:spTree']) {
      const shapes = node['p:sld']['p:cSld']['p:spTree']['p:sp'];
      if (Array.isArray(shapes)) {
        shapes.forEach((shape: any) => {
          if (shape['p:nvSpPr'] && shape['p:nvSpPr']['p:nvPr'] && 
              shape['p:nvSpPr']['p:nvPr']['p:ph'] && 
              shape['p:nvSpPr']['p:nvPr']['p:ph']['@_type'] === 'body') {
            
            if (shape['p:txBody'] && shape['p:txBody']['a:p']) {
              const paragraphs = shape['p:txBody']['a:p'];
              if (Array.isArray(paragraphs)) {
                paragraphs.forEach((para: any) => {
                  if (para['a:r'] && para['a:r']['a:t']) {
                    para['a:r']['a:t'] = chartInfo;
                  }
                });
              } else if (paragraphs && paragraphs['a:r'] && paragraphs['a:r']['a:t']) {
                paragraphs['a:r']['a:t'] = chartInfo;
              }
            }
          }
        });
      }
    }
  }

  /**
   * Update presentation.xml to include new slides
   */
  private async updatePresentationXml(currentSlideCount: number, newSlidesCount: number): Promise<void> {
    const presFile = this.zip.file('ppt/presentation.xml');
    if (!presFile) return;

    const presContent = await presFile.async('string');
    const parsed = this.parser.parse(presContent);

    // Update slide count
    if (parsed['p:presentation'] && parsed['p:presentation']['p:sldIdLst']) {
      const slideIdList = parsed['p:presentation']['p:sldIdLst']['p:sldId'];
      if (Array.isArray(slideIdList)) {
        const lastSlideId = slideIdList[slideIdList.length - 1];
        const lastId = parseInt(lastSlideId['@_id']);
        
        // Add new slide IDs
        for (let i = 1; i <= newSlidesCount; i++) {
          const newSlideId = {
            '@_id': lastId + i,
            '@_r:id': `rId${currentSlideCount + i + 2}` // +2 to account for slide master and theme relationships
          };
          slideIdList.push(newSlideId);
        }
      }
    }

    this.zip.file('ppt/presentation.xml', this.builder.build(parsed));
  }

  /**
   * Update [Content_Types].xml to include new slide types
   */
  private async updateContentTypes(currentSlideCount: number, newSlidesCount: number): Promise<void> {
    const contentTypesFile = this.zip.file('[Content_Types].xml');
    if (!contentTypesFile) return;

    const contentTypesContent = await contentTypesFile.async('string');
    const parsed = this.parser.parse(contentTypesContent);

    // Add new slide overrides
    if (parsed['Types'] && parsed['Types']['Override']) {
      const overrides = parsed['Types']['Override'];
      if (!Array.isArray(overrides)) {
        parsed['Types']['Override'] = [overrides];
      }
      
      // Add slide content types for new slides
      for (let i = 1; i <= newSlidesCount; i++) {
        const slideNum = currentSlideCount + i;
        parsed['Types']['Override'].push({
          '@_PartName': `/ppt/slides/slide${slideNum}.xml`,
          '@_ContentType': 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml'
        });
      }
    }

    this.zip.file('[Content_Types].xml', this.builder.build(parsed));
  }

  /**
   * Update slide relationships to include new slides
   */
  private async updateSlideRelationships(currentSlideCount: number, newSlidesCount: number): Promise<void> {
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