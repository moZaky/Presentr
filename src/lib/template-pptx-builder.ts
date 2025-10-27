/**
 * Template PPTX Builder - Preserves original design, only replaces placeholders
 */

import * as JSZip from 'jszip';

export interface PPTXData {
  [key: string]: any;
}

export class TemplatePPTXBuilder {
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
   * Replace placeholders in a specific slide
   */
  private async replaceInSlide(slideFile: string): Promise<void> {
    const slideXml = await this.zip.file(slideFile)?.async('string');
    if (!slideXml) {
      return;
    }

    let updatedXml = slideXml;

    // Handle agenda table format: {#agenda}{index} {title}{/agenda}
    updatedXml = this.processAgendaTemplate(updatedXml);

    // Replace regular placeholders (agenda will be handled by processAgendaTemplate if template format exists)
    Object.entries(this.data).forEach(([key, value]) => {
      // Skip agenda if it was already processed as a template
      if (key === 'agenda' && updatedXml.includes('{#agenda}')) {
        return;
      }
      
      const placeholder = `{${key}}`;
      if (updatedXml.includes(placeholder)) {
        const escapedValue = this.escapeXmlContent(String(value || ''));
        // Use a more precise replacement that preserves XML structure
        const regex = new RegExp(placeholder.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'g');
        updatedXml = updatedXml.replace(regex, escapedValue);
      }
    });

    // Handle special placeholders
    updatedXml = this.handleSpecialPlaceholders(updatedXml);

    // Update the slide content
    this.zip.file(slideFile, updatedXml);
  }

  /**
   * Escape XML special characters in content
   */
  private escapeXmlContent(content: string): string {
    if (typeof content !== 'string') {
      content = String(content || '');
    }
    return content
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  }

  /**
   * Escape XML special characters in attributes (more strict)
   */
  private escapeXmlAttribute(content: string): string {
    if (typeof content !== 'string') {
      content = String(content || '');
    }
    return content
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;')
      .replace(/\n/g, ' ')
      .replace(/\r/g, ' ')
      .replace(/\t/g, ' ');
  }

  /**
   * Handle special placeholders like date
   */
  private handleSpecialPlaceholders(xml: string): string {
    // Auto-generate common placeholders if not provided
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
        xml = xml.replace(new RegExp(`{${key}}`, 'g'), generator());
      }
    });
    
    return xml;
  }

  /**
   * Process agenda template with table format
   */
  private processAgendaTemplate(xml: string): string {
    // Check if there's an agenda template block
    const agendaTemplateRegex = /\{#agenda\}([\s\S]*?)\{\/agenda\}/g;
    const hasTemplateBlock = agendaTemplateRegex.test(xml);
    
    if (hasTemplateBlock) {
      // Process template format: {#agenda}{index} {title}{/agenda}
      return xml.replace(agendaTemplateRegex, (match, templateContent) => {
        const agendaText = this.data.agenda || '';
        const agendaItems = agendaText.split('\n').filter(item => item.trim());
        
        if (agendaItems.length === 0) {
          return ''; // Remove the entire agenda block if no items
        }

        // Process each agenda item
        const processedItems = agendaItems.map((item, index) => {
          let itemContent = templateContent;
          
          // Replace {index} with 1-based index
          itemContent = itemContent.replace(/\{index\}/g, (index + 1).toString());
          
          // Replace {title} with the agenda item text (escaped)
          const escapedTitle = this.escapeXmlContent(item.trim());
          itemContent = itemContent.replace(/\{title\}/g, escapedTitle);
          
          return itemContent;
        });

        return processedItems.join('');
      });
    } else {
      // Simple {agenda} replacement - format as numbered list
      return xml.replace(/\{agenda\}/g, (match) => {
        const agendaText = this.data.agenda || '';
        const agendaItems = agendaText.split('\n').filter(item => item.trim());
        
        if (agendaItems.length === 0) {
          return '';
        }

        // Format as numbered list with escaped content
        const numberedItems = agendaItems.map((item, index) => {
          const escapedItem = this.escapeXmlContent(item.trim());
          return `${(index + 1).toString().padStart(2, '0')}. ${escapedItem}`;
        });

        return numberedItems.join('\n');
      });
    }
  }

  /**
   * Validate basic XML structure
   */
  private isValidXmlStructure(xml: string): boolean {
    if (!xml || typeof xml !== 'string') {
      return false;
    }
    
    // Very basic validation - just check for obvious XML structure issues
    // PowerPoint slide XMLs are fragments, not full documents
    
    // Check for basic XML structure requirements
    const hasOpenTags = xml.includes('<');
    const hasCloseTags = xml.includes('>');
    
    if (!hasOpenTags || !hasCloseTags) {
      return false;
    }
    
    // Basic well-formedness check - no obvious broken tags at the end
    const brokenTags = xml.match(/<[^>]*$/g);
    if (brokenTags && brokenTags.length > 0) {
      return false;
    }
    
    // Only check for critical unescaped ampersands that would break XML parsing
    // Be more lenient - this is a common issue but PowerPoint can handle some cases
    const criticalUnescapedAmps = xml.match(/&(?![a-zA-Z#])/g);
    if (criticalUnescapedAmps && criticalUnescapedAmps.length > 0) {
      return false;
    }
    
    return true;
  }

  /**
   * Process chart data if available
   */
  async processChartData(chartData: any[]): Promise<void> {
    if (!chartData || !Array.isArray(chartData) || chartData.length === 0) {
      return;
    }

    // Group charts by title dynamically
    const chartGroups: { [key: string]: any[] } = {};
    chartData.forEach(chart => {
      // Find the title field dynamically (could be ChartTitle, title, name, etc.)
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

    // Find chart placeholder slides dynamically
    const placeholderSlides = await this.findChartPlaceholderSlides();
    
    // For each chart group, create a new slide
    for (const [chartTitle, charts] of Object.entries(chartGroups)) {
      await this.createChartSlide(chartTitle, charts, placeholderSlides);
    }
  }

  /**
   * Find all slides that contain chart-related placeholders
   */
  private async findChartPlaceholderSlides(): Promise<string[]> {
    const slideFiles = Object.keys(this.zip.files)
      .filter(filename => filename.startsWith('ppt/slides/slide') && filename.endsWith('.xml'));

    const placeholderSlides: string[] = [];
    
    for (const slideFile of slideFiles) {
      const content = await this.zip.file(slideFile)?.async('string');
      if (content) {
        // Look for any chart-related placeholders dynamically
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
   * Create a new chart slide by duplicating a placeholder slide
   */
  private async createChartSlide(chartTitle: string, charts: any[], placeholderSlides: string[]): Promise<void> {
    if (placeholderSlides.length === 0) {
      return;
    }

    // Use the first placeholder slide as template
    const placeholderSlideFile = placeholderSlides[0];
    const placeholderSlideContent = await this.zip.file(placeholderSlideFile)?.async('string');
    
    if (!placeholderSlideContent) {
      return;
    }

    // Get the next slide number
    const slideFiles = Object.keys(this.zip.files)
      .filter(filename => filename.startsWith('ppt/slides/slide') && filename.endsWith('.xml'));
    
    const slideNumbers = slideFiles.map(file => {
      const match = file.match(/slide(\d+)\.xml/);
      return match ? parseInt(match[1]) : 0;
    });
    const nextSlideNumber = Math.max(...slideNumbers) + 1;

    // Create new slide file
    const newSlideFile = `ppt/slides/slide${nextSlideNumber}.xml`;
    let newSlideContent = placeholderSlideContent;

    // Replace all placeholders dynamically
    newSlideContent = this.replaceAllPlaceholders(newSlideContent, {
      ...this.data,
      // Add chart-specific data
      [this.findChartTitlePlaceholder(placeholderSlideContent)]: chartTitle
    });

    // Process agenda template in chart slide if present
    newSlideContent = this.processAgendaTemplate(newSlideContent);

    // Handle special placeholders
    newSlideContent = this.handleSpecialPlaceholders(newSlideContent);

    // Add chart data
    const chartDataText = this.formatChartDataForSlide(charts);
    newSlideContent = this.replacePlaceholderText(newSlideContent, chartDataText);

    // Remove any remaining placeholder text
    newSlideContent = newSlideContent.replace(/{[^}]*}/g, '');

    this.zip.file(newSlideFile, newSlideContent);

    // Update presentation files
    await this.updatePresentationXml(nextSlideNumber);
    await this.updatePresentationRels(nextSlideNumber);
  }

  /**
   * Find the chart title placeholder in slide content
   */
  private findChartTitlePlaceholder(slideContent: string): string {
    const chartPlaceholders = slideContent.match(/\{[^}]*\}/g) || [];
    
    // Look for chart-related placeholders first
    for (const placeholder of chartPlaceholders) {
      const cleanName = placeholder.replace(/[{}]/g, '').toLowerCase();
      if (cleanName.includes('chart') || cleanName.includes('title')) {
        return cleanName;
      }
    }
    
    // Return the first placeholder if no chart-specific one found
    return chartPlaceholders.length > 0 ? chartPlaceholders[0].replace(/[{}]/g, '') : 'ChartTitle';
  }

  /**
   * Replace all placeholders in content with provided data
   */
  private replaceAllPlaceholders(content: string, data: Record<string, any>): string {
    let updatedContent = content;
    
    Object.entries(data).forEach(([key, value]) => {
      const placeholder = `{${key}}`;
      const escapedValue = this.escapeXmlContent(String(value || ''));
      updatedContent = updatedContent.replace(new RegExp(placeholder.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'g'), escapedValue);
    });
    
    return updatedContent;
  }

  /**
   * Safely replace text content in XML preserving structure
   */
  private safeReplaceTextContent(xml: string, searchText: string, replaceText: string): string {
    const escapedText = this.escapeXmlContent(replaceText);
    
    // Try to match text within <a:t> tags (PowerPoint text runs)
    const textRunRegex = /(<a:t[^>]*>)([^<]*?)(<\/a:t>)/g;
    
    return xml.replace(textRunRegex, (match, openTag, textContent, closeTag) => {
      if (textContent.includes(searchText)) {
        const newContent = textContent.replace(new RegExp(searchText.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'g'), escapedText);
        return openTag + newContent + closeTag;
      }
      return match;
    });
  }

  /**
   * Replace placeholder text with chart data
   */
  private replacePlaceholderText(content: string, chartDataText: string): string {
    // Look for common placeholder text patterns
    const placeholderPatterns = [
      /This is a placeholder slide\.[\s\S]*?Charts defined in[^.]*\.?/g,
      /Placeholder[^.]*chart[^.]*\./gi,
      /Chart data will be[^.]*\./gi
    ];
    
    let updatedContent = content;
    placeholderPatterns.forEach(pattern => {
      updatedContent = updatedContent.replace(pattern, chartDataText);
    });
    
    return updatedContent;
  }

  /**
   * Format chart data for display in slide
   */
  private formatChartDataForSlide(charts: any[]): string {
    if (!charts || charts.length === 0) {
      return 'No chart data available.';
    }

    // Create a formatted table representation
    const lines = ['Chart Data Summary', ''];
    
    // Group by series for better organization
    const groupedData: { [key: string]: any[] } = {};
    charts.forEach(chart => {
      const series = chart.Series || chart.series || 'Data';
      if (!groupedData[series]) {
        groupedData[series] = [];
      }
      groupedData[series].push(chart);
    });

    // Format each series with escaped content
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
   * Update presentation.xml to include new slide
   */
  private async updatePresentationXml(newSlideNumber: number): Promise<void> {
    const presFile = 'ppt/presentation.xml';
    const presXml = await this.zip.file(presFile)?.async('string');
    if (!presXml) return;

    // Find the slideIdLst section and add new slide ID
    const slideIdListMatch = presXml.match(/<p:sldIdLst>([\s\S]*?)<\/p:sldIdLst>/);
    if (!slideIdListMatch) return;

    // Get the highest existing ID
    const idMatches = presXml.match(/<p:sldId id="(\d+)"/g);
    const existingIds = idMatches ? idMatches.map(match => parseInt(match.match(/id="(\d+)"/)![1])) : [];
    const newId = existingIds.length > 0 ? Math.max(...existingIds) + 1 : 256;

    // Get the highest existing rId
    const rIdMatches = presXml.match(/r:id="rId(\d+)"/g);
    const existingRIds = rIdMatches ? rIdMatches.map(match => parseInt(match.match(/r:id="rId(\d+)"/)![1])) : [];
    const newRId = existingRIds.length > 0 ? Math.max(...existingRIds) + 1 : 2;

    // Create new slide ID entry
    const newSlideId = `        <p:sldId id="${newId}" r:id="rId${newRId}"/>`;

    // Insert the new slide ID
    const updatedSlideIdList = slideIdListMatch[1] + '\n' + newSlideId;
    const updatedPresXml = presXml.replace(
      /<p:sldIdLst>[\s\S]*?<\/p:sldIdLst>/,
      `<p:sldIdLst>${updatedSlideIdList}\n    </p:sldIdLst>`
    );

    this.zip.file(presFile, updatedPresXml);
  }

  /**
   * Update presentation.xml.rels to include new slide relationship
   */
  private async updatePresentationRels(newSlideNumber: number): Promise<void> {
    const relsFile = 'ppt/_rels/presentation.xml.rels';
    const relsXml = await this.zip.file(relsFile)?.async('string');
    if (!relsXml) return;

    // Get the highest existing rId
    const rIdMatches = relsXml.match(/Id="rId(\d+)"/g);
    const existingRIds = rIdMatches ? rIdMatches.map(match => parseInt(match.match(/Id="rId(\d+)"/)![1])) : [];
    const newRId = existingRIds.length > 0 ? Math.max(...existingRIds) + 1 : 2;

    // Create new relationship
    const newRelationship = `    <Relationship Id="rId${newRId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${newSlideNumber}.xml"/>`;

    // Insert the new relationship before the closing Relationships tag
    const updatedRelsXml = relsXml.replace(
      /<\/Relationships>/,
      `${newRelationship}\n<\/Relationships>`
    );

    this.zip.file(relsFile, updatedRelsXml);
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
      console.log('Starting PPTX build process...');
      
      // Validate input data
      if (!this.data || Object.keys(this.data).length === 0) {
        console.warn('No data provided for placeholder replacement');
      }
      
      // Process placeholders
      await this.replacePlaceholders();
      
      // Process chart data if available
      if (chartData && Array.isArray(chartData) && chartData.length > 0) {
        await this.processChartData(chartData);
      }
      
      // Final cleanup: remove any remaining placeholders
      await this.cleanupRemainingPlaceholders();
      
      // Generate the final PPTX
      return await this.generate();
      
    } catch (error) {
      console.error('Error during PPTX build:', error);
      throw new Error(`PPTX build failed: ${(error as Error).message}`);
    }
  }

  /**
   * Validate the ZIP structure before generation
   */
  private async validateZipStructure(): Promise<boolean> {
    try {
      console.log('Validating ZIP structure...');
      
      // Check essential PPTX files
      const essentialFiles = [
        '[Content_Types].xml',
        'ppt/presentation.xml',
        'ppt/_rels/presentation.xml.rels'
      ];
      
      for (const file of essentialFiles) {
        if (!this.zip.files[file]) {
          console.warn(`Missing essential file: ${file}`);
          // Don't fail validation for missing files, just warn
        } else {
          console.log(`Found essential file: ${file}`);
        }
      }
      
      // Validate slide files exist
      const slideFiles = Object.keys(this.zip.files)
        .filter(filename => filename.startsWith('ppt/slides/slide') && filename.endsWith('.xml'));
      
      if (slideFiles.length === 0) {
        console.error('No slide files found');
        return false;
      }
      
      console.log(`Found ${slideFiles.length} slide files: ${slideFiles.join(', ')}`);
      
      // Validate each slide XML structure (but don't fail if some are invalid)
      let validSlides = 0;
      for (const slideFile of slideFiles) {
        const slideContent = await this.zip.file(slideFile)?.async('string');
        if (slideContent && this.isValidXmlStructure(slideContent)) {
          validSlides++;
          console.log(`Valid slide: ${slideFile}`);
        } else {
          console.warn(`Invalid slide: ${slideFile}`);
        }
      }
      
      if (validSlides === 0) {
        console.error('No valid slides found');
        return false;
      }
      
      console.log(`ZIP validation passed: ${validSlides}/${slideFiles.length} slides are valid`);
      return true;
      
    } catch (error) {
      console.error('Error validating ZIP structure:', error);
      return false;
    }
  }

  /**
   * Clean up any remaining placeholders that weren't replaced
   */
  private async cleanupRemainingPlaceholders(): Promise<void> {
    const slideFiles = Object.keys(this.zip.files)
      .filter(filename => filename.startsWith('ppt/slides/slide') && filename.endsWith('.xml'));

    for (const slideFile of slideFiles) {
      const slideXml = await this.zip.file(slideFile)?.async('string');
      if (!slideXml) return;

      // Remove any remaining placeholder patterns dynamically
      let cleanedXml = slideXml
        .replace(/\{[^}]*\}/g, '') // Remove any remaining {placeholder}
        .replace(/\{#[^}]*\}[\s\S]*?\{\/[^}]*\}/g, ''); // Remove any remaining template blocks

      this.zip.file(slideFile, cleanedXml);
    }
  }
}