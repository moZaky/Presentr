/**
 * Ultra-Safe PPTX Builder - Fallback implementation with maximum compatibility
 * This is the most conservative approach possible
 */

import JSZip from 'jszip';

export interface PPTXData {
  [key: string]: any;
}

export class UltraSafePPTXBuilder {
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
   * Build the final PPTX with data applied - ultra conservative approach
   */
  async build(): Promise<ArrayBuffer> {
    // Only process text content, leave everything else untouched
    const slideFiles = this.zip.file(/ppt\/slides\/slide\d+\.xml/);
    
    for (const file of slideFiles) {
      const content = await file.async('string');
      const processedContent = this.processSlideContentUltraSafe(content);
      this.zip.file(file.name, processedContent);
    }

    // Generate the buffer
    return await this.zip.generateAsync({ type: 'arraybuffer' });
  }

  /**
   * Build with charts - simplified approach
   */
  async buildWithCharts(chartData: any[]): Promise<ArrayBuffer> {
    // First process existing slides
    const result = await this.build();
    
    // If we have chart data, add it as text-only slides
    if (chartData && chartData.length > 0) {
      // For now, just return the basic processed slides
      // Chart slides can be added in a future iteration
      console.log('Chart data detected but not processed in ultra-safe mode');
    }
    
    return result;
  }

  /**
   * Ultra-safe slide content processing
   * Only replaces text in <a:t> tags, leaves all other XML structure intact
   */
  private processSlideContentUltraSafe(xmlContent: string): string {
    // Find all text content in <a:t> tags and replace placeholders
    // This is the most conservative approach possible
    
    let processedContent = xmlContent;
    
    // Process each data item
    for (const [key, value] of Object.entries(this.data)) {
      const placeholder = `{${key}}`;
      if (processedContent.includes(placeholder)) {
        const safeValue = this.makeTextSafe(String(value || ''));
        // Only replace within <a:t> tags to avoid breaking XML structure
        processedContent = this.replaceInTextTags(processedContent, placeholder, safeValue);
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
        const safeValue = this.makeTextSafe(String(value));
        processedContent = this.replaceInTextTags(processedContent, placeholderPattern, safeValue);
      }
    }

    return processedContent;
  }

  /**
   * Replace text only within <a:t> tags
   */
  private replaceInTextTags(content: string, placeholder: string, replacement: string): string {
    // Use a regex to find <a:t>...</a:t> tags and replace within them
    return content.replace(/<a:t[^>]*>([^<]*)<\/a:t>/g, (match, textContent) => {
      if (textContent.includes(placeholder)) {
        const newContent = textContent.split(placeholder).join(replacement);
        return match.replace(textContent, newContent);
      }
      return match;
    });
  }

  /**
   * Make text safe for XML - minimal escaping
   */
  private makeTextSafe(text: string): string {
    // Only escape the most critical characters that would break XML
    return text
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;');
  }
}