/**
 * PPTX Validator - Validates generated PPTX files
 */

export class PPTXValidator {
  /**
   * Validate PPTX file structure
   */
  static validatePPTXStructure(arrayBuffer: ArrayBuffer): { isValid: boolean; errors: string[] } {
    const errors: string[] = [];
    
    try {
      // Convert ArrayBuffer to string for basic validation
      const buffer = new Uint8Array(arrayBuffer);
      const decoder = new TextDecoder('utf-8');
      const content = decoder.decode(buffer);
      
      // Check for PPTX file signature
      if (!content.includes('Content_Types')) {
        errors.push('Missing Content_Types.xml');
      }
      
      if (!content.includes('presentation.xml')) {
        errors.push('Missing presentation.xml');
      }
      
      if (!content.includes('slide')) {
        errors.push('No slides found');
      }
      
      // Check for required XML namespaces
      if (!content.includes('http://schemas.openxmlformats.org/presentationml/2006/main')) {
        errors.push('Missing PPTX namespace');
      }
      
      if (!content.includes('http://schemas.openxmlformats.org/drawingml/2006/main')) {
        errors.push('Missing DrawingML namespace');
      }
      
      // Check for basic slide structure
      if (!content.includes('<p:sld>')) {
        errors.push('Invalid slide structure');
      }
      
      return {
        isValid: errors.length === 0,
        errors
      };
      
    } catch (error) {
      errors.push(`Validation error: ${(error as Error).message}`);
      return {
        isValid: false,
        errors
      };
    }
  }
  
  /**
   * Get PPTX file info
   */
  static getFileInfo(arrayBuffer: ArrayBuffer): { size: number; slides: number; hasCharts: boolean } {
    const buffer = new Uint8Array(arrayBuffer);
    const decoder = new TextDecoder('utf-8');
    const content = decoder.decode(buffer);
    
    // Count slides
    const slideMatches = content.match(/<p:sld>/g);
    const slideCount = slideMatches ? slideMatches.length : 0;
    
    // Check for charts
    const hasCharts = content.includes('chart') || content.includes('Chart');
    
    return {
      size: arrayBuffer.byteLength,
      slides: slideCount,
      hasCharts
    };
  }
}