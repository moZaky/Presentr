/**
 * PPTX Diagnostic Tool - Helps analyze PPTX file structure and identify issues
 */

import JSZip from 'jszip';

export interface PPTXDiagnosticResult {
  isValid: boolean;
  slideCount: number;
  hasContentTypes: boolean;
  hasPresentation: boolean;
  hasSlideRelations: boolean;
  slideFiles: string[];
  issues: string[];
  warnings: string[];
}

export class PPTXDiagnostic {
  /**
   * Analyze a PPTX file and return diagnostic information
   */
  static async analyze(buffer: ArrayBuffer): Promise<PPTXDiagnosticResult> {
    const diagnostic: PPTXDiagnosticResult = {
      isValid: true,
      slideCount: 0,
      hasContentTypes: false,
      hasPresentation: false,
      hasSlideRelations: false,
      slideFiles: [],
      issues: [],
      warnings: []
    };

    try {
      const zip = await JSZip.loadAsync(buffer);
      
      // Check for required files
      const contentTypesFile = zip.file('[Content_Types].xml');
      if (!contentTypesFile) {
        diagnostic.issues.push('Missing [Content_Types].xml');
        diagnostic.isValid = false;
      } else {
        diagnostic.hasContentTypes = true;
      }

      const presentationFile = zip.file('ppt/presentation.xml');
      if (!presentationFile) {
        diagnostic.issues.push('Missing ppt/presentation.xml');
        diagnostic.isValid = false;
      } else {
        diagnostic.hasPresentation = true;
      }

      // Check slide files
      const slideFiles = zip.file(/ppt\/slides\/slide\d+\.xml/);
      diagnostic.slideFiles = slideFiles.map(f => f.name);
      diagnostic.slideCount = slideFiles.length;

      if (slideFiles.length === 0) {
        diagnostic.issues.push('No slide files found');
        diagnostic.isValid = false;
      }

      // Check slide relationship files
      const slideRelsFiles = zip.file(/ppt\/slides\/_rels\/slide\d+\.xml\.rels/);
      diagnostic.hasSlideRelations = slideRelsFiles.length > 0;

      if (slideFiles.length > 0 && slideRelsFiles.length === 0) {
        diagnostic.warnings.push('No slide relationship files found');
      }

      // Analyze slide content for common issues
      for (const slideFile of slideFiles) {
        const content = await slideFile.async('string');
        
        // Check for XML well-formedness
        if (!content.startsWith('<?xml')) {
          diagnostic.warnings.push(`Slide ${slideFile.name} doesn't start with XML declaration`);
        }

        // Check for unescaped characters
        const unescapedPatterns = [
          /&(?!(amp|lt|gt|quot|apos);)/g,
          /<(?!\/?[a-z])/gi,
          /[^"'&<>](?=>)/g
        ];

        unescapedPatterns.forEach((pattern, index) => {
          const matches = content.match(pattern);
          if (matches) {
            diagnostic.warnings.push(`Slide ${slideFile.name} has potentially unescaped characters (pattern ${index + 1})`);
          }
        });

        // Check for placeholder patterns
        const placeholderMatches = content.match(/\{[^}]+\}/g);
        if (placeholderMatches) {
          diagnostic.warnings.push(`Slide ${slideFile.name} contains unreplaced placeholders: ${placeholderMatches.join(', ')}`);
        }
      }

    } catch (error) {
      diagnostic.isValid = false;
      diagnostic.issues.push(`Failed to analyze PPTX: ${(error as Error).message}`);
    }

    return diagnostic;
  }

  /**
   * Generate a human-readable report
   */
  static generateReport(diagnostic: PPTXDiagnosticResult): string {
    let report = '=== PPTX Diagnostic Report ===\n';
    report += `Valid: ${diagnostic.isValid ? '✅' : '❌'}\n`;
    report += `Slide Count: ${diagnostic.slideCount}\n`;
    report += `Has Content Types: ${diagnostic.hasContentTypes ? '✅' : '❌'}\n`;
    report += `Has Presentation: ${diagnostic.hasPresentation ? '✅' : '❌'}\n`;
    report += `Has Slide Relations: ${diagnostic.hasSlideRelations ? '✅' : '❌'}\n`;
    
    if (diagnostic.slideFiles.length > 0) {
      report += '\nSlide Files:\n';
      diagnostic.slideFiles.forEach(file => {
        report += `  - ${file}\n`;
      });
    }

    if (diagnostic.issues.length > 0) {
      report += '\n🚨 Issues:\n';
      diagnostic.issues.forEach(issue => {
        report += `  - ${issue}\n`;
      });
    }

    if (diagnostic.warnings.length > 0) {
      report += '\n⚠️ Warnings:\n';
      diagnostic.warnings.forEach(warning => {
        report += `  - ${warning}\n`;
      });
    }

    return report;
  }
}