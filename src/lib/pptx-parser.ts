/**
 * PowerPoint Template Parser
 * Extracts design, layout, and styling information from PowerPoint templates
 */

export interface PPTXSlide {
  id: string;
  elements: PPTXElement[];
  background: PPTXBackground;
  layout: PPTXLayout;
  transition?: PPTXTransition;
}

export interface PPTXElement {
  id: string;
  type: 'text' | 'image' | 'shape' | 'chart' | 'table';
  content: string | any;
  position: {
    x: number;  // Percentage (0-100)
    y: number;  // Percentage (0-100)
  };
  size: {
    width: number;   // Percentage (0-100)
    height: number;  // Percentage (0-100)
  };
  style: PPTXStyle;
  placeholder?: PPTXPlaceholder;
}

export interface PPTXStyle {
  fontFamily?: string;
  fontSize?: number;
  fontWeight?: 'normal' | 'bold';
  fontStyle?: 'normal' | 'italic';
  textAlign?: 'left' | 'center' | 'right' | 'justify';
  color?: string;
  backgroundColor?: string;
  borderColor?: string;
  borderWidth?: number;
  borderRadius?: number;
  opacity?: number;
  lineHeight?: number;
  letterSpacing?: number;
  textShadow?: string;
  boxShadow?: string;
  gradient?: PPTXGradient;
  fill?: PPTXFill;
}

export interface PPTXGradient {
  type: 'linear' | 'radial';
  colors: Array<{ color: string; position: number }>;
  angle?: number;
}

export interface PPTXFill {
  type: 'solid' | 'gradient' | 'image' | 'pattern';
  color?: string;
  gradient?: PPTXGradient;
  image?: string;
  pattern?: string;
}

export interface PPTXBackground {
  type: 'solid' | 'gradient' | 'image';
  color?: string;
  gradient?: PPTXGradient;
  image?: string;
}

export interface PPTXLayout {
  type: 'title' | 'content' | 'section' | 'comparison' | 'custom';
  name: string;
  placeholders: PPTXPlaceholder[];
}

export interface PPTXPlaceholder {
  id: string;
  type: 'title' | 'content' | 'image' | 'chart' | 'table';
  position: { x: number; y: number };
  size: { width: number; height: number };
  style: PPTXStyle;
}

export interface PPTXTransition {
  type: string;
  duration: number;
  direction?: string;
}

export interface PPTXTemplate {
  slides: PPTXSlide[];
  theme: PPTXTheme;
  masterSlides: PPTXSlide[];
  layoutTemplates: Record<string, PPTXLayout>;
}

export interface PPTXTheme {
  name: string;
  colors: {
    primary: string;
    secondary: string;
    accent1: string;
    accent2: string;
    accent3: string;
    background: string;
    text: string;
  };
  fonts: {
    heading: { latin: string; asian?: string; complex?: string };
    body: { latin: string; asian?: string; complex?: string };
  };
  effects: Record<string, any>;
}

export class PPTXParser {
  private template: PPTXTemplate | null = null;

  /**
   * Parse PowerPoint template from file buffer
   */
  async parseTemplate(fileBuffer: ArrayBuffer): Promise<PPTXTemplate> {
    try {
      // For now, create a mock template structure
      // In a real implementation, this would parse the actual PPTX XML structure
      const template = await this.extractTemplateStructure(fileBuffer);
      this.template = template;
      return template;
    } catch (error) {
      console.error('Error parsing PPTX template:', error);
      throw new Error('Failed to parse PowerPoint template');
    }
  }

  /**
   * Extract template structure from PPTX file
   */
  private async extractTemplateStructure(fileBuffer: ArrayBuffer): Promise<PPTXTemplate> {
    // Mock implementation - in reality, this would:
    // 1. Unzip the PPTX file
    // 2. Parse slide XML files
    // 3. Extract theme, layout, and styling information
    // 4. Parse placeholder positions and styles
    
    return {
      slides: this.generateMockSlides(),
      theme: this.generateMockTheme(),
      masterSlides: [],
      layoutTemplates: this.generateMockLayouts()
    };
  }

  /**
   * Generate mock slides for demonstration
   */
  private generateMockSlides(): PPTXSlide[] {
    return [
      {
        id: 'slide1',
        elements: [
          {
            id: 'title1',
            type: 'text',
            content: '{project_title}',
            position: { x: 10, y: 10 },
            size: { width: 80, height: 15 },
            style: {
              fontFamily: 'Calibri',
              fontSize: 44,
              fontWeight: 'bold',
              textAlign: 'center',
              color: '#2E74B5',
              textShadow: '0px 0px 0px rgba(0,0,0,0)'
            },
            placeholder: {
              id: 'title1',
              type: 'title',
              position: { x: 10, y: 10 },
              size: { width: 80, height: 15 },
              style: {}
            }
          },
          {
            id: 'subtitle1',
            type: 'text',
            content: 'Status Report prepared by: {name}',
            position: { x: 10, y: 30 },
            size: { width: 80, height: 10 },
            style: {
              fontFamily: 'Calibri',
              fontSize: 20,
              fontWeight: 'normal',
              textAlign: 'center',
              color: '#595959'
            },
            placeholder: {
              id: 'subtitle1',
              type: 'content',
              position: { x: 10, y: 30 },
              size: { width: 80, height: 10 },
              style: {}
            }
          }
        ],
        background: {
          type: 'gradient',
          gradient: {
            type: 'linear',
            colors: [
              { color: '#FFFFFF', position: 0 },
              { color: '#F2F2F2', position: 100 }
            ],
            angle: 90
          }
        },
        layout: {
          type: 'title',
          name: 'Title Slide',
          placeholders: []
        }
      },
      {
        id: 'slide2',
        elements: [
          {
            id: 'title2',
            type: 'text',
            content: 'Overview',
            position: { x: 5, y: 8 },
            size: { width: 90, height: 12 },
            style: {
              fontFamily: 'Calibri',
              fontSize: 36,
              fontWeight: 'bold',
              textAlign: 'left',
              color: '#2E74B5'
            },
            placeholder: {
              id: 'title2',
              type: 'title',
              position: { x: 5, y: 8 },
              size: { width: 90, height: 12 },
              style: {}
            }
          },
          {
            id: 'content2',
            type: 'text',
            content: '{status_update}',
            position: { x: 5, y: 25 },
            size: { width: 90, height: 60 },
            style: {
              fontFamily: 'Calibri',
              fontSize: 24,
              fontWeight: 'normal',
              textAlign: 'left',
              color: '#000000',
              lineHeight: 1.5
            },
            placeholder: {
              id: 'content2',
              type: 'content',
              position: { x: 5, y: 25 },
              size: { width: 90, height: 60 },
              style: {}
            }
          }
        ],
        background: {
          type: 'solid',
          color: '#FFFFFF'
        },
        layout: {
          type: 'content',
          name: 'Title and Content',
          placeholders: []
        }
      },
      {
        id: 'slide3',
        elements: [
          {
            id: 'title3',
            type: 'text',
            content: 'Agenda',
            position: { x: 5, y: 8 },
            size: { width: 90, height: 12 },
            style: {
              fontFamily: 'Calibri',
              fontSize: 36,
              fontWeight: 'bold',
              textAlign: 'left',
              color: '#2E74B5'
            },
            placeholder: {
              id: 'title3',
              type: 'title',
              position: { x: 5, y: 8 },
              size: { width: 90, height: 12 },
              style: {}
            }
          },
          {
            id: 'agenda3',
            type: 'table',
            content: {
              headers: ['Topic'],
              rows: [['{agenda}']],
              style: {
                headerBackground: '#2E74B5',
                headerColor: '#FFFFFF',
                borderColor: '#D9D9D9',
                rowBackground: '#FFFFFF',
                alternateRowBackground: '#F2F2F2',
                fontSize: 20,
                fontFamily: 'Calibri'
              }
            },
            position: { x: 5, y: 25 },
            size: { width: 90, height: 60 },
            style: {
              fontFamily: 'Calibri',
              fontSize: 20,
              color: '#000000'
            },
            placeholder: {
              id: 'agenda3',
              type: 'table',
              position: { x: 5, y: 25 },
              size: { width: 90, height: 60 },
              style: {}
            }
          }
        ],
        background: {
          type: 'solid',
          color: '#FFFFFF'
        },
        layout: {
          type: 'content',
          name: 'Title and Content',
          placeholders: []
        }
      },
      {
        id: 'slide4',
        elements: [
          {
            id: 'title4',
            type: 'text',
            content: '{ChartTitle}',
            position: { x: 5, y: 8 },
            size: { width: 90, height: 12 },
            style: {
              fontFamily: 'Calibri',
              fontSize: 36,
              fontWeight: 'bold',
              textAlign: 'left',
              color: '#2E74B5'
            },
            placeholder: {
              id: 'title4',
              type: 'title',
              position: { x: 5, y: 8 },
              size: { width: 90, height: 12 },
              style: {}
            }
          },
          {
            id: 'chart4',
            type: 'chart',
            content: {
              type: 'placeholder',
              title: 'Chart will be generated from Excel data',
              style: {
                backgroundColor: '#F2F2F2',
                borderColor: '#D9D9D9',
                borderWidth: 1
              }
            },
            position: { x: 5, y: 25 },
            size: { width: 90, height: 60 },
            style: {
              backgroundColor: '#F2F2F2',
              borderColor: '#D9D9D9',
              borderWidth: 1
            },
            placeholder: {
              id: 'chart4',
              type: 'chart',
              position: { x: 5, y: 25 },
              size: { width: 90, height: 60 },
              style: {}
            }
          }
        ],
        background: {
          type: 'solid',
          color: '#FFFFFF'
        },
        layout: {
          type: 'content',
          name: 'Title and Content',
          placeholders: []
        }
      }
    ];
  }

  /**
   * Generate mock theme
   */
  private generateMockTheme(): PPTXTheme {
    return {
      name: 'Office Theme',
      colors: {
        primary: '#2E74B5',
        secondary: '#595959',
        accent1: '#FFC000',
        accent2: '#ED7D31',
        accent3: '#A5A5A5',
        background: '#FFFFFF',
        text: '#000000'
      },
      fonts: {
        heading: { latin: 'Calibri' },
        body: { latin: 'Calibri' }
      },
      effects: {}
    };
  }

  /**
   * Generate mock layouts
   */
  private generateMockLayouts(): Record<string, PPTXLayout> {
    return {
      'Title Slide': {
        type: 'title',
        name: 'Title Slide',
        placeholders: []
      },
      'Title and Content': {
        type: 'content',
        name: 'Title and Content',
        placeholders: []
      }
    };
  }

  /**
   * Get parsed template
   */
  getTemplate(): PPTXTemplate | null {
    return this.template;
  }

  /**
   * Get slide by ID
   */
  getSlide(slideId: string): PPTXSlide | null {
    if (!this.template) return null;
    return this.template.slides.find(slide => slide.id === slideId) || null;
  }

  /**
   * Get all slides
   */
  getSlides(): PPTXSlide[] {
    return this.template?.slides || [];
  }

  /**
   * Apply data to template
   */
  applyData(data: Record<string, any>): PPTXSlide[] {
    if (!this.template) return [];

    return this.template.slides.map(slide => ({
      ...slide,
      elements: slide.elements.map(element => ({
        ...element,
        content: this.processElementContent(element.content, data)
      }))
    }));
  }

  /**
   * Process element content with data
   */
  private processElementContent(content: any, data: Record<string, any>): any {
    if (typeof content === 'string') {
      // Replace placeholders with actual data
      return content.replace(/\{([^}]+)\}/g, (match, key) => {
        return data[key] !== undefined ? String(data[key]) : match;
      });
    }

    if (content && typeof content === 'object') {
      // Handle table content
      if (content.rows) {
        return {
          ...content,
          rows: content.rows.map((row: string[]) => 
            row.map(cell => 
              typeof cell === 'string' 
                ? cell.replace(/\{([^}]+)\}/g, (match, key) => 
                    data[key] !== undefined ? String(data[key]) : match
                  )
                : cell
            )
          )
        };
      }
    }

    return content;
  }
}