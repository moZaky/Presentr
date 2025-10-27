/**
 * Real PowerPoint Template Parser
 * Extracts actual design, layout, and styling information from PowerPoint templates
 */

import JSZip from 'jszip';

export interface RealPPTXSlide {
  id: string;
  elements: RealPPTXElement[];
  background: RealPPTXBackground;
  layout: RealPPTXLayout;
  transition?: RealPPTXTransition;
}

export interface RealPPTXElement {
  id: string;
  type: 'text' | 'image' | 'shape' | 'chart' | 'table';
  content: string | any;
  position: {
    x: number;  // EMU units (English Metric Units)
    y: number;  // EMU units
  };
  size: {
    width: number;   // EMU units
    height: number;  // EMU units
  };
  style: RealPPTXStyle;
  placeholder?: RealPPTXPlaceholder;
}

export interface RealPPTXStyle {
  fontFamily?: string;
  fontSize?: number;  // in points
  fontWeight?: 'normal' | 'bold';
  fontStyle?: 'normal' | 'italic';
  textAlign?: 'left' | 'center' | 'right' | 'justify';
  color?: string;  // hex color
  backgroundColor?: string;
  borderColor?: string;
  borderWidth?: number;
  borderRadius?: number;
  opacity?: number;
  lineHeight?: number;
  letterSpacing?: number;
  textShadow?: string;
  boxShadow?: string;
  gradient?: RealPPTXGradient;
  fill?: RealPPTXFill;
}

export interface RealPPTXGradient {
  type: 'linear' | 'radial';
  colors: Array<{ color: string; position: number }>;
  angle?: number;
}

export interface RealPPTXFill {
  type: 'solid' | 'gradient' | 'image' | 'pattern';
  color?: string;
  gradient?: RealPPTXGradient;
  image?: string;
  pattern?: string;
}

export interface RealPPTXBackground {
  type: 'solid' | 'gradient' | 'image';
  color?: string;
  gradient?: RealPPTXGradient;
  image?: string;
}

export interface RealPPTXLayout {
  type: 'title' | 'content' | 'section' | 'comparison' | 'custom';
  name: string;
  placeholders: RealPPTXPlaceholder[];
}

export interface RealPPTXPlaceholder {
  id: string;
  type: 'title' | 'content' | 'image' | 'chart' | 'table';
  position: { x: number; y: number };
  size: { width: number; height: number };
  style: RealPPTXStyle;
}

export interface RealPPTXTransition {
  type: string;
  duration: number;
  direction?: string;
}

export interface RealPPTXTemplate {
  slides: RealPPTXSlide[];
  theme: RealPPTXTheme;
  masterSlides: RealPPTXSlide[];
  layoutTemplates: Record<string, RealPPTXLayout>;
}

export interface RealPPTXTheme {
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

export class RealPPTXParser {
  private template: RealPPTXTemplate | null = null;
  private zip: JSZip | null = null;

  /**
   * Parse PowerPoint template from file buffer
   */
  async parseTemplate(fileBuffer: ArrayBuffer): Promise<RealPPTXTemplate> {
    try {
      // Load the PPTX file as ZIP
      this.zip = await JSZip.loadAsync(fileBuffer);
      
      // Extract template structure
      const template = await this.extractRealTemplateStructure();
      this.template = template;
      return template;
    } catch (error) {
      console.error('Error parsing PPTX template:', error);
      throw new Error('Failed to parse PowerPoint template');
    }
  }

  /**
   * Extract real template structure from PPTX file
   */
  private async extractRealTemplateStructure(): Promise<RealPPTXTemplate> {
    if (!this.zip) {
      throw new Error('PPTX file not loaded');
    }

    // Parse presentation XML
    const presentationXml = await this.zip.file('ppt/presentation.xml')?.async('string');
    if (!presentationXml) {
      throw new Error('Invalid PPTX file: presentation.xml not found');
    }

    // Parse slide master and layout files
    const slideMasters = await this.parseSlideMasters();
    const slideLayouts = await this.parseSlideLayouts();
    const theme = await this.parseTheme();

    // Parse individual slides
    const slides = await this.parseSlides(slideMasters, slideLayouts);

    return {
      slides,
      theme,
      masterSlides: slideMasters,
      layoutTemplates: slideLayouts
    };
  }

  /**
   * Parse slide masters
   */
  private async parseSlideMasters(): Promise<RealPPTXSlide[]> {
    if (!this.zip) return [];

    const masters: RealPPTXSlide[] = [];
    const masterFiles = this.zip.file(/ppt\/slideMasters\/slideMaster\d+\.xml/);

    for (const file of masterFiles) {
      const content = await file.async('string');
      const master = this.parseSlideMasterXml(content);
      if (master) {
        masters.push(master);
      }
    }

    return masters;
  }

  /**
   * Parse slide layouts
   */
  private async parseSlideLayouts(): Promise<Record<string, RealPPTXLayout>> {
    if (!this.zip) return {};

    const layouts: Record<string, RealPPTXLayout> = {};
    const layoutFiles = this.zip.file(/ppt\/slideLayouts\/slideLayout\d+\.xml/);

    for (const file of layoutFiles) {
      const content = await file.async('string');
      const layout = this.parseSlideLayoutXml(content);
      if (layout) {
        layouts[layout.name] = layout;
      }
    }

    return layouts;
  }

  /**
   * Parse theme
   */
  private async parseTheme(): Promise<RealPPTXTheme> {
    if (!this.zip) {
      return this.getDefaultTheme();
    }

    const themeFile = this.zip.file('ppt/theme/theme1.xml');
    if (!themeFile) {
      return this.getDefaultTheme();
    }

    const content = await themeFile.async('string');
    return this.parseThemeXml(content) || this.getDefaultTheme();
  }

  /**
   * Parse slides
   */
  private async parseSlides(masters: RealPPTXSlide[], layouts: Record<string, RealPPTXLayout>): Promise<RealPPTXSlide[]> {
    if (!this.zip) return [];

    const slides: RealPPTXSlide[] = [];
    const slideFiles = this.zip.file(/ppt\/slides\/slide\d+\.xml/);

    for (let i = 0; i < slideFiles.length; i++) {
      const file = slideFiles[i];
      const content = await file.async('string');
      const slide = await this.parseSlideXml(content, i + 1, masters, layouts);
      if (slide) {
        slides.push(slide);
      }
    }

    return slides;
  }

  /**
   * Parse slide master XML
   */
  private parseSlideMasterXml(xml: string): RealPPTXSlide | null {
    try {
      const parser = new DOMParser();
      const doc = parser.parseFromString(xml, 'text/xml');
      
      // Extract background
      const background = this.extractBackground(doc);
      
      // Extract elements
      const elements = this.extractElements(doc);
      
      return {
        id: 'master',
        elements,
        background,
        layout: {
          type: 'custom',
          name: 'Master',
          placeholders: []
        }
      };
    } catch (error) {
      console.error('Error parsing slide master:', error);
      return null;
    }
  }

  /**
   * Parse slide layout XML
   */
  private parseSlideLayoutXml(xml: string): RealPPTXLayout | null {
    try {
      const parser = new DOMParser();
      const doc = parser.parseFromString(xml, 'text/xml');
      
      const layoutNode = doc.querySelector('p:slideLayout');
      if (!layoutNode) return null;

      const name = layoutNode.getAttribute('name') || 'Unknown Layout';
      const type = this.inferLayoutType(name);
      
      // Extract placeholders
      const placeholders = this.extractPlaceholders(doc);

      return {
        type,
        name,
        placeholders
      };
    } catch (error) {
      console.error('Error parsing slide layout:', error);
      return null;
    }
  }

  /**
   * Parse theme XML
   */
  private parseThemeXml(xml: string): RealPPTXTheme | null {
    try {
      const parser = new DOMParser();
      const doc = parser.parseFromString(xml, 'text/xml');
      
      // Extract color scheme
      const colorScheme = this.extractColorScheme(doc);
      
      // Extract font scheme
      const fontScheme = this.extractFontScheme(doc);

      return {
        name: 'PowerPoint Theme',
        colors: colorScheme,
        fonts: fontScheme,
        effects: {}
      };
    } catch (error) {
      console.error('Error parsing theme:', error);
      return null;
    }
  }

  /**
   * Parse slide XML
   */
  private async parseSlideXml(
    xml: string, 
    slideNumber: number, 
    masters: RealPPTXSlide[], 
    layouts: Record<string, RealPPTXLayout>
  ): Promise<RealPPTXSlide | null> {
    try {
      const parser = new DOMParser();
      const doc = parser.parseFromString(xml, 'text/xml');
      
      // Extract background (from master or slide)
      const background = this.extractBackground(doc) || 
                       (masters[0]?.background) || 
                       this.getDefaultBackground();
      
      // Extract elements
      const elements = this.extractElements(doc);
      
      // Determine layout
      const layoutName = this.getSlideLayoutName(doc);
      const layout = layouts[layoutName] || this.getDefaultLayout();

      return {
        id: `slide${slideNumber}`,
        elements,
        background,
        layout
      };
    } catch (error) {
      console.error(`Error parsing slide ${slideNumber}:`, error);
      return null;
    }
  }

  /**
   * Extract background from XML
   */
  private extractBackground(doc: Document): RealPPTXBackground | null {
    const bgNode = doc.querySelector('p:bg, p:bgPr');
    if (!bgNode) return null;

    // Check for solid fill
    const solidFill = bgNode.querySelector('a:solidFill a:srgbClr');
    if (solidFill) {
      const color = '#' + solidFill.getAttribute('val');
      return { type: 'solid', color };
    }

    // Check for gradient fill
    const gradientFill = bgNode.querySelector('a:gradFill');
    if (gradientFill) {
      const colors = this.extractGradientColors(gradientFill);
      return { type: 'gradient', gradient: colors };
    }

    return null;
  }

  /**
   * Extract elements from XML
   */
  private extractElements(doc: Document): RealPPTXElement[] {
    const elements: RealPPTXElement[] = [];
    
    // Extract text shapes
    const textShapes = doc.querySelectorAll('p:sp');
    textShapes.forEach((shape, index) => {
      const element = this.extractTextElement(shape, index);
      if (element) elements.push(element);
    });

    // Extract images
    const images = doc.querySelectorAll('p:pic');
    images.forEach((image, index) => {
      const element = this.extractImageElement(image, index);
      if (element) elements.push(element);
    });

    // Extract tables
    const tables = doc.querySelectorAll('p:tbl');
    tables.forEach((table, index) => {
      const element = this.extractTableElement(table, index);
      if (element) elements.push(element);
    });

    return elements;
  }

  /**
   * Extract text element
   */
  private extractTextElement(shape: Element, index: number): RealPPTXElement | null {
    const textNode = shape.querySelector('a:t');
    if (!textNode) return null;

    const content = textNode.textContent || '';
    const position = this.extractPosition(shape);
    const size = this.extractSize(shape);
    const style = this.extractTextStyle(shape);

    return {
      id: `text_${index}`,
      type: 'text',
      content,
      position,
      size,
      style
    };
  }

  /**
   * Extract image element
   */
  private extractImageElement(image: Element, index: number): RealPPTXElement | null {
    const position = this.extractPosition(image);
    const size = this.extractSize(image);
    
    return {
      id: `image_${index}`,
      type: 'image',
      content: 'image_placeholder',
      position,
      size,
      style: {}
    };
  }

  /**
   * Extract table element
   */
  private extractTableElement(table: Element, index: number): RealPPTXElement | null {
    const position = this.extractPosition(table);
    const size = this.extractSize(table);
    
    // Extract table data
    const rows: string[][] = [];
    const rowNodes = table.querySelectorAll('a:tr');
    
    rowNodes.forEach(row => {
      const cells: string[] = [];
      const cellNodes = row.querySelectorAll('a:tc');
      cellNodes.forEach(cell => {
        const text = cell.querySelector('a:t')?.textContent || '';
        cells.push(text);
      });
      rows.push(cells);
    });

    return {
      id: `table_${index}`,
      type: 'table',
      content: {
        headers: rows[0] || [],
        rows: rows.slice(1),
        style: {
          headerBackground: '#2E74B5',
          headerColor: '#FFFFFF',
          borderColor: '#D9D9D9',
          rowBackground: '#FFFFFF',
          alternateRowBackground: '#F2F2F2',
          fontSize: 18,
          fontFamily: 'Calibri'
        }
      },
      position,
      size,
      style: {}
    };
  }

  /**
   * Extract position from element
   */
  private extractPosition(element: Element): { x: number; y: number } {
    const xform = element.querySelector('a:xfrm');
    if (!xform) return { x: 0, y: 0 };

    const off = xform.querySelector('a:off');
    if (!off) return { x: 0, y: 0 };

    const x = parseInt(off.getAttribute('x') || '0');
    const y = parseInt(off.getAttribute('y') || '0');

    return { x, y };
  }

  /**
   * Extract size from element
   */
  private extractSize(element: Element): { width: number; height: number } {
    const xform = element.querySelector('a:xfrm');
    if (!xform) return { width: 0, height: 0 };

    const ext = xform.querySelector('a:ext');
    if (!ext) return { width: 0, height: 0 };

    const width = parseInt(ext.getAttribute('cx') || '0');
    const height = parseInt(ext.getAttribute('cy') || '0');

    return { width, height };
  }

  /**
   * Extract text style
   */
  private extractTextStyle(element: Element): RealPPTXStyle {
    const style: RealPPTXStyle = {};

    // Extract font properties
    const font = element.querySelector('a:rPr');
    if (font) {
      const fontSize = font.getAttribute('sz');
      if (fontSize) style.fontSize = parseInt(fontSize) / 100; // Convert from hundredths of a point

      const bold = font.getAttribute('b');
      if (bold) style.fontWeight = 'bold';

      const italic = font.getAttribute('i');
      if (italic) style.fontStyle = 'italic';

      const fontFamily = font.querySelector('a:latin')?.getAttribute('typeface');
      if (fontFamily) style.fontFamily = fontFamily;

      const color = font.querySelector('a:srgbClr')?.getAttribute('val');
      if (color) style.color = '#' + color;
    }

    // Extract alignment
    const bodyPr = element.querySelector('a:bodyPr');
    if (bodyPr) {
      const align = bodyPr.getAttribute('algn');
      if (align) {
        style.textAlign = align as 'left' | 'center' | 'right';
      }
    }

    return style;
  }

  /**
   * Extract gradient colors
   */
  private extractGradientColors(gradientFill: Element): RealPPTXGradient {
    const colors: Array<{ color: string; position: number }> = [];
    const stops = gradientFill.querySelectorAll('a:gsLst a:gs');

    stops.forEach(stop => {
      const color = stop.querySelector('a:srgbClr')?.getAttribute('val');
      const position = parseInt(stop.getAttribute('pos') || '0');
      if (color) {
        colors.push({ color: '#' + color, position: position / 1000 });
      }
    });

    return {
      type: 'linear',
      colors,
      angle: 90
    };
  }

  /**
   * Extract color scheme
   */
  private extractColorScheme(doc: Document): RealPPTXTheme['colors'] {
    const defaultColors = {
      primary: '#2E74B5',
      secondary: '#595959',
      accent1: '#FFC000',
      accent2: '#ED7D31',
      accent3: '#A5A5A5',
      background: '#FFFFFF',
      text: '#000000'
    };

    // Try to extract actual colors from theme
    const colorScheme = doc.querySelector('a:clrScheme');
    if (!colorScheme) return defaultColors;

    // Map PowerPoint color names to our theme colors
    const colorMap: Record<string, keyof RealPPTXTheme['colors']> = {
      'a:dk1': 'primary',
      'a:lt1': 'background',
      'a:dk2': 'secondary',
      'a:lt2': 'text',
      'a:accent1': 'accent1',
      'a:accent2': 'accent2',
      'a:accent3': 'accent3'
    };

    const colors = { ...defaultColors };

    Object.entries(colorMap).forEach(([selector, key]) => {
      const colorNode = colorScheme.querySelector(selector + ' a:srgbClr');
      if (colorNode) {
        const color = colorNode.getAttribute('val');
        if (color) {
          colors[key] = '#' + color;
        }
      }
    });

    return colors;
  }

  /**
   * Extract font scheme
   */
  private extractFontScheme(doc: Document): RealPPTXTheme['fonts'] {
    const defaultFonts = {
      heading: { latin: 'Calibri' },
      body: { latin: 'Calibri' }
    };

    const fontScheme = doc.querySelector('a:fontScheme');
    if (!fontScheme) return defaultFonts;

    const fonts = { ...defaultFonts };

    // Extract heading font
    const headingFont = fontScheme.querySelector('a:majorFont a:latin')?.getAttribute('typeface');
    if (headingFont) {
      fonts.heading.latin = headingFont;
    }

    // Extract body font
    const bodyFont = fontScheme.querySelector('a:minorFont a:latin')?.getAttribute('typeface');
    if (bodyFont) {
      fonts.body.latin = bodyFont;
    }

    return fonts;
  }

  /**
   * Extract placeholders
   */
  private extractPlaceholders(doc: Document): RealPPTXPlaceholder[] {
    const placeholders: RealPPTXPlaceholder[] = [];
    const placeholderNodes = doc.querySelectorAll('p:spPr');

    placeholderNodes.forEach((node, index) => {
      const type = this.inferPlaceholderType(node);
      const position = this.extractPosition(node);
      const size = this.extractSize(node);
      const style = this.extractTextStyle(node);

      placeholders.push({
        id: `placeholder_${index}`,
        type,
        position: { x: position.x / 914400, y: position.y / 914400 }, // Convert EMU to percentage
        size: { width: size.width / 914400, height: size.height / 914400 },
        style
      });
    });

    return placeholders;
  }

  /**
   * Get slide layout name
   */
  private getSlideLayoutName(doc: Document): string {
    const layoutNode = doc.querySelector('p:sldLayoutId');
    return layoutNode?.getAttribute('r:id') || 'Unknown';
  }

  /**
   * Infer layout type from name
   */
  private inferLayoutType(name: string): RealPPTXLayout['type'] {
    const lowerName = name.toLowerCase();
    if (lowerName.includes('title')) return 'title';
    if (lowerName.includes('content')) return 'content';
    if (lowerName.includes('section')) return 'section';
    if (lowerName.includes('comparison')) return 'comparison';
    return 'custom';
  }

  /**
   * Infer placeholder type
   */
  private inferPlaceholderType(node: Element): RealPPTXPlaceholder['type'] {
    const type = node.getAttribute('type');
    switch (type) {
      case 'title': return 'title';
      case 'body': return 'content';
      case 'pic': return 'image';
      case 'chart': return 'chart';
      case 'tbl': return 'table';
      default: return 'content';
    }
  }

  /**
   * Get default theme
   */
  private getDefaultTheme(): RealPPTXTheme {
    return {
      name: 'Default Theme',
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
   * Get default layout
   */
  private getDefaultLayout(): RealPPTXLayout {
    return {
      type: 'content',
      name: 'Title and Content',
      placeholders: []
    };
  }

  /**
   * Get default background
   */
  private getDefaultBackground(): RealPPTXBackground {
    return {
      type: 'solid',
      color: '#FFFFFF'
    };
  }

  /**
   * Get parsed template
   */
  getTemplate(): RealPPTXTemplate | null {
    return this.template;
  }

  /**
   * Get slide by ID
   */
  getSlide(slideId: string): RealPPTXSlide | null {
    if (!this.template) return null;
    return this.template.slides.find(slide => slide.id === slideId) || null;
  }

  /**
   * Get all slides
   */
  getSlides(): RealPPTXSlide[] {
    return this.template?.slides || [];
  }

  /**
   * Apply data to template
   */
  applyData(data: Record<string, any>): RealPPTXSlide[] {
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