/**
 * PowerPoint Template Generator
 * Generates PPTX files while preserving original template design
 */

import JSZip from 'jszip';
import { RealPPTXTemplate, RealPPTXSlide, RealPPTXElement } from './real-pptx-parser';

export interface PPTXGenerationOptions {
  template: RealPPTXTemplate;
  data: Record<string, any>;
  outputPath?: string;
}

export class PPTXGenerator {
  private zip: JSZip;
  private template: RealPPTXTemplate;
  private data: Record<string, any>;

  constructor(options: PPTXGenerationOptions) {
    this.zip = new JSZip();
    this.template = options.template;
    this.data = options.data;
  }

  /**
   * Generate PPTX file with preserved template design
   */
  async generate(): Promise<Blob> {
    try {
      // Create PPTX structure
      await this.createPPTXStructure();
      
      // Generate content types
      await this.generateContentTypes();
      
      // Generate relationships
      await this.generateRelationships();
      
      // Generate presentation
      await this.generatePresentation();
      
      // Generate slides with applied data
      await this.generateSlides();
      
      // Generate theme
      await this.generateTheme();
      
      // Generate slide master and layout
      await this.generateSlideMaster();
      await this.generateSlideLayouts();
      
      // Generate the blob
      return await this.zip.generateAsync({ type: 'blob' });
      
    } catch (error) {
      console.error('Error generating PPTX:', error);
      throw new Error('Failed to generate PPTX file');
    }
  }

  /**
   * Create basic PPTX folder structure
   */
  private async createPPTXStructure(): Promise<void> {
    // Create folders
    this.zip.folder('ppt');
    this.zip.folder('ppt/slides');
    this.zip.folder('ppt/slideMasters');
    this.zip.folder('ppt/slideLayouts');
    this.zip.folder('ppt/theme');
    this.zip.folder('ppt/media');
    this.zip.folder('docProps');
    
    // Create [Content_Types].xml
    const contentTypes = this.generateContentTypesXML();
    this.zip.file('[Content_Types].xml', contentTypes);
    
    // Create _rels folder and .rels
    this.zip.folder('_rels');
    const rels = this.generateRootRels();
    this.zip.file('_rels/.rels', rels);
  }

  /**
   * Generate content types XML
   */
  private generateContentTypesXML(): string {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml" ContentType="application/xml"/>
    <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-presentationml.presentation.main+xml"/>
    <Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-presentationml.slideMaster+xml"/>
    <Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-presentationml.slideLayout+xml"/>
    <Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-presentationml.theme+xml"/>
    ${this.template.slides.map((_, index) => 
      `<Override PartName="/ppt/slides/slide${index + 1}.xml" ContentType="application/vnd.openxmlformats-presentationml.slide+xml"/>`
    ).join('\n    ')}
</Types>`;
  }

  /**
   * Generate root relationships
   */
  private generateRootRels(): string {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
    <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
    <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>`;
  }

  /**
   * Generate presentation XML
   */
  private async generatePresentation(): Promise<void> {
    const slideIds = this.template.slides.map((_, index) => 
      `<p:sldId id="${256 + index}" r:id="rId${2 + index}"/>`
    ).join('\n        ');

    const presentationXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
    <p:sldMasterIdLst>
        <p:sldMasterId id="2147483648" r:id="rId1"/>
    </p:sldMasterIdLst>
    <p:sldIdLst>
        ${slideIds}
    </p:sldIdLst>
    <p:sldSz cx="9144000" cy="5143500"/>
    <p:notesSz cx="6858000" cy="9144000"/>
    <p:defaultTextStyle>
        <a:defPPr>
            <a:defRPr lang="en-US"/>
        </a:defPPr>
    </p:defaultTextStyle>
</p:presentation>`;

    this.zip.file('ppt/presentation.xml', presentationXml);
    
    // Generate presentation relationships
    const slideRels = this.template.slides.map((_, index) => 
      `<Relationship Id="rId${2 + index}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${index + 1}.xml"/>`
    ).join('\n    ');

    const presentationRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>
    ${slideRels}
</Relationships>`;

    this.zip.folder('ppt/_rels');
    this.zip.file('ppt/_rels/presentation.xml.rels', presentationRels);
  }

  /**
   * Generate slides with applied data
   */
  private async generateSlides(): Promise<void> {
    const appliedSlides = this.applyDataToTemplate();
    
    for (let i = 0; i < appliedSlides.length; i++) {
      const slide = appliedSlides[i];
      const slideXml = this.generateSlideXML(slide, i + 1);
      this.zip.file(`ppt/slides/slide${i + 1}.xml`, slideXml);
      
      // Generate slide relationships
      const slideRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
</Relationships>`;
      
      this.zip.folder(`ppt/slides/_rels`);
      this.zip.file(`ppt/slides/_rels/slide${i + 1}.xml.rels`, slideRels);
    }
  }

  /**
   * Generate individual slide XML
   */
  private generateSlideXML(slide: RealPPTXSlide, slideNumber: number): string {
    const elements = slide.elements.map(element => this.generateElementXML(element)).join('\n        ');
    
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
            ${elements}
        </p:spTree>
    </p:cSld>
    <p:clrMapOvr>
        <a:masterClrMapping/>
    </p:clrMapOvr>
</p:sld>`;
  }

  /**
   * Generate element XML
   */
  private generateElementXML(element: RealPPTXElement): string {
    if (element.type === 'text') {
      return this.generateTextElementXML(element);
    } else if (element.type === 'table') {
      return this.generateTableElementXML(element);
    } else if (element.type === 'chart') {
      return this.generateChartElementXML(element);
    }
    return '';
  }

  /**
   * Generate text element XML
   */
  private generateTextElementXML(element: RealPPTXElement): string {
    const content = typeof element.content === 'string' ? element.content : '';
    const fontSize = element.style.fontSize || 18;
    const color = element.style.color || '#000000';
    const fontFamily = element.style.fontFamily || 'Calibri';
    const bold = element.style.fontWeight === 'bold';
    const italic = element.style.fontStyle === 'italic';
    const align = element.style.textAlign || 'left';

    return `<p:sp>
    <p:nvSpPr>
        <p:cNvPr id="${this.getRandomId()}" name="Text Shape ${this.getRandomId()}"/>
        <p:cNvSpPr>
            <a:spLocks noGrp="1"/>
        </p:cNvSpPr>
        <p:nvPr>
            <p:ph type="body"/>
        </p:nvPr>
    </p:nvSpPr>
    <p:spPr>
        <a:xfrm>
            <a:off x="${element.position.x}" y="${element.position.y}"/>
            <a:ext cx="${element.size.width}" cy="${element.size.height}"/>
        </a:xfrm>
        <a:prstGeom prst="rect">
            <a:avLst/>
        </a:prstGeom>
    </p:spPr>
    <p:txBody>
        <a:bodyPr vertOverflow="clip" horzOverflow="clip" vert="horz" wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0" anchor="t">
            <a:spAutoFit/>
        </a:bodyPr>
        <a:lstStyle/>
        <a:p>
            <a:pPr algn="${align}">
                <a:defRPr sz="${fontSize * 100}" b="${bold ? '1' : '0'}" i="${italic ? '1' : '0'}" lang="en-US">
                    <a:solidFill>
                        <a:srgbClr val="${color.replace('#', '')}"/>
                    </a:solidFill>
                    <a:latin typeface="${fontFamily}"/>
                </a:defRPr>
            </a:pPr>
            <a:r>
                <a:rPr lang="en-US">
                    <a:solidFill>
                        <a:srgbClr val="${color.replace('#', '')}"/>
                    </a:solidFill>
                    <a:latin typeface="${fontFamily}"/>
                </a:rPr>
                <a:t>${this.escapeXml(content)}</a:t>
            </a:r>
        </a:p>
    </p:txBody>
</p:sp>`;
  }

  /**
   * Generate table element XML
   */
  private generateTableElementXML(element: RealPPTXElement): string {
    const tableContent = element.content as any;
    if (!tableContent || !tableContent.rows) return '';

    const rows = tableContent.rows.map((row: string[], rowIndex: number) => {
      const cells = row.map((cell: string, cellIndex: number) => {
        const bgColor = rowIndex % 2 === 0 
          ? (tableContent.style?.rowBackground || '#FFFFFF')
          : (tableContent.style?.alternateRowBackground || '#F2F2F2');
        
        return `<a:tc>
            <a:txBody>
                <a:bodyPr/>
                <a:lstStyle/>
                <a:p>
                    <a:pPr>
                        <a:defRPr sz="${(tableContent.style?.fontSize || 18) * 100}" lang="en-US">
                            <a:solidFill>
                                <a:srgbClr val="${(tableContent.style?.headerColor || '#000000').replace('#', '')}"/>
                            </a:solidFill>
                            <a:latin typeface="${tableContent.style?.fontFamily || 'Calibri'}"/>
                        </a:defRPr>
                    </a:pPr>
                    <a:r>
                        <a:rPr lang="en-US">
                            <a:solidFill>
                                <a:srgbClr val="${(tableContent.style?.headerColor || '#000000').replace('#', '')}"/>
                            </a:solidFill>
                            <a:latin typeface="${tableContent.style?.fontFamily || 'Calibri'}"/>
                        </a:rPr>
                        <a:t>${this.escapeXml(cell)}</a:t>
                    </a:r>
                </a:p>
            </a:txBody>
            <a:tcPr>
                <a:fill>
                    <a:solidFill>
                        <a:srgbClr val="${bgColor.replace('#', '')}"/>
                    </a:solidFill>
                </a:fill>
                <a:ln w="12700" cmpd="sng">
                    <a:solidFill>
                        <a:srgbClr val="${(tableContent.style?.borderColor || '#D9D9D9').replace('#', '')}"/>
                    </a:solidFill>
                </a:ln>
            </a:tcPr>
        </a:tc>`;
      }).join('');
      
      return `<a:tr h="370840">${cells}</a:tr>`;
    }).join('');

    const headers = tableContent.headers?.map((header: string, index: number) => {
      return `<a:tc>
            <a:txBody>
                <a:bodyPr/>
                <a:lstStyle/>
                <a:p>
                    <a:pPr>
                        <a:defRPr sz="${(tableContent.style?.fontSize || 18) * 100}" b="1" lang="en-US">
                            <a:solidFill>
                                <a:srgbClr val="${(tableContent.style?.headerColor || '#FFFFFF').replace('#', '')}"/>
                            </a:solidFill>
                            <a:latin typeface="${tableContent.style?.fontFamily || 'Calibri'}"/>
                        </a:defRPr>
                    </a:pPr>
                    <a:r>
                        <a:rPr b="1" lang="en-US">
                            <a:solidFill>
                                <a:srgbClr val="${(tableContent.style?.headerColor || '#FFFFFF').replace('#', '')}"/>
                            </a:solidFill>
                            <a:latin typeface="${tableContent.style?.fontFamily || 'Calibri'}"/>
                        </a:rPr>
                        <a:t>${this.escapeXml(header)}</a:t>
                    </a:r>
                </a:p>
            </a:txBody>
            <a:tcPr>
                <a:fill>
                    <a:solidFill>
                        <a:srgbClr val="${(tableContent.style?.headerBackground || '#2E74B5').replace('#', '')}"/>
                    </a:solidFill>
                </a:fill>
                <a:ln w="12700" cmpd="sng">
                    <a:solidFill>
                        <a:srgbClr val="${(tableContent.style?.borderColor || '#D9D9D9').replace('#', '')}"/>
                    </a:solidFill>
                </a:ln>
            </a:tcPr>
        </a:tc>`;
    }).join('');

    return `<p:graphicFrame>
    <p:nvGraphicFramePr>
        <p:cNvPr id="${this.getRandomId()}" name="Table ${this.getRandomId()}"/>
        <p:cNvGraphicFramePr>
            <a:graphicFrameLocks noGrp="1"/>
        </p:cNvGraphicFramePr>
        <p:nvPr/>
    </p:nvGraphicFramePr>
    <p:xfrm>
        <a:off x="${element.position.x}" y="${element.position.y}"/>
        <a:ext cx="${element.size.width}" cy="${element.size.height}"/>
    </p:xfrm>
    <a:graphic>
        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
            <a:tbl>
                <a:tblPr firstRow="1" bandRow="1">
                    <a:tableStyleId="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}</a:tableStyleId>
                </a:tblPr>
                <a:tblGrid>
                    ${tableContent.headers?.map(() => '<a:gridCol w="2286000"/>').join('')}
                </a:tblGrid>
                <a:tr h="370840">${headers}</a:tr>
                ${rows}
            </a:tbl>
        </a:graphicData>
    </a:graphic>
</p:graphicFrame>`;
  }

  /**
   * Generate chart element XML
   */
  private generateChartElementXML(element: RealPPTXElement): string {
    return `<p:graphicFrame>
    <p:nvGraphicFramePr>
        <p:cNvPr id="${this.getRandomId()}" name="Chart ${this.getRandomId()}"/>
        <p:cNvGraphicFramePr>
            <a:graphicFrameLocks noGrp="1"/>
        </p:cNvGraphicFramePr>
        <p:nvPr/>
    </p:nvGraphicFramePr>
    <p:xfrm>
        <a:off x="${element.position.x}" y="${element.position.y}"/>
        <a:ext cx="${element.size.width}" cy="${element.size.height}"/>
    </p:xfrm>
    <a:graphic>
        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
            <c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rId1"/>
        </a:graphicData>
    </a:graphic>
</p:graphicFrame>`;
  }

  /**
   * Generate theme XML
   */
  private async generateTheme(): Promise<void> {
    const themeXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">
    <a:themeElements>
        <a:clrScheme name="Office">
            <a:dk1>
                <a:sysClr val="windowText" lastClr="000000"/>
            </a:dk1>
            <a:lt1>
                <a:sysClr val="window" lastClr="FFFFFF"/>
            </a:lt1>
            <a:dk2>
                <a:srgbClr val="${this.template.colors.primary.replace('#', '')}"/>
            </a:dk2>
            <a:lt2>
                <a:srgbClr val="${this.template.colors.background.replace('#', '')}"/>
            </a:lt2>
            <a:accent1>
                <a:srgbClr val="${this.template.colors.accent1.replace('#', '')}"/>
            </a:accent1>
            <a:accent2>
                <a:srgbClr val="${this.template.colors.accent2.replace('#', '')}"/>
            </a:accent2>
            <a:accent3>
                <a:srgbClr val="${this.template.colors.accent3.replace('#', '')}"/>
            </a:accent3>
            <a:hlink>
                <a:srgbClr val="0563C1"/>
            </a:hlink>
            <a:folHlink>
                <a:srgbClr val="954F72"/>
            </a:folHlink>
        </a:clrScheme>
        <a:fontScheme name="Office">
            <a:majorFont>
                <a:latin typeface="${this.template.fonts.heading.latin}"/>
                <a:ea typeface=""/>
                <a:cs typeface=""/>
            </a:majorFont>
            <a:minorFont>
                <a:latin typeface="${this.template.fonts.body.latin}"/>
                <a:ea typeface=""/>
                <a:cs typeface=""/>
            </a:minorFont>
        </a:fontScheme>
        <a:fmtScheme name="Office">
            <a:fillStyleLst>
                <a:solidFill>
                    <a:schemeClr val="phClr"/>
                </a:solidFill>
                <a:gradFill rotWithShape="1">
                    <a:gsLst>
                        <a:gs pos="0">
                            <a:schemeClr val="phClr">
                                <a:tint val="50000"/>
                                <a:satMod val="300000"/>
                            </a:schemeClr>
                        </a:gs>
                        <a:gs pos="35000">
                            <a:schemeClr val="phClr">
                                <a:tint val="37000"/>
                                <a:satMod val="300000"/>
                            </a:schemeClr>
                        </a:gs>
                        <a:gs pos="100000">
                            <a:schemeClr val="phClr">
                                <a:tint val="15000"/>
                                <a:satMod val="350000"/>
                            </a:schemeClr>
                        </a:gs>
                    </a:gsLst>
                    <a:lin ang="16200000" scaled="1"/>
                </a:gradFill>
                <a:gradFill rotWithShape="1">
                    <a:gsLst>
                        <a:gs pos="0">
                            <a:schemeClr val="phClr">
                                <a:tint val="100000"/>
                                <a:shade val="100000"/>
                                <a:satMod val="130000"/>
                            </a:schemeClr>
                        </a:gs>
                        <a:gs pos="100000">
                            <a:schemeClr val="phClr">
                                <a:tint val="50000"/>
                                <a:shade val="100000"/>
                                <a:satMod val="350000"/>
                            </a:schemeClr>
                        </a:gs>
                    </a:gsLst>
                    <a:lin ang="16200000" scaled="0"/>
                </a:gradFill>
            </a:fillStyleLst>
            <a:lnStyleLst>
                <a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">
                    <a:solidFill>
                        <a:schemeClr val="phClr">
                            <a:shade val="95000"/>
                            <a:satMod val="105000"/>
                        </a:schemeClr>
                    </a:solidFill>
                    <a:prstDash val="solid"/>
                </a:ln>
                <a:ln w="25400" cap="flat" cmpd="sng" algn="ctr">
                    <a:solidFill>
                        <a:schemeClr val="phClr"/>
                    </a:solidFill>
                    <a:prstDash val="solid"/>
                </a:ln>
                <a:ln w="38100" cap="flat" cmpd="sng" algn="ctr">
                    <a:solidFill>
                        <a:schemeClr val="phClr"/>
                    </a:solidFill>
                    <a:prstDash val="solid"/>
                </a:ln>
            </a:lnStyleLst>
            <a:effectStyleLst>
                <a:effectStyle>
                    <a:effectLst>
                        <a:outerShdw blurRad="40000" dist="20000" dir="5400000" rotWithShape="0">
                            <a:srgbClr val="000000">
                                <a:alpha val="38000"/>
                            </a:srgbClr>
                        </a:outerShdw>
                    </a:effectLst>
                </a:effectStyle>
                <a:effectStyle>
                    <a:effectLst>
                        <a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0">
                            <a:srgbClr val="000000">
                                <a:alpha val="35000"/>
                            </a:srgbClr>
                        </a:outerShdw>
                    </a:effectLst>
                </a:effectStyle>
                <a:effectStyle>
                    <a:effectLst>
                        <a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0">
                            <a:srgbClr val="000000">
                                <a:alpha val="35000"/>
                            </a:srgbClr>
                        </a:outerShdw>
                    </a:effectLst>
                </a:effectStyle>
            </a:effectStyleLst>
            <a:bgFillStyleLst>
                <a:solidFill>
                    <a:schemeClr val="phClr"/>
                </a:solidFill>
                <a:gradFill rotWithShape="1">
                    <a:gsLst>
                        <a:gs pos="0">
                            <a:schemeClr val="phClr">
                                <a:tint val="40000"/>
                                <a:satMod val="350000"/>
                            </a:schemeClr>
                        </a:gs>
                        <a:gs pos="40000">
                            <a:schemeClr val="phClr">
                                <a:tint val="45000"/>
                                <a:shade val="99000"/>
                                <a:satMod val="350000"/>
                            </a:schemeClr>
                        </a:gs>
                        <a:gs pos="100000">
                            <a:schemeClr val="phClr">
                                <a:shade val="20000"/>
                                <a:satMod val="255000"/>
                            </a:schemeClr>
                        </a:gs>
                    </a:gsLst>
                    <a:lin ang="16200000" scaled="0"/>
                </a:gradFill>
                <a:gradFill rotWithShape="1">
                    <a:gsLst>
                        <a:gs pos="0">
                            <a:schemeClr val="phClr">
                                <a:tint val="80000"/>
                                <a:satMod val="300000"/>
                            </a:schemeClr>
                        </a:gs>
                        <a:gs pos="100000">
                            <a:schemeClr val="phClr">
                                <a:shade val="80000"/>
                                <a:satMod val="300000"/>
                            </a:schemeClr>
                        </a:gs>
                    </a:gsLst>
                    <a:lin ang="16200000" scaled="0"/>
                </a:gradFill>
            </a:bgFillStyleLst>
        </a:fmtScheme>
    </a:themeElements>
    <a:objectDefaults/>
    <a:extraClrSchemeLst/>
</a:theme>`;

    this.zip.file('ppt/theme/theme1.xml', themeXml);
  }

  /**
   * Generate slide master
   */
  private async generateSlideMaster(): Promise<void> {
    const slideMasterXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
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
        </p:spTree>
    </p:cSld>
    <p:clrMap>
        <a:extLst>
            <a:ext uri="{BB962522-66D1-4E39-9F3F-89F3731E6A1F}">
                <a14:backgroundClrMap xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" bg1="lt1" bg2="lt2" tx1="dk1" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
            </a:ext>
        </a:extLst>
    </p:clrMap>
    <p:sldLayoutIdLst>
        <p:sldLayoutId id="2147483649" r:id="rId1"/>
    </p:sldLayoutIdLst>
    <p:txStyles>
        <a:titleStyle>
            <a:defPPr>
                <a:defRPr lang="en-US"/>
            </a:defPPr>
        </a:titleStyle>
        <a:bodyStyle>
            <a:defPPr>
                <a:defRPr lang="en-US"/>
            </a:defPPr>
        </a:bodyStyle>
        <a:otherStyle>
            <a:defPPr>
                <a:defRPr lang="en-US"/>
            </a:defPPr>
        </a:otherStyle>
    </p:txStyles>
</p:sldMaster>`;

    this.zip.file('ppt/slideMasters/slideMaster1.xml', slideMasterXml);
    
    // Generate slide master relationships
    const slideMasterRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
    <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>
</Relationships>`;

    this.zip.folder('ppt/slideMasters/_rels');
    this.zip.file('ppt/slideMasters/_rels/slideMaster1.xml.rels', slideMasterRels);
  }

  /**
   * Generate slide layouts
   */
  private async generateSlideLayouts(): Promise<void> {
    const slideLayoutXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="title" preserve="1">
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
        </p:spTree>
    </p:cSld>
    <p:clrMapOvr>
        <a:masterClrMapping/>
    </p:clrMapOvr>
</p:sldLayout>`;

    this.zip.file('ppt/slideLayouts/slideLayout1.xml', slideLayoutXml);
  }

  /**
   * Apply data to template
   */
  private applyDataToTemplate(): RealPPTXSlide[] {
    return this.template.slides.map(slide => ({
      ...slide,
      elements: slide.elements.map(element => ({
        ...element,
        content: this.processElementContent(element.content, this.data)
      }))
    }));
  }

  /**
   * Process element content with data
   */
  private processElementContent(content: any, data: Record<string, any>): any {
    if (typeof content === 'string') {
      return content.replace(/\{([^}]+)\}/g, (match, key) => {
        return data[key] !== undefined ? String(data[key]) : match;
      });
    }

    if (content && typeof content === 'object') {
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

  /**
   * Generate content types
   */
  private async generateContentTypes(): Promise<void> {
    // Already handled in createPPTXStructure
  }

  /**
   * Generate relationships
   */
  private async generateRelationships(): Promise<void> {
    // Already handled in createPPTXStructure
  }

  /**
   * Helper methods
   */
  private getRandomId(): number {
    return Math.floor(Math.random() * 1000000) + 1;
  }

  private escapeXml(text: string): string {
    return text
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  }
}