/**
 * Improved PPTX Builder - Creates fully compliant PPTX files
 */

import JSZip from 'jszip';

export interface PPTXData {
  [key: string]: any;
}

export class ImprovedPPTXBuilder {
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
      console.error('Error loading template, creating default:', error);
      this.createCompleteDefaultTemplate();
    }
  }

  /**
   * Set data for replacement
   */
  setData(data: PPTXData): void {
    this.data = data;
  }

  /**
   * Create a complete default template with all required files
   */
  private createCompleteDefaultTemplate(): void {
    this.zip = new JSZip();
    
    // 1. Content Types
    this.zip.file('[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml" ContentType="application/xml"/>
    <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-presentationml.presentation.main+xml"/>
    <Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-presentationml.slideMaster+xml"/>
    <Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-presentationml.slideLayout+xml"/>
    <Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-presentationml.theme+xml"/>
    <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-presentationml.slide+xml"/>
    <Override PartName="/ppt/slides/slide2.xml" ContentType="application/vnd.openxmlformats-presentationml.slide+xml"/>
    <Override PartName="/ppt/slides/slide3.xml" ContentType="application/vnd.openxmlformats-presentationml.slide+xml"/>
    <Override PartName="/ppt/slides/slide4.xml" ContentType="application/vnd.openxmlformats-presentationml.slide+xml"/>
</Types>`);

    // 2. Root relationships
    this.zip.file('_rels/.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>`);

    // 3. Presentation XML
    this.zip.file('ppt/presentation.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
    <p:sldMasterIdLst>
        <p:sldMasterId id="2147483648" r:id="rId1"/>
    </p:sldMasterIdLst>
    <p:sldIdLst>
        <p:sldId id="256" r:id="rId2"/>
        <p:sldId id="257" r:id="rId3"/>
        <p:sldId id="258" r:id="rId4"/>
        <p:sldId id="259" r:id="rId5"/>
    </p:sldIdLst>
    <p:sldSz cx="9144000" cy="5143500"/>
    <p:notesSz cx="6858000" cy="9144000"/>
    <p:defaultTextStyle>
        <a:defPPr>
            <a:defRPr lang="en-US"/>
        </a:defPPr>
    </p:defaultTextStyle>
</p:presentation>`);

    // 4. Presentation relationships
    this.zip.file('ppt/_rels/presentation.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>
    <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
    <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide2.xml"/>
    <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide3.xml"/>
    <Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide4.xml"/>
</Relationships>`);

    // 5. Theme
    this.zip.file('ppt/theme/theme1.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">
    <a:themeElements>
        <a:clrScheme name="Office">
            <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
            <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
            <a:dk2><a:srgbClr val="44546A"/></a:dk2>
            <a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>
            <a:accent1><a:srgbClr val="4472C4"/></a:accent1>
            <a:accent2><a:srgbClr val="ED7D31"/></a:accent2>
            <a:accent3><a:srgbClr val="A5A5A5"/></a:accent3>
            <a:accent4><a:srgbClr val="FFC000"/></a:accent4>
            <a:accent5><a:srgbClr val="5B9BD5"/></a:accent5>
            <a:accent6><a:srgbClr val="70AD47"/></a:accent6>
        </a:clrScheme>
        <a:fontScheme name="Office">
            <a:majorFont>
                <a:latin typeface="Calibri Light" panose="020F0302020204030204"/>
                <a:ea typeface=""/>
                <a:cs typeface=""/>
            </a:majorFont>
            <a:minorFont>
                <a:latin typeface="Calibri" panose="020F0502020204030204"/>
                <a:ea typeface=""/>
                <a:cs typeface=""/>
            </a:minorFont>
        </a:fontScheme>
        <a:fmtScheme name="Office">
            <a:fillStyleLst>
                <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
                <a:gradFill rotWithShape="1">
                    <a:gsLst>
                        <a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="50000"/><a:satMod val="300000"/></a:schemeClr></a:gs>
                        <a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="50000"/><a:satMod val="300000"/></a:schemeClr></a:gs>
                    </a:gsLst>
                    <a:lin ang="540000" scaled="1"/>
                </a:gradFill>
            </a:fillStyleLst>
            <a:lnStyleLst>
                <a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">
                    <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
                    <a:prstDash val="solid"/>
                    <a:miter lim="800000"/>
                </a:ln>
            </a:lnStyleLst>
            <a:effectStyleLst>
                <a:effectStyle>
                    <a:effectLst>
                        <a:outerShdw blurRad="57150" dist="19050" dir="5400000" algn="ctr" rotWithShape="0">
                            <a:srgbClr val="000000"><a:alpha val="32000"/></a:srgbClr>
                        </a:outerShdw>
                    </a:effectLst>
                </a:effectStyle>
            </a:effectStyleLst>
            <a:bgFillStyleLst>
                <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
                <a:gradFill rotWithShape="1">
                    <a:gsLst>
                        <a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="50000"/><a:satMod val="300000"/></a:schemeClr></a:gs>
                        <a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="50000"/><a:satMod val="300000"/></a:schemeClr></a:gs>
                    </a:gsLst>
                    <a:lin ang="540000" scaled="1"/>
                </a:gradFill>
            </a:bgFillStyleLst>
        </a:fmtScheme>
    </a:themeElements>
    <a:objectDefaults/>
    <a:extraClrSchemeLst/>
    <a:extLst>
        <a:ext uri="{05A44725-979E-4E34-AE85-DB4B8770C9B3}">
            <a14:themeDefaults xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main">
                <a14:clrMru>
                    <a14:clr idx="0"><a:srgbClr val="FFFFFF"/></a14:clr>
                    <a14:clr idx="1"><a:srgbClr val="000000"/></a14:clr>
                    <a14:clr idx="2"><a:srgbClr val="EEECE1"/></a14:clr>
                    <a14:clr idx="3"><a:srgbClr val="1F497D"/></a14:clr>
                    <a14:clr idx="4"><a:srgbClr val="4F81BD"/></a14:clr>
                    <a14:clr idx="5"><a:srgbClr val="C0504D"/></a14:clr>
                    <a14:clr idx="6"><a:srgbClr val="9BBB59"/></a14:clr>
                    <a14:clr idx="7"><a:srgbClr val="8064A2"/></a14:clr>
                    <a14:clr idx="8"><a:srgbClr val="4BACC6"/></a14:clr>
                    <a14:clr idx="9"><a:srgbClr val="F79646"/></a14:clr>
                </a14:clrMru>
            </a14:themeDefaults>
        </a:ext>
    </a:extLst>
</a:theme>`);

    // 6. Slide Master
    this.zip.file('ppt/slideMasters/slideMaster1.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
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
    <p:clrMapOvr>
        <a:masterClrMapping/>
    </p:clrMapOvr>
    <p:sldLayoutIdLst>
        <p:sldLayoutId id="2147483649" r:id="rId1"/>
    </p:sldLayoutIdLst>
</p:sldMaster>`);

    // 7. Slide Master relationships
    this.zip.file('ppt/slideMasters/_rels/slideMaster1.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
    <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>
</Relationships>`);

    // 8. Slide Layout
    this.zip.file('ppt/slideLayouts/slideLayout1.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
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
                <p:spPr/>
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
                <p:spPr/>
            </p:sp>
        </p:spTree>
    </p:cSld>
    <p:clrMapOvr>
        <a:masterClrMapping/>
    </p:clrMapOvr>
</p:sldLayout>`);

    // 9. Slide Layout relationships
    this.zip.file('ppt/slideLayouts/_rels/slideLayout1.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
</Relationships>`);

    // 10. Create basic slides
    this.createDefaultSlides();
  }

  /**
   * Create default slides
   */
  private createDefaultSlides(): void {
    // Slide 1 - Title
    this.zip.file('ppt/slides/slide1.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
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
                            <a:rPr lang="en-US" sz="4400" b="1">
                                <a:solidFill>
                                    <a:srgbClr val="363636"/>
                                </a:solidFill>
                                <a:latin typeface="Calibri"/>
                            </a:rPr>
                            <a:t>{project_title}</a:t>
                        </a:r>
                    </a:p>
                </p:txBody>
            </p:sp>
            <p:sp>
                <p:nvSpPr>
                    <p:cNvPr id="3" name="Subtitle"/>
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
                        <a:ext cx="8229600" cy="2286000"/>
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
                            <a:rPr lang="en-US" sz="2000">
                                <a:solidFill>
                                    <a:srgbClr val="666666"/>
                                </a:solidFill>
                                <a:latin typeface="Calibri"/>
                            </a:rPr>
                            <a:t>Status Report prepared by: {name} - {date}</a:t>
                        </a:r>
                    </a:p>
                </p:txBody>
            </p:sp>
        </p:spTree>
    </p:cSld>
    <p:clrMapOvr>
        <a:masterClrMapping/>
    </p:clrMapOvr>
</p:sld>`);

    // Slide 1 relationships
    this.zip.file('ppt/slides/_rels/slide1.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
</Relationships>`);

    // Slide 2 - Overview
    this.zip.file('ppt/slides/slide2.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
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
                        <a:pPr>
                            <a:buNone/>
                        </a:pPr>
                        <a:r>
                            <a:rPr lang="en-US" sz="3600" b="1">
                                <a:solidFill>
                                    <a:srgbClr val="363636"/>
                                </a:solidFill>
                                <a:latin typeface="Calibri"/>
                            </a:rPr>
                            <a:t>Overview</a:t>
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
                        <a:ext cx="8229600" cy="4572000"/>
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
                            <a:rPr lang="en-US" sz="1800">
                                <a:solidFill>
                                    <a:srgbClr val="444444"/>
                                </a:solidFill>
                                <a:latin typeface="Calibri"/>
                            </a:rPr>
                            <a:t>{status_update}</a:t>
                        </a:r>
                    </a:p>
                </p:txBody>
            </p:sp>
        </p:spTree>
    </p:cSld>
    <p:clrMapOvr>
        <a:masterClrMapping/>
    </p:clrMapOvr>
</p:sld>`);

    // Slide 2 relationships
    this.zip.file('ppt/slides/_rels/slide2.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
</Relationships>`);

    // Slide 3 - Agenda
    this.zip.file('ppt/slides/slide3.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
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
                        <a:pPr>
                            <a:buNone/>
                        </a:pPr>
                        <a:r>
                            <a:rPr lang="en-US" sz="3600" b="1">
                                <a:solidFill>
                                    <a:srgbClr val="363636"/>
                                </a:solidFill>
                                <a:latin typeface="Calibri"/>
                            </a:rPr>
                            <a:t>Agenda</a:t>
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
                        <a:ext cx="8229600" cy="4572000"/>
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
                            <a:rPr lang="en-US" sz="1800">
                                <a:solidFill>
                                    <a:srgbClr val="444444"/>
                                </a:solidFill>
                                <a:latin typeface="Calibri"/>
                            </a:rPr>
                            <a:t>{agenda}</a:t>
                        </a:r>
                    </a:p>
                </p:txBody>
            </p:sp>
        </p:spTree>
    </p:cSld>
    <p:clrMapOvr>
        <a:masterClrMapping/>
    </p:clrMapOvr>
</p:sld>`);

    // Slide 3 relationships
    this.zip.file('ppt/slides/_rels/slide3.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
</Relationships>`);

    // Slide 4 - Chart
    this.zip.file('ppt/slides/slide4.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
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
                        <a:pPr>
                            <a:buNone/>
                        </a:pPr>
                        <a:r>
                            <a:rPr lang="en-US" sz="3600" b="1">
                                <a:solidFill>
                                    <a:srgbClr val="363636"/>
                                </a:solidFill>
                                <a:latin typeface="Calibri"/>
                            </a:rPr>
                            <a:t>{ChartTitle}</a:t>
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
                        <a:ext cx="8229600" cy="4572000"/>
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
                            <a:rPr lang="en-US" sz="1800">
                                <a:solidFill>
                                    <a:srgbClr val="444444"/>
                                </a:solidFill>
                                <a:latin typeface="Calibri"/>
                            </a:rPr>
                            <a:t>This is a placeholder slide. The title above will be filled from your Excel data. Charts defined in the 'ChartData' sheet will be added as new slides following this one.</a:t>
                        </a:r>
                    </a:p>
                </p:txBody>
            </p:sp>
        </p:spTree>
    </p:cSld>
    <p:clrMapOvr>
        <a:masterClrMapping/>
    </p:clrMapOvr>
</p:sld>`);

    // Slide 4 relationships
    this.zip.file('ppt/slides/_rels/slide4.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
</Relationships>`);
  }

  /**
   * Simple text replacement
   */
  private simpleTextReplacement(content: string): string {
    let result = content;
    
    // Replace common placeholders
    const replacements = {
      '{project_title}': this.data.project_title || 'Project Title',
      '{name}': this.data.name || 'Presenter Name',
      '{date}': this.data.date || new Date().toLocaleDateString(),
      '{status_update}': this.data.status_update || 'Status update will appear here.',
      '{agenda}': this.data.agenda || 'Agenda items will appear here.',
      '{ChartTitle}': this.data.ChartTitle || 'Chart Title'
    };

    for (const [placeholder, value] of Object.entries(replacements)) {
      result = result.replace(new RegExp(placeholder.replace(/[{}]/g, '\\$&'), 'g'), this.escapeXml(value));
    }

    return result;
  }

  /**
   * Escape XML special characters
   */
  private escapeXml(text: string): string {
    return text
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  }

  /**
   * Build the final PPTX
   */
  async build(): Promise<ArrayBuffer> {
    try {
      // Process existing slides
      const slideFiles = this.zip.file(/ppt\/slides\/slide\d+\.xml/);
      
      for (const file of slideFiles) {
        const content = await file.async('string');
        const processedContent = this.simpleTextReplacement(content);
        this.zip.file(file.name, processedContent);
      }

      // Generate the buffer
      return await this.zip.generateAsync({ type: 'arraybuffer' });
    } catch (error) {
      console.error('Error building PPTX:', error);
      throw new Error('Failed to build PPTX file');
    }
  }

  /**
   * Build PPTX with charts
   */
  async buildWithCharts(chartData: any[]): Promise<ArrayBuffer> {
    try {
      // First, process existing slides
      const slideFiles = this.zip.file(/ppt\/slides\/slide\d+\.xml/);
      
      for (const file of slideFiles) {
        const content = await file.async('string');
        const processedContent = this.simpleTextReplacement(content);
        this.zip.file(file.name, processedContent);
      }

      // Add chart slides if needed
      if (chartData && chartData.length > 0) {
        await this.addSimpleChartSlides(chartData, slideFiles.length);
      }

      // Generate the buffer
      return await this.zip.generateAsync({ type: 'arraybuffer' });
    } catch (error) {
      console.error('Error building PPTX with charts:', error);
      throw new Error('Failed to build PPTX file with charts');
    }
  }

  /**
   * Add simple chart slides
   */
  private async addSimpleChartSlides(chartData: any[], currentSlideCount: number): Promise<void> {
    try {
      // Group chart data by title
      const chartGroups = this.groupChartData(chartData);
      
      let slideNum = currentSlideCount + 1;
      for (const [chartTitle, chartGroup] of Object.entries(chartGroups)) {
        // Create a simple chart slide
        const chartSlideContent = this.createSimpleChartSlide(chartTitle, chartGroup as any[]);
        const slideFileName = `ppt/slides/slide${slideNum}.xml`;
        
        this.zip.file(slideFileName, chartSlideContent);
        
        // Create relationship file
        const relsContent = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
</Relationships>`;
        
        this.zip.file(`ppt/slides/_rels/slide${slideNum}.xml.rels`, relsContent);
        
        slideNum++;
      }

      // Update presentation.xml
      await this.updatePresentationXml(currentSlideCount, Object.keys(chartGroups).length);
      
      // Update content types
      await this.updateContentTypes(currentSlideCount, Object.keys(chartGroups).length);
      
    } catch (error) {
      console.error('Error adding chart slides:', error);
    }
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
   * Create simple chart slide
   */
  private createSimpleChartSlide(chartTitle: string, chartGroup: any[]): string {
    const chartInfo = chartGroup.map(item => `${item.Label || ''}: ${item.Value || ''}`).join(', ');
    
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
              <a:rPr lang="en-US" sz="4400" b="1">
                <a:solidFill>
                  <a:srgbClr val="363636"/>
                </a:solidFill>
                <a:latin typeface="Calibri"/>
              </a:rPr>
              <a:t>${this.escapeXml(chartTitle)}</a:t>
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
            <a:ext cx="8229600" cy="4572000"/>
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
              <a:rPr lang="en-US" sz="1800">
                <a:solidFill>
                  <a:srgbClr val="444444"/>
                </a:solidFill>
                <a:latin typeface="Calibri"/>
              </a:rPr>
              <a:t>Chart Data: ${this.escapeXml(chartInfo)}</a:t>
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
   * Update presentation.xml
   */
  private async updatePresentationXml(currentSlideCount: number, newSlidesCount: number): Promise<void> {
    try {
      const presFile = this.zip.file('ppt/presentation.xml');
      if (!presFile) return;

      const presContent = await presFile.async('string');
      
      // Simple string-based update for slide IDs
      let updatedContent = presContent;
      
      // Find the last slide ID and add new ones
      const slideIdMatch = presContent.match(/<p:sldId[^>]*id="(\d+)"[^>]*>/g);
      if (slideIdMatch && slideIdMatch.length > 0) {
        const lastSlideMatch = slideIdMatch[slideIdMatch.length - 1];
        const lastIdMatch = lastSlideMatch.match(/id="(\d+)"/);
        if (lastIdMatch) {
          const lastId = parseInt(lastIdMatch[1]);
          const sldIdLstEnd = presContent.indexOf('</p:sldIdLst>');
          
          let newSlideIds = '';
          let newSlideRels = '';
          
          for (let i = 1; i <= newSlidesCount; i++) {
            const newId = lastId + i;
            const newRelId = 5 + i; // Starting from rId6 since we have rId1-5 already
            newSlideIds += `        <p:sldId id="${newId}" r:id="rId${newRelId}"/>\n`;
          }
          
          updatedContent = presContent.slice(0, sldIdLstEnd) + '\n' + newSlideIds + presContent.slice(sldIdLstEnd);
          
          // Update presentation relationships
          const presRelsFile = this.zip.file('ppt/_rels/presentation.xml.rels');
          if (presRelsFile) {
            const presRelsContent = await presRelsFile.async('string');
            const relsEnd = presRelsContent.indexOf('</Relationships>');
            let newRels = '';
            
            for (let i = 1; i <= newSlidesCount; i++) {
              const newRelId = 5 + i;
              const slideNum = currentSlideCount + i;
              newRels += `    <Relationship Id="rId${newRelId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${slideNum}.xml"/>\n`;
            }
            
            const updatedRelsContent = presRelsContent.slice(0, relsEnd) + '\n' + newRels + presRelsContent.slice(relsEnd);
            this.zip.file('ppt/_rels/presentation.xml.rels', updatedRelsContent);
          }
        }
      }
      
      this.zip.file('ppt/presentation.xml', updatedContent);
    } catch (error) {
      console.error('Error updating presentation.xml:', error);
    }
  }

  /**
   * Update content types
   */
  private async updateContentTypes(currentSlideCount: number, newSlidesCount: number): Promise<void> {
    try {
      const contentTypesFile = this.zip.file('[Content_Types].xml');
      if (!contentTypesFile) return;

      const contentTypesContent = await contentTypesFile.async('string');
      const typesEnd = contentTypesContent.indexOf('</Types>');
      
      let newTypes = '';
      
      for (let i = 1; i <= newSlidesCount; i++) {
        const slideNum = currentSlideCount + i;
        newTypes += `    <Override PartName="/ppt/slides/slide${slideNum}.xml" ContentType="application/vnd.openxmlformats-presentationml.slide+xml"/>\n`;
      }
      
      const updatedContentTypesContent = contentTypesContent.slice(0, typesEnd) + '\n' + newTypes + contentTypesContent.slice(typesEnd);
      this.zip.file('[Content_Types].xml', updatedContentTypesContent);
    } catch (error) {
      console.error('Error updating content types:', error);
    }
  }
}