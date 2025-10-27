'use client'

import React, { useState, useEffect, useCallback } from 'react'
import { Card, CardContent, CardHeader, CardTitle } from '@/components/ui/card'
import { Button } from '@/components/ui/button'
import { Badge } from '@/components/ui/badge'
import { Progress } from '@/components/ui/progress'
import { 
  Eye, 
  Edit3, 
  RefreshCw, 
  Download, 
  ZoomIn, 
  ZoomOut,
  Move,
  Type,
  BarChart3,
  Image as ImageIcon
} from 'lucide-react'
import { DynamicTemplateEngine, defaultTemplateConfig } from '@/lib/dynamic-templates'
import { RealPPTXParser, RealPPTXSlide, RealPPTXElement } from '@/lib/real-pptx-parser'

interface SlideElement {
  id: string
  type: 'text' | 'chart' | 'image' | 'shape' | 'table'
  content: any
  position: { x: number; y: number }
  size: { width: number; height: number }
  style: Record<string, any>
}

interface Slide {
  id: string
  elements: SlideElement[]
  layout: string
  background: string
}

interface RealtimePreviewProps {
  template: string
  data: Record<string, any>
  pptxFile?: File | null
  onTemplateChange?: (template: string) => void
  onDataChange?: (data: Record<string, any>) => void
}

export default function RealtimePreview({ 
  template, 
  data, 
  pptxFile,
  onTemplateChange, 
  onDataChange 
}: RealtimePreviewProps) {
  const [slides, setSlides] = useState<Slide[]>([])
  const [selectedSlide, setSelectedSlide] = useState(0)
  const [selectedElement, setSelectedElement] = useState<string | null>(null)
  const [isEditing, setIsEditing] = useState(false)
  const [zoom, setZoom] = useState(100)
  const [isProcessing, setIsProcessing] = useState(false)
  const [processingProgress, setProcessingProgress] = useState(0)
  const [editMode, setEditMode] = useState<'select' | 'text' | 'chart' | 'image'>('select')
  const [pptxParser] = useState(() => new RealPPTXParser())
  const [originalTemplate, setOriginalTemplate] = useState<RealPPTXSlide[]>([])

  // Process template and generate slides
  const generateSlides = useCallback(async () => {
    setIsProcessing(true)
    setProcessingProgress(0)

    try {
      let processedSlides: Slide[] = []

      // Process PPTX file if available
      if (pptxFile) {
        setProcessingProgress(10)
        
        // Parse the PPTX file
        const arrayBuffer = await pptxFile.arrayBuffer()
        const template = await pptxParser.parseTemplate(arrayBuffer)
        setOriginalTemplate(template.slides)
        setProcessingProgress(30)
        
        // Apply data to the template
        const dataAppliedSlides = pptxParser.applyData(data)
        setProcessingProgress(50)
        
        // Convert PPTX slides to our internal format
        processedSlides = await convertPPTXSlidesToInternalFormat(dataAppliedSlides)
        setProcessingProgress(70)
      } else {
        // Fallback to text-based processing if no PPTX file
        const engine = new DynamicTemplateEngine(defaultTemplateConfig)
        
        setProcessingProgress(20)
        await new Promise(resolve => setTimeout(resolve, 300))
        
        const processedTemplate = engine.processTemplate(template, data)
        setProcessingProgress(40)
        await new Promise(resolve => setTimeout(resolve, 300))
        
        const generatedSlides = parseSlidesFromTemplate(processedTemplate)
        setProcessingProgress(60)
        await new Promise(resolve => setTimeout(resolve, 300))
        
        const enhancedSlides = await enhanceSlidesWithInteractivity(generatedSlides)
        setProcessingProgress(80)
        await new Promise(resolve => setTimeout(resolve, 300))
        
        processedSlides = enhancedSlides
      }
      
      setSlides(processedSlides)
      setProcessingProgress(100)
      
    } catch (error) {
      console.error('Error generating slides:', error)
      // Create a fallback slide if there's an error
      const fallbackSlide = [{
        id: 'slide-error',
        layout: 'default',
        background: 'gradient',
        elements: [{
          id: 'element-error-0',
          type: 'text',
          content: 'Preview Error',
          position: { x: 1, y: 2 },
          size: { width: 8, height: 2 },
          style: {
            fontSize: 24,
            bold: true,
            align: 'center',
            color: '#dc2626'
          }
        }, {
          id: 'element-error-1',
          type: 'text',
          content: 'There was an error generating the preview. Please check your template and data.',
          position: { x: 1, y: 4 },
          size: { width: 8, height: 1.5 },
          style: {
            fontSize: 16,
            align: 'center',
            color: '#666666'
          }
        }]
      }]
      setSlides(fallbackSlide)
      setProcessingProgress(100)
    } finally {
      setIsProcessing(false)
      setTimeout(() => setProcessingProgress(0), 1000)
    }
  }, [template, data, pptxFile, pptxParser])

  // Parse slides from template
  const parseSlidesFromTemplate = (processedTemplate: string): Slide[] => {
    if (!processedTemplate || processedTemplate.trim() === '') {
      return [{
        id: 'slide-0',
        layout: 'default',
        background: 'gradient',
        elements: [{
          id: 'element-0-0',
          type: 'text',
          content: 'No content available',
          position: { x: 1, y: 2 },
          size: { width: 8, height: 2 },
          style: {
            fontSize: 24,
            bold: true,
            align: 'center',
            color: '#666666'
          }
        }]
      }]
    }

    // Split by double newlines to create slides
    const sections = processedTemplate.split(/\n\s*\n/).filter(section => section.trim())
    
    if (sections.length === 0) {
      return [{
        id: 'slide-0',
        layout: 'default',
        background: 'gradient',
        elements: [{
          id: 'element-0-0',
          type: 'text',
          content: 'No content available',
          position: { x: 1, y: 2 },
          size: { width: 8, height: 2 },
          style: {
            fontSize: 24,
            bold: true,
            align: 'center',
            color: '#666666'
          }
        }]
      }]
    }
    
    return sections.map((section, index) => ({
      id: `slide-${index}`,
      layout: 'default',
      background: 'gradient', // Use gradient background
      elements: parseElementsFromSection(section.trim(), index)
    }))
  }

  // Parse elements from slide section
  const parseElementsFromSection = (section: string, slideIndex: number): SlideElement[] => {
    const elements: SlideElement[] = []
    const lines = section.split('\n').filter(line => line.trim())
    
    // Check if this is an agenda slide
    const isAgendaSlide = lines.some(line => 
      line.trim().toLowerCase().includes('agenda') || 
      (slideIndex > 0 && lines.length > 3)
    )
    
    if (isAgendaSlide && lines.length > 1) {
      // Create agenda title
      const titleLine = lines.find(line => 
        line.trim().toLowerCase().includes('agenda')
      ) || 'Agenda'
      
      elements.push({
        id: `element-${slideIndex}-title`,
        type: 'text',
        content: titleLine.trim(),
        position: { x: 0.5, y: 1 },
        size: { width: 9, height: 1.5 },
        style: {
          fontSize: 36,
          bold: true,
          align: 'center',
          color: '#000000',
          fontFamily: 'Arial, sans-serif'
        }
      })
      
      // Parse agenda items and create table
      const agendaItems = lines.filter(line => 
        !line.trim().toLowerCase().includes('agenda') && line.trim()
      )
      
      if (agendaItems.length > 0) {
        // Combine all agenda lines and process them
        const combinedAgenda = agendaItems.join(' ').trim()
        
        // Better agenda parsing
        const items = combinedAgenda
          .split(/(?=[A-Z][a-z])|(?=\d+\.)|\s+(?=and\s+)|\s+(?=or\s+)|,\s*/)
          .filter(item => item.trim() && item.length > 2)
          .map(item => item.trim().replace(/^[,\s]+|[,\s]+$/g, ''))
        
        // If we still have very few items, try alternative splitting
        if (items.length < 2) {
          const altItems = combinedAgenda.split(/\s+/).filter(item => item.trim())
          if (altItems.length > items.length) {
            items.length = 0
            items.push(...altItems)
          }
        }
        
        // Create table structure
        elements.push({
          id: `element-${slideIndex}-table`,
          type: 'table',
          content: {
            headers: ['#', 'Topic'],
            rows: items.map((item, index) => [`${index + 1}.`, item]),
            style: {
              headerBackground: '#2E74B5',
              headerColor: '#FFFFFF',
              borderColor: '#D9D9D9',
              rowBackground: '#FFFFFF',
              alternateRowBackground: '#F2F2F2',
              fontSize: 18,
              fontFamily: 'Arial, sans-serif'
            }
          },
          position: { x: 1, y: 3 },
          size: { width: 8, height: 4 },
          style: {
            fontSize: 18,
            align: 'left',
            color: '#000000'
          }
        })
      }
    } else {
      // Regular text processing for non-agenda slides
      lines.forEach((line, index) => {
        const trimmedLine = line.trim()
        if (trimmedLine) {
          // Simple text styling based on position
          const isTitle = slideIndex === 0 && index === 0
          const isSubtitle = slideIndex === 0 && index === 1
          
          elements.push({
            id: `element-${slideIndex}-${index}`,
            type: 'text',
            content: trimmedLine,
            position: {
              x: 0.5,
              y: 1 + (index * 1.2)
            },
            size: {
              width: 9,
              height: isTitle ? 1.5 : 1
            },
            style: {
              fontSize: isTitle ? 36 : isSubtitle ? 24 : 18,
              bold: isTitle,
              align: 'left',
              color: '#000000',
              fontFamily: 'Arial, sans-serif'
            }
          })
        }
      })
    }
    
    return elements
  }

  // Enhance slides with interactive elements
  const enhanceSlidesWithInteractivity = async (slides: Slide[]): Promise<Slide[]> => {
    return slides.map((slide, slideIndex) => ({
      ...slide,
      elements: slide.elements.map(element => ({
        ...element,
        interactive: true,
        editable: true,
        draggable: true,
        resizable: true
      }))
    }))
  }

  // Handle element selection
  const handleElementClick = (elementId: string) => {
    setSelectedElement(elementId)
    setEditMode('select')
  }

  // Handle element drag
  const handleElementDrag = (elementId: string, newPosition: { x: number; y: number }) => {
    setSlides(prevSlides => 
      prevSlides.map(slide => ({
        ...slide,
        elements: slide.elements.map(element =>
          element.id === elementId 
            ? { ...element, position: newPosition }
            : element
        )
      }))
    )
  }

  // Add new element
  const addElement = (type: 'text' | 'chart' | 'image') => {
    const newElement: SlideElement = {
      id: `element-new-${Date.now()}`,
      type,
      content: type === 'text' ? 'New Text' : type === 'chart' ? {} : '',
      position: { x: 1, y: 2 },
      size: { width: 3, height: 2 },
      style: {
        fontSize: 18,
        color: '#444444'
      }
    }

    setSlides(prevSlides =>
      prevSlides.map((slide, index) =>
        index === selectedSlide
          ? { ...slide, elements: [...slide.elements, newElement] }
          : slide
      )
    )
  }

  // Update element content
  const updateElementContent = (elementId: string, content: any) => {
    setSlides(prevSlides =>
      prevSlides.map(slide => ({
        ...slide,
        elements: slide.elements.map(element =>
          element.id === elementId 
            ? { ...element, content }
            : element
        )
      }))
    )
  }

  // Download edited content as text file
  const downloadEditedContent = () => {
    const content = slides.map((slide, index) => {
      const slideContent = slide.elements.map(element => element.content).join('\n')
      return `--- Slide ${index + 1} ---\n${slideContent}\n`
    }).join('\n')

    const blob = new Blob([content], { type: 'text/plain' })
    const url = URL.createObjectURL(blob)
    const a = document.createElement('a')
    a.href = url
    a.download = 'edited-presentation-content.txt'
    document.body.appendChild(a)
    a.click()
    document.body.removeChild(a)
    URL.revokeObjectURL(url)
  }

  // Get slide background style
  const getSlideBackgroundStyle = (slide: Slide) => {
    if (originalTemplate.length > 0) {
      // Use original template background
      const originalSlide = originalTemplate.find(s => s.id === slide.id)
      if (originalSlide) {
        const bg = originalSlide.background
        if (bg.type === 'gradient' && bg.gradient) {
          return {
            background: `linear-gradient(${bg.gradient.angle || 90}deg, ${bg.gradient.colors.map(c => `${c.color} ${c.position}%`).join(', ')})`
          }
        } else if (bg.type === 'solid' && bg.color) {
          return { backgroundColor: bg.color }
        } else if (bg.type === 'image' && bg.image) {
          return { 
            backgroundImage: `url(${bg.image})`,
            backgroundSize: 'cover',
            backgroundPosition: 'center'
          }
        }
      }
    }
    
    // Fallback to gradient background for better visual appeal
    return {
      background: 'linear-gradient(135deg, #667eea 0%, #764ba2 100%)'
    }
  }

  // Convert PPTX slides to internal format
  const convertPPTXSlidesToInternalFormat = async (pptxSlides: RealPPTXSlide[]): Promise<Slide[]> => {
    return pptxSlides.map((pptxSlide, index) => ({
      id: pptxSlide.id,
      layout: pptxSlide.layout.type,
      background: pptxSlide.background.type,
      elements: pptxSlide.elements.map((pptxElement, elementIndex) => {
        // Convert PPTX element to internal format
        const internalElement: SlideElement = {
          id: pptxElement.id,
          type: pptxElement.type,
          content: pptxElement.content,
          position: {
            x: pptxElement.position.x / 914400, // Convert EMU to percentage (914400 EMU = 1 inch)
            y: pptxElement.position.y / 914400
          },
          size: {
            width: pptxElement.size.width / 914400,
            height: pptxElement.size.height / 914400
          },
          style: {
            fontSize: pptxElement.style.fontSize || 18,
            bold: pptxElement.style.fontWeight === 'bold',
            italic: pptxElement.style.fontStyle === 'italic',
            align: pptxElement.style.textAlign || 'left',
            color: pptxElement.style.color || '#000000',
            fontFamily: pptxElement.style.fontFamily || 'Arial, sans-serif',
            backgroundColor: pptxElement.style.backgroundColor,
            borderColor: pptxElement.style.borderColor,
            borderWidth: pptxElement.style.borderWidth,
            borderRadius: pptxElement.style.borderRadius,
            lineHeight: pptxElement.style.lineHeight,
            letterSpacing: pptxElement.style.letterSpacing,
            textShadow: pptxElement.style.textShadow,
            boxShadow: pptxElement.style.boxShadow
          }
        }

        // Special handling for agenda tables
        if (pptxElement.type === 'table' && typeof pptxElement.content === 'object') {
          const tableContent = pptxElement.content as any
          if (tableContent.rows && tableContent.rows.length > 0) {
            // Check if this is an agenda table that needs processing
            const firstRow = tableContent.rows[0]
            if (firstRow.length > 0 && typeof firstRow[0] === 'string') {
              const agendaText = firstRow[0]
              if (agendaText.includes('{agenda}')) {
                // Process agenda text into table format
                const agendaItems = agendaText.replace('{agenda}', '').trim()
                if (agendaItems) {
                  // Better agenda parsing - split by capital letters, numbers, or common separators
                  const items = agendaItems
                    .split(/(?=[A-Z][a-z])|(?=\d+\.)|\s+(?=and\s+)|\s+(?=or\s+)|,\s*/)
                    .filter(item => item.trim() && item.length > 2)
                    .map(item => item.trim().replace(/^[,\s]+|[,\s]+$/g, ''))
                  
                  // If we still have very few items, try alternative splitting
                  if (items.length < 2) {
                    const altItems = agendaItems.split(/\s+/).filter(item => item.trim())
                    if (altItems.length > items.length) {
                      items.length = 0
                      items.push(...altItems)
                    }
                  }
                  
                  internalElement.content = {
                    headers: tableContent.headers || ['#', 'Topic'],
                    rows: items.map((item, idx) => [`${idx + 1}.`, item]),
                    style: tableContent.style || {
                      headerBackground: '#2E74B5',
                      headerColor: '#FFFFFF',
                      borderColor: '#D9D9D9',
                      rowBackground: '#FFFFFF',
                      alternateRowBackground: '#F2F2F2',
                      fontSize: 18,
                      fontFamily: 'Calibri'
                    }
                  }
                }
              }
            }
          }
        }

        return internalElement
      })
    }))
  }

  // Auto-generate slides when template or data changes
  useEffect(() => {
    if (template && Object.keys(data).length > 0) {
      generateSlides()
    } else if (template && Object.keys(data).length === 0) {
      // Generate with empty data if template exists but no data
      generateSlides()
    }
  }, [template, data, generateSlides])

  return (
    <div className="w-full h-full flex flex-col min-h-0">
      {/* Header */}
      <Card className="mb-4">
        <CardHeader className="pb-3">
          <div className="flex items-center justify-between">
            <CardTitle className="flex items-center gap-2">
              <Eye className="w-5 h-5" />
              Real-time Preview
              <Badge variant={isProcessing ? "secondary" : "default"}>
                {isProcessing ? 'Processing...' : 'Live'}
              </Badge>
            </CardTitle>
            
            <div className="flex items-center gap-2">
              {/* Zoom Controls */}
              <div className="flex items-center gap-1 border rounded-md">
                <Button
                  variant="ghost"
                  size="sm"
                  onClick={() => setZoom(Math.max(50, zoom - 10))}
                >
                  <ZoomOut className="w-4 h-4" />
                </Button>
                <span className="px-2 text-sm font-medium">{zoom}%</span>
                <Button
                  variant="ghost"
                  size="sm"
                  onClick={() => setZoom(Math.min(200, zoom + 10))}
                >
                  <ZoomIn className="w-4 h-4" />
                </Button>
              </div>
            </div>
          </div>
        </CardHeader>
        
        {/* Processing Progress */}
        {isProcessing && (
          <CardContent className="pt-0">
            <Progress value={processingProgress} className="w-full" />
            <p className="text-sm text-gray-600 mt-2">
              {processingProgress < 20 && 'Initializing...'}
              {processingProgress >= 20 && processingProgress < 40 && 'Processing template...'}
              {processingProgress >= 40 && processingProgress < 60 && 'Generating slides...'}
              {processingProgress >= 60 && processingProgress < 80 && 'Adding interactivity...'}
              {processingProgress >= 80 && processingProgress < 100 && 'Finalizing...'}
              {processingProgress === 100 && 'Complete!'}
            </p>
          </CardContent>
        )}
      </Card>

      {/* Main Content */}
      <div className="flex-1 flex gap-4 min-h-0">
        {/* Sidebar - Slide Navigation & Tools */}
        <div className="w-64 space-y-4 flex-shrink-0">
          {/* Slide Navigation */}
          <Card>
            <CardHeader className="pb-3">
              <CardTitle className="text-sm">Slides ({slides.length})</CardTitle>
            </CardHeader>
            <CardContent className="space-y-2">
              {slides.map((slide, index) => (
                <div
                  key={slide.id}
                  className={`p-2 rounded cursor-pointer transition-colors ${
                    selectedSlide === index 
                      ? 'bg-blue-100 border-blue-300 border' 
                      : 'hover:bg-gray-100'
                  }`}
                  onClick={() => setSelectedSlide(index)}
                >
                  <div className="flex items-center justify-between">
                    <span className="text-sm font-medium">Slide {index + 1}</span>
                    <Badge variant="outline" className="text-xs">
                      {slide.elements.length} elements
                    </Badge>
                  </div>
                </div>
              ))}
            </CardContent>
          </Card>

          {/* Editing Tools */}
          {isEditing && (
            <Card>
              <CardHeader className="pb-3">
                <CardTitle className="text-sm">Tools</CardTitle>
              </CardHeader>
              <CardContent className="space-y-2">
                <div className="grid grid-cols-2 gap-2">
                  <Button
                    variant={editMode === 'select' ? 'default' : 'outline'}
                    size="sm"
                    onClick={() => setEditMode('select')}
                  >
                    <Move className="w-4 h-4" />
                  </Button>
                  <Button
                    variant={editMode === 'text' ? 'default' : 'outline'}
                    size="sm"
                    onClick={() => {
                      setEditMode('text')
                      addElement('text')
                    }}
                  >
                    <Type className="w-4 h-4" />
                  </Button>
                  <Button
                    variant={editMode === 'chart' ? 'default' : 'outline'}
                    size="sm"
                    onClick={() => {
                      setEditMode('chart')
                      addElement('chart')
                    }}
                  >
                    <BarChart3 className="w-4 h-4" />
                  </Button>
                  <Button
                    variant={editMode === 'image' ? 'default' : 'outline'}
                    size="sm"
                    onClick={() => {
                      setEditMode('image')
                      addElement('image')
                    }}
                  >
                    <ImageIcon className="w-4 h-4" />
                  </Button>
                </div>
              </CardContent>
            </Card>
          )}
        </div>

        {/* Preview Area */}
        <Card className="flex-1 min-h-0">
          <CardContent className="p-6 h-full overflow-auto">
            {slides.length > 0 ? (
              <div 
                className="w-full h-full flex items-center justify-center min-h-[500px]"
                style={{ transform: `scale(${zoom / 100})` }}
              >
                <div 
                  className="relative shadow-lg border border-gray-200 rounded-lg overflow-hidden"
                  style={{
                    width: '720px',
                    height: '405px', // 16:9 ratio
                    minHeight: '405px',
                    ...getSlideBackgroundStyle(slides[selectedSlide])
                  }}
                >
                  {slides[selectedSlide]?.elements.map(element => (
                    <div
                      key={element.id}
                      className={`absolute cursor-pointer transition-all ${
                        selectedElement === element.id ? 'ring-2 ring-blue-500' : ''
                      } ${isEditing ? 'hover:ring-2 hover:ring-gray-400' : ''}`}
                      style={{
                        left: `${element.position.x * 80}px`,
                        top: `${element.position.y * 80}px`,
                        width: `${element.size.width * 80}px`,
                        height: `${element.size.height * 80}px`,
                        fontFamily: element.style?.fontFamily || 'Arial, sans-serif',
                        fontSize: `${element.style?.fontSize || 16}px`,
                        fontWeight: element.style?.bold ? 'bold' : 'normal',
                        fontStyle: element.style?.italic ? 'italic' : 'normal',
                        textAlign: element.style?.align || 'left',
                        color: element.style?.color || '#000000',
                        backgroundColor: element.style?.backgroundColor || 'transparent',
                        border: element.style?.borderWidth 
                          ? `${element.style.borderWidth}px solid ${element.style?.borderColor || '#000000'}`
                          : 'none',
                        borderRadius: element.style?.borderRadius ? `${element.style.borderRadius}px` : '0',
                        lineHeight: element.style?.lineHeight || 1.2,
                        letterSpacing: element.style?.letterSpacing || 'normal',
                        textShadow: element.style?.textShadow || 'none',
                        boxShadow: element.style?.boxShadow || 'none'
                      }}
                      onClick={() => isEditing && handleElementClick(element.id)}
                    >
                      {element.type === 'text' && (
                        <div
                          contentEditable={isEditing && selectedElement === element.id}
                          onBlur={(e) => updateElementContent(element.id, e.target.textContent)}
                          suppressContentEditableWarning
                          className="w-full h-full outline-none"
                          style={{
                            fontFamily: 'inherit',
                            fontSize: 'inherit',
                            fontWeight: 'inherit',
                            fontStyle: 'inherit',
                            textAlign: 'inherit',
                            color: 'inherit',
                            lineHeight: 'inherit',
                            letterSpacing: 'inherit',
                            textShadow: 'inherit',
                            display: 'flex',
                            alignItems: 'center',
                            justifyContent: element.style?.align === 'center' ? 'center' : 
                                           element.style?.align === 'right' ? 'flex-end' : 'flex-start',
                            padding: '4px'
                          }}
                        >
                          {element.content}
                        </div>
                      )}
                      {element.type === 'chart' && (
                        <div 
                          className="w-full h-full bg-gray-50 rounded-lg flex flex-col items-center justify-center border-2 border-dashed border-gray-300"
                          style={{
                            backgroundColor: element.style?.backgroundColor || '#f8fafc',
                            border: element.style?.border || '2px dashed #cbd5e1'
                          }}
                        >
                          <BarChart3 className="w-12 h-12 text-gray-400 mb-2" />
                          <p className="text-xs text-gray-500 text-center px-2">
                            {element.content?.title || 'Chart Placeholder'}
                          </p>
                          {element.content?.type === 'placeholder' && (
                            <p className="text-xs text-blue-500 mt-1">Auto-generated from data</p>
                          )}
                        </div>
                      )}
                      {element.type === 'image' && (
                        <div className="w-full h-full bg-gray-100 rounded flex items-center justify-center">
                          <ImageIcon className="w-8 h-8 text-gray-400" />
                        </div>
                      )}
                      {element.type === 'table' && (
                        <div className="w-full h-full overflow-hidden">
                          <table 
                            className="w-full h-full border-collapse" 
                            style={{ 
                              fontSize: `${element.content?.style?.fontSize || 14}px`,
                              fontFamily: element.content?.style?.fontFamily || 'Arial, sans-serif'
                            }}
                          >
                            <thead>
                              <tr style={{ 
                                backgroundColor: element.content?.style?.headerBackground || '#2E74B5',
                                color: element.content?.style?.headerColor || '#FFFFFF'
                              }}>
                                {element.content?.headers?.map((header: string, index: number) => (
                                  <th
                                    key={index}
                                    className="border font-semibold text-left px-2 py-1"
                                    style={{
                                      borderColor: element.content?.style?.borderColor || '#D9D9D9',
                                      fontSize: 'inherit',
                                      fontFamily: 'inherit',
                                      fontWeight: 'bold'
                                    }}
                                  >
                                    {header}
                                  </th>
                                ))}
                              </tr>
                            </thead>
                            <tbody>
                              {element.content?.rows?.map((row: string[], rowIndex: number) => (
                                <tr
                                  key={rowIndex}
                                  style={{
                                    backgroundColor: rowIndex % 2 === 0 
                                      ? element.content?.style?.rowBackground || '#FFFFFF'
                                      : element.content?.style?.alternateRowBackground || '#F2F2F2',
                                    fontSize: 'inherit',
                                    fontFamily: 'inherit'
                                  }}
                                >
                                  {row.map((cell: string, cellIndex: number) => (
                                    <td
                                      key={cellIndex}
                                      className="border px-2 py-1"
                                      style={{
                                        borderColor: element.content?.style?.borderColor || '#D9D9D9',
                                        fontSize: 'inherit',
                                        fontFamily: 'inherit',
                                        color: element.style?.color || '#000000'
                                      }}
                                    >
                                      {cell}
                                    </td>
                                  ))}
                                </tr>
                              ))}
                            </tbody>
                          </table>
                        </div>
                      )}
                    </div>
                  ))}
                </div>
              </div>
            ) : (
              <div className="w-full h-full flex items-center justify-center min-h-[400px]">
                <div className="text-center">
                  <Eye className="w-16 h-16 text-gray-300 mx-auto mb-4" />
                  <p className="text-gray-500">No slides to preview</p>
                  <p className="text-sm text-gray-400">Upload template and data to generate slides</p>
                </div>
              </div>
            )}
          </CardContent>
        </Card>
      </div>
    </div>
  )
}