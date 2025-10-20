import { NextRequest, NextResponse } from 'next/server'
import { writeFile, readFile, unlink } from 'fs/promises'
import { join } from 'path'
import { existsSync, mkdirSync } from 'fs'
import * as XLSX from 'xlsx'
import PptxGenJS from 'pptxgenjs'

// Ensure uploads directory exists
const uploadsDir = join(process.cwd(), 'uploads')
if (!existsSync(uploadsDir)) {
  mkdirSync(uploadsDir, { recursive: true })
}

export async function POST(request: NextRequest) {
  try {
    const formData = await request.formData()
    const pptFile = formData.get('ppt') as File
    const excelFile = formData.get('excel') as File

    if (!pptFile || !excelFile) {
      return NextResponse.json(
        { error: 'Both PowerPoint and Excel files are required' },
        { status: 400 }
      )
    }

    // Validate file types
    if (!pptFile.name.endsWith('.pptx') || !excelFile.name.endsWith('.xlsx')) {
      return NextResponse.json(
        { error: 'Invalid file types. Please upload .pptx and .xlsx files' },
        { status: 400 }
      )
    }

    // Save uploaded files
    const pptPath = join(uploadsDir, pptFile.name)
    const excelPath = join(uploadsDir, excelFile.name)

    await writeFile(pptPath, Buffer.from(await pptFile.arrayBuffer()))
    await writeFile(excelPath, Buffer.from(await excelFile.arrayBuffer()))

    // Read Excel data
    const excelBuffer = await readFile(excelPath)
    const workbook = XLSX.read(excelBuffer, { type: 'buffer' })
    
    // Get the main sheet data
    const sheetName = workbook.SheetNames[0]
    const worksheet = workbook.Sheets[sheetName]
    const excelData = XLSX.utils.sheet_to_json(worksheet)

    if (excelData.length === 0) {
      return NextResponse.json(
        { error: 'Excel file is empty or invalid' },
        { status: 400 }
      )
    }

    // Get the first row of data for replacement
    const dataRow = excelData[0] as Record<string, any>
    
    // Also try to read ChartData sheet if it exists
    let chartData = []
    
    if (workbook.SheetNames.includes('ChartData')) {
      const chartSheet = workbook.Sheets['ChartData']
      chartData = XLSX.utils.sheet_to_json(chartSheet)
    }
    
    // Add chart data to the main data row
    if (chartData.length > 0) {
      dataRow.ChartData = chartData
    }

    // Create a new presentation with PPTXGenJS
    const processedBuffer = await createPresentationWithCharts(dataRow)

    // Return processed file
    return new NextResponse(processedBuffer, {
      headers: {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
        'Content-Disposition': `attachment; filename="processed_${pptFile.name}"`,
      },
    })

  } catch (error) {
    console.error('Processing error:', error)
    return NextResponse.json(
      { error: 'Failed to process files: ' + (error as Error).message },
      { status: 500 }
    )
  }
}

async function createPresentationWithCharts(dataRow: Record<string, any>): Promise<Buffer> {
  try {
    // Create a new presentation
    const pptx = new PptxGenJS()
    
    // Set presentation properties
    pptx.author = 'PowerPoint AutoFill'
    pptx.company = 'AutoFill System'
    pptx.title = dataRow.project_title || 'AutoFill Presentation'
    
    // Use default layout (16:9)
    
    // Helper function to replace placeholders
    const replacePlaceholders = (text: string): string => {
      if (!text) return text
      const defaultValues = {
        date: new Date().toLocaleDateString(),
        presenter: 'Presenter',
        company: 'Company Name'
      }
      
      return text.replace(/\{([^}]+)\}/g, (match, key) => {
        const trimmedKey = key.trim()
        const value = dataRow[trimmedKey] !== undefined ? dataRow[trimmedKey] : defaultValues[trimmedKey as keyof typeof defaultValues]
        return value !== undefined ? String(value) : match
      })
    }
    
    // Slide 1: Title Slide
    const slide1 = pptx.addSlide()
    slide1.addText(
      replacePlaceholders('{project_title}'),
      { x: 0.5, y: 2, w: 9, h: 1.5, fontSize: 44, bold: true, color: '363636', align: 'center' }
    )
    
    slide1.addText(
      replacePlaceholders('Status Report prepared by: {name} - {date}'),
      { x: 0.5, y: 3.5, w: 9, h: 1, fontSize: 24, color: '666666', align: 'center' }
    )
    
    // Slide 2: Overview
    const slide2 = pptx.addSlide()
    slide2.addText(
      'Overview',
      { x: 0.5, y: 0.5, w: 9, h: 1, fontSize: 36, bold: true, color: '363636' }
    )
    
    slide2.addText(
      replacePlaceholders('{status_update}'),
      { x: 0.5, y: 1.5, w: 9, h: 4, fontSize: 18, color: '444444', valign: 'top', wrapText: true }
    )
    
    // Slide 3: Agenda
    const slide3 = pptx.addSlide()
    slide3.addText(
      'Agenda',
      { x: 0.5, y: 0.5, w: 9, h: 1, fontSize: 36, bold: true, color: '363636' }
    )
    
    slide3.addText(
      replacePlaceholders('{agenda}'),
      { x: 0.5, y: 1.5, w: 9, h: 4, fontSize: 18, color: '444444', valign: 'top', wrapText: true }
    )
    
    // Slide 4: Chart Title Placeholder
    const slide4 = pptx.addSlide()
    slide4.addText(
      replacePlaceholders('{ChartTitle}'),
      { x: 0.5, y: 0.5, w: 9, h: 1, fontSize: 36, bold: true, color: '363636' }
    )
    
    slide4.addText(
      'Chart data will be displayed in the following slides based on your Excel data.',
      { x: 0.5, y: 1.5, w: 9, h: 1, fontSize: 16, color: '666666' }
    )
    
    // Process chart data and create chart slides
    if (dataRow.ChartData && Array.isArray(dataRow.ChartData)) {
      const chartGroups = groupChartData(dataRow.ChartData)
      
      Object.entries(chartGroups).forEach(([chartTitle, chartGroup]) => {
        const chartSlide = pptx.addSlide()
        chartSlide.addText(
          chartTitle,
          { x: 0.5, y: 0.5, w: 9, h: 1, fontSize: 36, bold: true, color: '363636' }
        )
        
        // Create chart based on chart type
        const chartType = chartGroup[0].ChartType?.toLowerCase()
        const chartData = formatChartData(chartGroup)
        
        if (chartType === 'bar' || chartType === 'column') {
          chartSlide.addChart(pptx.charts.BAR, chartData, {
            x: 0.5, y: 1.5, w: 9, h: 5,
            barDir: 'col',
            catAxisLabelRotate: 45,
            showLegend: true,
            legendPos: 'b'
          })
        } else if (chartType === 'line') {
          chartSlide.addChart(pptx.charts.LINE, chartData, {
            x: 0.5, y: 1.5, w: 9, h: 5,
            lineDataSymbol: 'circle',
            showLegend: true,
            legendPos: 'b'
          })
        } else if (chartType === 'pie') {
          chartSlide.addChart(pptx.charts.PIE, chartData, {
            x: 0.5, y: 1.5, w: 9, h: 5,
            showLegend: true,
            legendPos: 'b',
            dataLabelFormatCode: '#,##0.0%'
          })
        } else {
          // Default to bar chart if type is not recognized
          chartSlide.addChart(pptx.charts.BAR, chartData, {
            x: 0.5, y: 1.5, w: 9, h: 5,
            barDir: 'col',
            showLegend: true,
            legendPos: 'b'
          })
        }
      })
    }
    
    // Generate the presentation using the working method
    // Create temp file first, then read as buffer (Method 3 from our test)
    const tempPath = join(process.cwd(), 'uploads', `temp_${Date.now()}.pptx`)
    await pptx.writeFile({ fileName: tempPath })
    
    // Read the file as buffer
    const buffer = await readFile(tempPath)
    
    // Clean up temp file
    try {
      await unlink(tempPath)
    } catch (error) {
      // Ignore cleanup errors
    }
    
    return buffer
    
  } catch (error) {
    console.error('Error creating presentation:', error)
    throw new Error('Failed to create presentation: ' + (error as Error).message)
  }
}

// Group chart data by ChartTitle
function groupChartData(chartData: any[]): Record<string, any[]> {
  const grouped: Record<string, any[]> = {}
  
  chartData.forEach(item => {
    const title = item.ChartTitle || 'Chart'
    if (!grouped[title]) {
      grouped[title] = []
    }
    grouped[title].push(item)
  })
  
  return grouped
}

// Format chart data for PPTXGenJS
function formatChartData(chartGroup: any[]): any[] {
  const labels: string[] = []
  const dataPoints: number[] = []
  const series: string[] = []
  
  // Group by series if multiple series exist
  const seriesGroups: Record<string, any[]> = {}
  
  chartGroup.forEach(item => {
    const seriesName = item.Series || 'Data'
    if (!seriesGroups[seriesName]) {
      seriesGroups[seriesName] = []
    }
    seriesGroups[seriesName].push(item)
  })
  
  // Format for PPTXGenJS
  const chartData = []
  
  Object.entries(seriesGroups).forEach(([seriesName, seriesData]) => {
    const labels = seriesData.map(item => item.Label || '')
    const values = seriesData.map(item => Number(item.Value) || 0)
    
    chartData.push({
      name: seriesName,
      labels: labels,
      values: values
    })
  })
  
  return chartData
}