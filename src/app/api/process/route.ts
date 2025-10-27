import { NextRequest, NextResponse } from 'next/server'
import { writeFile, readFile, unlink } from 'fs/promises'
import { join } from 'path'
import { existsSync, mkdirSync } from 'fs'
import * as XLSX from 'xlsx'
import { SafePPTXBuilder } from '@/lib/safe-pptx-builder'
import { UltraSafePPTXBuilder } from '@/lib/ultra-safe-pptx-builder'
import { PPTXDiagnostic } from '@/lib/pptx-diagnostic'

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

    // Create PPTX builder that preserves original template design
    let processedBuffer: ArrayBuffer;
    
    try {
      // Try the safe builder first
      const pptxBuilder = new SafePPTXBuilder()
      await pptxBuilder.loadTemplate(Buffer.from(await pptFile.arrayBuffer()))
      pptxBuilder.setData(dataRow)
      
      // Build the final presentation with charts if they exist
      processedBuffer = await pptxBuilder.buildWithCharts(chartData)
      
      // Validate the result
      const diagnostic = await PPTXDiagnostic.analyze(processedBuffer)
      if (!diagnostic.isValid) {
        console.warn('SafePPTXBuilder produced invalid file, falling back to ultra-safe mode')
        console.warn('Issues:', diagnostic.issues)
        
        // Fallback to ultra-safe builder
        const ultraSafeBuilder = new UltraSafePPTXBuilder()
        await ultraSafeBuilder.loadTemplate(Buffer.from(await pptFile.arrayBuffer()))
        ultraSafeBuilder.setData(dataRow)
        processedBuffer = await ultraSafeBuilder.buildWithCharts(chartData)
        
        // Validate the fallback result
        const fallbackDiagnostic = await PPTXDiagnostic.analyze(processedBuffer)
        console.log('Ultra-safe builder diagnostic:', PPTXDiagnostic.generateReport(fallbackDiagnostic))
      }
    } catch (error) {
      console.error('SafePPTXBuilder failed, using ultra-safe fallback:', error)
      
      // Fallback to ultra-safe builder
      const ultraSafeBuilder = new UltraSafePPTXBuilder()
      await ultraSafeBuilder.loadTemplate(Buffer.from(await pptFile.arrayBuffer()))
      ultraSafeBuilder.setData(dataRow)
      processedBuffer = await ultraSafeBuilder.buildWithCharts(chartData)
    }

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