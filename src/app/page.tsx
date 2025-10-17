'use client'

import { useState } from 'react'
import { Button } from '@/components/ui/button'
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from '@/components/ui/card'
import { Badge } from '@/components/ui/badge'
import { Progress } from '@/components/ui/progress'
import { 
  Upload, 
  FileText, 
  FileSpreadsheet, 
  Download, 
  CheckCircle,
  AlertCircle,
  ArrowRight,
  Zap
} from 'lucide-react'

export default function Home() {
  const [pptFile, setPptFile] = useState<File | null>(null)
  const [excelFile, setExcelFile] = useState<File | null>(null)
  const [isProcessing, setIsProcessing] = useState(false)
  const [processedFile, setProcessedFile] = useState<string | null>(null)
  const [progress, setProgress] = useState(0)
  const [status, setStatus] = useState<'idle' | 'processing' | 'completed' | 'error'>('idle')

  const handlePptUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0]
    if (file && (file.type === 'application/vnd.openxmlformats-officedocument.presentationml.presentation' || file.name.endsWith('.pptx'))) {
      setPptFile(file)
      setStatus('idle')
    }
  }

  const handleExcelUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0]
    if (file && (file.type === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' || file.name.endsWith('.xlsx'))) {
      setExcelFile(file)
      setStatus('idle')
    }
  }

  const processFiles = async () => {
    if (!pptFile || !excelFile) return

    setIsProcessing(true)
    setStatus('processing')
    setProgress(0)

    try {
      const formData = new FormData()
      formData.append('ppt', pptFile)
      formData.append('excel', excelFile)

      // Simulate progress
      const progressInterval = setInterval(() => {
        setProgress(prev => {
          if (prev >= 90) {
            clearInterval(progressInterval)
            return 90
          }
          return prev + 10
        })
      }, 200)

      const response = await fetch('/api/process', {
        method: 'POST',
        body: formData
      })

      clearInterval(progressInterval)
      setProgress(100)

      if (response.ok) {
        const blob = await response.blob()
        const url = URL.createObjectURL(blob)
        setProcessedFile(url)
        setStatus('completed')
      } else {
        setStatus('error')
      }
    } catch (error) {
      setStatus('error')
    } finally {
      setIsProcessing(false)
    }
  }

  const downloadProcessedFile = () => {
    if (processedFile) {
      const a = document.createElement('a')
      a.href = processedFile
      a.download = `processed_${pptFile?.name}`
      document.body.appendChild(a)
      a.click()
      document.body.removeChild(a)
    }
  }

  const reset = () => {
    setPptFile(null)
    setExcelFile(null)
    setProcessedFile(null)
    setProgress(0)
    setStatus('idle')
  }

  return (
    <div className="min-h-screen bg-gradient-to-br from-slate-50 to-slate-100">
      {/* Header */}
      <header className="border-b bg-white/80 backdrop-blur-sm">
        <div className="container mx-auto px-4 py-4 flex items-center justify-between">
          <div className="flex items-center gap-2">
            <FileText className="h-8 w-8 text-primary" />
            <span className="text-2xl font-bold">PPT AutoFill</span>
          </div>
          <Badge variant="secondary">
            <Zap className="w-3 h-3 mr-1" />
            AI-Powered
          </Badge>
        </div>
      </header>

      {/* Main Content */}
      <main className="container mx-auto px-4 py-12">
        <div className="max-w-4xl mx-auto">
          {/* Hero Section */}
          <div className="text-center mb-12">
            <h1 className="text-4xl font-bold mb-4 bg-gradient-to-r from-slate-900 to-slate-600 bg-clip-text text-transparent">
              Auto-Fill PowerPoint Templates
            </h1>
            <p className="text-xl text-muted-foreground mb-8">
              Upload your PowerPoint template with {'{placeholder}'} tags and Excel data. 
              We'll extract your template, replace placeholders with matching column values, and generate a new PowerPoint file.
            </p>
            <div className="flex items-center justify-center gap-4">
              <Button variant="outline" onClick={() => window.location.href = '/samples'}>
                <Download className="w-4 h-4 mr-2" />
                Download Sample Files
              </Button>
            </div>
          </div>

          {/* Upload Section */}
          <Card className="mb-8">
            <CardHeader>
              <CardTitle className="flex items-center gap-2">
                <Upload className="h-5 w-5" />
                Upload Files
              </CardTitle>
              <CardDescription>
                Upload your PowerPoint template and Excel data file
              </CardDescription>
            </CardHeader>
            <CardContent className="space-y-6">
              <div className="grid md:grid-cols-2 gap-6">
                {/* PPT Upload */}
                <div className="border-2 border-dashed border-gray-300 rounded-lg p-6 text-center hover:border-primary transition-colors">
                  <input
                    type="file"
                    accept=".pptx"
                    onChange={handlePptUpload}
                    className="hidden"
                    id="ppt-upload"
                  />
                  <label htmlFor="ppt-upload" className="cursor-pointer">
                    <FileText className="h-12 w-12 mx-auto mb-4 text-gray-400" />
                    <p className="font-medium mb-2">PowerPoint Template</p>
                    <p className="text-sm text-gray-500 mb-4">
                      Upload .pptx file with {'{placeholder}'} tags (e.g., {'{name}'}, {'{project_title}'})
                    </p>
                    {pptFile ? (
                      <div className="flex items-center justify-center gap-2 text-green-600">
                        <CheckCircle className="h-4 w-4" />
                        <span className="text-sm">{pptFile.name}</span>
                      </div>
                    ) : (
                      <Button variant="outline" size="sm">
                        Choose File
                      </Button>
                    )}
                  </label>
                </div>

                {/* Excel Upload */}
                <div className="border-2 border-dashed border-gray-300 rounded-lg p-6 text-center hover:border-primary transition-colors">
                  <input
                    type="file"
                    accept=".xlsx"
                    onChange={handleExcelUpload}
                    className="hidden"
                    id="excel-upload"
                  />
                  <label htmlFor="excel-upload" className="cursor-pointer">
                    <FileSpreadsheet className="h-12 w-12 mx-auto mb-4 text-gray-400" />
                    <p className="font-medium mb-2">Excel Data</p>
                    <p className="text-sm text-gray-500 mb-4">
                      Upload .xlsx file with columns matching placeholder names
                    </p>
                    {excelFile ? (
                      <div className="flex items-center justify-center gap-2 text-green-600">
                        <CheckCircle className="h-4 w-4" />
                        <span className="text-sm">{excelFile.name}</span>
                      </div>
                    ) : (
                      <Button variant="outline" size="sm">
                        Choose File
                      </Button>
                    )}
                  </label>
                </div>
              </div>

              {/* Process Button */}
              <div className="flex justify-center">
                <Button
                  size="lg"
                  onClick={processFiles}
                  disabled={!pptFile || !excelFile || isProcessing}
                  className="px-8"
                >
                  {isProcessing ? 'Processing...' : 'Process Files'}
                  <ArrowRight className="w-5 h-5 ml-2" />
                </Button>
              </div>
            </CardContent>
          </Card>

          {/* Progress Section */}
          {status === 'processing' && (
            <Card className="mb-8">
              <CardContent className="p-6">
                <div className="flex items-center justify-between mb-2">
                  <span className="text-sm font-medium">Processing files...</span>
                  <span className="text-sm text-gray-500">{progress}%</span>
                </div>
                <Progress value={progress} className="w-full" />
              </CardContent>
            </Card>
          )}

          {/* Error Section */}
          {status === 'error' && (
            <Card className="mb-8 border-red-200 bg-red-50">
              <CardContent className="p-6">
                <div className="flex items-center gap-3 text-red-600">
                  <AlertCircle className="h-5 w-5" />
                  <div>
                    <p className="font-medium">Processing Failed</p>
                    <p className="text-sm text-red-500">
                      Please check your files and try again. Make sure your PowerPoint contains {'{placeholder}'} tags and Excel has matching column names.
                    </p>
                  </div>
                </div>
                <Button variant="outline" size="sm" className="mt-4" onClick={reset}>
                  Try Again
                </Button>
              </CardContent>
            </Card>
          )}

          {/* Success Section */}
          {status === 'completed' && processedFile && (
            <Card className="mb-8 border-green-200 bg-green-50">
              <CardContent className="p-6">
                <div className="flex items-center justify-between">
                  <div className="flex items-center gap-3 text-green-600">
                    <CheckCircle className="h-5 w-5" />
                    <div>
                      <p className="font-medium">Processing Complete!</p>
                      <p className="text-sm text-green-500">
                        Your PowerPoint has been updated with data from Excel
                      </p>
                    </div>
                  </div>
                  <Button onClick={downloadProcessedFile}>
                    <Download className="w-4 h-4 mr-2" />
                    Download
                  </Button>
                </div>
              </CardContent>
            </Card>
          )}

          {/* Instructions */}
          <Card>
            <CardHeader>
              <CardTitle>How It Works</CardTitle>
            </CardHeader>
            <CardContent>
              <div className="space-y-4 text-sm">
                <div className="flex items-start gap-3">
                  <div className="w-6 h-6 rounded-full bg-primary text-white flex items-center justify-center text-xs font-bold">1</div>
                  <div>
                    <p className="font-medium">Create PowerPoint Template</p>
                    <p className="text-gray-600">Add placeholders like {'{name}'}, {'{company}'}, {'{date}'} in your slides</p>
                  </div>
                </div>
                <div className="flex items-start gap-3">
                  <div className="w-6 h-6 rounded-full bg-primary text-white flex items-center justify-center text-xs font-bold">2</div>
                  <div>
                    <p className="font-medium">Prepare Excel Data</p>
                    <p className="text-gray-600">Create columns with matching names (name, company, date, etc.)</p>
                  </div>
                </div>
                <div className="flex items-start gap-3">
                  <div className="w-6 h-6 rounded-full bg-primary text-white flex items-center justify-center text-xs font-bold">3</div>
                  <div>
                    <p className="font-medium">Upload & Process</p>
                    <p className="text-gray-600">We'll automatically replace all placeholders with Excel data</p>
                  </div>
                </div>
                <div className="flex items-start gap-3">
                  <div className="w-6 h-6 rounded-full bg-primary text-white flex items-center justify-center text-xs font-bold">4</div>
                  <div>
                    <p className="font-medium">Download Result</p>
                    <p className="text-gray-600">Get your updated PowerPoint with all data filled in</p>
                  </div>
                </div>
              </div>
            </CardContent>
          </Card>
        </div>
      </main>
    </div>
  )
}