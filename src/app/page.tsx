'use client'

import { useState, useEffect } from 'react'
import { Button } from '@/components/ui/button'
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from '@/components/ui/card'
import { Badge } from '@/components/ui/badge'
import { Progress } from '@/components/ui/progress'
import { useToast } from '@/hooks/use-toast'
import { Toaster } from '@/components/ui/toaster'
import RealtimePreview from '@/components/realtime-preview'
import { 
  Upload, 
  FileText, 
  FileSpreadsheet, 
  Download, 
  CheckCircle,
  AlertCircle,
  ArrowRight,
  Zap,
  BarChart3,
  Clock,
  Shield,
  Eye,
  Settings,
  Play,
  Pause,
  Layers,
  Database,
  RefreshCw,
  X
} from 'lucide-react'
import { DynamicTemplateEngine, defaultTemplateConfig } from '@/lib/dynamic-templates'

export default function Home() {
  const [pptFile, setPptFile] = useState<File | null>(null)
  const [excelFile, setExcelFile] = useState<File | null>(null)
  const [isProcessing, setIsProcessing] = useState(false)
  const [processedFile, setProcessedFile] = useState<string | null>(null)
  const [progress, setProgress] = useState(0)
  const [status, setStatus] = useState<'idle' | 'processing' | 'completed' | 'error'>('idle')
  const [activeTab, setActiveTab] = useState('basic')
  const [template, setTemplate] = useState('')
  const [data, setData] = useState<Record<string, any>>({})
  const [isAutomationEnabled, setIsAutomationEnabled] = useState(false)
  const [showPreview, setShowPreview] = useState(false)
  const { toast } = useToast()

  // Initialize advanced systems
  const [templateEngine] = useState(() => new DynamicTemplateEngine(defaultTemplateConfig))

  const handlePptUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0]
    if (file && (file.type === 'application/vnd.openxmlformats-officedocument.presentationml.presentation' || file.name.endsWith('.pptx'))) {
      setPptFile(file)
      setStatus('idle')
      
      // Create a simple template for preview
      const sampleTemplate = `{project_title}
Status Report prepared by: {name}

Overview
{status_update}

Agenda
{agenda}

{ChartTitle}
This is a placeholder slide. The title above will be filled from your Excel data.`
      
      setTemplate(sampleTemplate)
      setShowPreview(true)
      
      toast({
        title: "PowerPoint uploaded",
        description: `${file.name} has been uploaded successfully.`,
      })
    } else if (file) {
      toast({
        title: "Invalid file type",
        description: "Please upload a .pptx file.",
        variant: "destructive",
      })
    }
  }

  const handleExcelUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0]
    if (file && (file.type === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' || file.name.endsWith('.xlsx'))) {
      setExcelFile(file)
      setStatus('idle')
      
      // Use sample data for preview based on the provided document
      const sampleData = {
        project_title: 'DGE Review',
        name: 'zinab',
        status_update: 'All financial models have been updated, and the initial report draft is complete. We are on schedule for the board meeting next week.',
        agenda: 'Technology Stack System Context Design Decisions Mobile Application Features',
        ChartTitle: 'Analysis of Key Metrics'
      }
      setData(sampleData)
      
      toast({
        title: "Excel file uploaded",
        description: `${file.name} has been uploaded successfully.`,
      })
    } else if (file) {
      toast({
        title: "Invalid file type",
        description: "Please upload a .xlsx file.",
        variant: "destructive",
      })
    }
  }

  const processFiles = async () => {
    if (!pptFile || !excelFile) {
      toast({
        title: "Files missing",
        description: "Please upload both PowerPoint and Excel files.",
        variant: "destructive",
      })
      return
    }

    setIsProcessing(true)
    setStatus('processing')
    setProgress(0)

    toast({
      title: "Processing started",
      description: "Your files are being processed...",
    })

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
        toast({
          title: "Processing complete!",
          description: "Your PowerPoint has been generated successfully.",
        })
      } else {
        setStatus('error')
        toast({
          title: "Processing failed",
          description: "There was an error processing your files. Please try again.",
          variant: "destructive",
        })
      }
    } catch (error) {
      setStatus('error')
      toast({
        title: "Processing failed",
        description: "Network error. Please check your connection and try again.",
        variant: "destructive",
      })
    } finally {
      setIsProcessing(false)
    }
  }

  const reset = () => {
    setPptFile(null)
    setExcelFile(null)
    setProcessedFile(null)
    setProgress(0)
    setStatus('idle')
    setTemplate('')
    setData({})
    setShowPreview(false)
  }

  const toggleAutomation = () => {
    setIsAutomationEnabled(!isAutomationEnabled)
    if (!isAutomationEnabled) {
      toast({
        title: "Automation Enabled",
        description: "Automatic processing and monitoring has been started.",
      })
    } else {
      toast({
        title: "Automation Disabled",
        description: "Automatic processing has been stopped.",
      })
    }
  }

  const processWithDynamicFeatures = async () => {
    if (!pptFile || !excelFile) {
      toast({
        title: "Files missing",
        description: "Please upload both PowerPoint and Excel files.",
        variant: "destructive",
      })
      return
    }

    setIsProcessing(true)
    setStatus('processing')
    setProgress(0)

    try {
      // Process with dynamic template engine
      const processedTemplate = templateEngine.processTemplate(template, data)
      setProgress(30)

      // Continue with normal processing
      const formData = new FormData()
      formData.append('ppt', pptFile)
      formData.append('excel', excelFile)
      formData.append('processedTemplate', processedTemplate)

      const response = await fetch('/api/process', {
        method: 'POST',
        body: formData
      })

      setProgress(100)

      if (response.ok) {
        const blob = await response.blob()
        const url = URL.createObjectURL(blob)
        setProcessedFile(url)
        setStatus('completed')
        toast({
          title: "Processing complete!",
          description: "Your PowerPoint has been generated with dynamic features.",
        })
      } else {
        throw new Error('Processing failed')
      }
    } catch (error) {
      setStatus('error')
      toast({
        title: "Processing failed",
        description: "There was an error processing your files.",
        variant: "destructive",
      })
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
      
      toast({
        title: "Download started",
        description: "Your PowerPoint file is being downloaded.",
      })
    }
  }

  const loadSampleData = () => {
    // Load sample template
    const sampleTemplate = `{project_title}
Status Report prepared by: {name}

Overview
{status_update}

Agenda
{agenda}

{ChartTitle}
This is a placeholder slide. The title above will be filled from your Excel data.`
    
    setTemplate(sampleTemplate)
    
    // Load sample data
    const sampleData = {
      project_title: 'DGE Review',
      name: 'zinab',
      status_update: 'All financial models have been updated, and the initial report draft is complete. We are on schedule for the board meeting next week.',
      agenda: 'Technology Stack System Context Design Decisions Mobile Application Features',
      ChartTitle: 'Analysis of Key Metrics'
    }
    setData(sampleData)
    setShowPreview(true)
    
    toast({
      title: "Sample data loaded",
      description: "Sample template and data have been loaded for preview.",
    })
  }

  const downloadSampleFile = (type: 'pptx' | 'xlsx') => {
    const fileName = type === 'pptx' ? 'sample-template.pptx' : 'sample-data.xlsx'
    const url = `/samples/${fileName}`
    
    const a = document.createElement('a')
    a.href = url
    a.download = fileName
    document.body.appendChild(a)
    a.click()
    document.body.removeChild(a)
    
    toast({
      title: "Download started",
      description: `Sample ${type === 'pptx' ? 'PowerPoint template' : 'Excel data file'} is being downloaded.`,
    })
  }

  return (
    <div className="min-h-screen bg-gradient-to-br from-indigo-500 via-purple-500 to-pink-500">
      {/* Header */}
      <header className="bg-white/10 backdrop-blur-md border-b border-white/20">
        <div className="container mx-auto px-4 py-4 flex items-center justify-between">
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 bg-gradient-to-r from-blue-500 to-purple-600 rounded-lg flex items-center justify-center">
              <FileText className="h-6 w-6 text-white" />
            </div>
            <span className="text-2xl font-bold text-white">Presentr Pro</span>
          </div>
          <div className="flex items-center gap-4">
            <span className="text-sm text-white/80">Advanced Mode</span>
            <div className="w-2 h-2 bg-green-400 rounded-full animate-pulse"></div>
          </div>
        </div>
      </header>

      {/* Main Content */}
      <main className="container mx-auto px-4 py-8">
        <div className="max-w-4xl mx-auto">
          {/* Hero Section */}
          <div className="text-center mb-8">
            <h1 className="text-4xl md:text-5xl font-bold text-white mb-4">
              PowerPoint AutoFill Pro
            </h1>
            <p className="text-xl text-white/90 mb-8 max-w-2xl mx-auto">
              Transform your Excel data into professional PowerPoint presentations with AI-powered dynamic templates and real-time preview
            </p>
            <div className="flex justify-center flex-wrap gap-4">
              <div className="bg-white/20 backdrop-blur-sm rounded-lg px-4 py-2">
                <span className="text-white font-semibold flex items-center gap-2">
                  <Shield className="w-4 h-4" />
                  Server Processing
                </span>
              </div>
              <div className="bg-white/20 backdrop-blur-sm rounded-lg px-4 py-2">
                <span className="text-white font-semibold flex items-center gap-2">
                  <BarChart3 className="w-4 h-4" />
                  Visual Charts
                </span>
              </div>
              <div className="bg-white/20 backdrop-blur-sm rounded-lg px-4 py-2">
                <span className="text-white font-semibold flex items-center gap-2">
                  <Zap className="w-4 h-4" />
                  Instant Processing
                </span>
              </div>
              <div className="bg-white/20 backdrop-blur-sm rounded-lg px-4 py-2">
                <span className="text-white font-semibold flex items-center gap-2">
                  <Layers className="w-4 h-4" />
                  Dynamic Templates
                </span>
              </div>
              <div className="bg-white/20 backdrop-blur-sm rounded-lg px-4 py-2">
                <span className="text-white font-semibold flex items-center gap-2">
                  <Database className="w-4 h-4" />
                  Automation
                </span>
              </div>
            </div>
          </div>

          {/* Quick Start Section */}
          {!pptFile && !excelFile && (
            <Card className="bg-white/95 backdrop-blur-sm border-0 shadow-2xl mb-8">
              <CardContent className="p-6">
                <div className="text-center">
                  <div className="w-16 h-16 bg-gradient-to-r from-blue-500 to-purple-600 rounded-full flex items-center justify-center mx-auto mb-4">
                    <Zap className="h-8 w-8 text-white" />
                  </div>
                  <h3 className="text-xl font-semibold text-gray-800 mb-2">Quick Start Guide</h3>
                  <p className="text-gray-600 mb-6 max-w-2xl mx-auto">
                    Get started in seconds! You can either upload your own files or try our sample files to see the magic happen.
                  </p>
                  <div className="flex justify-center gap-4 mb-6">
                    <Button
                      onClick={loadSampleData}
                      className="bg-gradient-to-r from-blue-500 to-purple-600 hover:from-blue-600 hover:to-purple-700 text-white"
                    >
                      <Eye className="w-4 h-4 mr-2" />
                      Try Sample Preview
                    </Button>
                    <Button
                      variant="outline"
                      onClick={() => {
                        downloadSampleFile('pptx')
                        setTimeout(() => downloadSampleFile('xlsx'), 500)
                      }}
                    >
                      <Download className="w-4 h-4 mr-2" />
                      Download Sample Files
                    </Button>
                  </div>
                  <div className="grid md:grid-cols-3 gap-4 text-sm">
                    <div className="flex items-center gap-2 text-gray-600">
                      <div className="w-8 h-8 bg-blue-100 rounded-full flex items-center justify-center">
                        <span className="text-blue-600 font-semibold">1</span>
                      </div>
                      <span>Upload or load sample files</span>
                    </div>
                    <div className="flex items-center gap-2 text-gray-600">
                      <div className="w-8 h-8 bg-blue-100 rounded-full flex items-center justify-center">
                        <span className="text-blue-600 font-semibold">2</span>
                      </div>
                      <span>Enable advanced features</span>
                    </div>
                    <div className="flex items-center gap-2 text-gray-600">
                      <div className="w-8 h-8 bg-blue-100 rounded-full flex items-center justify-center">
                        <span className="text-blue-600 font-semibold">3</span>
                      </div>
                      <span>Process & download result</span>
                    </div>
                  </div>
                </div>
              </CardContent>
            </Card>
          )}

          {/* Upload Section */}
          <Card className="bg-white/95 backdrop-blur-sm border-0 shadow-2xl mb-8">
            <CardHeader>
              <CardTitle className="text-2xl text-gray-800">Upload Your Files</CardTitle>
              <CardDescription className="text-gray-600">
                Upload your PowerPoint template and Excel data file to get started
              </CardDescription>
            </CardHeader>
            <CardContent className="space-y-6">
              <div className="grid md:grid-cols-2 gap-6">
                {/* PPT Upload */}
                <div className="border-2 border-dashed border-gray-300 rounded-lg p-6 text-center hover:border-purple-400 transition-colors bg-gray-50/50">
                  <input
                    type="file"
                    accept=".pptx"
                    onChange={handlePptUpload}
                    className="hidden"
                    id="ppt-upload"
                  />
                  <label htmlFor="ppt-upload" className="cursor-pointer">
                    <FileText className="h-12 w-12 mx-auto mb-4 text-gray-400" />
                    <p className="font-medium mb-2 text-gray-700">PowerPoint Template</p>
                    <p className="text-sm text-gray-500 mb-4">
                      Upload .pptx file with {'{placeholder}'} tags
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
                <div className="border-2 border-dashed border-gray-300 rounded-lg p-6 text-center hover:border-purple-400 transition-colors bg-gray-50/50">
                  <input
                    type="file"
                    accept=".xlsx"
                    onChange={handleExcelUpload}
                    className="hidden"
                    id="excel-upload"
                  />
                  <label htmlFor="excel-upload" className="cursor-pointer">
                    <FileSpreadsheet className="h-12 w-12 mx-auto mb-4 text-gray-400" />
                    <p className="font-medium mb-2 text-gray-700">Excel Data</p>
                    <p className="text-sm text-gray-500 mb-4">
                      Upload .xlsx file with matching columns
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

              {/* Sample Files Section */}
              <div className="border-t pt-6">
                <div className="text-center mb-4">
                  <h3 className="text-lg font-semibold text-gray-800 mb-2">Don't have files ready?</h3>
                  <p className="text-sm text-gray-600 mb-4">Try our sample files to see how it works</p>
                </div>
                <div className="grid grid-cols-2 md:grid-cols-4 gap-4">
                  <Button
                    variant="outline"
                    size="sm"
                    onClick={() => downloadSampleFile('pptx')}
                    className="flex flex-col items-center gap-2 h-auto py-3"
                  >
                    <FileText className="w-6 h-6" />
                    <span className="text-xs">Sample PPTX</span>
                  </Button>
                  <Button
                    variant="outline"
                    size="sm"
                    onClick={() => downloadSampleFile('xlsx')}
                    className="flex flex-col items-center gap-2 h-auto py-3"
                  >
                    <FileSpreadsheet className="w-6 h-6" />
                    <span className="text-xs">Sample Excel</span>
                  </Button>
                  <Button
                    variant="outline"
                    size="sm"
                    onClick={loadSampleData}
                    className="flex flex-col items-center gap-2 h-auto py-3 bg-blue-50 border-blue-200 hover:bg-blue-100"
                  >
                    <Eye className="w-6 h-6 text-blue-600" />
                    <span className="text-xs text-blue-600">Load Sample Data</span>
                  </Button>
                  <Button
                    variant="outline"
                    size="sm"
                    onClick={reset}
                    className="flex flex-col items-center gap-2 h-auto py-3"
                  >
                    <X className="w-6 h-6" />
                    <span className="text-xs">Clear All</span>
                  </Button>
                </div>
              </div>

              {/* Process Button */}
              <div className="flex justify-center">
                <Button
                  size="lg"
                  onClick={activeTab === 'advanced' ? processWithDynamicFeatures : processFiles}
                  disabled={!pptFile || !excelFile || isProcessing}
                  className="px-8 bg-gradient-to-r from-blue-500 to-purple-600 hover:from-blue-600 hover:to-purple-700 text-white font-semibold"
                >
                  {isProcessing ? 'Processing...' : activeTab === 'advanced' ? 'Process with AI' : 'Process Files'}
                  <ArrowRight className="w-5 h-5 ml-2" />
                </Button>
              </div>
              
              {/* Advanced Features Toggle */}
              <div className="flex justify-center mt-4">
                <Button
                  variant="outline"
                  size="sm"
                  onClick={() => setActiveTab(activeTab === 'basic' ? 'advanced' : 'basic')}
                  className="text-sm"
                >
                  <Settings className="w-4 h-4 mr-2" />
                  {activeTab === 'basic' ? 'Enable Advanced Features' : 'Basic Mode'}
                </Button>
              </div>
            </CardContent>
          </Card>

          {/* Advanced Features Section */}
          {activeTab === 'advanced' && (
            <div className="grid md:grid-cols-2 gap-6 mb-8">
              {/* Real-time Preview */}
              <Card className="bg-white/95 backdrop-blur-sm border-0 shadow-xl">
                <CardHeader>
                  <CardTitle className="flex items-center gap-2">
                    <Eye className="w-5 h-5" />
                    Real-time Preview
                  </CardTitle>
                  <CardDescription>
                    See changes instantly as you edit
                  </CardDescription>
                </CardHeader>
                <CardContent>
                  <Button
                    variant="outline"
                    className="w-full"
                    onClick={() => setShowPreview(!showPreview)}
                  >
                    <Eye className="w-4 h-4 mr-2" />
                    {showPreview ? 'Hide Preview' : 'Show Preview'}
                  </Button>
                </CardContent>
              </Card>

              {/* Automation System */}
              <Card className="bg-white/95 backdrop-blur-sm border-0 shadow-xl">
                <CardHeader>
                  <CardTitle className="flex items-center gap-2">
                    <Play className="w-5 h-5" />
                    Automation System
                  </CardTitle>
                  <CardDescription>
                    Enable automatic processing and monitoring
                  </CardDescription>
                </CardHeader>
                <CardContent>
                  <Button
                    variant={isAutomationEnabled ? "default" : "outline"}
                    className="w-full"
                    onClick={toggleAutomation}
                  >
                    {isAutomationEnabled ? <Pause className="w-4 h-4 mr-2" /> : <Play className="w-4 h-4 mr-2" />}
                    {isAutomationEnabled ? 'Stop Automation' : 'Start Automation'}
                  </Button>
                </CardContent>
              </Card>
            </div>
          )}

          {/* Real-time Preview Panel */}
          {showPreview && template && Object.keys(data).length > 0 && (
            <Card className="bg-white/95 backdrop-blur-sm border-0 shadow-xl mb-8">
              <CardHeader>
                <CardTitle className="flex items-center gap-2">
                  <Eye className="w-5 h-5" />
                  Live Preview
                </CardTitle>
              </CardHeader>
              <CardContent className="h-[600px]">
                <RealtimePreview
                  template={template}
                  data={data}
                  pptxFile={pptFile}
                  onTemplateChange={setTemplate}
                  onDataChange={setData}
                />
              </CardContent>
            </Card>
          )}

          {/* Progress Section */}
          {status === 'processing' && (
            <Card className="bg-white/95 backdrop-blur-sm border-0 shadow-2xl mb-8">
              <CardContent className="p-6">
                <div className="flex items-center justify-between mb-2">
                  <span className="text-sm font-medium text-gray-700">Processing files...</span>
                  <span className="text-sm text-gray-500">{progress}%</span>
                </div>
                <Progress value={progress} className="w-full h-2" />
                <div className="mt-2 text-sm text-gray-600">
                  {progress < 30 && 'Reading files...'}
                  {progress >= 30 && progress < 60 && 'Processing data...'}
                  {progress >= 60 && progress < 90 && 'Generating PowerPoint...'}
                  {progress >= 90 && 'Finalizing...'}
                </div>
              </CardContent>
            </Card>
          )}

          {/* Results Section */}
          {status === 'completed' && processedFile && (
            <Card className="bg-white/95 backdrop-blur-sm border-0 shadow-2xl">
              <CardContent className="p-6 text-center">
                <div className="w-16 h-16 bg-green-100 rounded-full flex items-center justify-center mx-auto mb-4">
                  <CheckCircle className="h-8 w-8 text-green-600" />
                </div>
                <h3 className="text-xl font-semibold text-gray-800 mb-2">Processing Complete!</h3>
                <p className="text-gray-600 mb-6">Your PowerPoint has been generated successfully.</p>
                <div className="flex justify-center gap-4">
                  <Button onClick={downloadProcessedFile} className="bg-gradient-to-r from-green-500 to-emerald-600 hover:from-green-600 hover:to-emerald-700">
                    <Download className="w-4 h-4 mr-2" />
                    Download PowerPoint
                  </Button>
                  <Button variant="outline" onClick={reset}>
                    Start Over
                  </Button>
                </div>
              </CardContent>
            </Card>
          )}

          {/* Error Section */}
          {status === 'error' && (
            <Card className="bg-white/95 backdrop-blur-sm border-0 shadow-2xl">
              <CardContent className="p-6 text-center">
                <div className="w-16 h-16 bg-red-100 rounded-full flex items-center justify-center mx-auto mb-4">
                  <AlertCircle className="h-8 w-8 text-red-600" />
                </div>
                <h3 className="text-xl font-semibold text-gray-800 mb-2">Processing Failed</h3>
                <p className="text-gray-600 mb-6">There was an error processing your files. Please try again.</p>
                <div className="flex justify-center gap-4">
                  <Button onClick={reset} variant="outline">
                    Try Again
                  </Button>
                </div>
              </CardContent>
            </Card>
          )}
        </div>
      </main>

      {/* Footer */}
      <footer className="bg-white/10 backdrop-blur-md border-t border-white/20 mt-16">
        <div className="container mx-auto px-4 py-6">
          <div className="text-center text-white/80 text-sm">
            <p>© 2024 Presentr Pro. Advanced PowerPoint automation with AI.</p>
          </div>
        </div>
      </footer>

      <Toaster />
    </div>
  )
}