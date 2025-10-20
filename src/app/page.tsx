'use client'

import { useState } from 'react'
import { Button } from '@/components/ui/button'
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from '@/components/ui/card'
import { Badge } from '@/components/ui/badge'
import { Progress } from '@/components/ui/progress'
import { useToast } from '@/hooks/use-toast'
import { Toaster } from '@/components/ui/toaster'
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
  Shield
} from 'lucide-react'

export default function Home() {
  const [pptFile, setPptFile] = useState<File | null>(null)
  const [excelFile, setExcelFile] = useState<File | null>(null)
  const [isProcessing, setIsProcessing] = useState(false)
  const [processedFile, setProcessedFile] = useState<string | null>(null)
  const [progress, setProgress] = useState(0)
  const [status, setStatus] = useState<'idle' | 'processing' | 'completed' | 'error'>('idle')
  const { toast } = useToast()

  const handlePptUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0]
    if (file && (file.type === 'application/vnd.openxmlformats-officedocument.presentationml.presentation' || file.name.endsWith('.pptx'))) {
      setPptFile(file)
      setStatus('idle')
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

  const reset = () => {
    setPptFile(null)
    setExcelFile(null)
    setProcessedFile(null)
    setProgress(0)
    setStatus('idle')
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
            <span className="text-2xl font-bold text-white">Presentr</span>
          </div>
          <div className="flex items-center gap-4">
            <span className="text-sm text-white/80">Online Mode</span>
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
              PowerPoint AutoFill
            </h1>
            <p className="text-xl text-white/90 mb-8 max-w-2xl mx-auto">
              Transform your Excel data into professional PowerPoint presentations with visual charts
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
            </div>
          </div>

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

              {/* Process Button */}
              <div className="flex justify-center">
                <Button
                  size="lg"
                  onClick={processFiles}
                  disabled={!pptFile || !excelFile || isProcessing}
                  className="px-8 bg-gradient-to-r from-blue-500 to-purple-600 hover:from-blue-600 hover:to-purple-700 text-white font-semibold"
                >
                  {isProcessing ? 'Processing...' : 'Process Files'}
                  <ArrowRight className="w-5 h-5 ml-2" />
                </Button>
              </div>
            </CardContent>
          </Card>

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
                <h3 className="text-xl font-semibold text-gray-800 mb-2">
                  Processing Complete!
                </h3>
                <p className="text-gray-600 mb-6">
                  Your PowerPoint presentation has been generated successfully with charts and data.
                </p>
                <div className="flex justify-center gap-4">
                  <Button onClick={downloadProcessedFile} className="bg-gradient-to-r from-green-500 to-emerald-600 hover:from-green-600 hover:to-emerald-700">
                    <Download className="w-4 h-4 mr-2" />
                    Download PowerPoint
                  </Button>
                  <Button variant="outline" onClick={reset}>
                    Process Another File
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
                <h3 className="text-xl font-semibold text-gray-800 mb-2">
                  Processing Failed
                </h3>
                <p className="text-gray-600 mb-6">
                  There was an error processing your files. Please check your file formats and try again.
                </p>
                <Button variant="outline" onClick={reset}>
                  Try Again
                </Button>
              </CardContent>
            </Card>
          )}
        </div>
      </main>

      {/* Footer */}
      <footer className="bg-white/10 backdrop-blur-md border-t border-white/20 mt-16">
        <div className="container mx-auto px-4 py-6">
          <div className="flex flex-col md:flex-row justify-between items-center">
            <div className="flex items-center gap-2 mb-4 md:mb-0">
              <FileText className="h-5 w-5 text-white/80" />
              <span className="text-white/80">Presentr - PowerPoint AutoFill Tool</span>
            </div>
            <div className="flex items-center gap-6">
              <Button variant="ghost" size="sm" className="text-white/80 hover:text-white hover:bg-white/10" onClick={() => window.location.href = '/samples'}>
                <Download className="w-4 h-4 mr-2" />
                Sample Files
              </Button>
            </div>
          </div>
        </div>
      </footer>

      <Toaster />
    </div>
  )
}