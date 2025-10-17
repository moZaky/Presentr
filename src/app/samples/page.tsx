'use client'

import { Button } from '@/components/ui/button'
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from '@/components/ui/card'
import { Badge } from '@/components/ui/badge'
import { 
  Download, 
  FileText, 
  FileSpreadsheet, 
  ArrowLeft,
  CheckCircle,
  Info
} from 'lucide-react'

export default function SamplesPage() {
  const downloadFile = (filename: string) => {
    const link = document.createElement('a')
    link.href = `/samples/${filename}`
    link.download = filename
    document.body.appendChild(link)
    link.click()
    document.body.removeChild(link)
  }

  return (
    <div className="min-h-screen bg-gradient-to-br from-slate-50 to-slate-100">
      {/* Header */}
      <header className="border-b bg-white/80 backdrop-blur-sm">
        <div className="container mx-auto px-4 py-4 flex items-center justify-between">
          <div className="flex items-center gap-4">
            <Button variant="ghost" size="sm" onClick={() => window.history.back()}>
              <ArrowLeft className="w-4 h-4 mr-2" />
              Back
            </Button>
            <div className="flex items-center gap-2">
              <FileText className="h-8 w-8 text-primary" />
              <span className="text-2xl font-bold">Sample Files</span>
            </div>
          </div>
          <Badge variant="secondary">Ready to Test</Badge>
        </div>
      </header>

      {/* Main Content */}
      <main className="container mx-auto px-4 py-12">
        <div className="max-w-4xl mx-auto">
          {/* Hero Section */}
          <div className="text-center mb-12">
            <h1 className="text-4xl font-bold mb-4 bg-gradient-to-r from-slate-900 to-slate-600 bg-clip-text text-transparent">
              Test the PowerPoint AutoFill
            </h1>
            <p className="text-xl text-muted-foreground mb-8">
              Download these sample files to test the placeholder replacement functionality
            </p>
          </div>

          {/* Sample Files */}
          <div className="grid md:grid-cols-2 gap-6 mb-12">
            {/* PowerPoint Template */}
            <Card className="hover:shadow-lg transition-shadow">
              <CardHeader>
                <div className="flex items-center gap-3">
                  <div className="w-12 h-12 bg-blue-100 rounded-lg flex items-center justify-center">
                    <FileText className="w-6 h-6 text-blue-600" />
                  </div>
                  <div>
                    <CardTitle>PowerPoint Template</CardTitle>
                    <CardDescription>professional-template.pptx</CardDescription>
                  </div>
                </div>
              </CardHeader>
              <CardContent>
                <div className="space-y-4">
                  <div className="text-sm text-gray-600">
                    <p className="font-medium mb-2">Contains placeholders:</p>
                    <div className="flex flex-wrap gap-2">
                      {['{project_title}', '{name}', '{date}', '{status_update}', '{agenda}', '{ChartTitle}'].map((placeholder) => (
                        <Badge key={placeholder} variant="outline" className="text-xs">
                          {placeholder}
                        </Badge>
                      ))}
                    </div>
                  </div>
                  <Button 
                    className="w-full" 
                    onClick={() => downloadFile('professional-template.pptx')}
                  >
                    <Download className="w-4 h-4 mr-2" />
                    Download PowerPoint
                  </Button>
                </div>
              </CardContent>
            </Card>

            {/* Excel Data */}
            <Card className="hover:shadow-lg transition-shadow">
              <CardHeader>
                <div className="flex items-center gap-3">
                  <div className="w-12 h-12 bg-green-100 rounded-lg flex items-center justify-center">
                    <FileSpreadsheet className="w-6 h-6 text-green-600" />
                  </div>
                  <div>
                    <CardTitle>Excel Data</CardTitle>
                    <CardDescription>sample-data.xlsx</CardDescription>
                  </div>
                </div>
              </CardHeader>
              <CardContent>
                <div className="space-y-4">
                  <div className="text-sm text-gray-600">
                    <p className="font-medium mb-2">Contains:</p>
                    <ul className="space-y-1">
                      <li>• Sheet1: Main data (name, project_title, etc.)</li>
                      <li>• ChartData: Chart information (bar, line, pie)</li>
                    </ul>
                  </div>
                  <Button 
                    className="w-full" 
                    onClick={() => downloadFile('sample-data.xlsx')}
                  >
                    <Download className="w-4 h-4 mr-2" />
                    Download Excel
                  </Button>
                </div>
              </CardContent>
            </Card>
          </div>

          {/* Instructions */}
          <Card>
            <CardHeader>
              <CardTitle className="flex items-center gap-2">
                <Info className="w-5 h-5" />
                How to Test
              </CardTitle>
            </CardHeader>
            <CardContent>
              <div className="space-y-6">
                <div className="grid md:grid-cols-4 gap-4">
                  <div className="text-center">
                    <div className="w-10 h-10 bg-primary text-white rounded-full flex items-center justify-center mx-auto mb-2">
                      1
                    </div>
                    <p className="text-sm font-medium">Download both files</p>
                  </div>
                  <div className="text-center">
                    <div className="w-10 h-10 bg-primary text-white rounded-full flex items-center justify-center mx-auto mb-2">
                      2
                    </div>
                    <p className="text-sm font-medium">Go to main page</p>
                  </div>
                  <div className="text-center">
                    <div className="w-10 h-10 bg-primary text-white rounded-full flex items-center justify-center mx-auto mb-2">
                      3
                    </div>
                    <p className="text-sm font-medium">Upload both files</p>
                  </div>
                  <div className="text-center">
                    <div className="w-10 h-10 bg-primary text-white rounded-full flex items-center justify-center mx-auto mb-2">
                      4
                    </div>
                    <p className="text-sm font-medium">Download result</p>
                  </div>
                </div>

                <div className="bg-blue-50 border border-blue-200 rounded-lg p-4">
                  <div className="flex items-start gap-3">
                    <CheckCircle className="w-5 h-5 text-blue-600 mt-0.5" />
                    <div>
                      <p className="font-medium text-blue-900 mb-1">Expected Result</p>
                      <p className="text-sm text-blue-700">
                        The processed PowerPoint will have all placeholders replaced with data from Excel:
                        "Q4 2024 Performance Review" by "Sarah Johnson" with status updates and chart slides.
                      </p>
                    </div>
                  </div>
                </div>

                <div className="bg-gray-50 border border-gray-200 rounded-lg p-4">
                  <p className="text-sm text-gray-600">
                    <strong>Tip:</strong> The sample PowerPoint has 4 slides with different placeholders, and the Excel file contains matching data plus chart information that will create additional slides.
                  </p>
                </div>
              </div>
            </CardContent>
          </Card>
        </div>
      </main>
    </div>
  )
}