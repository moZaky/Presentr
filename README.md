# 🚀 PowerPoint AutoFill Application

A modern web application that automatically replaces placeholders in PowerPoint templates with data from Excel spreadsheets. Built with cutting-edge technologies and designed for production use.

## ✨ Features

### 🎯 Core Functionality
- **📄 PowerPoint Processing** - Automatic placeholder replacement in PPTX files
- **📊 Excel Integration** - Read data from XLSX files with multiple sheets
- **📈 Visual Chart Generation** - Create actual PowerPoint charts (bar, line, pie) from Excel data
- **🔄 Real-time Processing** - Fast file processing with instant download

### 🎨 User Experience
- **📱 Responsive Design** - Works perfectly on desktop and mobile devices
- **🎯 Modern UI** - Beautiful interface built with shadcn/ui components
- **⚡ Instant Feedback** - Real-time processing status and progress indicators
- **📂 Sample Files** - Ready-to-use templates and data for testing

## 🛠️ Technology Stack

### 🎯 Core Framework
- **⚡ Next.js 15** - React framework with App Router
- **📘 TypeScript 5** - Type-safe development
- **🎨 Tailwind CSS 4** - Modern utility-first CSS

### 🧩 UI Components
- **🧩 shadcn/ui** - High-quality accessible components
- **🎯 Lucide React** - Beautiful icon library
- **🌈 Framer Motion** - Smooth animations

### 📊 File Processing
- **📄 XLSX.js** - Excel file parsing
- **🗂️ JSZip** - ZIP file manipulation for PowerPoint
- **🎯 PPTXGenJS** - PowerPoint generation library

### 🔄 Backend & Data
- **🗄️ Prisma** - Type-safe database operations
- **🔐 NextAuth.js** - Authentication ready
- **🐻 Zustand** - Simple state management
- **🔄 TanStack Query** - Server state synchronization

## 🚀 Quick Start

```bash
# Install dependencies
npm install

# Start development server
npm run dev

# Build for production
npm run build

# Start production server
npm start
```

Open [http://localhost:3000](http://localhost:3000) to use the application.

## 📋 How to Use

### 🎯 Basic Usage
1. **Upload PowerPoint Template** - Choose a PPTX file with `{placeholder}` tags
2. **Upload Excel Data** - Select XLSX file with matching column names
3. **Process Files** - Click "Process Files" to replace placeholders
4. **Download Result** - Get your customized PowerPoint presentation

### 📄 Placeholders Supported
Any text wrapped in curly braces `{}` will be replaced with corresponding Excel data:
- `{project_title}` - Project or presentation title
- `{name}` - Presenter name
- `{date}` - Presentation date
- `{status_update}` - Status information
- `{agenda}` - Agenda items
- `{ChartTitle}` - Chart title

### 📊 Excel Data Structure
- **Sheet1** - Main data row with column names matching placeholders
- **ChartData** - Chart information (ChartTitle, ChartType, Series, Label, Value)

## 🧪 Test with Sample Files

The application includes ready-to-use sample files:

1. **Visit** http://localhost:3000
2. **Click** "View Sample Files" 
3. **Download** `professional-template.pptx` and `sample-data.xlsx`
4. **Upload** both files to test AutoFill functionality
5. **Download** the processed PowerPoint result

### 📄 Sample Template Features
- **4 Professional Slides** - Title, Overview, Agenda, Chart placeholder
- **6 Dynamic Placeholders** - Automatic content replacement
- **Chart Support** - Automatic chart slide generation
- **65,729 bytes** - Complete PowerPoint structure

## 🎯 Supported Chart Types

- **📊 Bar Charts** - Quarterly sales, performance metrics, comparisons
- **📈 Line Charts** - Trends over time, engagement metrics, growth patterns  
- **🥧 Pie Charts** - Demographics, distribution data, percentage breakdowns

**Chart Features:**
- **Visual PowerPoint Charts** - Not just text placeholders, but actual chart objects
- **Automatic Data Grouping** - Multiple data series automatically grouped and displayed
- **Legend Support** - Charts include legends for better data interpretation
- **Professional Styling** - Consistent colors, fonts, and sizing

## 🏗️ Project Structure

```
src/
├── app/                    # Next.js App Router pages
│   ├── api/               # API routes for file processing
│   ├── samples/           # Sample files page
│   └── page.tsx           # Main application page
├── components/            # React components
│   └── ui/               # shadcn/ui components
├── hooks/                # Custom React hooks
├── lib/                  # Utility functions
└── samples/              # Sample files and templates
```

## 🤖 Powered by Z.ai

This application is optimized for use with [Z.ai](https://chat.z.ai) - your AI assistant for:
- **💻 Feature Enhancement** - Add new functionality instantly
- **🎨 UI Improvements** - Create beautiful interfaces
- **🔧 Bug Fixing** - Intelligent issue resolution
- **📝 Documentation** - Auto-generate comprehensive docs
- **🚀 Optimization** - Performance improvements

## 🌟 Production Features

- **🔒 Type Safety** - End-to-end TypeScript with Zod validation
- **📱 Responsive Design** - Mobile-first approach
- **⚡ Performance** - Optimized file processing
- **🎨 Accessibility** - WCAG compliant components
- **🔐 Security** - Safe file handling and validation
- **📊 Analytics** - Processing status and feedback

Ready to enhance your PowerPoint workflow? Start building with AI assistance at [chat.z.ai](https://chat.z.ai)!

---

Built with ❤️ for productivity enhancement. Supercharged by [Z.ai](https://chat.z.ai) 🚀
