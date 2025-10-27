/**
 * Internationalization (i18n) System
 * Supports multiple languages, RTL support, and cultural adaptations
 */

export interface Language {
  code: string
  name: string
  nativeName: string
  direction: 'ltr' | 'rtl'
  locale: string
  dateFormat: string
  numberFormat: string
  currency: string
  pluralRules: PluralRule[]
  genderRules: GenderRule[]
}

export interface PluralRule {
  condition: string
  category: 'zero' | 'one' | 'two' | 'few' | 'many' | 'other'
}

export interface GenderRule {
  type: 'masculine' | 'feminine' | 'neuter'
  patterns: string[]
}

export interface TranslationNamespace {
  [key: string]: string | TranslationNamespace
}

export interface TranslationData {
  [languageCode: string]: {
    [namespace: string]: TranslationNamespace
  }
}

export interface CulturalAdaptation {
  dateFormats: { [key: string]: string }
  numberFormats: { [key: string]: Intl.NumberFormatOptions }
  currencyFormats: { [key: string]: Intl.NumberFormatOptions }
  colorSchemes: { [key: string]: string[] }
  imagery: {
    business: string[]
    education: string[]
    technology: string[]
    healthcare: string[]
  }
  textDirection: 'ltr' | 'rtl'
  readingStyle: 'left-to-right' | 'right-to-left' | 'top-to-bottom'
}

export interface I18nConfig {
  defaultLanguage: string
  fallbackLanguage: string
  supportedLanguages: string[]
  namespaces: string[]
  autoDetect: boolean
  persistChoice: boolean
  debug: boolean
}

/**
 * Internationalization Manager
 */
export class I18nManager {
  private config: I18nConfig
  private currentLanguage: string
  private translations: TranslationData = {}
  private languages: Map<string, Language> = new Map()
  private culturalAdaptations: Map<string, CulturalAdaptation> = new Map()
  private formatters: Map<string, Intl.DateTimeFormat | Intl.NumberFormat> = new Map()
  private listeners: Set<(language: string) => void> = new Set()

  constructor(config: I18nConfig) {
    this.config = config
    this.currentLanguage = config.defaultLanguage
    this.initializeLanguages()
    this.loadCulturalAdaptations()
    
    if (config.autoDetect) {
      this.detectLanguage()
    }
  }

  /**
   * Initialize supported languages
   */
  private initializeLanguages(): void {
    const languages: Language[] = [
      {
        code: 'en',
        name: 'English',
        nativeName: 'English',
        direction: 'ltr',
        locale: 'en-US',
        dateFormat: 'MM/DD/YYYY',
        numberFormat: 'en-US',
        currency: 'USD',
        pluralRules: [
          { condition: 'n === 1', category: 'one' },
          { condition: 'other', category: 'other' }
        ],
        genderRules: []
      },
      {
        code: 'es',
        name: 'Spanish',
        nativeName: 'Español',
        direction: 'ltr',
        locale: 'es-ES',
        dateFormat: 'DD/MM/YYYY',
        numberFormat: 'es-ES',
        currency: 'EUR',
        pluralRules: [
          { condition: 'n === 1', category: 'one' },
          { condition: 'other', category: 'other' }
        ],
        genderRules: [
          { type: 'masculine', patterns: ['el', 'un', 'buen'] },
          { type: 'feminine', patterns: ['la', 'una', 'buena'] }
        ]
      },
      {
        code: 'fr',
        name: 'French',
        nativeName: 'Français',
        direction: 'ltr',
        locale: 'fr-FR',
        dateFormat: 'DD/MM/YYYY',
        numberFormat: 'fr-FR',
        currency: 'EUR',
        pluralRules: [
          { condition: 'n === 0 || n === 1', category: 'one' },
          { condition: 'other', category: 'other' }
        ],
        genderRules: [
          { type: 'masculine', patterns: ['le', 'un', 'bon'] },
          { type: 'feminine', patterns: ['la', 'une', 'bonne'] }
        ]
      },
      {
        code: 'de',
        name: 'German',
        nativeName: 'Deutsch',
        direction: 'ltr',
        locale: 'de-DE',
        dateFormat: 'DD.MM.YYYY',
        numberFormat: 'de-DE',
        currency: 'EUR',
        pluralRules: [
          { condition: 'n === 1', category: 'one' },
          { condition: 'other', category: 'other' }
        ],
        genderRules: [
          { type: 'masculine', patterns: ['der', 'ein', 'guter'] },
          { type: 'feminine', patterns: ['die', 'eine', 'gute'] },
          { type: 'neuter', patterns: ['das', 'ein', 'gutes'] }
        ]
      },
      {
        code: 'zh',
        name: 'Chinese',
        nativeName: '中文',
        direction: 'ltr',
        locale: 'zh-CN',
        dateFormat: 'YYYY-MM-DD',
        numberFormat: 'zh-CN',
        currency: 'CNY',
        pluralRules: [
          { condition: 'other', category: 'other' }
        ],
        genderRules: []
      },
      {
        code: 'ja',
        name: 'Japanese',
        nativeName: '日本語',
        direction: 'ltr',
        locale: 'ja-JP',
        dateFormat: 'YYYY/MM/DD',
        numberFormat: 'ja-JP',
        currency: 'JPY',
        pluralRules: [
          { condition: 'other', category: 'other' }
        ],
        genderRules: []
      },
      {
        code: 'ar',
        name: 'Arabic',
        nativeName: 'العربية',
        direction: 'rtl',
        locale: 'ar-SA',
        dateFormat: 'DD/MM/YYYY',
        numberFormat: 'ar-SA',
        currency: 'SAR',
        pluralRules: [
          { condition: 'n === 0', category: 'zero' },
          { condition: 'n === 1', category: 'one' },
          { condition: 'n === 2', category: 'two' },
          { condition: 'n % 100 >= 3 && n % 100 <= 10', category: 'few' },
          { condition: 'n % 100 >= 11 && n % 100 <= 99', category: 'many' },
          { condition: 'other', category: 'other' }
        ],
        genderRules: [
          { type: 'masculine', patterns: ['ال', 'واحد', 'جيد'] },
          { type: 'feminine', patterns: ['ال', 'واحدة', 'جيدة'] }
        ]
      },
      {
        code: 'hi',
        name: 'Hindi',
        nativeName: 'हिन्दी',
        direction: 'ltr',
        locale: 'hi-IN',
        dateFormat: 'DD/MM/YYYY',
        numberFormat: 'hi-IN',
        currency: 'INR',
        pluralRules: [
          { condition: 'n === 1', category: 'one' },
          { condition: 'other', category: 'other' }
        ],
        genderRules: [
          { type: 'masculine', patterns: ['एक', 'अच्छा'] },
          { type: 'feminine', patterns: ['एक', 'अच्छी'] }
        ]
      },
      {
        code: 'pt',
        name: 'Portuguese',
        nativeName: 'Português',
        direction: 'ltr',
        locale: 'pt-BR',
        dateFormat: 'DD/MM/YYYY',
        numberFormat: 'pt-BR',
        currency: 'BRL',
        pluralRules: [
          { condition: 'n === 0 || n === 1', category: 'one' },
          { condition: 'other', category: 'other' }
        ],
        genderRules: [
          { type: 'masculine', patterns: ['o', 'um', 'bom'] },
          { type: 'feminine', patterns: ['a', 'uma', 'boa'] }
        ]
      },
      {
        code: 'ru',
        name: 'Russian',
        nativeName: 'Русский',
        direction: 'ltr',
        locale: 'ru-RU',
        dateFormat: 'DD.MM.YYYY',
        numberFormat: 'ru-RU',
        currency: 'RUB',
        pluralRules: [
          { condition: 'n % 10 === 1 && n % 100 !== 11', category: 'one' },
          { condition: 'n % 10 >= 2 && n % 10 <= 4 && (n % 100 < 10 || n % 100 >= 20)', category: 'few' },
          { condition: 'other', category: 'many' }
        ],
        genderRules: [
          { type: 'masculine', patterns: ['-', 'один', 'хороший'] },
          { type: 'feminine', patterns: ['-', 'одна', 'хорошая'] },
          { type: 'neuter', patterns: ['-', 'одно', 'хорошее'] }
        ]
      }
    ]

    languages.forEach(lang => this.languages.set(lang.code, lang))
  }

  /**
   * Load cultural adaptations
   */
  private loadCulturalAdaptations(): void {
    // Western cultures
    this.culturalAdaptations.set('en', {
      dateFormats: {
        short: 'MM/DD/YYYY',
        medium: 'MMM DD, YYYY',
        long: 'MMMM DD, YYYY',
        full: 'dddd, MMMM DD, YYYY'
      },
      numberFormats: {
        decimal: { style: 'decimal', minimumFractionDigits: 2 },
        currency: { style: 'currency', currency: 'USD' },
        percent: { style: 'percent' }
      },
      currencyFormats: {
        USD: { style: 'currency', currency: 'USD' }
      },
      colorSchemes: {
        primary: ['#3b82f6', '#ef4444', '#10b981', '#f59e0b'],
        business: ['#1e40af', '#dc2626', '#059669'],
        education: ['#7c3aed', '#ec4899', '#06b6d4']
      },
      imagery: {
        business: ['professional', 'corporate', 'office', 'meeting'],
        education: ['classroom', 'students', 'teacher', 'books'],
        technology: ['computer', 'code', 'innovation', 'digital'],
        healthcare: ['hospital', 'doctor', 'patient', 'medical']
      },
      textDirection: 'ltr',
      readingStyle: 'left-to-right'
    })

    // Arabic cultures
    this.culturalAdaptations.set('ar', {
      dateFormats: {
        short: 'DD/MM/YYYY',
        medium: 'DD MMM, YYYY',
        long: 'DD MMMM, YYYY',
        full: 'dddd, DD MMMM, YYYY'
      },
      numberFormats: {
        decimal: { style: 'decimal', minimumFractionDigits: 2, useGrouping: true },
        currency: { style: 'currency', currency: 'SAR' },
        percent: { style: 'percent' }
      },
      currencyFormats: {
        SAR: { style: 'currency', currency: 'SAR' }
      },
      colorSchemes: {
        primary: ['#059669', '#dc2626', '#f59e0b', '#7c3aed'],
        business: ['#047857', '#b91c1c', '#d97706'],
        education: ['#065f46', '#991b1b', '#92400e']
      },
      imagery: {
        business: ['تجاري', 'شركة', 'مكتب', 'اجتماع'],
        education: ['فصل', 'طلاب', 'معلم', 'كتب'],
        technology: ['كمبيوتر', 'برمجة', 'ابتكار', 'رقمي'],
        healthcare: ['مستشفى', 'طبيب', 'مريض', 'طبي']
      },
      textDirection: 'rtl',
      readingStyle: 'right-to-left'
    })

    // Japanese cultures
    this.culturalAdaptations.set('ja', {
      dateFormats: {
        short: 'YYYY/MM/DD',
        medium: 'YYYY年MM月DD日',
        long: 'YYYY年MM月DD日(ddd)',
        full: 'YYYY年MM月DD日dddd'
      },
      numberFormats: {
        decimal: { style: 'decimal', minimumFractionDigits: 0 },
        currency: { style: 'currency', currency: 'JPY' },
        percent: { style: 'percent' }
      },
      currencyFormats: {
        JPY: { style: 'currency', currency: 'JPY' }
      },
      colorSchemes: {
        primary: ['#dc2626', '#2563eb', '#059669', '#d97706'],
        business: ['#b91c1c', '#1d4ed8', '#047857'],
        education: ['#ef4444', '#3b82f6', '#10b981']
      },
      imagery: {
        business: ['ビジネス', '会社', 'オフィス', '会議'],
        education: ['教室', '学生', '教師', '本'],
        technology: ['コンピューター', 'コード', '革新', 'デジタル'],
        healthcare: ['病院', '医者', '患者', '医療']
      },
      textDirection: 'ltr',
      readingStyle: 'top-to-bottom'
    })
  }

  /**
   * Detect user's preferred language
   */
  private detectLanguage(): void {
    // Check browser language
    const browserLang = navigator.language.split('-')[0]
    
    // Check localStorage
    const savedLang = localStorage.getItem('preferred-language')
    
    // Check URL params
    const urlParams = new URLSearchParams(window.location.search)
    const urlLang = urlParams.get('lang')
    
    // Determine language
    let detectedLang = this.config.defaultLanguage
    
    if (urlLang && this.isLanguageSupported(urlLang)) {
      detectedLang = urlLang
    } else if (savedLang && this.isLanguageSupported(savedLang)) {
      detectedLang = savedLang
    } else if (browserLang && this.isLanguageSupported(browserLang)) {
      detectedLang = browserLang
    }
    
    this.setLanguage(detectedLang)
  }

  /**
   * Check if language is supported
   */
  private isLanguageSupported(languageCode: string): boolean {
    return this.config.supportedLanguages.includes(languageCode)
  }

  /**
   * Set current language
   */
  setLanguage(languageCode: string): void {
    if (!this.isLanguageSupported(languageCode)) {
      console.warn(`Language '${languageCode}' is not supported`)
      return
    }

    const previousLanguage = this.currentLanguage
    this.currentLanguage = languageCode

    // Update HTML attributes
    document.documentElement.lang = languageCode
    document.documentElement.dir = this.getLanguage(languageCode).direction

    // Save preference
    if (this.config.persistChoice) {
      localStorage.setItem('preferred-language', languageCode)
    }

    // Clear formatters cache
    this.formatters.clear()

    // Notify listeners
    if (previousLanguage !== languageCode) {
      this.listeners.forEach(listener => listener(languageCode))
    }
  }

  /**
   * Get current language
   */
  getCurrentLanguage(): string {
    return this.currentLanguage
  }

  /**
   * Get language info
   */
  getLanguage(languageCode?: string): Language {
    return this.languages.get(languageCode || this.currentLanguage)!
  }

  /**
   * Get supported languages
   */
  getSupportedLanguages(): Language[] {
    return this.config.supportedLanguages.map(code => this.languages.get(code)!)
  }

  /**
   * Load translations
   */
  async loadTranslations(languageCode: string, namespace: string): Promise<void> {
    try {
      // In a real app, this would load from API or JSON files
      const translations = await this.fetchTranslations(languageCode, namespace)
      
      if (!this.translations[languageCode]) {
        this.translations[languageCode] = {}
      }
      
      this.translations[languageCode][namespace] = translations
    } catch (error) {
      console.error(`Failed to load translations for ${languageCode}:${namespace}`, error)
    }
  }

  /**
   * Fetch translations (placeholder implementation)
   */
  private async fetchTranslations(languageCode: string, namespace: string): Promise<TranslationNamespace> {
    // This would be replaced with actual API calls
    return this.getDefaultTranslations(languageCode, namespace)
  }

  /**
   * Get default translations
   */
  private getDefaultTranslations(languageCode: string, namespace: string): TranslationNamespace {
    const translations: { [key: string]: TranslationNamespace } = {
      en: {
        common: {
          'app.title': 'PowerPoint AutoFill',
          'app.description': 'Transform your Excel data into professional PowerPoint presentations',
          'button.upload': 'Upload',
          'button.download': 'Download',
          'button.process': 'Process Files',
          'button.generate': 'Generate Presentation',
          'error.file.required': 'Please select both PowerPoint and Excel files',
          'error.file.type': 'Invalid file type. Please upload .pptx and .xlsx files',
          'success.processing': 'Processing complete! Your PowerPoint has been generated.',
          'status.uploading': 'Uploading files...',
          'status.processing': 'Processing files...',
          'status.generating': 'Generating presentation...',
          'status.complete': 'Complete!'
        },
        charts: {
          'title.revenue': 'Revenue',
          'title.expenses': 'Expenses',
          'title.profit': 'Profit',
          'axis.months': 'Months',
          'axis.amount': 'Amount',
          'legend.q1': 'Q1',
          'legend.q2': 'Q2',
          'legend.q3': 'Q3',
          'legend.q4': 'Q4'
        },
        placeholders: {
          'project_title': 'Project Title',
          'name': 'Presenter Name',
          'date': 'Date',
          'status_update': 'Status Update',
          'agenda': 'Agenda',
          'chart_title': 'Chart Title'
        }
      },
      es: {
        common: {
          'app.title': 'AutoFill PowerPoint',
          'app.description': 'Transforma tus datos de Excel en presentaciones PowerPoint profesionales',
          'button.upload': 'Subir',
          'button.download': 'Descargar',
          'button.process': 'Procesar Archivos',
          'button.generate': 'Generar Presentación',
          'error.file.required': 'Por favor selecciona archivos de PowerPoint y Excel',
          'error.file.type': 'Tipo de archivo inválido. Por favor sube archivos .pptx y .xlsx',
          'success.processing': '¡Procesamiento completo! Tu PowerPoint ha sido generado.',
          'status.uploading': 'Subiendo archivos...',
          'status.processing': 'Procesando archivos...',
          'status.generating': 'Generando presentación...',
          'status.complete': '¡Completo!'
        },
        charts: {
          'title.revenue': 'Ingresos',
          'title.expenses': 'Gastos',
          'title.profit': 'Beneficio',
          'axis.months': 'Meses',
          'axis.amount': 'Cantidad',
          'legend.q1': 'T1',
          'legend.q2': 'T2',
          'legend.q3': 'T3',
          'legend.q4': 'T4'
        },
        placeholders: {
          'project_title': 'Título del Proyecto',
          'name': 'Nombre del Presentador',
          'date': 'Fecha',
          'status_update': 'Actualización de Estado',
          'agenda': 'Agenda',
          'chart_title': 'Título del Gráfico'
        }
      },
      fr: {
        common: {
          'app.title': 'AutoFill PowerPoint',
          'app.description': 'Transformez vos données Excel en présentations PowerPoint professionnelles',
          'button.upload': 'Télécharger',
          'button.download': 'Télécharger',
          'button.process': 'Traiter les Fichiers',
          'button.generate': 'Générer la Présentation',
          'error.file.required': 'Veuillez sélectionner des fichiers PowerPoint et Excel',
          'error.file.type': 'Type de fichier invalide. Veuillez télécharger des fichiers .pptx et .xlsx',
          'success.processing': 'Traitement terminé! Votre PowerPoint a été généré.',
          'status.uploading': 'Téléchargement des fichiers...',
          'status.processing': 'Traitement des fichiers...',
          'status.generating': 'Génération de la présentation...',
          'status.complete': 'Terminé!'
        },
        charts: {
          'title.revenue': 'Revenus',
          'title.expenses': 'Dépenses',
          'title.profit': 'Profit',
          'axis.months': 'Mois',
          'axis.amount': 'Montant',
          'legend.q1': 'T1',
          'legend.q2': 'T2',
          'legend.q3': 'T3',
          'legend.q4': 'T4'
        },
        placeholders: {
          'project_title': 'Titre du Projet',
          'name': 'Nom du Présentateur',
          'date': 'Date',
          'status_update': 'Mise à jour du Statut',
          'agenda': 'Ordre du jour',
          'chart_title': 'Titre du Graphique'
        }
      },
      ar: {
        common: {
          'app.title': 'أوتوفيل باوربوينت',
          'app.description': 'حول بيانات Excel الخاصة بك إلى عروض تقديمية احترافية',
          'button.upload': 'رفع',
          'button.download': 'تحميل',
          'button.process': 'معالجة الملفات',
          'button.generate': 'إنشاء العرض التقديمي',
          'error.file.required': 'يرجى تحديد ملفات PowerPoint و Excel',
          'error.file.type': 'نوع الملف غير صالح. يرجى رفع ملفات .pptx و .xlsx',
          'success.processing': 'اكتملت المعالجة! تم إنشاء PowerPoint الخاص بك.',
          'status.uploading': 'جاري رفع الملفات...',
          'status.processing': 'جاري معالجة الملفات...',
          'status.generating': 'جاري إنشاء العرض التقديمي...',
          'status.complete': 'مكتمل!'
        },
        charts: {
          'title.revenue': 'الإيرادات',
          'title.expenses': 'المصروفات',
          'title.profit': 'الربح',
          'axis.months': 'الأشهر',
          'axis.amount': 'المبلغ',
          'legend.q1': 'ر1',
          'legend.q2': 'ر2',
          'legend.q3': 'ر3',
          'legend.q4': 'ر4'
        },
        placeholders: {
          'project_title': 'عنوان المشروع',
          'name': 'اسم العارض',
          'date': 'التاريخ',
          'status_update': 'تحديث الحالة',
          'agenda': 'جدول الأعمال',
          'chart_title': 'عنوان الرسم البياني'
        }
      }
    }

    return translations[languageCode]?.[namespace] || {}
  }

  /**
   * Translate text
   */
  t(key: string, namespace: string = 'common', params?: { [key: string]: any }): string {
    const translation = this.getTranslation(key, namespace)
    
    if (!translation) {
      if (this.config.debug) {
        console.warn(`Translation not found: ${namespace}:${key}`)
      }
      return key
    }

    // Replace parameters
    if (params) {
      return this.interpolateParams(translation, params)
    }

    return translation
  }

  /**
   * Get translation
   */
  private getTranslation(key: string, namespace: string): string | null {
    const keys = key.split('.')
    let current: any = this.translations[this.currentLanguage]?.[namespace]

    if (!current) {
      // Try fallback language
      current = this.translations[this.config.fallbackLanguage]?.[namespace]
    }

    if (!current) {
      return null
    }

    for (const k of keys) {
      if (current && typeof current === 'object' && k in current) {
        current = current[k]
      } else {
        return null
      }
    }

    return typeof current === 'string' ? current : null
  }

  /**
   * Interpolate parameters in translation
   */
  private interpolateParams(text: string, params: { [key: string]: any }): string {
    return text.replace(/\{\{(\w+)\}\}/g, (match, key) => {
      return params[key] !== undefined ? String(params[key]) : match
    })
  }

  /**
   * Format date
   */
  formatDate(date: Date, format?: string): string {
    const language = this.getLanguage()
    const cultural = this.culturalAdaptations.get(this.currentLanguage)
    
    const formatterKey = `date-${format || 'medium'}`
    let formatter = this.formatters.get(formatterKey) as Intl.DateTimeFormat

    if (!formatter) {
      const options: Intl.DateTimeFormatOptions = this.getDateFormatOptions(format || 'medium')
      formatter = new Intl.DateTimeFormat(language.locale, options)
      this.formatters.set(formatterKey, formatter)
    }

    return formatter.format(date)
  }

  /**
   * Get date format options
   */
  private getDateFormatOptions(format: string): Intl.DateTimeFormatOptions {
    const cultural = this.culturalAdaptations.get(this.currentLanguage)
    
    switch (format) {
      case 'short':
        return { year: 'numeric', month: '2-digit', day: '2-digit' }
      case 'medium':
        return { year: 'numeric', month: 'short', day: 'numeric' }
      case 'long':
        return { year: 'numeric', month: 'long', day: 'numeric' }
      case 'full':
        return { year: 'numeric', month: 'long', day: 'numeric', weekday: 'long' }
      default:
        return { year: 'numeric', month: 'short', day: 'numeric' }
    }
  }

  /**
   * Format number
   */
  formatNumber(number: number, options?: Intl.NumberFormatOptions): string {
    const language = this.getLanguage()
    const cultural = this.culturalAdaptations.get(this.currentLanguage)
    
    const formatterKey = `number-${JSON.stringify(options || {})}`
    let formatter = this.formatters.get(formatterKey) as Intl.NumberFormat

    if (!formatter) {
      const defaultOptions = cultural?.numberFormats.decimal || { style: 'decimal' }
      formatter = new Intl.NumberFormat(language.locale, { ...defaultOptions, ...options })
      this.formatters.set(formatterKey, formatter)
    }

    return formatter.format(number)
  }

  /**
   * Format currency
   */
  formatCurrency(amount: number, currency?: string): string {
    const language = this.getLanguage()
    const cultural = this.culturalAdaptations.get(this.currentLanguage)
    
    const currencyCode = currency || language.currency
    const options = cultural?.currencyFormats[currencyCode] || { 
      style: 'currency', 
      currency: currencyCode 
    }

    return this.formatNumber(amount, options)
  }

  /**
   * Format percentage
   */
  formatPercentage(number: number, options?: Intl.NumberFormatOptions): string {
    return this.formatNumber(number, { style: 'percent', ...options })
  }

  /**
   * Get plural form
   */
  getPluralForm(count: number): string {
    const language = this.getLanguage()
    
    for (const rule of language.pluralRules) {
      if (this.evaluateCondition(rule.condition, count)) {
        return rule.category
      }
    }
    
    return 'other'
  }

  /**
   * Evaluate plural condition
   */
  private evaluateCondition(condition: string, n: number): boolean {
    // Simple condition evaluation
    if (condition === 'other') return true
    
    try {
      // Replace 'n' with actual number
      const evalCondition = condition.replace(/n/g, n.toString())
      return Function(`"use strict"; return (${evalCondition})`)()
    } catch {
      return false
    }
  }

  /**
   * Add language change listener
   */
  addListener(listener: (language: string) => void): void {
    this.listeners.add(listener)
  }

  /**
   * Remove language change listener
   */
  removeListener(listener: (language: string) => void): void {
    this.listeners.delete(listener)
  }

  /**
   * Get cultural adaptation
   */
  getCulturalAdaptation(): CulturalAdaptation | null {
    return this.culturalAdaptations.get(this.currentLanguage) || null
  }

  /**
   * Get text direction
   */
  getTextDirection(): 'ltr' | 'rtl' {
    return this.getLanguage().direction
  }

  /**
   * Check if RTL language
   */
  isRTL(): boolean {
    return this.getTextDirection() === 'rtl'
  }
}

/**
 * React Hook for i18n
 */
export function useI18n(i18n: I18nManager) {
  const [language, setLanguage] = React.useState(i18n.getCurrentLanguage())

  React.useEffect(() => {
    const handleLanguageChange = (newLanguage: string) => {
      setLanguage(newLanguage)
    }

    i18n.addListener(handleLanguageChange)
    return () => i18n.removeListener(handleLanguageChange)
  }, [i18n])

  return {
    language,
    setLanguage: (lang: string) => i18n.setLanguage(lang),
    t: (key: string, namespace?: string, params?: { [key: string]: any }) => 
      i18n.t(key, namespace, params),
    formatDate: (date: Date, format?: string) => i18n.formatDate(date, format),
    formatNumber: (number: number, options?: Intl.NumberFormatOptions) => 
      i18n.formatNumber(number, options),
    formatCurrency: (amount: number, currency?: string) => 
      i18n.formatCurrency(amount, currency),
    formatPercentage: (number: number, options?: Intl.NumberFormatOptions) => 
      i18n.formatPercentage(number, options),
    isRTL: i18n.isRTL(),
    getTextDirection: () => i18n.getTextDirection(),
    getSupportedLanguages: () => i18n.getSupportedLanguages(),
    getCurrentLanguageInfo: () => i18n.getLanguage()
  }
}

/**
 * Default i18n configuration
 */
export const defaultI18nConfig: I18nConfig = {
  defaultLanguage: 'en',
  fallbackLanguage: 'en',
  supportedLanguages: ['en', 'es', 'fr', 'de', 'zh', 'ja', 'ar', 'hi', 'pt', 'ru'],
  namespaces: ['common', 'charts', 'placeholders'],
  autoDetect: true,
  persistChoice: true,
  debug: false
}

/**
 * Main I18n System class
 */
export class I18nSystem {
  private config: I18nConfig
  private currentLanguage: string
  private translations: Map<string, Map<string, string>> = new Map()
  private loadedLanguages: Set<string> = new Set()

  constructor(config?: Partial<I18nConfig>) {
    this.config = { ...defaultI18nConfig, ...config }
    this.currentLanguage = this.config.defaultLanguage
    this.loadLanguage(this.currentLanguage)
  }

  /**
   * Set current language
   */
  setLanguage(language: string): void {
    if (!this.config.supportedLanguages.includes(language)) {
      console.warn(`Language ${language} is not supported`)
      return
    }

    this.currentLanguage = language
    this.loadLanguage(language)
    
    if (this.config.persistChoice) {
      localStorage.setItem('language', language)
    }
  }

  /**
   * Get current language
   */
  getCurrentLanguage(): string {
    return this.currentLanguage
  }

  /**
   * Load language translations
   */
  private loadLanguage(language: string): void {
    if (this.loadedLanguages.has(language)) return

    // Mock translations - in real implementation, would load from files
    const mockTranslations = new Map<string, string>()
    mockTranslations.set('project_title', 'Project Title')
    mockTranslations.set('name', 'Name')
    mockTranslations.set('status_update', 'Status Update')
    mockTranslations.set('agenda', 'Agenda')
    mockTranslations.set('ChartTitle', 'Chart Title')

    this.translations.set(language, mockTranslations)
    this.loadedLanguages.add(language)
  }

  /**
   * Translate text
   */
  translate(key: string, language?: string): string {
    const lang = language || this.currentLanguage
    const translations = this.translations.get(lang)
    
    if (!translations) {
      console.warn(`Translations for ${lang} not found`)
      return key
    }

    const translation = translations.get(key)
    return translation || key
  }

  /**
   * Process template with translations
   */
  processTemplate(template: string, language?: string): string {
    const lang = language || this.currentLanguage
    
    // Simple placeholder replacement
    return template.replace(/\{(\w+)\}/g, (match, key) => {
      const translation = this.translate(key, lang)
      return translation !== key ? translation : match
    })
  }

  /**
   * Get supported languages
   */
  getSupportedLanguages(): string[] {
    return this.config.supportedLanguages
  }

  /**
   * Detect user's preferred language
   */
  detectLanguage(): string {
    if (typeof window === 'undefined') return this.config.defaultLanguage

    // Check localStorage first
    const saved = localStorage.getItem('language')
    if (saved && this.config.supportedLanguages.includes(saved)) {
      return saved
    }

    // Check browser language
    const browserLang = navigator.language.split('-')[0]
    if (this.config.supportedLanguages.includes(browserLang)) {
      return browserLang
    }

    return this.config.fallbackLanguage
  }

  /**
   * Format date according to locale
   */
  formatDate(date: Date, language?: string): string {
    const lang = language || this.currentLanguage
    return new Intl.DateTimeFormat(lang).format(date)
  }

  /**
   * Format number according to locale
   */
  formatNumber(number: number, language?: string): string {
    const lang = language || this.currentLanguage
    return new Intl.NumberFormat(lang).format(number)
  }
}