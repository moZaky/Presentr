/**
 * Template Marketplace System
 * Supports template creation, sharing, versioning, and distribution
 */

export interface Template {
  id: string
  name: string
  description: string
  category: TemplateCategory
  tags: string[]
  author: Author
  version: string
  createdAt: Date
  updatedAt: Date
  downloads: number
  rating: number
  reviews: Review[]
  price: number
  license: LicenseType
  preview: TemplatePreview
  files: TemplateFile[]
  metadata: TemplateMetadata
  compatibility: CompatibilityInfo
}

export interface TemplateCategory {
  id: string
  name: string
  description: string
  icon: string
  subcategories: string[]
}

export interface Author {
  id: string
  name: string
  email: string
  avatar?: string
  bio?: string
  website?: string
  verified: boolean
  followers: number
  templates: number
}

export interface Review {
  id: string
  userId: string
  userName: string
  rating: number
  comment: string
  createdAt: Date
  helpful: number
  verified: boolean
}

export interface TemplatePreview {
  thumbnail: string
  screenshots: string[]
  demoData: any
  video?: string
  interactive: boolean
}

export interface TemplateFile {
  name: string
  type: 'pptx' | 'json' | 'md' | 'assets'
  size: number
  url: string
  checksum: string
}

export interface TemplateMetadata {
  placeholders: PlaceholderInfo[]
  slides: number
  charts: number
  images: number
  animations: boolean
  aspectRatio: '16:9' | '4:3' | 'custom'
  colorScheme: string[]
  fonts: string[]
  requirements: string[]
  features: string[]
}

export interface CompatibilityInfo {
  powerpointVersions: string[]
  officeVersions: string[]
  platforms: ('windows' | 'mac' | 'web' | 'mobile')[]
  minPowerPointVersion: string
  testedVersions: string[]
}

export interface PlaceholderInfo {
  name: string
  type: 'text' | 'image' | 'chart' | 'date' | 'number'
  required: boolean
  description: string
  defaultValue?: any
  validation?: ValidationRule
}

export interface ValidationRule {
  pattern?: string
  minLength?: number
  maxLength?: number
  min?: number
  max?: number
  options?: string[]
}

export type LicenseType = 'free' | 'premium' | 'enterprise' | 'custom'

export interface MarketplaceFilters {
  category?: string
  tags?: string[]
  priceRange?: [number, number]
  rating?: number
  author?: string
  sortBy?: 'popularity' | 'newest' | 'rating' | 'price' | 'downloads'
  compatibility?: string[]
  features?: string[]
}

export interface MarketplaceSearch {
  query: string
  filters: MarketplaceFilters
  page: number
  limit: number
}

export interface SearchResult {
  templates: Template[]
  total: number
  page: number
  totalPages: number
  facets: SearchFacets
}

export interface SearchFacets {
  categories: { [key: string]: number }
  tags: { [key: string]: number }
  authors: { [key: string]: number }
  priceRanges: { [key: string]: number }
  ratings: { [key: string]: number }
}

/**
 * Template Marketplace Manager
 */
export class TemplateMarketplace {
  private templates: Map<string, Template> = new Map()
  private categories: Map<string, TemplateCategory> = new Map()
  private authors: Map<string, Author> = new Map()

  constructor() {
    this.initializeDefaultData()
  }

  /**
   * Search templates with filters
   */
  async searchTemplates(search: MarketplaceSearch): Promise<SearchResult> {
    let filteredTemplates = Array.from(this.templates.values())

    // Apply text search
    if (search.query) {
      const query = search.query.toLowerCase()
      filteredTemplates = filteredTemplates.filter(template =>
        template.name.toLowerCase().includes(query) ||
        template.description.toLowerCase().includes(query) ||
        template.tags.some(tag => tag.toLowerCase().includes(query))
      )
    }

    // Apply filters
    if (search.filters.category) {
      filteredTemplates = filteredTemplates.filter(template =>
        template.category.id === search.filters.category
      )
    }

    if (search.filters.tags && search.filters.tags.length > 0) {
      filteredTemplates = filteredTemplates.filter(template =>
        search.filters.tags!.some(tag => template.tags.includes(tag))
      )
    }

    if (search.filters.priceRange) {
      const [min, max] = search.filters.priceRange
      filteredTemplates = filteredTemplates.filter(template =>
        template.price >= min && template.price <= max
      )
    }

    if (search.filters.rating) {
      filteredTemplates = filteredTemplates.filter(template =>
        template.rating >= search.filters.rating!
      )
    }

    if (search.filters.author) {
      filteredTemplates = filteredTemplates.filter(template =>
        template.author.id === search.filters.author
      )
    }

    // Sort results
    filteredTemplates = this.sortTemplates(filteredTemplates, search.filters.sortBy)

    // Paginate
    const startIndex = (search.page - 1) * search.limit
    const endIndex = startIndex + search.limit
    const paginatedTemplates = filteredTemplates.slice(startIndex, endIndex)

    // Generate facets
    const facets = this.generateFacets(filteredTemplates)

    return {
      templates: paginatedTemplates,
      total: filteredTemplates.length,
      page: search.page,
      totalPages: Math.ceil(filteredTemplates.length / search.limit),
      facets
    }
  }

  /**
   * Get template by ID
   */
  async getTemplate(id: string): Promise<Template | null> {
    return this.templates.get(id) || null
  }

  /**
   * Upload new template
   */
  async uploadTemplate(templateData: Omit<Template, 'id' | 'createdAt' | 'updatedAt' | 'downloads' | 'rating' | 'reviews'>): Promise<Template> {
    const template: Template = {
      ...templateData,
      id: this.generateId(),
      createdAt: new Date(),
      updatedAt: new Date(),
      downloads: 0,
      rating: 0,
      reviews: []
    }

    this.templates.set(template.id, template)
    return template
  }

  /**
   * Update template
   */
  async updateTemplate(id: string, updates: Partial<Template>): Promise<Template | null> {
    const template = this.templates.get(id)
    if (!template) return null

    const updatedTemplate = {
      ...template,
      ...updates,
      updatedAt: new Date(),
      version: this.incrementVersion(template.version)
    }

    this.templates.set(id, updatedTemplate)
    return updatedTemplate
  }

  /**
   * Delete template
   */
  async deleteTemplate(id: string): Promise<boolean> {
    return this.templates.delete(id)
  }

  /**
   * Download template (increment download count)
   */
  async downloadTemplate(id: string): Promise<TemplateFile[]> {
    const template = this.templates.get(id)
    if (!template) throw new Error('Template not found')

    // Increment download count
    template.downloads += 1
    this.templates.set(id, template)

    return template.files
  }

  /**
   * Rate template
   */
  async rateTemplate(id: string, userId: string, rating: number, comment?: string): Promise<Review> {
    const template = this.templates.get(id)
    if (!template) throw new Error('Template not found')

    const review: Review = {
      id: this.generateId(),
      userId,
      userName: 'User', // Would get from user service
      rating,
      comment: comment || '',
      createdAt: new Date(),
      helpful: 0,
      verified: false
    }

    template.reviews.push(review)
    
    // Update average rating
    template.rating = template.reviews.reduce((sum, review) => sum + review.rating, 0) / template.reviews.length

    this.templates.set(id, template)
    return review
  }

  /**
   * Get categories
   */
  async getCategories(): Promise<TemplateCategory[]> {
    return Array.from(this.categories.values())
  }

  /**
   * Get featured templates
   */
  async getFeaturedTemplates(limit: number = 10): Promise<Template[]> {
    return Array.from(this.templates.values())
      .filter(template => template.rating >= 4.5 && template.downloads >= 100)
      .sort((a, b) => b.downloads - a.downloads)
      .slice(0, limit)
  }

  /**
   * Get popular templates
   */
  async getPopularTemplates(limit: number = 10): Promise<Template[]> {
    return Array.from(this.templates.values())
      .sort((a, b) => b.downloads - a.downloads)
      .slice(0, limit)
  }

  /**
   * Get new templates
   */
  async getNewTemplates(limit: number = 10): Promise<Template[]> {
    return Array.from(this.templates.values())
      .sort((a, b) => b.createdAt.getTime() - a.createdAt.getTime())
      .slice(0, limit)
  }

  /**
   * Sort templates
   */
  private sortTemplates(templates: Template[], sortBy?: string): Template[] {
    switch (sortBy) {
      case 'popularity':
        return templates.sort((a, b) => b.downloads - a.downloads)
      case 'newest':
        return templates.sort((a, b) => b.createdAt.getTime() - a.createdAt.getTime())
      case 'rating':
        return templates.sort((a, b) => b.rating - a.rating)
      case 'price':
        return templates.sort((a, b) => a.price - b.price)
      case 'downloads':
        return templates.sort((a, b) => b.downloads - a.downloads)
      default:
        return templates
    }
  }

  /**
   * Generate search facets
   */
  private generateFacets(templates: Template[]): SearchFacets {
    const facets: SearchFacets = {
      categories: {},
      tags: {},
      authors: {},
      priceRanges: {},
      ratings: {}
    }

    templates.forEach(template => {
      // Categories
      facets.categories[template.category.id] = (facets.categories[template.category.id] || 0) + 1

      // Tags
      template.tags.forEach(tag => {
        facets.tags[tag] = (facets.tags[tag] || 0) + 1
      })

      // Authors
      facets.authors[template.author.id] = (facets.authors[template.author.id] || 0) + 1

      // Price ranges
      const priceRange = this.getPriceRange(template.price)
      facets.priceRanges[priceRange] = (facets.priceRanges[priceRange] || 0) + 1

      // Ratings
      const ratingRange = this.getRatingRange(template.rating)
      facets.ratings[ratingRange] = (facets.ratings[ratingRange] || 0) + 1
    })

    return facets
  }

  /**
   * Get price range category
   */
  private getPriceRange(price: number): string {
    if (price === 0) return 'Free'
    if (price < 10) return '$1-10'
    if (price < 25) return '$10-25'
    if (price < 50) return '$25-50'
    return '$50+'
  }

  /**
   * Get rating range category
   */
  private getRatingRange(rating: number): string {
    if (rating >= 4.5) return '4.5+'
    if (rating >= 4.0) return '4.0-4.5'
    if (rating >= 3.5) return '3.5-4.0'
    if (rating >= 3.0) return '3.0-3.5'
    return 'Below 3.0'
  }

  /**
   * Generate unique ID
   */
  private generateId(): string {
    return Math.random().toString(36).substr(2, 9)
  }

  /**
   * Increment version
   */
  private incrementVersion(version: string): string {
    const parts = version.split('.')
    const patch = parseInt(parts[2] || '0') + 1
    return `${parts[0]}.${parts[1]}.${patch}`
  }

  /**
   * Initialize default data
   */
  private initializeDefaultData(): void {
    // Initialize categories
    this.categories.set('business', {
      id: 'business',
      name: 'Business',
      description: 'Professional business presentations',
      icon: 'briefcase',
      subcategories: ['marketing', 'sales', 'finance', 'strategy']
    })

    this.categories.set('education', {
      id: 'education',
      name: 'Education',
      description: 'Educational and academic presentations',
      icon: 'graduation-cap',
      subcategories: ['lecture', 'research', 'training', 'e-learning']
    })

    this.categories.set('creative', {
      id: 'creative',
      name: 'Creative',
      description: 'Creative and artistic presentations',
      icon: 'palette',
      subcategories: ['portfolio', 'art', 'design', 'photography']
    })

    // Initialize sample templates
    this.initializeSampleTemplates()
  }

  /**
   * Initialize sample templates
   */
  private initializeSampleTemplates(): void {
    const sampleTemplate: Template = {
      id: 'sample-business-pro',
      name: 'Professional Business Report',
      description: 'A comprehensive business report template with charts and professional styling',
      category: this.categories.get('business')!,
      tags: ['business', 'report', 'charts', 'professional'],
      author: {
        id: 'author-1',
        name: 'Design Pro',
        email: 'design@example.com',
        verified: true,
        followers: 1250,
        templates: 15
      },
      version: '1.0.0',
      createdAt: new Date('2024-01-15'),
      updatedAt: new Date('2024-01-15'),
      downloads: 342,
      rating: 4.7,
      reviews: [],
      price: 0,
      license: 'free',
      preview: {
        thumbnail: '/templates/business-pro-thumb.jpg',
        screenshots: ['/templates/business-pro-1.jpg', '/templates/business-pro-2.jpg'],
        demoData: {
          project_title: 'Q4 2024 Business Report',
          name: 'John Doe',
          status_update: 'Excellent progress this quarter',
          agenda: 'Financial Results, Strategic Initiatives, Team Updates'
        },
        interactive: true
      },
      files: [
        {
          name: 'business-report-template.pptx',
          type: 'pptx',
          size: 2048576,
          url: '/templates/business-pro.pptx',
          checksum: 'abc123'
        }
      ],
      metadata: {
        placeholders: [
          { name: 'project_title', type: 'text', required: true, description: 'Main project title' },
          { name: 'name', type: 'text', required: true, description: 'Presenter name' },
          { name: 'status_update', type: 'text', required: false, description: 'Status information' },
          { name: 'agenda', type: 'text', required: false, description: 'Agenda items' }
        ],
        slides: 8,
        charts: 3,
        images: 5,
        animations: true,
        aspectRatio: '16:9',
        colorScheme: ['#1e40af', '#dc2626', '#059669'],
        fonts: ['Inter', 'Roboto'],
        requirements: ['PowerPoint 2016 or later'],
        features: ['Animated transitions', 'Interactive charts', 'Professional layout']
      },
      compatibility: {
        powerpointVersions: ['2016', '2019', '2021', '365'],
        officeVersions: ['2016', '2019', '2021', '365'],
        platforms: ['windows', 'mac', 'web'],
        minPowerPointVersion: '2016',
        testedVersions: ['2016', '2019', '2021', '365']
      }
    }

    this.templates.set(sampleTemplate.id, sampleTemplate)
  }
}

/**
 * Template Builder for creating new templates
 */
export class TemplateBuilder {
  private template: Partial<Template> = {}

  /**
   * Set basic template info
   */
  setInfo(name: string, description: string, category: TemplateCategory): TemplateBuilder {
    this.template.name = name
    this.template.description = description
    this.template.category = category
    return this
  }

  /**
   * Set author
   */
  setAuthor(author: Author): TemplateBuilder {
    this.template.author = author
    return this
  }

  /**
   * Set pricing
   */
  setPricing(price: number, license: LicenseType): TemplateBuilder {
    this.template.price = price
    this.template.license = license
    return this
  }

  /**
   * Add tags
   */
  addTags(...tags: string[]): TemplateBuilder {
    this.template.tags = [...(this.template.tags || []), ...tags]
    return this
  }

  /**
   * Set preview
   */
  setPreview(preview: TemplatePreview): TemplateBuilder {
    this.template.preview = preview
    return this
  }

  /**
   * Add file
   */
  addFile(file: TemplateFile): TemplateBuilder {
    this.template.files = [...(this.template.files || []), file]
    return this
  }

  /**
   * Set metadata
   */
  setMetadata(metadata: TemplateMetadata): TemplateBuilder {
    this.template.metadata = metadata
    return this
  }

  /**
   * Set compatibility
   */
  setCompatibility(compatibility: CompatibilityInfo): TemplateBuilder {
    this.template.compatibility = compatibility
    return this
  }

  /**
   * Build template
   */
  build(): Omit<Template, 'id' | 'createdAt' | 'updatedAt' | 'downloads' | 'rating' | 'reviews'> {
    if (!this.template.name || !this.template.description || !this.template.category) {
      throw new Error('Missing required template information')
    }

    return {
      name: this.template.name,
      description: this.template.description,
      category: this.template.category,
      tags: this.template.tags || [],
      author: this.template.author!,
      version: '1.0.0',
      price: this.template.price || 0,
      license: this.template.license || 'free',
      preview: this.template.preview!,
      files: this.template.files || [],
      metadata: this.template.metadata!,
      compatibility: this.template.compatibility!
    }
  }
}

/**
 * Template Validator
 */
export class TemplateValidator {
  /**
   * Validate template data
   */
  static validateTemplate(template: Partial<Template>): { valid: boolean; errors: string[] } {
    const errors: string[] = []

    if (!template.name || template.name.trim().length === 0) {
      errors.push('Template name is required')
    }

    if (!template.description || template.description.trim().length === 0) {
      errors.push('Template description is required')
    }

    if (!template.category) {
      errors.push('Template category is required')
    }

    if (!template.author) {
      errors.push('Template author is required')
    }

    if (!template.preview) {
      errors.push('Template preview is required')
    }

    if (!template.files || template.files.length === 0) {
      errors.push('At least one template file is required')
    }

    if (!template.metadata) {
      errors.push('Template metadata is required')
    }

    if (!template.compatibility) {
      errors.push('Template compatibility information is required')
    }

    // Validate files
    if (template.files) {
      template.files.forEach((file, index) => {
        if (!file.name || file.name.trim().length === 0) {
          errors.push(`File ${index + 1}: Name is required`)
        }
        if (!file.type) {
          errors.push(`File ${index + 1}: Type is required`)
        }
        if (!file.url) {
          errors.push(`File ${index + 1}: URL is required`)
        }
      })
    }

    // Validate metadata
    if (template.metadata) {
      if (!template.metadata.placeholders || template.metadata.placeholders.length === 0) {
        errors.push('At least one placeholder is required')
      }

      if (template.metadata.slides <= 0) {
        errors.push('Number of slides must be greater than 0')
      }
    }

    return {
      valid: errors.length === 0,
      errors
    }
  }

  /**
   * Validate placeholder
   */
  static validatePlaceholder(placeholder: PlaceholderInfo): { valid: boolean; errors: string[] } {
    const errors: string[] = []

    if (!placeholder.name || placeholder.name.trim().length === 0) {
      errors.push('Placeholder name is required')
    }

    if (!placeholder.type) {
      errors.push('Placeholder type is required')
    }

    if (!placeholder.description || placeholder.description.trim().length === 0) {
      errors.push('Placeholder description is required')
    }

    // Validate validation rules
    if (placeholder.validation) {
      const validation = placeholder.validation
      if (validation.minLength !== undefined && validation.minLength < 0) {
        errors.push('Minimum length cannot be negative')
      }
      if (validation.maxLength !== undefined && validation.maxLength < 0) {
        errors.push('Maximum length cannot be negative')
      }
      if (validation.min !== undefined && validation.max !== undefined && validation.min > validation.max) {
        errors.push('Minimum value cannot be greater than maximum value')
      }
    }

    return {
      valid: errors.length === 0,
      errors
    }
  }
}