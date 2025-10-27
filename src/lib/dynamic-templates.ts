/**
 * Dynamic Template System for PowerPoint AutoFill
 * Supports advanced placeholders, conditional logic, and data transformation
 */

export interface TemplateConfig {
  placeholders: Record<string, PlaceholderConfig>;
  conditionals: ConditionalConfig[];
  loops: LoopConfig[];
  calculations: CalculationConfig[];
  formatting: FormattingConfig;
}

export interface PlaceholderConfig {
  type: 'text' | 'number' | 'date' | 'image' | 'chart';
  required: boolean;
  default?: any;
  format?: string;
  transform?: (value: any) => any;
}

export interface ConditionalConfig {
  id: string;
  condition: string;
  trueContent: string;
  falseContent?: string;
}

export interface LoopConfig {
  id: string;
  dataSource: string;
  template: string;
  separator?: string;
}

export interface CalculationConfig {
  id: string;
  formula: string;
  alias: string;
  format?: string;
}

export interface FormattingConfig {
  dateFormat: string;
  numberFormat: string;
  currency: string;
  locale: string;
}

export class DynamicTemplateEngine {
  private config: TemplateConfig;
  private data: Record<string, any>;

  constructor(config: TemplateConfig) {
    this.config = config;
  }

  /**
   * Process template with dynamic data
   */
  processTemplate(template: string, data: Record<string, any>): string {
    this.data = data;
    
    let processed = template;
    
    // Process conditionals first
    processed = this.processConditionals(processed);
    
    // Process loops
    processed = this.processLoops(processed);
    
    // Process calculations
    processed = this.processCalculations(processed);
    
    // Process basic placeholders
    processed = this.processPlaceholders(processed);
    
    // Process formatting
    processed = this.processFormatting(processed);
    
    return processed;
  }

  /**
   * Process conditional placeholders {if:condition}content{endif}
   */
  private processConditionals(template: string): string {
    const conditionalRegex = /\{if:([^}]+)\}([\s\S]*?)\{endif\}/g;
    
    return template.replace(conditionalRegex, (match, condition, content) => {
      const hasElse = content.includes('{else}');
      const [trueContent, falseContent] = hasElse 
        ? content.split(/\{else\}/) 
        : [content, ''];
      
      return this.evaluateCondition(condition.trim()) ? trueContent : falseContent;
    });
  }

  /**
   * Process loop placeholders {loop:dataSource}template{endloop}
   */
  private processLoops(template: string): string {
    const loopRegex = /\{loop:([^}]+)\}([\s\S]*?)\{endloop\}/g;
    
    return template.replace(loopRegex, (match, dataSource, loopTemplate) => {
      const data = this.getDataValue(dataSource.trim());
      
      if (!Array.isArray(data)) {
        console.warn(`Loop data source '${dataSource}' is not an array`);
        return '';
      }
      
      const separator = this.config.loops
        .find(loop => loop.dataSource === dataSource.trim())
        ?.separator || '\n';
      
      return data
        .map(item => this.processLoopTemplate(loopTemplate, item))
        .join(separator);
    });
  }

  /**
   * Process calculation placeholders {calc:formula}
   */
  private processCalculations(template: string): string {
    const calcRegex = /\{calc:([^}]+)\}/g;
    
    return template.replace(calcRegex, (match, formula) => {
      try {
        const result = this.evaluateFormula(formula.trim());
        return this.formatValue(result, 'number');
      } catch (error) {
        console.error(`Calculation error for formula '${formula}':`, error);
        return '0';
      }
    });
  }

  /**
   * Process basic placeholders {placeholder}
   */
  private processPlaceholders(template: string): string {
    const placeholderRegex = /\{([^}]+)\}/g;
    
    return template.replace(placeholderRegex, (match, placeholder) => {
      // Skip already processed special placeholders
      if (placeholder.startsWith('if:') || 
          placeholder.startsWith('loop:') || 
          placeholder.startsWith('calc:') ||
          placeholder.startsWith('date:') ||
          placeholder.startsWith('image:')) {
        return match;
      }
      
      const value = this.getDataValue(placeholder.trim());
      return value !== undefined ? String(value) : match;
    });
  }

  /**
   * Process formatted placeholders {date:format='YYYY-MM-DD'}
   */
  private processFormatting(template: string): string {
    // Date formatting
    const dateRegex = /\{date:format='([^']+)'\}/g;
    template = template.replace(dateRegex, (match, format) => {
      return this.formatDate(new Date(), format);
    });
    
    // Image placeholders
    const imageRegex = /\{image:([^:]+):?([^}]*)\}/g;
    template = template.replace(imageRegex, (match, imageKey, options) => {
      return this.processImagePlaceholder(imageKey, options);
    });
    
    return template;
  }

  /**
   * Evaluate conditional expressions
   */
  private evaluateCondition(condition: string): boolean {
    // Simple condition evaluation (can be enhanced)
    const operators = ['==', '!=', '>', '<', '>=', '<=', 'has', 'empty'];
    
    for (const op of operators) {
      if (condition.includes(op)) {
        const [left, right] = condition.split(op).map(s => s.trim());
        const leftValue = this.getDataValue(left);
        const rightValue = right.startsWith("'") ? right.slice(1, -1) : this.getDataValue(right);
        
        switch (op) {
          case '==': return leftValue == rightValue;
          case '!=': return leftValue != rightValue;
          case '>': return Number(leftValue) > Number(rightValue);
          case '<': return Number(leftValue) < Number(rightValue);
          case '>=': return Number(leftValue) >= Number(rightValue);
          case '<=': return Number(leftValue) <= Number(rightValue);
          case 'has': return Array.isArray(leftValue) && leftValue.includes(rightValue);
          case 'empty': return !leftValue || leftValue.length === 0;
        }
      }
    }
    
    // Simple truthy/falsy evaluation
    return Boolean(this.getDataValue(condition));
  }

  /**
   * Process loop template with item data
   */
  private processLoopTemplate(template: string, item: any): string {
    const itemData = { ...this.data, ...item, item };
    const tempEngine = new DynamicTemplateEngine(this.config);
    return tempEngine.processTemplate(template, itemData);
  }

  /**
   * Evaluate mathematical formulas
   */
  private evaluateFormula(formula: string): number {
    // Safe formula evaluation (can be enhanced with proper parser)
    const safeFormula = formula.replace(/[^0-9+\-*/().\s]/g, '');
    
    try {
      // Replace data references in formula
      const processedFormula = safeFormula.replace(/\b([a-zA-Z_][a-zA-Z0-9_]*)\b/g, (match) => {
        const value = this.getDataValue(match);
        return isNaN(value) ? '0' : String(value);
      });
      
      return Function(`"use strict"; return (${processedFormula})`)();
    } catch (error) {
      throw new Error(`Invalid formula: ${formula}`);
    }
  }

  /**
   * Get value from data using dot notation
   */
  private getDataValue(path: string): any {
    const parts = path.split('.');
    let current = this.data;
    
    for (const part of parts) {
      if (current === null || current === undefined) {
        return undefined;
      }
      current = current[part];
    }
    
    return current;
  }

  /**
   * Format value based on type
   */
  private formatValue(value: any, type: string): string {
    if (value === null || value === undefined) {
      return '';
    }
    
    switch (type) {
      case 'number':
        return Number(value).toLocaleString(this.config.formatting.locale);
      case 'currency':
        return new Intl.NumberFormat(this.config.formatting.locale, {
          style: 'currency',
          currency: this.config.formatting.currency
        }).format(Number(value));
      case 'date':
        return this.formatDate(new Date(value), this.config.formatting.dateFormat);
      default:
        return String(value);
    }
  }

  /**
   * Format date with custom format
   */
  private formatDate(date: Date, format: string): string {
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const day = String(date.getDate()).padStart(2, '0');
    const hours = String(date.getHours()).padStart(2, '0');
    const minutes = String(date.getMinutes()).padStart(2, '0');
    
    return format
      .replace('YYYY', String(year))
      .replace('MM', month)
      .replace('DD', day)
      .replace('HH', hours)
      .replace('mm', minutes);
  }

  /**
   * Process image placeholders
   */
  private processImagePlaceholder(imageKey: string, options: string): string {
    const imageData = this.getDataValue(imageKey);
    
    if (!imageData) {
      return `[Image: ${imageKey}]`;
    }
    
    // Parse options (width, height, etc.)
    const opts = this.parseOptions(options);
    
    return `[Image: ${imageKey} ${JSON.stringify(opts)}]`;
  }

  /**
   * Parse options string
   */
  private parseOptions(optionsStr: string): Record<string, any> {
    const options: Record<string, any> = {};
    
    if (!optionsStr) {
      return options;
    }
    
    optionsStr.split(',').forEach(opt => {
      const [key, value] = opt.split('=').map(s => s.trim());
      if (key && value) {
        options[key] = value.replace(/['"]/g, '');
      }
    });
    
    return options;
  }
}

/**
 * Default template configuration
 */
export const defaultTemplateConfig: TemplateConfig = {
  placeholders: {
    project_title: { type: 'text', required: true },
    name: { type: 'text', required: true },
    date: { type: 'date', required: false },
    status_update: { type: 'text', required: false },
    agenda: { type: 'text', required: false },
    ChartTitle: { type: 'text', required: false }
  },
  conditionals: [],
  loops: [],
  calculations: [],
  formatting: {
    dateFormat: 'YYYY-MM-DD',
    numberFormat: 'en-US',
    currency: 'USD',
    locale: 'en-US'
  }
};

/**
 * Example usage
 */
export const exampleUsage = {
  template: `
    {project_title}
    
    {if:has_charts}
    Charts Available:
    {loop:charts}• {name}: {value}{endloop}
    {else}
    No charts available
    {endif}
    
    Total: {calc:revenue * 1.2}
    Date: {date:format='MMMM DD, YYYY'}
  `,
  
  data: {
    project_title: 'Q4 2024 Review',
    has_charts: true,
    charts: [
      { name: 'Revenue', value: 150000 },
      { name: 'Costs', value: 80000 }
    ],
    revenue: 100000
  }
};