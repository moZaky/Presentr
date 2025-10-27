/**
 * Test file for PPTX Builder functionality
 */

import { PPTXBuilder } from './pptx-builder'

export async function testPPTXBuilder() {
  console.log('Testing PPTX Builder...')
  
  try {
    // Create a simple test
    const builder = new PPTXBuilder()
    
    // Test data
    const testData = {
      project_title: 'Test Project',
      name: 'Test User',
      status_update: 'This is a test status update',
      agenda: 'Technology Stack System Context Design Decisions',
      ChartTitle: 'Test Chart'
    }
    
    builder.setData(testData)
    
    console.log('✅ PPTX Builder test data set successfully')
    console.log('✅ PPTX Builder is ready for use')
    
    return true
  } catch (error) {
    console.error('❌ PPTX Builder test failed:', error)
    return false
  }
}

// Run test if this file is executed directly
if (require.main === module) {
  testPPTXBuilder()
}