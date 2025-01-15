import { SpreadsheetService } from './spreadsheet.service';

interface PromptConfig {
  modelType: 'language' | 'search' | 'image';
  inputSheet: string;
  inputColumn: string;
  outputSheet: string;
  outputColumn: string;
  startRow: number;
  rowMode: 'fixed' | 'all' | '3rows';
  rowCount?: number;
  prompt: string;
  systemInstructions?: string;
  model: string;
}

interface PromptResult {
  success: boolean;
  message?: string;
  results?: string[];
}

export class PromptService {
  /**
   * Processes a custom prompt for a range of spreadsheet cells
   */
  static async processCustomPrompt(config: PromptConfig): Promise<PromptResult> {
    try {
      // Determine if this should be a web search based on prompt content
      const shouldUseSearch = config.prompt.toLowerCase().includes('search the web');
      const effectiveModelType = shouldUseSearch ? 'search' : config.modelType;
      
      // If using search, ensure a valid search model is selected
      if (shouldUseSearch && !config.model.includes('llama-3.1-sonar')) {
        // Default to small model if no valid search model is selected
        config.model = 'llama-3.1-sonar-small-128k-online';
      }

      // Validate model type
      if (!['language', 'search', 'image'].includes(effectiveModelType)) {
        throw new Error('Invalid model type');
      }

      // Validate start row
      if (!config.startRow || config.startRow < 1) {
        throw new Error('Invalid start row');
      }

      console.log('Processing prompt with config:', {
        startRow: config.startRow,
        rowCount: config.rowCount,
        rowMode: config.rowMode,
        modelType: effectiveModelType,
        model: config.model,
        isWebSearch: shouldUseSearch
      });

      // Determine effective row count based on mode
      let effectiveRowCount: number | 'all';
      if (config.rowMode === 'all') {
        effectiveRowCount = 'all';
      } else if (config.rowMode === '3rows') {
        effectiveRowCount = 3;
      } else {
        // Fixed mode - use rowCount or default to 1
        effectiveRowCount = Math.max(1, config.rowCount || 1);
      }

      console.log('Effective row count:', effectiveRowCount);

      // Get input data from spreadsheet using the fixed start row
      const inputData = SpreadsheetService.getDataFromRange({
        sheet: config.inputSheet,
        column: config.inputColumn,
        startRow: config.startRow, // Use the fixed start row value
        rowCount: effectiveRowCount
      });

      console.log('Retrieved input data from row', config.startRow, ':', inputData);

      if (!inputData.length) {
        return {
          success: false,
          message: 'No input data found in the specified range'
        };
      }

      // Process each input with the AI model - only process the number of rows specified
      const rowsToProcess = effectiveRowCount === 'all' ? inputData.length : effectiveRowCount;
      const results = await Promise.all(
        inputData.slice(0, rowsToProcess).map(async (input, index) => {
          const currentRow = config.startRow + index;
          console.log(`Processing row ${currentRow} (${index + 1} of ${rowsToProcess})`);
          
          // Combine the user's prompt with the cell content
          const cellContent = input.trim();
          const combinedPrompt = `${config.prompt}\n\nContent to process: ${cellContent}`;
          console.log('Combined prompt for row', currentRow, ':', combinedPrompt);

          // Call AI model with combined prompt
          const result = await this.callAIModel(
            combinedPrompt,
            config.systemInstructions,
            config.model,
            effectiveModelType
          );

          return result;
        })
      );

      console.log('Processed results:', results);

      // Write results back to spreadsheet using the same fixed start row
      const writeResult = SpreadsheetService.writeDataToRange(
        {
          sheet: config.outputSheet,
          column: config.outputColumn,
          startRow: config.startRow, // Use the same fixed start row
          rowCount: effectiveRowCount
        },
        results
      );

      if (!writeResult.success) {
        throw new Error(writeResult.message);
      }

      return {
        success: true,
        message: `Successfully processed ${results.length} rows starting from row ${config.startRow}`,
        results
      };
    } catch (error) {
      console.error('Error processing custom prompt:', error);
      return {
        success: false,
        message: error instanceof Error ? error.message : 'Unknown error occurred'
      };
    }
  }

  /**
   * Calls the AI model with the given prompt
   * This is a placeholder - implement actual AI model call
   */
  private static async callAIModel(
    prompt: string,
    systemInstructions?: string,
    model?: string,
    modelType: string = 'language'
  ): Promise<string> {
    // Ensure model is defined
    if (!model) {
      throw new Error('Model must be specified');
    }

    switch (modelType) {
      case 'language':
        return this.callLanguageModel(prompt, systemInstructions, model);
      case 'search':
        return this.callSearchModel(prompt, systemInstructions, model);
      case 'image':
        // TODO: Implement image model
        throw new Error('Image model not implemented yet');
      default:
        throw new Error(`Unknown model type: ${modelType}`);
    }
  }

  private static async callLanguageModel(
    prompt: string,
    systemInstructions: string = "You are a helpful assistant.",
    model: string
  ): Promise<string> {
    try {
      // Debug: Log all script properties
      const scriptProperties = PropertiesService.getScriptProperties();
      const allProperties = scriptProperties.getProperties();
      console.log('Available script properties:', Object.keys(allProperties));
      
      // Try both property names
      const openaiKey = scriptProperties.getProperty('OPENAI_API_KEY');
      const apiKey = scriptProperties.getProperty('API_KEY');
      
      console.log('OPENAI_API_KEY exists:', !!openaiKey);
      console.log('API_KEY exists:', !!apiKey);
      
      // Use OPENAI_API_KEY if available, otherwise try API_KEY
      const API_KEY = openaiKey || apiKey || '';
      
      if (!API_KEY) {
        throw new Error('OpenAI API key not found in script properties. Please check both OPENAI_API_KEY and API_KEY properties.');
      }

      // Validate inputs
      if (!prompt || prompt.trim() === '') {
        throw new Error('Prompt cannot be empty');
      }

      // Ensure systemInstructions has a default value if null/undefined
      const safeSystemInstructions = systemInstructions || "You are a helpful assistant.";

      // Use provided model or default to gpt-4o-mini
      const modelToUse = model || 'gpt-4o-mini';

      // Debug log the request payload
      const payload = {
        model: modelToUse,
        messages: [
          { role: "system", content: safeSystemInstructions },
          { role: "user", content: prompt.trim() }
        ],
        temperature: 0.7,
        max_tokens: 1000
      };
      console.log('Request payload:', JSON.stringify(payload));

      const response = await UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', {
        method: 'post',
        headers: {
          'Authorization': `Bearer ${API_KEY}`,
          'Content-Type': 'application/json'
        },
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      });

      const result = JSON.parse(response.getContentText());
      
      if (result.error) {
        console.error('OpenAI API error:', result.error);
        throw new Error(result.error.message);
      }

      return result.choices[0].message.content;
    } catch (error) {
      console.error('Error calling OpenAI API:', error);
      throw new Error(`AI Model Error: ${error instanceof Error ? error.message : 'Unknown error'}`);
    }
  }

  private static async callSearchModel(
    prompt: string,
    systemInstructions: string = "You are a helpful assistant.",
    model: string
  ): Promise<string> {
    try {
      // Get Perplexity API key
      const scriptProperties = PropertiesService.getScriptProperties();
      const perplexityKey = scriptProperties.getProperty('PERPLEXITY_API_KEY');
      
      if (!perplexityKey) {
        throw new Error('Perplexity API key not found in script properties. Please add PERPLEXITY_API_KEY to script properties.');
      }

      // Validate inputs
      if (!prompt || prompt.trim() === '') {
        throw new Error('Prompt cannot be empty');
      }

      // Validate model is one of the supported Perplexity models
      const validModels = [
        'llama-3.1-sonar-small-128k-online',
        'llama-3.1-sonar-large-128k-online',
        'llama-3.1-sonar-huge-128k-online'
      ];
      
      if (!validModels.includes(model)) {
        throw new Error(`Invalid Perplexity model. Must be one of: ${validModels.join(', ')}`);
      }

      // Note: According to docs, search models don't attend to system prompts
      // but we'll include it in the request for consistency
      const payload = {
        model: model,
        messages: [
          { role: "system", content: systemInstructions },
          { role: "user", content: prompt.trim() }
        ],
        temperature: 0.7,
        max_tokens: 1000
      };

      console.log('Perplexity request payload:', JSON.stringify(payload));

      const response = await UrlFetchApp.fetch('https://api.perplexity.ai/chat/completions', {
        method: 'post',
        headers: {
          'Authorization': `Bearer ${perplexityKey}`,
          'Content-Type': 'application/json'
        },
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      });

      const result = JSON.parse(response.getContentText());
      
      if (result.error) {
        console.error('Perplexity API error:', result.error);
        throw new Error(result.error.message);
      }

      return result.choices[0].message.content;
    } catch (error) {
      console.error('Error calling Perplexity API:', error);
      throw new Error(`Search Model Error: ${error instanceof Error ? error.message : 'Unknown error'}`);
    }
  }
} 