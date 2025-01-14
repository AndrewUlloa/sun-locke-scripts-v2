import { SpreadsheetService } from './spreadsheet.service';

interface PromptConfig {
  inputSheet: string;
  inputColumn: string;
  outputSheet: string;
  outputColumn: string;
  startRow: number;
  rowCount: number;
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
      // Get input data from spreadsheet
      const inputData = SpreadsheetService.getDataFromRange({
        sheet: config.inputSheet,
        column: config.inputColumn,
        startRow: config.startRow,
        rowCount: config.rowCount
      });

      if (!inputData.length) {
        return {
          success: false,
          message: 'No input data found in the specified range'
        };
      }

      // Process each input with the AI model
      const results = await Promise.all(
        inputData.map(async (input) => {
          // Replace variables in prompt
          const processedPrompt = config.prompt.replace(/\{\{input\}\}/g, input);

          // Call AI model (this is a placeholder - implement actual AI call)
          const result = await this.callAIModel(
            processedPrompt,
            config.systemInstructions,
            config.model
          );

          return result;
        })
      );

      // Write results back to spreadsheet
      const writeResult = SpreadsheetService.writeDataToRange(
        {
          sheet: config.outputSheet,
          column: config.outputColumn,
          startRow: config.startRow,
          rowCount: config.rowCount
        },
        results
      );

      if (!writeResult.success) {
        throw new Error(writeResult.message);
      }

      return {
        success: true,
        message: `Successfully processed ${results.length} rows`,
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
    model?: string
  ): Promise<string> {
    // TODO: Implement actual AI model call
    // For now, return a mock response
    return `AI response for: ${prompt}`;
  }
} 