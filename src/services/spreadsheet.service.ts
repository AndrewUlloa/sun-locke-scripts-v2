import 'google-apps-script';

interface SpreadsheetRange {
  sheet: string;
  column: string;
  startRow: number | 'auto';
  rowCount: number | 'all';
}

interface ProcessingResult {
  success: boolean;
  message?: string;
  data?: any[];
}

export class SpreadsheetService {
  /**
   * Gets all sheet names from the active spreadsheet
   */
  static getSheetNames(): string[] {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    return spreadsheet.getSheets().map(sheet => sheet.getName());
  }

  /**
   * Gets all column letters from a specific sheet
   */
  static getColumnLetters(sheetName: string): string[] {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) return [];
    
    const lastColumn = sheet.getLastColumn();
    return Array.from({ length: lastColumn }, (_, i) => 
      this.columnToLetter(i + 1)
    );
  }

  /**
   * Gets data from a specific range
   */
  static getDataFromRange(range: SpreadsheetRange): string[] {
    try {
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      const sheet = spreadsheet.getSheetByName(range.sheet);
      if (!sheet) throw new Error(`Sheet ${range.sheet} not found`);

      // Get the header row value
      const headerRow = this.getHeaderRow(range.sheet);
      
      // If in auto mode, start from header row + 1
      const effectiveStartRow = range.startRow === 'auto' ? headerRow + 1 : range.startRow;

      // Calculate the actual number of rows to process
      const lastRow = sheet.getLastRow();
      const availableRows = lastRow - effectiveStartRow + 1;
      const effectiveRowCount = range.rowCount === 'all' 
        ? availableRows 
        : Math.min(range.rowCount as number, availableRows);

      // Get the column data
      const columnIndex = this.letterToColumn(range.column);
      const rangeA1 = sheet.getRange(
        effectiveStartRow,
        columnIndex,
        effectiveRowCount,
        1
      );

      return rangeA1.getValues().map(row => row[0]?.toString() || '');
    } catch (error: unknown) {
      console.error('Error getting data from range:', error);
      return [];
    }
  }

  /**
   * Gets the header row number from sheet properties or defaults to 1
   */
  private static getHeaderRow(sheetName: string): number {
    try {
      const userProperties = PropertiesService.getUserProperties();
      const headerRowProp = userProperties.getProperty(`${sheetName}_headerRow`);
      return headerRowProp ? parseInt(headerRowProp) : 1;
    } catch (error) {
      console.error('Error getting header row:', error);
      return 1;
    }
  }

  /**
   * Sets the header row number for a sheet
   */
  static setHeaderRow(sheetName: string, headerRow: number): void {
    try {
      const userProperties = PropertiesService.getUserProperties();
      userProperties.setProperty(`${sheetName}_headerRow`, headerRow.toString());
    } catch (error) {
      console.error('Error setting header row:', error);
    }
  }

  /**
   * Writes data to a specific range
   */
  static writeDataToRange(range: SpreadsheetRange, data: string[]): ProcessingResult {
    try {
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      const sheet = spreadsheet.getSheetByName(range.sheet);
      if (!sheet) throw new Error(`Sheet ${range.sheet} not found`);

      // Get the header row value
      const headerRow = this.getHeaderRow(range.sheet);
      
      // If in auto mode, start from header row + 1
      const effectiveStartRow = range.startRow === 'auto' ? headerRow + 1 : range.startRow;

      // Calculate the actual number of rows to process
      const lastRow = sheet.getLastRow();
      const availableRows = lastRow - effectiveStartRow + 1;
      const effectiveRowCount = range.rowCount === 'all' 
        ? availableRows 
        : Math.min(range.rowCount as number, availableRows);

      const columnIndex = this.letterToColumn(range.column);
      const rangeA1 = sheet.getRange(
        effectiveStartRow,
        columnIndex,
        Math.min(effectiveRowCount, data.length),
        1
      );

      const values = data.slice(0, effectiveRowCount).map(value => [value]);
      rangeA1.setValues(values);

      return {
        success: true,
        message: `Successfully wrote ${values.length} rows of data`
      };
    } catch (error: unknown) {
      console.error('Error writing data to range:', error);
      return {
        success: false,
        message: error instanceof Error ? error.message : 'Unknown error occurred'
      };
    }
  }

  /**
   * Gets column headers from a specific sheet
   */
  static getColumnHeaders(sheetName: string, headerRow: number = 1): Map<string, string> {
    try {
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      const sheet = spreadsheet.getSheetByName(sheetName);
      if (!sheet) return new Map();

      const lastColumn = sheet.getLastColumn();
      const headers = sheet.getRange(headerRow, 1, 1, lastColumn).getValues()[0];
      
      return new Map(
        headers.map((header, index) => [
          this.columnToLetter(index + 1),
          header?.toString() || ''
        ])
      );
    } catch (error: unknown) {
      console.error('Error getting column headers:', error);
      return new Map();
    }
  }

  /**
   * Converts column letter to number (e.g., 'A' -> 1, 'B' -> 2)
   */
  private static letterToColumn(letter: string): number {
    let column = 0;
    const length = letter.length;
    for (let i = 0; i < length; i++) {
      column += (letter.charCodeAt(i) - 64) * Math.pow(26, length - i - 1);
    }
    return column;
  }

  /**
   * Converts column number to letter (e.g., 1 -> 'A', 2 -> 'B')
   */
  private static columnToLetter(column: number): string {
    let temp = column;
    let letter = '';
    while (temp > 0) {
      const remainder = (temp - 1) % 26;
      letter = String.fromCharCode(65 + remainder) + letter;
      temp = (temp - remainder - 1) / 26;
    }
    return letter;
  }
} 