// Get API key from script properties
const API_KEY = PropertiesService.getScriptProperties().getProperty("API_KEY");
const SEARCH_ENGINE_ID = "f791de3ddf13c4413";

/**
 * Gets search suggestions from Google Custom Search API
 * @param {string} query - The search query
 * @return {Array} Array of suggestion strings
 */
function getSearchSuggestions(query) {
  if (!query || query.length < 2) return [];
  if (!API_KEY) {
    console.error("API_KEY not found in script properties");
    return [];
  }

  try {
    const url = `https://www.googleapis.com/customsearch/v1?key=${API_KEY}&cx=${SEARCH_ENGINE_ID}&q=${encodeURIComponent(
      query
    )}`;

    const response = UrlFetchApp.fetch(url);
    const data = JSON.parse(response.getContentText());

    // Extract relevant suggestions from the search results
    const suggestions = [];

    if (data.items) {
      data.items.forEach((item) => {
        // Extract title and snippet, clean them up
        const title = item.title.replace(/\s+/g, " ").trim();
        const snippet = item.snippet?.replace(/\s+/g, " ").trim();

        // Add if they're relevant prompt-like phrases
        if (title.toLowerCase().includes(query.toLowerCase())) {
          // Only add if it looks like a prompt (starts with action verbs)
          if (
            /^(write|create|generate|translate|summarize|extract|analyze|format|describe)/i.test(
              title
            )
          ) {
            suggestions.push(title);
          }
        }
        if (snippet && snippet.toLowerCase().includes(query.toLowerCase())) {
          // Split into sentences and add relevant ones that look like prompts
          const sentences = snippet
            .split(/[.!?]+/)
            .filter((s) => s.trim().length > 0);
          sentences.forEach((sentence) => {
            const trimmedSentence = sentence.trim();
            if (
              trimmedSentence.toLowerCase().includes(query.toLowerCase()) &&
              /^(write|create|generate|translate|summarize|extract|analyze|format|describe)/i.test(
                trimmedSentence
              )
            ) {
              suggestions.push(trimmedSentence);
            }
          });
        }
      });
    }

    // Add spelling suggestions if available and relevant
    if (
      data.spelling?.correctedQuery &&
      /^(write|create|generate|translate|summarize|extract|analyze|format|describe)/i.test(
        data.spelling.correctedQuery
      )
    ) {
      suggestions.push(data.spelling.correctedQuery);
    }

    // Clean up suggestions
    return [...new Set(suggestions)]
      .filter((s) => {
        // Must be longer than query
        if (s.length <= query.length) return false;
        // Must look like a prompt
        if (
          !/^(write|create|generate|translate|summarize|extract|analyze|format|describe)/i.test(
            s
          )
        )
          return false;
        // Remove any HTML or special characters
        return s.replace(/[<>]/g, "").length === s.length;
      })
      .map((s) => s.replace(/\s+/g, " ").trim()) // Clean up whitespace
      .slice(0, 5); // Limit to top 5
  } catch (error) {
    console.error("Error fetching Google suggestions:", error);
    return [];
  }
}

/**
 * Gets all sheet names from the active spreadsheet
 */
function getSheetNames() {
  return SpreadsheetService.getSheetNames();
}

/**
 * Gets all column letters from a specific sheet
 */
function getColumnLetters(sheetName) {
  return SpreadsheetService.getColumnLetters(sheetName);
}

/**
 * Gets column headers and letters from a specific sheet
 * @param {string} sheetName - Name of the sheet
 * @param {number} headerRow - Row number containing headers
 * @return {Object} Object containing headers and column letters
 */
function getColumnHeaders(sheetName, headerRow) {
  const headers = SpreadsheetService.getColumnHeaders(sheetName, headerRow);
  const letters = SpreadsheetService.getColumnLetters(sheetName);

  return {
    headers: Object.fromEntries(headers),
    letters: letters,
  };
}

/**
 * Processes a custom prompt for a range of spreadsheet cells
 */
function processCustomPrompt(config) {
  return PromptService.processCustomPrompt(config);
}

/**
 * Needed to expose the function to the client-side code
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile("sidebar")
    .setTitle("Sun Locke")
    .setWidth(300);
}

/**
 * Gets all sheet data including names and columns in one call
 * @return {Object} Object containing all sheet data
 */
function getAllSheetData() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = spreadsheet.getSheets();

    return {
      sheets: sheets.map((sheet) => ({
        name: sheet.getName(),
        columns: SpreadsheetService.getColumnLetters(sheet.getName()),
      })),
    };
  } catch (error) {
    console.error("Error getting all sheet data:", error);
    throw new Error("Failed to load spreadsheet data");
  }
}

function setOpenAIKey() {
  // DO NOT hardcode your API key here. Instead, use the Google Apps Script Project Properties
  // To set your API key:
  // 1. Go to your Google Apps Script project
  // 2. Click on "Project Settings" (the gear icon)
  // 3. Click on "Script Properties"
  // 4. Click "Add Script Property"
  // 5. Set "Property" as "OPENAI_API_KEY"
  // 6. Set "Value" as your OpenAI API key
  // 7. Click "Save"

  // This function is just for documentation. The actual key should be set in Project Settings.
  const apiKey =
    PropertiesService.getScriptProperties().getProperty("OPENAI_API_KEY");
  if (!apiKey) {
    throw new Error(
      "OpenAI API key not found in Script Properties. Please add it in Project Settings."
    );
  }
  return true;
}

// Add a test function
function testLanguageModel() {
  const config = {
    modelType: "language",
    inputSheet: "Sheet1", // Replace with your actual sheet name
    inputColumn: "A",
    outputSheet: "Sheet1",
    outputColumn: "B",
    startRow: 2,
    rowCount: 1,
    prompt: "Write a short greeting",
    systemInstructions: "You are a friendly assistant.",
    model: "gpt-4",
  };

  try {
    const result = PromptService.processCustomPrompt(config);
    Logger.log("Test result:", result);
    return result;
  } catch (error) {
    Logger.log("Test error:", error);
    return { success: false, message: error.toString() };
  }
}
