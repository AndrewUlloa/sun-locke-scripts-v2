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
 * Needed to expose the function to the client-side code
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile("sidebar")
    .setTitle("Sun Locke")
    .setWidth(300);
}
