// @ts-nocheck
import 'google-apps-script';

/**
 * Runs when the add-on is installed.
 */
function onInstall(e) {
  onOpen(e);
}

/**
 * Runs when a spreadsheet that has this add-on is opened.
 */
function onOpen(e) {
  SpreadsheetApp.getUi()
    .createAddonMenu()
    .addItem('Open Sidebar', 'showSidebar')
    .addToUi();
}

function showSidebar() {
  const html = HtmlService.createTemplateFromFile('sidebar')
    .evaluate()
    .setTitle('AI For Sheets™ by Sun Locke')
    .setWidth(300);
  
  SpreadsheetApp.getUi().showSidebar(html);
}

// Get the user's selected models
function getSelectedModels() {
  const userProperties = PropertiesService.getUserProperties();
  return {
    llm: userProperties.getProperty('selectedLLM') || 'gpt-4',
    search: userProperties.getProperty('selectedSearch') || 'llama-3.1-sonar-small-128k-online'
  };
}

// Save the user's model selection
function setSelectedModel(type: string, modelId: string) {
  const userProperties = PropertiesService.getUserProperties();
  if (type === 'llm') {
    userProperties.setProperty('selectedLLM', modelId);
  } else if (type === 'search') {
    userProperties.setProperty('selectedSearch', modelId);
  }
  return true;
}

function showModelDialog(type: string, models: string[]) {
  const ui = SpreadsheetApp.getUi();
  const title = type === 'llm' ? 'Select Language Model' : 'Select Search Model';
  
  // Create the dialog message with model options
  const message = models.map((model, index) => 
    `${index + 1}. ${model}`
  ).join('\n');

  const response = ui.prompt(
    title,
    message + '\n\nEnter the number of your selection:',
    ui.ButtonSet.OK_CANCEL
  );

  // Handle the user's response
  if (response.getSelectedButton() === ui.Button.OK) {
    const selection = parseInt(response.getResponseText());
    if (selection > 0 && selection <= models.length) {
      return models[selection - 1];
    }
  }
  
  return null;
}

// Types for saved prompts
interface SavedPrompt {
  id: string;
  name: string;
  content: string;
  category: string;
  createdAt: string;
  lastUsed?: string;
}

// Get all saved prompts
function getSavedPrompts(): SavedPrompt[] {
  const userProperties = PropertiesService.getUserProperties();
  const savedPromptsJson = userProperties.getProperty('savedPrompts');
  if (!savedPromptsJson) return [];
  return JSON.parse(savedPromptsJson);
}

// Save a new prompt
function savePrompt(name: string, content: string, category: string = 'custom'): SavedPrompt {
  const prompts = getSavedPrompts();
  const newPrompt: SavedPrompt = {
    id: Utilities.getUuid(),
    name,
    content,
    category,
    createdAt: new Date().toISOString()
  };
  
  prompts.push(newPrompt);
  PropertiesService.getUserProperties().setProperty('savedPrompts', JSON.stringify(prompts));
  return newPrompt;
}

// Delete a saved prompt
function deletePrompt(promptId: string): boolean {
  const prompts = getSavedPrompts();
  const index = prompts.findIndex(p => p.id === promptId);
  if (index === -1) return false;
  
  prompts.splice(index, 1);
  PropertiesService.getUserProperties().setProperty('savedPrompts', JSON.stringify(prompts));
  return true;
}

// Update last used timestamp for a prompt
function updatePromptUsage(promptId: string): void {
  const prompts = getSavedPrompts();
  const prompt = prompts.find(p => p.id === promptId);
  if (!prompt) return;
  
  prompt.lastUsed = new Date().toISOString();
  PropertiesService.getUserProperties().setProperty('savedPrompts', JSON.stringify(prompts));
}

// Get a specific prompt by ID
function getPromptById(promptId: string): SavedPrompt | null {
  const prompts = getSavedPrompts();
  return prompts.find(p => p.id === promptId) || null;
}

// Search saved prompts
function searchPrompts(query: string, category?: string): SavedPrompt[] {
  const prompts = getSavedPrompts();
  return prompts.filter(prompt => {
    const matchesQuery = query ? 
      (prompt.name.toLowerCase().includes(query.toLowerCase()) || 
       prompt.content.toLowerCase().includes(query.toLowerCase())) 
      : true;
    
    const matchesCategory = category && category !== 'all' ? 
      prompt.category === category 
      : true;
    
    return matchesQuery && matchesCategory;
  });
}

// Sort saved prompts
function sortPrompts(prompts: SavedPrompt[], sortBy: string): SavedPrompt[] {
  return [...prompts].sort((a, b) => {
    switch (sortBy) {
      case 'name':
        return a.name.localeCompare(b.name);
      case 'category':
        return a.category.localeCompare(b.category);
      case 'newest':
        return new Date(b.createdAt).getTime() - new Date(a.createdAt).getTime();
      case 'oldest':
        return new Date(a.createdAt).getTime() - new Date(b.createdAt).getTime();
      case 'lastUsed':
        const aTime = a.lastUsed ? new Date(a.lastUsed).getTime() : 0;
        const bTime = b.lastUsed ? new Date(b.lastUsed).getTime() : 0;
        return bTime - aTime;
      default:
        return 0;
    }
  });
}

// Get filtered and sorted prompts
function getFilteredPrompts(query: string = '', category: string = 'all', sortBy: string = 'newest'): SavedPrompt[] {
  const filtered = searchPrompts(query, category === 'all' ? undefined : category);
  return sortPrompts(filtered, sortBy);
} 