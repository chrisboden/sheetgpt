function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('GPT Settings')
    .addItem('Set API Key', 'setApiKey')
    .addItem('Set Model', 'setModel')
    .addItem('Add Messages Note', 'addMessagesNote')
    .addItem('Clear Cache', 'clearCache')
    .addToUi();
}

function setApiKey() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('Enter your OpenRouter API Key');
  const apiKey = response.getResponseText();
  PropertiesService.getUserProperties().setProperty('OPENROUTER_API_KEY', apiKey);
  ui.alert('API Key has been set.');
}

function setModel() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('Enter the GPT Model');
  const model = response.getResponseText();
  PropertiesService.getUserProperties().setProperty('MODEL', model);
  ui.alert('Model has been set to: ' + model);
}

function GPT(prompt, model, bypassCache = false) {
  const properties = PropertiesService.getUserProperties();
  const apiKey = properties.getProperty('OPENROUTER_API_KEY');
  const selectedModel = model || properties.getProperty('MODEL');

  if (!apiKey || !selectedModel) {
    Logger.log("Error: API Key or Model is not set.");
    return 'Error: API Key or Model is not set. Please set them using the GPT Settings menu.';
  }

  // Generate a unique cache key based on the prompt and model
  const cacheKey = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, prompt + selectedModel));
  
  // Check if the result is already in the cache and bypass is not requested
  if (!bypassCache) {
    const cachedResult = CacheService.getUserCache().get(cacheKey);
    if (cachedResult != null) {
      return cachedResult;
    }
  }

  let messages;
  if (/system\s*:/i.test(prompt) || /user\s*:/i.test(prompt) || /assistant\s*:/i.test(prompt)) {
    messages = parseMessages(prompt);
  } else {
    messages = [{
      "role": "user",
      "content": prompt.trim()
    }];
  }

  storeMessagesInProperties(messages);

  const payload = {
    "model": selectedModel,
    "messages": messages
  };

  const options = {
    "method": "post",
    "contentType": "application/json",
    "headers": {
      "Authorization": `Bearer ${apiKey}`
    },
    "payload": JSON.stringify(payload)
  };

  try {
    Logger.log("API Request Payload: " + JSON.stringify(payload, null, 2));
    const response = UrlFetchApp.fetch("https://openrouter.ai/api/v1/chat/completions", options);
    const jsonResponse = JSON.parse(response.getContentText());
    Logger.log("API Response: " + JSON.stringify(jsonResponse, null, 2));
    const result = jsonResponse.choices[0].message.content.trim();
    
    // Cache the result for future use
    CacheService.getUserCache().put(cacheKey, result, 21600); // Cache for 6 hours
    
    return result;
  } catch (error) {
    Logger.log("Error: " + error.message);
    return `Error: ${error.message}`;
  }
}

function parseMessages(prompt) {
  const messages = [];

  // Define the role patterns with case-insensitivity and optional space after the colon
  const rolePatterns = {
    "system": /^system\s*:\s*/i,
    "user": /^user\s*:\s*/i,
    "assistant": /^assistant\s*:\s*/i
  };

  // Split the prompt into parts based on the roles
  const parts = prompt.split(';');

  parts.forEach(part => {
    part = part.trim();
    let matched = false;

    for (const [role, pattern] of Object.entries(rolePatterns)) {
      if (pattern.test(part)) {
        messages.push({
          "role": role,
          "content": part.replace(pattern, '').trim()
        });
        matched = true;
        break;
      }
    }

    // If no specific role is matched, treat the entire part as a user message
    if (!matched) {
      messages.push({
        "role": "user",
        "content": part.trim()
      });
    }
  });

  return messages;
}

function storeMessagesInProperties(messages) {
  const activeCell = SpreadsheetApp.getActiveSpreadsheet().getActiveRange();
  const cellAddress = activeCell.getA1Notation();
  const messagesJson = JSON.stringify(messages, null, 2); // Pretty-print JSON
  PropertiesService.getDocumentProperties().setProperty(cellAddress, messagesJson);
}

function addMessagesNote() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  const cell = sheet.getActiveRange();
  const cellAddress = cell.getA1Notation();
  const messagesJson = PropertiesService.getDocumentProperties().getProperty(cellAddress);

  // Fetch the execution logs
  const logs = Logger.getLog();
  const lastLogEntry = logs.split('\n').pop(); // Get the last log entry
  
  if (messagesJson) {
    const noteContent = `Messages JSON: ${messagesJson}\n\nLast Log Entry:\n${lastLogEntry}`;
    cell.setNote(noteContent);
  } else {
    SpreadsheetApp.getUi().alert("No message array found for this cell.");
  }
}

function clearCache() {
  CacheService.getUserCache().removeAll();
  SpreadsheetApp.getUi().alert('Cache has been cleared.');
}
