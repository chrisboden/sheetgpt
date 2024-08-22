function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('GPT Settings')
    .addItem('Set API Key', 'setApiKey')
    .addItem('Set Model', 'setModel')
    .addItem('Add Messages Note', 'addMessagesNote')
    .addToUi();
}

function setApiKey() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('Enter your OpenRouter API Key');
  const apiKey = response.getResponseText();
  PropertiesService.getScriptProperties().setProperty('OPENROUTER_API_KEY', apiKey);
  ui.alert('API Key has been set.');
}

function setModel() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('Enter the GPT Model');
  const model = response.getResponseText();
  PropertiesService.getScriptProperties().setProperty('MODEL', model);
  ui.alert('Model has been set to: ' + model);
}

function GPT(prompt, model) {
  // Fetch API key from properties
  const properties = PropertiesService.getScriptProperties();
  const apiKey = properties.getProperty('OPENROUTER_API_KEY');
  
  // Use the provided model or fall back to the default model
  const selectedModel = model || properties.getProperty('MODEL');

  if (!apiKey || !selectedModel) {
    return 'Error: API Key or Model is not set. Please set them using the GPT Settings menu.';
  }

  // If no specific role is provided, assume 'user' role for the prompt
  let messages;
  if (prompt.includes("system:") || prompt.includes("user:") || prompt.includes("assistant:")) {
    messages = parseMessages(prompt);
  } else {
    messages = [{
      "role": "user",
      "content": prompt.trim()
    }];
  }

  // Store the messages array as a note using a separate function
  storeMessagesInProperties(messages);

  // Construct the API request payload
  const payload = {
    "model": selectedModel,
    "messages": messages
  };

  // Make the API request
  const options = {
    "method": "post",
    "contentType": "application/json",
    "headers": {
      "Authorization": `Bearer ${apiKey}`
    },
    "payload": JSON.stringify(payload)
  };

  try {
    const response = UrlFetchApp.fetch("https://openrouter.ai/api/v1/chat/completions", options);
    const jsonResponse = JSON.parse(response.getContentText());

    // Return the AI response content
    return jsonResponse.choices[0].message.content.trim();
  } catch (error) {
    return `Error: ${error.message}`;
  }
}

function parseMessages(prompt) {
  const messages = [];
  
  // Split the prompt based on the role delimiters
  const parts = prompt.split(';');
  
  for (let i = 0; i < parts.length; i++) {
    const part = parts[i].trim();
    
    if (part.startsWith("system:")) {
      messages.push({
        "role": "system",
        "content": part.replace("system:", "").trim()
      });
    } else if (part.startsWith("user:")) {
      messages.push({
        "role": "user",
        "content": part.replace("user:", "").trim()
      });
    } else if (part.startsWith("assistant:")) {
      messages.push({
        "role": "assistant",
        "content": part.replace("assistant:", "").trim()
      });
    }
  }

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
  
  if (messagesJson) {
    cell.setNote(messagesJson);
  } else {
    SpreadsheetApp.getUi().alert("No message array found for this cell.");
  }
}
