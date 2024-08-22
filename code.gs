function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('GPT Settings')
    .addItem('Set API Key', 'setApiKey')
    .addItem('Set Model', 'setModel')
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

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('GPT Settings')
    .addItem('Set API Key', 'setApiKey')
    .addItem('Set Model', 'setModel')
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

  // Construct the API request payload
  const payload = {
    "model": selectedModel,
    "messages": [
      {"role": "user", "content": prompt}
    ]
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
