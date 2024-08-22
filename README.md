# GPT Google Sheets Integration

This Google Apps Script allows you to leverage AI models like GPT-3.5 and GPT-4 directly within Google Sheets. You can set a default model and API key through the custom menu and override the default model for specific prompts as needed.

## Features

- **Dynamic AI Responses**: Generate AI-driven responses directly in your Google Sheets using a simple `GPT` function.
- **Multiple Model Support**: Use different AI models for different prompts by specifying the model in the function.
- **Default Model Configuration**: Set a default model that will be used across all prompts unless explicitly overridden.
- **Customizable Prompts**: Build complex prompts dynamically by concatenating cell values.

## Installation

1. **Open Google Sheets**:
   - Go to `Extensions` > `Apps Script`.

2. **Add the Script**:
   - Replace the existing code in the `Code.gs` file with the provided script.

3. **Save the Script**:
   - Click the floppy disk icon or press `Ctrl+S` to save the script.

4. **Authorize the Script**:
   - You may be prompted to authorize the script to run. Follow the on-screen instructions to complete the authorization.

## Configuration

1. **Set API Key**:
   - In Google Sheets, click `GPT Settings` > `Set API Key`.
   - Enter your OpenRouter API key.

2. **Set Default Model**:
   - In Google Sheets, click `GPT Settings` > `Set Model`.
   - Enter the desired default model (e.g., `openai/gpt-4`).

## Usage

### Basic Usage with Default Model

Use the default model set in the settings:

```plaintext
=GPT("Provide the total number of Olympic medals at the Paris Olympics for " & A8 & ".")
```

### Override the Model for Specific Prompts

Specify a different model for a specific prompt:

```plaintext
=GPT("Provide the total number of Olympic medals at the Paris Olympics for " & A8 & ".", "perplexity/llama-3.1-sonar-huge-128k-online")
```

### Example Prompts

- **Simple Prompt**:
  ```plaintext
  =GPT("Summarize the following article: " & A1)
  ```
- **Complex Prompt with Model Override**:
  ```plaintext
  =GPT("Analyze the financial data for the year " & B2 & ".", "openai/gpt-3.5-turbo")
  ```

## Troubleshooting

- **Invalid or Missing API Key/Model**: Ensure that you have set the API key and model correctly in the `GPT Settings` menu.
- **Error Messages**: If the function returns an error, check the authorization status or the validity of the model specified.

## License

This script is free to use and modify under the MIT License.
