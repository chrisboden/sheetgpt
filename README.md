# GPT Google Sheets Integration

This Google Apps Script allows you to leverage all of the LLM models aggregated by Openrouter from directly within Google Sheets. You can set a default model and API key through the custom menu and override the default model for specific prompts as needed. Openrouter usually has free models you can use. Or try the excellent and very cheap 'openai/gpt-4o-mini-2024-07-18'

## Features

- **Dynamic AI Responses**: Generate AI-driven responses directly in your Google Sheets using a simple `GPT` function.
- **Multiple Model Support**: Use different AI models for different prompts by specifying the model in the function.
- **Default Model Configuration**: Set a default model that will be used across all prompts unless explicitly overridden.
- **Customizable Prompts with Multiple Roles**: Craft complex prompts dynamically by concatenating cell values and utilizing `system`, `user`, and `assistant` roles.

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
   - Enter the desired default model (e.g., `openai/gpt-4o-mini-2024-07-18`). Browse all openrouter models [here](https://openrouter.ai/models)

## Usage

### Basic Usage with Default Model

Use the default model set in the settings:

```plaintext
=GPT("system:you are an expert historian.;user:list the 5 major geopolitical events for the year "&A8)
```

### Override the Model for a Specific Prompt

Specify a different model for a specific prompt:

```plaintext
=GPT("system:you are an olympics data analyst;user:provide the total number of Olympic medals at the Paris Olympics for "&A8, "perplexity/llama-3.1-sonar-huge-128k-online")
```

### Multi-Message Array Example

The `GPT` function supports multi-message arrays, allowing you to simulate a conversation with the AI by using `system`, `user`, and `assistant` roles.

**Example 1: Basic Conversation**

When you don't specify a role, the script assume your prompt is for the user role. Example

```plaintext
=GPT("list the 5 major geopolitical events for the year "&A8)
```

**Example 2: Simple Multi-Message Conversation**

```plaintext
=GPT("system:you are a pirate;user:Write a poem about "&A8)
```

**Example 3: Dynamic Multi-Message Array with Cell References**

Dynamically create a multi-message array using data from cells:

```plaintext
=GPT("system:you are an expert stock analyst;user:what is the latest news for "&A7&"and is it a buy or sell rating from you? Answer strictly and exactly with either 'buy' or 'sell'","perplexity/llama-3.1-sonar-huge-128k-online")
```

**Example 4: Prompt chaining

Chain the result of one prompt to another:

| A      | B                                                                                  | C |
|--------------|---------------------------------------------------------------------------------------------------|-------------|
| South Korea  | South Korea won a total of 32 medals at the 2024 Paris Olympics, consisting of 13 gold, 9 silver, and 10 bronze medals. | 13          |



```plaintext
=GPT("user:how many olympic gold medals did "&A10&"win at the Paris Olympics in 2024. Reply only with the correct number","perplexity/llama-3.1-sonar-large-128k-online")
```

```plaintext
=GPT("system:you are a world class data cleanser;user:parse out this statement to find the number of gold medals. reply only with the number. Statement: "&B14)
```

These examples demonstrate the flexibility of the `GPT` function in handling both simple and complex interactions directly within your Google Sheets.

## Troubleshooting

- **Invalid or Missing API Key/Model**: Ensure that you have set the API key and model correctly in the `GPT Settings` menu.
- **Error Messages**: If the function returns an error, check the authorization status or the validity of the model specified.

## License

This script is free to use and modify under the MIT License.

# Prompt Playground

The script has been integrated into [this Google Sheet](https://docs.google.com/spreadsheets/d/196kX19rz7vH-aRvuiavzXR8KxlI3ukT0G3RHzKDX8w0/edit?usp=sharing) that teaches you how to do advanced prompting. The sheet is adapted from the Claude for Sheets workbook from Anthropic. You can use the claude LLM's via OpenRouter if you choose.
