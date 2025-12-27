// Storage for chat history shared across function calls
let chatHistory = [];

// Register the custom function
CustomFunctions.associate("ASKADDIN", askAddin);
console.log("functions.js LOADED!");

CustomFunctions.associate("ASKADDIN", askAddin);
console.log("ASKADDIN function associated!");

/**
 * Ask the add-in AI a question with chat history
 * @customfunction
 * @param {string} question The question to ask
 * @returns {Promise<string>} The AI's response
 */
async function askAddin(question) {
  try {
    console.log("askAddin called with:", question);
    return "Response: " + question;
    // Get user's Microsoft token for security
    const token = await OfficeRuntime.auth.getAccessToken({
      allowSignInPrompt: true,
      allowConsentPrompt: true
    });

    // Call your backend API with chat history
    const response = await fetch("https://your-backend-api.com/api/endpoint", {
      method: "POST",
      headers: {
        "Authorization": `Bearer ${token}`,
        "Content-Type": "application/json"
      },
      body: JSON.stringify({
        USER_INPUT: question,
        CONVERSATION_HISTORY: chatHistory
      })
    });

    if (!response.ok) {
      throw new Error(`API Error: ${response.status}`);
    }

    const data = await response.json();

    // Extract just the AI reply
    const aiReply = data.DATA.ai_reply;

    // Update chat history for next call
    chatHistory = data.DATA.conversation_history || chatHistory;

    // Return clean response to cell
    return aiReply;

  } catch (error) {
    console.error("Custom function error:", error);
    return `#ERROR: ${error.message}`;
  }
}

/**
 * Reset the chat history
 * @customfunction
 * @returns {string} Confirmation message
 */
function resetChat() {
  chatHistory = [];
  return "Chat history cleared";
}

// Register reset function too
CustomFunctions.associate("RESETCHAT", resetChat);