// Send message and call API
async function sendMessage() {
  const text = messageInput.value.trim();
  if (!text && attachedFiles.length === 0) return;

  // Show user message in chat
  addMessageToChat('user', text, attachedFiles.map(f => f.name));

  try {
    console.log("Starting to process files...", attachedFiles);

    // Prepare files with base64 content
    const filesData = [];
    for (const file of attachedFiles) {
      try {
        console.log("Converting file:", file.name);
        const base64Content = await fileToBase64(file);
        console.log("File converted successfully, length:", base64Content.length);
        
        filesData.push({
          name: file.name,
          size: file.size,
          type: file.type,
          content: base64Content
        });
      } catch (fileError) {
        console.error("Error converting file:", file.name, fileError);
        addMessageToChat('assistant', `Error reading file ${file.name}: ${fileError.message}`);
        return; // Stop if file conversion fails
      }
    }

    console.log("All files converted, count:", filesData.length);

    // Add to conversation history
    conversationHistory.push({
      role: 'user',
      message: text,
      files: filesData
    });

    console.log("Files being sent:", filesData);

    // Call your backend API
    const response = await fetch("https://www.misrut.com/papi/opn", {
      method: "POST",
      headers: {
        "Accept": "application/json",
        "Content-Type": "application/json"
      },
      body: JSON.stringify({
        workflow: "addin",
        action: "testing",
        USER_INPUT: text,
        files: filesData,
        conversation_history: conversationHistory
      })
    });

    const data = await response.json();
    console.log("API response:", data);

    // Extract payload
    const payload = data.DATA;

    // Add AI response to conversation history and insert into Excel
    if (payload.ai_reply) {
      conversationHistory.push({
        role: 'assistant',
        message: payload.ai_reply
      });

      // Insert ai_reply into Excel - split by newlines and put each in separate rows
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        
        // Get the used range to find the next empty row
        const usedRange = sheet.getUsedRange();
        usedRange.load("rowCount");
        await context.sync();
        
        // Split the AI reply by newlines to get individual items
        const items = payload.ai_reply
          .split('\n')
          .map(item => item.trim())
          .filter(item => item.length > 0);

        // Insert each item in a new row, column A
        items.forEach((item, index) => {
          const nextRow = usedRange.rowCount + index;
          const targetCell = sheet.getCell(nextRow, 0); // Column A (index 0)
          targetCell.values = [[item]];
        });
        
        await context.sync();
      });

      // Show AI reply in chat
      addMessageToChat('assistant', payload.ai_reply);
    }

  } catch (error) {
    console.error("Error:", error);
    addMessageToChat('assistant', `Error: ${error.message}`);
  }

  // Clear input
  messageInput.value = '';
  messageInput.style.height = 'auto';
  attachedFiles = [];
  renderAttachedFiles();
  updateSendButton();
}