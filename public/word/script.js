let attachedFiles = [];

Office.onReady(() => {
  const messageInput = document.getElementById('messageInput');
  const sendButton = document.getElementById('sendButton');
  const fileButton = document.getElementById('fileButton');
  const fileInput = document.getElementById('fileInput');
  const inputWrapper = document.getElementById('inputWrapper');
  const attachedFilesContainer = document.getElementById('attachedFiles');
  const chatContainer = document.getElementById('chatContainer');

  // Auto-resize textarea
  messageInput.addEventListener('input', () => {
    messageInput.style.height = 'auto';
    messageInput.style.height = messageInput.scrollHeight + 'px';
    updateSendButton();
  });

  // Enable/disable send button
  function updateSendButton() {
    const hasText = messageInput.value.trim().length > 0;
    const hasFiles = attachedFiles.length > 0;
    sendButton.disabled = !hasText && !hasFiles;
  }

  // File button click
  fileButton.addEventListener('click', () => {
    fileInput.click();
  });

  // File input change
  fileInput.addEventListener('change', (e) => {
    handleFiles(Array.from(e.target.files));
    fileInput.value = '';
  });

  // Drag and drop
  inputWrapper.addEventListener('dragover', (e) => {
    e.preventDefault();
    inputWrapper.classList.add('drag-over');
  });

  inputWrapper.addEventListener('dragleave', () => {
    inputWrapper.classList.remove('drag-over');
  });

  inputWrapper.addEventListener('drop', (e) => {
    e.preventDefault();
    inputWrapper.classList.remove('drag-over');
    handleFiles(Array.from(e.dataTransfer.files));
  });

  // Handle files
  function handleFiles(files) {
    files.forEach(file => {
      attachedFiles.push(file);
    });
    renderAttachedFiles();
    updateSendButton();
  }

  // Render attached files
  function renderAttachedFiles() {
    attachedFilesContainer.innerHTML = '';
    attachedFiles.forEach((file, index) => {
      const chip = document.createElement('div');
      chip.className = 'file-chip';
      chip.innerHTML = `
        <svg fill="none" stroke="currentColor" viewBox="0 0 24 24">
          <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M7 21h10a2 2 0 002-2V9.414a1 1 0 00-.293-.707l-5.414-5.414A1 1 0 0012.586 3H7a2 2 0 00-2 2v14a2 2 0 002 2z"/>
        </svg>
        <span>${file.name}</span>
        <button class="remove-file" data-index="${index}">
          <svg fill="none" stroke="currentColor" viewBox="0 0 24 24">
            <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M6 18L18 6M6 6l12 12"/>
          </svg>
        </button>
      `;
      attachedFilesContainer.appendChild(chip);
    });

    // Add remove listeners
    document.querySelectorAll('.remove-file').forEach(btn => {
      btn.addEventListener('click', (e) => {
        const index = parseInt(e.currentTarget.dataset.index);
        attachedFiles.splice(index, 1);
        renderAttachedFiles();
        updateSendButton();
      });
    });
  }

  // Add message to chat UI
  function addMessageToChat(role, text, files = []) {
    const messageDiv = document.createElement('div');
    messageDiv.className = `message ${role}`;
    
    let content = text;
    if (files.length > 0) {
      const filesList = files.map(name => 
        `<div class="file-preview ${role === 'user' ? '' : 'dark'}">
          <svg fill="none" stroke="currentColor" viewBox="0 0 24 24" style="width:16px;height:16px">
            <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M7 21h10a2 2 0 002-2V9.414a1 1 0 00-.293-.707l-5.414-5.414A1 1 0 0012.586 3H7a2 2 0 00-2 2v14a2 2 0 002 2z"/>
          </svg>
          ${name}
        </div>`
      ).join('');
      content = filesList + (text ? '<div style="margin-top:8px">' + text + '</div>' : '');
    }
    
    messageDiv.innerHTML = content;
    chatContainer.appendChild(messageDiv);
    chatContainer.scrollTop = chatContainer.scrollHeight;
  }

  // Send message and call API
  async function sendMessage() {
    const text = messageInput.value.trim();
    if (!text && attachedFiles.length === 0) return;

    // Show user message in chat
    addMessageToChat('user', text, attachedFiles.map(f => f.name));

    try {
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
          message: text,
          files: attachedFiles.map(f => ({
            name: f.name,
            size: f.size,
            type: f.type
          }))
        })
      });

      const data = await response.json();
      console.log("API response:", data);

      // Extract payload
      const payload = data.DATA;

      // Insert ai_reply into Word document
      if (payload.ai_reply) {
        await Word.run(async (context) => {
          const body = context.document.body;
          body.insertParagraph(payload.ai_reply, Word.InsertLocation.end);
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

  // Send on Enter (without Shift)
  messageInput.addEventListener('keydown', (e) => {
    if (e.key === 'Enter' && !e.shiftKey) {
      e.preventDefault();
      if (!sendButton.disabled) {
        sendMessage();
      }
    }
  });

  // Send button click
  sendButton.addEventListener('click', sendMessage);
});