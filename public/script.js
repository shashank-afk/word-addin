Office.onReady((info) => {
  console.log("Office.js is ready");
  console.log("Host:", info.host); // Should show "Word"
  console.log("Platform:", info.platform);

  // Make sure the button exists before adding event listener
  const btn = document.getElementById("insertTextBtn");
  if (btn) {
    btn.onclick = async () => {
      console.log("Button clicked!");
      try {
        await Word.run(async (context) => {
          const body = context.document.body;
          body.insertParagraph("Hello from your test add-in!", Word.InsertLocation.end);
          await context.sync();
          console.log("Text inserted successfully!");
        });
      } catch (error) {
        console.error("Error inserting text:", error);
        alert("Error: " + error.message);
      }
    };
  } else {
    console.error("Button not found!");
  }
});