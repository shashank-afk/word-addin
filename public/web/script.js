Office.onReady(() => {
  console.log("Office.js is ready");

  document.getElementById("insertTextBtn").onclick = async () => {
    try {
      await Word.run(async (context) => {
        const body = context.document.body;
        body.insertParagraph("Hello from your test add-in!", Word.InsertLocation.end);
        await context.sync();
      });
    } catch (error) {
      console.error("Error inserting text:", error);
    }
  };
});
