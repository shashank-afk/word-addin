Office.onReady(() => {
  const btn = document.getElementById("insertTextBtn");
  if (!btn) return;

  btn.onclick = async () => {
    try {
      const response = await fetch("https://www.misrut.com/papi/opn", {
        method: "POST",
        headers: {
          "Accept": "application/json",
          "Content-Type": "application/json"
        },
        body: JSON.stringify({
          workflow: "addin",
          action: "testing"
        })
      });

      const data = await response.json();
      console.log("API response:", data);

      // Extract actual payload
      const payload = data.DATA;

      await Word.run(async (context) => {
        const body = context.document.body;
        
        // Insert each line as a separate paragraph
        body.insertParagraph(
          `Message: ${payload.message}`,
          Word.InsertLocation.end
        );
        
        body.insertParagraph(
          `Timestamp: ${payload.timestamp}`,
          Word.InsertLocation.end
        );
        
        body.insertParagraph(
          `Random: ${payload.random}`,
          Word.InsertLocation.end
        );
        
        body.insertParagraph(
          `Final Message: ${payload["final message"]}`,
          Word.InsertLocation.end
        );
        
        await context.sync();
      });

    } catch (error) {
      console.error("Error:", error);
      alert(error.message);
    }
  };
});