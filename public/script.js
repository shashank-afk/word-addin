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

      // 🔑 extract actual payload
      const payload = data.DATA;

      const textToInsert = `
Message: ${payload.message}
Timestamp: ${payload.timestamp}
Random: ${payload.random}
Final Message: ${payload["final message"]}
`;

      await Word.run(async (context) => {
        context.document.body.insertParagraph(
          textToInsert,
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
