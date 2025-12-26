Office.onReady((info) => {
  console.log("Office.js is ready");

  const btn = document.getElementById("insertTextBtn");
  if (!btn) {
    console.error("Button not found!");
    return;
  }

  btn.onclick = async () => {
    console.log("Button clicked!");

    try {
      // 1️⃣ Call your backend API
      const response = await fetch("https://www.misrut.com/papi/opn", {
        method: "POST",
        headers: {
          "Accept": "application/json",
          "Content-Type": "application/json",

        },
        body: JSON.stringify({
          workflow: "addin",
          action: "testing"
        })
      });

      if (!response.ok) {
        throw new Error(`API failed with status ${response.status}`);
      }

      const data = await response.json();
      console.log("API response:", data);

      // 2️⃣ Decide what text you want to insert
      const textToInsert = `
Message: ${data.message}
Timestamp: ${data.timestamp}
Random: ${data.random}
Final Message: ${data["final message"]}
`;


      // 3️⃣ Insert into Word
      await Word.run(async (context) => {
        context.document.body.insertParagraph(
          textToInsert,
          Word.InsertLocation.end
        );
        await context.sync();
      });

    } catch (error) {
      console.error("Error:", error);
      alert("Something went wrong: " + error.message);
    }
  };
});
