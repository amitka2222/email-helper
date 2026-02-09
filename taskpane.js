/* global document, Office, fetch, localStorage, window */

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    // 1. Hook up the buttons
    document.getElementById("save-key-btn").onclick = saveApiKey;
    document.getElementById("clear-key-btn").onclick = clearApiKey;
    document.getElementById("reply-btn").onclick = function() { runAI("reply"); };
    document.getElementById("new-mail-btn").onclick = function() { runAI("new"); };
    document.getElementById("insert-btn").onclick = insertText;

    // 2. Check if we already have a key
    checkApiKey(); 
  }
});

// --- SETTINGS ---
function saveApiKey() {
    const keyInput = document.getElementById("api-key-input");
    const key = keyInput.value;
    
    if (key && key.trim() !== "") {
        localStorage.setItem("myGeminiKey", key.trim());
        checkApiKey(); // Refresh the view
    } else {
        keyInput.style.border = "2px solid red";
    }
}

function checkApiKey() {
    const key = localStorage.getItem("myGeminiKey");
    if (key) {
        document.getElementById("settings-area").style.display = "none";
        document.getElementById("main-area").style.display = "block";
    } else {
        document.getElementById("settings-area").style.display = "block";
        document.getElementById("main-area").style.display = "none";
    }
}

function clearApiKey() {
    localStorage.removeItem("myGeminiKey");
    location.reload();
}

// --- AI LOGIC ---
async function runAI(mode) {
  const resultArea = document.getElementById("result-area");
  resultArea.value = "Thinking...";

  Office.context.mailbox.item.body.getAsync("text", async function (result) {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      const textContext = result.value;
      const apiKey = localStorage.getItem("myGeminiKey");

      let systemInstruction = `
        You are a busy professional assistant.
        STRICT RULES:
        1. Length: Short, concise, direct.
        2. Language: South African English (colour, programme, centre).
        3. Tone: Professional but human. No "I hope this finds you well".
        4. No Signature: The user has an auto-signature.
        5. Formatting: Plain text only.
      `;

      let userPrompt = "";
      if (mode === "reply") {
          userPrompt = `Task: Write a direct reply. Incoming Email: "${textContext}"`;
      } else {
          userPrompt = `Task: Write a NEW email based on notes: "${textContext}". 
          Format: Subject: [Subject Here] \n\n [Body Here]`;
      }

      const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=${apiKey}`;

      try {
        const response = await fetch(url, {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({
            contents: [{ parts: [{ text: systemInstruction + "\n\n" + userPrompt }] }]
          })
        });

        const data = await response.json();

        if (data.error) {
            resultArea.value = "API Error: " + data.error.message;
            return;
        }
        
        if (!data.candidates || !data.candidates.length) {
             resultArea.value = "No response.";
             return;
        }

        let finalText = data.candidates[0].content.parts[0].text;

        if (mode === "new" && finalText.includes("Subject:")) {
            const match = finalText.match(/Subject: (.*)/);
            if (match) {
                Office.context.mailbox.item.subject.setAsync(match[1]);
                finalText = finalText.replace(/Subject: .*\n+/, "").trim();
            }
        }
        resultArea.value = finalText.trim();

      } catch (error) {
        resultArea.value = "Network Error: " + error.message;
      }
    }
  });
}

function insertText() {
  const text = document.getElementById("result-area").value;
  Office.context.mailbox.item.body.setSelectedDataAsync(text, { coercionType: Office.CoercionType.Text });
}