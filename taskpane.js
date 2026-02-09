/* global document, Office, fetch, localStorage, window */

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    checkApiKey(); // Check if user has already saved a key
    document.getElementById("save-key-btn").onclick = saveApiKey;
    document.getElementById("clear-key-btn").onclick = clearApiKey;
    
    // Button Event Listeners
    document.getElementById("reply-btn").onclick = () => runAI("reply");
    document.getElementById("new-mail-btn").onclick = () => runAI("new");
    document.getElementById("insert-btn").onclick = insertText;
  }
});

// --- SETTINGS: API KEY MANAGEMENT ---
function saveApiKey() {
    const key = document.getElementById("api-key-input").value;
    if (key && key.trim() !== "") {
        localStorage.setItem("myGeminiKey", key.trim());
        checkApiKey();
    } else {
        document.getElementById("api-key-input").style.border = "1px solid red";
    }
}

function checkApiKey() {
    const key = localStorage.getItem("myGeminiKey");
    if (key) {
        document.getElementById("settings-area").classList.add("hidden");
        document.getElementById("main-area").classList.remove("hidden");
    } else {
        document.getElementById("settings-area").classList.remove("hidden");
        document.getElementById("main-area").classList.add("hidden");
    }
}

function clearApiKey() {
    localStorage.removeItem("myGeminiKey");
    window.location.reload();
}

// --- AI GENERATION LOGIC ---
export async function runAI(mode) {
  const resultArea = document.getElementById("result-area");
  resultArea.value = "Thinking...";

  // Get content (either the email reply body OR your bullet points)
  Office.context.mailbox.item.body.getAsync("text", async function (result) {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      const textContext = result.value;
      const apiKey = localStorage.getItem("myGeminiKey");

      // --- SYSTEM INSTRUCTIONS (The "Brain") ---
      let systemInstruction = `
        You are a busy professional assistant.
        STRICT RULES:
        1.  **Length:** Extremely short, concise, and direct. No filler words.
        2.  **Language:** South African English (use 'colour', 'organise', 'programme', 'centre').
        3.  **Tone:** Professional but human. Avoid "AI" phrases like "I hope this email finds you well."
        4.  **Formatting:** Plain text only. NO bold (**), NO italics, NO markdown headers (#).
        5.  **Sign-off:** Do NOT include a signature (e.g., 'Kind regards', 'Name'). The user has an auto-signature.
        6.  **Structure:** Use bullet points if listing items.
      `;

      let userPrompt = "";

      if (mode === "reply") {
          userPrompt = `Task: Write a direct reply to this email.
          Incoming Email: "${textContext}"`;
      } else {
          userPrompt = `Task: Write a NEW email based on these rough notes: "${textContext}".
          Output Format:
          Subject: [Write a clear Subject Line here]
          
          [Write Email Body here]`;
      }

      // Using the Free Tier Model (1.5 Flash)
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
        
        if (!data.candidates || data.candidates.length === 0) {
             resultArea.value = "No response generated.";
             return;
        }

        let finalText = data.candidates[0].content.parts[0].text;

        // Clean up Subject Line if in "New Mail" mode
        if (mode === "new" && finalText.includes("Subject:")) {
            const subjectMatch = finalText.match(/Subject: (.*)/);
            if (subjectMatch) {
                const subjectLine = subjectMatch[1];
                finalText = finalText.replace(/Subject: .*\n+/, "").trim(); // Remove subject from body
                
                // Auto-fill the actual Subject field in Outlook
                Office.context.mailbox.item.subject.setAsync(subjectLine);
            }
        }

        resultArea.value = finalText.trim();

      } catch (error) {
        console.error(error);
        resultArea.value = "Network Error. Check internet connection.";
      }
    } else {
        resultArea.value = "Error: Could not read email content.";
    }
  });
}

function insertText() {
  const text = document.getElementById("result-area").value;
  Office.context.mailbox.item.body.setSelectedDataAsync(text, { coercionType: Office.CoercionType.Text }, function(asyncResult){
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          document.getElementById("result-area").value += `\n\n[ERROR]: ${asyncResult.error.message}\n(Tip: Click 'Reply' or 'New Email' in Outlook first!)`;
      }
  });
}