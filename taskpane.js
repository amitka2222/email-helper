/* global document, Office, fetch, localStorage, window */

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("save-key-btn").onclick = saveSettings;
    document.getElementById("clear-key-btn").onclick = clearSettings;
    document.getElementById("reply-btn").onclick = function() { runAI("reply"); };
    document.getElementById("new-mail-btn").onclick = function() { runAI("new"); };
    document.getElementById("insert-btn").onclick = insertHtml;

    checkSettings(); 
  }
});

// --- SETTINGS ---
function saveSettings() {
    const key = document.getElementById("api-key-input").value;
    const name = document.getElementById("user-name-input").value;
    
    if (key && key.trim() !== "") {
        localStorage.setItem("myGeminiKey", key.trim());
        localStorage.setItem("myUserName", name.trim());
        checkSettings();
    } else {
        alert("Please enter an API Key.");
    }
}

function checkSettings() {
    const key = localStorage.getItem("myGeminiKey");
    if (key) {
        document.getElementById("settings-area").style.display = "none";
        document.getElementById("main-area").style.display = "block";
    } else {
        document.getElementById("settings-area").style.display = "block";
        document.getElementById("main-area").style.display = "none";
    }
}

function clearSettings() {
    localStorage.removeItem("myGeminiKey");
    localStorage.removeItem("myUserName");
    location.reload();
}

// --- AI LOGIC ---
async function runAI(mode) {
  const previewBox = document.getElementById("preview-box");
  const hiddenResult = document.getElementById("hidden-result");
  previewBox.innerHTML = "<i>Reading your email...</i>";

  Office.context.mailbox.item.getSelectedDataAsync(Office.CoercionType.Text, async function (asyncResult) {
      let userNotes = "";
      let fullContext = "";

      if (asyncResult.status === Office.AsyncResultStatus.Succeeded && asyncResult.value.data) {
          userNotes = asyncResult.value.data;
      }

      Office.context.mailbox.item.body.getAsync("text", async function (result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
            fullContext = result.value;
            
            if (!userNotes) {
                userNotes = "No specific notes highlighted. Infer intent from email thread.";
            }

            const apiKey = localStorage.getItem("myGeminiKey");
            const userName = localStorage.getItem("myUserName") || "[My Name]";

            // --- PROMPT ENGINEERING (FIXED FOR NO BOLD) ---
            let systemInstruction = `
              You are an expert email assistant.
              RULES:
              1. **Output Format:** Plain HTML (<p>, <br>). 
              2. **NO Markdown:** Do not use bold (**) or markdown. 
              3. **Formatting:** Use standard paragraphs.
              4. **Greeting:** Start with "Hi [Name]" or "Hi All".
              5. **Sign-off:** End strictly with: <br><br>Kind regards,<br>${userName}
              6. **Style:** Professional South African English. Concise.
            `;

            let userPrompt = "";
            if (mode === "reply") {
                userPrompt = `
                CONTEXT: I am replying to an email chain.
                MY ROUGH NOTES: "${userNotes}"
                FULL HISTORY: "${fullContext.substring(0, 2000)}"
                TASK: Write a reply.
                `;
            } else {
                userPrompt = `
                TASK: Write a NEW email based on: "${fullContext}".
                FORMAT: Return Subject in <h1>, then body.
                `;
            }

            const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-flash-latest:generateContent?key=${apiKey}`;

            try {
                const response = await fetch(url, {
                  method: "POST",
                  headers: { "Content-Type": "application/json" },
                  body: JSON.stringify({
                    contents: [{ parts: [{ text: systemInstruction + "\n\n" + userPrompt }] }]
                  })
                });

                const data = await response.json();
                
                if (!data.candidates || !data.candidates.length) {
                     previewBox.innerHTML = "Error: No response from AI.";
                     return;
                }

                let finalHtml = data.candidates[0].content.parts[0].text;
                
                // Cleanup Markdown to be safe
                finalHtml = finalHtml.replace(/\*\*/g, ""); // Removes bold ** markers
                finalHtml = finalHtml.replace(/```html/g, "").replace(/```/g, "");

                if (mode === "new" && finalHtml.includes("<h1>")) {
                    const subjectMatch = finalHtml.match(/<h1>(.*?)<\/h1>/);
                    if (subjectMatch) {
                        Office.context.mailbox.item.subject.setAsync(subjectMatch[1]);
                        finalHtml = finalHtml.replace(/<h1>.*?<\/h1>/, "");
                    }
                }

                previewBox.innerHTML = finalHtml;
                hiddenResult.value = finalHtml;

            } catch (error) {
                previewBox.innerHTML = "Network Error: " + error.message;
            }
        }
      });
  });
}

function insertHtml() {
  const html = document.getElementById("hidden-result").value;
  if (!html) return;

  Office.context.mailbox.item.body.setSelectedDataAsync(
      html, 
      { coercionType: Office.CoercionType.Html }, 
      function(result) {
          if (result.status === Office.AsyncResultStatus.Failed) {
              Office.context.mailbox.item.body.setAsync(
                  html, 
                  { coercionType: Office.CoercionType.Html }
              );
          }
      }
  );
}