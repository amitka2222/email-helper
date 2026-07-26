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
    
    // Allow empty key for Local AI routing
    Office.context.roamingSettings.set("myGeminiKey", key.trim());
    Office.context.roamingSettings.set("myUserName", name.trim());
    Office.context.roamingSettings.saveAsync(function (result) {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
            console.error("Failed to save settings: " + result.error.message);
        }
    });
    checkSettings();
}

function checkSettings() {
    // As long as they clicked save, let them into the main area
    const name = Office.context.roamingSettings.get("myUserName");
    if (name) { 
        document.getElementById("settings-area").style.display = "none";
        document.getElementById("main-area").style.display = "block";
    } else {
        document.getElementById("settings-area").style.display = "block";
        document.getElementById("main-area").style.display = "none";
    }
}

function clearSettings() {
    Office.context.roamingSettings.remove("myGeminiKey");
    Office.context.roamingSettings.remove("myUserName");
    Office.context.roamingSettings.saveAsync(function () {
        location.reload();
    });
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

            const apiKey = Office.context.roamingSettings.get("myGeminiKey");
            const userName = Office.context.roamingSettings.get("myUserName") || "[My Name]";

            // --- STRICT PROMPT RULES ---
            let systemInstruction = `
              You are an expert email assistant writing on behalf of ${userName}. 
              Your goal is to write highly natural, human-sounding emails in professional South African English.
              
              CRITICAL RULES FOR TONE AND STYLE:
              1. NO AI TROPES: Never use phrases like "I hope this email finds you well", "Please do not hesitate to reach out", "Delve", "Moreover", "In conclusion", or "As per our previous".
              2. SOUTH AFRICAN ENGLISH: Use British/SA spelling (e.g., 'organise', 'colour', 'programme'). Be polite and warm, but direct and concise.
              3. HUMAN CONVERSATIONAL FLOW: Write as if you are quickly typing a reply. Do not sound like a robotic corporate template. Avoid overly complex vocabulary.
              4. Output Format: Return only Plain HTML (<p>, <br>).
              5. NO Markdown: Do NOT use markdown (* or **) or code blocks (\`\`\`).
              6. Greeting: Start naturally, e.g., "Hi [Name]", "Morning [Name]", or "Afternoon [Name]".
              7. Sign-off: End strictly with: <br><br>Kind regards,<br>${userName}
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

            // --- HYBRID ROUTING LOGIC ---
            let url;
            let fetchOptions;

            if (apiKey && apiKey !== "") {
                // Route to Google Cloud (For your friends)
                url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-flash-latest:generateContent?key=${apiKey}`;
                fetchOptions = {
                    method: "POST",
                    headers: { "Content-Type": "application/json" },
                    body: JSON.stringify({
                        contents: [{ parts: [{ text: systemInstruction + "\n\n" + userPrompt }] }]
                    })
                };
            } else {
                // Route to Local Ollama (For you)
                url = "http://127.0.0.1:11434/api/generate";
                fetchOptions = {
                    method: "POST",
                    headers: { "Content-Type": "application/json" },
                    body: JSON.stringify({
                        model: "gemma3:12b", 
                        prompt: systemInstruction + "\n\n" + userPrompt,
                        stream: false
                    })
                };
            }

            try {
                const response = await fetch(url, fetchOptions);
                const data = await response.json();
                
                let finalHtml = "";

                // Parse response based on which API was used
                if (apiKey && apiKey !== "") {
                    if (!data.candidates || !data.candidates.length) throw new Error("No response from Gemini.");
                    finalHtml = data.candidates[0].content.parts[0].text;
                } else {
                    if (!data.response) throw new Error("No response from Local AI.");
                    finalHtml = data.response;
                }
                
                // --- CLEANUP SCRIPT ---
                finalHtml = finalHtml.replace(/\*\*/g, ""); 
                finalHtml = finalHtml.replace(/```html/g, "").replace(/```/g, "");

                if (mode === "new" && finalHtml.includes("<h1>")) {
                    const subjectMatch = finalHtml.match(/<h1>(.*?)<\/h1>/);
                    if (subjectMatch) {
                        Office.context.mailbox.item.subject.setAsync(subjectMatch[1]);
                        finalHtml = finalHtml.replace(/<h1>.*?<\/h1>/, "");
                    }
                }

                const cleanHtml = typeof DOMPurify !== "undefined" ? DOMPurify.sanitize(finalHtml) : finalHtml;
                previewBox.innerHTML = cleanHtml;
                hiddenResult.value = cleanHtml;

            } catch (error) {
                previewBox.innerHTML = "Error processing request. If using Local AI, ensure Ollama is running and CORS is configured. Details: " + error.message;
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