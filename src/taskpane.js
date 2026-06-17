Office.onReady(function (info) {
  if (info.host === Office.HostType.Outlook) {
    init();
  }
});

function init() {
  var providerSelect = document.getElementById("provider");
  var openaiKeyInput = document.getElementById("openai-key");
  var geminiKeyInput = document.getElementById("gemini-key");
  var jargonSelect = document.getElementById("jargon-level");
  var insertModeSelect = document.getElementById("insert-mode");
  var translateBtn = document.getElementById("translate-btn");
  var insertBtn = document.getElementById("insert-btn");
  var copyBtn = document.getElementById("copy-btn");

  openaiKeyInput.value = localStorage.getItem("corp-xlator-openai-key") || "";
  geminiKeyInput.value = localStorage.getItem("corp-xlator-gemini-key") || "";
  providerSelect.value = localStorage.getItem("corp-xlator-provider") || "openai";
  jargonSelect.value = localStorage.getItem("corp-xlator-jargon-level") || "medium";
  insertModeSelect.value = localStorage.getItem("corp-xlator-insert-mode") || "both";

  toggleKeyFields(providerSelect.value);

  providerSelect.addEventListener("change", function () {
    localStorage.setItem("corp-xlator-provider", providerSelect.value);
    toggleKeyFields(providerSelect.value);
  });

  openaiKeyInput.addEventListener("change", function () {
    localStorage.setItem("corp-xlator-openai-key", openaiKeyInput.value);
  });

  geminiKeyInput.addEventListener("change", function () {
    localStorage.setItem("corp-xlator-gemini-key", geminiKeyInput.value);
  });

  jargonSelect.addEventListener("change", function () {
    localStorage.setItem("corp-xlator-jargon-level", jargonSelect.value);
  });

  insertModeSelect.addEventListener("change", function () {
    localStorage.setItem("corp-xlator-insert-mode", insertModeSelect.value);
  });

  translateBtn.addEventListener("click", handleTranslate);
  insertBtn.addEventListener("click", handleInsert);
  copyBtn.addEventListener("click", handleCopy);
}

function toggleKeyFields(provider) {
  document.getElementById("openai-key-field").classList.toggle("hidden", provider !== "openai");
  document.getElementById("gemini-key-field").classList.toggle("hidden", provider !== "gemini");
}

function getSelectedText() {
  return new Promise(function (resolve, reject) {
    var item = Office.context.mailbox.item;
    if (!item) {
      reject(new Error("No email item found."));
      return;
    }

    if (item.body) {
      item.body.getAsync(Office.CoercionType.Text, function (result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve(result.value);
        } else {
          reject(new Error("Could not read email body: " + result.error.message));
        }
      });
    } else {
      reject(new Error("No email body available."));
    }
  });
}

function getSystemPrompt(level) {
  var prompts = {
    mild: "Rewrite the following text in light corporate jargon. Keep it mostly readable but swap in a few business buzzwords. Be concise.",
    medium: "Rewrite the following text in heavy corporate doublespeak. Use phrases like 'leverage synergies', 'cross-functional alignment', 'circle back', 'touch base', 'move the needle', and 'bandwidth'. Be concise.",
    extreme: "Rewrite the following text in extreme corporate jargon. Every sentence must contain at least two meaningless buzzwords. Use phrases like 'paradigm shift', 'holistic ecosystem', 'value-added deliverables', 'actioning the roadmap', and 'driving stakeholder engagement'. Be concise but dense with jargon.",
    unhinged: "Rewrite the following text in absolutely unhinged corporate doublespeak. It should be barely comprehensible. Pack every sentence with maximum buzzwords. Use made-up compound nouns, gratuitous acronyms, and circular metaphors. The reader should have no idea what is actually being said. Be concise but devastating."
  };
  return prompts[level] || prompts.medium;
}

function callOpenAI(apiKey, systemPrompt, emailText) {
  return fetch("https://api.openai.com/v1/chat/completions", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      "Authorization": "Bearer " + apiKey
    },
    body: JSON.stringify({
      model: "gpt-4o-mini",
      messages: [
        { role: "system", content: systemPrompt },
        { role: "user", content: emailText }
      ],
      temperature: 0.8,
      max_tokens: 1000
    })
  }).then(function (response) {
    if (!response.ok) {
      return response.json().catch(function () { return {}; }).then(function (err) {
        throw new Error(err.error?.message || "OpenAI request failed (status " + response.status + ")");
      });
    }
    return response.json();
  }).then(function (data) {
    return data.choices[0].message.content.trim();
  });
}

function callGemini(apiKey, systemPrompt, emailText) {
  var url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=" + apiKey;

  return fetch(url, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      systemInstruction: { parts: [{ text: systemPrompt }] },
      contents: [{ parts: [{ text: emailText }] }],
      generationConfig: { temperature: 0.8, maxOutputTokens: 1000 }
    })
  }).then(function (response) {
    if (!response.ok) {
      return response.json().catch(function () { return {}; }).then(function (err) {
        throw new Error(err.error?.message || "Gemini request failed (status " + response.status + ")");
      });
    }
    return response.json();
  }).then(function (data) {
    return data.candidates[0].content.parts[0].text.trim();
  });
}

async function handleTranslate() {
  var provider = document.getElementById("provider").value;
  var jargonLevel = document.getElementById("jargon-level").value;
  var translateBtn = document.getElementById("translate-btn");

  var apiKey;
  if (provider === "openai") {
    apiKey = document.getElementById("openai-key").value.trim();
    if (!apiKey) { showStatus("Please enter your OpenAI API key.", "error"); return; }
  } else {
    apiKey = document.getElementById("gemini-key").value.trim();
    if (!apiKey) { showStatus("Please enter your Gemini API key.", "error"); return; }
  }

  translateBtn.disabled = true;
  translateBtn.classList.add("loading");
  translateBtn.textContent = "Translating...";

  try {
    var emailText = await getSelectedText();

    if (!emailText || emailText.trim().length === 0) {
      showStatus("Email body is empty. Nothing to translate.", "error");
      return;
    }

    showStatus("Sending to " + (provider === "openai" ? "OpenAI" : "Gemini") + "...", "info");

    var systemPrompt = getSystemPrompt(jargonLevel);
    var translated = provider === "openai"
      ? await callOpenAI(apiKey, systemPrompt, emailText)
      : await callGemini(apiKey, systemPrompt, emailText);

    document.getElementById("original-text").textContent = emailText.substring(0, 500);
    document.getElementById("translated-text").textContent = translated;
    document.getElementById("preview-panel").classList.remove("hidden");
    document.getElementById("insert-btn").classList.remove("hidden");
    document.getElementById("copy-btn").classList.remove("hidden");

    window._translatedText = translated;
    window._originalText = emailText;
    showStatus("Translation complete!", "success");
  } catch (err) {
    showStatus(err.message, "error");
  } finally {
    translateBtn.disabled = false;
    translateBtn.classList.remove("loading");
    translateBtn.textContent = "Translate to Corporate";
  }
}

function buildOutputText(translated, mode) {
  if (mode === "corporate") {
    return translated;
  }
  var original = window._originalText || "";
  return original + "\n\n--- Corporate Translation ---\n\n" + translated;
}

function handleInsert() {
  var translated = window._translatedText;
  if (!translated) {
    showStatus("Nothing to insert. Translate first.", "error");
    return;
  }

  var mode = document.getElementById("insert-mode").value;
  var outputText = buildOutputText(translated, mode);

  var item = Office.context.mailbox.item;
  if (item.body) {
    item.body.setAsync(
      outputText,
      { coercionType: Office.CoercionType.Text },
      function (result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          showStatus(mode === "corporate" ? "Corporate translation inserted!" : "Original + translation inserted!", "success");
        } else {
          showStatus("Failed to insert: " + result.error.message, "error");
        }
      }
    );
  }
}

function handleCopy() {
  var translated = window._translatedText;
  if (!translated) return;

  var mode = document.getElementById("insert-mode").value;
  var outputText = buildOutputText(translated, mode);

  navigator.clipboard.writeText(outputText).then(function () {
    showStatus("Copied to clipboard!", "success");
  }).catch(function () {
    showStatus("Failed to copy.", "error");
  });
}

function showStatus(message, type) {
  var el = document.getElementById("status");
  el.textContent = message;
  el.className = type;
  el.classList.remove("hidden");
}
