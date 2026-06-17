Office.onReady(function (info) {
  if (info.host === Office.HostType.Outlook) {
    init();
  }
});

function init() {
  var apiKeyInput = document.getElementById("api-key");
  var jargonSelect = document.getElementById("jargon-level");
  var insertModeSelect = document.getElementById("insert-mode");
  var translateBtn = document.getElementById("translate-btn");
  var insertBtn = document.getElementById("insert-btn");
  var copyBtn = document.getElementById("copy-btn");

  var savedKey = localStorage.getItem("corp-translator-api-key");
  if (savedKey) {
    apiKeyInput.value = savedKey;
  }

  var savedLevel = localStorage.getItem("corp-translator-jargon-level");
  if (savedLevel) {
    jargonSelect.value = savedLevel;
  }

  var savedMode = localStorage.getItem("corp-translator-insert-mode");
  if (savedMode) {
    insertModeSelect.value = savedMode;
  }

  apiKeyInput.addEventListener("change", function () {
    localStorage.setItem("corp-translator-api-key", apiKeyInput.value);
  });

  jargonSelect.addEventListener("change", function () {
    localStorage.setItem("corp-translator-jargon-level", jargonSelect.value);
  });

  insertModeSelect.addEventListener("change", function () {
    localStorage.setItem("corp-translator-insert-mode", insertModeSelect.value);
  });

  translateBtn.addEventListener("click", handleTranslate);
  insertBtn.addEventListener("click", handleInsert);
  copyBtn.addEventListener("click", handleCopy);
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

async function handleTranslate() {
  var apiKey = document.getElementById("api-key").value.trim();
  var jargonLevel = document.getElementById("jargon-level").value;
  var translateBtn = document.getElementById("translate-btn");

  if (!apiKey) {
    showStatus("Please enter your OpenAI API key.", "error");
    return;
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

    showStatus("Sending to AI...", "info");

    var response = await fetch("https://api.openai.com/v1/chat/completions", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "Authorization": "Bearer " + apiKey
      },
      body: JSON.stringify({
        model: "gpt-4o-mini",
        messages: [
          { role: "system", content: getSystemPrompt(jargonLevel) },
          { role: "user", content: emailText }
        ],
        temperature: 0.8,
        max_tokens: 1000
      })
    });

    if (!response.ok) {
      var err = await response.json().catch(function () { return {}; });
      throw new Error(err.error?.message || "API request failed with status " + response.status);
    }

    var data = await response.json();
    var translated = data.choices[0].message.content.trim();

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
