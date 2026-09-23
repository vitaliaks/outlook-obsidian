"use strict";

const CONFIG = Object.freeze({
  BUILD_VERSION: "1.1.0",
  DEFAULT_TARGET_FOLDER: "Inbox/Email",
  STORAGE_KEY: "outlook-to-obsidian-settings-v1",
  MAX_HANDOFF_URL_LENGTH: 8000
});

const saveButton = document.getElementById("save-button");
const statusElement = document.getElementById("status");
const settingsForm = document.getElementById("settings-form");
const vaultNameInput = document.getElementById("vault-name");
const targetFolderInput = document.getElementById("target-folder");
const resetSettingsButton = document.getElementById("reset-settings");
const copyAndContinueButton = document.getElementById("copy-and-continue");
const continueLink = document.getElementById("continue-link");
const debugOutput = document.getElementById("debug-output");
const copyDebugButton = document.getElementById("copy-debug");
const appVersionElement = document.getElementById("app-version");
const debugEvents = [];

appVersionElement.textContent = CONFIG.BUILD_VERSION;
debugLog("Script loaded", { build: CONFIG.BUILD_VERSION });

let settings = loadSettings();
let officeReady = false;
let pendingHandoff = null;
showSettings(settings);
debugLog("Settings loaded", { configured: Boolean(settings) });
updateSaveButton();

settingsForm.addEventListener("submit", saveSettings);
resetSettingsButton.addEventListener("click", resetSettings);
copyDebugButton.addEventListener("click", copyDiagnostics);
copyAndContinueButton.addEventListener("click", copyPendingNoteAndContinue);
saveButton.addEventListener("click", saveCurrentEmail);

window.addEventListener("error", (event) => {
  if (event.target && event.target !== window) {
    debugLog("Resource load failed", {
      element: textOrEmpty(event.target.tagName).toLowerCase() || "unknown",
      resource: safeResourceUrl(event.target.src || event.target.href)
    });
    return;
  }
  debugLog("Unhandled error", safeError(event.error || event.message));
}, true);

window.addEventListener("unhandledrejection", (event) => {
  debugLog("Unhandled promise rejection", safeError(event.reason));
});

const officeApiDetected = typeof Office !== "undefined" && typeof Office.onReady === "function";
debugLog("Office API detected", { detected: officeApiDetected });

if (officeApiDetected) {
  Office.initialize = (reason) => {
    debugLog("Office.initialize called", { reason: String(reason || "unknown") });
    markOfficeReady("Office.initialize");
  };

  Office.onReady((info) => {
    officeReady = info.host === Office.HostType.Outlook || Boolean(Office.context.mailbox);
    debugLog("Office ready", {
      source: "Office.onReady",
      host: textOrEmpty(info.host) || "unknown",
      platform: textOrEmpty(info.platform) || "unknown",
      outlookReady: officeReady,
      openBrowserWindow: Boolean(Office.context.ui && typeof Office.context.ui.openBrowserWindow === "function")
    });
    if (!officeReady) {
      debugLog("Unsupported host");
      setStatus("This page must be opened from Outlook.", true);
      return;
    }

    setStatus(settings ? "Ready" : "Enter and save your Obsidian settings.");
  }).catch((error) => {
    debugLog("Office.onReady rejected", safeError(error));
    setStatus("Office.js initialization failed. See Diagnostics.", true);
  });
} else {
  setStatus("Office.js did not load. See Diagnostics.", true);
}

window.setTimeout(() => {
  if (!officeReady) {
    debugLog("Office initialization timeout", { milliseconds: 15000 });
    setStatus("Office.js is not ready. See Diagnostics.", true);
  }
}, 15000);

async function saveCurrentEmail() {
  debugLog("Save button selected");
  if (!officeReady) {
    debugLog("Save stopped", { reason: "Office is not ready" });
    setStatus("Office.js is not ready. See Diagnostics.", true);
    return;
  }
  if (!settings) {
    debugLog("Save stopped", { reason: "settings missing" });
    setStatus("Enter and save your Obsidian settings first.", true);
    return;
  }

  saveButton.disabled = true;
  copyAndContinueButton.hidden = true;
  continueLink.hidden = true;
  pendingHandoff = null;
  setStatus("Reading email...");

  try {
    const item = Office.context.mailbox && Office.context.mailbox.item;
    debugLog("Outlook item inspected", {
      present: Boolean(item),
      type: item && item.itemType ? String(item.itemType) : "unknown",
      hasBody: Boolean(item && item.body)
    });
    if (!item || item.itemType !== Office.MailboxEnums.ItemType.Message || !item.body) {
      throw new Error("The current Outlook item is not a readable email message.");
    }

    debugLog("Body read requested", { coercion: "text" });
    const body = await getBodyAsText(item.body);
    debugLog("Body read succeeded", { characters: textOrEmpty(body).length });
    const email = {
      subject: textOrEmpty(item.subject),
      from: formatAddress(item.from || item.sender),
      to: formatAddresses(item.to),
      cc: formatAddresses(item.cc),
      date: formatDate(item.dateTimeCreated),
      body: textOrEmpty(body)
    };

    const markdown = createMarkdown(email);
    const fileName = createFileName(email.date, email.subject);
    const uri = createObsidianUri(fileName, settings);
    const bridgeUrl = createBridgeUrl(uri);
    debugLog("Note prepared", {
      markdownCharacters: markdown.length,
      obsidianUriCharacters: uri.length,
      handoffUrlCharacters: bridgeUrl.length,
      limit: CONFIG.MAX_HANDOFF_URL_LENGTH
    });

    if (bridgeUrl.length > CONFIG.MAX_HANDOFF_URL_LENGTH) {
      debugLog("Save stopped", { reason: "handoff URL exceeds limit" });
      setStatus("The vault or file path is too long to open safely.", true);
      return;
    }

    const copied = await writeClipboardText(markdown);
    if (!copied) {
      pendingHandoff = { markdown, bridgeUrl };
      copyAndContinueButton.hidden = false;
      debugLog("Clipboard retry offered");
      setStatus("Outlook blocked automatic clipboard access. Select Copy note and continue.", true);
      return;
    }

    continueWithHandoff(bridgeUrl);
  } catch (error) {
    console.error("Unable to create Obsidian note:", error);
    debugLog("Save failed", safeError(error));
    setStatus("Unable to read this Outlook item.", true);
  } finally {
    saveButton.disabled = false;
  }
}

function saveSettings(event) {
  event.preventDefault();

  try {
    const nextSettings = validateSettings({
      vaultName: vaultNameInput.value,
      targetFolder: targetFolderInput.value
    });
    localStorage.setItem(CONFIG.STORAGE_KEY, JSON.stringify(nextSettings));
    settings = nextSettings;
    clearPendingHandoff();
    showSettings(settings);
    updateSaveButton();
    setStatus("Settings saved locally.");
    debugLog("Settings saved", { configured: true });
  } catch (error) {
    console.error("Unable to save settings:", error);
    setStatus(error.message || "Unable to save settings locally.", true);
    debugLog("Settings save failed", safeError(error));
  }
}

function resetSettings() {
  try {
    localStorage.removeItem(CONFIG.STORAGE_KEY);
  } catch (error) {
    console.error("Unable to clear settings:", error);
  }

  settings = null;
  clearPendingHandoff();
  showSettings(null);
  updateSaveButton();
  setStatus("Settings reset.");
  debugLog("Settings reset");
}

function loadSettings() {
  try {
    const stored = localStorage.getItem(CONFIG.STORAGE_KEY);
    return stored ? validateSettings(JSON.parse(stored)) : null;
  } catch (error) {
    console.error("Unable to load settings:", error);
    return null;
  }
}

function validateSettings(value) {
  const vaultName = textOrEmpty(value && value.vaultName).trim();
  const targetFolder = normalizeFolder(textOrEmpty(value && value.targetFolder));

  if (!vaultName) throw new Error("Vault name is required.");
  if (/[\u0000-\u001f\u007f]/.test(vaultName)) throw new Error("Vault name contains invalid characters.");
  if (!targetFolder) throw new Error("Target folder is invalid.");

  return { vaultName, targetFolder };
}

function normalizeFolder(value) {
  const folder = value.trim().replace(/\\/g, "/").replace(/^\/+|\/+$/g, "");
  if (!folder || folder.split("/").some((part) =>
    !part || part === "." || part === ".." || /[<>:"|?*\u0000-\u001f\u007f]/.test(part)
  )) {
    return "";
  }
  return folder;
}

function showSettings(value) {
  vaultNameInput.value = value ? value.vaultName : "";
  targetFolderInput.value = value ? value.targetFolder : CONFIG.DEFAULT_TARGET_FOLDER;
}

function updateSaveButton() {
  saveButton.disabled = !settings;
}

function clearPendingHandoff() {
  pendingHandoff = null;
  copyAndContinueButton.hidden = true;
  continueLink.hidden = true;
}

function getBodyAsText(body) {
  return new Promise((resolve, reject) => {
    body.getAsync(Office.CoercionType.Text, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value);
      } else {
        reject(new Error(result.error ? result.error.message : "Body read failed."));
      }
    });
  });
}

function createMarkdown(email) {
  const lines = [
    "---",
    "type: email",
    `date: ${yamlString(email.date)}`,
    `from: ${yamlString(email.from)}`,
    ...yamlList("to", email.to),
    ...yamlList("cc", email.cc),
    `subject: ${yamlString(email.subject)}`,
    "---",
    "",
    `# ${markdownHeading(email.subject || "(No subject)")}`,
    "",
    "## Email",
    "",
    email.body,
    "",
    "## Notes",
    "",
    "",
    "## Actions",
    "",
    "- [ ]",
    ""
  ];

  return lines.join("\n");
}

function yamlList(key, values) {
  if (!values.length) return [`${key}: []`];
  return [`${key}:`, ...values.map((value) => `  - ${yamlString(value)}`)];
}

function yamlString(value) {
  return `"${textOrEmpty(value)
    .replace(/\\/g, "\\\\")
    .replace(/"/g, '\\"')
    .replace(/\r\n|\r|\n/g, "\\n")
    .replace(/\t/g, "\\t")
    .replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f\u007f]/g, "")}"`;
}

function markdownHeading(value) {
  return textOrEmpty(value).replace(/\r\n|\r|\n/g, " ").replace(/^#+\s*/, "").trim();
}

function formatAddress(address) {
  if (!address) return "";
  const name = textOrEmpty(address.displayName).trim();
  const email = textOrEmpty(address.emailAddress).trim();
  if (name && email) return `${name} <${email}>`;
  return name || email;
}

function formatAddresses(addresses) {
  return Array.isArray(addresses) ? addresses.map(formatAddress).filter(Boolean) : [];
}

function formatDate(value) {
  if (!value) return "";
  const date = value instanceof Date ? value : new Date(value);
  return Number.isNaN(date.getTime()) ? "" : date.toISOString();
}

function createFileName(isoDate, subject) {
  const date = /^\d{4}-\d{2}-\d{2}/.exec(isoDate || "");
  const safeSubject = (subject || "No subject")
    .normalize("NFKC")
    .replace(/[<>:"/\\|?*\u0000-\u001f\u007f]/g, "-")
    .replace(/\.\.+/g, ".")
    .replace(/^\.+|[. ]+$/g, "")
    .replace(/\s+/g, " ")
    .trim()
    .slice(0, 120) || "No subject";
  return `${date ? date[0] : "Unknown date"} - ${safeSubject}`;
}

function createObsidianUri(fileName, currentSettings) {
  const params = new URLSearchParams({
    vault: currentSettings.vaultName,
    file: `${currentSettings.targetFolder}/${fileName}.md`,
    clipboard: "true",
    overwrite: "false"
  });
  return `obsidian://new?${params.toString()}`;
}

function createBridgeUrl(obsidianUri) {
  const bridge = new URL("open-obsidian.html?v=1.1.0", window.location.href);
  bridge.hash = encodeURIComponent(obsidianUri);
  return bridge.toString();
}

async function copyPendingNoteAndContinue() {
  if (!pendingHandoff) return;

  copyAndContinueButton.disabled = true;
  const copied = await writeClipboardText(pendingHandoff.markdown);
  copyAndContinueButton.disabled = false;
  if (!copied) {
    setStatus("Clipboard access is blocked by Outlook or your browser settings.", true);
    return;
  }

  const bridgeUrl = pendingHandoff.bridgeUrl;
  pendingHandoff = null;
  copyAndContinueButton.hidden = true;
  continueWithHandoff(bridgeUrl);
}

function continueWithHandoff(bridgeUrl) {
  continueLink.href = bridgeUrl;
  continueLink.hidden = false;
  debugLog("Markdown copied and fallback link displayed");
  openBridge(bridgeUrl);
}

async function writeClipboardText(text) {
  if (navigator.clipboard && typeof navigator.clipboard.writeText === "function") {
    try {
      await navigator.clipboard.writeText(text);
      debugLog("Clipboard write succeeded", { method: "async-api", characters: text.length });
      return true;
    } catch (error) {
      debugLog("Async clipboard write failed", safeError(error));
    }
  }

  const textArea = document.createElement("textarea");
  textArea.value = text;
  textArea.setAttribute("readonly", "");
  textArea.setAttribute("aria-hidden", "true");
  textArea.style.position = "fixed";
  textArea.style.left = "-9999px";
  textArea.style.top = "0";
  document.body.appendChild(textArea);
  textArea.focus();
  textArea.select();

  try {
    const copied = document.execCommand("copy");
    debugLog(copied ? "Clipboard write succeeded" : "Legacy clipboard write rejected", {
      method: "exec-command",
      characters: text.length
    });
    return copied;
  } catch (error) {
    debugLog("Legacy clipboard write failed", safeError(error));
    return false;
  } finally {
    textArea.remove();
  }
}

function openBridge(bridgeUrl) {
  debugLog("Browser handoff requested", {
    officeApiAvailable: Boolean(Office.context.ui && typeof Office.context.ui.openBrowserWindow === "function")
  });
  if (Office.context.ui && typeof Office.context.ui.openBrowserWindow === "function") {
    try {
      Office.context.ui.openBrowserWindow(bridgeUrl);
      debugLog("Office openBrowserWindow returned without throwing");
      setStatus("Continue in the browser window to open Obsidian.");
      return;
    } catch (error) {
      console.error("Unable to open the browser automatically:", error);
      debugLog("Office openBrowserWindow threw", safeError(error));
    }
  }

  setStatus("Select Continue in browser to open Obsidian.");
  debugLog("Manual browser handoff required");
}

async function copyDiagnostics() {
  try {
    await navigator.clipboard.writeText(debugOutput.textContent);
    setStatus("Diagnostics copied.");
  } catch (error) {
    debugLog("Copy diagnostics failed", safeError(error));
    setStatus("Unable to copy diagnostics. Take a screenshot instead.", true);
  }
}

function debugLog(event, details) {
  const time = new Date().toISOString().slice(11, 23);
  const suffix = details === undefined ? "" : ` ${JSON.stringify(details)}`;
  debugEvents.push(`${time} ${event}${suffix}`);
  if (debugEvents.length > 40) debugEvents.shift();
  if (debugOutput) debugOutput.textContent = debugEvents.join("\n");
}

function safeError(error) {
  const value = error && typeof error === "object" ? error : { message: String(error || "Unknown error") };
  return {
    name: textOrEmpty(value.name).slice(0, 80) || "Error",
    code: value.code === undefined ? "" : String(value.code).slice(0, 80),
    message: textOrEmpty(value.message)
      .replace(/[\w.+-]+@[\w.-]+/g, "[email redacted]")
      .slice(0, 240) || "No error message"
  };
}

function safeResourceUrl(value) {
  try {
    const url = new URL(String(value || ""), window.location.href);
    return `${url.origin}${url.pathname}`;
  } catch (error) {
    return "unknown";
  }
}

function markOfficeReady(source) {
  officeReady = Boolean(Office.context && Office.context.mailbox);
  debugLog("Office initialization fallback", {
    source,
    mailboxAvailable: officeReady,
    openBrowserWindow: Boolean(Office.context && Office.context.ui && typeof Office.context.ui.openBrowserWindow === "function")
  });
  setStatus(officeReady ? (settings ? "Ready" : "Enter and save your Obsidian settings.") : "Outlook mailbox API is unavailable.", !officeReady);
}

function textOrEmpty(value) {
  return typeof value === "string" ? value : "";
}

function setStatus(message, isError = false) {
  statusElement.textContent = message;
  statusElement.classList.toggle("error", isError);
}
