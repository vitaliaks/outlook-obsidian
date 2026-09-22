"use strict";

const CONFIG = Object.freeze({
  DEFAULT_TARGET_FOLDER: "Inbox/Email",
  STORAGE_KEY: "outlook-to-obsidian-settings-v1",
  MAX_OBSIDIAN_URI_LENGTH: 8000
});

const saveButton = document.getElementById("save-button");
const statusElement = document.getElementById("status");
const settingsForm = document.getElementById("settings-form");
const vaultNameInput = document.getElementById("vault-name");
const targetFolderInput = document.getElementById("target-folder");
const resetSettingsButton = document.getElementById("reset-settings");
const continueLink = document.getElementById("continue-link");

let settings = loadSettings();
showSettings(settings);

settingsForm.addEventListener("submit", saveSettings);
resetSettingsButton.addEventListener("click", resetSettings);

Office.onReady((info) => {
  if (info.host !== Office.HostType.Outlook) {
    setStatus("This page must be opened from Outlook.", true);
    return;
  }

  updateSaveButton();
  setStatus(settings ? "Ready" : "Enter and save your Obsidian settings.");
  saveButton.addEventListener("click", saveCurrentEmail);
});

async function saveCurrentEmail() {
  if (!settings) {
    setStatus("Enter and save your Obsidian settings first.", true);
    return;
  }

  saveButton.disabled = true;
  continueLink.hidden = true;
  setStatus("Reading email...");

  try {
    const item = Office.context.mailbox && Office.context.mailbox.item;
    if (!item || item.itemType !== Office.MailboxEnums.ItemType.Message || !item.body) {
      throw new Error("The current Outlook item is not a readable email message.");
    }

    const body = await getBodyAsText(item.body);
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
    const uri = createObsidianUri(fileName, markdown, settings);
    const bridgeUrl = createBridgeUrl(uri);

    if (bridgeUrl.length > CONFIG.MAX_OBSIDIAN_URI_LENGTH) {
      setStatus("Email is too large for safe URI transfer. Nothing was truncated or sent.", true);
      return;
    }

    continueLink.href = bridgeUrl;
    continueLink.hidden = false;

    openBridge(bridgeUrl);
  } catch (error) {
    console.error("Unable to create Obsidian note:", error);
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
    showSettings(settings);
    updateSaveButton();
    setStatus("Settings saved locally.");
  } catch (error) {
    console.error("Unable to save settings:", error);
    setStatus(error.message || "Unable to save settings locally.", true);
  }
}

function resetSettings() {
  try {
    localStorage.removeItem(CONFIG.STORAGE_KEY);
  } catch (error) {
    console.error("Unable to clear settings:", error);
  }

  settings = null;
  showSettings(null);
  updateSaveButton();
  setStatus("Settings reset.");
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

function createObsidianUri(fileName, content, currentSettings) {
  const params = new URLSearchParams({
    vault: currentSettings.vaultName,
    file: `${currentSettings.targetFolder}/${fileName}.md`,
    content,
    overwrite: "false"
  });
  return `obsidian://new?${params.toString()}`;
}

function createBridgeUrl(obsidianUri) {
  const bridge = new URL("open-obsidian.html?v=1.0.1", window.location.href);
  bridge.hash = encodeURIComponent(obsidianUri);
  return bridge.toString();
}

function openBridge(bridgeUrl) {
  if (Office.context.ui && typeof Office.context.ui.openBrowserWindow === "function") {
    try {
      Office.context.ui.openBrowserWindow(bridgeUrl);
      setStatus("Continue in the browser window to open Obsidian.");
      return;
    } catch (error) {
      console.error("Unable to open the browser automatically:", error);
    }
  }

  setStatus("Select Continue in browser to open Obsidian.");
}

function textOrEmpty(value) {
  return typeof value === "string" ? value : "";
}

function setStatus(message, isError = false) {
  statusElement.textContent = message;
  statusElement.classList.toggle("error", isError);
}
