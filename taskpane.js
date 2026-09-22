"use strict";

// Configuration: set the exact, case-sensitive name of your local Obsidian vault.
const CONFIG = Object.freeze({
  VAULT_NAME: "CHANGE_ME",
  TARGET_FOLDER: "Inbox/Email",
  // Conservative cross-client safety guard; the URI is never truncated.
  MAX_OBSIDIAN_URI_LENGTH: 8000
});

const saveButton = document.getElementById("save-button");
const statusElement = document.getElementById("status");

Office.onReady((info) => {
  if (info.host !== Office.HostType.Outlook) {
    setStatus("This page must be opened from Outlook.", true);
    return;
  }

  saveButton.disabled = false;
  setStatus("Ready");
  saveButton.addEventListener("click", saveCurrentEmail);
});

async function saveCurrentEmail() {
  if (CONFIG.VAULT_NAME === "CHANGE_ME" || !CONFIG.VAULT_NAME.trim()) {
    setStatus("Set VAULT_NAME in taskpane.js before using the add-in.", true);
    return;
  }

  saveButton.disabled = true;
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
    const uri = createObsidianUri(fileName, markdown);

    if (uri.length > CONFIG.MAX_OBSIDIAN_URI_LENGTH) {
      setStatus("Email is too large for safe URI transfer. Nothing was truncated or sent.", true);
      return;
    }

    setStatus("Opening Obsidian...");
    window.location.assign(uri);
    setStatus("Email saved/opened in Obsidian.");
  } catch (error) {
    console.error("Unable to create Obsidian note:", error);
    setStatus("Unable to read this Outlook item.", true);
  } finally {
    saveButton.disabled = false;
  }
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

function createObsidianUri(fileName, content) {
  const folder = CONFIG.TARGET_FOLDER.replace(/\\/g, "/").replace(/^\/+|\/+$/g, "");
  if (!folder || folder.split("/").some((part) => !part || part === "." || part === "..")) {
    throw new Error("TARGET_FOLDER is invalid.");
  }

  const params = new URLSearchParams({
    vault: CONFIG.VAULT_NAME.trim(),
    file: `${folder}/${fileName}.md`,
    content,
    overwrite: "false"
  });
  return `obsidian://new?${params.toString()}`;
}

function textOrEmpty(value) {
  return typeof value === "string" ? value : "";
}

function setStatus(message, isError = false) {
  statusElement.textContent = message;
  statusElement.classList.toggle("error", isError);
}
