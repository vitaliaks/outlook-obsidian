"use strict";

const openObsidianLink = document.getElementById("open-obsidian");
const bridgeStatus = document.getElementById("bridge-status");
const bridgeDebug = document.getElementById("bridge-debug");

function bridgeLog(message) {
  bridgeDebug.textContent += `\n${message}`;
}

try {
  bridgeLog(`Fragment received: ${window.location.hash.length > 1 ? "yes" : "no"}`);
  bridgeLog(`Fragment characters: ${Math.max(0, window.location.hash.length - 1)}`);
  const uri = decodeURIComponent(window.location.hash.slice(1));
  if (!uri.startsWith("obsidian://new?")) throw new Error("Invalid Obsidian URI.");
  bridgeLog("Obsidian URI validated");
  window.history.replaceState(null, "", window.location.pathname);
  bridgeLog("Fragment removed from address bar");

  openObsidianLink.href = uri;
  openObsidianLink.hidden = false;
  bridgeStatus.textContent = "Select Open Obsidian and approve the browser prompt.";
  bridgeStatus.classList.remove("error");
  openObsidianLink.addEventListener("click", () => {
    bridgeStatus.textContent = "Opening Obsidian… If no prompt appears, verify that Obsidian is installed.";
    bridgeLog("Open Obsidian selected");
  });
} catch (error) {
  console.error("Unable to prepare Obsidian link:", error);
  bridgeStatus.textContent = "The Obsidian link is missing or invalid. Return to Outlook and try again.";
  bridgeLog(`Failed: ${error && error.message ? error.message : "unknown error"}`);
}
