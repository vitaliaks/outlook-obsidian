"use strict";

const openObsidianLink = document.getElementById("open-obsidian");
const bridgeStatus = document.getElementById("bridge-status");

try {
  const uri = decodeURIComponent(window.location.hash.slice(1));
  if (!uri.startsWith("obsidian://new?")) throw new Error("Invalid Obsidian URI.");
  window.history.replaceState(null, "", window.location.pathname);

  openObsidianLink.href = uri;
  openObsidianLink.hidden = false;
  bridgeStatus.textContent = "Select Open Obsidian and approve the browser prompt.";
  bridgeStatus.classList.remove("error");
} catch (error) {
  console.error("Unable to prepare Obsidian link:", error);
  bridgeStatus.textContent = "The Obsidian link is missing or invalid. Return to Outlook and try again.";
}
