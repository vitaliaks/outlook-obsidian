# Outlook to Obsidian

A minimal Outlook add-in that reads the currently opened email, generates Markdown entirely in the task pane, and opens a new note in a local Obsidian vault through Obsidian's supported URI scheme.

> Email content is processed locally in the Outlook add-in and is not transmitted to the GitHub Pages hosting server. GitHub Pages only serves the static HTML, CSS and JavaScript files.

## Architecture

The project has no build step and no backend. It uses static HTML and CSS, vanilla JavaScript, Microsoft Office.js, the system clipboard, and a direct `obsidian://new` link. It has no Node.js runtime, package dependencies, database, analytics, telemetry, cookies, or other application APIs.

The only external runtime resource is Microsoft Office.js, loaded from Microsoft's official CDN at `https://appsforoffice.microsoft.com/lib/1/hosted/office.js`, as required for Office add-ins. Icons, styles, and application code are served from this repository's GitHub Pages site.

## Privacy and security

- The add-in only reads `Office.context.mailbox.item`, the item currently open in Outlook.
- It does not enumerate the mailbox, access attachments, or modify an Outlook item.
- It makes no `fetch`, XMLHttpRequest, WebSocket, or similar application network call.
- It stores no email data in localStorage, sessionStorage, IndexedDB, cookies, or remote storage. Only the user-entered vault name and target folder are kept in localStorage.
- Markdown and YAML are generated locally. Dynamic content is never inserted into task-pane HTML.
- The generated Markdown is placed in the system clipboard and Obsidian reads it through the supported `clipboard=true` URI parameter. This avoids URL-length limits but temporarily replaces the user's clipboard contents.
- The task pane opens the local `obsidian://` protocol directly. Neither the note nor its destination passes through an intermediate web page.
- There are no secrets, credentials, analytics, or telemetry.

The manifest requests `ReadItem`, the minimum permission needed to read the current message body and metadata. It does not request `ReadWriteMailbox` or `ReadWriteItem`.

See [privacy.html](privacy.html) for the user-facing privacy statement.

## Files

- `manifest.xml` — sideloadable Outlook add-in manifest.
- `index.html` — GitHub Pages landing page.
- `taskpane.html`, `taskpane.js`, `styles.css` — task pane UI and local conversion logic.
- `privacy.html` — privacy statement.
- `icons/` — local PNG manifest icons.
- `.nojekyll` — prevents Jekyll processing on GitHub Pages.

## Configure Obsidian

Open the add-in task pane and enter:

- **Vault name** — the exact, case-sensitive name of the local Obsidian vault.
- **Target folder** — the vault-relative destination folder; the default is `Inbox/Email`.

Choose **Save settings**. These two values are stored only in localStorage for the GitHub Pages task-pane origin. Email content is never stored. Use **Reset** to remove the saved configuration. Do not use `.` or `..` path segments in the target folder.

## Configure and enable GitHub Pages

The configured repository is `https://github.com/vitaliaks/outlook-obsidian.git`, so the expected Pages URL is `https://vitaliaks.github.io/outlook-obsidian/`.

1. Confirm the repository remote:

   ```sh
   git remote -v
   ```

2. Commit and push `main` only when you are ready.
3. In GitHub, open **Settings → Pages**.
4. Under **Build and deployment**, choose **Deploy from a branch**.
5. Select branch **main**, folder **/(root)**, then save.
6. Confirm the final URL shown in Pages settings. If a custom domain or enterprise configuration changes it, update every hosted URL in `manifest.xml` and the source link in `index.html`.
7. Wait for deployment, then open the task-pane URL directly and confirm it returns over HTTPS:

   `https://vitaliaks.github.io/outlook-obsidian/taskpane.html`

## Sideload in Outlook

After the site is deployed and `manifest.xml` contains the final HTTPS URLs:

1. In Outlook, open **My add-ins** (the exact menu location varies by Outlook version).
2. Choose **Custom Addins → Add a custom add-in → Add from file**.
3. Select `manifest.xml` and accept the prompt.
4. Open a received email in its read form.
5. Choose **Save to Obsidian** from the message add-in commands.

This XML add-in manifest targets Outlook on the web and Outlook on Mac where the client and administrator allow sideloaded Office web add-ins. Organization policies can disable sideloading or custom add-ins.

## Test

1. Confirm Obsidian is installed, the configured vault exists, and the `obsidian://` protocol opens Obsidian.
2. Open an email containing quotes, colons, non-ASCII characters, multiple To/CC recipients, and a multiline body.
3. Open the add-in and select **Save to Obsidian**.
4. If Outlook blocks the first clipboard attempt, select **Copy note and continue**.
5. If Obsidian does not open automatically, select **Open Obsidian** in the task pane and approve the client prompt.
6. Verify the note appears under `Inbox/Email`, its YAML parses, and Outlook's item remains unchanged.
7. Repeat with missing CC, a missing sender/date where possible, and a long email well beyond 8,000 characters.
8. In browser developer tools, verify there are no application network requests containing email data. Loading Office.js from Microsoft's CDN is expected.

## Known limitations

- Version 1 handles only opened email messages in read mode. It does not handle meetings, tasks, attachments, conversation-specific processing, templates, tagging, or AI.
- Outlook's plain-text body conversion can lose formatting from HTML messages. The add-in deliberately requests `Office.CoercionType.Text`.
- File names are sanitized and the subject portion is limited to 120 characters for filesystem compatibility; the email subject inside the note and the email body are not truncated.
- `overwrite=false` is requested. How a name collision is presented or resolved depends on the installed Obsidian version.
- Outlook client and tenant policies determine whether custom add-ins and direct custom-protocol navigation are allowed.
- Saving temporarily replaces the system clipboard with the generated Markdown. Clipboard permissions and corporate browser policies can block this operation.

### Large emails and clipboard transfer

The note content is transferred through the system clipboard rather than the URL, so long emails are not subject to browser or protocol-handler URL limits. The compact Obsidian URI is still capped at 8,000 characters as a defensive check; in practice only an unusually long vault, folder, or filename can reach that limit. If automatic clipboard access is rejected, the add-in displays **Copy note and continue** so the user can retry from an explicit click.

## Troubleshooting

- **Enter and save your Obsidian settings** — provide the vault name and target folder in the task pane, then choose **Save settings**.
- **Add-in does not appear** — confirm the message is opened in read mode, the manifest has real HTTPS URLs, and custom add-ins are permitted by the Microsoft 365 administrator.
- **Task pane is blank** — open the configured `taskpane.html` URL directly, verify GitHub Pages deployment, HTTPS, and the browser console. Office.js only initializes fully inside Office.
- **Unable to read this Outlook item** — confirm the current item is a received email rather than an appointment, compose form, or unsupported item type.
- **Obsidian does not open** — select **Open Obsidian** in the task pane; then confirm Obsidian is installed, the OS registered the `obsidian://` handler, the configured vault name is exact, and Outlook allows custom-protocol navigation.
- **Clipboard access is blocked** — select **Copy note and continue**. If it still fails, allow clipboard access for the add-in or review the organization's browser and Outlook policies.
- **Changes do not appear** — allow time for GitHub Pages and the Outlook webview cache to refresh; then close and reopen the task pane.
- **Nothing happens after Save to Obsidian** — select **Save to Obsidian** again, then use the displayed **Open Obsidian** link. Confirm that clipboard access and the `obsidian://` protocol are allowed.

## Security review checklist

Before release, verify that the manifest still requests only `ReadItem`, no email data is written to browser storage, all hosted URLs are HTTPS and point to this repository, and only Office.js is externally loaded.
