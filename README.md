# IRM Encryption Check — Outlook Add-in

An Outlook Add-in that detects sensitive information and prompts users to apply a **Microsoft IRM / Sensitivity Label** before sending. Works in **Outlook Desktop (Classic Win32)** and **Outlook on the Web (OWA)**.

---

## What it does

1. **Scan** — user clicks the ribbon button to scan the email body and subject for sensitive patterns (card numbers, TFNs, passwords, API keys, health records, CONFIDENTIAL markers, etc.)
2. **Review** — a task pane shows any detected patterns and asks whether the email contains sensitive information
3. **Label** — user picks a Sensitivity Label: **General**, **Confidential**, or **Highly Confidential**
4. **Apply** — the add-in writes `X-IRM-Label` and `X-Reviewed-By-Addin` headers; if your org uses Microsoft Purview / AIP labels, the user is prompted to also confirm the matching label in the ribbon
5. **Block on send** *(requires admin policy)* — if the email contains sensitive patterns and the user has NOT completed the review, Outlook blocks the send with an actionable notification

---

## Files

```
outlook-encrypt-addin/
├── manifest.xml      ← Add-in registration (update localhost URLs for production)
├── taskpane.html     ← Task pane UI: scan, review, label selection
├── commands.html     ← Background function file (event-based, not shown to users)
├── commands.js       ← onItemSend handler
├── assets/           ← Place icon-16.png, icon-32.png, icon-64.png, icon-80.png, icon-128.png here
└── README.md
```

---

## Setup

### 1. Add icons

Place PNG icons in the `assets/` folder:

| File | Size |
|---|---|
| icon-16.png | 16×16 |
| icon-32.png | 32×32 |
| icon-64.png | 64×64 |
| icon-80.png | 80×80 |
| icon-128.png | 128×128 |

A simple shield or lock icon works well. Free icons: [Microsoft Fluent UI Icons](https://github.com/microsoft/fluentui-system-icons).

### 2. Host over HTTPS

Office.js requires HTTPS. For local development:

```bash
# Install a trusted local cert (one-time)
npx office-addin-dev-certs install

# Serve files
npx http-server . -p 3000 --cors -S -C ~/.office-addin-dev-certs/localhost.crt -K ~/.office-addin-dev-certs/localhost.key
```

### 3. Sideload the manifest

**Outlook Desktop (Windows):**
- File → Options → Trust Center → Trust Center Settings → Trusted Add-in Catalogs
- Or: Home → Get Add-ins → My Add-ins → + Add a custom add-in → Add from file → select `manifest.xml`

**Outlook on the Web:**
- Settings (⚙) → View all Outlook settings → Mail → Customize actions → Manage add-ins
- Click **+** → Add from file → upload `manifest.xml`

**Microsoft 365 Admin Centre (organisation-wide):**
- Settings → Integrated Apps → Upload custom app → Office Add-in → upload `manifest.xml`

---

## Production deployment

1. Replace all `https://localhost:3000` in `manifest.xml` with your hosted URL (e.g. `https://addin.yourcompany.com`)
2. Host files on any HTTPS server (Azure Static Web Apps, IIS, nginx, etc.)
3. Re-upload the updated `manifest.xml` to the Admin Centre

---

## Enabling on-send blocking (Exchange Online)

The automatic send-block requires an Exchange admin to enable on-send add-ins:

```powershell
# Exchange Online PowerShell
Set-OwaMailboxPolicy -Identity OwaMailboxPolicy-Default -OnSendAddinsEnabled $true
```

Without this, the task pane scan and IRM labelling still work — only the automatic block is inactive.

---

## Sensitivity labels and Microsoft Purview

The add-in writes `X-IRM-Label` and `X-Sensitivity` internet headers to record the user's choice.

If your organisation has **Microsoft Purview Information Protection** (formerly AIP) deployed, the Purview Office add-in reads these headers and can auto-apply the corresponding label. Alternatively, the task pane prompts the user to confirm the label in the Purview ribbon themselves.

To integrate directly with the Purview labelling API (applying labels programmatically without the ribbon step), you can extend `commands.js` to call the [Microsoft Information Protection SDK](https://learn.microsoft.com/en-us/information-protection/develop/overview) via a backend proxy.

---

## Extending sensitive data patterns

Edit the `PATTERNS` array in both `taskpane.html` and `commands.js` to add custom regexes:

```js
{ label: 'Employee ID', re: /\bEMP\d{6}\b/ },
{ label: 'Project code', re: /\bPROJ-[A-Z]{3}-\d{4}\b/i },
```

---

## Security note

All scanning runs **locally** inside Outlook / the user's browser. No email content is transmitted to any external server by this add-in. The only data written is custom internet headers on the outgoing message itself.
