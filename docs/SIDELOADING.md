# 🔌 Sideloading Guide

How to install the LMStudio AI add-in into each Office application.

## Prerequisites

The webpack dev server must be running before sideloading:

```bash
cd office-addin
npm run dev-server
# Server starts at https://localhost:3000
```

---

## Method 1 — Launcher Document (Recommended)

A pre-built `.docx` with the add-in embedded is created by the setup script:

**`C:\Users\<You>\Desktop\LMStudio-AI.docx`**

1. Double-click `LMStudio-AI.docx`
2. The task pane opens automatically on the right
3. If prompted, click **Enable Content** / **Trust**

---

## Method 2 — Registry Developer Sideload

The launcher writes to:

```
HKCU\Software\Microsoft\Office\16.0\WEF\Developer\{GUID}
  ManifestPath = <path to manifest.xml>
  ManifestType = 1
```

This is the official dev-sideload mechanism used by `office-addin-debugging`.
After writing the key, restart Word — the add-in is active.

---

## Word / Excel / PowerPoint

### Via Developer Tab
1. **File → Options → Customize Ribbon** → check **Developer** → OK
2. **Developer** tab → **Office Add-ins**
3. **MY ADD-INS** tab → **Upload My Add-in**
4. Browse to `office-addin/manifest.xml` → **Upload**

### Via Insert Tab
1. **Insert** → **Get Add-ins**
2. Click **MY ORGANIZATION** tab (requires catalog registration)
3. Select **LMStudio AI Assistant** → **Add**

---

## Outlook

Sideload uses the separate `manifest-outlook.xml`:

1. In Outlook: **File → Manage Add-ins** (opens browser)
2. Click **+** → **Add from file**
3. Upload `office-addin/manifest-outlook.xml`

---

## Troubleshooting

| Problem | Solution |
|---|---|
| Task pane shows blank/error | Ensure `npm run dev-server` is running |
| "Cannot connect to LMStudio" | Start LMStudio, enable Local Server on port 1234 |
| Add-in not in ribbon | Restart Word after sideloading |
| Certificate error in task pane | Run `mkcert -install` to trust the local CA |
