# Rounded Corner Lock (Office.js PowerPoint Add-in)

An Office.js task pane add-in for **PowerPoint** that lets you type a corner radius in **pixels** (e.g., `16`) and keeps it **fixed after resize**. DPI is **auto-detected** from the current slide.

## Quick start

1. **Enable GitHub Pages** for this repo (Settings → Pages). Use the main branch, root folder.
2. Edit `manifest.xml` and replace `<your-username>` with your GitHub account name.
3. Open PowerPoint → **Insert** → **Add-ins** → **Upload my Add-in** (sideload) and choose your local `manifest.xml`.
4. In the task pane, enter px and click **Apply to selected shapes**. Toggle **Lock radius after resize**.

## Structure
```
rounded-corner-lock/
├─ manifest.xml
├─ assets/
│  └─ icon-32.png
└─ taskpane/
   ├─ taskpane.html
   ├─ taskpane.css
   └─ taskpane.js
```

## Notes
- Requires **PowerPointApi 1.10**.
- Works on Windows/Mac/Web. Selection events may be flaky on some Mac builds; a polling fallback is included.
