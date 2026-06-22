# 🗺️ Mind Map Studio — OneNote Add-in

A full-featured interactive mind map builder that runs directly inside Microsoft OneNote as a Task Pane add-in.

---

## ✨ Features

| Feature | Details |
|---|---|
| **Topic creation** | Add root topics and unlimited child branches |
| **Shapes** | Rounded rectangle or ellipse per node |
| **Full styling** | Background color, text color, border color, font family, size, bold, italic |
| **5 color themes** | Classic · Pastel Dream · Pink Dream · Deep Ocean · Enchanted Forest |
| **Collapse / Expand** | Click the **−/+** circle at the tip of any branch arrow to hide/show its children |
| **Note & Association** | Attach a private text note to any node; hover the 🟠 dot to read it |
| **OneNote page link** | Link any topic to another page in your notebook; click ⛓ to navigate |
| **Pan & Zoom** | Drag canvas to pan · Scroll wheel to zoom · Fit-to-screen button |
| **Save / Load** | Export map as `.json` file; reload it at any time |
| **Insert into page** | Embeds the map as an image directly into the current OneNote page |

---

## 🚀 Setup & Installation

### 1. Prerequisites

- **Node.js** v18+ — [nodejs.org](https://nodejs.org)
- **Microsoft OneNote** desktop app (Windows) or OneNote on the web
- **HTTPS certificate** for localhost (required by Office)

### 2. Install dependencies

```bash
cd "c:\Visual Studio Code Projects\OneNoteExtension"
npm install
```

### 3. Install dev HTTPS certificates (one-time)

```bash
npm run install-certs
```

This installs a trusted self-signed certificate for `localhost:3000` so Office will accept the add-in.

### 4. Start the server

```bash
npm start
```

The server runs at **https://localhost:3000**.

### 5. Sideload the add-in into OneNote

**OneNote Desktop (Windows):**
1. Open OneNote → **File → Options → Add-ins**
2. In the *Manage* dropdown choose **Office Add-ins** → Go
3. Click **Upload My Add-in** → Browse to `manifest.xml` → **Upload**
4. The **Mind Map** button now appears in the **Home** ribbon tab

**OneNote on the Web:**
1. Open any page in OneNote online
2. **Insert → Office Add-ins → Upload My Add-in**
3. Browse to `manifest.xml` → **Upload**

---

## 🎨 Keyboard Shortcuts

| Shortcut | Action |
|---|---|
| `Ctrl + Enter` | Add a new root topic |
| `Tab` | Add a child branch to the selected node |
| `F2` | Edit the selected node's label |
| `Delete` | Delete the selected node (with confirmation) |
| `Escape` | Deselect / close panel |
| `Ctrl + +` / `Ctrl + -` | Zoom in / zoom out |
| `Ctrl + 0` | Fit map to screen |

---

## 🗂️ Project Structure

```
OneNoteExtension/
├── manifest.xml               # Office Add-in manifest
├── package.json
├── server.js                  # HTTPS dev server (Express)
├── assets/
│   ├── icon-16.svg / icon-32.svg / icon-80.svg
├── src/
│   ├── mindmap/
│   │   ├── themes.js          # 5 color themes
│   │   └── mindmap.js         # SVG mind map engine (MindMap class)
│   └── taskpane/
│       ├── taskpane.html      # Main UI
│       ├── taskpane.css       # Styles
│       └── taskpane.js        # UI controller + Office.js integration
```

---

## 🛠️ Development

The task pane (`taskpane.html`) can also be opened directly in a browser for development:

```
https://localhost:3000/taskpane/taskpane.html
```

The app detects whether Office.js is available and falls back gracefully when running in a plain browser, using mock OneNote page data.

---

## 🌸 Themes

| Theme | Description |
|---|---|
| **Classic** | Dark navy root, deep blues and greens |
| **Pastel Dream** | Soft, gentle pastel colors — easy on the eyes |
| **Pink Dream** | Full spectrum of pinks from deep fuchsia to blush |
| **Deep Ocean** | Dark navy to light aqua gradient blues |
| **Enchanted Forest** | Deep forest greens and earthy tones |

---

## 📋 Requirements from the user (Hebrew → implemented)

- ✅ Separate toolbar at the top of the add-in
- ✅ Choose shape (rectangle or ellipse) per node
- ✅ Choose font, size, and color
- ✅ Add branches from any node with the **+** button
- ✅ Collapse branches via the **−** circle at the end of each arrow
- ✅ Individual editing of each node's color, font, etc.
- ✅ 4–5 automatic color themes (including pastel and pink)
- ✅ Link any topic to another OneNote page (click ⛓ to navigate)
- ✅ Hover note/association per node (hover the 🟠 dot to read)
