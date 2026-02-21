# Copy to Markdown - Excel Add-In

Bidirectional Excel ↔ Markdown table converter. Copy Excel ranges to Markdown format and paste Markdown tables back into Excel.

![Excel Online Compatible](https://img.shields.io/badge/Excel%20Online-Compatible-green)
![Excel Desktop](https://img.shields.io/badge/Excel%20Desktop-2016%2B-blue)
![Office.js](https://img.shields.io/badge/Office.js-1.1-orange)

## 🚀 Features

### Copy to Markdown
- Select any range in Excel
- Right-click → "Copy to Markdown"
- Get properly formatted Markdown table
- Works with dates, numbers, formulas (displays values)

### Paste from Markdown
- Right-click → "Paste from Markdown"
- Paste any Markdown table
- Automatically inserts into Excel cells
- Handles alignment rows and complex tables

## 📦 Installation

### Excel Online (Temporary - Re-add each session)

1. Download [manifest.xml](https://ravenous47.github.io/copy-to-markdown-addin-web/manifest.xml)
2. In Excel Online: **Home** → **Add-ins** → **More Add-ins**
3. Click **Upload My Add-in** (bottom left)
4. Upload the manifest file

### Desktop Excel (Permanent)

#### Windows:
1. Download `manifest.xml` to: `C:\OfficeAddIns\CopyToMarkdown\`
2. Excel → **File** → **Options** → **Trust Center** → **Trust Center Settings**
3. **Trusted Add-in Catalogs** → Add catalog URL: `C:\OfficeAddIns\CopyToMarkdown\`
4. Check ☑ **Show in Menu**
5. Restart Excel
6. **Insert** → **My Add-ins** → **SHARED FOLDER** → Select add-in

#### Mac:
1. Download `manifest.xml` to: `~/OfficeAddIns/CopyToMarkdown/`
2. Similar steps as Windows
3. Use path: `file:///Users/YourUsername/OfficeAddIns/CopyToMarkdown/`

## 🎯 Usage

### Copy to Markdown

1. **Select cells** in Excel
2. **Right-click** → **"Copy to Markdown"**
3. Dialog opens with Markdown table
4. **Copy** from dialog (text is pre-selected)
5. **Paste** into GitHub, documentation, forums, etc.

**Example:**

Excel:
```
| Name  | Age |
| John  | 25  |
| Jane  | 30  |
```

Becomes:
```markdown
| Name | Age |
|------|-----|
| John | 25  |
| Jane | 30  |
```

### Paste from Markdown

1. **Copy** a Markdown table from anywhere
2. **Click** where you want the table in Excel
3. **Right-click** → **"Paste from Markdown"**
4. **Paste** Markdown in dialog
5. **Click "Insert into Excel"**
6. Data appears in cells!

## 🏗️ Architecture

### Tech Stack

- **Office.js API** - Excel integration
- **markdown-it** (118KB) - Robust Markdown parser
- **markdown-table** (2KB) - Table generator
- **GitHub Pages** - Static hosting
- **Pure JavaScript** - No build tools needed

### File Structure

```
copy-to-markdown-addin-web/
├── manifest.xml              # Add-in manifest
├── index.html               # Landing page
├── dialog.html              # Copy dialog
├── paste-dialog.html        # Paste dialog
├── Functions/
│   ├── FunctionFile.html    # Background page
│   └── FunctionFile.js      # Main logic
├── libs/
│   ├── markdown-it.min.js   # Markdown parser
│   └── markdown-table.min.js # Table generator
├── images/                   # Icons
└── Scripts/                  # Office.js libraries
```

### How It Works

#### Copy to Markdown:
1. `copyToMarkdown()` function triggered from context menu
2. Reads selected Excel range using `Excel.run()`
3. Converts cell data to 2D array
4. Uses `markdownTable()` library to generate Markdown
5. Opens dialog with formatted text
6. User copies from dialog (clipboard APIs blocked in Excel Online)

#### Paste from Markdown:
1. `pasteFromMarkdown()` opens dialog
2. User pastes Markdown table
3. Dialog uses `markdown-it` to parse table
4. Sends parsed data to parent via `messageParent()`
5. Parent receives data and inserts using `Excel.run()`
6. Data appears in Excel cells

### Why Dialogs?

Excel Online blocks direct clipboard access due to browser security policies. Dialogs provide a reliable cross-platform solution:
- ✅ Works in Excel Online and Desktop
- ✅ No permission prompts
- ✅ User has full control
- ✅ Visual confirmation

## 🔧 Development

### Local Testing

1. Clone repository:
   ```bash
   git clone https://github.com/Ravenous47/copy-to-markdown-addin-web.git
   cd copy-to-markdown-addin-web
   ```

2. Serve files locally:
   ```bash
   python3 -m http.server 8000
   ```

3. Update `manifest.xml` URLs to `http://localhost:8000/`

4. Sideload in Excel

### Deployment

Automatically deployed via GitHub Pages:
- **URL**: https://ravenous47.github.io/copy-to-markdown-addin-web/
- **Updates**: Push to `main` branch → Live in 1-2 minutes

### Manifest Configuration

Key settings in `manifest.xml`:

```xml
<!-- Context menu items -->
<Control xsi:type="Button" id="Contoso.TaskpaneButton">
  <Action xsi:type="ExecuteFunction">
    <FunctionName>copyToMarkdown</FunctionName>
  </Action>
</Control>

<Control xsi:type="Button" id="Contoso.PasteButton">
  <Action xsi:type="ExecuteFunction">
    <FunctionName>pasteFromMarkdown</FunctionName>
  </Action>
</Control>
```

## 📋 Limitations

### Current:
- **Data only** - No formatting, colors, or borders (yet)
- **Excel Online** - Add-in doesn't persist between sessions (Microsoft limitation)
- **Single table** - Processes one table at a time

### Planned:
- Alignment preservation (left/center/right)
- Cell styling (bold, italic)
- Multi-table support
- Settings panel

## 🐛 Troubleshooting

### Add-in doesn't appear in context menu
- Re-upload manifest in Excel Online
- Check manifest.xml URLs are correct
- Clear Excel cache and reload

### Copy/Paste doesn't work
- Check browser console (F12) for errors
- Verify GitHub Pages is serving files
- Try hard refresh (Ctrl+Shift+R)

### "Permissions policy" error
- Expected in Excel Online - dialogs are the solution
- Make sure you're using the dialog, not direct clipboard

## 📜 License

Based on [CopyToMarkdownAddIn](https://github.com/nuitsjp/CopyToMarkdownAddIn) by nuits.jp

Free for personal and commercial use.

## 🤝 Contributing

Issues and pull requests welcome!

## 📞 Support

- **GitHub Issues**: [Report a bug](https://github.com/Ravenous47/copy-to-markdown-addin-web/issues)
- **Original Project**: [nuitsjp/CopyToMarkdownAddIn](https://github.com/nuitsjp/CopyToMarkdownAddIn)

---

**Built with ❤️ using Office.js and modern web standards**
