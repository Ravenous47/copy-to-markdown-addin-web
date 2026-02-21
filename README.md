# Copy to Markdown Excel Add-In (Web Version)

Excel web add-in that copies selected cells to Markdown table format.

## Features

- Copy Excel ranges to Markdown tables
- Preserves formatting (alignment, text)
- Works in Excel Online and Desktop Excel (2016+)

## Installation

### Excel Online

1. Open Excel Online
2. Go to **Insert** → **Office Add-ins**
3. Click **Upload My Add-in**
4. Upload the `manifest.xml` file

### Desktop Excel (Windows/Mac)

1. Save `manifest.xml` to a network share or local folder
2. Add the folder to **File** → **Options** → **Trust Center** → **Trust Center Settings** → **Trusted Add-in Catalogs**
3. Restart Excel
4. Go to **Insert** → **My Add-ins** → **Shared Folder**
5. Select "Copy to Markdown"

## Usage

1. Select cells in Excel
2. Right-click
3. Choose **"Copy to Markdown"**
4. Paste anywhere!

## Hosting

This add-in is hosted via GitHub Pages at:
`https://ravenous47.github.io/copy-to-markdown-addin-web/`

## Credits

Based on [CopyToMarkdownAddIn](https://github.com/nuitsjp/CopyToMarkdownAddIn) by nuits.jp
