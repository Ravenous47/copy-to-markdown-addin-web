# How to Sideload the Add-In

## Method 1: Excel Online (Browser)

### If "Upload My Add-in" is Available:
1. Go to https://office.live.com/start/Excel.aspx
2. Create or open a workbook
3. Click **Home** tab → **Add-ins** → **Get Add-ins**
4. Click **Upload My Add-in** (bottom left corner)
5. Click **Browse** and select `manifest.xml`
6. Click **Upload**

### If "Upload My Add-in" is NOT Available:
Your organization may have disabled this feature. Try Desktop Excel instead.

---

## Method 2: Desktop Excel (Windows/Mac)

### Quick Method - Manual Sideload:
1. Download `manifest.xml` from:
   https://ravenous47.github.io/copy-to-markdown-addin-web/manifest.xml

2. Save it to a local folder (e.g., `C:\OfficeAddIns\` or `~/OfficeAddIns/`)

3. Open Excel Desktop

4. Go to **File** → **Options** → **Trust Center** → **Trust Center Settings**

5. Click **Trusted Add-in Catalogs**

6. In "Catalog Url" field, paste your folder path:
   - Windows: `C:\OfficeAddIns\`
   - Mac: `file:///Users/YourName/OfficeAddIns/`

7. Check ☑ "Show in Menu"

8. Click **Add Catalog**, then **OK**

9. Restart Excel

10. Go to **Insert** tab → **My Add-ins** → **SHARED FOLDER**

11. Select "Copy to Markdown"

---

## Method 3: Using Network Share (Windows)

1. Create a network share folder:
   ```
   \\YourComputer\OfficeAddIns\
   ```

2. Place `manifest.xml` in this folder

3. In Excel: **File** → **Options** → **Trust Center** → **Trust Center Settings** → **Trusted Add-in Catalogs**

4. Add the network path: `\\YourComputer\OfficeAddIns\`

5. Restart Excel

6. **Insert** → **My Add-ins** → **SHARED FOLDER**

---

## Method 4: AppSource (For Testing - Development Mode)

If you have Microsoft 365 developer account:

1. Upload to AppSource Partner Center
2. Submit for validation
3. Install from AppSource

---

## Troubleshooting

### "Upload My Add-in" is grayed out or missing:
- Your organization's admin has disabled sideloading
- You need to use Desktop Excel with Trusted Catalogs method
- Contact your IT admin to enable Office Add-in sideloading

### Add-in doesn't appear after upload:
- Wait 30 seconds and refresh
- Clear Excel cache (close all Excel windows)
- Check browser console for errors (F12)
- Verify manifest.xml URLs are correct

### Add-in loads but doesn't work:
- Check that GitHub Pages is serving files correctly
- Open browser console (F12) to see JavaScript errors
- Verify HTTPS URLs in manifest.xml
- Make sure repository is Public

### Excel cache location (to clear):
- Windows: `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`
- Mac: `~/Library/Containers/com.microsoft.Excel/Data/Library/Caches/`

---

## Quick Test URL

To test if your add-in files are accessible:

1. Open browser and visit:
   - https://ravenous47.github.io/copy-to-markdown-addin-web/Home.html
   - https://ravenous47.github.io/copy-to-markdown-addin-web/Functions/FunctionFile.html

2. You should see the add-in interface (not a 404 error)

3. Check browser console (F12) for any errors
