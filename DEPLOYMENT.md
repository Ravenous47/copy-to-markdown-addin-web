# Deployment Instructions

## Step 1: Create GitHub Repository

1. Go to https://github.com/new
2. Repository name: `copy-to-markdown-addin-web`
3. Keep it Public (required for free GitHub Pages)
4. Click "Create repository"

## Step 2: Push to GitHub

Run these commands from the project directory:

```bash
git remote add origin https://github.com/ravenous47/copy-to-markdown-addin-web.git
git branch -M main
git push -u origin main
```

## Step 3: Enable GitHub Pages

1. Go to your repository: https://github.com/ravenous47/copy-to-markdown-addin-web
2. Click **Settings** tab
3. Click **Pages** in the left sidebar
4. Under "Source", select:
   - Branch: `main`
   - Folder: `/ (root)`
5. Click **Save**
6. Wait 1-2 minutes for deployment

## Step 4: Verify Deployment

Your add-in will be live at:
**https://ravenous47.github.io/copy-to-markdown-addin-web/**

Test these URLs:
- https://ravenous47.github.io/copy-to-markdown-addin-web/Home.html
- https://ravenous47.github.io/copy-to-markdown-addin-web/Functions/FunctionFile.html
- https://ravenous47.github.io/copy-to-markdown-addin-web/manifest.xml

## Step 5: Install the Add-In

### For Excel Online:

1. Open Excel Online
2. Go to **Insert** → **Office Add-ins**
3. Click **Upload My Add-in**
4. Upload `manifest.xml` from this repository

### For Desktop Excel (Windows/Mac):

1. Download `manifest.xml` from:
   https://ravenous47.github.io/copy-to-markdown-addin-web/manifest.xml
2. Save it to a trusted location
3. In Excel: **File** → **Options** → **Trust Center** → **Trust Center Settings** → **Trusted Add-in Catalogs**
4. Add your folder path
5. Restart Excel
6. **Insert** → **My Add-ins** → **Shared Folder**

## Troubleshooting

### If GitHub Pages shows 404:
- Wait a few more minutes
- Check that branch is set to `main` in Settings → Pages
- Verify files are in the root directory (not in a subdirectory)

### If add-in doesn't load:
- Check browser console for CORS or HTTPS errors
- Verify all URLs in manifest.xml are correct
- Make sure repository is Public (not Private)
- Clear Excel cache and restart

### CORS Issues:
GitHub Pages automatically serves with correct CORS headers for Office Add-ins.
If you see CORS errors, verify the manifest URLs match your GitHub Pages URL exactly.
