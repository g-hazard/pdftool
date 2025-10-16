# Distribution Guide

## How to Make This Available via One-Line Install

### Step 1: Upload to GitHub

1. **Create a GitHub account** if you don't have one: https://github.com/join

2. **Create a new repository:**

   - Go to https://github.com/new
   - Name: `pdftool` (or any name you prefer)
   - Description: "PDF Merge Tool for Windows with context menu integration"
   - Public or Private (must be Public for raw file access)
   - Don't initialize with README (you already have one)
   - Click "Create repository"

3. **Upload your files to GitHub:**
   - On your new repository page, click "uploading an existing file"
   - Drag and drop these files:
     - `merge_pdfs.py`
     - `merge_pdf_handler.py`
     - `register_context_menu.ps1`
     - `unregister_context_menu.ps1`
     - `setup.ps1`
     - `install.ps1`
     - `README.md`
   - Add commit message: "Initial commit"
   - Click "Commit changes"

### Step 2: Update install.ps1

1. Edit `install.ps1` on GitHub (click the file → pencil icon)
2. Replace `YOUR_USERNAME` with your actual GitHub username
3. Change this line:
   ```powershell
   $baseUrl = "https://raw.githubusercontent.com/YOUR_USERNAME/pdftool/main"
   ```
   To:
   ```powershell
   $baseUrl = "https://raw.githubusercontent.com/YOUR_ACTUAL_USERNAME/pdftool/main"
   ```
4. Commit the changes

### Step 3: Share Your One-Line Installer

Your one-line installer command will be:

```powershell
irm https://raw.githubusercontent.com/YOUR_USERNAME/pdftool/main/install.ps1 | iex
```

Replace `YOUR_USERNAME` with your GitHub username.

### Example

If your GitHub username is `johnsmith`, the command would be:

```powershell
irm https://raw.githubusercontent.com/johnsmith/pdftool/main/install.ps1 | iex
```

## Alternative Distribution Methods

### Option 2: Using GitHub Releases

1. Go to your repository → Releases → "Create a new release"
2. Tag: `v1.0`
3. Title: `PDF Merge Tool v1.0`
4. Upload a ZIP file containing all the scripts
5. Users can download and extract

### Option 3: Using Gist

For simpler distribution of just the installer:

1. Go to https://gist.github.com
2. Create a new gist with `install.ps1`
3. Click "Create public gist"
4. Click "Raw" button
5. Copy the URL
6. Share: `irm YOUR_GIST_RAW_URL | iex`

### Option 4: Self-Hosted

If you have a web server:

1. Upload all files to your server
2. Update `$baseUrl` in `install.ps1` to your server URL
3. Make sure files are accessible via HTTPS
4. Share: `irm https://yoursite.com/pdftool/install.ps1 | iex`

## Security Notes

**For users downloading your tool:**

- They should review the install script before running
- PowerShell may require execution policy bypass
- Recommend running as Administrator for best results

**For you as the distributor:**

- Keep your repository updated
- Tag releases for version control
- Include a LICENSE file (MIT recommended for open source)
- Keep README updated with any changes

## Updating the Tool

When you make changes:

1. Update files on GitHub
2. Users can run the installer again to get the latest version
3. Or create a new release tag (v1.1, v1.2, etc.)

## Making It Even Easier

Create a short link using a service like:

- bit.ly
- tinyurl.com
- GitHub's own URL shortener

Example:

```powershell
irm bit.ly/pdfmerge | iex
```

Much easier to remember and share!
