# 🚀 LabFolder Backup Tool - Quick Start Guide

Welcome! This guide will help you get started with the LabFolder Backup Tool in just a few minutes.



## ⚖️ Important Disclaimer

**This tool is specifically designed for LMU Medical Faculty's LabFolder instance.** Other institutions may need to modify the code based on their API configuration. 

**Use at your own risk.** Always verify the completeness and accuracy of your backups. This tool is provided as-is without any warranties.



## 📋 Prerequisites

Before you begin, you'll need:

- **Google Chrome browser** (other browsers may work but are untested)
- **LabFolder account credentials** (username and password)
- **Basic familiarity** with your operating system's file system



## ⚙️ Configuration

### Step 1: Create Your Configuration File

1. Download the file named `.env.example` in the tool directory (https://raw.githubusercontent.com/jiehua1995/LabFolder_Backup/refs/heads/main/.env.example)
2. Make a copy and rename it to `.env` (remove the `.example` part)
3. Open `.env` in a text editor
4. Update the following fields with your credentials:
   ```bash
   LABFOLDER_USERNAME=your_username
   LABFOLDER_PASSWORD=your_password
   DOWNLOAD_DIR=./Download
   ```
5. **Leave all other settings unchanged** unless you have specific requirements
6. Save the file

⚠️ **Security Note**: Keep your `.env` file private! Never share it.



## 🎯 Choose Your Method

### Option 1: Executable File (Easiest) 

**Recommended for non-technical users**

#### Requirements:
- ✅ Google Chrome installed
- ✅ Windows 10 or 11
- ✅ `.env` file configured (see above)

#### Steps:
1. **Download** the executable file:
   - 📥 [Download LabFolder Backup Tool (Windows)](https://github.com/jiehua1995/LabFolder_Backup/releases/download/v0.1.8/LabFolder.Archive_win_v0.1.8.exe)
2. **Place** the executable in a folder where you want to store backups
3. **Ensure** your `.env` file is in the same directory as the executable
4. **Double-click** the executable file to run it
5. **Follow the GUI instructions**, including clicking the "Run Backup" button to start the backup process.
6. **Sit back** and let it work! ☕

The tool will create a folder named as what you specified in the `DOWNLOAD_DIR`.



### Option 2: Python Script (Advanced)

**For users comfortable with command-line interfaces**

#### Requirements:
- ✅ Python 3.10 or higher installed
- ✅ Google Chrome installed
- ✅ `.env` file configured (see above)

#### Steps:

1. **Create a virtual environment** (recommended):
   ```bash
   conda create -n labfolder_backup python=3.10
   conda activate labfolder_backup
   ```

2. **Install required packages**:
   ```bash
   pip install -r requirements.txt
   ```

3. **Run the backup script**:
   
   **Command-line version** (detailed output):
   ```bash
   python labfolder_backup_v4.py
   ```
   
   **OR**
   
   **GUI version** (graphical interface with buttons):
   ```bash
   python labfolder_gui.py
   ```



## 📊 What to Expect: Step-by-Step Process

The backup tool performs **8 major steps**. Here's what happens at each stage:

### [1/8] 🔐 Logging in
- Authenticates with your LabFolder instance
- Verifies your credentials
- **Expected output**: `✓ Login successful`

### [2/8] 📥 Fetching data
- Retrieves list of all projects and entries
- Queries the LabFolder API
- **Expected output**: Number of projects and entries found

### [2.5/8] 🔐 Capturing Authorization Token
- **⚠️ BROWSER WINDOW ALERT**: A Chrome window will open automatically
- Extracts authentication token for API access
- **Duration**: ~30 seconds
- **⚠️ IMPORTANT**: 
  - **DO NOT** click or interact with the browser window
  - **MOVE** your mouse outside the browser window
  - Let the automation complete on its own

### [3/8] 📦 Processing entries
- Downloads all content from each entry:
  - 📝 Text blocks (HTML content)
  - 🖼️ Images (original + annotated versions)
  - 📎 Files (attachments)
  - 📊 Tables (JSON data)
  - 🧪 Well Plates (JSON data)
  - 📈 Data elements
  - ✏️ Sketches
- **Expected output**: Progress bar showing download status

### [4/8] 📥 Downloading XLSX files
- **⚠️ BROWSER WINDOW ALERT**: A new Chrome window will open
- Automates browser to download Excel files for tables and wellplates
- Uses smart scrolling to reveal download buttons
- **Duration**: ~30 seconds per file
- **⚠️ IMPORTANT**: 
  - **DO NOT** move your mouse into the browser window
  - **DO NOT** click anything
  - **MOVE** your mouse to:
    - Another monitor (if available)
    - Outside the browser boundaries
    - The browser's title bar (safe zone)
  - The script will display: "🚀 Starting in 3 seconds..." before launching
- **Expected output**: Progress bar showing XLSX downloads

### [5/8] 🎨 Generating HTML reports
- Creates HTML pages for each entry
- Embeds text, pictures and files directly in HTML
- Generates main index page with project grouping
- **Expected output**: "✓ Created X entry HTML pages"

### [6/7] 📋 Exporting to ELN format
- Creates RO-Crate compliant ELN packages
- Calculates SHA256 hashes for file integrity
- Packages each entry as a `.eln` file
- **Expected output**: "✓ Created X ELN packages"

### [7/8] ✨ Finishing up
- Cleans up temporary files
- Calculates final statistics
- Verifies backup integrity

### [8/8] ✨ Backup Complete!
- Displays comprehensive summary
- Shows total backup size and location
- Provides next steps



## 🎯 After Backup Completes

### What You'll Have:

```
labfolder_backup/
├── Project_Name_1/
│   ├── 12345_Entry_Title/
│   │   ├── index.html          (View this in browser!)
│   │   ├── image_001.png
│   │   ├── file_attachment.pdf
│   │   ├── table_123.json
│   │   └── table_123_..._Table.xlsx
│   └── 67890_Another_Entry/
│       └── ...
├── Project_Name_2/
│   └── ...
├── eln_export/
│   ├── 12345_Entry_Title.eln   (RO-Crate format)
│   └── 67890_Another_Entry.eln
└── index.html                  (⭐ START HERE!)
```

### Next Steps:

1. **View your backup**:
   - Open `labfolder_backup/index.html` in any web browser
   - Browse through your entries with a beautiful interface
   - All tables and wellplates are embedded and viewable

2. **Verify completeness**:
   - Check that all entries are present
   - Verify XLSX files downloaded correctly
   - Look for any warning messages in the output

3. **Archive for long-term storage**:
   - Copy the entire `labfolder_backup` folder to external storage
   - Consider creating a ZIP archive
   - Store in multiple locations for redundancy

4. **ELN files**:
   - Find your `.eln` files in `labfolder_backup/eln_export/`
   - These can be imported into other ELN systems
   - Each file is RO-Crate compliant



## ⚙️ Advanced Configuration

### Adjusting Scroll Offset (If Buttons Are Hidden)

If you notice that download buttons are obscured during XLSX downloads:

1. Open your `.env` file
2. Modify this line:
   ```bash
   SCROLL_NUDGE_PIXELS=150
   ```
3. Try different values (120, 150, 200) until buttons appear correctly. The resolution of your screen could affect this. I only tested on 1920x1080.

### Custom Download Directory

To change where backups are saved:

1. Open `.env`
2. Modify:
   ```bash
   DOWNLOAD_DIR=./customized_name
   ```
### Build binary executable yourself
If you want to build the executable file on your own machine:
1. Ensure you have PyInstaller installed:
   ```bash
   pip install pyinstaller
   ```
2. Run PyInstaller to create the executable:
   ```bash
   pyinstaller labfolder_gui.py --onefile --name "LabFolder Archive" --icon image.ico --add-data "labfolder_backup_v4.py;."
   ```
You can then find the executable in the `dist` folder created by PyInstaller. Your can also change the image.ico to your own icon file if desired.


## 🐛 Troubleshooting

### "Login failed" error
- ✅ Verify your credentials in `.env`
- ✅ Check that your LabFolder instance URL is correct (only supported for LMU Medical Faculty. do not change it.)
- ✅ Ensure you can log in manually via web browser

### Chrome crashes or automation fails
- ✅ Update Chrome to the latest version
- ✅ Ensure you're not interacting with the browser during automation
- ✅ Try increasing `SCROLL_NUDGE_PIXELS` if buttons are hidden

### Missing XLSX files
- ✅ Check the console output for specific error messages
- ✅ Some entries may not have downloadable tables
- ✅ Verify that the failed downloads are noted in the summary

### "Selenium not installed" warning
- ✅ Reinstall requirements: `pip install -r requirements.txt`
- ✅ Ensure you're using the virtual environment






## 📄 Disclaimer

**This tool is provided “as is” without any warranty of any kind, express or implied. The author, lab, and LMU Munich are not responsible for any data loss, system errors, or other damages arising from the use of this software. Users should verify all outputs and use the tool at your own risk.**






## ✨ Tips for Best Results

- 🕐 **Run during off-peak hours** for faster downloads
- ☕ **Be patient** - large notebooks may take 30+ minutes
- 💾 **Keep multiple backup copies** in different locations
- 🔄 **Try different tools** to backup your data.
- 🔍 **Verify the data** after backup.




## 🧾 License

This project is licensed under the **MIT License with Commercial Restriction**.

- 🧑‍🔬 Academic, personal, and non-commercial use: **Free**
- 💼 Commercial use: **Contact me**
