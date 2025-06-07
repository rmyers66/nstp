# QR Badges App for macOS and Windows

Double-click **QR Badges.app** to launch.

## Usage

1. **Select your CSV**  
   Must have columns `Preferred,Last,Code` (Slate QR URLs).

2. **Generate & Open**  
   The app creates `yourfile_nametags.docx` alongside your CSV and opens it in Word.

3. **Print**
   Use Avery 5395 label sheets (2 cols × 4 rows per page).
4. **Optional Config**
   Provide a JSON or YAML file with `-c myconfig.yml` to override defaults like margins or label sizes.

## Version

Current app version: **3.6.1-postfix**

## Updating

- The app checks for updates at launch.  
- If a newer version is found, you’ll be prompted to download and install.

## Troubleshooting

- **No dialog?** Make sure you opened the app via its `.app` bundle, not the executable inside.
- **Permission issues?** Grant **Full Disk Access** to the app in **System Settings → Privacy & Security**.
- **Gatekeeper block?** Right-click the app, choose **Open**, then confirm.
- **Exit codes**: When run from the command line, the script returns `1` for
  failures such as missing dependencies or invalid input.

## License

This project is licensed under the [MIT License](LICENSE).

## PNG Assets

The UI now loads logo images from external PNG files. Place
`GT_full_logo.png`, `GT_small_logo.png`, and `GT_ribbon.png` in the same
directory as `generate_qr_badges_final.py`. When packaging with PyInstaller,
these files are included via the `datas` section of the spec files so they are
bundled with the application.

## Running from Source

If you prefer the command-line version, install the Python dependencies and run
the script directly:

```bash
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt

# Generate badges from your CSV file
python3 generate_qr_badges_final.py -i attendees.csv
```

The output Word document will be created alongside your input CSV.

## Building Standalone Apps

PyInstaller spec files are provided for macOS (`QR_Badges_mac.spec`) and Windows (`QR_Badges_win.spec`).

```
# macOS universal build
pyinstaller QR_Badges_mac.spec

# Windows build
pyinstaller QR_Badges_win.spec
```

Each build produces a standalone application bundle. macOS users can create a DMG using `hdiutil` after running the spec.

---

5z4xbu-codex/add-license-file-and-reference-in-readme
*Developed by R. Clark Myers, M.Ed.*
*Contact: rmyers66@gatech.edu*

