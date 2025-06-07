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

## Converting PNG to Base64

The script `generate_qr_badges_final.py` embeds logo images as Base64 strings.
Convert your PNG files using the `base64` command (the `--wrap=0` flag keeps the
output on one line):

```bash
base64 --wrap=0 GT_full_logo.png > full_logo.b64
base64 --wrap=0 GT_small_logo.png > small_logo.b64
```

Copy the contents of these files into the `FULL_LOGO_B64` and `SMALL_LOGO_B64`
variables near the top of `generate_qr_badges_final.py` to embed your custom
logos.

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
If your Python environment is single-arch, use the macOS spec as-is (or set
`target_arch` to `arm64` or `x86_64`) and run:

```bash
# macOS build (single-arch)
pyinstaller QR_Badges_mac.spec
```

```
# Windows build
pyinstaller QR_Badges_win.spec
```

Each build produces a standalone application bundle. macOS users can create a DMG using `hdiutil` after running the spec.

---

5z4xbu-codex/add-license-file-and-reference-in-readme
*Developed by R. Clark Myers, M.Ed.*
*Contact: rmyers66@gatech.edu*

