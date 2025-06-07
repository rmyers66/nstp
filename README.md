# QR Badges App for macOS

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

Current app version: **1.0.0**

## Updating

- The app checks for updates at launch.  
- If a newer version is found, you’ll be prompted to download and install.

## Troubleshooting

- **No dialog?** Make sure you opened the app via its `.app` bundle, not the executable inside.  
- **Permission issues?** Grant **Full Disk Access** to the app in **System Settings → Privacy & Security**.  
- **Gatekeeper block?** Right-click the app, choose **Open**, then confirm.

---

*Developed by Your Name*  
*Contact: you@example.com*
