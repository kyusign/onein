# Community Crawler Packaging Guide

This repository contains resources and notes for packaging the **Community Crawler** GUI
application into a standalone Windows executable.

## Requirements

The bundled application embeds all required Python libraries:

- `selenium`
- `webdriver-manager`
- `pandas`
- `openpyxl`
- `cryptography`

The end user must have Google Chrome installed. `webdriver-manager` will download the
matching ChromeDriver automatically. For offline environments, a copy of
`chromedriver.exe` is included in the package and used as a fallback.

## Building (PyInstaller)

```batch
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
pip install pyinstaller

pyinstaller ^
  --onefile --noconsole ^
  --name CommunityCrawler ^
  --add-binary "chromedriver.exe;." ^
  --icon icon.ico ^
  --collect-all pandas ^
  --collect-all openpyxl ^
  community_crawler_gui_hours.py
```

The resulting executable is written to `dist/CommunityCrawler.exe`.

## Building (Nuitka)

```batch
pip install nuitka
python -m nuitka ^
  --onefile ^
  --enable-plugin=tk-inter ^
  --windows-disable-console ^
  --include-data-files=chromedriver.exe=chromedriver.exe ^
  --include-data-files=icon.ico=icon.ico ^
  community_crawler_gui_hours.py
```

Rename the output to `CommunityCrawler.exe` for distribution.

## License File Lookup

The application first checks the standard application data directory for
`license.lic`. For portable use, placing `license.lic` beside the executable is
also supported.

## User Guide

1. Run `CommunityCrawler.exe`.
2. If prompted, load `license.lic` using **라이선스 불러오기**.
3. Select community, provide the list URL and time window, then choose an output
   path and click **실행**.
4. The results are saved to an Excel file with embedded watermark metadata.

