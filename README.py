from pathlib import Path
from typing import Any, Dict, Optional
from urllib.parse import urlparse, unquote
import os
import sys
import subprocess


def _load_config(config_path: Optional[str] = None) -> Optional[Dict[str, Any]]:
    if config_path is None:
        config_file = Path(__file__).resolve().parent.parent / "config" / "config.yaml"
    else:
        config_file = Path(config_path)

    if not config_file.exists():
        print(f"Konfigurationsdatei nicht gefunden: {config_file}")
        return None

    data: Dict[str, Any] = {}
    with config_file.open("r", encoding="utf-8") as handle:
        for raw_line in handle:
            line = raw_line.strip()
            if not line or line.startswith("#") or ":" not in line:
                continue

            key, value = line.split(":", 1)
            data[key.strip()] = value.strip().strip('"').strip("'")

    return data


def _create_client_context(config: Dict[str, Any]):
    try:
        from office365.sharepoint.client_context import ClientContext
        from office365.runtime.auth.client_credential import ClientCredential
    except ModuleNotFoundError:
        print("SharePoint-Abhaengigkeit fehlt: office365-rest-python-client")
        print("Bitte installieren: pip install -r requirements.txt")
        return None

    site_url = str(config.get("site_url", "")).strip()
    client_id = str(config.get("client_id", "")).strip()
    client_secret = str(config.get("client_secret", "")).strip()

    if not site_url:
        print("config.yaml: 'site_url' fehlt.")
        return None

    if not client_id or not client_secret:
        print("config.yaml: 'client_id' und/oder 'client_secret' fehlen.")
        return None

    credentials = ClientCredential(client_id, client_secret)
    return ClientContext(site_url).with_credentials(credentials)


def _build_web_url(site_url: str, server_relative_url: str) -> str:
    parsed = urlparse(site_url)
    return f"{parsed.scheme}://{parsed.netloc}{server_relative_url}?web=1"


def _build_direct_excel_url(config: Dict[str, Any]) -> str:
    site_url = str(config.get("site_url", "")).strip()
    folder_relative_url = str(config.get("folder_server_relative_url", "")).strip().rstrip("/")
    target_excel_name = str(config.get("target_excel_name", "")).strip()

    if not site_url or not folder_relative_url or not target_excel_name:
        return ""

    return _build_web_url(site_url, f"{folder_relative_url}/{target_excel_name}")


def _open_with_excel_desktop(url: str, sheet_name: Optional[str] = None) -> bool:
    if not url:
        return False

    try:
        if sys.platform == "win32":
            # ms-excel:ofe|u|<url> zwingt das Oeffnen in Excel Desktop (Open For Edit)
            # und haelt die Datei mit SharePoint verbunden (Speichern geht nach SharePoint).
            excel_protocol_url = f"ms-excel:ofe|u|{url}"
            os.startfile(excel_protocol_url)
            print(f"Excel Desktop geoeffnet: {url}")
            # Das Aktivieren des Blattes und die Bearbeitung erfolgen in planning_progress.py
            # ueber PowerShell-COM (kein pywin32 noetig).
        else:
            subprocess.Popen(["open", url])
            print(f"Excel Desktop geoeffnet: {url}")

        return True
    except Exception as exc:
        print(f"Fehler beim Oeffnen mit Excel: {exc}")
        return False


def _open_direct_excel_url(config: Dict[str, Any]) -> bool:
    direct_excel_url = str(config.get("direct_excel_url", "")).strip()
    target_excel_name = str(config.get("target_excel_name", "")).strip().lower()
    target_sheet_name = str(config.get("target_sheet_name", "")).strip() or None

    if direct_excel_url and target_excel_name and target_excel_name in direct_excel_url.lower():
        return _open_with_excel_desktop(direct_excel_url, target_sheet_name)

    built_excel_url = _build_direct_excel_url(config)
    if built_excel_url:
        return _open_with_excel_desktop(built_excel_url, target_sheet_name)

    folder_web_url = str(config.get("folder_web_url", "")).strip()
    if folder_web_url:
        print(f"SharePoint-Ordner wird geoeffnet (keine direkte Excel-Datei gefunden): {folder_web_url}")
        return _open_with_excel_desktop(folder_web_url, None)

    if not str(config.get("site_url", "")).strip():
        print("config.yaml: 'site_url' fehlt.")
        return False

    if not str(config.get("folder_server_relative_url", "")).strip():
        print("config.yaml: 'folder_server_relative_url' fehlt.")
        return False

    if not str(config.get("target_excel_name", "")).strip():
        print("config.yaml: 'target_excel_name' fehlt.")
        return False

    print("Kein gueltiger Excel-Link in config.yaml gefunden.")
    return False


def open_excel_from_sharepoint(config_path: Optional[str] = None) -> bool:
    config = _load_config(config_path)
    if config is None:
        return False

    folder_relative_url = str(config.get("folder_server_relative_url", "")).strip()
    target_excel_name = str(config.get("target_excel_name", "")).strip()
    site_url = str(config.get("site_url", "")).strip()

    if not folder_relative_url:
        print("config.yaml: 'folder_server_relative_url' fehlt.")
        return False

    if not target_excel_name:
        print("config.yaml: 'target_excel_name' fehlt.")
        return False

    client_id = str(config.get("client_id", "")).strip()
    client_secret = str(config.get("client_secret", "")).strip()

    if not client_id or not client_secret:
        print("Kein Azure-App-Zugriff vorhanden. Direkter Excel-Link wird verwendet.")
        return _open_direct_excel_url(config)

    context = _create_client_context(config)
    if context is None:
        return False

    try:
        folder = context.web.get_folder_by_server_relative_url(folder_relative_url)
        files = folder.files
        context.load(files)
        context.execute_query()
    except Exception as exc:
        print(f"Fehler beim Zugriff auf SharePoint-Ordner: {exc}")
        return False

    target = None
    for item in files:
        name = str(item.properties.get("Name", ""))
        if name.lower() == target_excel_name.lower():
            target = item
            break

    if target is None:
        available = [str(item.properties.get("Name", "")) for item in files]
        print(f"Excel-Datei nicht gefunden: {target_excel_name}")
        print(f"Verfuegbare Dateien: {', '.join(available)}")
        return False

    server_relative_url = str(target.properties.get("ServerRelativeUrl", "")).strip()
    if not server_relative_url:
        print("ServerRelativeUrl konnte nicht gelesen werden.")
        return False

    target_sheet_name = str(config.get("target_sheet_name", "")).strip() or None
    open_url = _build_web_url(site_url, server_relative_url)
    return _open_with_excel_desktop(open_url, target_sheet_name)


def download_files() -> bool:
    return open_excel_from_sharepoint()


anderes File
from datetime import date
from pathlib import Path
from typing import Any, Dict, Optional
import subprocess
import tempfile
import os


def _load_config(config_path: Optional[str] = None) -> Dict[str, Any]:
    if config_path is None:
        config_file = Path(__file__).resolve().parent.parent / "config" / "config.yaml"
    else:
        config_file = Path(config_path)

    data: Dict[str, Any] = {}
    if not config_file.exists():
        return data

    with config_file.open("r", encoding="utf-8") as handle:
        for raw_line in handle:
            line = raw_line.strip()
            if not line or line.startswith("#") or ":" not in line:
                continue

            key, value = line.split(":", 1)
            data[key.strip()] = value.strip().strip('"').strip("'")

    return data


def _ps_quote(value: str) -> str:
    # Einfache Anfuehrungszeichen fuer PowerShell escapen.
    return str(value).replace("'", "''")


def _build_powershell_script(sheet_name: str, target_name: str, g25_new_value: str) -> str:
    sheet_q = _ps_quote(sheet_name)
    target_q = _ps_quote(target_name)
    g25_q = _ps_quote(g25_new_value)

    return f"""
$ErrorActionPreference = 'Stop'

$targetName = '{target_q}'
$sheetName = '{sheet_q}'

# Warten, bis Excel + Arbeitsmappe + Blatt bereit sind (SharePoint-Download dauert ggf.).
$excel = $null
$wb = $null
$sheet = $null
for ($attempt = 0; $attempt -lt 40; $attempt++) {{
    try {{
        $excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
    }} catch {{
        Start-Sleep -Seconds 2
        continue
    }}

    $wb = $null
    foreach ($w in $excel.Workbooks) {{
        if ($targetName -eq '' -or $w.Name.ToLower().Contains($targetName.ToLower())) {{
            $wb = $w
            break
        }}
    }}
    if ($null -eq $wb) {{ $wb = $excel.ActiveWorkbook }}

    if ($null -ne $wb) {{
        $sheet = $null
        foreach ($s in $wb.Worksheets) {{
            if ($s.Name -eq $sheetName) {{ $sheet = $s; break }}
        }}
        if ($null -ne $sheet) {{ break }}
    }}

    Start-Sleep -Seconds 2
}}

if ($null -eq $excel) {{
    Write-Output 'ERROR: Keine laufende Excel-Instanz gefunden. Bitte Datei in Excel Desktop oeffnen.'
    exit 2
}}
if ($null -eq $wb) {{
    Write-Output 'ERROR: Keine offene Arbeitsmappe gefunden.'
    exit 3
}}
if ($null -eq $sheet) {{
    $names = ($wb.Worksheets | ForEach-Object {{ $_.Name }}) -join ', '
    Write-Output "ERROR: Blatt '$sheetName' nicht gefunden. Verfuegbar: $names"
    exit 4
}}

# Feste Werte aus G25:G30 einlesen (Value2 = berechneter Wert, keine Formel).
$vals = @()
for ($r = 25; $r -le 30; $r++) {{
    $vals += $sheet.Range("G$r").Value2
}}
$originalG25 = $vals[0]

# I24 = urspruenglicher Wert aus G25.
$sheet.Range('I24').Value2 = $originalG25

# I25:I30 = feste Werte aus G25:G30.
for ($i = 0; $i -lt 6; $i++) {{
    $row = 25 + $i
    $sheet.Range("I$row").Value2 = $vals[$i]
}}

# G25 mit neuem Text ueberschreiben.
$sheet.Range('G25').Value2 = '{g25_q}'

$sheet.Activate()
$wb.Save()
Write-Output 'OK: Planning_Progress aktualisiert und gespeichert.'
exit 0
"""


def update_planning_progress(config_path: Optional[str] = None) -> bool:
    config = _load_config(config_path)
    sheet_name = str(config.get("target_sheet_name", "Planning_Progress")).strip() or "Planning_Progress"
    planning_label = str(config.get("planning_label", "Planung")).strip() or "Planung"
    target_excel_name = str(config.get("target_excel_name", "")).strip()

    today = date.today().strftime("%d.%m.%Y")
    g25_new_value = f"{planning_label} {today}"

    script = _build_powershell_script(sheet_name, target_excel_name, g25_new_value)

    tmp_path = None
    try:
        with tempfile.NamedTemporaryFile(
            mode="w", suffix=".ps1", delete=False, encoding="utf-8-sig"
        ) as tmp:
            tmp.write(script)
            tmp_path = tmp.name

        result = subprocess.run(
            [
                "powershell.exe",
                "-NoProfile",
                "-ExecutionPolicy",
                "Bypass",
                "-File",
                tmp_path,
            ],
            capture_output=True,
            text=True,
            timeout=120,
        )

        output = (result.stdout or "").strip()
        error = (result.stderr or "").strip()

        if output:
            print(output)
        if result.returncode != 0:
            if error:
                print(f"PowerShell-Fehler: {error}")
            return False

        return True
    except FileNotFoundError:
        print("PowerShell wurde nicht gefunden (nur unter Windows verfuegbar).")
        return False
    except subprocess.TimeoutExpired:
        print("Zeitueberschreitung beim Aktualisieren der Excel-Datei.")
        return False
    finally:
        if tmp_path:
            try:
                os.remove(tmp_path)
            except OSError:
                pass


if __name__ == "__main__":
    update_planning_progress()


