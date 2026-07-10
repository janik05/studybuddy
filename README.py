from datetime import date
from pathlib import Path
from typing import Any, Dict, Optional
import subprocess
import tempfile
import os


def _load_config(config_path: Optional[str] = None) -> Dict[str, Any]:
    config_file = Path(config_path) if config_path else Path(__file__).resolve().parent.parent / "config" / "config.yaml"
    data: Dict[str, Any] = {}
    if not config_file.exists():
        return data
    with config_file.open("r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#") or ":" not in line:
                continue
            key, value = line.split(":", 1)
            data[key.strip()] = value.strip().strip('"').strip("'")
    return data


def _ps_quote(value: str) -> str:
    return str(value).replace("'", "''")


def _build_powershell_script(sheet_name: str, target_name: str, new_header_value: str) -> str:
    sheet_q = _ps_quote(sheet_name)
    target_q = _ps_quote(target_name)
    header_q = _ps_quote(new_header_value)

    return f"""
$ErrorActionPreference = 'Stop'
$targetName = '{target_q}'
$sheetName  = '{sheet_q}'

$excel = $null; $wb = $null; $sheet = $null
for ($i = 0; $i -lt 40; $i++) {{
    try {{ $excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application') }}
    catch {{ Start-Sleep -Seconds 2; continue }}

    $wb = $null
    foreach ($w in $excel.Workbooks) {{
        if ($targetName -eq '' -or $w.Name.ToLower().Contains($targetName.ToLower())) {{ $wb = $w; break }}
    }}
    if ($null -eq $wb) {{ $wb = $excel.ActiveWorkbook }}

    if ($null -ne $wb) {{
        foreach ($s in $wb.Worksheets) {{ if ($s.Name -eq $sheetName) {{ $sheet = $s; break }} }}
        if ($null -ne $sheet) {{ break }}
    }}
    Start-Sleep -Seconds 2
}}

if ($null -eq $excel) {{ Write-Output 'ERROR: Keine laufende Excel-Instanz gefunden.'; exit 2 }}
if ($null -eq $wb)    {{ Write-Output 'ERROR: Keine offene Arbeitsmappe gefunden.'; exit 3 }}
if ($null -eq $sheet) {{ Write-Output "ERROR: Blatt '$sheetName' nicht gefunden."; exit 4 }}

# G24:G25 ist verbunden -> der Wert steckt NUR in der Anker-Zelle G24.
$oldHeader = $sheet.Range('G24').Value2

# Alten Inhalt von G24:G25 nach I24 und I25 sichern.
$sheet.Range('I24').Value2 = $oldHeader
$sheet.Range('I25').Value2 = $oldHeader

# G26:G30 als feste Werte nach I26:I30 kopieren.
for ($r = 26; $r -le 30; $r++) {{
    $sheet.Range("I$r").Value2 = $sheet.Range("G$r").Value2
}}

# Neuen Wert IMMER ueber die Anker-Zelle G24 in die verbundene Zelle schreiben.
$sheet.Range('G24').Value2 = '{header_q}'

$sheet.Activate()
$wb.Save()
Write-Output 'OK: Planning_Progress aktualisiert und gespeichert.'
exit 0
"""


def update_planning_progress(config_path: Optional[str] = None) -> bool:
    config = _load_config(config_path)
    sheet_name = str(config.get("target_sheet_name", "Planning_Progress")).strip() or "Planning_Progress"
    target_excel_name = str(config.get("target_excel_name", "")).strip()
    planning_label = str(config.get("planning_label", "")).strip()

    today = date.today().strftime("%d.%m.%Y")
    new_value = f"{planning_label} {today}".strip()

    script = _build_powershell_script(sheet_name, target_excel_name, new_value)

    tmp_path = None
    try:
        with tempfile.NamedTemporaryFile(mode="w", suffix=".ps1", delete=False, encoding="utf-8-sig") as tmp:
            tmp.write(script)
            tmp_path = tmp.name

        result = subprocess.run(
            ["powershell.exe", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", tmp_path],
            capture_output=True, text=True, timeout=120,
        )
        if result.stdout.strip():
            print(result.stdout.strip())
        if result.returncode != 0:
            if result.stderr.strip():
                print(f"PowerShell-Fehler: {result.stderr.strip()}")
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
