# studybuddy
from playwright.sync_api import sync_playwright
import os
import re
import urllib.parse
import requests
 
URL = "https://gss.bmwgroup.net/board/14324/meetings"
DOWNLOAD_DIR = os.path.join(os.path.expanduser("~"), "Downloads")
 
 
def click_safe(page, selector):
    sel = selector.strip()
    if not sel.startswith("xpath="):
        sel = "xpath=" + sel
 
    loc = page.locator(sel)
    loc.wait_for(state="visible", timeout=30000)
    loc.scroll_into_view_if_needed(timeout=5000)
    try:
        loc.click(timeout=8000)
    except Exception:
        loc.click(timeout=8000, force=True)
 
 
def download_with_cookies(page, download_url):
    """Download-URL mit den Browser-Cookies per requests herunterladen."""
    cookies = page.context.cookies()
    session = requests.Session()
    for c in cookies:
        session.cookies.set(c["name"], c["value"], domain=c.get("domain", ""))
 
    resp = session.get(download_url, stream=True, timeout=120, allow_redirects=True)
    resp.raise_for_status()
 
    # Dateinamen aus Content-Disposition Header
    filename = None
    content_disp = resp.headers.get("Content-Disposition", "")
    if content_disp:
        # filename*=UTF-8''Name.zip
        match = re.search(r"filename\*=(?:UTF-8''|utf-8'')(.+)", content_disp, re.IGNORECASE)
        if match:
            filename = urllib.parse.unquote(match.group(1).strip().strip('"'))
        else:
            # filename="Name.zip"
            match = re.search(r'filename="?([^";\r\n]+)"?', content_disp)
            if match:
                filename = match.group(1).strip()
 
    # Fallback: aus URL
    if not filename:
        filename = download_url.split("/")[-1].split("?")[0]
 
    if not filename or not filename.endswith(".zip"):
        filename = (filename or "download") + ".zip"
 
    save_path = os.path.join(DOWNLOAD_DIR, filename)
    with open(save_path, "wb") as f:
        for chunk in resp.iter_content(chunk_size=8192):
            f.write(chunk)
 
    return save_path
 
 
def go_back_to_january_2021(page, max_clicks=80):
    prev_btn = page.locator(
        "xpath=/html/body/gss-app-root/div/div/div[2]/div/div[2]/gss-board-meetings/div[1]/div[1]/div/"
        "gss-shared-calendar/div/div[4]/gss-half-year-view/div/div[1]/div[1]/div/p-button[1]//button"
    )
 
    prev_btn.wait_for(state="visible", timeout=60000)
 
    for _ in range(max_clicks):
        try:
            if page.locator("text=Januar 2021").first.is_visible(timeout=800):
                return
        except Exception:
            pass
 
        try:
            prev_btn.click(timeout=8000)
        except Exception:
            prev_btn.click(timeout=8000, force=True)
 
        page.wait_for_load_state("domcontentloaded")
        page.wait_for_timeout(500)
 
    raise RuntimeError("Januar 2021 nicht gefunden (max_clicks erreicht).")
 
 
with sync_playwright() as p:
    ctx = p.chromium.launch_persistent_context(
        user_data_dir="profile",
        channel="chrome",
        headless=False,
        accept_downloads=True,
    )
    page = ctx.new_page()
    page.goto(URL, wait_until="domcontentloaded")
 
    # Automatisch bis Januar 2021 zurückgehen
    go_back_to_january_2021(page)
 
    # Merke die URL der Meeting-Liste
    meetings_url = page.url
 
    print("\n>>> Klicke jetzt einen Meeting-Eintrag im Browser an.")
    print(">>> Das Script erkennt deinen Klick automatisch und lädt die ZIP herunter.")
    print(">>> Schließe den Browser zum Beenden.\n")
 
    while True:
        try:
            # Warte bis der User einen Meeting-Eintrag anklickt
            # => Die URL ändert sich oder das Meeting-Agenda-Element erscheint
            agenda_selector = (
                "xpath=/html/body/gss-app-root/div/div/div[2]/div/div[2]/gss-meeting-agenda"
                "/div/div[1]/div/div[2]/gss-meeting-documents-menu/div/p-button/button/span"
            )
            print("⏳ Warte auf deinen Klick auf einen Meeting-Eintrag...")
            page.locator(agenda_selector).wait_for(state="visible", timeout=0)
            print("✅ Meeting-Eintrag erkannt! Starte automatischen Download...")
 
            page.wait_for_timeout(1000)
 
            # Dokument-Menü öffnen
            click_safe(
                page,
                "/html/body/gss-app-root/div/div/div[2]/div/div[2]/gss-meeting-agenda"
                "/div/div[1]/div/div[2]/gss-meeting-documents-menu/div/p-button/button/span",
            )
            page.wait_for_timeout(2000)
 
            # Menüpunkt auswählen (9. Eintrag)
            click_safe(
                page,
                "/html/body/gss-app-root/div/div/div[2]/div/div[2]/gss-meeting-agenda"
                "/div/div[1]/div/div[2]/gss-meeting-documents-menu/p-menu/div/ul/li[9]",
            )
            page.wait_for_timeout(2000)
 
            # Download-Button klicken
            click_safe(
                page,
                "/html/body/gss-app-root/div/div/div[2]/div/div[2]/gss-meeting-documents-map"
                "/div/div/div[1]/p-button/button/span[2]",
            )
            page.wait_for_timeout(2000)
 
            # Download-URL abfangen statt Playwright-Download
            download_url = [None]
 
            def intercept_download(route):
                download_url[0] = route.request.url
                route.abort()  # Download in Playwright abbrechen
 
            # Alle Requests auf typische Download-Muster abfangen
            page.route("**/*", intercept_download)
 
            click_safe(page, "/html/body/div/div/div[4]/p-button[2]/button")
            page.wait_for_timeout(3000)
 
            page.unroute("**/*")
 
            if download_url[0]:
                print(f"📥 Download-URL gefunden, lade herunter...")
                save_path = download_with_cookies(page, download_url[0])
                print(f"\n✅ ZIP gespeichert: {save_path}")
            else:
                print("\n❌ Keine Download-URL gefunden.")
 
            # Zurück zur Meeting-Liste
            page.goto(meetings_url, wait_until="domcontentloaded")
            page.wait_for_timeout(1000)
 
            print("\n>>> Klicke den nächsten Meeting-Eintrag an...")
 
        except Exception as e:
            if "Target page, context or browser has been closed" in str(e):
                print("\n🛑 Browser geschlossen. Script beendet.")
                break
            print(f"\n❌ Fehler: {e}")
            try:
                page.goto(meetings_url, wait_until="domcontentloaded")
                page.wait_for_timeout(1000)
            except Exception:
                break
 
    ctx.close()
    
