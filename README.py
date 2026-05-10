from playwright.sync_api import sync_playwright
from datetime import datetime
import os

URL = "https://gss.bmwgroup.net/board/14324/meetings"
DOWNLOAD_DIR = os.path.join(os.path.expanduser("~"), "Downloads")
os.makedirs(DOWNLOAD_DIR, exist_ok=True)

PREV_BTN_XPATH = (
    "xpath=/html/body/gss-app-root/div/div/div[2]/div/div[2]/gss-board-meetings/div[1]/div[1]/div/"
    "gss-shared-calendar/div/div[4]/gss-half-year-view/div/div[1]/div[1]/div/p-button[1]//button"
)


def click_safe(page, selector, timeout=10000):
    sel = selector.strip()
    if not sel.startswith("xpath="):
        sel = "xpath=" + sel
    loc = page.locator(sel)
    loc.wait_for(state="visible", timeout=30000)
    loc.scroll_into_view_if_needed(timeout=5000)
    try:
        loc.click(timeout=timeout)
    except Exception:
        loc.click(timeout=timeout, force=True)


def go_back_n_times(page, n):
    """Drückt den Zurück-Button genau n mal."""
    if n <= 0:
        print("📅 Kein Zurücknavigieren nötig.")
        return

    prev_btn = page.locator(PREV_BTN_XPATH)
    prev_btn.wait_for(state="visible", timeout=60000)

    for i in range(n):
        try:
            prev_btn.click(timeout=8000)
        except Exception:
            prev_btn.click(timeout=8000, force=True)
        page.wait_for_load_state("domcontentloaded")
        page.wait_for_timeout(500)

    print(f"📅 {n}x zurückgedrückt.")


def go_back_to_january_2021_and_count(page, max_clicks=80):
    """
    Navigiert zurück bis Januar 2021 und gibt zurück
    wie oft der Button gedrückt wurde.
    """
    prev_btn = page.locator(PREV_BTN_XPATH)
    prev_btn.wait_for(state="visible", timeout=60000)

    clicks = 0
    for _ in range(max_clicks):
        try:
            if page.locator("text=Januar 2021").first.is_visible(timeout=800):
                print(f"✅ Januar 2021 erreicht nach {clicks} Klicks.")
                return clicks
        except Exception:
            pass
        try:
            prev_btn.click(timeout=8000)
        except Exception:
            prev_btn.click(timeout=8000, force=True)
        page.wait_for_load_state("domcontentloaded")
        page.wait_for_timeout(500)
        clicks += 1

    raise RuntimeError("Januar 2021 nicht gefunden.")


def navigate_back(page, clicks):
    """Lädt die Seite neu, wartet bis Angular fertig ist, dann clicks-mal zurück."""
    page.goto(URL, wait_until="domcontentloaded")
    # Warten bis Angular-App vollständig initialisiert ist
    page.wait_for_timeout(2500)
    # Warten bis der Kalender-Button wirklich im DOM und klickbar ist
    prev_btn = page.locator(PREV_BTN_XPATH)
    prev_btn.wait_for(state="visible", timeout=30000)
    page.wait_for_timeout(1000)
    go_back_n_times(page, clicks)


def download_meeting(page):
    agenda_selector = (
        "xpath=/html/body/gss-app-root/div/div/div[2]/div/div[2]/gss-meeting-agenda"
        "/div/div[1]/div/div[2]/gss-meeting-documents-menu/div/p-button/button/span"
    )

    print("\n⏳ Warte auf deinen Klick auf einen Meeting-Eintrag...")
    try:
        page.locator(agenda_selector).wait_for(state="visible", timeout=0)
    except Exception as e:
        if "closed" in str(e).lower():
            return None
        raise

    print("✅ Meeting erkannt — starte Download-Flow...")
    page.wait_for_timeout(1000)

    click_safe(
        page,
        "/html/body/gss-app-root/div/div/div[2]/div/div[2]/gss-meeting-agenda"
        "/div/div[1]/div/div[2]/gss-meeting-documents-menu/div/p-button/button/span",
    )
    page.wait_for_timeout(2000)

    click_safe(
        page,
        "/html/body/gss-app-root/div/div/div[2]/div/div[2]/gss-meeting-agenda"
        "/div/div[1]/div/div[2]/gss-meeting-documents-menu/p-menu/div/ul/li[9]",
    )
    page.wait_for_timeout(2000)

    click_safe(
        page,
        "/html/body/gss-app-root/div/div/div[2]/div/div[2]/gss-meeting-documents-map"
        "/div/div/div[1]/p-button/button/span[2]",
    )
    page.wait_for_timeout(2000)

    print("📥 Starte Download...")
    try:
        with page.expect_download(timeout=60000) as download_info:
            click_safe(page, "/html/body/div/div/div[4]/p-button[2]/button")
            page.evaluate("""
                () => {
                    const origCreate = URL.createObjectURL;
                    URL.createObjectURL = function(blob) {
                        const url = origCreate.call(URL, blob);
                        const a = document.createElement('a');
                        a.href = url;
                        a.download = 'meeting.zip';
                        document.body.appendChild(a);
                        a.click();
                        document.body.removeChild(a);
                        URL.createObjectURL = origCreate;
                        return url;
                    };
                }
            """)

        download = download_info.value

        path = download.path()
        while path is None:
            page.wait_for_timeout(500)
            path = download.path()

        filename = download.suggested_filename or ""
        if not filename or not filename.endswith(".zip"):
            filename = f"meeting_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip"

        save_path = os.path.join(DOWNLOAD_DIR, filename)
        download.save_as(save_path)

        with open(save_path, "rb") as f:
            magic = f.read(2)

        if magic == b"PK":
            print(f"✅ Gültige ZIP gespeichert: {save_path}")
        else:
            print(f"⚠️ Datei gespeichert aber kein ZIP-Format (Magic: {magic.hex()})")

        return True

    except Exception as e:
        print(f"❌ Download fehlgeschlagen: {e}")
        return False


with sync_playwright() as p:
    ctx = p.chromium.launch_persistent_context(
        user_data_dir="profile",
        channel="chrome",
        headless=False,
        accept_downloads=True,
        downloads_path=DOWNLOAD_DIR,
        args=[
            "--disable-blink-features=AutomationControlled",
            "--no-sandbox",
            "--disable-infobars",
            "--start-maximized",
            f"--download-default-directory={DOWNLOAD_DIR}",
        ],
        ignore_default_args=["--enable-automation"],
    )

    page = ctx.new_page()

    client = page.context.new_cdp_session(page)
    client.send("Browser.setDownloadBehavior", {
        "behavior": "allow",
        "downloadPath": DOWNLOAD_DIR,
        "eventsEnabled": True,
    })

    page.add_init_script("""
        Object.defineProperty(navigator, 'webdriver', { get: () => undefined });
        Object.defineProperty(navigator, 'plugins', { get: () => [1, 2, 3] });
        window.chrome = { runtime: {} };
    """)

    # Erster Start: zu Januar 2021 navigieren und Klicks zählen
    page.goto(URL, wait_until="domcontentloaded")
    page.wait_for_timeout(1000)
    basis_klicks = go_back_to_january_2021_and_count(page)
    # basis_klicks = wie oft zurückgedrückt wurde um Jan 2021 zu erreichen
    # aktuelle_klicks = basis_klicks - (download_count // 8)
    # => alle 8 Downloads einen Klick weniger = einen Monat weiter vor

    print(f"\n📌 Basis: {basis_klicks} Klicks bis Januar 2021")
    print(">>> Klicke einen Meeting-Eintrag im Browser an.")
    print(">>> Alle 8 Downloads springt der Kalender einen Monat vor.")
    print(">>> Browser schließen zum Beenden.\n")

    download_count = 0

    while True:
        try:
            result = download_meeting(page)

            if result is None:
                print("\n🛑 Browser geschlossen. Script beendet.")
                break

            if result:
                download_count += 1
                monate_vor = download_count // 8
                aktuelle_klicks = basis_klicks - monate_vor
                print(f"📊 Downloads: {download_count} | Monat-Offset: +{monate_vor} | Zurück-Klicks: {aktuelle_klicks}")

                navigate_back(page, aktuelle_klicks)
                print(">>> Klicke den nächsten Meeting-Eintrag an...")

            else:
                print("\n⚠️ Übersprungen — navigiere zurück...")
                monate_vor = download_count // 8
                aktuelle_klicks = basis_klicks - monate_vor
                navigate_back(page, aktuelle_klicks)

        except Exception as e:
            if "closed" in str(e).lower():
                print("\n🛑 Browser geschlossen. Script beendet.")
                break
            print(f"\n❌ Fehler: {e}")
            try:
                monate_vor = download_count // 8
                aktuelle_klicks = basis_klicks - monate_vor
                navigate_back(page, aktuelle_klicks)
            except Exception:
                break

    ctx.close()
