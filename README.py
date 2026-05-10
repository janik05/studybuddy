from playwright.sync_api import sync_playwright
from datetime import datetime
import os

URL = "https://gss.bmwgroup.net/board/14324/meetings"
DOWNLOAD_DIR = os.path.join(os.path.expanduser("~"), "Downloads")
os.makedirs(DOWNLOAD_DIR, exist_ok=True)

NEXT_BTN_XPATH = (
    "xpath=/html/body/gss-app-root/div/div/div[2]/div/div[2]/gss-board-meetings/div[1]/div[1]/div/"
    "gss-shared-calendar/div/div[4]/gss-half-year-view/div/div[1]/div[1]/div/p-button[2]//button"
)
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


def go_back_to_january_2021(page, max_clicks=80):
    """Navigiert zurück bis Januar 2021."""
    prev_btn = page.locator(PREV_BTN_XPATH)
    prev_btn.wait_for(state="visible", timeout=60000)

    for _ in range(max_clicks):
        try:
            if page.locator("text=Januar 2021").first.is_visible(timeout=800):
                print("✅ Januar 2021 erreicht.")
                return
        except Exception:
            pass
        try:
            prev_btn.click(timeout=8000)
        except Exception:
            prev_btn.click(timeout=8000, force=True)
        page.wait_for_load_state("domcontentloaded")
        page.wait_for_timeout(500)

    raise RuntimeError("Januar 2021 nicht gefunden.")


def go_forward_months(page, months):
    """Springt N Monate vorwärts vom aktuellen Kalenderstand."""
    if months <= 0:
        return
    next_btn = page.locator(NEXT_BTN_XPATH)
    next_btn.wait_for(state="visible", timeout=10000)
    for i in range(months):
        try:
            next_btn.click(timeout=8000)
        except Exception:
            next_btn.click(timeout=8000, force=True)
        page.wait_for_load_state("domcontentloaded")
        page.wait_for_timeout(600)
    print(f"📅 {months} Monat(e) vorgesprungen.")


def navigate_to_month(page, target_month_offset):
    """
    Geht immer zuerst zu Januar 2021 zurück,
    dann target_month_offset Monate vor.
    """
    page.goto(URL, wait_until="domcontentloaded")
    page.wait_for_timeout(1000)
    go_back_to_january_2021(page)
    if target_month_offset > 0:
        go_forward_months(page, target_month_offset)


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

    # Starte bei Januar 2021
    current_month_offset = 0  # 0 = Januar 2021
    navigate_to_month(page, current_month_offset)

    print("\n>>> Klicke einen Meeting-Eintrag im Browser an.")
    print(">>> Nach jedem 2. Download springt der Kalender automatisch einen Monat vor.")
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
                print(f"📊 Downloads: {download_count} | Aktueller Monat-Offset: +{current_month_offset} ab Jan 2021")

                # Nach jedem 2. Download: einen Monat vor
                if download_count % 2 == 0:
                    current_month_offset += 1
                    print(f"⏭️  Springe zu Monat +{current_month_offset} ab Januar 2021...")
                    navigate_to_month(page, current_month_offset)
                else:
                    # Zurück zum gleichen Monat
                    navigate_to_month(page, current_month_offset)
                    print(">>> Klicke den nächsten Meeting-Eintrag an...")

            else:
                print("\n⚠️ Übersprungen — zurück zum gleichen Monat...")
                navigate_to_month(page, current_month_offset)

        except Exception as e:
            if "closed" in str(e).lower():
                print("\n🛑 Browser geschlossen. Script beendet.")
                break
            print(f"\n❌ Fehler: {e}")
            try:
                navigate_to_month(page, current_month_offset)
            except Exception:
                break

    ctx.close()
