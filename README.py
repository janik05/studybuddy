from playwright.sync_api import sync_playwright
import os

URL = "https://gss.bmwgroup.net/board/14324/meetings"
DOWNLOAD_DIR = os.path.join(os.path.expanduser("~"), "Downloads")
os.makedirs(DOWNLOAD_DIR, exist_ok=True)


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
    prev_btn = page.locator(
        "xpath=/html/body/gss-app-root/div/div/div[2]/div/div[2]/gss-board-meetings/div[1]/div[1]/div/"
        "gss-shared-calendar/div/div[4]/gss-half-year-view/div/div[1]/div[1]/div/p-button[1]//button"
    )
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


def download_meeting(page, meetings_url):
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

    # Schritt 1: Dokument-Menü öffnen
    click_safe(
        page,
        "/html/body/gss-app-root/div/div/div[2]/div/div[2]/gss-meeting-agenda"
        "/div/div[1]/div/div[2]/gss-meeting-documents-menu/div/p-button/button/span",
    )
    page.wait_for_timeout(2000)

    # Schritt 2: Menüpunkt 9 wählen
    click_safe(
        page,
        "/html/body/gss-app-root/div/div/div[2]/div/div[2]/gss-meeting-agenda"
        "/div/div[1]/div/div[2]/gss-meeting-documents-menu/p-menu/div/ul/li[9]",
    )
    page.wait_for_timeout(2000)

    # Schritt 3: Download-Button
    click_safe(
        page,
        "/html/body/gss-app-root/div/div/div[2]/div/div[2]/gss-meeting-documents-map"
        "/div/div/div[1]/p-button/button/span[2]",
    )
    page.wait_for_timeout(2000)

    # Schritt 4: ZIP herunterladen — expect_download fängt auch Blob-Downloads ab
    try:
        with page.expect_download(timeout=60000) as download_info:
            click_safe(page, "/html/body/div/div/div[4]/p-button[2]/button")

        download = download_info.value
        filename = download.suggested_filename or "meeting.zip"
        if not filename.endswith(".zip"):
            filename += ".zip"

        save_path = os.path.join(DOWNLOAD_DIR, filename)
        download.save_as(save_path)
        print(f"✅ ZIP gespeichert: {save_path}")
        return True

    except Exception as e:
        print(f"❌ Download fehlgeschlagen: {e}")

        # Fallback: fetch() im Browser-Kontext
        print("🔄 Versuche Fallback...")
        captured = {"url": None}

        def intercept(route):
            url = route.request.url
            if any(x in url for x in [".zip", "download", "export"]):
                captured["url"] = url
            route.continue_()

        page.route("**/*", intercept)
        try:
            click_safe(page, "/html/body/div/div/div[4]/p-button[2]/button")
            page.wait_for_timeout(5000)
        except Exception:
            pass
        page.unroute("**/*")

        if captured["url"]:
            result = page.evaluate(
                """async (url) => {
                    const resp = await fetch(url, { credentials: 'include' });
                    if (!resp.ok) return null;
                    const blob = await resp.blob();
                    const buf = await blob.arrayBuffer();
                    return Array.from(new Uint8Array(buf));
                }""",
                captured["url"],
            )
            if result:
                save_path = os.path.join(DOWNLOAD_DIR, "meeting_fallback.zip")
                with open(save_path, "wb") as f:
                    f.write(bytes(result))
                print(f"✅ ZIP gespeichert (Fallback): {save_path}")
                return True

        print("❌ Kein Download möglich.")
        return False


with sync_playwright() as p:
    ctx = p.chromium.launch_persistent_context(
        user_data_dir="profile",
        channel="chrome",
        headless=False,
        accept_downloads=True,
        downloads_path=DOWNLOAD_DIR,
        # Unterdrückt "Chrome wird automatisiert gesteuert"-Banner
        args=[
            "--disable-blink-features=AutomationControlled",
            "--no-sandbox",
            "--disable-infobars",
            "--start-maximized",
        ],
        ignore_default_args=["--enable-automation"],
    )

    page = ctx.new_page()

    # Bot-Erkennung umgehen
    page.add_init_script("""
        Object.defineProperty(navigator, 'webdriver', { get: () => undefined });
        Object.defineProperty(navigator, 'plugins', { get: () => [1, 2, 3] });
        window.chrome = { runtime: {} };
    """)

    page.goto(URL, wait_until="domcontentloaded")

    go_back_to_january_2021(page)
    meetings_url = page.url

    print("\n>>> Klicke einen Meeting-Eintrag im Browser an.")
    print(">>> Script erkennt deinen Klick und lädt die ZIP automatisch.")
    print(">>> Browser schließen zum Beenden.\n")

    while True:
        try:
            result = download_meeting(page, meetings_url)

            if result is None:
                print("\n🛑 Browser geschlossen. Script beendet.")
                break
            elif result:
                print("\n>>> Klicke den nächsten Meeting-Eintrag an...")
            else:
                print("\n⚠️ Übersprungen. Klicke den nächsten Eintrag an...")

            try:
                page.goto(meetings_url, wait_until="domcontentloaded")
                page.wait_for_timeout(1500)
            except Exception:
                print("🛑 Konnte nicht zurücknavigieren.")
                break

        except Exception as e:
            if "closed" in str(e).lower():
                print("\n🛑 Browser geschlossen. Script beendet.")
                break
            print(f"\n❌ Fehler: {e}")
            try:
                page.goto(meetings_url, wait_until="domcontentloaded")
                page.wait_for_timeout(1000)
            except Exception:
                break

    ctx.close()
