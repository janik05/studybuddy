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

# Selektor für alle "Protokoll verteilt"-Einträge im Kalender
PROTOKOLL_SELEKTOR = (
    "xpath=//gss-half-year-view//table//tbody//tr//td//div"
    "[contains(., 'Protokoll verteilt')]"
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


def go_back_to_january_2021_and_count(page, max_clicks=80):
    """Navigiert zurück bis Januar 2021, gibt Anzahl der Klicks zurück."""
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


def go_back_n_times(page, n):
    """Drückt den Zurück-Button genau n mal."""
    if n <= 0:
        return
    prev_btn = page.locator(PREV_BTN_XPATH)
    prev_btn.wait_for(state="visible", timeout=60000)
    for _ in range(n):
        try:
            prev_btn.click(timeout=8000)
        except Exception:
            prev_btn.click(timeout=8000, force=True)
        page.wait_for_load_state("domcontentloaded")
        page.wait_for_timeout(500)
    print(f"📅 {n}x zurückgedrückt.")


def navigate_to_monat(page, klicks):
    """Lädt Hauptseite neu und navigiert mit klicks-mal zurück zum Zielmonat."""
    page.goto(URL, wait_until="domcontentloaded")
    page.wait_for_timeout(2500)
    prev_btn = page.locator(PREV_BTN_XPATH)
    prev_btn.wait_for(state="visible", timeout=30000)
    page.wait_for_timeout(1000)
    go_back_n_times(page, klicks)


def hole_protokoll_texte(page):
    """
    Gibt eine Liste der Texte aller Protokoll-Einträge zurück.
    So können wir nach dem Neuladen prüfen welche wir schon hatten.
    """
    eintraege = page.locator(PROTOKOLL_SELEKTOR)
    count = eintraege.count()
    texte = []
    for i in range(count):
        try:
            text = eintraege.nth(i).inner_text().strip()
            texte.append(text)
        except Exception:
            texte.append(f"Eintrag_{i}")
    return texte


def download_fuer_eintrag(page, index):
    """
    Klickt den index-ten Protokoll-Eintrag an und führt den Download durch.
    Gibt True/False/None zurück.
    """
    eintraege = page.locator(PROTOKOLL_SELEKTOR)
    count = eintraege.count()

    if index >= count:
        print(f"⚠️ Index {index} nicht mehr vorhanden (nur {count} Einträge).")
        return False

    eintrag = eintraege.nth(index)
    print(f"🖱️  Klicke Eintrag {index + 1} von {count}...")

    try:
        eintrag.scroll_into_view_if_needed(timeout=5000)
        eintrag.click(timeout=8000)
    except Exception:
        try:
            eintrag.click(timeout=8000, force=True)
        except Exception as e:
            print(f"❌ Klick fehlgeschlagen: {e}")
            return False

    # Warten bis Meeting-Detailansicht geladen
    agenda_selector = (
        "xpath=/html/body/gss-app-root/div/div/div[2]/div/div[2]/gss-meeting-agenda"
        "/div/div[1]/div/div[2]/gss-meeting-documents-menu/div/p-button/button/span"
    )
    try:
        page.locator(agenda_selector).wait_for(state="visible", timeout=15000)
    except Exception as e:
        if "closed" in str(e).lower():
            return None
        print(f"⚠️ Meeting-Ansicht nicht geladen: {e}")
        return False

    print("✅ Meeting geladen — starte Download...")
    page.wait_for_timeout(1000)

    # Dokument-Menü öffnen
    click_safe(
        page,
        "/html/body/gss-app-root/div/div/div[2]/div/div[2]/gss-meeting-agenda"
        "/div/div[1]/div/div[2]/gss-meeting-documents-menu/div/p-button/button/span",
    )
    page.wait_for_timeout(2000)

    # Menüpunkt 9
    click_safe(
        page,
        "/html/body/gss-app-root/div/div/div[2]/div/div[2]/gss-meeting-agenda"
        "/div/div[1]/div/div[2]/gss-meeting-documents-menu/p-menu/div/ul/li[9]",
    )
    page.wait_for_timeout(2000)

    # Download-Button
    click_safe(
        page,
        "/html/body/gss-app-root/div/div/div[2]/div/div[2]/gss-meeting-documents-map"
        "/div/div/div[1]/p-button/button/span[2]",
    )
    page.wait_for_timeout(2000)

    # ZIP herunterladen
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
            print(f"✅ ZIP gespeichert: {save_path}")
        else:
            print(f"⚠️ Kein ZIP-Format (Magic: {magic.hex()}): {save_path}")

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

    print(f"\n📌 Basis: {basis_klicks} Klicks bis Januar 2021")
    print(">>> Script läuft vollautomatisch.")
    print(">>> Alle 'Protokoll verteilt'-Einträge werden der Reihe nach heruntergeladen.")
    print(">>> Browser schließen zum Beenden.\n")

    download_count = 0
    # monat_offset: wie viele Klicks weniger als basis_klicks
    # 0 = Januar 2021, 1 = Februar 2021, usw.
    monat_offset = 0

    while True:
        try:
            aktuelle_klicks = basis_klicks - monat_offset
            print(f"\n🗓️  Monat-Offset: +{monat_offset} (Klicks: {aktuelle_klicks})")

            # Alle Protokoll-Einträge dieses Monats holen
            texte = hole_protokoll_texte(page)
            count = len(texte)
            print(f"📋 {count} Einträge mit 'Protokoll verteilt' in diesem Monat.")

            if count == 0:
                print("📭 Kein Eintrag — nächster Monat...")
                monat_offset += 1
                navigate_to_monat(page, basis_klicks - monat_offset)
                continue

            # Jeden Eintrag einzeln abarbeiten
            # Nach jedem Download: zurück zum Monat, nächsten Index anklicken
            for i in range(count):
                print(f"\n▶️  Verarbeite Eintrag {i + 1} von {count}...")

                result = download_fuer_eintrag(page, i)

                if result is None:
                    print("\n🛑 Browser geschlossen. Script beendet.")
                    ctx.close()
                    exit()

                if result:
                    download_count += 1
                    print(f"📊 Downloads gesamt: {download_count}")

                # Zurück zum aktuellen Monat für nächsten Eintrag
                navigate_to_monat(page, aktuelle_klicks)

            # Alle Einträge dieses Monats erledigt → nächster Monat
            print(f"\n✅ Monat +{monat_offset} komplett ({count} Downloads). Weiter zum nächsten Monat...")
            monat_offset += 1
            navigate_to_monat(page, basis_klicks - monat_offset)

        except Exception as e:
            if "closed" in str(e).lower():
                print("\n🛑 Browser geschlossen. Script beendet.")
                break
            print(f"\n❌ Fehler: {e}")
            try:
                navigate_to_monat(page, basis_klicks - monat_offset)
            except Exception:
                break

    ctx.close()
