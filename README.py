from pathlib import Path
from typing import Any, Dict, Optional
from datetime import date
import subprocess
import sys
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


def _open_powerpoint(url: str) -> bool:
    if not url:
        return False
    try:
        if sys.platform == "win32":
            # ms-powerpoint:ofe|u|<url> zwingt das Oeffnen in PowerPoint Desktop
            # (Open For Edit) und haelt die Datei mit SharePoint verbunden.
            os.startfile(f"ms-powerpoint:ofe|u|{url}")
        else:
            subprocess.Popen(["open", url])
        print(f"PowerPoint Desktop geoeffnet: {url}")
        return True
    except Exception as exc:
        print(f"Fehler beim Oeffnen mit PowerPoint: {exc}")
        return False


def _get_presentation(target_name: str):
    """Uebernimmt die laufende PowerPoint-Instanz und sucht die passende
    Praesentation (per Teilstring). Wartet, falls die Datei gerade erst aus
    SharePoint geoeffnet wird. Gibt (app, presentation) oder (None, None) zurueck."""
    import time
    import win32com.client as win32
    import pythoncom

    target = (target_name or "").lower()
    for _ in range(40):
        try:
            app = win32.GetActiveObject("PowerPoint.Application")
        except pythoncom.com_error:
            time.sleep(2)
            continue

        pres = None
        try:
            for p in app.Presentations:
                if target == "" or target in p.Name.lower():
                    pres = p
                    break
        except pythoncom.com_error:
            pres = None
        if pres is None:
            try:
                pres = app.ActivePresentation
            except pythoncom.com_error:
                pres = None

        if pres is not None:
            # Sicherstellen, dass die Folien zugaenglich sind (nicht geschuetzte Ansicht).
            try:
                _ = pres.Slides.Count
                return app, pres
            except pythoncom.com_error:
                pass
        time.sleep(2)
    return None, None


def _walk_shapes(shapes):
    """Liefert alle Shapes einer Sammlung, inkl. der Elemente in Gruppen (rekursiv)."""
    import pythoncom
    for sh in list(shapes):
        yield sh
        try:
            if sh.Type == 6:  # 6 = msoGroup
                for child in _walk_shapes(sh.GroupItems):
                    yield child
        except pythoncom.com_error:
            pass


def _shape_link_format(sh):
    """Gibt das LinkFormat zurueck, wenn das Shape tatsaechlich verknuepft ist
    (unabhaengig vom Shape-Type), sonst None."""
    import pythoncom
    try:
        lf = sh.LinkFormat
        _ = lf.SourceFullName  # Zugriff bestaetigt eine echte Verknuepfung
        return lf
    except pythoncom.com_error:
        return None
    except Exception:
        return None


def _remove_all_links(slide) -> int:
    """Entfernt auf einer Folie ALLE Verknuepfungen: OLE-/Bild-Links werden
    aufgebrochen (Inhalt bleibt statisch) UND saemtliche Hyperlinks geloescht."""
    import pythoncom

    removed = 0

    # 1. Verknuepfungen aufbrechen - Erkennung ueber LinkFormat (robust, auch in Gruppen).
    for sh in _walk_shapes(slide.Shapes):
        lf = _shape_link_format(sh)
        if lf is not None:
            try:
                lf.BreakLink()
                removed += 1
            except pythoncom.com_error as exc:
                print(f"WARN: Verknuepfung konnte nicht aufgebrochen werden: {exc}")

    # 2. Alle Hyperlinks der Folie entfernen (rueckwaerts, da Delete die Collection aendert).
    try:
        links = slide.Hyperlinks
        for i in range(links.Count, 0, -1):
            try:
                links.Item(i).Delete()
                removed += 1
            except pythoncom.com_error:
                pass
    except pythoncom.com_error:
        pass

    return removed


def _has_thinkcell(slide) -> bool:
    """True, wenn die Folie noch aktive think-cell-Elemente enthaelt.
    Jede think-cell-Folie besitzt einen Anker-Shape 'think-cell data - do not delete'
    bzw. ein OLE-Objekt mit ProgID 'TCLayout...'."""
    for sh in _walk_shapes(slide.Shapes):
        try:
            if "think-cell" in (sh.Name or "").lower():
                return True
        except Exception:
            pass
        try:
            if "tclayout" in str(sh.OLEFormat.ProgID).lower():
                return True
        except Exception:
            pass
    return False


def _freeze_slide_as_picture(pres, slide) -> bool:
    """Friert eine Folie als statisches Bild ein. Da think-cell das Loeschen
    einzelner Chart-/Anker-Shapes blockiert ('do not delete'), wird stattdessen
    eine NEUE leere Folie direkt hinter der Originalfolie erzeugt, dort nur das
    Standbild (Enhanced Metafile) eingefuegt und die komplette Originalfolie
    geloescht. Eine ganze Folie kann think-cell nicht schuetzen -> keine Links mehr.
    Die neue Folie liegt danach an derselben Position und in derselben Sektion."""
    import pythoncom
    import time

    try:
        idx = slide.SlideIndex
    except pythoncom.com_error:
        return False

    # Bounding-Box der Original-Shapes fuer exakte Positionierung merken.
    try:
        rng = slide.Shapes.Range()
        box = (rng.Left, rng.Top, rng.Width, rng.Height)
    except pythoncom.com_error:
        box = None

    # 1. Neue LEERE Folie direkt hinter der Originalfolie (gleiche Sektion).
    try:
        new_slide = pres.Slides.Add(idx + 1, 12)  # 12 = ppLayoutBlank
    except pythoncom.com_error as exc:
        print(f"WARN: Neue Folie fuer Einfrieren konnte nicht erstellt werden: {exc}")
        return False

    # 2. Alle Original-Shapes als Bild in die Zwischenablage kopieren.
    try:
        slide.Shapes.Range().Copy()
    except pythoncom.com_error as exc:
        print(f"WARN: Kopieren fuer Einfrieren fehlgeschlagen: {exc}")
        try:
            new_slide.Delete()
        except pythoncom.com_error:
            pass
        return False
    time.sleep(0.4)

    # 3. Als Bild in die neue Folie einfuegen.
    pic = None
    for data_type in (2, 7):  # 2 = ppPasteEnhancedMetafile, 7 = ppPastePNG (Fallback)
        try:
            pasted = new_slide.Shapes.PasteSpecial(DataType=data_type)
            pic = pasted.Item(1)
            break
        except pythoncom.com_error:
            pic = None
    if pic is None:
        print("WARN: Einfuegen als Bild fehlgeschlagen - Folie bleibt unveraendert.")
        try:
            new_slide.Delete()
        except pythoncom.com_error:
            pass
        return False

    try:
        pic.Name = "__FROZEN_SNAPSHOT__"
    except pythoncom.com_error:
        pass
    if box is not None:
        try:
            pic.Left, pic.Top, pic.Width, pic.Height = box
        except pythoncom.com_error:
            pass

    # 4. Komplette Originalfolie loeschen (think-cell kann eine ganze Folie nicht
    #    schuetzen). Die neue Standbild-Folie rueckt auf die Originalposition.
    try:
        slide.Delete()
    except pythoncom.com_error as exc:
        print(f"WARN: Originalfolie konnte nicht geloescht werden: {exc}")
        return False
    return True


def _update_ppt_via_com(target_name: str, section_name: str, current_section: str, today: str) -> bool:
    try:
        import win32com.client  # noqa: F401
        import pythoncom
    except ImportError:
        print("pywin32 (win32com) ist nicht installiert - PowerPoint-Automatisierung nicht moeglich.")
        return False
    import re

    app, pres = _get_presentation(target_name)
    if app is None:
        print("ERROR: Keine laufende PowerPoint-Instanz gefunden.")
        return False
    if pres is None:
        print("ERROR: Keine passende Praesentation gefunden.")
        return False

    try:
        app.DisplayAlerts = 1  # 1 = ppAlertsNone
    except pythoncom.com_error:
        pass

    if pres.Slides.Count < 1:
        print("ERROR: Praesentation enthaelt keine Folien.")
        return False

    # 1. Erste Folie duplizieren. Das Duplikat liegt danach direkt hinter Folie 1.
    first = pres.Slides.Item(1)
    dup_range = first.Duplicate()
    try:
        dup_index = dup_range.SlideIndex
    except pythoncom.com_error:
        dup_index = 2
    dup = pres.Slides.Item(dup_index)
    print(f"Folie 1 dupliziert -> neues Duplikat an Position {dup_index}.")

    # 2. Abschnitt 'Previous Status' suchen.
    sp = pres.SectionProperties
    if sp.Count < 1:
        print("ERROR: Praesentation enthaelt keine Abschnitte.")
        return False
    sec_index = 0
    for s in range(1, sp.Count + 1):
        if sp.Name(s) == section_name:
            sec_index = s
            break
    if sec_index == 0:
        print(f"ERROR: Abschnitt '{section_name}' nicht gefunden.")
        return False

    # Zielposition: Duplikat ans ENDE des Zielabschnitts verschieben, damit es
    # zuverlaessig in diesem Abschnitt landet (Abschnittsgrenzen sind heikel).
    sec_first = sp.FirstSlide(sec_index)
    sec_count = sp.SlidesCount(sec_index)
    if sec_count >= 1:
        target_pos = sec_first + sec_count - 1
    else:
        target_pos = sec_first
        if target_pos < 1:
            target_pos = 1
        if target_pos > pres.Slides.Count:
            target_pos = pres.Slides.Count

    dup.MoveTo(target_pos)
    dup_idx = dup.SlideIndex

    # Kontrolle: in welchem Abschnitt liegt das Duplikat jetzt?
    landed = ""
    for s in range(1, sp.Count + 1):
        f = sp.FirstSlide(s)
        c = sp.SlidesCount(s)
        if f <= dup_idx < (f + c):
            landed = sp.Name(s)
            break
    print(f"Duplikat verschoben -> Position {dup_idx}, Abschnitt: '{landed}'.")

    # 3. Auf ALLEN Folien AUSSER Folie 1 die think-cell-Verknuepfungen dauerhaft
    #    einfrieren: die Folie wird als statisches Bild (Enhanced Metafile) fixiert.
    #    So bleiben keine aktiven Datenverknuepfungen (Previous Status = eingefroren).
    #    Bereits statische Folien werden uebersprungen (idempotent).
    frozen = 0
    skipped = 0
    for pos in range(2, pres.Slides.Count + 1):
        try:
            sld = pres.Slides.Item(pos)
        except pythoncom.com_error:
            continue
        if not _has_thinkcell(sld):
            # Bereits statisch: nur klassische OLE-Links/Hyperlinks entfernen.
            _remove_all_links(sld)
            skipped += 1
            continue
        # think-cell vorhanden -> Folie durch statisches Standbild ersetzen.
        if _freeze_slide_as_picture(pres, sld):
            frozen += 1
    print(f"Eingefroren (ausser Folie 1): {frozen} Folien als Standbild, {skipped} bereits statisch.")

    # ==================== TEIL B: Current State ====================
    cur_index = 0
    for s in range(1, sp.Count + 1):
        if sp.Name(s) == current_section:
            cur_index = s
            break
    if cur_index == 0:
        print(f"WARN: Abschnitt '{current_section}' nicht gefunden - Teil B uebersprungen.")
    else:
        cur_first = sp.FirstSlide(cur_index)
        cur_count = sp.SlidesCount(cur_index)
        if cur_count < 1:
            print(f"WARN: Abschnitt '{current_section}' enthaelt keine Folien - Teil B uebersprungen.")
        else:
            cur_slide = pres.Slides.Item(cur_first)
            print(f"Current State: Folie {cur_first} wird bearbeitet.")

            # 5. Manuell dargestelltes Datum 'oben links' aktualisieren: unter allen
            #    Textfeldern mit Datum (TT.MM.JJJJ) das am weitesten oben-links.
            date_pattern = re.compile(r"\b\d{1,2}\.\d{1,2}\.\d{4}\b")
            date_shape = None
            best_score = float("inf")
            for sh in cur_slide.Shapes:
                has_text = False
                try:
                    has_text = sh.HasTextFrame == -1 and sh.TextFrame.HasText == -1
                except pythoncom.com_error:
                    has_text = False
                if has_text:
                    try:
                        txt = str(sh.TextFrame.TextRange.Text)
                        if date_pattern.search(txt):
                            score = float(sh.Top) + float(sh.Left)
                            if score < best_score:
                                best_score = score
                                date_shape = sh
                    except pythoncom.com_error:
                        pass
            if date_shape is not None:
                txt = str(date_shape.TextFrame.TextRange.Text)
                new_txt = date_pattern.sub(today, txt)
                date_shape.TextFrame.TextRange.Text = new_txt
                print(f"Current State: Datum oben-links aktualisiert in '{date_shape.Name}' -> {today}")
            else:
                print("WARN: Kein Datumsfeld (TT.MM.JJJJ) auf der Current-Status-Folie gefunden.")

            # 6. Folie 1 bleibt live: die think-cell-Elemente bleiben mit dem
            #    Quell-Excel verbunden. think-cell haelt/aktualisiert diese, solange
            #    die Quell-Arbeitsmappe geoeffnet ist (PowerPoint-COM erfasst sie nicht).
            tc_count = 0
            for sh in _walk_shapes(cur_slide.Shapes):
                try:
                    if "think-cell" in (sh.Name or "").lower():
                        tc_count += 1
                except Exception:
                    pass
            print(f"Current State: Folie 1 bleibt live ({tc_count} think-cell-Anker verbunden).")

    # 7. Klassische (Nicht-think-cell) Datenverknuepfungen best-effort aktualisieren.
    #    think-cell-Links werden hier nicht erfasst und aktualisieren sich ueber
    #    think-cell selbst (solange die Quell-Excel geoeffnet ist).
    try:
        pres.UpdateLinks()
    except pythoncom.com_error:
        pass

    try:
        pres.Save()
    except pythoncom.com_error as exc:
        print(f"ERROR: Speichern fehlgeschlagen: {exc}")
        return False
    print("OK: Previous Status eingefroren (Standbild), Folie 1 bleibt live, Datum aktualisiert, gespeichert.")
    return True



def _reopen_for_thinkcell_update(target_name: str, url: str) -> bool:
    """Schliesst die (bereits gespeicherte) Praesentation und oeffnet sie ueber
    die SharePoint-URL neu. think-cell aktualisiert die Excel-Datenverknuepfungen
    automatisch beim Laden, sofern die Quell-Arbeitsmappe geoeffnet/erreichbar ist.
    Danach wird erneut gespeichert, damit die aktualisierten Werte persistiert sind."""
    import time
    import pythoncom

    if not url:
        print("WARN: Keine PPT-URL fuer Neu-Oeffnen - think-cell-Update uebersprungen.")
        return False

    app, pres = _get_presentation(target_name)
    if pres is None:
        print("WARN: Fuer think-cell-Update keine Praesentation gefunden.")
        return False

    try:
        app.DisplayAlerts = 1  # ppAlertsNone -> keine Dialoge beim Schliessen
    except pythoncom.com_error:
        pass

    # Sicherstellen, dass alles gespeichert ist, dann schliessen.
    try:
        pres.Save()
    except pythoncom.com_error:
        pass
    try:
        pres.Close()
        print("think-cell-Update: Praesentation geschlossen.")
    except pythoncom.com_error as exc:
        print(f"WARN: Schliessen fehlgeschlagen: {exc}")
        return False

    time.sleep(2)

    # Neu oeffnen (Open For Edit) -> think-cell aktualisiert beim Laden.
    _open_powerpoint(url)
    app2, pres2 = _get_presentation(target_name)
    if pres2 is None:
        print("WARN: Neu geoeffnete Praesentation nicht gefunden - Update evtl. unvollstaendig.")
        return False

    # think-cell braucht nach dem Laden einen Moment, um die Links zu aktualisieren.
    time.sleep(12)

    try:
        app2.DisplayAlerts = 1
    except pythoncom.com_error:
        pass
    try:
        pres2.Save()
        print("OK: think-cell hat Folie 1 beim Neu-Oeffnen aktualisiert und gespeichert.")
        return True
    except pythoncom.com_error as exc:
        print(f"WARN: Speichern nach think-cell-Update fehlgeschlagen: {exc}")
        return False


def _find_ppt_frame_hwnd():
    """Handle des sichtbaren PowerPoint-Hauptfensters (PPTFrameClass)."""
    import win32gui
    found = []

    def cb(h, _):
        try:
            if win32gui.IsWindowVisible(h) and win32gui.GetClassName(h) == "PPTFrameClass":
                found.append(h)
        except Exception:
            pass

    win32gui.EnumWindows(cb, None)
    return found[0] if found else None


def _find_datalinks_dialog_rect():
    """Rechteck (left, top, right, bottom) des think-cell-Dialogs
    'Datenverknuepfungen'. Prozess muss DPI-UNAWARE bleiben (logische Koordinaten)."""
    import win32gui
    res = []

    def cb(h, _):
        try:
            if (win32gui.IsWindowVisible(h)
                    and win32gui.GetWindowText(h) == "Datenverkn\u00fcpfungen"
                    and win32gui.GetClassName(h).startswith("ATL:")):
                res.append(win32gui.GetWindowRect(h))
        except Exception:
            pass

    win32gui.EnumWindows(cb, None)
    return res[0] if res else None


def _find_datalinks_dialog_hwnd():
    """Handle des think-cell-Dialogs 'Datenverknuepfungen'."""
    import win32gui
    res = []

    def cb(h, _):
        try:
            if (win32gui.IsWindowVisible(h)
                    and win32gui.GetWindowText(h) == "Datenverkn\u00fcpfungen"
                    and win32gui.GetClassName(h).startswith("ATL:")):
                res.append(h)
        except Exception:
            pass

    win32gui.EnumWindows(cb, None)
    return res[0] if res else None


def _get_dpi_scale(hwnd) -> float:
    """DPI-Skalierungsfaktor relativ zur Kalibrierung der Klick-Offsets (150 % = 144 DPI).
    Faellt auf 1.0 zurueck, wenn GetDpiForWindow nicht verfuegbar ist (aeltere Windows-Version)."""
    import ctypes
    try:
        dpi = ctypes.windll.user32.GetDpiForWindow(hwnd)
        if dpi:
            return dpi / 144.0
    except Exception:
        pass
    return 1.0


def _update_thinkcell_via_dialog(target_name: str) -> bool:
    """Bildet exakt den manuellen think-cell-Ablauf nach:
    Einfuegen-Tab -> 'Datenverknuepfungen...' oeffnen -> 'Alle verknuepften
    Elemente auswaehlen' -> Update-Button (Kreispfeil) klicken -> 'Schliessen'.

    Oeffnen/Auswaehlen/Schliessen laufen ueber UIA (InvokePattern, funktioniert
    unabhaengig von Fokus/Monitor). Der Update-Button ist eigen-gezeichnet und
    nicht UIA-ansprechbar -> Koordinaten-Klick relativ zum Dialog-Rechteck,
    DPI-skaliert und mit explizitem Foreground-Fenster.
    """
    import time
    import win32api
    import win32con
    import win32gui

    try:
        from pywinauto.application import Application
    except ImportError:
        print("WARN: pywinauto nicht installiert - Dialog-Update nicht moeglich.")
        return False

    # Offsets des Update-Buttons (Kreispfeil) vom Dialogrand, in LOGISCHEN Pixeln
    # (gemessen bei 150 %, DPI-unaware). Rechtsbuendige Mini-Toolbar, erste Quellzeile.
    # Werden zur Laufzeit per _get_dpi_scale() an die tatsaechliche Skalierung angepasst.
    OFFSET_FROM_RIGHT = 69
    OFFSET_FROM_TOP = 176

    ppt_hwnd = _find_ppt_frame_hwnd()
    if ppt_hwnd is None:
        print("WARN: PowerPoint-Fenster nicht gefunden - Dialog-Update abgebrochen.")
        return False

    try:
        app = Application(backend="uia").connect(handle=ppt_hwnd)
        main = app.window(handle=ppt_hwnd)
    except Exception as exc:
        print(f"WARN: UIA-Verbindung zu PowerPoint fehlgeschlagen: {exc}")
        return False

    # 1) Einfuegen-Tab aktivieren, damit der Ribbon-Button verfuegbar ist.
    try:
        main.child_window(title="Einf\u00fcgen", control_type="TabItem").wrapper_object().select()
        time.sleep(1)
    except Exception as exc:
        print(f"Hinweis: Einfuegen-Tab nicht aktiviert ({exc}) - versuche trotzdem weiter.")

    # 2) Ribbon-Button 'Datenverknuepfungen...' aufrufen -> Dialog oeffnet.
    try:
        main.child_window(title_re="Datenverkn.*", control_type="Button").wrapper_object().invoke()
    except Exception as exc:
        print(f"WARN: 'Datenverknuepfungen' konnte nicht geoeffnet werden: {exc}")
        return False

    # Auf den Dialog warten.
    rect = None
    for _ in range(20):
        time.sleep(0.5)
        rect = _find_datalinks_dialog_rect()
        if rect is not None:
            break
    if rect is None:
        print("WARN: Datenverknuepfungen-Dialog ist nicht erschienen.")
        return False

    dlg_hwnd = _find_datalinks_dialog_hwnd()
    if dlg_hwnd is None:
        print("WARN: Handle des Datenverknuepfungen-Dialogs nicht gefunden.")
        return False

    # 3) 'Alle verknuepften Elemente auswaehlen' aufrufen (UIA).
    try:
        dlg_app = Application(backend="uia").connect(handle=dlg_hwnd)
        dlg = dlg_app.window(handle=dlg_hwnd)
        dlg.child_window(title_re="Alle verkn.*", control_type="Button").wrapper_object().invoke()
        time.sleep(1.5)
    except Exception as exc:
        print(f"WARN: 'Alle verknuepften Elemente auswaehlen' fehlgeschlagen: {exc}")
        return False

    # 4) Dialog in den Vordergrund holen, damit der synthetische Klick auch ankommt
    #    (ohne Fokus/Vordergrund verpuffen SetCursorPos/mouse_event oft wirkungslos).
    try:
        win32gui.SetForegroundWindow(dlg_hwnd)
        time.sleep(0.3)
    except Exception as exc:
        print(f"Hinweis: Dialog konnte nicht in den Vordergrund geholt werden: {exc}")

    # 5) Update-Button (Kreispfeil) per Koordinaten-Klick. Rect frisch lesen,
    #    Offsets an die tatsaechliche DPI-Skalierung anpassen (Kalibrierung war 150 %).
    rect = _find_datalinks_dialog_rect() or rect
    left, top, right, bottom = rect
    scale = _get_dpi_scale(dlg_hwnd)
    x = right - int(OFFSET_FROM_RIGHT * scale)
    y = top + int(OFFSET_FROM_TOP * scale)
    print(f"DEBUG: dpi_scale={scale:.2f}, rect={rect}, click=({x},{y})")
    print(f"think-cell-Update: Klick auf Aktualisieren bei ({x},{y}) - Dialog {rect}.")
    try:
        win32api.SetCursorPos((x, y))
        time.sleep(0.3)
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
        time.sleep(0.05)
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
    except Exception as exc:
        print(f"WARN: Klick auf Update-Button fehlgeschlagen: {exc}")
        return False

    # think-cell braucht einen Moment fuer die Aktualisierung.
    time.sleep(6)

    # 6) Dialog schliessen (UIA).
    try:
        dlg.child_window(title="Schlie\u00dfen", control_type="Button").wrapper_object().invoke()
    except Exception:
        # Fallback: Fenster direkt schliessen.
        try:
            win32gui.PostMessage(dlg_hwnd, win32con.WM_CLOSE, 0, 0)
        except Exception:
            pass

    time.sleep(1)

    # 7) Speichern.
    try:
        import pythoncom
        _app, pres = _get_presentation(target_name)
        if pres is not None:
            try:
                _app.DisplayAlerts = 1
            except pythoncom.com_error:
                pass
            pres.Save()
            print("OK: think-cell aktualisiert (Dialog: alle auswaehlen + aktualisieren) und gespeichert.")
            return True
    except Exception as exc:
        print(f"WARN: Speichern nach Dialog-Update fehlgeschlagen: {exc}")

    return True


def _build_ppt_url(config: Dict[str, Any]) -> str:
    ppt_url = str(config.get("ppt_url", "")).strip()
    if ppt_url:
        return ppt_url

    site_url = str(config.get("site_url", "")).strip().rstrip("/")
    folder_relative_url = str(config.get("folder_server_relative_url", "")).strip().rstrip("/")
    ppt_name = str(config.get("ppt_name", "")).strip()
    if not site_url or not folder_relative_url or not ppt_name:
        return ""

    from urllib.parse import urlparse
    parsed = urlparse(site_url)
    file_name = ppt_name if ppt_name.lower().endswith(".pptx") else f"{ppt_name}.pptx"
    return f"{parsed.scheme}://{parsed.netloc}{folder_relative_url}/{file_name}?web=1"


def update_powerpoint(config_path: Optional[str] = None) -> bool:
    config = _load_config(config_path)
    ppt_name = str(config.get("ppt_name", "")).strip()
    section_name = str(config.get("ppt_previous_section", "Previous Status")).strip() or "Previous Status"
    current_section = str(config.get("ppt_current_section", "Current State")).strip() or "Current State"
    today = date.today().strftime("%d.%m.%Y")

    url = _build_ppt_url(config)
    if url:
        _open_powerpoint(url)

    ok = _update_ppt_via_com(ppt_name, section_name, current_section, today)

    # Nach dem Freeze: think-cell-Datenverknuepfungen auf Folie 1 aktualisieren -
    # exakt der manuelle Ablauf: Datenverknuepfungen oeffnen -> alle auswaehlen ->
    # aktualisieren -> schliessen. Faellt bei Problemen auf Neu-Oeffnen zurueck.
    if ok:
        if not _update_thinkcell_via_dialog(ppt_name):
            print("Hinweis: Dialog-Update nicht erfolgreich - versuche Neu-Oeffnen als Fallback.")
            _reopen_for_thinkcell_update(ppt_name, url)

    return ok


if __name__ == "__main__":
    update_powerpoint()
