# export_col_final.py
# AZtec 4.x — Export Word reports ONLY for sites whose name contains "col" (case-insensitive).
# Movement per row: focus-click -> Down once (no mouse wheel, no scrollbar arrow).
# Name read: Context Menu (Apps key) on CURRENT SELECTION; fallback: mouse right-click at Y (your old way).
# Uses your Save Report / Save button coordinates.

import time, os, re
import pyautogui as pag
import pyperclip, tkinter as tk
from tkinter import simpledialog

try:
    import keyboard
    HAVE_KBD = True
except:
    HAVE_KBD = False

# ============= Coordinates (buttons only) =============
SAVE_BTN_X, SAVE_BTN_Y = 1441, 105   # "Save Report"
DLG_SAVE_X, DLG_SAVE_Y = 1744, 1004  # "Save" in dialog

# ============= Settings =============
START_DELAY_SECONDS   = 3.0
DEFAULT_ROWS_TO_SCAN  = 20
ROW_HEIGHT_PX         = 24   # adjust ±1–2 if misaligned
WAIT_MENU_OPEN        = 0.50
WAIT_RENAME_FOCUS     = 1.60  # bump to ~2.0 if you ever see <blank>
COPY_SELECT_WAIT      = 0.15
COPY_CLIP_WAIT        = 0.25
FOCUS_DELAY           = 0.12
AFTER_DOWN_DELAY      = 0.14
PAUSE_BETWEEN_ROWS    = 0.10
WAIT_AFTER_SAVE       = 2.00

OUTPUT_DIR = r""   # set to skip already-saved, "" disables

pag.FAILSAFE = True
pag.PAUSE = 0.04

# ============= Helpers =============
def ask_int(title, prompt, default, minv=1, maxv=2000):
    root = tk.Tk(); root.withdraw()
    try:
        val = simpledialog.askinteger(title, prompt, initialvalue=default,
                                      minvalue=minv, maxvalue=maxv, parent=root)
    finally:
        try: root.destroy()
        except: pass
    return val or default

def capture_point(label, secs=3):
    print(f"\nHover your mouse over: {label}")
    for i in range(secs,0,-1):
        print(f"  capturing in {i}…"); time.sleep(1)
    pos = pag.position()
    print(f"-> Captured {label}: ({pos.x}, {pos.y})")
    return pos.x, pos.y

def normalize_name(txt):
    if not txt: return ""
    s = txt.strip().strip('"').strip()
    s = (s.replace('\u2013','-').replace('\u2014','-')
           .replace('\u2212','-').replace('\u00A0',' ')
           .replace('\u2009',' ').replace('\u200A',' ')
           .replace('\u202F',' '))
    return re.sub(r'\s+',' ',s)

def name_has_col_tag(txt):  # NEW: your tag rule
    return 'col' in (txt or '').lower()

def sanitize_filename(stem: str) -> str:
    stem = os.path.splitext(stem)[0]
    return re.sub(r'[\\/:*?"<>|]', '_', stem)

def already_saved(name):
    if not OUTPUT_DIR: return False
    safe = sanitize_filename(normalize_name(name)) + ".docx"
    return os.path.exists(os.path.join(OUTPUT_DIR, safe))

# ---- Old menu-driving for rename (mouse) ----
def right_click_menu_open_at(x,y):
    pag.moveTo(x,y,duration=0.05)
    pag.click(button='right')
    time.sleep(WAIT_MENU_OPEN)

def try_rename_from_open_menu():
    try: pyperclip.copy("")
    except: pass
    pag.press('home'); time.sleep(0.06)
    for idx in (0,1,2,3,4,5,6):
        if idx>0:
            for _ in range(idx):
                pag.press('down'); time.sleep(0.03)
        pag.press('enter'); time.sleep(WAIT_RENAME_FOCUS)
        pag.hotkey('ctrl','a'); time.sleep(COPY_SELECT_WAIT)
        pag.hotkey('ctrl','c'); time.sleep(COPY_CLIP_WAIT)
        name = (pyperclip.paste() or "").strip()
        pag.press('esc'); time.sleep(0.10)
        if name: return name
        pag.click(button='right'); time.sleep(WAIT_MENU_OPEN)
        pag.press('home'); time.sleep(0.06)
    return ""

def read_row_name_at(x,y):  # fallback only
    right_click_menu_open_at(x,y)
    return normalize_name(try_rename_from_open_menu())

# ---- NEW: Rename via Apps key on CURRENT SELECTION (selection-based, no Y guess) ----
def read_current_selection_name(fallback_x: int, fallback_y: int) -> str:
    try: pyperclip.copy("")
    except: pass

    # Apps key opens context menu on the highlighted row
    pag.press('apps'); time.sleep(WAIT_MENU_OPEN)
    pag.press('home'); time.sleep(0.05)

    for idx in (0,1,2,3,4,5,6):
        if idx>0:
            for _ in range(idx):
                pag.press('down'); time.sleep(0.03)
        pag.press('enter'); time.sleep(WAIT_RENAME_FOCUS)

        pag.hotkey('ctrl','a'); time.sleep(COPY_SELECT_WAIT)
        pag.hotkey('ctrl','c'); time.sleep(COPY_CLIP_WAIT)
        name = (pyperclip.paste() or "").strip()
        pag.press('esc'); time.sleep(0.10)
        if name:
            return normalize_name(name)

        # Try next menu item
        pag.press('apps'); time.sleep(WAIT_MENU_OPEN)
        pag.press('home'); time.sleep(0.05)

    # If Apps key path failed entirely, fall back to your old right-click-at-Y once
    return read_row_name_at(fallback_x, fallback_y)

def save_current_map():
    pag.moveTo(SAVE_BTN_X,SAVE_BTN_Y,duration=0.15)
    time.sleep(0.18); pag.click()
    time.sleep(0.18)
    pag.moveTo(DLG_SAVE_X,DLG_SAVE_Y,duration=0.15)
    pag.click(); time.sleep(WAIT_AFTER_SAVE)

# ============= Main =============
def main():
    rows_to_scan = ask_int("AZtec Export (col)","How many rows to scan?",DEFAULT_ROWS_TO_SCAN)
    visible_rows = ask_int("AZtec Export (col)","How many rows are visible at once?",10)

    print(f"\nYou have {START_DELAY_SECONDS:.1f}s to HOVER the first row.")
    time.sleep(START_DELAY_SECONDS)
    anchor_x, anchor_y = pag.position()
    print(f"[Anchor at x={anchor_x}, y={anchor_y}]")

    focus_x, focus_y = capture_point("SAFE FOCUS SPOT inside the list",secs=3)
    bottom_y = anchor_y + (visible_rows-1)*ROW_HEIGHT_PX  # kept for fallback only

    # Select Row 1 once
    pag.moveTo(anchor_x,anchor_y,duration=0.05); pag.click(); time.sleep(0.12)

    for i in range(1, rows_to_scan+1):
        if HAVE_KBD and keyboard.is_pressed('esc'):
            print("\n[STOP] ESC pressed."); break

        # 1) Ensure list focus (important after Save / dialogs)
        pag.moveTo(focus_x,focus_y,duration=0.02); pag.click(); time.sleep(FOCUS_DELAY)

        # 2) Compute fallback Y in case Apps key is ignored
        raw_y = anchor_y + (i-1)*ROW_HEIGHT_PX
        row_y_fallback = min(raw_y, bottom_y)

        # 3) Read CURRENT highlighted row name (Apps key preferred)
        nm = read_current_selection_name(anchor_x, row_y_fallback)

        # 4) Decide & Save
        if nm and name_has_col_tag(nm):
            if OUTPUT_DIR and already_saved(nm):
                print(f"[{i:02d}] {nm} → col | already saved")
            else:
                print(f"[{i:02d}] {nm} → col | saving…")
                save_current_map()
                # Re-focus list after Save (Word may steal focus)
                pag.moveTo(focus_x,focus_y,duration=0.02); pag.click(); time.sleep(FOCUS_DELAY)
        else:
            print(f"[{i:02d}] {nm if nm else '<blank>'} → skip")

        # 5) Move exactly ONE row forward (same as your second script)
        if i < rows_to_scan:
            pag.press('down'); time.sleep(AFTER_DOWN_DELAY)

        time.sleep(PAUSE_BETWEEN_ROWS)

    print("\nDone.")

if __name__=="__main__":
    main()
