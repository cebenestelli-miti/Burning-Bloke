"""
BB Schedule — admin UI + fullscreen display. Paths are relative to the executable directory.
"""
from __future__ import annotations

import html
import json
import os
import shutil
import subprocess
import sys
from collections import defaultdict
import time
from datetime import date, datetime, timedelta
from typing import Any, Callable, Dict, List, Optional, Tuple

import tkinter as tk
from tkinter import filedialog, font as tkfont, messagebox

import customtkinter as ctk

try:
    from tkcalendar import DateEntry
except ImportError:
    DateEntry = None  # type: ignore

# --- Theme & constants -------------------------------------------------
# Admin UI stack: CustomTkinter 5, set_appearance_mode("dark") +
# set_default_color_theme("dark-blue"); transparent CTkFrames; subtitle/hint
# labels use text_color=("gray30","gray65"); main title CTkFont 18 bold.
# Non-CTk: tkcalendar DateEntry (picker button #1f538d / #DCE4EE); time fields
# are tk.Spinbox with dark-gray styling to match dark-blue frames; fullscreen
# DisplayWindow stays tk.Canvas + tk.Toplevel.
SUBTLE_TEXT_COLOR = ("gray30", "gray65")
# CTkLabel.configure() rejects text_color=None; use explicit theme-like default.
LABEL_TEXT_COLOR = ("gray14", "gray90")

DISPLAY_CLOCK_FONT = ("Segoe UI", 48, "bold")
DISPLAY_MSG_FONT = ("Segoe UI", 28)

# Activity row image filename display (full path kept in config; label is truncated).
_IMAGE_NAME_DISPLAY_MAX = 24

DEFAULT_ACTIVITY_TYPES = [
    "Breakfast & Coffee",
    "Lunch & Camp Setup",
    "Lunch",
    "Dinner, Campfire & Drinks",
    "Dinner & Bloke Burning",
    "Meditation",
    "Swimming & Kayaking",
    "Touch Football",
    "Shooting",
    "Rally Driving",
    "Sauna & Ice Bath",
    "Time Capsule",
    "Free Time",
]

PROTECTED_ACTIVITY_TYPES: frozenset[str] = frozenset(DEFAULT_ACTIVITY_TYPES)

DEFAULT_LOCATION_TYPES = [
    "The Barn",
    "Fire",
    "Sauna",
    "Watch Tower",
    "Bayside",
    "Dam Forest",
    "Touch Football",
    "Shooting",
]

PROTECTED_LOCATION_TYPES: frozenset[str] = frozenset(DEFAULT_LOCATION_TYPES)

CONFIG_NAME = "config.json"
HTML_EXPORT_NAME = "event_status.html"
ICON_ICO = "app_icon.ico"
BACKGROUND_IMAGE_NAME = "Background.jpg"
MAP_IMAGE_CANDIDATES = [
    os.path.join("Map", "BB map3.png"),
    "site_map_art_illustrated2.png",
    "site_map_art_illustrated.png",
    "site_map_art_display.jpg",
    "site_map_art.jpg",
]

# --- Paths -------------------------------------------------------------


def app_directory() -> str:
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


def icon_ico_path() -> Optional[str]:
    if getattr(sys, "frozen", False):
        meipass = getattr(sys, "_MEIPASS", None)
        if meipass:
            p = os.path.join(meipass, ICON_ICO)
            if os.path.isfile(p):
                return p
        p = os.path.join(app_directory(), ICON_ICO)
        return p if os.path.isfile(p) else None
    base = os.path.dirname(os.path.abspath(__file__))
    p = os.path.join(base, ICON_ICO)
    return p if os.path.isfile(p) else None


def apply_window_icon(win: tk.Misc) -> None:
    """Title bar / taskbar icon (ICO preferred; PhotoImage fallback for stubborn hosts)."""
    p = icon_ico_path()
    if not p:
        return
    p = os.path.abspath(os.path.normpath(p))
    if not os.path.isfile(p):
        return

    def with_bitmap() -> None:
        try:
            win.iconbitmap(p)
        except tk.TclError:
            pass

    with_bitmap()
    try:
        win.after(1, with_bitmap)
    except Exception:
        pass

    try:
        from PIL import Image, ImageTk

        im = Image.open(p)
        if im.width != 32 or im.height != 32:
            try:
                resample = Image.Resampling.LANCZOS  # type: ignore[attr-defined]
            except AttributeError:
                resample = Image.LANCZOS
            im = im.resize((32, 32), resample)
        photo = ImageTk.PhotoImage(im)
        try:
            win.iconphoto(True, photo)
        except tk.TclError:
            pass
        setattr(win, "_bb_icon_photo_ref", photo)
    except Exception:
        pass


def _center_toplevel_on_parent(
    win: Any,
    parent: Optional[tk.Misc],
    min_width: int,
    min_height: int,
) -> None:
    win.update_idletasks()
    w = max(int(win.winfo_reqwidth()), int(win.winfo_width()), min_width)
    h = max(int(win.winfo_reqheight()), int(win.winfo_height()), min_height)
    if parent is not None:
        try:
            px = int(parent.winfo_rootx())
            py = int(parent.winfo_rooty())
            pw = int(parent.winfo_width())
            ph = int(parent.winfo_height())
            x = px + max((pw - w) // 2, 0)
            y = py + max((ph - h) // 2, 0)
        except Exception:
            x = (win.winfo_screenwidth() - w) // 2
            y = (win.winfo_screenheight() - h) // 2
    else:
        x = (win.winfo_screenwidth() - w) // 2
        y = (win.winfo_screenheight() - h) // 2
    win.geometry(f"{w}x{h}+{x}+{y}")


def themed_message(
    parent: Optional[tk.Misc],
    title: str,
    message: str,
    kind: str = "info",
    detail: Optional[str] = None,
) -> None:
    """Themed modal dialog (replaces native messagebox)."""
    try:
        win = ctk.CTkToplevel(parent) if parent is not None else ctk.CTkToplevel()
    except Exception:
        win = ctk.CTkToplevel()
    win.title(title)
    win.resizable(False, False)
    apply_window_icon(win)

    try:
        win.transient(parent)  # type: ignore[arg-type]
    except Exception:
        pass
    try:
        win.grab_set()
    except Exception:
        pass
    try:
        win.lift()
    except Exception:
        pass

    outer_pad = 20
    inner_w = 500
    wrap = inner_w - outer_pad * 2

    card = ctk.CTkFrame(win, corner_radius=10)
    card.pack(fill="both", expand=True, padx=outer_pad, pady=outer_pad)

    kind_norm = (kind or "info").lower()
    if kind_norm == "error":
        accent = "#ff7875"
    elif kind_norm == "warning":
        accent = "#ffc069"
    elif kind_norm in ("success", "ok"):
        accent = "#73d13d"
    else:
        accent = "#40a9ff"

    ctk.CTkLabel(
        card,
        text=title,
        font=ctk.CTkFont(size=19, weight="bold"),
        text_color=accent,
    ).pack(anchor="w", padx=outer_pad, pady=(outer_pad, 6))

    msg_lbl = ctk.CTkLabel(
        card,
        text=message,
        justify="left",
        anchor="w",
        wraplength=wrap,
    )
    msg_lbl.pack(fill="x", padx=outer_pad, pady=(0, 6))

    if detail:
        ctk.CTkLabel(
            card,
            text=detail,
            text_color=SUBTLE_TEXT_COLOR,
            justify="left",
            anchor="w",
            wraplength=wrap,
            font=ctk.CTkFont(size=13),
        ).pack(fill="x", padx=outer_pad, pady=(0, outer_pad))

    btns = ctk.CTkFrame(card, fg_color="transparent")
    btns.pack(fill="x", padx=outer_pad, pady=(8, outer_pad))
    btns.grid_columnconfigure(0, weight=1)
    btns.grid_columnconfigure(2, weight=1)

    def close() -> None:
        try:
            win.grab_release()
        except Exception:
            pass
        win.destroy()

    ok = ctk.CTkButton(btns, text="OK", width=120, command=close)
    ok.grid(row=0, column=1, pady=(4, 0))

    win.bind("<Escape>", lambda _e: close())
    win.bind("<Return>", lambda _e: close())

    win.update_idletasks()
    _center_toplevel_on_parent(win, parent, min_width=inner_w + outer_pad * 2, min_height=140)
    win.update_idletasks()
    _center_toplevel_on_parent(win, parent, min_width=inner_w + outer_pad * 2, min_height=140)
    try:
        ok.focus_set()
    except Exception:
        pass
    try:
        win.wait_window()
    except Exception:
        pass


def config_path() -> str:
    return os.path.join(app_directory(), CONFIG_NAME)


def find_git_repo_root(start_dir: str) -> Optional[str]:
    """Find nearest parent directory containing .git."""
    cur = os.path.abspath(start_dir)
    while True:
        if os.path.isdir(os.path.join(cur, ".git")):
            return cur
        parent = os.path.dirname(cur)
        if parent == cur:
            return None
        cur = parent


def sort_activity_types(vals: List[str]) -> List[str]:
    return sorted([str(v).strip() for v in (vals or []) if str(v).strip()], key=lambda s: s.casefold())


def normalize_activity_types_from_config(raw: Any) -> List[str]:
    if not isinstance(raw, list):
        return sort_activity_types(list(DEFAULT_ACTIVITY_TYPES))
    result = list(DEFAULT_ACTIVITY_TYPES)
    seen = set(result)
    for x in raw:
        s = str(x).strip()
        if s and s not in seen:
            result.append(s)
            seen.add(s)
    return sort_activity_types(result)


def normalize_location_types_from_config(raw: Any) -> List[str]:
    if not isinstance(raw, list):
        return sort_activity_types(list(DEFAULT_LOCATION_TYPES))
    result = list(DEFAULT_LOCATION_TYPES)
    seen = set(result)
    for x in raw:
        s = str(x).strip()
        if s and s not in seen:
            result.append(s)
            seen.add(s)
    return sort_activity_types(result)


# --- Time helpers ------------------------------------------------------


def parse_hhmm(s: str) -> Optional[Tuple[int, int]]:
    s = (s or "").strip()
    if not s:
        return None
    parts = s.replace(".", ":").split(":")
    if len(parts) != 2:
        return None
    try:
        h, m = int(parts[0]), int(parts[1])
    except ValueError:
        return None
    if not (0 <= h <= 23 and 0 <= m <= 59):
        return None
    return h, m


def hhmm(h: int, m: int) -> str:
    return f"{h:02d}:{m:02d}"


def now_seconds_since_midnight() -> int:
    n = datetime.now()
    return n.hour * 3600 + n.minute * 60 + n.second


def time_tuple_to_seconds(h: int, m: int, sec: int = 0) -> int:
    return h * 3600 + m * 60 + sec


def fmt_ampm(h: int, m: int) -> str:
    hh = h % 12
    if hh == 0:
        hh = 12
    suffix = "AM" if h < 12 else "PM"
    return f"{hh}:{m:02d} {suffix}"


def schedule_slot_spec(sh: int, sm: int, eh: int, em: int, g1: str, g2: str) -> Optional[Tuple[str, List[str]]]:
    """One timeslot for fullscreen overlay: (time_prefix, body_lines). Group 2 line under Group 1 (pixel-aligned when drawn)."""
    g1 = (g1 or "").strip()
    g2 = (g2 or "").strip()
    if not g1 and not g2:
        return None
    time_prefix = f"{fmt_ampm(sh, sm)} - {fmt_ampm(eh, em)}  "
    if g1 and g2 and g1.casefold() == g2.casefold():
        return (time_prefix, [g1])
    if g1 and g2:
        return (time_prefix, [f"Group 1 - {g1}", f"Group 2 - {g2}"])
    if g1:
        return (time_prefix, [g1])
    return (time_prefix, [g2])


def activity_label_for_bucket(g1: str, g2: str) -> str:
    g1 = (g1 or "").strip()
    g2 = (g2 or "").strip()
    if g1 and g2 and g1.casefold() == g2.casefold():
        return g1
    if g1 and g2:
        return f"Group 1 - {g1}\nGroup 2 - {g2}"
    if g1:
        return g1
    return g2


def infer_locations_from_text(text: str) -> Tuple[str, str]:
    s = (text or "").strip().casefold()
    loc1 = ""
    loc2 = ""
    if "breakfast" in s:
        loc1 = "The Barn"
    if "kayaking" in s:
        loc1 = "Bayside"
    if "meditation" in s:
        loc1 = "Dam Forest"
    if "sauna" in s or "ice bath" in s or "icebath" in s:
        loc1 = "Sauna"
    if "rally driving" in s:
        loc1 = "Watch Tower"
    if "campfire" in s or "bloke burning" in s:
        loc1 = "Fire"
    if "touch football" in s:
        loc1 = "Touch Football"
    if "shooting" in s:
        loc1 = "Shooting"
    return loc1, loc2


def daterange_inclusive(start: date, end: date) -> List[date]:
    out: List[date] = []
    d = start
    while d <= end:
        out.append(d)
        d += timedelta(days=1)
    return out


def sort_activities(rows: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    def key(r: Dict[str, Any]) -> int:
        st = parse_hhmm(r.get("start", ""))
        if not st:
            return 0
        return time_tuple_to_seconds(st[0], st[1])

    return sorted(rows, key=key)


def normalize_activity_group(val: Any) -> str:
    """Return '1' or '2' for legacy rows that used name + group."""
    if val is None:
        return "1"
    s = str(val).strip().lower()
    if s in ("2", "group 2", "g2"):
        return "2"
    if s in ("1", "group 1", "g1"):
        return "1"
    try:
        if int(float(s)) == 2:
            return "2"
    except ValueError:
        pass
    return "1"


def activity_row_to_contribution(
    a: Dict[str, Any],
) -> Optional[Tuple[Tuple[int, int], str, str, str, str, str]]:
    """One row's times + group activities + image. Legacy: name + optional group."""
    st = parse_hhmm(a.get("start", ""))
    et = parse_hhmm(a.get("end", ""))
    if not st or not et:
        return None
    ssec = time_tuple_to_seconds(st[0], st[1])
    esec = time_tuple_to_seconds(et[0], et[1])
    if esec <= ssec:
        return None
    key = (ssec, esec)
    img = (a.get("image") or "").strip()
    loc1 = (a.get("loc_g1") or "").strip()
    loc2 = (a.get("loc_g2") or "").strip()
    if "name_g1" in a or "name_g2" in a:
        g1 = (a.get("name_g1") or "").strip()
        g2 = (a.get("name_g2") or "").strip()
    else:
        nm = (a.get("name") or "").strip()
        if not nm:
            g1, g2 = "", ""
        elif normalize_activity_group(a.get("group")) == "2":
            g1, g2 = "", nm
        else:
            g1, g2 = nm, ""
    return (key, g1, g2, img, loc1, loc2)


def normalize_day_activities(activities: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """One entry per time slot: shared start/end, name_g1, name_g2, one image."""
    slot_map: Dict[Tuple[int, int], Dict[str, str]] = defaultdict(
        lambda: {"name_g1": "", "name_g2": "", "image": "", "loc_g1": "", "loc_g2": ""}
    )
    for a in sort_activities(list(activities or [])):
        c = activity_row_to_contribution(a)
        if not c:
            continue
        key, g1, g2, img, loc1, loc2 = c
        b = slot_map[key]
        if g1:
            b["name_g1"] = g1
        if g2:
            b["name_g2"] = g2
        if loc1:
            b["loc_g1"] = loc1
        if loc2:
            b["loc_g2"] = loc2
        if img and not b["image"]:
            b["image"] = img
    out: List[Dict[str, Any]] = []
    for ssec, esec in sorted(slot_map.keys(), key=lambda t: t[0]):
        b = slot_map[(ssec, esec)]
        sh, sm = ssec // 3600, (ssec % 3600) // 60
        eh, em = esec // 3600, (esec % 3600) // 60
        out.append(
            {
                "start": hhmm(sh, sm),
                "end": hhmm(eh, em),
                "name_g1": b["name_g1"],
                "name_g2": b["name_g2"],
                "loc_g1": b["loc_g1"],
                "loc_g2": b["loc_g2"],
                "image": b["image"],
            }
        )
    return out


def config_to_public_schedule(cfg: Dict[str, Any]) -> List[Dict[str, Any]]:
    """Shape for embedded JSON in exported event_status.html."""
    try:
        start_d = date.fromisoformat((cfg.get("event_start") or "")[:10])
        end_d = date.fromisoformat((cfg.get("event_end") or "")[:10])
    except ValueError:
        return []
    if end_d < start_d:
        return []
    days = cfg.get("days") or {}
    alt = ("sat", "sun", "mon")
    out: List[Dict[str, Any]] = []
    for i, d in enumerate(daterange_inclusive(start_d, end_d)):
        iso = d.isoformat()
        day_data = days.get(iso, {})
        activities = normalize_day_activities(list(day_data.get("activities") or []))

        blocks: List[Dict[str, str]] = []
        for a in activities:
            st = parse_hhmm(a.get("start", ""))
            et = parse_hhmm(a.get("end", ""))
            if not st or not et:
                continue
            ssec = time_tuple_to_seconds(st[0], st[1])
            esec = time_tuple_to_seconds(et[0], et[1])
            if esec <= ssec:
                continue
            g1 = (a.get("name_g1") or "").strip()
            g2 = (a.get("name_g2") or "").strip()
            l1 = (a.get("loc_g1") or "").strip()
            l2 = (a.get("loc_g2") or "").strip()
            if not g1 and not g2:
                continue
            if g1 and not l1:
                l1, _ = infer_locations_from_text(g1)
            if g2 and not l2:
                l2, _ = infer_locations_from_text(g2)
            if g1 and g2:
                if g1.casefold() == g2.casefold():
                    activity = g1
                else:
                    activity = f"Group 1 - {g1}\nGroup 2 - {g2}"
            elif g1:
                activity = g1
            else:
                activity = g2
            sh, sm = ssec // 3600, (ssec % 3600) // 60
            eh, em = esec // 3600, (esec % 3600) // 60
            blocks.append(
                {
                    "start": f"{sh:02d}{sm:02d}",
                    "end": f"{eh:02d}{em:02d}",
                    "activity": activity,
                    "location_g1": l1,
                    "location_g2": l2,
                }
            )
        out.append(
            {
                "dateIso": iso,
                "dayLabel": d.strftime("%A %d %b"),
                "className": alt[i % len(alt)],
                "blocks": blocks,
            }
        )
    return out


def resolve_map_image_filename(base_dir: str) -> str:
    for name in MAP_IMAGE_CANDIDATES:
        if os.path.isfile(os.path.join(base_dir, name)):
            return name
    return MAP_IMAGE_CANDIDATES[0]


def ensure_export_assets(out_dir: str, base_dir: str) -> str:
    map_name = resolve_map_image_filename(base_dir)
    src = os.path.join(base_dir, map_name)
    dst = os.path.join(out_dir, map_name)
    if os.path.isfile(src):
        try:
            if os.path.abspath(src) != os.path.abspath(dst):
                dst_parent = os.path.dirname(dst)
                if dst_parent:
                    os.makedirs(dst_parent, exist_ok=True)
                shutil.copy2(src, dst)
        except OSError:
            pass
    return map_name


def _map_image_pixel_size(map_path: str) -> Tuple[int, int]:
    """Intrinsic pixels for <img width/height>; overlay viewBox stays 767×1024."""
    try:
        from PIL import Image

        with Image.open(map_path) as im:
            w, h = im.size
            if w > 0 and h > 0:
                return w, h
    except Exception:
        pass
    return 767, 1024


def build_event_status_html(
    cfg: Dict[str, Any],
    map_image_name: str = "site_map_art_illustrated2.png",
    map_image_width: int = 767,
    map_image_height: int = 1024,
) -> str:
    schedule = config_to_public_schedule(cfg)
    payload = json.dumps(schedule, ensure_ascii=False)
    payload = payload.replace("<", "\\u003c")
    location_types = normalize_location_types_from_config(cfg.get("location_types"))
    locations_payload = json.dumps(location_types, ensure_ascii=False).replace("<", "\\u003c")
    map_src = map_image_name.replace("\\", "/")
    _gen = datetime.now().astimezone()
    _tz = (_gen.tzname() or "").strip()
    _gen_line = _gen.strftime("%d %b %Y, %I:%M %p").strip()
    if _tz:
        _gen_line = f"{_gen_line} ({_tz})"
    html_generated_label = html.escape(_gen_line)

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0"/>
  <title>Event Live Status</title>
  <style>
    html, body {{
      margin: 0;
      padding: 0;
      overflow-x: hidden;
      overflow-y: auto;
      height: auto;
      min-height: 100%;
      background: transparent;
      -ms-overflow-style: none;
      scrollbar-width: none;
    }}
    html::-webkit-scrollbar, body::-webkit-scrollbar {{
      display: none;
    }}
    body {{
      font-family: Arial, sans-serif;
      max-width: 900px;
      margin: 24px auto;
      padding: 0 20px;
      color: #333;
      min-height: auto;
    }}

    body::before {{
      content: none;
    }}

    h1 {{
      text-align: center;
      margin-bottom: 20px;
      color: #fff;
      text-shadow: 0 2px 8px rgba(0,0,0,0.7);
    }}
    h2 {{
      text-align: center;
      font-size: 2em;
      margin: 50px 0 30px;
      color: #fff;
      border-bottom: 2px solid rgba(255,255,255,0.2);
      padding-bottom: 10px;
      text-shadow: 0 1px 6px rgba(0,0,0,0.6);
    }}
    #current {{
      background: rgba(230, 247, 255, 0.94);
      border: 3px solid #1890ff;
      border-radius: 12px;
      padding: 40px 25px;
      margin-bottom: 50px;
      text-align: center;
      font-size: 2.4em;
      line-height: 1.4;
      box-shadow: 0 4px 16px rgba(0,0,0,0.3);
      backdrop-filter: blur(4px);
    }}
    #time {{
      font-size: 1.6em;
      font-weight: bold;
      margin-bottom: 20px;
      color: #0050b3;
      word-break: break-word;
      overflow-wrap: anywhere;
    }}
    #status strong {{
      color: #d4380d;
      display: block;
      margin: 12px 0 8px;
    }}
    #next-info {{
      font-size: 0.825em;
      color: #d46b08;
      font-weight: 500;
      margin-top: 20px;
      padding-top: 15px;
      border-top: 1px solid #91d5ff88;
      display: none;
    }}
    #next-info.visible {{
      display: block;
    }}
    .day-section {{
      margin-bottom: 50px;
      background: rgba(255, 255, 255, 0.97);
      border-radius: 12px;
      overflow: hidden;
      box-shadow: 0 4px 12px rgba(0,0,0,0.25);
      backdrop-filter: blur(3px);
    }}
    .day-header {{
      font-size: 1.6em;
      font-weight: bold;
      text-align: center;
      background: #444;
      color: white;
      padding: 20px;
      cursor: pointer;
      user-select: none;
      position: relative;
    }}
    .day-header::after {{
      content: "▼";
      position: absolute;
      right: 30px;
      font-size: 1.2em;
      transition: transform 0.25s ease;
    }}
    .day-header.expanded::after {{
      transform: rotate(180deg);
    }}
    .day-header:hover {{ background: #555; }}
    table.day-table {{
      width: 100%;
      border-collapse: collapse;
      background: inherit;
    }}
    td {{
      border: 1px solid #ddd8;
      padding: 16px 20px;
      vertical-align: top;
    }}
    tr:hover td {{ background: rgba(248,249,250,0.6); }}
    .time-col {{
      width: 160px;
      font-weight: 600;
      color: #444;
      white-space: nowrap;
    }}
    .activity-col {{
      line-height: 1.6;
    }}
    .location-chips {{
      margin-top: 8px;
      display: flex;
      gap: 8px;
      flex-wrap: wrap;
    }}
    .loc-chip {{
      border: 1px solid #90caf9;
      background: #e3f2fd;
      color: #0d47a1;
      border-radius: 999px;
      padding: 4px 10px;
      font-size: 0.8rem;
      font-weight: 700;
      cursor: pointer;
    }}
    .loc-chip:hover {{ background: #bbdefb; }}
    .map-card {{
      margin: 20px 0 48px;
      background: rgba(255,255,255,0.97);
      border-radius: 12px;
      box-shadow: 0 4px 12px rgba(0,0,0,0.25);
      overflow: hidden;
      padding: 14px 12px 16px;
    }}
    .map-links {{
      display: flex;
      flex-wrap: wrap;
      gap: 8px;
      margin: 8px 0 12px;
      justify-content: center;
    }}
    .map-link {{
      border: 1px solid #ffd180;
      background: #fff3e0;
      color: #bf360c;
      border-radius: 999px;
      padding: 5px 11px;
      font-size: 0.82em;
      font-weight: 700;
      cursor: pointer;
    }}
    .map-wrap {{
      position: relative;
      width: 100%;
      max-width: 767px;
      margin: 0 auto;
      line-height: 0;
    }}
    .map-wrap img {{
      width: 100%;
      height: auto;
      display: block;
    }}
    .map-target {{
      position: absolute;
      left: 50%;
      top: 50%;
      width: 20px;
      height: 20px;
      margin-left: -10px;
      margin-top: -10px;
      border-radius: 999px;
      background: #e53935;
      border: 2px solid #fff;
      box-shadow: 0 0 0 0 rgba(229,57,53,0.8);
      pointer-events: none;
      opacity: 0;
      z-index: 4;
      transform: scale(0.85);
    }}
    .map-target.active {{
      opacity: 1;
      animation: target-pulse 1s ease-out infinite;
      transform: scale(1);
    }}
    @keyframes target-pulse {{
      0% {{ box-shadow: 0 0 0 0 rgba(229,57,53,0.75); }}
      70% {{ box-shadow: 0 0 0 20px rgba(229,57,53,0); }}
      100% {{ box-shadow: 0 0 0 0 rgba(229,57,53,0); }}
    }}
    .map-wrap svg {{
      position: absolute;
      inset: 0;
      width: 100%;
      height: 100%;
      pointer-events: none;
    }}
    .map-wrap a {{
      pointer-events: auto;
      cursor: pointer;
    }}
    .map-hint {{
      text-align: center;
      color: #48606f;
      margin: 10px 0 0;
      font-size: 0.92em;
      min-height: 1.2em;
    }}
    .map-hint--generated {{
      margin-top: 4px;
      margin-bottom: 0;
      font-size: 0.8em;
      opacity: 0.92;
    }}
    .sat {{ background: rgba(255,242,230,0.9); }}
    .sun {{ background: rgba(240,230,255,0.9); }}
    .mon {{ background: rgba(230,255,251,0.9); }}
    .past {{ color: #888; opacity: 0.75; }}
    .current {{ background: rgba(217,243,190,0.9) !important; font-weight: bold; }}

    @media (max-width: 640px) {{
      body {{
        margin: 14px auto;
        padding: 0 10px;
      }}
      h1 {{
        font-size: 1.8em;
      }}
      h2 {{
        font-size: 1.6em;
        margin: 30px 0 18px;
      }}
      #current {{
        padding: 22px 14px;
        font-size: 1.75em;
        line-height: 1.3;
      }}
      #time {{
        font-size: 1.35em;
        line-height: 1.15;
        margin-bottom: 14px;
      }}
      #status strong {{
        margin: 8px 0 6px;
      }}
      .day-header {{
        font-size: 1.25em;
        padding: 14px 16px;
      }}
      .day-header::after {{
        right: 14px;
      }}
      td {{
        padding: 12px 10px;
      }}
      .time-col {{
        width: 120px;
      }}
      .map-card {{
        margin: 12px 0 32px;
        padding: 12px 8px 12px;
      }}
    }}
  </style>
</head>
<body>

<h1>Event Live Status</h1>

<div id="current">
  <div id="time"></div>
  <div id="status">Loading schedule...</div>
  <div id="next-info"></div>
</div>

<h2>Full Schedule</h2>
<div id="schedule-container"></div>

<div id="map-section" class="map-card">
  <h2 style="margin: 8px 0 12px;">Site Map</h2>
  <div id="map-links" class="map-links"></div>
  <div id="map-wrap" class="map-wrap">
    <img src="{map_src}" width="{map_image_width}" height="{map_image_height}" alt="Burning Bloke site map"/>
    <div id="map-target" class="map-target" aria-hidden="true"></div>
    <svg viewBox="0 0 767 1024" preserveAspectRatio="xMidYMid meet" aria-hidden="true">
      <a href="#the-barn" aria-label="The Barn"><rect id="zone-the-barn" x="25" y="40" width="275" height="185" fill="rgba(255,255,255,0.001)"/></a>
      <a href="#fire-1" aria-label="Fire"><rect id="zone-fire-1" x="178" y="220" width="205" height="145" rx="16" fill="rgba(255,255,255,0.001)"/></a>
      <a href="#sauna" aria-label="Sauna"><ellipse id="zone-sauna" cx="124" cy="648" rx="72" ry="58" fill="rgba(255,255,255,0.001)"/></a>
      <a href="#watch-tower" aria-label="Watch Tower"><rect id="zone-watch-tower" x="514" y="472" width="170" height="170" rx="14" fill="rgba(255,255,255,0.001)"/></a>
      <a href="#race-track" aria-label="Race Track"><polygon id="zone-race-track" points="405,150 750,150 750,700 465,700 405,640" fill="rgba(255,255,255,0.001)"/></a>
      <a href="#dam-forest" aria-label="Dam Forest"><rect id="zone-dam-forest" x="406" y="792" width="155" height="208" rx="26" fill="rgba(255,255,255,0.001)"/></a>
      <a href="#bayside" aria-label="Bayside"><rect id="zone-bayside" x="158" y="740" width="552" height="225" rx="65" fill="rgba(255,255,255,0.001)"/></a>
      <a href="#touch-football" aria-label="Touch Football"><ellipse id="zone-touch-football" cx="380" cy="400" rx="70" ry="70" fill="rgba(255,255,255,0.001)"/></a>
      <a href="#shooting" aria-label="Shooting"><ellipse id="zone-shooting" cx="210" cy="720" rx="72" ry="62" fill="rgba(255,255,255,0.001)"/></a>
    </svg>
  </div>
  <p id="map-status" class="map-hint">Click schedule/location links to target the map.</p>
  <p id="map-generated" class="map-hint map-hint--generated">HTML generated {html_generated_label}</p>
</div>

<script>
const schedule = {payload};
const locationTypes = {locations_payload};
const mapTarget = document.getElementById("map-target");
const mapStatus = document.getElementById("map-status");
const TARGET_VISIBLE_MS = 5200;
const targetOverrides = {{
  "fire-1": {{ xPct: 32.6, yPct: 27.0 }},
  "watch-tower": {{ xPct: 68.8, yPct: 53.8 }},
  "bayside": {{ xPct: 60.2, yPct: 71.8 }},
  "touch-football": {{ xPct: 64.8, yPct: 22.0 }},
  "shooting": {{ xPct: 24.8, yPct: 70.0 }},
  "sauna": {{ xPct: 16.2, yPct: 63.3 }},
  "dam-forest": {{ xPct: 63.0, yPct: 86.0 }}
}};

function locationToId(name) {{
  const s = (name || "").trim().toLowerCase();
  const map = {{
    "the barn": "the-barn",
    "fire": "fire-1",
    "sauna": "sauna",
    "watch tower": "watch-tower",
    "race track": "race-track",
    "bayside": "bayside",
    "bay side": "bayside",
    "touch football": "touch-football",
    "shooting": "shooting",
    "dam forest": "dam-forest"
  }};
  return map[s] || "";
}}

function todayIso() {{
  const n = new Date();
  const y = n.getFullYear();
  const m = String(n.getMonth() + 1).padStart(2, "0");
  const d = String(n.getDate()).padStart(2, "0");
  return `${{y}}-${{m}}-${{d}}`;
}}

function toMinutes(timeStr) {{
  const hh = parseInt(timeStr.slice(0, 2), 10);
  const mm = parseInt(timeStr.slice(2, 4), 10);
  return hh * 60 + mm;
}}

function fmtHm(timeStr) {{
  return timeStr.slice(0, 2) + ":" + timeStr.slice(2, 4);
}}

function getCurrentLocal() {{
  const n = new Date();
  const wd = n.toLocaleDateString("en-US", {{ weekday: "long" }});
  const hh = n.getHours();
  const mm = n.getMinutes();
  return {{
    dateIso: todayIso(),
    dayOfWeek: wd,
    currentMinutes: hh * 60 + mm,
    timeStr: `${{wd}} ${{String(hh).padStart(2, "0")}}:${{String(mm).padStart(2, "0")}}`
  }};
}}

function getCurrentAndNext() {{
  const {{ dateIso, currentMinutes }} = getCurrentLocal();
  const daySchedule = schedule.find(s => s.dateIso === dateIso);
  if (!daySchedule) {{
    return {{ statusHtml: "No event scheduled today", nextText: "" }};
  }}
  let current = null;
  let nextBlock = null;
  for (const block of daySchedule.blocks) {{
    const s = toMinutes(block.start);
    const e = toMinutes(block.end);
    if (currentMinutes >= s && currentMinutes < e) {{
      current = block.activity;
    }}
    if (currentMinutes < s && (!nextBlock || s < toMinutes(nextBlock.start))) {{
      nextBlock = block;
    }}
  }}
  let statusHtml = current
    ? `Currently:<br><strong>${{current.replace(/\\n/g, "<br>")}}</strong>`
    : "No activity right now";
  let nextText = "";
  if (nextBlock) {{
    nextText = `Next: ${{nextBlock.activity.split("\\n")[0].trim()}} — ${{fmtHm(nextBlock.start)}}–${{fmtHm(nextBlock.end)}}`;
  }}
  return {{ statusHtml, nextText }};
}}

function updateDisplay() {{
  const {{ timeStr }} = getCurrentLocal();
  const {{ statusHtml, nextText }} = getCurrentAndNext();
  document.getElementById("time").textContent = timeStr;
  document.getElementById("status").innerHTML = statusHtml;
  const nextEl = document.getElementById("next-info");
  if (nextText) {{
    nextEl.textContent = nextText;
    nextEl.classList.add("visible");
  }} else {{
    nextEl.classList.remove("visible");
  }}
}}

function buildSchedule() {{
  const container = document.getElementById("schedule-container");
  const today = todayIso();
  const {{ currentMinutes }} = getCurrentLocal();
  schedule.forEach(dayObj => {{
    const section = document.createElement("div");
    section.className = "day-section";
    const header = document.createElement("div");
    header.className = "day-header expanded";
    header.textContent = dayObj.dayLabel;
    header.addEventListener("click", () => {{
      const table = section.querySelector("table");
      const isHidden = table.style.display === "none";
      table.style.display = isHidden ? "" : "none";
      header.classList.toggle("expanded", isHidden);
    }});
    section.appendChild(header);
    const table = document.createElement("table");
    table.className = "day-table";
    const tbody = document.createElement("tbody");
    dayObj.blocks.forEach(block => {{
      const row = document.createElement("tr");
      row.classList.add(dayObj.className);
      const sMin = toMinutes(block.start);
      const eMin = toMinutes(block.end);
      if (dayObj.dateIso === today) {{
        if (currentMinutes >= sMin && currentMinutes < eMin) row.classList.add("current");
        else if (currentMinutes >= eMin) row.classList.add("past");
      }}
      const act = block.activity.replace(/\\n/g, "<br>");
      const loc1 = (block.location_g1 || "").trim();
      const loc2 = (block.location_g2 || "").trim();
      const locSet = [];
      if (loc1) locSet.push(loc1);
      if (loc2 && loc2.toLowerCase() !== loc1.toLowerCase()) locSet.push(loc2);
      let chipsHtml = "";
      if (locSet.length) {{
        chipsHtml = `<div class="location-chips">` + locSet.map((loc) => {{
          const id = locationToId(loc);
          if (!id) return "";
          return `<button type="button" class="loc-chip" data-map-target="${{id}}">Location: ${{loc}}</button>`;
        }}).join("") + `</div>`;
      }}
      row.innerHTML = `<td class="time-col">${{fmtHm(block.start)}} – ${{fmtHm(block.end)}}</td><td class="activity-col">${{act}}${{chipsHtml}}</td>`;
      tbody.appendChild(row);
    }});
    table.appendChild(tbody);
    section.appendChild(table);
    container.appendChild(section);
  }});
  container.addEventListener("click", (event) => {{
    const btn = event.target && event.target.closest ? event.target.closest(".loc-chip") : null;
    if (!btn) return;
    const targetId = btn.getAttribute("data-map-target") || "";
    if (targetId) {{
      focusMapLocation(targetId, btn.textContent.replace("Location: ", "").trim());
    }}
  }});
}}

function buildLocationLinks() {{
  const wrap = document.getElementById("map-links");
  if (!wrap) return;
  wrap.innerHTML = "";
  locationTypes.forEach((loc) => {{
    const id = locationToId(loc);
    if (!id) return;
    const b = document.createElement("button");
    b.type = "button";
    b.className = "map-link";
    b.textContent = loc;
    b.addEventListener("click", () => focusMapLocation(id, loc));
    wrap.appendChild(b);
  }});
}}

/** Visible viewport height (handles mobile URL bar / iframe). */
function viewportHeightPx() {{
  if (window.visualViewport && window.visualViewport.height) return window.visualViewport.height;
  return window.innerHeight || 0;
}}

function scrollMainDocumentToElement(el, block) {{
  if (!el) return;
  const b = block || "center";
  const step = () => {{
    try {{
      const ae = document.activeElement;
      if (ae && typeof ae.blur === "function") ae.blur();
    }} catch (_x) {{}}
    const se = document.scrollingElement || document.documentElement;
    const r = el.getBoundingClientRect();
    const y0 = window.pageYOffset || se.scrollTop || 0;
    const vh = viewportHeightPx();
    if (!vh) return;
    let top;
    if (b === "start") {{
      top = r.top + y0 - 10;
    }} else {{
      const mid = r.top + y0 + r.height / 2;
      top = mid - vh / 2;
      const overshootPx = Math.min(80, Math.max(24, Math.round(vh * 0.05)));
      top -= overshootPx;
    }}
    top = Math.max(0, top);
    const mq =
      typeof window.matchMedia === "function"
        ? window.matchMedia("(max-width: 768px), (pointer: coarse)")
        : null;
    const preferInstant = mq && mq.matches;
    const behavior = preferInstant ? "auto" : "smooth";
    const apply = (beh) => {{
      try {{
        se.scrollTo({{ left: 0, top: top, behavior: beh }});
      }} catch (_e) {{
        try {{
          window.scrollTo(0, top);
        }} catch (_e2) {{
          se.scrollTop = top;
        }}
      }}
    }};
    apply(behavior);
    if (!preferInstant) {{
      setTimeout(() => {{
        const y1 = window.pageYOffset || se.scrollTop || 0;
        if (Math.abs(y1 - top) > 56) apply("auto");
      }}, 160);
    }}
  }};
  requestAnimationFrame(() => requestAnimationFrame(step));
}}

/**
 * When this page is in an iframe (e.g. Squarespace), the *host* page scrolls — not the iframe's window.
 * Standalone scroll uses scrollMainDocumentToElement; embeds need the parent to listen for this message.
 * Parent listens for postMessage type bb-schedule-scroll-to-map (mapMidY, mapTop, iframeViewportH).
 */
function notifyParentScrollToMap(el) {{
  if (!el) return;
  let embedded = false;
  try {{
    embedded = window.self !== window.top;
  }} catch (_e) {{
    embedded = true;
  }}
  if (!embedded) return;
  const measure = () => {{
    try {{
      const se = document.scrollingElement || document.documentElement;
      const r = el.getBoundingClientRect();
      const y0 = window.pageYOffset || se.scrollTop || 0;
      const mapTop = r.top + y0;
      const mapMidY = mapTop + r.height / 2;
      window.parent.postMessage(
        {{
          type: "bb-schedule-scroll-to-map",
          mapMidY: mapMidY,
          mapTop: mapTop,
          iframeViewportH: viewportHeightPx() || window.innerHeight || 0,
        }},
        "*"
      );
    }} catch (_e) {{}}
  }};
  requestAnimationFrame(() => requestAnimationFrame(measure));
}}

function focusMapLocation(id, label) {{
  const mapSection = document.getElementById("map-section");
  const mapFocusEl = document.getElementById("map-wrap") || mapSection;
  const zone = document.getElementById("zone-" + id);
  scrollMainDocumentToElement(mapFocusEl, "center");
  notifyParentScrollToMap(mapFocusEl);
  setTimeout(() => notifyParentScrollToMap(mapFocusEl), 120);
  setTimeout(() => notifyParentScrollToMap(mapFocusEl), 450);
  setTimeout(() => {{
    if (!zone || !mapTarget) return;
    const override = targetOverrides[id];
    const box = zone.getBBox();
    const cx = override ? override.xPct : ((box.x + box.width / 2) / 767) * 100;
    const cy = override ? override.yPct : ((box.y + box.height / 2) / 1024) * 100;
    mapTarget.style.left = `${{cx}}%`;
    mapTarget.style.top = `${{cy}}%`;
    mapTarget.classList.remove("active");
    void mapTarget.offsetWidth;
    mapTarget.classList.add("active");
    setTimeout(() => mapTarget.classList.remove("active"), TARGET_VISIBLE_MS);
  }}, 450);
  mapStatus.textContent = `Targeted: ${{label}}`;
  try {{
    if (window.history && window.history.replaceState) {{
      window.history.replaceState(null, "", "#" + id);
    }}
  }} catch (_e) {{}}
  notifyParentHeight();
  setTimeout(() => notifyParentScrollToMap(mapFocusEl), 220);
}}

function bindMapClicks() {{
  document.querySelectorAll("#map-wrap svg a").forEach((anchor) => {{
    anchor.addEventListener("click", (event) => {{
      event.preventDefault();
      const href = anchor.getAttribute("href") || "";
      if (!href.startsWith("#")) return;
      const id = href.slice(1);
      const label = (anchor.getAttribute("aria-label") || id).trim();
      focusMapLocation(id, label);
    }});
  }});
}}

function notifyParentHeight() {{
  const h = Math.max(
    document.body.scrollHeight,
    document.documentElement.scrollHeight,
    document.body.offsetHeight,
    document.documentElement.offsetHeight
  );
  try {{
    window.parent.postMessage({{ type: "bb-schedule-height", height: h }}, "*");
  }} catch (_e) {{}}
}}

function bindHeightObservers() {{
  const mapImg = document.querySelector("#map-wrap img");
  if (mapImg) {{
    if (mapImg.complete) {{
      notifyParentHeight();
    }} else {{
      mapImg.addEventListener("load", notifyParentHeight, {{ once: true }});
      mapImg.addEventListener("error", notifyParentHeight, {{ once: true }});
    }}
  }}
  if (typeof ResizeObserver !== "undefined") {{
    try {{
      const ro = new ResizeObserver(() => notifyParentHeight());
      ro.observe(document.body);
      ro.observe(document.documentElement);
    }} catch (_e) {{}}
  }}
  window.addEventListener("load", notifyParentHeight);
}}

buildSchedule();
buildLocationLinks();
bindMapClicks();
bindHeightObservers();
updateDisplay();
setInterval(updateDisplay, 60000);
notifyParentHeight();
setTimeout(notifyParentHeight, 50);
setTimeout(notifyParentHeight, 400);
window.addEventListener("resize", notifyParentHeight);
</script>
</body>
</html>
"""


# --- Image -------------------------------------------------------------


def load_scaled_photo(
    path: str, max_w: int, max_h: int
) -> Tuple[Optional[Any], Optional[Any]]:
    """PIL is imported here so the admin window starts without loading Pillow/NumPy."""
    from PIL import Image, ImageTk

    try:
        im = Image.open(path).convert("RGBA")
    except OSError:
        return None, None
    iw, ih = im.size
    if iw <= 0 or ih <= 0:
        return None, None
    scale = min(max_w / iw, max_h / ih)
    nw, nh = max(1, int(iw * scale)), max(1, int(ih * scale))
    im = im.resize((nw, nh), Image.Resampling.LANCZOS)
    return ImageTk.PhotoImage(im), im


# --- Fullscreen display -------------------------------------------------


class DisplayWindow(tk.Toplevel):
    DISPLAY_TICK_MS = 250
    DISPLAY_SCHEDULE_PHASE_S = 15.0
    DISPLAY_ACTIVITY_IMAGE_PHASE_S = 10.0
    SCHEDULE_CHECK_MS = 5000
    CLOCK_FONT = DISPLAY_CLOCK_FONT
    CLOCK_PAD = 30
    QR_PAD = 30

    def __init__(
        self,
        master: tk.Misc,
        get_config: Callable[[], Dict[str, Any]],
        on_close: Optional[Callable[[], None]] = None,
    ):
        super().__init__(master)
        self._get_config = get_config
        self._on_close = on_close
        self.title("Display")
        self.configure(bg="black")
        self.attributes("-fullscreen", True)
        self.overrideredirect(False)
        try:
            self.attributes("-topmost", True)
        except tk.TclError:
            pass
        try:
            self.config(cursor="none")
        except tk.TclError:
            pass
        apply_window_icon(self)

        self._canvas = tk.Canvas(self, bg="black", highlightthickness=0, bd=0)
        self._canvas.pack(fill="both", expand=True)

        self._img_tk: Optional[Any] = None
        self._pil_ref: Optional[Any] = None
        self._qr_tk: Optional[Any] = None
        self._qr_cache_key: Optional[Tuple[str, int]] = None

        self._clock_shadow_id = self._canvas.create_text(
            0,
            0,
            text=datetime.now().strftime("%H:%M:%S"),
            anchor="se",
            fill="#000000",
            font=self.CLOCK_FONT,
            tags="clock",
        )
        self._clock_fg_id = self._canvas.create_text(
            0,
            0,
            text=datetime.now().strftime("%H:%M:%S"),
            anchor="se",
            fill="#ffffff",
            font=self.CLOCK_FONT,
            tags="clock",
        )

        self._bucket: Tuple[Any, ...] = ("init",)
        self._last_bucket_for_alt: Optional[Tuple[Any, ...]] = None
        self._alt_anchor_mono: float = time.monotonic()
        self._tick_after: Optional[str] = None
        self._schedule_after: Optional[str] = None
        self._clock_after: Optional[str] = None

        self._canvas.bind("<Configure>", self._on_canvas_configure)

        self.bind("<Escape>", lambda e: self.close())
        self.protocol("WM_DELETE_WINDOW", self.close)

        self._refresh_schedule_bucket()
        self._schedule_after = self.after(self.SCHEDULE_CHECK_MS, self._schedule_check_loop)
        self._clock_tick()
        self._display_tick()

    def _on_canvas_configure(self, _evt=None):
        self._layout_clock()
        w = max(self._canvas.winfo_width(), 1)
        h = max(self._canvas.winfo_height(), 1)
        self._sync_qr_overlay(w, h)

    def _layout_clock(self):
        w = max(self._canvas.winfo_width(), 1)
        h = max(self._canvas.winfo_height(), 1)
        cx = w - self.CLOCK_PAD
        cy = h - self.CLOCK_PAD
        self._canvas.coords(self._clock_shadow_id, cx + 2, cy + 2)
        self._canvas.coords(self._clock_fg_id, cx, cy)
        self._canvas.tag_raise("clock")

    def _sync_qr_overlay(self, w: int, h: int) -> None:
        cfg = self._get_config()
        fn = (cfg.get("qr_code_image") or "").strip()
        edge = min(w, h)
        side = max(160, min(int(edge * 0.19), 240))
        x, y = self.QR_PAD, h - self.QR_PAD
        if not fn:
            self._canvas.delete("qr")
            self._qr_tk = None
            self._qr_cache_key = None
            self._canvas.tag_raise("clock")
            return
        path = os.path.join(app_directory(), fn)
        key: Tuple[str, int] = (fn, side)
        if key == self._qr_cache_key and self._qr_tk is not None:
            for iid in self._canvas.find_withtag("qr"):
                self._canvas.coords(iid, x, y)
            self._canvas.tag_raise("qr")
            self._canvas.tag_raise("clock")
            return
        if not os.path.isfile(path):
            self._canvas.delete("qr")
            self._qr_tk = None
            self._qr_cache_key = None
            self._canvas.tag_raise("clock")
            return
        photo, _pil = load_scaled_photo(path, side, side)
        if not photo:
            self._canvas.delete("qr")
            self._qr_tk = None
            self._qr_cache_key = None
            self._canvas.tag_raise("clock")
            return
        self._canvas.delete("qr")
        self._qr_tk = photo
        self._qr_cache_key = key
        self._canvas.create_image(x, y, image=photo, anchor="sw", tags="qr")
        self._canvas.tag_raise("qr")
        self._canvas.tag_raise("clock")

    def _clock_tick(self):
        t = datetime.now().strftime("%H:%M:%S")
        self._canvas.itemconfig(self._clock_shadow_id, text=t)
        self._canvas.itemconfig(self._clock_fg_id, text=t)
        self._layout_clock()
        self._clock_after = self.after(1000, self._clock_tick)

    def close(self):
        if self._tick_after:
            try:
                self.after_cancel(self._tick_after)
            except tk.TclError:
                pass
            self._tick_after = None
        if self._schedule_after:
            try:
                self.after_cancel(self._schedule_after)
            except tk.TclError:
                pass
            self._schedule_after = None
        if self._clock_after:
            try:
                self.after_cancel(self._clock_after)
            except tk.TclError:
                pass
            self._clock_after = None
        if self._on_close:
            self._on_close()
        self.destroy()

    def _schedule_check_loop(self):
        self._refresh_schedule_bucket()
        self._schedule_after = self.after(self.SCHEDULE_CHECK_MS, self._schedule_check_loop)

    def _refresh_schedule_bucket(self):
        cfg = self._get_config()
        self._bucket = self._compute_bucket(cfg)

    def _compute_bucket(self, cfg: Dict[str, Any]) -> Tuple[Any, ...]:
        today = date.today()
        start_s = cfg.get("event_start") or ""
        end_s = cfg.get("event_end") or ""
        try:
            start_d = date.fromisoformat(start_s) if start_s else None
            end_d = date.fromisoformat(end_s) if end_s else None
        except ValueError:
            return ("bad_config",)

        if not start_d or not end_d or today < start_d or today > end_d:
            return ("out_of_range", today.isoformat())

        day_key = today.isoformat()
        days = cfg.get("days") or {}
        day = days.get(day_key) or {}
        activities = normalize_day_activities(list(day.get("activities") or []))

        now_sec = now_seconds_since_midnight()

        if not activities:
            return ("no_activities", day_key)

        parsed: List[Tuple[int, int, str, str, str]] = []
        for a in activities:
            st = parse_hhmm(a.get("start", ""))
            et = parse_hhmm(a.get("end", ""))
            if not st or not et:
                continue
            ssec = time_tuple_to_seconds(st[0], st[1])
            esec = time_tuple_to_seconds(et[0], et[1])
            if esec <= ssec:
                continue
            g1 = (a.get("name_g1") or "").strip()
            g2 = (a.get("name_g2") or "").strip()
            if not g1 and not g2:
                continue
            img = (a.get("image") or "").strip()
            parsed.append((ssec, esec, g1, g2, img))

        if not parsed:
            return ("no_valid_activities", day_key)

        first_s = parsed[0][0]
        last_e = parsed[-1][1]

        if now_sec < first_s:
            return ("before_day", day_key)
        if now_sec >= last_e:
            return ("after_day", day_key)

        for ssec, esec, cg1, cg2, img in parsed:
            if ssec <= now_sec < esec:
                day_label = today.strftime("%A")
                schedule_specs: List[Tuple[str, List[str]]] = []
                for ss, ee, gg1, gg2, _i in parsed:
                    sh, sm = ss // 3600, (ss % 3600) // 60
                    eh, em = ee // 3600, (ee % 3600) // 60
                    row = schedule_slot_spec(sh, sm, eh, em, gg1, gg2)
                    if row:
                        schedule_specs.append(row)
                name = activity_label_for_bucket(cg1, cg2)
                return ("in_activity", day_key, day_label, schedule_specs, name, img, ssec, esec)

        day_label = today.strftime("%A")
        schedule_specs = []
        for ss, ee, gg1, gg2, _i in parsed:
            sh, sm = ss // 3600, (ss % 3600) // 60
            eh, em = ee // 3600, (ee % 3600) // 60
            row = schedule_slot_spec(sh, sm, eh, em, gg1, gg2)
            if row:
                schedule_specs.append(row)
        return ("gap", day_key, day_label, schedule_specs)

    def _clear_visual_layer(self):
        self._canvas.delete("display_image")
        self._canvas.delete("message_text")
        self._canvas.delete("schedule_overlay")

    def _paint_image_centered(self, photo: Any):
        self._clear_visual_layer()
        w = max(self._canvas.winfo_width(), 1)
        h = max(self._canvas.winfo_height(), 1)
        self._img_tk = photo
        self._canvas.create_image(w // 2, h // 2, image=photo, anchor="center", tags="display_image")
        self._canvas.tag_lower("display_image")
        self._canvas.tag_raise("clock")
        self._sync_qr_overlay(w, h)

    def _paint_message_centered(self, msg: str):
        self._clear_visual_layer()
        w = max(self._canvas.winfo_width(), 1)
        h = max(self._canvas.winfo_height(), 1)
        self._img_tk = None
        self._canvas.create_text(
            w // 2,
            h // 2,
            text=msg,
            fill="#ffffff",
            font=DISPLAY_MSG_FONT,
            tags="message_text",
        )
        self._canvas.tag_raise("clock")
        self._sync_qr_overlay(w, h)

    def _display_tick(self):
        bucket = self._bucket

        w = max(self._canvas.winfo_width(), 1)
        h = max(self._canvas.winfo_height(), 1)

        base = app_directory()
        default_names = ["default.jpg", "default.png"]

        def path_for(filename: str) -> str:
            return os.path.join(base, filename) if filename else ""

        def show_file(filename: str, alt_text: Optional[str] = None):
            p = path_for(filename)
            if p and os.path.isfile(p):
                photo, pil = load_scaled_photo(p, w, h)
                if photo:
                    self._pil_ref = pil
                    self._paint_image_centered(photo)
                    return
            self._pil_ref = None
            self._paint_message_centered(alt_text or "")

        def show_default_or_black(msg: str):
            for dn in default_names:
                p = os.path.join(base, dn)
                if os.path.isfile(p):
                    show_file(dn)
                    return
            self._pil_ref = None
            self._paint_message_centered(msg)

        def show_schedule_background(day_label: str, specs: List[Tuple[str, List[str]]]):
            # Background image (preferred) else default/black.
            bg = os.path.join(base, BACKGROUND_IMAGE_NAME)
            if os.path.isfile(bg):
                photo, pil = load_scaled_photo(bg, w, h)
                if photo:
                    self._pil_ref = pil
                    self._paint_image_centered(photo)
                else:
                    show_default_or_black("")
            else:
                show_default_or_black("")

            # Overlay text in the "time" font family (Segoe UI).
            self._canvas.delete("schedule_overlay")
            title = f"{day_label}'s Schedule"
            title_font = ("Segoe UI", 68, "bold")
            row_font = ("Segoe UI", 34, "bold")
            row_line_spacing = 76
            title_to_activities_gap = 176

            x_center = w // 2
            ty = int(h * 0.10)
            self._canvas.create_text(
                x_center,
                ty,
                text=title,
                fill="#ffffff",
                font=title_font,
                tags="schedule_overlay",
            )

            row_f = tkfont.Font(family="Segoe UI", size=34, weight="bold")
            max_activity_lines = 14
            total_lines = sum(len(bodies) for _, bodies in specs)
            render_specs: List[Tuple[str, List[str]]] = []
            used = 0
            for time_p, bodies in specs:
                n = len(bodies)
                if used + n > max_activity_lines:
                    break
                render_specs.append((time_p, bodies))
                used += n
            leftover_lines = total_lines - used

            y = float(ty + title_to_activities_gap)
            for time_p, bodies in render_specs:
                tw = row_f.measure(time_p)
                b0 = bodies[0]
                block_w = tw + row_f.measure(b0)
                x0 = x_center - block_w / 2
                x_body = x0 + tw
                self._canvas.create_text(
                    x0,
                    y,
                    text=time_p,
                    anchor="w",
                    fill="#ffffff",
                    font=row_font,
                    tags="schedule_overlay",
                )
                self._canvas.create_text(
                    x_body,
                    y,
                    text=b0,
                    anchor="w",
                    fill="#ffffff",
                    font=row_font,
                    tags="schedule_overlay",
                )
                y += row_line_spacing
                for bline in bodies[1:]:
                    self._canvas.create_text(
                        x_body,
                        y,
                        text=bline,
                        anchor="w",
                        fill="#ffffff",
                        font=row_font,
                        tags="schedule_overlay",
                    )
                    y += row_line_spacing

            if leftover_lines > 0:
                self._canvas.create_text(
                    x_center,
                    y,
                    text=f"... and {leftover_lines} more",
                    fill="#ffffff",
                    font=row_font,
                    tags="schedule_overlay",
                )

            self._canvas.tag_raise("schedule_overlay")
            self._canvas.tag_raise("clock")
            self._sync_qr_overlay(w, h)

        kind = bucket[0]

        if kind in ("bad_config", "out_of_range", "init", "before_day", "after_day"):
            show_default_or_black("Event not active")
            self._tick_after = self.after(self.DISPLAY_TICK_MS, self._display_tick)
            return

        if kind == "no_activities":
            show_default_or_black("Event not active")
            self._tick_after = self.after(self.DISPLAY_TICK_MS, self._display_tick)
            return

        if kind == "no_valid_activities":
            show_default_or_black("Event not active")
            self._tick_after = self.after(self.DISPLAY_TICK_MS, self._display_tick)
            return

        if kind == "gap":
            _, _day_key, day_label, specs = bucket
            show_schedule_background(day_label, specs)
            self._tick_after = self.after(self.DISPLAY_TICK_MS, self._display_tick)
            return

        if kind == "in_activity":
            _, _day_key, day_label, specs, _act_name, act_img, _ssec, _esec = bucket
            if bucket != self._last_bucket_for_alt:
                self._last_bucket_for_alt = bucket
                self._alt_anchor_mono = time.monotonic()

            elapsed = time.monotonic() - self._alt_anchor_mono
            cycle = self.DISPLAY_SCHEDULE_PHASE_S + self.DISPLAY_ACTIVITY_IMAGE_PHASE_S
            phase = elapsed % cycle
            if phase < self.DISPLAY_SCHEDULE_PHASE_S:
                show_schedule_background(day_label, specs)
            else:
                fn = act_img
                if fn:
                    show_file(fn)
                else:
                    show_default_or_black("Event not active")
        else:
            show_default_or_black("Event not active")

        self._tick_after = self.after(self.DISPLAY_TICK_MS, self._display_tick)


# --- Admin UI -----------------------------------------------------------


class ActivityRow(ctk.CTkFrame):
    _TIME_ENTRY_W = 52

    def __init__(
        self,
        parent: ctk.CTkFrame | ctk.CTkScrollableFrame,
        on_remove: Callable[[Any], None],
        admin: "AdminApp",
    ):
        super().__init__(parent, fg_color="transparent")
        self.on_remove = on_remove
        self._admin = admin

        ctk.CTkLabel(self, text="Start", text_color=SUBTLE_TEXT_COLOR).grid(row=0, column=0, sticky="w")
        self.start_h = self._make_time_entry(0, 23)
        self.start_m = self._make_time_entry(0, 59)
        self.start_h.grid(row=1, column=0, padx=(0, 2), sticky="nw")
        ctk.CTkLabel(self, text=":", text_color=SUBTLE_TEXT_COLOR).grid(row=1, column=1, sticky="w")
        self.start_m.grid(row=1, column=2, padx=(2, 12), sticky="nw")

        ctk.CTkLabel(self, text="End", text_color=SUBTLE_TEXT_COLOR).grid(row=0, column=3, sticky="w")
        self.end_h = self._make_time_entry(0, 23)
        self.end_m = self._make_time_entry(0, 59)
        self.end_h.grid(row=1, column=3, padx=(0, 2), sticky="nw")
        ctk.CTkLabel(self, text=":", text_color=SUBTLE_TEXT_COLOR).grid(row=1, column=4, sticky="w")
        self.end_m.grid(row=1, column=5, padx=(2, 8), sticky="nw")

        types = admin.get_activity_types()
        combo_vals = ([""] + types) if types else [""]

        ctk.CTkLabel(self, text="Group 1", text_color=SUBTLE_TEXT_COLOR).grid(
            row=0, column=6, sticky="w"
        )
        self.combo_g1 = ctk.CTkComboBox(
            self,
            values=combo_vals,
            state="readonly",
            width=168,
        )
        self.combo_g1.grid(row=1, column=6, padx=(0, 6), sticky="nw")
        self.combo_g1.set(combo_vals[0])
        admin.register_activity_combo(self.combo_g1)

        ctk.CTkLabel(self, text="Group 2", text_color=SUBTLE_TEXT_COLOR).grid(
            row=0, column=7, sticky="w"
        )
        self.combo_g2 = ctk.CTkComboBox(
            self,
            values=combo_vals,
            state="readonly",
            width=168,
        )
        self.combo_g2.grid(row=1, column=7, padx=(0, 6), sticky="nw")
        self.combo_g2.set(combo_vals[0])
        admin.register_activity_combo(self.combo_g2)

        loc_types = admin.get_location_types()
        loc_vals = ([""] + loc_types) if loc_types else [""]

        ctk.CTkLabel(self, text="Loc 1", text_color=SUBTLE_TEXT_COLOR).grid(
            row=0, column=8, sticky="w"
        )
        self.combo_l1 = ctk.CTkComboBox(
            self,
            values=loc_vals,
            state="readonly",
            width=130,
        )
        self.combo_l1.grid(row=1, column=8, padx=(0, 6), sticky="nw")
        self.combo_l1.set(loc_vals[0])
        admin.register_location_combo(self.combo_l1)

        ctk.CTkLabel(self, text="Loc 2", text_color=SUBTLE_TEXT_COLOR).grid(
            row=0, column=9, sticky="w"
        )
        self.combo_l2 = ctk.CTkComboBox(
            self,
            values=loc_vals,
            state="readonly",
            width=130,
        )
        self.combo_l2.grid(row=1, column=9, padx=(0, 6), sticky="nw")
        self.combo_l2.set(loc_vals[0])
        admin.register_location_combo(self.combo_l2)

        ctk.CTkLabel(self, text="Image", text_color=SUBTLE_TEXT_COLOR).grid(
            row=0, column=10, sticky="w"
        )
        self.img_path = ""
        self.img_lbl = ctk.CTkLabel(
            self,
            text="(no image)",
            text_color=SUBTLE_TEXT_COLOR,
            width=112,
            anchor="w",
            justify="left",
        )
        self.img_lbl.grid(row=1, column=10, padx=(0, 4), sticky="nw")

        ctk.CTkButton(self, text="Image", width=72, command=self._browse_image).grid(
            row=1, column=11, padx=(0, 4), sticky="nw"
        )
        ctk.CTkButton(self, text="✕", width=32, command=self._clear_image).grid(
            row=1, column=12, padx=(0, 4), sticky="nw"
        )
        ctk.CTkButton(self, text="Remove Activity", width=110, command=lambda: self.on_remove(self)).grid(
            row=1, column=13, padx=(0, 4), sticky="nw"
        )

    def _make_time_entry(self, lo: int, hi: int) -> ctk.CTkEntry:
        e = ctk.CTkEntry(self, width=self._TIME_ENTRY_W, justify="center")

        def clamp(_evt=None) -> None:
            raw = e.get().strip()
            try:
                v = int(raw)
            except ValueError:
                v = lo
            v = max(lo, min(hi, v))
            e.delete(0, "end")
            e.insert(0, f"{v:02d}")

        e.bind("<FocusOut>", clamp)
        e.bind("<Return>", clamp)
        return e

    @staticmethod
    def _display_image_name(filename: str) -> str:
        if not filename:
            return "(no image)"
        n = _IMAGE_NAME_DISPLAY_MAX
        if len(filename) <= n:
            return filename
        return filename[: n - 1] + "…"

    def _refresh_image_label(self) -> None:
        self.img_lbl.configure(
            text=self._display_image_name(self.img_path),
            text_color=LABEL_TEXT_COLOR if self.img_path else SUBTLE_TEXT_COLOR,
        )

    def _clear_image(self) -> None:
        self.img_path = ""
        self._refresh_image_label()

    def _browse_image(self):
        p = filedialog.askopenfilename(
            title="Select activity image",
            filetypes=[("Images", "*.jpg *.jpeg *.png"), ("All", "*.*")],
        )
        if not p:
            return
        self.img_path = os.path.basename(p)
        dest = os.path.join(app_directory(), self.img_path)
        try:
            if os.path.abspath(p) != os.path.abspath(dest):
                shutil.copy2(p, dest)
        except OSError as e:
            themed_message(self, "Copy failed", str(e), kind="error")
            return
        self._refresh_image_label()

    def get_data(self) -> Dict[str, str]:
        def parse_cell(ent: ctk.CTkEntry, lo: int, hi: int) -> int:
            raw = ent.get().strip()
            try:
                v = int(raw)
            except ValueError:
                v = lo
            return max(lo, min(hi, v))

        sh = parse_cell(self.start_h, 0, 23)
        sm = parse_cell(self.start_m, 0, 59)
        eh = parse_cell(self.end_h, 0, 23)
        em = parse_cell(self.end_m, 0, 59)
        return {
            "start": hhmm(sh, sm),
            "end": hhmm(eh, em),
            "name_g1": self.combo_g1.get().strip(),
            "name_g2": self.combo_g2.get().strip(),
            "loc_g1": self.combo_l1.get().strip(),
            "loc_g2": self.combo_l2.get().strip(),
            "image": self.img_path,
        }

    def set_data(self, data: Dict[str, Any]):
        st = parse_hhmm(data.get("start", "07:00")) or (7, 0)
        et = parse_hhmm(data.get("end", "08:00")) or (8, 0)
        self.start_h.delete(0, "end")
        self.start_h.insert(0, f"{st[0]:02d}")
        self.start_m.delete(0, "end")
        self.start_m.insert(0, f"{st[1]:02d}")
        self.end_h.delete(0, "end")
        self.end_h.insert(0, f"{et[0]:02d}")
        self.end_m.delete(0, "end")
        self.end_m.insert(0, f"{et[1]:02d}")
        types = self._admin.get_activity_types()
        combo_vals = ([""] + types) if types else [""]
        self.combo_g1.configure(values=combo_vals)
        self.combo_g2.configure(values=combo_vals)
        loc_types = self._admin.get_location_types()
        loc_vals = ([""] + loc_types) if loc_types else [""]
        self.combo_l1.configure(values=loc_vals)
        self.combo_l2.configure(values=loc_vals)

        def resolve_choice(stored: str) -> str:
            s = (stored or "").strip()
            return s if s in combo_vals else combo_vals[0]

        g1 = (data.get("name_g1") or "").strip()
        g2 = (data.get("name_g2") or "").strip()
        if "name_g1" not in data and "name_g2" not in data:
            nm = (data.get("name") or "").strip()
            if normalize_activity_group(data.get("group")) == "2":
                g1, g2 = "", nm
            else:
                g1, g2 = nm, ""

        self.combo_g1.set(resolve_choice(g1))
        self.combo_g2.set(resolve_choice(g2))
        l1 = (data.get("loc_g1") or "").strip()
        l2 = (data.get("loc_g2") or "").strip()
        if g1 and not l1:
            l1, _ = infer_locations_from_text(g1)
        if g2 and not l2:
            l2, _ = infer_locations_from_text(g2)
        self.combo_l1.set(l1 if l1 in loc_vals else loc_vals[0])
        self.combo_l2.set(l2 if l2 in loc_vals else loc_vals[0])
        self.img_path = (data.get("image") or "").strip()
        self._refresh_image_label()


class DaySection(ctk.CTkFrame):
    def __init__(self, parent: ctk.CTkScrollableFrame, day: date, admin: "AdminApp"):
        super().__init__(parent, fg_color="transparent")
        self.day = day
        self.iso = day.isoformat()
        self._admin = admin

        card = ctk.CTkFrame(self)
        card.pack(fill="x", pady=(0, 8))

        title = day.strftime("%A %d %b")
        ctk.CTkLabel(
            card,
            text=title,
            font=ctk.CTkFont(size=15, weight="bold"),
        ).pack(anchor="w", padx=8, pady=(8, 4))

        self.act_frame = ctk.CTkFrame(card, fg_color="transparent")
        self.act_frame.pack(fill="x", padx=4, pady=(8, 4))
        self.rows: List[ActivityRow] = []

        ctk.CTkButton(card, text="+ Add Activity", width=140, command=self.add_row).pack(
            anchor="w", padx=8, pady=(4, 8)
        )

        ctk.CTkFrame(card, height=1, fg_color=("gray70", "gray35")).pack(fill="x", pady=(4, 0))

    def add_row(self, data: Optional[Dict[str, Any]] = None):
        row = ActivityRow(self.act_frame, self._remove_row, self._admin)
        row.pack(fill="x", pady=4)
        self.rows.append(row)
        if data:
            row.set_data(data)

    def _remove_row(self, row: ActivityRow):
        if row in self.rows:
            self.rows.remove(row)
        row.destroy()

    def get_data(self) -> Dict[str, Any]:
        acts = [r.get_data() for r in self.rows]
        return {"activities": acts}


class AdminApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("BB Schedule — Admin")
        self.geometry("1380x780")
        self.minsize(1280, 640)
        apply_window_icon(self)

        self._display_win: Optional[DisplayWindow] = None
        self._day_sections: Dict[str, DaySection] = {}
        self._activity_types: List[str] = sort_activity_types(list(DEFAULT_ACTIVITY_TYPES))
        self._activity_combos: List[ctk.CTkComboBox] = []
        self._location_types: List[str] = sort_activity_types(list(DEFAULT_LOCATION_TYPES))
        self._location_combos: List[ctk.CTkComboBox] = []
        self._qr_code_filename = ""
        self._startup_days_from_file: Optional[Dict[str, Any]] = None

        self._build_ui()
        self._load_config_at_start()

    def get_activity_types(self) -> List[str]:
        return self._activity_types

    def get_location_types(self) -> List[str]:
        return self._location_types

    def register_activity_combo(self, cb: ctk.CTkComboBox) -> None:
        self._activity_combos.append(cb)

    def register_location_combo(self, cb: ctk.CTkComboBox) -> None:
        self._location_combos.append(cb)

    def refresh_all_activity_combos(self) -> None:
        vals = sort_activity_types(self._activity_types)
        combo_vals = ([""] + vals) if vals else [""]
        alive: List[ctk.CTkComboBox] = []
        for cb in self._activity_combos:
            try:
                if not cb.winfo_exists():
                    continue
                alive.append(cb)
                cur = cb.get()
                cb.configure(values=combo_vals)
                if cur in combo_vals:
                    cb.set(cur)
                elif combo_vals:
                    cb.set(combo_vals[0])
                else:
                    cb.set("")
            except tk.TclError:
                continue
        self._activity_combos = alive

    def refresh_all_location_combos(self) -> None:
        vals = sort_activity_types(self._location_types)
        combo_vals = ([""] + vals) if vals else [""]
        alive: List[ctk.CTkComboBox] = []
        for cb in self._location_combos:
            try:
                if not cb.winfo_exists():
                    continue
                alive.append(cb)
                cur = cb.get()
                cb.configure(values=combo_vals)
                if cur in combo_vals:
                    cb.set(cur)
                elif combo_vals:
                    cb.set(combo_vals[0])
                else:
                    cb.set("")
            except tk.TclError:
                continue
        self._location_combos = alive

    def _refresh_manage_activities_list(self) -> None:
        for w in self._manage_list_inner.winfo_children():
            w.destroy()
        for name in sort_activity_types(self._activity_types):
            row = ctk.CTkFrame(self._manage_list_inner)
            row.pack(fill="x", pady=4)
            ctk.CTkLabel(row, text=name, anchor="w").pack(side="left", fill="x", expand=True, padx=8, pady=6)
            if name in PROTECTED_ACTIVITY_TYPES:
                ctk.CTkLabel(row, text="(protected)", text_color=SUBTLE_TEXT_COLOR).pack(
                    side="right", padx=8, pady=6
                )
            else:
                ctk.CTkButton(
                    row,
                    text="✕",
                    width=36,
                    command=lambda n=name: self._remove_activity_type(n),
                ).pack(side="right", padx=6, pady=4)

    def _add_activity_type(self) -> None:
        name = self._new_activity_entry.get().strip()
        if not name:
            return
        if name in self._activity_types:
            themed_message(self, "Duplicate", "That activity type is already in the list.", kind="info")
            return
        if name in PROTECTED_ACTIVITY_TYPES:
            themed_message(self, "Built-in", "That name is already a built-in activity type.", kind="info")
            return
        self._activity_types.append(name)
        self._activity_types = sort_activity_types(self._activity_types)
        self._new_activity_entry.delete(0, "end")
        self._refresh_manage_activities_list()
        self.refresh_all_activity_combos()

    def _remove_activity_type(self, name: str) -> None:
        if name in PROTECTED_ACTIVITY_TYPES:
            return
        if name not in self._activity_types:
            return
        self._activity_types.remove(name)
        self._refresh_manage_activities_list()
        self.refresh_all_activity_combos()

    def _refresh_manage_locations_list(self) -> None:
        for w in self._manage_locations_inner.winfo_children():
            w.destroy()
        for name in sort_activity_types(self._location_types):
            row = ctk.CTkFrame(self._manage_locations_inner)
            row.pack(fill="x", pady=4)
            ctk.CTkLabel(row, text=name, anchor="w").pack(side="left", fill="x", expand=True, padx=8, pady=6)
            if name in PROTECTED_LOCATION_TYPES:
                ctk.CTkLabel(row, text="(protected)", text_color=SUBTLE_TEXT_COLOR).pack(
                    side="right", padx=8, pady=6
                )
            else:
                ctk.CTkButton(
                    row,
                    text="✕",
                    width=36,
                    command=lambda n=name: self._remove_location_type(n),
                ).pack(side="right", padx=6, pady=4)

    def _add_location_type(self) -> None:
        name = self._new_location_entry.get().strip()
        if not name:
            return
        if name in self._location_types:
            themed_message(self, "Duplicate", "That location is already in the list.", kind="info")
            return
        self._location_types.append(name)
        self._location_types = sort_activity_types(self._location_types)
        self._new_location_entry.delete(0, "end")
        self._refresh_manage_locations_list()
        self.refresh_all_location_combos()

    def _remove_location_type(self, name: str) -> None:
        if name in PROTECTED_LOCATION_TYPES:
            return
        if name not in self._location_types:
            return
        self._location_types.remove(name)
        self._refresh_manage_locations_list()
        self.refresh_all_location_combos()

    def _build_ui(self) -> None:
        save_bar = ctk.CTkFrame(self, fg_color="transparent")
        save_bar.pack(side="bottom", fill="x", padx=20, pady=(8, 16))
        save_row = ctk.CTkFrame(save_bar, fg_color="transparent")
        save_row.pack()
        ctk.CTkButton(save_row, text="Save", width=120, command=self._save).pack(
            side="left", padx=(0, 10)
        )
        ctk.CTkButton(save_row, text="Export HTML", width=140, command=self._export_html).pack(
            side="left", padx=(0, 10)
        )
        ctk.CTkButton(save_row, text="Push to Git", width=130, command=self._push_to_git).pack(
            side="left"
        )

        ctk.CTkFrame(self, height=1, fg_color=("gray70", "gray35")).pack(side="bottom", fill="x")

        title_bar = ctk.CTkFrame(self, fg_color="transparent")
        title_bar.pack(side="top", fill="x", padx=16, pady=(12, 4))
        ctk.CTkLabel(
            title_bar,
            text="BB Schedule",
            font=ctk.CTkFont(size=18, weight="bold"),
        ).pack(anchor="w")

        self._tabview = ctk.CTkTabview(self)
        self._tabview.pack(fill="both", expand=True, padx=12, pady=(4, 8))

        tab_home = self._tabview.add("Home")
        tab_activities = self._tabview.add("Activities")

        header = ctk.CTkFrame(tab_home, fg_color="transparent")
        header.pack(fill="x", padx=4, pady=(4, 8))

        ctk.CTkLabel(
            header,
            text="Event Date Range",
            font=ctk.CTkFont(size=14, weight="bold"),
        ).pack(anchor="w")

        rng = ctk.CTkFrame(header, fg_color="transparent")
        rng.pack(fill="x", pady=(8, 0))

        ctk.CTkLabel(rng, text="Start Date", text_color=SUBTLE_TEXT_COLOR).pack(side="left")
        self.start_cal = DateEntry(
            rng,
            width=12,
            background="#1f538d",
            foreground="#DCE4EE",
            borderwidth=0,
            date_pattern="yyyy-mm-dd",
        )
        self.start_cal.pack(side="left", padx=(8, 28))

        ctk.CTkLabel(rng, text="End Date", text_color=SUBTLE_TEXT_COLOR).pack(side="left")
        self.end_cal = DateEntry(
            rng,
            width=12,
            background="#1f538d",
            foreground="#DCE4EE",
            borderwidth=0,
            date_pattern="yyyy-mm-dd",
        )
        self.end_cal.pack(side="left", padx=(8, 0))

        self.start_cal.bind("<<DateEntrySelected>>", lambda e: self._rebuild_days())
        self.end_cal.bind("<<DateEntrySelected>>", lambda e: self._rebuild_days())

        btns = ctk.CTkFrame(tab_home, fg_color="transparent")
        btns.pack(fill="x", padx=4, pady=(0, 8))

        ctk.CTkButton(btns, text="Start Display", width=140, command=self._start_display).pack(
            side="left", padx=(0, 10)
        )
        ctk.CTkButton(btns, text="Stop Display", width=120, command=self._stop_display).pack(
            side="left"
        )

        ctk.CTkLabel(
            tab_home,
            text="Display QR code (bottom-left on fullscreen; sized for phone scanning)",
            text_color=SUBTLE_TEXT_COLOR,
            wraplength=520,
            justify="left",
        ).pack(anchor="w", padx=4, pady=(2, 4))
        qr_pick = ctk.CTkFrame(tab_home, fg_color="transparent")
        qr_pick.pack(fill="x", padx=4, pady=(0, 8))
        self._qr_lbl = ctk.CTkLabel(qr_pick, text="(none)", text_color=SUBTLE_TEXT_COLOR, anchor="w")
        self._qr_lbl.pack(side="left", padx=(0, 12))
        ctk.CTkButton(qr_pick, text="Choose QR image…", width=150, command=self._browse_qr_code).pack(
            side="left"
        )
        ctk.CTkButton(qr_pick, text="Clear", width=80, command=self._clear_qr_code).pack(
            side="left", padx=(8, 0)
        )

        self.days_parent = ctk.CTkScrollableFrame(tab_home, fg_color="transparent")
        self.days_parent.pack(fill="both", expand=True, padx=2, pady=(0, 6))

        add_type_row = ctk.CTkFrame(tab_activities, fg_color="transparent")
        add_type_row.pack(fill="x", padx=4, pady=(8, 4))
        ctk.CTkLabel(add_type_row, text="New activity type", text_color=SUBTLE_TEXT_COLOR).pack(
            anchor="w", pady=(0, 4)
        )
        entry_row = ctk.CTkFrame(tab_activities, fg_color="transparent")
        entry_row.pack(fill="x", padx=4, pady=(0, 8))
        self._new_activity_entry = ctk.CTkEntry(entry_row, width=320)
        self._new_activity_entry.pack(side="left", padx=(0, 10))
        ctk.CTkButton(entry_row, text="Add Activity Type", width=150, command=self._add_activity_type).pack(
            side="left"
        )

        ctk.CTkLabel(
            tab_activities,
            text="Activity types",
            font=ctk.CTkFont(size=14, weight="bold"),
        ).pack(anchor="w", padx=4, pady=(8, 6))
        self._manage_list_inner = ctk.CTkScrollableFrame(tab_activities, fg_color="transparent")
        self._manage_list_inner.pack(fill="both", expand=True, padx=4, pady=(0, 12))
        self._refresh_manage_activities_list()

        ctk.CTkFrame(tab_activities, height=1, fg_color=("gray70", "gray35")).pack(fill="x", padx=4, pady=(2, 10))

        ctk.CTkLabel(
            tab_activities,
            text="Location types",
            font=ctk.CTkFont(size=14, weight="bold"),
        ).pack(anchor="w", padx=4, pady=(0, 6))
        loc_row = ctk.CTkFrame(tab_activities, fg_color="transparent")
        loc_row.pack(fill="x", padx=4, pady=(0, 8))
        self._new_location_entry = ctk.CTkEntry(loc_row, width=320)
        self._new_location_entry.pack(side="left", padx=(0, 10))
        ctk.CTkButton(loc_row, text="Add Location Type", width=150, command=self._add_location_type).pack(side="left")

        self._manage_locations_inner = ctk.CTkScrollableFrame(tab_activities, fg_color="transparent", height=170)
        self._manage_locations_inner.pack(fill="x", padx=4, pady=(0, 12))
        self._refresh_manage_locations_list()

    def _get_config(self) -> Dict[str, Any]:
        return self._collect_config()

    def _collect_config(self) -> Dict[str, Any]:
        try:
            start_d = self.start_cal.get_date()
            end_d = self.end_cal.get_date()
        except Exception:
            start_d = end_d = date.today()
        out: Dict[str, Any] = {
            "event_start": start_d.isoformat(),
            "event_end": end_d.isoformat(),
            "activity_types": list(self._activity_types),
            "location_types": list(self._location_types),
            "qr_code_image": (self._qr_code_filename or "").strip(),
            "days": {},
        }
        for iso, sec in self._day_sections.items():
            out["days"][iso] = sec.get_data()
        return out

    def _rebuild_days(self):
        self._activity_combos.clear()
        self._location_combos.clear()
        for w in self.days_parent.winfo_children():
            w.destroy()
        self._day_sections.clear()

        try:
            start_d = self.start_cal.get_date()
            end_d = self.end_cal.get_date()
        except Exception:
            start_d = end_d = date.today()

        if end_d < start_d:
            themed_message(self, "Date range", "End date is before start date.", kind="warning")
            return

        for d in daterange_inclusive(start_d, end_d):
            sec = DaySection(self.days_parent, d, self)
            sec.pack(fill="x", padx=8)
            self._day_sections[d.isoformat()] = sec

    def _load_config_at_start(self) -> None:
        """Read config quickly; build day/activity rows on idle so the window appears first."""
        path = config_path()
        if not os.path.isfile(path):
            today = date.today()
            self.start_cal.set_date(today)
            self.end_cal.set_date(today)
            self._startup_days_from_file = {}
            self._refresh_manage_activities_list()
            self._refresh_manage_locations_list()
            self.after_idle(self._finish_heavy_startup)
            return
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
        except (OSError, json.JSONDecodeError) as e:
            themed_message(self, "Config", f"Could not load config: {e}", kind="error")
            today = date.today()
            self.start_cal.set_date(today)
            self.end_cal.set_date(today)
            self._startup_days_from_file = {}
            self._refresh_manage_activities_list()
            self._refresh_manage_locations_list()
            self.after_idle(self._finish_heavy_startup)
            return

        self._activity_types = normalize_activity_types_from_config(data.get("activity_types"))
        self._location_types = normalize_location_types_from_config(data.get("location_types"))
        self._qr_code_filename = (data.get("qr_code_image") or "").strip()
        self._update_qr_label()

        es = data.get("event_start") or ""
        ee = data.get("event_end") or ""
        try:
            if es:
                self.start_cal.set_date(date.fromisoformat(es))
            if ee:
                self.end_cal.set_date(date.fromisoformat(ee))
        except ValueError:
            pass

        self._startup_days_from_file = data.get("days") or {}
        self._refresh_manage_activities_list()
        self._refresh_manage_locations_list()
        self.after_idle(self._finish_heavy_startup)

    def _finish_heavy_startup(self) -> None:
        days = self._startup_days_from_file
        self._startup_days_from_file = None
        if days is None:
            return
        self._rebuild_days()
        if days:
            for iso, sec in self._day_sections.items():
                ddata = days.get(iso)
                if not ddata:
                    continue
                for a in normalize_day_activities(list(ddata.get("activities") or [])):
                    sec.add_row(a)
        self.refresh_all_activity_combos()
        self.refresh_all_location_combos()

    def _update_qr_label(self) -> None:
        fn = (self._qr_code_filename or "").strip()
        if fn:
            self._qr_lbl.configure(text=fn, text_color=LABEL_TEXT_COLOR)
        else:
            self._qr_lbl.configure(text="(none)", text_color=SUBTLE_TEXT_COLOR)

    def _browse_qr_code(self) -> None:
        p = filedialog.askopenfilename(
            title="QR code image",
            filetypes=[
                ("Images", "*.png *.jpg *.jpeg *.gif *.webp"),
                ("All", "*.*"),
            ],
        )
        if not p:
            return
        self._qr_code_filename = os.path.basename(p)
        dest = os.path.join(app_directory(), self._qr_code_filename)
        try:
            if os.path.abspath(p) != os.path.abspath(dest):
                shutil.copy2(p, dest)
        except OSError as e:
            themed_message(self, "Copy failed", str(e), kind="error")
            return
        self._update_qr_label()

    def _clear_qr_code(self) -> None:
        self._qr_code_filename = ""
        self._update_qr_label()

    def _save(self):
        path = config_path()
        cfg = self._collect_config()
        try:
            with open(path, "w", encoding="utf-8") as f:
                json.dump(cfg, f, indent=2)
        except OSError as e:
            themed_message(self, "Save failed", str(e), kind="error")
            return
        themed_message(self, "Saved", path, kind="success")

    def _export_html(self) -> None:
        cfg = self._collect_config()
        out_path = os.path.join(app_directory(), HTML_EXPORT_NAME)
        out_dir = os.path.dirname(out_path)
        try:
            map_name = ensure_export_assets(out_dir, app_directory())
            map_path = os.path.join(app_directory(), map_name)
            mw, mh = _map_image_pixel_size(map_path)
            html = build_event_status_html(cfg, map_image_name=map_name, map_image_width=mw, map_image_height=mh)
            with open(out_path, "w", encoding="utf-8") as f:
                f.write(html)
        except OSError as e:
            themed_message(self, "Export failed", str(e), kind="error")
            return
        themed_message(
            self,
            "Exported",
            out_path,
            kind="success",
            detail="Open this file in a browser. Live status uses the viewer's clock.",
        )

    def _push_to_git(self) -> None:
        repo = find_git_repo_root(app_directory())
        if not repo:
            themed_message(
                self,
                "Git error",
                "No git repository found in this app folder or its parent folders.",
                kind="error",
            )
            return

        # Save config first so latest app edits are included.
        path = os.path.join(repo, CONFIG_NAME)
        cfg = self._collect_config()
        try:
            with open(path, "w", encoding="utf-8") as f:
                json.dump(cfg, f, indent=2)
        except OSError as e:
            themed_message(self, "Save failed", str(e), kind="error")
            return

        def run_git(args: List[str], allow_fail: bool = False) -> subprocess.CompletedProcess[str]:
            cp = subprocess.run(
                ["git", "-C", repo] + args,
                check=False,
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
            )
            if not allow_fail and cp.returncode != 0:
                raise RuntimeError((cp.stderr or cp.stdout or "Git command failed").strip())
            return cp

        try:
            run_git(["add", "-A"])
            status = run_git(["status", "--porcelain"], allow_fail=False)
            committed = False
            if status.stdout.strip():
                msg = f"Update BB Schedule project ({datetime.now().strftime('%Y-%m-%d %H:%M')})"
                run_git(["commit", "-m", msg])
                committed = True

            push = run_git(["push", "origin", "main"], allow_fail=False)
            detail = "Changes committed and pushed." if committed else "No local changes to commit; pushed current main."
            out = (push.stdout or "").strip()
            if out:
                detail += f"\n\n{out}"
            themed_message(self, "Pushed", "Git push completed successfully.", kind="success", detail=detail)
        except RuntimeError as e:
            themed_message(self, "Git push failed", str(e), kind="error")

    def _start_display(self):
        if self._display_win is not None:
            try:
                if self._display_win.winfo_exists():
                    self._display_win.lift()
                    self._display_win.focus_force()
                    return
            except tk.TclError:
                self._display_win = None

        def clear_ref():
            self._display_win = None

        self._display_win = DisplayWindow(self, self._get_config, on_close=clear_ref)

    def _stop_display(self):
        if self._display_win is not None:
            try:
                self._display_win.close()
            except tk.TclError:
                pass
        self._display_win = None


def main():
    if DateEntry is None:
        root = tk.Tk()
        root.withdraw()
        themed_message(
            None,
            "Missing dependency",
            "tkcalendar is required. Install with: pip install tkcalendar",
            kind="error",
        )
        return
    ctk.set_appearance_mode("dark")
    ctk.set_default_color_theme("dark-blue")
    app = AdminApp()
    app.mainloop()


if __name__ == "__main__":
    main()
