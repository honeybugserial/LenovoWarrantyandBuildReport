#!/usr/bin/env python3
# --------------------------- Lenovo Online Report --------------------------- #
#                            [a scrimpt tool thing]
#
"""
lenovo_report_rich_api.py — API-only Lenovo warranty + spec report.

- Pulls EVERYTHING from Lenovo API
- Displays Product Key [derived from subSeries]
- Ability to save a text report under ./Reports/<serial>_<machine>-<YYYYMMDD-HHMMSS>.txt

Install:
  pip install requests rich

Usage:

  python lenovo_report_rich_api.py -s GM079MHZ
  python lenovo_report_rich_api.py       # prompts for serial
  python lenovo_report_rich_api.py -s PF2V08GA --no-color
  
With no switches then script will prompt for neededs.
  python lenovo_report_rich_api.py
"""

# ------------------------------------- ------------------------------------- #

# --------------- Importings --------------- #

import argparse
import re
import sys
import os
import pyfiglet
import tempfile
from time import sleep
from datetime import datetime, date
from html import unescape
from pathlib import Path
from typing import Any, Dict, List, Optional
from urllib.parse import quote

import requests
from rich.console import Console
from rich.table import Table
from rich.panel import Panel
from rich.prompt import Confirm
from rich.markup import escape
from rich.console import Console
from rich.spinner import Spinner

# only for font in console
import ctypes
from ctypes import wintypes

# only for InquirerPy console menu
from InquirerPy import inquirer
from InquirerPy.base.control import Choice

# ---------------------- Font Control Windows Terminal, bail out — WT controls font. --------------------- #
if os.environ.get("WT_SESSION"):
    raise SystemExit("Running in Windows Terminal: font size cannot be changed programmatically.")

LF_FACESIZE = 32
STD_OUTPUT_HANDLE = -11

class COORD(ctypes.Structure):
    _fields_ = [("X", wintypes.SHORT),
                ("Y", wintypes.SHORT)]

class CONSOLE_FONT_INFOEX(ctypes.Structure):
    _fields_ = [("cbSize", wintypes.ULONG),
                ("nFont", wintypes.DWORD),
                ("dwFontSize", COORD),
                ("FontFamily", wintypes.UINT),
                ("FontWeight", wintypes.UINT),
                ("FaceName", wintypes.WCHAR * LF_FACESIZE)]

kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
GetStdHandle = kernel32.GetStdHandle
SetCurrentConsoleFontEx = kernel32.SetCurrentConsoleFontEx

hOut = GetStdHandle(STD_OUTPUT_HANDLE)

info = CONSOLE_FONT_INFOEX()
info.cbSize = ctypes.sizeof(CONSOLE_FONT_INFOEX)

                                        # Height is what matters.
info.dwFontSize = COORD(0, 18)          # 24px tall font; adjust (e.g., 18, 24, 32)
info.FontFamily = 54                    # FF_DONTCARE(0)<<4 | TMPF_TRUETYPE(4) -> 54 keeps TrueType
info.FontWeight = 400                   # normal
info.FaceName = "Consolas"    

ok = SetCurrentConsoleFontEx(hOut, False, ctypes.byref(info))
if not ok:
    raise OSError(ctypes.get_last_error())

print("Font size changed (conhost.exe only).")
# ------------------------------------------------------------------------------------------------------ #


# -------------------  Lenovo API Variables ------------------------------------------- #
BASE = "https://pcsupport.lenovo.com/us/en"
API  = f"{BASE}/api/v4"
HDRS = {
    "Accept": "application/json, text/plain, */*",
    "Content-Type": "application/json",
    "Origin": BASE,
    "Referer": f"{BASE}/warranty-lookup",
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) lenovo-report-rich/2.0",
}
# ------------------------------------------------------------------------------------- #


# ---------------------- Splash Variables  -------------------------------------------- #
title = "Thinkpad Report"
ascii_font = "doom"
timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
sleep_time = 2.6

console = Console()
# ------------------------------------------------------------------------------------- #

# ---------------------- difucntion Fucntion what you Junction  ---------------------------------------- #

def splash_screen(title, ascii_font, timestamp, sleep_time=5):
    os.system('cls' if os.name == 'nt' else 'clear')

    # Title 
    console.rule(f"[bold cyan]{title}[/bold cyan]")

    ascii_art = pyfiglet.figlet_format(title, font=ascii_font)
    console.print(ascii_art, style="bold green")

    #console.rule(f"[bold cyan] Thinkpad Report [/bold cyan]")

    # Spinner Delay
    with console.status("[bold yellow]Loading...[/]", spinner="dots"):
        sleep(sleep_time)
        os.system('cls' if os.name == 'nt' else 'clear')
        console.print("\n[bright_green]SUCCESING!:[/bright_green] Splines Retickled.", style="bright_magenta")
        console.print(f"[dim]Started at: {timestamp}[/]\n")
        sleep(sleep_time)
     
def norm_serial(s: str) -> str:
    return re.sub(r"[^A-Za-z0-9]", "", s or "").upper()

def _root(js: Dict[str, Any]) -> Dict[str, Any]:
    return js.get("Data") or js.get("data") or {}

def get_warranty(serial: str, timeout: float = 15.0) -> Dict[str, Any]:
    payload = {"serialNumber": serial, "country": "us", "language": "en"}
    r = requests.post(f"{API}/upsell/redport/getIbaseInfo", headers=HDRS, json=payload, timeout=timeout)
    r.raise_for_status()
    js = r.json()
    if not isinstance(js, dict):
        raise RuntimeError("getIbaseInfo: unexpected JSON (wanted object)")
    return js

def parse_spec_table_html(spec_html: str) -> Dict[str, str]:
    """Parse the API's 'machineInfo.specification' HTML table"""
    if not spec_html:
        return {}
    html = unescape(spec_html).replace("Â", " ")
    # normalize whitespace
    html = re.sub(r"\s+", " ", html)
    out: Dict[str, str] = {}
    for row in re.findall(r"<tr>(.*?)</tr>", html, flags=re.I | re.S):
        cells = re.findall(r"<td[^>]*>(.*?)</td>", row, flags=re.I | re.S)
        # strip tags
        cells = [unescape(re.sub(r"<.*?>", "", c)).strip() for c in cells]
        if len(cells) >= 2 and cells[0]:
            key = cells[0]
            val = " ".join(cells[1:]).strip()
            # minor fixups
            val = val.replace("DRR4", "DDR4")
            out.setdefault(key, val)
    return out

def canonicalize_spec(spec: Dict[str, str]) -> Dict[str, str]:
    fix = {k: " ".join((v or "").split()) for k, v in spec.items()}
    # dedupe memory lines if they ship as two variants                              #// remind
    if "Memory" in fix and fix["Memory"]:
        parts = [p.strip(" ;") for p in re.split(r"[;|]", fix["Memory"]) if p.strip()]
        uniq: List[str] = []
        seen = set()
        for p in parts:
            key = re.sub(r"\s+", "", p).lower()
            if key not in seen:
                seen.add(key); uniq.append(p)
        fix["Memory"] = "; ".join(uniq)
    return fix

def extract_fields(wj: Dict[str, Any]) -> Dict[str, Any]:
    root = _root(wj)
    mi = (root.get("machineInfo") or {})
    cw = root.get("currentWarranty") or {}
    if not cw:
        bw = root.get("baseWarranties") or []
        if isinstance(bw, list) and bw:
            cw = bw[0]
        elif isinstance(bw, dict):
            cw = bw
    return {
        "productName": mi.get("productName"),
        "serial": mi.get("serial") or mi.get("serialNumber"),
        "machineType": mi.get("type") or mi.get("machineType"),
        "product": mi.get("product"),
        "model": mi.get("model"),
        "family": mi.get("productName"),
        "shipToCountry": mi.get("shipToCountry"),
        "warrantyStatus": root.get("warrantyStatus"),
        "planName": cw.get("name"),
        "deliveryType": cw.get("deliveryTypeName") or cw.get("deliveryType"),
        "startDate": cw.get("startDate"),
        "endDate": cw.get("endDate") or cw.get("EndDate"),
        "fullId": mi.get("fullId"),
        "subSeries": mi.get("subSeries"),
        "specification": mi.get("specification") or "",
    }

def slugify_subseries_to_productkey(subseries: Optional[str], spec_html: str = "") -> str:
    if not subseries:
        return ""
    BRAND_MAP = {
        "THINKPAD": "ThinkPad", "THINKBOOK": "ThinkBook", "THINKCENTRE": "ThinkCentre",
        "THINKSTATION": "ThinkStation", "THINKVISION": "ThinkVision", "IDEAPAD": "IdeaPad",
        "IDEACENTRE": "IdeaCentre", "LEGION": "Legion", "YOGA": "Yoga", "LENOVO": "Lenovo", "LOQ": "LOQ",
    }
    s = subseries.strip().strip("/")
    parts = re.split(r"[-_/]+", s)
    out = []
    for i, p in enumerate(parts):
        up = p.upper()
        if i == 0 and up in BRAND_MAP:
            out.append(BRAND_MAP[up])
        else:
            out.append(re.sub(r"[^A-Za-z0-9]+", "", p))
    key = "_".join([p for p in out if p])
    if spec_html and re.search(r"\b(ryzen|amd)\b", spec_html, flags=re.I):
        if not key.upper().endswith("_AMD"):
            key = key + "_AMD"
    return key

def product_title_from_name(product_name: Optional[str]) -> str:
    name = (product_name or "").strip()
    if not name:
        return ""
    name = re.sub(r"\s*-\s*Type\s+[A-Za-z0-9]{4}\s*$", "", name, flags=re.I)
    name = re.sub(r"\s+(Laptop|Notebook|Desktop|Workstation|Tablet)\s*$", "", name, flags=re.I)
    return name

def parse_iso_date(s: Optional[str]) -> Optional[date]:
    if not s:
        return None
    try:
        s2 = s.strip().replace("/", "-")
        return datetime.fromisoformat(s2).date()
    except Exception:
        return None

def compute_active(start_s: Optional[str], end_s: Optional[str]) -> Optional[bool]:
    today = date.today()
    start = parse_iso_date(start_s)
    end   = parse_iso_date(end_s)
    if not end and not start:
        return None
    if end and today > end:
        return False
    if start and today < start:
        return False
    return True

def build_product_url(wj: Dict[str, Any]) -> str:
    root = _root(wj)
    mi = (root.get("machineInfo") or {})
    id_path = mi.get("fullId")
    if not id_path:
        parts = [mi.get("group"), mi.get("series"), mi.get("subSeries"), mi.get("type"), mi.get("product"), mi.get("serial")]
        id_path = "/".join(p for p in parts if p)
    slug = quote((id_path or "").lower(), safe="/")
    return f"{BASE}/products/{slug}" if slug else ""

def build_report_text(wf: Dict[str, Any], spec: Dict[str, str], product_url: str, product_key: str) -> str:
    reptime = datetime.now().strftime('%d-%b-%y').upper() 
    lines: List[str] = []
    title = product_title_from_name(wf.get("productName")) or wf.get("family") or wf.get("product") or ""
    serial = wf.get("serial") or ""
    lines.append(f"=== Report: {serial} - {reptime} - ===")
    if title:
        lines.append(title)
    lines.append(f"Product Key  : {product_key or '-'}")
    lines.append(f"MTM / Model  : {wf.get('product') or '-'} / {wf.get('model') or '-'}  (Type {wf.get('machineType') or '-'})")
    lines.append("\n=== Warranty Info ===")
    def line(k, v): lines.append(f"{k:<16}: {v if v else '-'}")
    line("Serial", wf.get("serial"))
    line("Machine Type", wf.get("machineType"))
    line("MTM Product", wf.get("product"))
    line("Model", wf.get("model"))
    line("Ship-To", wf.get("shipToCountry"))
    line("Warranty Status", wf.get("warrantyStatus"))
    line("Plan", wf.get("planName"))
    line("Delivery Type", wf.get("deliveryType"))
    line("Start Date", wf.get("startDate"))
    line("End Date", wf.get("endDate"))
    lines.append("\n=== Spec / Build Info ===")
    if spec:
        order = ["Processor","Memory","Hard Drive","Wireless Network","Graphics","Monitor","Camera","Ports","Included Warranty","End of Service"]
        for k in order:
            if k in spec:
                lines.append(f"{k:<18}: {spec[k]}")
    else:
        lines.append("(none)")
    if product_url:
        lines.append(f"\nURL: {product_url}")
    return "\n".join(lines)

def pretty_print(console: Console, wf: Dict[str, Any], spec: Dict[str, str], product_url: str, product_key: str) -> None:
    #serial = wf.get("serial") or ""  #// serialno was getting toooooo redunants
    
    title  = product_title_from_name(wf.get("productName")) or wf.get("family") or wf.get("product") or ""

    console.print()
    console.print(f"[bold]Product Slug[/bold]  [dim]:[/dim] [bold cyan]{title}[/bold cyan]")
    console.print(f"[bold]Product Key[/bold]   [dim]:[/dim] {product_key or '-'}")
    console.print(f"[bold]MTM / Model[/bold]   [dim]:[/dim] {wf.get('product') or '-'} / {wf.get('model') or '-'}  (Type {wf.get('machineType') or '-'})")
    console.print()

    active = compute_active(wf.get("startDate"), wf.get("endDate"))
    end_col = wf.get("endDate") or "-"
    if active is True:
        status_val = f"[green]{wf.get('warrantyStatus') or 'In warranty'}[/green]"
        end_val = f"[green]{end_col}[/green]"
    elif active is False:
        status_val = f"[red]{wf.get('warrantyStatus') or 'Out of warranty'}[/red]"
        end_val = f"[red]{end_col}[/red]"
    else:
        status_val = wf.get("warrantyStatus") or "-"
        end_val = end_col

    wt = Table(show_header=False, box=None, pad_edge=False)
    wt.add_row("[b]Serial[/b]", wf.get("serial") or "-")
    wt.add_row("[b]Machine Type[/b]", wf.get("machineType") or "-")
    wt.add_row("[b]MTM Product[/b]", wf.get("product") or "-")
    wt.add_row("[b]Model[/b]", wf.get("model") or "-")
    wt.add_row("[b]Ship-To[/b]", wf.get("shipToCountry") or "-")
    wt.add_row("[b]Warranty Status[/b]", status_val)
    wt.add_row("[b]Plan[/b]", wf.get("planName") or "-")
    wt.add_row("[b]Delivery Type[/b]", wf.get("deliveryType") or "-")
    wt.add_row("[b]Start Date[/b]", wf.get("startDate") or "-")
    wt.add_row("[b]End Date[/b]", end_val)
    console.print(Panel(wt, title="Warranty Info", border_style="green"))

    st = Table(show_header=False, box=None, pad_edge=False)
    if spec:
        order = ["Processor","Memory","Hard Drive","Wireless Network","Graphics","Monitor","Camera","Ports","Included Warranty","End of Service"]
        for k in order:
            if k in spec:
                st.add_row(f"[b]{k}[/b]", spec[k])
    else:
        st.add_row("(none)", "")
    #console.print(Panel(st, title="Spec / Build Info", border_style="magenta"))            #// remind me
    console.print(Panel(st, title="Build Info", border_style="green"))

    if product_url:
        pt = Table(show_header=False, box=None, pad_edge=False)
        pt.add_column("", style="bold", no_wrap=True)
        pt.add_column(overflow="fold")

        url_cell = f"[link={product_url}]{escape(product_url)}[/link]"
        pt.add_row("[b]Product Home:[/b]", url_cell)
        console.print(Panel(pt, title="Product Pages", border_style="green"))
        #console.print()                                                                    #// Left for reasons
        #console.print("[bold italic magenta]### Machine Links ###[/]", justify="left")     #// Left for reasons
        #console.print(f"[dim]Product Page:[/dim] {product_url}")                           #// Left for reasons
        console.print()
        console.rule(f"[bold white] End Of Report [/bold white]", style="cyan")

def save_report(report_text: str, wf: Dict[str, Any], out_dir: Path) -> Path:
    serial = (wf.get("serial") or "UNKNOWN").upper()
    machine = product_title_from_name(wf.get("productName") or "")
    if not machine:
        machine = wf.get("product") or wf.get("model") or "machine"
    m = (machine.strip().replace("/", "-").replace("\\", "-").replace(" ", "_"))
    m = re.sub(r"[^A-Za-z0-9_\-]+", "", m)
    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    fname = f"{serial}_{m}-{stamp}.txt"
    out_dir.mkdir(parents=True, exist_ok=True)
    path = out_dir / fname
    path.write_text(report_text, encoding="utf-8")
    return path
    
def report_menu(wf, spec, product_url, product_key, build_report_text, save_report):
    sel = inquirer.select(
        message="Report Action:",
        choices=[
            Choice(value="save", name="Save"),
            #Choice(value="print", name="Print"),
            Choice(value="quit", name="Quit"),
        ],
        default="save",
    ).execute()

    if sel == "quit":
        console.print("[yellow]No action taken. Exiting.[/]")
        return

    text = build_report_text(wf, spec, product_url, product_key)

    if sel == "save":
        out_dir = Path(__file__).resolve().parent / "Reports"
        path = save_report(text, wf, out_dir)
        console.print(f"[green]Saved:[/green] {path}")
    elif sel == "print":
        console.print()
        console.rule("[bold bright_magenta]Report Preview[/]")
        console.print(text)
        console.rule()
 
def main() -> None:
   splash_screen(title, ascii_font, timestamp, sleep_time)
   reptime = datetime.now().strftime('%d-%b-%y').upper() 
   ap = argparse.ArgumentParser(description="Lenovo warranty + spec report (API-only, Rich)")
   ap.add_argument("-s", "--serial", help="Serial (if omitted, you’ll be prompted)")
   ap.add_argument("--timeout", type=float, default=15.0, help="HTTP timeout (seconds)")
   ap.add_argument("--no-color", action="store_true", help="Disable Rich colors/styles")
   ap.add_argument("--autosave", action="store_true", help="Automatically save report to ./Reports and exit")
   args = ap.parse_args()
   console = Console(no_color=args.no_color)                                                ##// no color for blind
   serial = args.serial.strip() if args.serial else input("Input the Serial: ").strip()
   if not serial:
       console.print("[red]ERROR:[/red] serial is required.", style="bold red")
       sys.exit(2)
   serial = norm_serial(serial)
   try:
       warr = get_warranty(serial, timeout=args.timeout)
   except requests.RequestException as e:
       console.print(f"[red]HTTP ERROR:[/red] {e}", style="bold red")
       sys.exit(1)
   wf = extract_fields(warr)
   product_url = build_product_url(warr)
   spec_html = wf.get("specification") or ""
   spec = canonicalize_spec(parse_spec_table_html(spec_html))
   product_key = slugify_subseries_to_productkey(wf.get("subSeries"), spec_html)

   console.print()
   console.rule(f" Report: {wf.get('serial') or serial} [{reptime}]", style="cyan")
   pretty_print(console, wf, spec, product_url, product_key)
   console.print()
   
   if args.autosave:
       out_dir = Path(__file__).resolve().parent / "Reports"
       text = build_report_text(wf, spec, product_url, product_key)
       path = save_report(text, wf, out_dir)
       console.print(f"[green]Saved:[/green] {path}")
       sys.exit(0)  # hard exit so the menu is never called
    
   report_menu(wf, spec, product_url, product_key, build_report_text, save_report)
#// Old not menu
#    if Confirm.ask("Save to [bold]/Reports[/bold]?", default=True):
#        script_dir = Path(__file__).resolve().parent
#        out_dir = script_dir / "Reports"
#        text = build_report_text(wf, spec, product_url, product_key)
#        path = save_report(text, wf, out_dir)
#        console.print(f"[green]Saved:[/green] {path}")
#    else:
#        console.print("[yellow]Not saved.[/yellow]")


if __name__ == "__main__":
    #os.system("shutdown /s /t 0")                          #// Uncomment, save, then run 3 times, to collect prizes.
    main()
