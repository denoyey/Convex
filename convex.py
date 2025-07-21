# -*- coding: utf-8 -*-

# Author: DENOYEY
# Github: github.com/denoyey
# Date: 2025-07-21 (upgraded)
# Version: 2.0.1
# Description: Convert CSV⇄Excel/JSON, Excel⇄CSV, support .xls, multi-sheet, drag&drop, encoding-detect, recent-file history.
# License: MIT


import os, sys, json, time, fnmatch, readline, datetime, subprocess
import pandas as pd
import platform
from termcolor import colored
import chardet
try:
    import xlrd
except ImportError:
    xlrd = None

COLOR_EXIT = 'light_red'
COLOR_ALT = 'light_green'
COLOR_WARN = 'light_yellow'
COLOR_MENU = 'light_grey'
COLOR_HIGHLIGHT = 'bold'

ENGINE = 'openpyxl'
USE_INDEX = False
HISTORY_FILE = os.path.join(os.path.expanduser("~"), ".fileconvert_history.json")

def clear_screen():
    cmd = 'cls' if os.name=='nt' else 'clear'
    subprocess.run(cmd, shell=True, check=False)

def copyright():
    year = datetime.datetime.now().year
    text = f"Github: github.com/denoyey || © {year} All Rights Reserved."
    return colored(text, 'light_grey', attrs=['reverse'])

def logo():
    clear_screen()
    print(colored(f"""
██▄   ▄███▄      ▄   ████▄ ▀▄    ▄ ▄███▄ ▀▄    ▄ 
█  █  █▀   ▀      █  █   █   █  █  █▀   ▀  █  █  
█   █ ██▄▄    ██   █ █   █    ▀█   ██▄▄     ▀█   
█  █  █▄   ▄▀ █ █  █ ▀████    █    █▄   ▄▀  █    
███▀  ▀███▀   █  █ █        ▄▀     ▀███▀  ▄▀     
              █   ██                                                                            
{copyright()}
    """, 'light_grey', attrs=['bold']))

def is_safe_filename(name):
    return name.strip() and all(x not in name for x in ['..','/','\\','~','$',';'])

def is_file_readable(p): return os.path.isfile(p) and os.access(p, os.R_OK)
def save_history(path):
    try:
        h = []
        if os.path.exists(HISTORY_FILE):
            h = json.load(open(HISTORY_FILE))
        if path not in h:
            h.insert(0, path)
        json.dump(h[:5], open(HISTORY_FILE,'w'))
    except: pass

def show_history():
    if os.path.exists(HISTORY_FILE):
        h = json.load(open(HISTORY_FILE))
        if h:
            print(colored("\n🕘 Recent Files:", "cyan", attrs=['bold']))
            for i, f in enumerate(h,1): print(f"[{i}] {f}")

def detect_encoding(file_path):
    with open(file_path,'rb') as f:
        result = chardet.detect(f.read(10000))
    return result.get('encoding','utf-8')

def preview(df, max_rows=5):
    print(colored(f"\n🔍 Preview {min(len(df), max_rows)} rows:\n", 'magenta', attrs=['bold']))
    print(df.head(max_rows).to_string(), "\n")

def read_csv(path):
    enc = detect_encoding(path)
    return pd.read_csv(path, encoding=enc)

def validate_preview_csv(p):
    if not is_file_readable(p):
        print(colored(f"File CSV '{p}' tidak ada!", "red")); return None
    try:
        df = read_csv(p); preview(df); return df
    except Exception as e:
        print(colored(f"Gagal baca CSV: {e}", "red")); return None

def validate_preview_json(p):
    if not is_file_readable(p):
        print(colored(f"File JSON '{p}' tidak ada!", "red")); return None
    try:
        data = json.load(open(p))
        df = pd.json_normalize(data); preview(df); return df
    except Exception as e:
        print(colored(f"Gagal baca JSON: {e}", "red")); return None

def validate_preview_excel(p, sheet=None):
    if not is_file_readable(p):
        print(colored(f"File Excel '{p}' tidak ada!", "red")); return None
    try:
        eng = 'xlrd' if p.lower().endswith('.xls') else ENGINE
        df = pd.read_excel(p, sheet_name=sheet, engine=eng)
        preview(df); return df
    except Exception as e:
        print(colored(f"Gagal baca Excel: {e}", "red")); return None

def select_output_folder():
    print(colored("\n[1] Folder saat ini\n[2] Folder lain", 'light_grey'))
    if input(colored("Pilih: ", 'light_grey'))=='2':
        path = input(colored("Masukkan path: ", 'light_grey')).strip()
        os.makedirs(path, exist_ok=True)
        return path
    return os.getcwd()

def batch(folder, ext, func, out_folder, **k):
    files = [f for f in os.listdir(folder) if f.lower().endswith(ext.lower())]
    if not files: print(colored(f"Tidak ada *{ext} di {folder}", "yellow")); return
    for f in files:
        inp = os.path.join(folder,f)
        base = os.path.splitext(f)[0] + k.get('output_ext','')
        out = os.path.join(out_folder, base)
        func(inp, out, **k); save_history(inp)

def csv2excel(inp, out, preview=False):
    df = read_csv(inp) if preview else pd.read_csv(detect_encoding(inp))
    if preview: preview(df)
    df.to_excel(out, index=USE_INDEX, engine=ENGINE)
    print(colored(f"Converted ✅ {out}", "green"))
    save_history(inp)

def json2csv(inp, out, preview=False):
    data = json.load(open(inp))
    df = pd.json_normalize(data)
    if preview: preview(df)
    df.to_csv(out, index=USE_INDEX)
    print(colored(f"Converted ✅ {out}", "green"))
    save_history(inp)

def excel2csv(inp, out, sheet=None, preview=False):
    eng = 'xlrd' if inp.lower().endswith('.xls') else ENGINE
    df = pd.read_excel(inp, sheet_name=sheet, engine=eng)
    if preview: preview(df)
    df.to_csv(out, index=USE_INDEX)
    print(colored(f"Converted ✅ {out}", "green"))
    save_history(inp)

def excel_multi(inp, out_folder):
    eng = 'xlrd' if inp.lower().endswith('.xls') else ENGINE
    xls = pd.ExcelFile(inp, engine=eng)
    for sheet in xls.sheet_names:
        df = xls.parse(sheet)
        fn = os.path.join(out_folder, f"{os.path.splitext(os.path.basename(inp))[0]}_{sheet}.csv")
        df.to_csv(fn, index=USE_INDEX)
        print(colored(f"Sheet '{sheet}' → {fn}", "green"))
        save_history(inp)

def csv2json(inp, out, preview=False):
    df = pd.read_csv(inp, encoding=detect_encoding(inp))
    if preview: preview(df)
    json.dump(df.to_dict(orient='records'), open(out,'w'), indent=4)
    print(colored(f"Converted ✅ {out}", "green"))
    save_history(inp)

def input_drag(text):
    raw = input(colored(text, 'light_grey')).strip().strip('"')
    return ('folder',raw) if os.path.isdir(raw) else ('file',raw)

def setup_tab_completion():
    import rlcompleter

    def complete_path(text, state):
        line = readline.get_line_buffer().split()
        if not line:
            return [c + os.sep if os.path.isdir(c) else c
                    for c in os.listdir('.')][state]
        else:
            path = os.path.expanduser(text)
            dirname = os.path.dirname(path) or '.'
            matches = []
            try:
                entries = os.listdir(dirname)
                matches = [os.path.join(dirname, e) for e in entries if e.startswith(os.path.basename(path))]
                matches = [m + os.sep if os.path.isdir(m) else m for m in matches]
            except Exception:
                pass
            return matches[state] if state < len(matches) else None

    readline.set_completer_delims(' \t\n;')
    readline.parse_and_bind("tab: complete")
    readline.set_completer(complete_path)


def main():
    if platform.system() != 'Windows':
        setup_tab_completion()

    while True:
        try:
            logo()
            show_history()
            print(colored("\n[ === MENU CONVERT === ]\n", 'light_grey', attrs=['bold']))
            print(colored("[1] CSV → Excel", 'light_green', attrs=['bold']))
            print(colored("[2] JSON → CSV", 'light_yellow', attrs=['bold']))
            print(colored("[3] Excel → CSV", 'light_green', attrs=['bold']))
            print(colored("[4] Excel Multi-Sheet → CSV", 'light_yellow', attrs=['bold']))
            print(colored("[5] CSV → JSON", 'light_green', attrs=['bold']))
            print(colored("[0] Exit", 'light_red', attrs=['bold']))

            c = input(colored("\n✨ Pilih Menu: ", "cyan", attrs=['bold'])).strip()
            if c == '0':
                clear_screen(); logo()
                print(colored("\nTerima kasih telah menggunakan alat ini. 👋\n", 'light_red', attrs=['bold']))
                break

            mapping = {
                '1': ('.csv', '.xlsx', csv2excel),
                '2': ('.json', '.csv', json2csv),
                '3': ('.xls;.xlsx', '.csv', excel2csv),
                '5': ('.csv', '.json', csv2json)
            }

            if c == '4':
                inp = input_drag(colored("\n📂 Drop file .xls/.xlsx → ", 'light_grey', attrs=['bold']))
                out = select_output_folder()
                excel_multi(inp[1], out)
                time.sleep(1.5)
                continue

            if c not in mapping:
                print(colored("\n❗ Input tidak valid. Coba lagi.\n", 'light_red', attrs=['bold']))
                time.sleep(1.5)
                continue

            ext_in, ext_out, func = mapping[c]
            typ, inp = input_drag(f"\n📁 Drop file/folder {ext_in} → ")
            outf = select_output_folder()

            if typ == 'folder':
                batch(inp, ext_in.split(';')[0], func, outf, output_ext=ext_out, preview=False)
            else:
                if not is_file_readable(inp):
                    print(colored("❌ File tidak ditemukan atau tidak bisa dibaca!", "red", attrs=['bold']))
                    time.sleep(2)
                    continue

                if c == '1':
                    df = validate_preview_csv(inp)
                elif c == '2':
                    df = validate_preview_json(inp)
                elif c == '3':
                    sheet = None
                    df = validate_preview_excel(inp, sheet)
                elif c == '5':
                    df = validate_preview_csv(inp)

                if df is None:
                    time.sleep(1.5)
                    continue

                outname = input(colored("📝 Nama file output (tanpa ekstensi): ", "cyan", attrs=['bold'])).strip()
                if not is_safe_filename(outname):
                    print(colored("❗ Nama file output tidak valid!", "red", attrs=['bold']))
                    time.sleep(1.5)
                    continue

                out_path = os.path.join(outf, outname + ext_out)
                if c == '3':
                    sheet = None
                    func(inp, out_path, sheet=sheet, preview=True)
                else:
                    func(inp, out_path, preview=True)

            time.sleep(1.5)

        except KeyboardInterrupt:
            print(colored("\n\n[!] Program dibatalkan. Keluar...\n", "red", attrs=['bold']))
            sys.exit(0)


if __name__=='__main__':
    try:
        main()
    except Exception as e:
        print(colored(f"Terjadi kesalahan: {e}", "red", attrs=['bold']))
        sys.exit(0)