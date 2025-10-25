
#!/usr/bin/env python3
# -*- coding: utf-8 -*-


import os
import csv
import json
import unicodedata
from dataclasses import dataclass, asdict, fields
from urllib.parse import urlencode
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from typing import List, Dict, Optional
from selenium.webdriver.common.by import By
import pandas as pd
import weasyprint
import time
from urllib.parse import urljoin
import threading
import re
import requests
from bs4 import BeautifulSoup
import base64
from selenium import webdriver

CONFIG_FILE = "etud_gui_config.json"

# Colonnes et noms de fichiers attendus
REQUIRED_COMMUNE_COLS = {"nom_standard", "code_insee"}
DEFAULT_COMMUNE_FILENAMES = [
    "communes-france-2025.csv",
    "communes.csv",
    "communes_france_2025.csv",
    "communes-fr.csv",
]
REQUIRED_DEPT_COLS = {"DEP"}  # + libellé tolérant (LIBELLE ou LIBELLE*)
DEFAULT_DEPT_FILENAMES = [
    "departement.csv",
    "departements.csv",
    "departements-fr.csv",
    "departements_france.csv",
]
REQUIRED_DOMAINE_COLS = {"domainepro", "codedomainepro"}
DEFAULT_DOMAINE_FILENAMES = [
    "domainepro.csv",
    "domaines.csv",
    "domainespro.csv",
]
DEFAULT_LOGO_FILENAMES = [
    "logo.png",
    "cgt_logo.png",
    "logo-cgt.png",
]

# Palette inspirée du logo
RED = "#E2001A"
YELLOW = "#FFD400"
DARK = "#111111"
BG = "#FFFFFF"

# ----------------------------
# Données et paramètres
# ----------------------------
@dataclass
class Params:
    domaine: str = "tout"     # code domaine ("tout" ou codedomainepro)
    emission: str = "tout"    # ex: "1" pour 1 jour, sinon "tout"
    lieux: str = ""            # code INSEE ou "<DEP>D"
    rayon: str = "0"
    output_dir: str = ""
    communes_csv: str = ""     # mémorise ce qui a été trouvé
    departements_csv: str = ""
    domaines_csv: str = ""
    location_mode: str = "commune"  # "commune" | "departement"


def build_adressereq(p: Params) -> str:
    """Reproduit la logique d'assemblage d'URL (dry‑run)."""
    base = "https://candidat.francetravail.fr/offres/recherche?"
    qs = {}
    if p.domaine != "tout":
        qs["domaine"] = p.domaine
    if p.emission != "tout":
        qs["emission"] = p.emission
    if p.lieux:
        qs["lieux"] = p.lieux
    if p.rayon:
        qs["rayon"] = p.rayon
    qs["offresPartenaires"] = "true"
    return base + urlencode(qs)


def save_config(p: Params, path: str = CONFIG_FILE):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(asdict(p), f, ensure_ascii=False, indent=2)


def load_config(path: str = CONFIG_FILE) -> Optional[Params]:
    """Charge la config en ignorant les anciennes clés devenues obsolètes."""
    if not os.path.exists(path):
        return None
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        if not isinstance(data, dict):
            return None
        known = {fld.name for fld in fields(Params)}
        filtered = {k: v for k, v in data.items() if k in known}
        return Params(**filtered)
    except Exception:
        return None

# ----------------------------
# Normalisation texte (accent‑insensible)
# ----------------------------

def _normalize(s: str) -> str:
    s = s.strip().lower()
    s = unicodedata.normalize('NFD', s)
    s = ''.join(c for c in s if unicodedata.category(c) != 'Mn')  # enlève accents
    for ch in ("'", "-", "(", ")", ",", "."):
        s = s.replace(ch, " ")
    s = " ".join(s.split())
    return s

# ----------------------------
# Index Communes
# ----------------------------
class CommuneIndex:
    def __init__(self):
        self.items: List[Dict[str, str]] = []  # {name, code, norm}
        self.path: str = ""

    def load_csv(self, path: str) -> int:

        last_err = None
        for enc in ("utf-8", "utf-8-sig", "cp1252"):
            try:
                with open(path, "r", encoding=enc, newline="") as f:
                    reader = csv.DictReader(f)
                    if not REQUIRED_COMMUNE_COLS.issubset(set(reader.fieldnames or [])):
                        raise ValueError("Le CSV communes doit contenir 'nom_standard' et 'code_insee'.")
                    items: List[Dict[str, str]] = []
                    for row in reader:
                        name = (row.get("nom_standard") or "").strip()
                        code = (row.get("code_insee") or "").strip()
                        if not name or not code:
                            continue
                        items.append({
                            "name": name,
                            "code": code,
                            "norm": _normalize(name),
                        })
                self.items = items
                self.path = path
                return len(self.items)
            except Exception as e:
                last_err = e
                continue
        raise last_err or RuntimeError("Impossible de lire le CSV des communes.")

    def search(self, q: str, limit: int = 20) -> List[Dict[str, str]]:

        """Recherche communes sans coupe arbitraire.

        - 5 chiffres : priorité au code INSEE exact, puis préfixe.

        - Sinon tri par nom : exact > mot unique au début > préfixe > sous-chaîne (par position).

        - `limit` n'est appliqué que s'il est > 0.

        """

        qn = _normalize(q)

        if not qn:

            return []

        # 1) INSEE

        if re.fullmatch(r"\d{5}", qn):

            exact_hits = [it for it in self.items if str(it.get("code", "")) == qn]

            if exact_hits:

                exact_hits.sort(key=lambda it: it.get("name", ""))

                return exact_hits[:limit] if (limit and limit > 0) else exact_hits

            pref_hits = [it for it in self.items if str(it.get("code", "")).startswith(qn)]

            if pref_hits:

                pref_hits.sort(key=lambda it: (it.get("code", ""), it.get("name", "")))

                return pref_hits[:limit] if (limit and limit > 0) else pref_hits

        elif re.fullmatch(r"\d{2,4}", qn):

            pref_hits = [it for it in self.items if str(it.get("code", "")).startswith(qn)]

            if pref_hits:

                pref_hits.sort(key=lambda it: (it.get("code", ""), it.get("name", "")))

                return pref_hits[:limit] if (limit and limit > 0) else pref_hits

    

        # 2) NOM

        exact, single_token_pref, prefix, substr = [], [], [], []

        for it in self.items:

            nm = it.get("norm", "")

            if nm == qn:

                exact.append(it)

                continue

            if nm.startswith(qn):

                toks = nm.split()

                if toks and toks[0] == qn and len(toks) == 1:

                    single_token_pref.append(it)  # ex: 'laval' > 'laval-en-...'

                else:

                    prefix.append(it)

                continue

            pos = nm.find(qn)

            if pos >= 0:

                substr.append((pos, it))

    

        substr.sort(key=lambda t: (t[0], t[1].get("name", "")))

        ordered = exact + single_token_pref + prefix + [it for _, it in substr]

        return ordered[:limit] if (limit and limit > 0) else ordered
        if not qn:
            return []
        pref = [it for it in self.items if it["norm"].startswith(qn)]
        if len(pref) >= limit:
            return pref[:limit]
        cont = [it for it in self.items if qn in it["norm"] and it not in pref]
        return (pref + cont)[:limit]

# ----------------------------
# Index Départements (implémentation propre)
# ----------------------------
class DepartementIndex:
    def __init__(self):
        self.items: List[Dict[str, str]] = []  # {name, dep, norm}
        self.path: str = ""
        self.lib_col: str = "LIBELLE"  # détecté

    def load_csv(self, path: str) -> int:

        last_err = None
        for enc in ("utf-8", "utf-8-sig", "cp1252"):
            try:
                with open(path, "r", encoding=enc, newline="") as f:
                    reader = csv.DictReader(f)
                    flds = [c.strip() for c in (reader.fieldnames or [])]
                    if "DEP" not in flds:
                        raise ValueError("Le CSV départements doit contenir la colonne 'DEP'.")
                    lib_candidates = [c for c in flds if c.upper().startswith("LIBELLE")]
                    if not lib_candidates:
                        raise ValueError("Le CSV départements doit contenir 'LIBELLE' (ou 'LIBELLE*').")
                    self.lib_col = lib_candidates[0]
                    items: List[Dict[str, str]] = []
                    for row in reader:
                        name = (row.get(self.lib_col) or "").strip()
                        dep = (row.get("DEP") or "").strip()
                        if not name or not dep:
                            continue
                        items.append({"name": name, "dep": dep, "norm": _normalize(name)})
                self.items = items
                self.path = path
                return len(self.items)
            except Exception as e:
                last_err = e
                continue
        raise last_err or RuntimeError("Impossible de lire le CSV des départements.")

    def search(self, q: str, limit: int = 20) -> List[Dict[str, str]]:
        qn = _normalize(q)
        if not qn:
            return []
        pref = [it for it in self.items if it["norm"].startswith(qn)]
        if len(pref) >= limit:
            return pref[:limit]
        cont = [it for it in self.items if qn in it["norm"] and it not in pref]
        return (pref + cont)[:limit]

# ----------------------------
# Index Domaines
# ----------------------------
class DomaineIndex:
    def __init__(self):
        self.items: List[Dict[str, str]] = []  # {name, code}
        self.path: str = ""
        self.by_code: Dict[str, str] = {}
        self.by_name: Dict[str, str] = {}

    def load_csv(self, path: str) -> int:

        last_err = None
        for enc in ("utf-8", "utf-8-sig", "cp1252"):
            try:
                with open(path, "r", encoding=enc, newline="") as f:
                    reader = csv.DictReader(f)
                    flds = set(reader.fieldnames or [])
                    if not REQUIRED_DOMAINE_COLS.issubset(flds):
                        raise ValueError("Le CSV domaines doit contenir 'domainepro' et 'codedomainepro'.")
                    items: List[Dict[str, str]] = []
                    by_code, by_name = {}, {}
                    for row in reader:
                        name = (row.get("domainepro") or "").strip()
                        code = (row.get("codedomainepro") or "").strip()
                        if not name or not code:
                            continue
                        items.append({"name": name, "code": code})
                        by_code[code] = name
                        by_name[name] = code
                items.sort(key=lambda x: _normalize(x["name"]))
                self.items = items
                self.path = path
                self.by_code = by_code
                self.by_name = by_name
                return len(self.items)
            except Exception as e:
                last_err = e
                continue
        raise last_err or RuntimeError("Impossible de lire le CSV des domaines.")

# ----------------------------
# Widget d'auto‑complétion générique
# ----------------------------
class AutoCompleteEntry(ttk.Entry):
    def __init__(self, master, search_fn, on_select, *args, **kwargs):
        super().__init__(master, *args, **kwargs)
        self.search_fn = search_fn
        self.on_select = on_select
        self.popup: Optional[tk.Toplevel] = None
        self.listbox: Optional[tk.Listbox] = None
        self.suggestions: List[Dict[str, str]] = []
        self.bind('<KeyRelease>', self._on_key)
        self.bind('<Down>', self._move_down)
        self.bind('<Up>', self._move_up)
        self.bind('<Return>', self._confirm)
        self.bind('<Escape>', lambda e: self._hide())
        self.bind('<FocusOut>', lambda e: self.after(150, self._hide))

    def _on_key(self, event):
        if event.keysym in ("Down", "Up", "Return", "Escape"):
            return
        query = self.get()
        sugs = self.search_fn(query)
        if sugs:
            self._show(sugs)
        else:
            self._hide()

    def _show(self, suggestions: List[Dict[str, str]]):
        if not self.popup:
            self.popup = tk.Toplevel(self)
            self.popup.wm_overrideredirect(True)
            self.popup.attributes('-topmost', True)
            self.listbox = tk.Listbox(self.popup, height=min(12, max(3, len(suggestions))))
            self.listbox.pack(fill='both', expand=True)
            self.listbox.bind('<ButtonRelease-1>', self._click)
            self.listbox.bind('<Return>', self._confirm)
        self.suggestions = suggestions
        self.listbox.delete(0, 'end')
        for it in suggestions:
            extra = it.get('dep') or it.get('code') or ""
            label = f"{it['name']}" + (f" — {extra}" if extra else "")
            self.listbox.insert('end', label)
        self._place_popup()

    def _place_popup(self):
        if not self.popup:
            return
        x = self.winfo_rootx()
        y = self.winfo_rooty() + self.winfo_height()
        w = max(self.winfo_width(), 340)
        h = self.listbox.size()*18 + 6
        self.popup.geometry(f"{w}x{h}+{x}+{y}")

    def _hide(self):
        if self.popup:
            self.popup.destroy()
            self.popup = None
            self.listbox = None
            self.suggestions = []

    def _move_down(self, event):
        if not self.listbox:
            return
        i = self.listbox.curselection()
        idx = (i[0] + 1) if i else 0
        idx = min(idx, self.listbox.size() - 1)
        self.listbox.selection_clear(0, 'end')
        self.listbox.selection_set(idx)
        self.listbox.activate(idx)

    def _move_up(self, event):
        if not self.listbox:
            return
        i = self.listbox.curselection()
        idx = (i[0] - 1) if i else 0
        idx = max(idx, 0)
        self.listbox.selection_clear(0, 'end')
        self.listbox.selection_set(idx)
        self.listbox.activate(idx)

    def _click(self, event):
        self._confirm(event)

    def _confirm(self, event):
        if not self.listbox or not self.suggestions:
            return
        i = self.listbox.curselection()
        if not i:
            i = (0,)
        idx = i[0]
        if 0 <= idx < len(self.suggestions):
            chosen = self.suggestions[idx]
            self.delete(0, 'end')
            self.insert(0, chosen['name'])
            self._hide()
            self.on_select(chosen)

# ----------------------------
# Journal simple
# ----------------------------
class LogText(ttk.Frame):
    def __init__(self, master):
        super().__init__(master)
        self.text = tk.Text(self, height=14, wrap="word", bg=BG, fg=DARK)
        self.text.configure(state="disabled")
        vsb = ttk.Scrollbar(self, command=self.text.yview)
        self.text["yscrollcommand"] = vsb.set
        self.text.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

    def write(self, msg: str):
        self.text.configure(state="normal")
        self.text.insert("end", msg + "\n")
        self.text.see("end")
        self.text.configure(state="disabled")

    def clear(self):
        self.text.configure(state="normal")
        self.text.delete("1.0", "end")
        self.text.configure(state="disabled")

# ----------------------------
# Helpers de recherche de fichiers / logo
# ----------------------------

def candidate_dirs() -> List[str]:
    script_dir = os.path.dirname(os.path.abspath(__file__))
    cwd = os.getcwd()
    dirs = [
        script_dir,
        cwd,
        os.path.join(script_dir, "data"),
        os.path.join(cwd, "data"),
    ]
    out = []
    seen = set()
    for d in dirs:
        if d and d not in seen and os.path.isdir(d):
            out.append(d)
            seen.add(d)
    return out


def find_csv(filenames: List[str], required_cols: set, extra_col_prefix: Optional[str] = None) -> Optional[str]:
    # 1) noms connus
    for d in candidate_dirs():
        for name in filenames:
            p = os.path.join(d, name)
            if os.path.exists(p):
                return p
    # 2) scan
    for d in candidate_dirs():
        try:
            for fname in os.listdir(d):
                if not fname.lower().endswith('.csv'):
                    continue
                p = os.path.join(d, fname)
                for enc in ("utf-8", "utf-8-sig", "cp1252"):
                    try:
                        with open(p, "r", encoding=enc, newline="") as f:
                            reader = csv.DictReader(f)
                            flds = set(reader.fieldnames or [])
                            if not required_cols.issubset(flds):
                                continue
                            if extra_col_prefix:
                                if not any(c.upper().startswith(extra_col_prefix.upper()) for c in flds):
                                    continue
                            return p
                    except Exception:
                        continue
        except Exception:
            continue
    return None


def find_logo() -> Optional[str]:
    # 1) variable d'env
    envp = os.environ.get("APP_LOGO")
    if envp and os.path.exists(envp):
        return envp
    # 2) fichiers connus
    for d in candidate_dirs():
        for name in DEFAULT_LOGO_FILENAMES:
            p = os.path.join(d, name)
            if os.path.exists(p):
                return p
    return None


# ----------------------------
# Application
# ----------------------------
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Étude d'offres — Interface")
        self.geometry("980x640")
        self.minsize(860, 520)

        self.browser = None
        # Styles (ttk)
        style = ttk.Style(self)
        try:
            style.theme_use('clam')
        except Exception:
            pass
        style.configure('TFrame', background=BG)
        style.configure('TLabelframe', background=BG, foreground=DARK)
        style.configure('TLabelframe.Label', background=BG, foreground=DARK)
        style.configure('TLabel', background=BG, foreground=DARK)
        style.configure('Accent.TButton', foreground='white', background=RED)
        style.map('Accent.TButton', background=[('active', RED)], foreground=[('active', 'white')])

        container = ttk.Frame(self, padding=12, style='TFrame')
        container.pack(fill="both", expand=True)

        # --- Ligne supérieure : logo (col 0) + cadre Paramètres (col 1) ---
        top = ttk.Frame(container, style='TFrame')
        top.pack(fill='x', side='top', pady=(0, 8))
        top.grid_columnconfigure(1, weight=1)

        # Logo à gauche
        self.logo_img = None
        logo_holder = ttk.Frame(top, style='TFrame')
        logo_holder.grid(row=0, column=0, sticky='nw', padx=(0, 12))
        self.logo_label = tk.Label(logo_holder, bg=BG)
        self.logo_label.grid(row=0, column=0, sticky='nw')
        self._load_and_place_logo()

        # Formulaire à droite du logo
        form = ttk.LabelFrame(top, text="Paramètres", padding=8)
        form.grid(row=0, column=1, sticky='ew')

        # Index
        self.index_commune = CommuneIndex()
        self.index_dept = DepartementIndex()
        self.index_domaine = DomaineIndex()

        # Variables
        self.var_domaine_code = tk.StringVar(value="tout")  # code utilisé dans URL
        self.var_domaine_label = tk.StringVar(value="tout") # affiché dans la combo
        self.var_emission = tk.StringVar(value="tout")
        self.var_loc_mode = tk.StringVar(value="commune")  # "commune" | "departement"
        # Commune
        self.var_commune = tk.StringVar()
        self.var_lieux_commune = tk.StringVar()  # code INSEE (readonly)
        # Département
        self.var_dept = tk.StringVar()
        self.var_dept_code = tk.StringVar()      # DEP (readonly)
        # communs
        self.var_rayon = tk.StringVar(value="0")
        self.var_outdir = tk.StringVar()
        self.var_communes_csv = tk.StringVar()
        self.var_departements_csv = tk.StringVar()
        self.var_domaines_csv = tk.StringVar()

        # Ligne 0 — Domaine (combo) + Émission
        ttk.Label(form, text="Domaine").grid(row=0, column=0, sticky="w", padx=6, pady=6)
        self.cb_domaine = ttk.Combobox(form, textvariable=self.var_domaine_label, state="readonly", width=38)
        self.cb_domaine.grid(row=0, column=1, sticky="we", padx=6, pady=6)
        self.cb_domaine.bind('<<ComboboxSelected>>', self.on_domaine_selected)

        ttk.Label(form, text="Émission (jours)").grid(row=0, column=2, sticky="w", padx=6, pady=6)
        ttk.Entry(form, textvariable=self.var_emission, width=10).grid(row=0, column=3, sticky="w", padx=6, pady=6)

        # Ligne 1 — Mode lieu
        ttk.Label(form, text="Lieu par").grid(row=1, column=0, sticky="w", padx=6, pady=6)
        rb1 = ttk.Radiobutton(form, text="Commune", value="commune", variable=self.var_loc_mode, command=self.update_location_mode_ui)
        rb2 = ttk.Radiobutton(form, text="Département", value="departement", variable=self.var_loc_mode, command=self.update_location_mode_ui)
        rb1.grid(row=1, column=1, sticky="w", padx=6, pady=6)
        rb2.grid(row=1, column=2, sticky="w", padx=6, pady=6)

        # Ligne 2 — Commune
        self.row_commune = 2
        ttk.Label(form, text="Commune").grid(row=self.row_commune, column=0, sticky="w", padx=6, pady=6)
        self.entry_commune = AutoCompleteEntry(
            form,
            search_fn=lambda q: self.index_commune.search(q, 20),
            on_select=self.on_commune_chosen,
            textvariable=self.var_commune,
            width=42,
        )
        self.entry_commune.grid(row=self.row_commune, column=1, columnspan=3, sticky="we", padx=6, pady=6)
        ttk.Label(form, text="Code INSEE").grid(row=self.row_commune, column=4, sticky="w", padx=6, pady=6)
        self.entry_insee = ttk.Entry(form, textvariable=self.var_lieux_commune, width=16, state="readonly")
        self.entry_insee.grid(row=self.row_commune, column=5, sticky="w", padx=6, pady=6)

        # Ligne 3 — Département
        self.row_dept = 3
        ttk.Label(form, text="Département").grid(row=self.row_dept, column=0, sticky="w", padx=6, pady=6)
        self.entry_dept = AutoCompleteEntry(
            form,
            search_fn=lambda q: self.index_dept.search(q, 20),
            on_select=self.on_dept_chosen,
            textvariable=self.var_dept,
            width=42,
        )
        self.entry_dept.grid(row=self.row_dept, column=1, columnspan=3, sticky="we", padx=6, pady=6)
        ttk.Label(form, text="Code dép.").grid(row=self.row_dept, column=4, sticky="w", padx=6, pady=6)
        self.entry_dept_code = ttk.Entry(form, textvariable=self.var_dept_code, width=16, state="readonly")
        self.entry_dept_code.grid(row=self.row_dept, column=5, sticky="w", padx=6, pady=6)

        # Ligne 4 — Rayon
        ttk.Label(form, text="Rayon (km)").grid(row=4, column=0, sticky="w", padx=6, pady=6)
        ttk.Entry(form, textvariable=self.var_rayon, width=10).grid(row=4, column=1, sticky="w", padx=6, pady=6)

        # Ligne 5 — Dossier de sortie
        ttk.Label(form, text="Dossier de sortie").grid(row=5, column=0, sticky="w", padx=6, pady=6)
        outdir_entry = ttk.Entry(form, textvariable=self.var_outdir, width=42)
        outdir_entry.grid(row=5, column=1, columnspan=3, sticky="we", padx=6, pady=6)
        ttk.Button(form, text="Choisir…", command=self.choose_outdir, style='Accent.TButton').grid(row=5, column=4, sticky="w", padx=6, pady=6)

        for c in range(6):
            form.grid_columnconfigure(c, weight=1)

        # Boutons principaux (sans etud.py)
        btns = ttk.Frame(container, style='TFrame')
        btns.pack(fill="x", pady=(8, 0))
        ttk.Button(btns, text="Démarrer (dry‑run)", command=self.on_start, style='Accent.TButton').pack(side="left")
        ttk.Button(btns, text="Initialiser Chrome", command=self.on_init_chrome).pack(side="left", padx=8)
        ttk.Button(btns, text="Nb offres", command=self.on_count_offers).pack(side="left", padx=8)
        ttk.Button(btns, text="Lancer l’étude", command=self.on_run_study).pack(side="left", padx=8)
        ttk.Button(btns, text="finaliser", command=self.on_end_study).pack(side="left", padx=8)
        ttk.Button(btns, text="Effacer le journal", command=self.on_clear).pack(side="left", padx=8)
        ttk.Button(btns, text="Sauver paramètres", command=self.on_save).pack(side="left")
        ttk.Button(btns, text="Quitter", command=self.on_close).pack(side="right")

        # Journal
        self.log = LogText(container)
        self.log.pack(fill="both", expand=True, pady=(8, 0))

        # Chargement config + auto‑load références
        cfg = load_config()
        if cfg:
            self.apply_params(cfg)
        self.try_autoload_domaines()
        self.try_autoload_communes()
        self.try_autoload_departements()
        self.update_location_mode_ui()

    # ---------------- Helpers UI ----------------
    def _load_and_place_logo(self):
        path = find_logo()
        if not path:
            guess = os.path.join(os.getcwd(), "logo.png")
            path = guess if os.path.exists(guess) else None
        if not path:
            return
        try:
            # Essayer Pillow pour un redimensionnement propre
            try:
                from PIL import Image, ImageTk  # type: ignore
                img = Image.open(path)
                target_w = 120
                w, h = img.size
                if w > target_w:
                    ratio = target_w / float(w)
                    new_size = (int(w * ratio), int(h * ratio))
                    img = img.resize(new_size, Image.LANCZOS)
                self.logo_img = ImageTk.PhotoImage(img)
            except Exception:
                im = tk.PhotoImage(file=path)
                w = im.width()
                target_w = 120
                factor = max(1, int(w / target_w))
                self.logo_img = im.subsample(factor, factor)
            self.logo_label.configure(image=self.logo_img)
        except Exception:
            pass

    def choose_outdir(self):
        path = filedialog.askdirectory(title="Choisir le dossier de sortie")
        if path:
            self.var_outdir.set(path)

    # --- Auto‑load CSV ---
    def _load_communes_path(self, path: str):
        try:
            n = self.index_commune.load_csv(path)
            self.var_communes_csv.set(path)
            self.log.write(f"Communes chargées depuis {path} ({n} lignes)")
        except Exception as e:
            self.log.write(f"Erreur CSV communes: {e}")

    def _load_departements_path(self, path: str):
        try:
            n = self.index_dept.load_csv(path)
            self.var_departements_csv.set(path)
            self.log.write(f"Départements chargés depuis {path} ({n} lignes)")
        except Exception as e:
            self.log.write(f"Erreur CSV départements: {e}")

    def _load_domaines_path(self, path: str):
        try:
            n = self.index_domaine.load_csv(path)
            self.var_domaines_csv.set(path)
            self.refresh_domaine_combobox()
            self.log.write(f"Domaines chargés depuis {path} ({n} lignes)")
        except Exception as e:
            self.log.write(f"Erreur CSV domaines: {e}")
            self.refresh_domaine_combobox()

    def try_autoload_communes(self):
        env_path = os.environ.get("COMMUNES_CSV")
        if env_path and os.path.exists(env_path):
            self._load_communes_path(env_path)
            return
        found = find_csv(DEFAULT_COMMUNE_FILENAMES, REQUIRED_COMMUNE_COLS)
        if found:
            self._load_communes_path(found)
            return
        self.log.write("Aucun CSV communes auto‑détecté : placez 'communes-france-2025.csv' (ou équivalent) à côté du script ou dans data/.")

    def try_autoload_departements(self):
        env_path = os.environ.get("DEPARTEMENTS_CSV")
        if env_path and os.path.exists(env_path):
            self._load_departements_path(env_path)
            return
        found = find_csv(DEFAULT_DEPT_FILENAMES, REQUIRED_DEPT_COLS, extra_col_prefix="LIBELLE")
        if found:
            self._load_departements_path(found)
        else:
            self.log.write("Aucun CSV départements auto‑détecté : placez 'departement.csv' (ou équivalent) à côté du script ou dans data/.")

    def try_autoload_domaines(self):
        env_path = os.environ.get("DOMAINES_CSV")
        if env_path and os.path.exists(env_path):
            self._load_domaines_path(env_path)
            return
        found = find_csv(DEFAULT_DOMAINE_FILENAMES, REQUIRED_DOMAINE_COLS)
        if found:
            self._load_domaines_path(found)
        else:
            self.refresh_domaine_combobox()
            self.log.write("Aucun CSV domaines auto‑détecté : placez 'domainepro.csv' (ou équivalent) à côté du script ou dans data/.")

    # --- Domaine combo logic ---
    def refresh_domaine_combobox(self):
        labels = [it["name"] for it in self.index_domaine.items]
        self.cb_domaine["values"] = ["tout"] + labels
        code = self.var_domaine_code.get() or "tout"
        if code == "tout":
            self.var_domaine_label.set("tout")
        else:
            name = self.index_domaine.by_code.get(code)
            self.var_domaine_label.set(name or "tout")
        if not self.cb_domaine.get():
            self.var_domaine_label.set("tout")

    def on_domaine_selected(self, event=None):
        label = self.var_domaine_label.get()
        if label == "tout":
            self.var_domaine_code.set("tout")
        else:
            code = self.index_domaine.by_name.get(label)
            self.var_domaine_code.set(code or "tout")

    # --- Callbacks AutoComplete ---
    def on_commune_chosen(self, item: Dict[str, str]):
        self.var_commune.set(item["name"])
        self.entry_insee.configure(state="normal")
        self.var_lieux_commune.set(item["code"])  # code INSEE
        self.entry_insee.configure(state="readonly")
        self.log.write(f"Commune sélectionnée: {item['name']} (INSEE {item['code']})")

    def on_dept_chosen(self, item: Dict[str, str]):
        self.var_dept.set(item["name"])
        self.entry_dept_code.configure(state="normal")
        self.var_dept_code.set(item["dep"])  # DEP
        self.entry_dept_code.configure(state="readonly")
        self.log.write(f"Département sélectionné: {item['name']} (DEP {item['dep']})")

    def update_location_mode_ui(self):
        mode = self.var_loc_mode.get()
        for w in (self.entry_commune, self.entry_insee):
            w.grid() if mode == "commune" else w.grid_remove()
        for child in self.entry_commune.master.grid_slaves(row=self.row_commune):
            if isinstance(child, ttk.Label):
                child.grid() if mode == "commune" else child.grid_remove()
        for w in (self.entry_dept, self.entry_dept_code):
            w.grid() if mode == "departement" else w.grid_remove()
        for child in self.entry_dept.master.grid_slaves(row=self.row_dept):
            if isinstance(child, ttk.Label):
                child.grid() if mode == "departement" else child.grid_remove()

    # --- Collecte des paramètres ---
    def params_from_ui(self) -> Optional[Params]:
        mode = self.var_loc_mode.get()
        lieux = ""
        if mode == "commune":
            if not self.var_lieux_commune.get().strip() and self.var_commune.get().strip() and self.index_commune.items:
                sug = self.index_commune.search(self.var_commune.get().strip(), limit=1)
                if sug:
                    self.on_commune_chosen(sug[0])
            lieux = self.var_lieux_commune.get().strip()
            if not lieux:
                messagebox.showerror("Code INSEE manquant", "Choisissez une commune dans la liste pour renseigner automatiquement le code INSEE.")
                return None
        else:
            if not self.var_dept_code.get().strip() and self.var_dept.get().strip() and self.index_dept.items:
                sug = self.index_dept.search(self.var_dept.get().strip(), limit=1)
                if sug:
                    self.on_dept_chosen(sug[0])
            dep = self.var_dept_code.get().strip()
            if not dep:
                messagebox.showerror("Département manquant", "Choisissez un département dans la liste.")
                return None
            lieux = f"{dep}D"  # règle demandée

        domaine_code = self.var_domaine_code.get().strip() or "tout"

        return Params(
            domaine=domaine_code,
            emission=self.var_emission.get().strip() or "tout",
            lieux=lieux,
            rayon=self.var_rayon.get().strip() or "0",
            output_dir=self.var_outdir.get().strip(),
            communes_csv=self.var_communes_csv.get().strip(),
            departements_csv=self.var_departements_csv.get().strip(),
            domaines_csv=self.var_domaines_csv.get().strip(),
            location_mode=mode,
        )

    def apply_params(self, p: Params):
        self.var_domaine_code.set(p.domaine if p.domaine else "tout")
        self.var_emission.set(p.emission)
        self.var_rayon.set(p.rayon)
        self.var_outdir.set(p.output_dir)
        self.var_communes_csv.set(p.communes_csv)
        self.var_departements_csv.set(p.departements_csv)
        self.var_domaines_csv.set(p.domaines_csv)
        if p.location_mode in ("commune", "departement"):
            self.var_loc_mode.set(p.location_mode)

    def on_clear(self):
        self.log.clear()

    def on_save(self):
        p = self.params_from_ui()
        if not p:
            return
        save_config(p)
        self.log.write("Paramètres sauvegardés dans %s" % CONFIG_FILE)

    def on_start(self):
        p = self.params_from_ui()
        if not p:
            return
        self.log.write("— Lancement DRY‑RUN —")
        self.log.write("Paramètres:")
        for k, v in asdict(p).items():
            if k in ("communes_csv", "departements_csv", "domaines_csv"):
                continue
            self.log.write(f"  • {k} = {v}")
        url = build_adressereq(p)
        self.log.write("")
        self.log.write("URL de recherche générée :")
        self.log.write(url)
        self.log.write("— Fin DRY‑RUN —")

    # --- Compteur d'offres ---
    def on_count_offers(self):

        p = self.params_from_ui()
        if not p:
            return
        adressereq = build_adressereq(p)
        self.log.write("Requête du nombre d'offres…")
        threading.Thread(target=self._count_offers_thread, args=(adressereq,), daemon=True).start()

    def _count_offers_thread(self, adressereq: str):


        try:
            url = adressereq + ("&" if "?" in adressereq else "?") + "range=0-19&tri=0"
            headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124 Safari/537.36"}
            response = requests.get(url, headers=headers, timeout=20)
            response.raise_for_status()
            soup = BeautifulSoup(response.content, "html.parser")
            zone = soup.find(class_="zone-resultats")
            if zone and zone.h1:
                result = zone.h1.get_text(" ", strip=True)
            else:
                h1 = soup.find("h1")
                result = h1.get_text(" ", strip=True) if h1 else ""
            cleaned = result.replace("\xa0", " ")
            nums = [int(t) for t in cleaned.split() if t.isdigit()]
            if nums:
                nb_result = nums[0]
            else:
                m = re.search(r"\d+", cleaned)
                nb_result = int(m.group()) if m else 0
            self.log.write(f"il y a {nb_result} offres")
        except Exception as e:
            self.log.write(f"Erreur lors du comptage d'offres : {e}")


    # --- Chrome (initialisation & PDF DevTools) ---
    def on_init_chrome(self):
        outdir = self.var_outdir.get().strip()
        if not outdir:
            from tkinter import messagebox
            messagebox.showerror("Dossier manquant", "Choisissez un dossier de sortie (Paramètres > Dossier de sortie).")
            return
        pdf_dir = os.path.join(outdir, "pdfpar")
        prof_dir = os.path.join(outdir, "chrome-profile")
        os.makedirs(pdf_dir, exist_ok=True)
        os.makedirs(prof_dir, exist_ok=True)

        try:
            options = self._build_chrome_options(download_dir=pdf_dir, profile_dir=prof_dir)
          # Selenium 4+ : Selenium Manager résout chromedriver automatiquement (pas besoin d'exe)
            self.browser = webdriver.Chrome(options=options)
            self.browser.get("https://www.francetravail.fr/")
            self.log.write(f"Chrome initialisé ✓  (profil: {prof_dir})")
        except Exception as e:
            self.log.write(f"Échec initialisation Chrome : {e}")

    def _build_chrome_options(self, download_dir: str, profile_dir: str):
        options = webdriver.ChromeOptions()

        # Préférences téléchargement / PDF
        settings = {
            "recentDestinations": [
                {"id": "Save as PDF", "origin": "local", "account": ""}
            ],
            "selectedDestinationId": "Save as PDF",
            "version": 2,
            "isHeaderFooterEnabled": True,
            "isLandscapeEnabled": False
        }
        prefs = {
            "savefile.default_directory": download_dir,
            "download.default_directory": download_dir,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": True,
            # Conserve le réglage d'impression (utile si tu utilises encore Ctrl+P)
            "printing.print_preview_sticky_settings.appState": json.dumps(settings)
        }
        options.add_experimental_option("prefs", prefs)

        # Profil persistant => cookies conservés
        options.add_argument(f"--user-data-dir={profile_dir}")

        # Option historique: impression sans dialogue (si tu utilises encore l'imprimante "Save as PDF")
        options.add_argument("--kiosk-printing")

        # (facultatif) Forcer un binaire Chrome précis via variable d'env
        #   set CHROME_BIN=C:\chemin\vers\chrome.exe
        chrome_bin = os.environ.get("CHROME_BIN")
        if chrome_bin and os.path.exists(chrome_bin):
            options.binary_location = chrome_bin

        return options

    def print_url_to_pdf(self, url: str, out_pdf: str):
        """
        Méthode 'propre' pour produire un PDF sans passer par 'Save as PDF':
        utilise l'API DevTools Page.printToPDF via Selenium 4.
        """
        if not getattr(self, "browser", None):
            self.log.write("Chrome non initialisé. Clique d'abord sur 'Initialiser Chrome'.")
            return
        try:
            import base64, time
            self.browser.get(url)
            time.sleep(1.0)  # laisse le temps de charger
            result = self.browser.execute_cdp_cmd("Page.printToPDF", {
                "printBackground": True,
                "landscape": False,
                "preferCSSPageSize": True
            })
            data = base64.b64decode(result.get("data", ""))
            with open(out_pdf, "wb") as f:
                f.write(data)
            self.log.write(f"PDF écrit : {out_pdf}")
        except Exception as e:
            self.log.write(f"Erreur impression PDF : {e}")

    def on_close(self):
        try:
            if getattr(self, "browser", None):
                self.browser.quit()
        except Exception:
            pass
        self.destroy()


    # --- Étude complète : comptage + collecte des références ---
    def on_run_study(self):
        p = self.params_from_ui()
        if not p:
            return
        adressereq = build_adressereq(p)
        self.log.write("— Lancer l’étude —")
        self.log.write("Préparation de la requête…")
        import threading
        threading.Thread(target=self._run_study_thread, args=(adressereq,), daemon=True).start()

    def _http_get(self, url: str):
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124 Safari/537.36"
        }
        r = requests.get(url, headers=headers, timeout=25)
        r.raise_for_status()
        return r

    def _count_offers(self, adressereq: str) -> int:
        # même logique que "Nb offres"
        url = adressereq + ("&" if "?" in adressereq else "?") + "range=0-19&tri=0"
        try:
            resp = self._http_get(url)
            soup = BeautifulSoup(resp.content, "html.parser")
            zone = soup.find(class_="zone-resultats")
            if zone and zone.h1:
                result = zone.h1.get_text(" ", strip=True)
            else:
                h1 = soup.find("h1")
                result = h1.get_text(" ", strip=True) if h1 else ""
            cleaned = result.replace("\xa0", " ")
            nums = [int(t) for t in cleaned.split() if t.isdigit()]
            if nums:
                return nums[0]
            m = re.search(r"\d+", cleaned)
            return int(m.group()) if m else 0
        except Exception as e:
            self.log.write(f"Erreur comptage des offres : {e}")
            return 0

    def _collect_refs_from_soup(self, soup, references, index_offres, nb_result_reel_ref):
        # Parcourt <li class="result"> et récupère l'attribut data-id-offre
        added = 0
        for li in soup.find_all("li", class_="result"):
            off_id = li.get("data-id-offre")
            if not off_id:
                continue
            references.append(off_id)
            index_offres.append(nb_result_reel_ref[0])
            nb_result_reel_ref[0] += 1
            added += 1
        return added

 


    # --- util HTTP ---
    def _http_get(self, url: str):
        headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124 Safari/537.36"}
        r = requests.get(url, headers=headers, timeout=25)
        r.raise_for_status()
        return r

    def _count_offers(self, adressereq: str) -> int:
        url = adressereq + ("&" if "?" in adressereq else "?") + "range=0-19&tri=0"
        try:
            resp = self._http_get(url)
            soup = BeautifulSoup(resp.content, "html.parser")
            zone = soup.find(class_="zone-resultats")
            if zone and zone.h1:
                txt = zone.h1.get_text(" ", strip=True)
            else:
                h1 = soup.find("h1")
                txt = h1.get_text(" ", strip=True) if h1 else ""
            cleaned = txt.replace("\xa0", " ")
            nums = [int(t) for t in cleaned.split() if t.isdigit()]
            if nums:
                return nums[0]
            m = re.search(r"\d+", cleaned)
            return int(m.group()) if m else 0
        except Exception:
            return 0

    def _collect_refs_from_soup(self, soup, references, index_offres, nb_result_reel_ref):
        # récupère <li class="result" data-id-offre="...">
        added = 0
        for li in soup.find_all("li", class_="result"):
            off_id = li.get("data-id-offre")
            if not off_id:
                continue
            references.append(off_id)
            index_offres.append(nb_result_reel_ref[0])
            nb_result_reel_ref[0] += 1
            added += 1
        return added

    # --- Impression PDF d'une offre + récupération du contrat ---
    def pepdf(self, idoffre: str, rang: int, out_pdf: str) -> str:
        try:
            # 1) page détail
            r = self._http_get(f'https://candidat.francetravail.fr/offres/recherche/detail/{idoffre}')
            soup = BeautifulSoup(r.content, "html.parser")

            # 2) contrat (ta logique)
            try:
                dd_list = soup.find_all("dd")
                contrat = (dd_list[0].contents)[0].strip('\n') if dd_list else "-"
            except Exception:
                contrat = "-"

            # 3) réduire la page pour le PDF
            try:
                soupe = soup.find(class_="modal-details-offre")
                sup = soupe.find(class_="other-offers-container")
                if sup is not None:
                    sup.clear()
                sup = soupe.find(class_="media-body media-middle")
                if sup is not None:
                    sup.clear()
            except Exception:
                soupe = soup  # fallback: toute la page

            # 4) PDF via WeasyPrint
            try:
                html = weasyprint.HTML(string=str(soupe))
                css = [weasyprint.CSS(string=f"@page {{ @top-center {{ content: 'offre n°{rang}'; }} }}")]
                # CSS optionnel ../none.css si présent
                try:
                    css_path = os.path.join(os.path.dirname(out_pdf), "..", "none.css")
                    if os.path.exists(css_path):
                        css.append(weasyprint.CSS(filename=css_path))
                except Exception:
                    pass
                html.write_pdf(out_pdf, stylesheets=css)
                self.log.write(f"PDF écrit : {out_pdf}")
            except Exception as e:
                self.log.write(f"Erreur WeasyPrint : {e}")

            return contrat
        except Exception as e:
            self.log.write(f"Erreur pepdf({idoffre}) : {e}")
            return "-"

    # --- Lancer l'étude ---
    def on_run_study(self):
        p = self.params_from_ui()
        if not p:
            return
        adressereq = build_adressereq(p)
        self.log.write("— Lancer l’étude —")
        self.log.write("Préparation de la requête…")
        threading.Thread(target=self._run_study_thread, args=(adressereq, p), daemon=True).start()

    def _run_study_thread(self, adressereq: str, p):
        import os
        # 1) nb offres
        nb_result = self._count_offers(adressereq)
        self.log.write(f"Nombre d'offres détecté : {nb_result}")
        if nb_result <= 0:
            self.log.write("Aucune offre ou échec du comptage. Arrêt.")
            return

        # 2) références + index_offres
        references, index_offres = [], []
        nb_result_reel = [1]
        # page 0-19
        url0 = adressereq + ("&" if "?" in adressereq else "?") + "range=0-19&tri=0"
        resp0 = self._http_get(url0)
        soup0 = BeautifulSoup(resp0.content, "html.parser")
        n0 = self._collect_refs_from_soup(soup0, references, index_offres, nb_result_reel)
        self.log.write(f"Page 0-19 : {n0} références")
        # pagination
        nb_range = nb_result // 20
        if nb_result > 20:
            etat = 1
            while etat <= nb_range:
                start, end = etat * 20, etat * 20 + 19
                page_url = adressereq + ("&" if "?" in adressereq else "?") + f"range={start}-{end}&tri=0"
                try:
                    resp = self._http_get(page_url)
                    soup = BeautifulSoup(resp.content, "html.parser")
                    n = self._collect_refs_from_soup(soup, references, index_offres, nb_result_reel)
                    self.log.write(f"Page {start}-{end} : {n} références")
                except Exception as e:
                    self.log.write(f"Erreur page {start}-{end} : {e}")
                etat += 1

        self.log.write(f"Total références collectées : {len(references)}")

        # 3) listes partenaires (par défaut 'pe' comme dans ton exemple)
        urlpartenaire, nompartenaire = [], []
        for refoffre in references:
            if not refoffre[-4:].isalpha():
                self.log.write("— partenaire —")  
                browser = getattr(self, "browser", None)
                nompart, lien = "pe", "pe"
                if browser:
                    try:
                        detail_url = f"https://candidat.francetravail.fr/offres/recherche/detail/{refoffre}"
                        browser.get(detail_url)
                        browser.implicitly_wait(60)
                        try:
                            browser.find_element(By.ID, "detail-apply").click()
                        except Exception:
                            self.log.write("clic 'Postuler' échoué, tentative refresh…")
                            browser.refresh()
                            time.sleep(1)
                            try:
                                browser.find_element(By.ID, "detail-apply").click()
                            except Exception:
                                pass
                        try:
                            nompar = browser.find_element(By.CSS_SELECTOR, "div.item div.media-body.media-middle h4")
                            nompart = nompar.text.strip()
                        except Exception:
                            try:
                                nompar = browser.find_element(By.CSS_SELECTOR, "div.item div.media-body media-middle h4")
                                nompart = nompar.text.strip()
                            except Exception:
                                nompart = "?"
                        try:
                            lien = browser.find_element(By.ID, "idLienPartenaire").get_attribute("href")
                        except Exception:
                            lien = ""
                    except Exception as e:
                        self.log.write(f"Erreur Selenium partenaire ({refoffre}) : {e}")
                    urlpartenaire.append(lien)
                    nompartenaire.append(nompart)
                else:
                    self.log.write("Chrome non initialisé — partenaire par défaut 'pe'.")

            else:
                urlpartenaire.append("pe")
                nompartenaire.append("pe")
        print(references)
        print(urlpartenaire)
        print(nompartenaire)

        # 4) Répertoires sortie
        outdir = p.output_dir or os.getcwd()
        pdf_dir = os.path.join(outdir, f"pdf-{p.lieux}-{p.emission}jours-{p.domaine}-rayon{p.rayon}km")
        os.makedirs(pdf_dir, exist_ok=True)

        # 5) DataFrame initial + CSV
        print(references)
        print(urlpartenaire)
        print(nompartenaire)

        df = pd.DataFrame({
            'index_etud': index_offres,
            'reference de l offre': references,
            'nom du partenaire': nompartenaire,
            'url partenaire': urlpartenaire
        })
        csv_path = os.path.join(outdir, f"etud-{p.lieux}-{p.emission}jours-{p.domaine}-rayon{p.rayon}km.csv")
        try:
            df.to_csv(csv_path, index=False, encoding="utf-8")
            self.log.write(f"CSV écrit : {csv_path}")
        except Exception as e:
            self.log.write(f"Erreur écriture CSV initial : {e}")

        # 6) Impression PDF + collecte typecontrat
        typecontrat = []
        for q in df.itertuples():
            idx_etud = q[1]  # index_etud
            ref = q[2]       # reference de l offre
            out_pdf = os.path.join(pdf_dir, f"{idx_etud:04d}-{ref}.pdf")
            contrat = self.pepdf(ref, idx_etud, out_pdf)
            typecontrat.append(contrat)

        # 7) Ajout de la colonne typedecontrat + CSV final
        try:
            df = df.assign(typedecontrat=typecontrat)
            df.to_csv(csv_path, index=False, encoding="utf-8")
            self.log.write(f"CSV mis à jour avec typedecontrat : {csv_path}")
        except Exception as e:
            self.log.write(f"Erreur mise à jour CSV : {e}")

        self.log.write("— Fin de l’étude —")
    def _build_main_interface(self):
        """
        (Re)construit l'écran d'accueil sans recréer les StringVar ni les index.
        Appelée par reload_main_interface() après avoir détruit les widgets.
        """
        # Container racine de l'UI
        container = ttk.Frame(self, padding=12, style='TFrame')
        container.pack(fill="both", expand=True)

        # --- Bandeau haut : logo + formulaire paramètres ---
        top = ttk.Frame(container, style='TFrame')
        top.pack(fill='x', side='top', pady=(0, 8))
        top.grid_columnconfigure(1, weight=1)

        # Logo (gauche)
        self.logo_img = None
        logo_holder = ttk.Frame(top, style='TFrame')
        logo_holder.grid(row=0, column=0, sticky='nw', padx=(0, 12))
        self.logo_label = tk.Label(logo_holder, bg=BG)
        self.logo_label.grid(row=0, column=0, sticky='nw')
        self._load_and_place_logo()

        # Formulaire (droite)
        form = ttk.LabelFrame(top, text="Paramètres", padding=8)
        form.grid(row=0, column=1, sticky='ew')

        # Ligne 0 — Domaine + Émission
        ttk.Label(form, text="Domaine").grid(row=0, column=0, sticky="w", padx=6, pady=6)
        self.cb_domaine = ttk.Combobox(form, textvariable=self.var_domaine_label, state="readonly", width=38)
        self.cb_domaine.grid(row=0, column=1, sticky="we", padx=6, pady=6)
        self.cb_domaine.bind('<<ComboboxSelected>>', self.on_domaine_selected)

        ttk.Label(form, text="Émission (jours)").grid(row=0, column=2, sticky="w", padx=6, pady=6)
        ttk.Entry(form, textvariable=self.var_emission, width=10).grid(row=0, column=3, sticky="w", padx=6, pady=6)

        # Ligne 1 — Mode lieu
        ttk.Label(form, text="Lieu par").grid(row=1, column=0, sticky="w", padx=6, pady=6)
        rb1 = ttk.Radiobutton(form, text="Commune", value="commune",
                              variable=self.var_loc_mode, command=self.update_location_mode_ui)
        rb2 = ttk.Radiobutton(form, text="Département", value="departement",
                              variable=self.var_loc_mode, command=self.update_location_mode_ui)
        rb1.grid(row=1, column=1, sticky="w", padx=6, pady=6)
        rb2.grid(row=1, column=2, sticky="w", padx=6, pady=6)

        # Ligne 2 — Commune
        self.row_commune = 2
        ttk.Label(form, text="Commune").grid(row=self.row_commune, column=0, sticky="w", padx=6, pady=6)
        self.entry_commune = AutoCompleteEntry(
            form,
            search_fn=lambda q: self.index_commune.search(q, 20),
            on_select=self.on_commune_chosen,
            textvariable=self.var_commune,
            width=42,
        )
        self.entry_commune.grid(row=self.row_commune, column=1, columnspan=3, sticky="we", padx=6, pady=6)
        ttk.Label(form, text="Code INSEE").grid(row=self.row_commune, column=4, sticky="w", padx=6, pady=6)
        self.entry_insee = ttk.Entry(form, textvariable=self.var_lieux_commune, width=16, state="readonly")
        self.entry_insee.grid(row=self.row_commune, column=5, sticky="w", padx=6, pady=6)

        # Ligne 3 — Département
        self.row_dept = 3
        ttk.Label(form, text="Département").grid(row=self.row_dept, column=0, sticky="w", padx=6, pady=6)
        self.entry_dept = AutoCompleteEntry(
            form,
            search_fn=lambda q: self.index_dept.search(q, 20),
            on_select=self.on_dept_chosen,
            textvariable=self.var_dept,
            width=42,
        )
        self.entry_dept.grid(row=self.row_dept, column=1, columnspan=3, sticky="we", padx=6, pady=6)
        ttk.Label(form, text="Code dép.").grid(row=self.row_dept, column=4, sticky="w", padx=6, pady=6)
        self.entry_dept_code = ttk.Entry(form, textvariable=self.var_dept_code, width=16, state="readonly")
        self.entry_dept_code.grid(row=self.row_dept, column=5, sticky="w", padx=6, pady=6)

        # Ligne 4 — Rayon
        ttk.Label(form, text="Rayon (km)").grid(row=4, column=0, sticky="w", padx=6, pady=6)
        ttk.Entry(form, textvariable=self.var_rayon, width=10).grid(row=4, column=1, sticky="w", padx=6, pady=6)

        # Ligne 5 — Dossier de sortie
        ttk.Label(form, text="Dossier de sortie").grid(row=5, column=0, sticky="w", padx=6, pady=6)
        outdir_entry = ttk.Entry(form, textvariable=self.var_outdir, width=42)
        outdir_entry.grid(row=5, column=1, columnspan=3, sticky="we", padx=6, pady=6)
        ttk.Button(form, text="Choisir…", command=self.choose_outdir, style='Accent.TButton')\
           .grid(row=5, column=4, sticky="w", padx=6, pady=6)

        for c in range(6):
            form.grid_columnconfigure(c, weight=1)

        # --- Boutons principaux
        btns = ttk.Frame(container, style='TFrame')
        btns.pack(fill="x", pady=(8, 0))
        ttk.Button(btns, text="Démarrer (dry-run)", command=self.on_start, style='Accent.TButton').pack(side="left")
        ttk.Button(btns, text="Initialiser Chrome", command=self.on_init_chrome).pack(side="left", padx=8)
        ttk.Button(btns, text="Nb offres", command=self.on_count_offers).pack(side="left", padx=8)
        ttk.Button(btns, text="Lancer l’étude", command=self.on_run_study).pack(side="left", padx=8)
        ttk.Button(btns, text="finaliser", command=self.on_end_study).pack(side="left", padx=8)
        ttk.Button(btns, text="Effacer le journal", command=self.on_clear).pack(side="left", padx=8)
        ttk.Button(btns, text="Sauver paramètres", command=self.on_save).pack(side="left")
        ttk.Button(btns, text="Quitter", command=self.on_close).pack(side="right")

        # Journal
        self.log = LogText(container)
        self.log.pack(fill="both", expand=True, pady=(8, 0))

        # Remise en cohérence des contrôles avec les valeurs actuelles
        try:
            self.refresh_domaine_combobox()
        except Exception:
            pass
        try:
            self.update_location_mode_ui()
        except Exception:
            pass

    def reload_main_interface(self):
        """Reconstruit l'interface d'accueil sans perdre les paramètres."""
        for widget in self.winfo_children():
            widget.destroy()
        self._build_main_interface()

    def on_end_study(self):
        from finalisation import handle_end_study

        mode = self.var_loc_mode.get()
        code_lieux = f"{self.var_dept_code.get()}D" if mode == "departement" else self.var_lieux_commune.get()

        p = Params(
            domaine=self.var_domaine_code.get(),
            emission=self.var_emission.get(),
            lieux=code_lieux,
            rayon=self.var_rayon.get(),
            output_dir=self.var_outdir.get()
        )

        handle_end_study(self, p)

        





if __name__ == "__main__":
    app = App()
    if app.index_domaine.items:
        app.refresh_domaine_combobox()
    app.mainloop()

