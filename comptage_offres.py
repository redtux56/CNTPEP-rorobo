#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""Comptage d'offres France Travail par lieu / rayon / domaine / période.

Programme complémentaire indépendant de l'étude d'offres.
Il réutilise les fichiers :
- communes.csv
- departement.csv
- domainepro.csv

Les recherches peuvent être sauvegardées dans recherches.json puis exportées en CSV.
"""

from __future__ import annotations

import csv
import json
import os
import re
import threading
import time
import unicodedata
from dataclasses import asdict, dataclass
from datetime import datetime
from typing import Callable, Dict, List, Optional
from urllib.parse import urlencode

import requests
from bs4 import BeautifulSoup

import tkinter as tk
from tkinter import filedialog, messagebox, ttk


# -----------------------------------------------------------------------------
# Configuration
# -----------------------------------------------------------------------------

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
COMMUNES_FILE = os.path.join(BASE_DIR, "communes.csv")
DEPARTEMENTS_FILE = os.path.join(BASE_DIR, "departement.csv")
DOMAINES_FILE = os.path.join(BASE_DIR, "domainepro.csv")
JSON_FILE = os.path.join(BASE_DIR, "recherches.json")

BASE_URL = "https://candidat.francetravail.fr/offres/recherche?"

# Valeurs proposées par France Travail pour le rayon autour d'une commune.
DISTANCES = ["0", "5", "10", "20", "30", "50", "100"]

# Paramètre URL "emission" utilisé par France Travail.
PERIODES = {
    "Un jour": "1",
    "Trois jours": "3",
    "Une semaine": "7",
    "Deux semaines": "14",
    "Un mois": "31",
    "Toutes les offres": "tout",
}

RED = "#E2001A"
YELLOW = "#FFD400"
DARK = "#111111"
BG = "#FFFFFF"

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/151.0 Safari/537.36"
    ),
    "Accept-Language": "fr-FR,fr;q=0.9,en;q=0.8",
}


# -----------------------------------------------------------------------------
# Modèle de données
# -----------------------------------------------------------------------------

@dataclass
class Recherche:
    statut: str = "À calculer"
    lieu: str = ""
    type_lieu: str = "commune"  # commune | departement
    code: str = ""               # INSEE commune ou DEP sans suffixe D
    distance: str = "10"
    domaine: str = ""
    code_domaine: str = ""
    periode: str = "Toutes les offres"
    emission: str = "tout"
    nombre_offres: Optional[int] = None
    date_calcul: str = ""

    @property
    def code_url_lieu(self) -> str:
        return f"{self.code}D" if self.type_lieu == "departement" else self.code


# -----------------------------------------------------------------------------
# Utilitaires
# -----------------------------------------------------------------------------

def normalize_text(value: str) -> str:
    value = (value or "").strip().lower()
    value = unicodedata.normalize("NFD", value)
    value = "".join(c for c in value if unicodedata.category(c) != "Mn")
    return " ".join(value.split())


def read_csv_dict(path: str) -> List[dict]:
    last_error = None
    for encoding in ("utf-8-sig", "utf-8", "cp1252"):
        try:
            with open(path, "r", encoding=encoding, newline="") as f:
                return list(csv.DictReader(f))
        except Exception as exc:
            last_error = exc
    raise last_error or RuntimeError(f"Impossible de lire {path}")


def build_url(recherche: Recherche) -> str:
    params = {
        "lieux": recherche.code_url_lieu,
        "rayon": recherche.distance,
        "offresPartenaires": "true",
        "range": "0-19",
        "tri": "0",
    }
    if recherche.code_domaine:
        params["domaine"] = recherche.code_domaine
    if recherche.emission and recherche.emission != "tout":
        params["emission"] = recherche.emission
    return BASE_URL + urlencode(params)


def extract_count_from_html(content: bytes | str) -> int:
    """Extrait le nombre exact affiché dans le H1 de France Travail."""
    if isinstance(content, bytes):
        content = content.decode("utf-8", errors="replace")
    soup = BeautifulSoup(content, "html.parser")

    zone = soup.find(class_="zone-resultats")
    if zone:
        h1 = zone.find("h1")
    else:
        h1 = soup.find("h1")

    if not h1:
        raise ValueError("Nombre d'offres introuvable dans la page")

    text = h1.get_text(" ", strip=True).replace("\xa0", " ")

    # Exemples : "2494 offres - Lorient" ou "21 398 offres - ..."
    match = re.search(r"([0-9][0-9\s.\u202f\u00a0]*)\s+offre", text, flags=re.I)
    if not match:
        raise ValueError(f"Compteur non reconnu : {text!r}")

    digits = re.sub(r"\D", "", match.group(1))
    if not digits:
        raise ValueError(f"Compteur vide : {text!r}")
    return int(digits)


def count_offers(recherche: Recherche, attempts: int = 2) -> int:
    url = build_url(recherche)
    last_error = None

    for attempt in range(1, attempts + 1):
        try:
            response = requests.get(url, headers=HEADERS, timeout=25)
            response.raise_for_status()
            return extract_count_from_html(response.content)
        except Exception as exc:
            last_error = exc
            if attempt < attempts:
                time.sleep(1.5)

    raise RuntimeError(str(last_error) if last_error else "Erreur inconnue")


# -----------------------------------------------------------------------------
# Autocomplétion
# -----------------------------------------------------------------------------

class AutoCompleteEntry(ttk.Entry):
    def __init__(
        self,
        master,
        search_fn: Callable[[str], List[dict]],
        label_fn: Callable[[dict], str],
        on_select: Callable[[dict], None],
        **kwargs,
    ):
        super().__init__(master, **kwargs)
        self.search_fn = search_fn
        self.label_fn = label_fn
        self.on_select = on_select
        self.popup: Optional[tk.Toplevel] = None
        self.listbox: Optional[tk.Listbox] = None
        self.suggestions: List[dict] = []

        self.bind("<KeyRelease>", self._on_key_release)
        self.bind("<Down>", self._move_down)
        self.bind("<Up>", self._move_up)
        self.bind("<Return>", self._confirm)
        self.bind("<Escape>", lambda _e: self._hide())
        self.bind("<FocusOut>", lambda _e: self.after(180, self._hide))

    def _on_key_release(self, event):
        if event.keysym in {"Down", "Up", "Return", "Escape"}:
            return
        self.event_generate("<<EntryChanged>>")
        query = self.get().strip()
        if not query:
            self._hide()
            return
        suggestions = self.search_fn(query)
        if suggestions:
            self._show(suggestions)
        else:
            self._hide()

    def _show(self, suggestions: List[dict]):
        self.suggestions = suggestions
        if self.popup is None:
            self.popup = tk.Toplevel(self)
            self.popup.wm_overrideredirect(True)
            self.popup.attributes("-topmost", True)
            self.listbox = tk.Listbox(self.popup, height=10)
            self.listbox.pack(fill="both", expand=True)
            self.listbox.bind("<ButtonRelease-1>", self._confirm)
            self.listbox.bind("<Return>", self._confirm)

        assert self.listbox is not None
        self.listbox.delete(0, "end")
        for item in suggestions:
            self.listbox.insert("end", self.label_fn(item))

        height = min(10, max(1, len(suggestions))) * 20 + 4
        width = max(self.winfo_width(), 380)
        x = self.winfo_rootx()
        y = self.winfo_rooty() + self.winfo_height()
        self.popup.geometry(f"{width}x{height}+{x}+{y}")

    def _hide(self):
        if self.popup is not None:
            self.popup.destroy()
            self.popup = None
            self.listbox = None
            self.suggestions = []

    def _move_down(self, _event):
        if not self.listbox:
            return "break"
        current = self.listbox.curselection()
        index = current[0] + 1 if current else 0
        index = min(index, self.listbox.size() - 1)
        self.listbox.selection_clear(0, "end")
        self.listbox.selection_set(index)
        self.listbox.activate(index)
        return "break"

    def _move_up(self, _event):
        if not self.listbox:
            return "break"
        current = self.listbox.curselection()
        index = current[0] - 1 if current else 0
        index = max(index, 0)
        self.listbox.selection_clear(0, "end")
        self.listbox.selection_set(index)
        self.listbox.activate(index)
        return "break"

    def _confirm(self, _event=None):
        if not self.listbox or not self.suggestions:
            return
        selection = self.listbox.curselection()
        index = selection[0] if selection else 0
        item = self.suggestions[index]
        self.on_select(item)
        self._hide()
        return "break"


# -----------------------------------------------------------------------------
# Application
# -----------------------------------------------------------------------------

class ComptageOffresApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Comptage des offres - France Travail")
        self.root.geometry("1280x780")
        self.root.minsize(1080, 650)
        self.root.configure(bg=BG)

        self.recherches: List[Recherche] = []
        self.modified = False
        self.calculating = False

        self.communes: List[dict] = []
        self.departements: List[dict] = []
        self.domaines: List[dict] = []

        self.selected_lieu: Optional[dict] = None
        self.selected_domaine: Optional[dict] = None
        self.editing_index: Optional[int] = None

        self.var_type_lieu = tk.StringVar(value="commune")
        self.var_lieu = tk.StringVar()
        self.var_distance = tk.StringVar(value="10")
        self.var_domaine = tk.StringVar()
        self.var_periode = tk.StringVar(value="Toutes les offres")
        self.var_progress = tk.StringVar(value="Prêt")
        self.var_total = tk.StringVar(value="TOTAL : indisponible")

        self._setup_style()
        self._load_reference_files()
        self._build_ui()
        self._refresh_table()

        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    # ----- style -------------------------------------------------------------

    def _setup_style(self):
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure("Title.TLabel", font=("Segoe UI", 18, "bold"), foreground=DARK, background=BG)
        style.configure("Section.TLabelframe", background=BG)
        style.configure("Section.TLabelframe.Label", font=("Segoe UI", 10, "bold"), foreground=DARK, background=BG)
        style.configure("Treeview", rowheight=25, font=("Segoe UI", 9))
        style.configure("Treeview.Heading", font=("Segoe UI", 9, "bold"))
        style.configure("Red.TButton", font=("Segoe UI", 9, "bold"))

    # ----- fichiers de référence --------------------------------------------

    def _load_reference_files(self):
        missing = [p for p in (COMMUNES_FILE, DEPARTEMENTS_FILE, DOMAINES_FILE) if not os.path.exists(p)]
        if missing:
            names = "\n".join(os.path.basename(p) for p in missing)
            messagebox.showerror(
                "Fichiers manquants",
                "Place les fichiers suivants dans le même dossier que comptage_offres.py :\n\n" + names,
            )
            return

        try:
            self.communes = []
            for row in read_csv_dict(COMMUNES_FILE):
                name = (row.get("nom_standard") or "").strip()
                code = (row.get("code_insee") or "").strip()
                if name and code:
                    self.communes.append({"name": name, "code": code, "norm": normalize_text(name)})

            dep_rows = read_csv_dict(DEPARTEMENTS_FILE)
            dep_fields = dep_rows[0].keys() if dep_rows else []
            lib_col = next((c for c in dep_fields if c.upper().startswith("LIBELLE")), None)
            if not lib_col:
                raise ValueError("Colonne LIBELLE introuvable dans departement.csv")
            self.departements = []
            for row in dep_rows:
                name = (row.get(lib_col) or "").strip()
                code = (row.get("DEP") or "").strip()
                if name and code:
                    self.departements.append({"name": name, "code": code, "norm": normalize_text(name)})

            self.domaines = []
            for row in read_csv_dict(DOMAINES_FILE):
                name = (row.get("domainepro") or "").strip()
                code = (row.get("codedomainepro") or "").strip()
                if name and code:
                    self.domaines.append({"name": name, "code": code, "norm": normalize_text(name)})
            self.domaines.sort(key=lambda x: x["norm"])
        except Exception as exc:
            messagebox.showerror("Erreur de lecture", str(exc))

    # ----- recherche autocomplete -------------------------------------------

    def _search_lieux(self, query: str) -> List[dict]:
        q = normalize_text(query)
        if not q:
            return []
        source = self.communes if self.var_type_lieu.get() == "commune" else self.departements

        exact_code = [x for x in source if x["code"].lower() == query.strip().lower()]
        if exact_code:
            return exact_code[:20]

        exact = [x for x in source if x["norm"] == q]
        prefix = [x for x in source if x["norm"].startswith(q) and x not in exact]
        contains = [x for x in source if q in x["norm"] and x not in exact and x not in prefix]
        return (exact + prefix + contains)[:20]

    def _search_domaines(self, query: str) -> List[dict]:
        q = normalize_text(query)
        if not q:
            return []
        exact_code = [x for x in self.domaines if x["code"].lower() == query.strip().lower()]
        exact = [x for x in self.domaines if x["norm"] == q]
        prefix = [x for x in self.domaines if x["norm"].startswith(q) and x not in exact]
        contains = [x for x in self.domaines if q in x["norm"] and x not in exact and x not in prefix]
        ordered = exact_code + exact + prefix + contains
        unique = []
        seen = set()
        for item in ordered:
            key = (item["name"], item["code"])
            if key not in seen:
                seen.add(key)
                unique.append(item)
        return unique[:20]

    def _select_lieu(self, item: dict):
        self.selected_lieu = item
        self.var_lieu.set(f"{item['name']} ({item['code']})")

    def _select_domaine(self, item: dict):
        self.selected_domaine = item
        self.var_domaine.set(f"{item['name']} — {item['code']}")

    def _invalidate_lieu(self, _event=None):
        typed = self.var_lieu.get().strip()
        if self.selected_lieu:
            expected = f"{self.selected_lieu['name']} ({self.selected_lieu['code']})"
            if typed != expected:
                self.selected_lieu = None

    def _invalidate_domaine(self, _event=None):
        typed = self.var_domaine.get().strip()
        if self.selected_domaine:
            expected = f"{self.selected_domaine['name']} — {self.selected_domaine['code']}"
            if typed != expected:
                self.selected_domaine = None

    # ----- interface ---------------------------------------------------------

    def _build_ui(self):
        container = ttk.Frame(self.root, padding=12)
        container.pack(fill="both", expand=True)

        header = ttk.Frame(container)
        header.pack(fill="x", pady=(0, 8))
        ttk.Label(header, text="Comptage des offres France Travail", style="Title.TLabel").pack(side="left")

        form = ttk.LabelFrame(container, text="Recherche", style="Section.TLabelframe", padding=10)
        form.pack(fill="x", pady=(0, 8))

        ttk.Label(form, text="Type de lieu").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        type_frame = ttk.Frame(form)
        type_frame.grid(row=1, column=0, sticky="w", padx=5)
        ttk.Radiobutton(type_frame, text="Commune", value="commune", variable=self.var_type_lieu,
                        command=self._on_type_lieu_changed).pack(side="left")
        ttk.Radiobutton(type_frame, text="Département", value="departement", variable=self.var_type_lieu,
                        command=self._on_type_lieu_changed).pack(side="left", padx=(8, 0))

        ttk.Label(form, text="Lieu").grid(row=0, column=1, sticky="w", padx=5, pady=5)
        self.ent_lieu = AutoCompleteEntry(
            form,
            search_fn=self._search_lieux,
            label_fn=lambda x: f"{x['name']} ({x['code']})",
            on_select=self._select_lieu,
            textvariable=self.var_lieu,
            width=34,
        )
        self.ent_lieu.grid(row=1, column=1, sticky="ew", padx=5)
        self.ent_lieu.bind("<<EntryChanged>>", self._invalidate_lieu, add="+")

        ttk.Label(form, text="Distance").grid(row=0, column=2, sticky="w", padx=5, pady=5)
        self.cb_distance = ttk.Combobox(
            form,
            values=[f"{d} km" for d in DISTANCES],
            state="readonly",
            width=10,
        )
        self.cb_distance.grid(row=1, column=2, sticky="w", padx=5)
        self.cb_distance.set("10 km")
        self.cb_distance.bind("<<ComboboxSelected>>", self._on_distance_changed)

        ttk.Label(form, text="Domaine").grid(row=0, column=3, sticky="w", padx=5, pady=5)
        self.ent_domaine = AutoCompleteEntry(
            form,
            search_fn=self._search_domaines,
            label_fn=lambda x: f"{x['name']} — {x['code']}",
            on_select=self._select_domaine,
            textvariable=self.var_domaine,
            width=37,
        )
        self.ent_domaine.grid(row=1, column=3, sticky="ew", padx=5)
        self.ent_domaine.bind("<<EntryChanged>>", self._invalidate_domaine, add="+")

        ttk.Label(form, text="Période").grid(row=0, column=4, sticky="w", padx=5, pady=5)
        self.cb_periode = ttk.Combobox(
            form,
            textvariable=self.var_periode,
            values=list(PERIODES.keys()),
            state="readonly",
            width=18,
        )
        self.cb_periode.grid(row=1, column=4, sticky="w", padx=5)

        form.columnconfigure(1, weight=1)
        form.columnconfigure(3, weight=1)

        edit_buttons = ttk.Frame(form)
        edit_buttons.grid(row=2, column=0, columnspan=5, sticky="w", padx=5, pady=(10, 0))
        self.btn_add = ttk.Button(edit_buttons, text="Ajouter", command=self.add_recherche)
        self.btn_add.pack(side="left")
        self.btn_update = ttk.Button(edit_buttons, text="Mettre à jour", command=self.update_recherche)
        self.btn_update.pack(side="left", padx=5)
        self.btn_cancel_edit = ttk.Button(edit_buttons, text="Annuler modification", command=self.clear_form)
        self.btn_cancel_edit.pack(side="left")

        table_frame = ttk.Frame(container)
        table_frame.pack(fill="both", expand=True)

        columns = (
            "statut", "lieu", "type", "code", "distance", "domaine", "code_domaine",
            "periode", "nombre", "date"
        )
        self.tree = ttk.Treeview(table_frame, columns=columns, show="headings", selectmode="browse")
        headings = {
            "statut": "Statut",
            "lieu": "Lieu",
            "type": "Type",
            "code": "Code",
            "distance": "Distance",
            "domaine": "Domaine",
            "code_domaine": "Code domaine",
            "periode": "Période",
            "nombre": "Nombre offres",
            "date": "Date calcul",
        }
        widths = {
            "statut": 105, "lieu": 170, "type": 90, "code": 70, "distance": 75,
            "domaine": 240, "code_domaine": 90, "periode": 115, "nombre": 95, "date": 90,
        }
        for col in columns:
            self.tree.heading(col, text=headings[col])
            self.tree.column(col, width=widths[col], minwidth=60, anchor="w")

        yscroll = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        xscroll = ttk.Scrollbar(table_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        yscroll.grid(row=0, column=1, sticky="ns")
        xscroll.grid(row=1, column=0, sticky="ew")
        table_frame.rowconfigure(0, weight=1)
        table_frame.columnconfigure(0, weight=1)
        self.tree.bind("<<TreeviewSelect>>", self._on_table_select)
        self.tree.bind("<Double-1>", lambda _e: self.load_selected_into_form())

        controls = ttk.Frame(container)
        controls.pack(fill="x", pady=(8, 0))

        self.btn_delete = ttk.Button(controls, text="Supprimer", command=self.delete_selected)
        self.btn_delete.pack(side="left")
        self.btn_duplicate = ttk.Button(controls, text="Dupliquer", command=self.duplicate_selected)
        self.btn_duplicate.pack(side="left", padx=4)
        self.btn_up = ttk.Button(controls, text="Monter", command=lambda: self.move_selected(-1))
        self.btn_up.pack(side="left", padx=(8, 4))
        self.btn_down = ttk.Button(controls, text="Descendre", command=lambda: self.move_selected(1))
        self.btn_down.pack(side="left")

        ttk.Separator(controls, orient="vertical").pack(side="left", fill="y", padx=10)
        self.btn_load = ttk.Button(controls, text="Charger", command=self.load_json)
        self.btn_load.pack(side="left")
        self.btn_save = ttk.Button(controls, text="Enregistrer", command=self.save_json)
        self.btn_save.pack(side="left", padx=4)
        self.btn_export = ttk.Button(controls, text="Exporter CSV", command=self.export_csv)
        self.btn_export.pack(side="left", padx=(4, 12))

        self.btn_calc = ttk.Button(controls, text="CALCULER TOUT", command=self.calculate_all)
        self.btn_calc.pack(side="right")

        bottom = ttk.Frame(container)
        bottom.pack(fill="x", pady=(8, 0))
        self.progress = ttk.Progressbar(bottom, mode="determinate")
        self.progress.pack(side="left", fill="x", expand=True)
        ttk.Label(bottom, textvariable=self.var_progress).pack(side="left", padx=10)
        ttk.Label(bottom, textvariable=self.var_total, font=("Segoe UI", 10, "bold")).pack(side="right")

        self._action_buttons = [
            self.btn_add, self.btn_update, self.btn_cancel_edit, self.btn_delete,
            self.btn_duplicate, self.btn_up, self.btn_down, self.btn_load,
            self.btn_save, self.btn_export, self.btn_calc,
        ]

    # ----- formulaire --------------------------------------------------------

    def _on_type_lieu_changed(self):
        self.selected_lieu = None
        self.var_lieu.set("")

    def _on_distance_changed(self, _event=None):
        value = self.cb_distance.get().replace("km", "").strip()
        if value in DISTANCES:
            self.var_distance.set(value)

    def clear_form(self):
        self.editing_index = None
        self.selected_lieu = None
        self.selected_domaine = None
        self.var_type_lieu.set("commune")
        self.var_lieu.set("")
        self.var_distance.set("10")
        self.cb_distance.set("10 km")
        self.var_domaine.set("")
        self.var_periode.set("Toutes les offres")

    def _form_to_recherche(self) -> Optional[Recherche]:
        if not self.selected_lieu:
            messagebox.showwarning("Lieu", "Sélectionne un lieu dans la liste proposée.")
            return None
        if not self.selected_domaine:
            messagebox.showwarning("Domaine", "Sélectionne un domaine dans la liste proposée.")
            return None

        distance = self.cb_distance.get().replace("km", "").strip()
        periode = self.var_periode.get()
        return Recherche(
            statut="À calculer",
            lieu=self.selected_lieu["name"],
            type_lieu=self.var_type_lieu.get(),
            code=self.selected_lieu["code"],
            distance=distance,
            domaine=self.selected_domaine["name"],
            code_domaine=self.selected_domaine["code"],
            periode=periode,
            emission=PERIODES[periode],
        )

    def add_recherche(self):
        recherche = self._form_to_recherche()
        if not recherche:
            return
        self.recherches.append(recherche)
        self._mark_modified()
        self._refresh_table(select_index=len(self.recherches) - 1)
        self.clear_form()

    def update_recherche(self):
        if self.editing_index is None:
            selected = self._selected_index()
            if selected is None:
                messagebox.showinfo("Modification", "Sélectionne d'abord une ligne à modifier.")
                return
            self.editing_index = selected

        new_recherche = self._form_to_recherche()
        if not new_recherche:
            return

        old = self.recherches[self.editing_index]
        # On conserve l'ancien résultat tant qu'un nouveau calcul n'a pas réussi.
        new_recherche.nombre_offres = old.nombre_offres
        new_recherche.date_calcul = old.date_calcul
        new_recherche.statut = "À recalculer" if old.nombre_offres is not None else "À calculer"

        index = self.editing_index
        self.recherches[index] = new_recherche
        self._mark_modified()
        self._refresh_table(select_index=index)
        self.clear_form()

    def load_selected_into_form(self):
        index = self._selected_index()
        if index is None:
            return
        r = self.recherches[index]
        self.editing_index = index
        self.var_type_lieu.set(r.type_lieu)
        self.selected_lieu = {"name": r.lieu, "code": r.code, "norm": normalize_text(r.lieu)}
        self.var_lieu.set(f"{r.lieu} ({r.code})")
        self.var_distance.set(r.distance)
        self.cb_distance.set(f"{r.distance} km")
        self.selected_domaine = {"name": r.domaine, "code": r.code_domaine, "norm": normalize_text(r.domaine)}
        self.var_domaine.set(f"{r.domaine} — {r.code_domaine}")
        self.var_periode.set(r.periode)

    # ----- tableau -----------------------------------------------------------

    def _selected_index(self) -> Optional[int]:
        selection = self.tree.selection()
        if not selection:
            return None
        try:
            return int(selection[0])
        except ValueError:
            return None

    def _on_table_select(self, _event=None):
        pass

    def _refresh_table(self, select_index: Optional[int] = None):
        for item in self.tree.get_children():
            self.tree.delete(item)

        for i, r in enumerate(self.recherches):
            nombre = "" if r.nombre_offres is None else str(r.nombre_offres)
            self.tree.insert(
                "",
                "end",
                iid=str(i),
                values=(
                    r.statut,
                    r.lieu,
                    "Commune" if r.type_lieu == "commune" else "Département",
                    r.code,
                    f"{r.distance} km",
                    r.domaine,
                    r.code_domaine,
                    r.periode,
                    nombre,
                    r.date_calcul,
                ),
            )

        if select_index is not None and 0 <= select_index < len(self.recherches):
            iid = str(select_index)
            self.tree.selection_set(iid)
            self.tree.focus(iid)
            self.tree.see(iid)

        self._update_total()

    def delete_selected(self):
        index = self._selected_index()
        if index is None:
            return
        del self.recherches[index]
        self._mark_modified()
        self._refresh_table(select_index=min(index, len(self.recherches) - 1) if self.recherches else None)
        self.clear_form()

    def duplicate_selected(self):
        index = self._selected_index()
        if index is None:
            return
        source = self.recherches[index]
        duplicate = Recherche(**asdict(source))
        duplicate.statut = "À calculer"
        duplicate.nombre_offres = None
        duplicate.date_calcul = ""
        self.recherches.insert(index + 1, duplicate)
        self._mark_modified()
        self._refresh_table(select_index=index + 1)
        self.load_selected_into_form()

    def move_selected(self, direction: int):
        index = self._selected_index()
        if index is None:
            return
        new_index = index + direction
        if new_index < 0 or new_index >= len(self.recherches):
            return
        self.recherches[index], self.recherches[new_index] = self.recherches[new_index], self.recherches[index]
        self._mark_modified()
        self._refresh_table(select_index=new_index)

    # ----- JSON --------------------------------------------------------------

    def save_json(self) -> bool:
        try:
            data = {
                "version": 1,
                "recherches": [asdict(r) for r in self.recherches],
            }
            with open(JSON_FILE, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            self.modified = False
            self.var_progress.set(f"Enregistré : {os.path.basename(JSON_FILE)}")
            return True
        except Exception as exc:
            messagebox.showerror("Enregistrement", f"Impossible d'enregistrer recherches.json :\n{exc}")
            return False

    def load_json(self):
        if self.modified:
            answer = messagebox.askyesnocancel(
                "Liste modifiée",
                "La liste actuelle a été modifiée. Voulez-vous l'enregistrer avant de charger ?",
            )
            if answer is None:
                return
            if answer and not self.save_json():
                return

        if not os.path.exists(JSON_FILE):
            messagebox.showinfo("Chargement", "recherches.json n'existe pas encore.")
            return

        try:
            with open(JSON_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
            rows = data.get("recherches", []) if isinstance(data, dict) else data
            loaded = []
            for row in rows:
                # Tolérance pour de futures/anciennes versions du JSON.
                allowed = Recherche.__dataclass_fields__.keys()
                filtered = {k: v for k, v in row.items() if k in allowed}
                loaded.append(Recherche(**filtered))
            self.recherches = loaded
            self.modified = False
            self.clear_form()
            self._refresh_table()
            self.var_progress.set(f"Chargé : {len(loaded)} recherche(s)")
        except Exception as exc:
            messagebox.showerror("Chargement", f"Impossible de charger recherches.json :\n{exc}")

    # ----- calcul ------------------------------------------------------------

    def calculate_all(self):
        if self.calculating:
            return
        if not self.recherches:
            messagebox.showinfo("Calcul", "Ajoute au moins une recherche.")
            return

        self.calculating = True
        self.progress["maximum"] = len(self.recherches)
        self.progress["value"] = 0
        self.var_total.set("TOTAL : indisponible")
        self._set_controls_enabled(False)

        thread = threading.Thread(target=self._calculate_worker, daemon=True)
        thread.start()

    def _calculate_worker(self):
        total = len(self.recherches)
        errors = 0

        for i, r in enumerate(self.recherches):
            r.statut = "Calcul en cours"
            self.root.after(0, self._refresh_table, i)
            self.root.after(0, self.var_progress.set, f"Recherche {i + 1} / {total} — {r.lieu} — {r.domaine}")

            try:
                number = count_offers(r, attempts=2)
                r.nombre_offres = number
                r.date_calcul = datetime.now().strftime("%d/%m/%Y")
                r.statut = "OK"
            except Exception:
                # On conserve un éventuel ancien résultat, mais le statut indique clairement l'échec.
                r.statut = "ERREUR"
                errors += 1

            self.root.after(0, self._progress_update, i + 1, i)

        self.modified = True
        self.root.after(0, self._calculation_finished, errors)

    def _progress_update(self, value: int, select_index: int):
        self.progress["value"] = value
        self._refresh_table(select_index=select_index)

    def _calculation_finished(self, errors: int):
        self.calculating = False
        self._set_controls_enabled(True)
        self._refresh_table()
        if errors:
            self.var_progress.set(f"Calcul incomplet — {errors} ligne(s) en erreur — TOTAL indisponible")
            self.var_total.set("TOTAL : indisponible")
        else:
            total = sum(r.nombre_offres or 0 for r in self.recherches)
            self.var_progress.set(f"Calcul terminé — {len(self.recherches)} recherche(s)")
            self.var_total.set(f"TOTAL : {total}")

    def _set_controls_enabled(self, enabled: bool):
        state = "normal" if enabled else "disabled"
        for button in self._action_buttons:
            button.configure(state=state)

    def _update_total(self):
        if not self.recherches:
            self.var_total.set("TOTAL : indisponible")
            return
        if any(r.statut != "OK" or r.nombre_offres is None for r in self.recherches):
            self.var_total.set("TOTAL : indisponible")
            return
        self.var_total.set(f"TOTAL : {sum(r.nombre_offres or 0 for r in self.recherches)}")

    # ----- CSV ---------------------------------------------------------------

    def export_csv(self):
        if not self.recherches:
            messagebox.showinfo("Export CSV", "Aucune recherche à exporter.")
            return
        if any(r.statut != "OK" or r.nombre_offres is None for r in self.recherches):
            messagebox.showwarning(
                "Export CSV",
                "Le CSV complet ne peut être exporté que lorsque toutes les lignes sont au statut OK.",
            )
            return

        default_name = f"comptage_offres_{datetime.now().strftime('%Y-%m-%d')}.csv"
        path = filedialog.asksaveasfilename(
            title="Exporter les résultats",
            initialdir=BASE_DIR,
            initialfile=default_name,
            defaultextension=".csv",
            filetypes=[("Fichier CSV", "*.csv"), ("Tous les fichiers", "*.*")],
        )
        if not path:
            return

        fieldnames = [
            "Statut", "Lieu", "Type", "Code", "Distance", "Domaine",
            "Code domaine", "Période", "Nombre offres", "Date calcul",
        ]

        try:
            with open(path, "w", encoding="utf-8-sig", newline="") as f:
                writer = csv.writer(f, delimiter=";")
                writer.writerow(fieldnames)
                for r in self.recherches:
                    writer.writerow([
                        r.statut,
                        r.lieu,
                        "Commune" if r.type_lieu == "commune" else "Département",
                        r.code,
                        f"{r.distance} km",
                        r.domaine,
                        r.code_domaine,
                        r.periode,
                        r.nombre_offres,
                        r.date_calcul,
                    ])
                writer.writerow(["TOTAL", "", "", "", "", "", "", "", sum(r.nombre_offres or 0 for r in self.recherches), ""])
            self.var_progress.set(f"CSV exporté : {os.path.basename(path)}")
        except Exception as exc:
            messagebox.showerror("Export CSV", f"Impossible d'écrire le CSV :\n{exc}")

    # ----- état / fermeture --------------------------------------------------

    def _mark_modified(self):
        self.modified = True

    def on_close(self):
        if self.calculating:
            messagebox.showinfo("Calcul en cours", "Le calcul doit se terminer avant de fermer le programme.")
            return
        if self.modified:
            answer = messagebox.askyesnocancel(
                "Liste modifiée",
                "La liste a été modifiée. Voulez-vous l'enregistrer avant de quitter ?",
            )
            if answer is None:
                return
            if answer and not self.save_json():
                return
        self.root.destroy()


# -----------------------------------------------------------------------------
# Démarrage
# -----------------------------------------------------------------------------

def main():
    root = tk.Tk()
    ComptageOffresApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
