#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import json
import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd

from interface import Params, BG, RED  # couleurs + options Chrome depuis interface.py

# =============================== CSV utils ===================================

def _load_df(path: str) -> pd.DataFrame:
    """
    Lecture robuste SANS en-tête (header=None), positions uniquement.
    Essaie plusieurs encodages/séparateurs, tolère les lignes douteuses.
    """
    encodings = ("utf-8", "utf-8-sig", "cp1252", "latin1")
    seps = (",", ";", "\t", None)  # None => autodétection (engine='python')
    tries = []
    for enc in encodings:
        for sep in seps:
            try:
                kwargs = dict(encoding=enc, dtype=str, na_filter=False, header=None)
                if sep is None:
                    df = pd.read_csv(path, engine="python", sep=None, on_bad_lines="skip", **kwargs)
                else:
                    df = pd.read_csv(path, engine="python", sep=sep, on_bad_lines="skip", **kwargs)
                if df is not None:
                    df = df.dropna(how="all").reset_index(drop=True)
                    return df
            except Exception as e:
                tries.append(f"{enc}/sep={sep or 'auto'}: {e}")
    messagebox.showerror("Lecture CSV",
                         "Impossible de lire :\n"
                         f"{path}\n\nEssais :\n- " + "\n- ".join(tries[:20]))
    return pd.DataFrame()


def _looks_like_header_row(values: list[str]) -> bool:
    """
    Retourne True si la ligne ressemble à un en-tête (>=2 hits sur mots-clés).
    """
    kw = ("index", "indice", "ref", "réf", "référence", "reference",
          "partenaire", "nom_partenaire", "url", "url_partenaire", "type", "contrat")
    hits = 0
    for v in values:
        s = str(v).strip().lower()
        if any(k in s for k in kw):
            hits += 1
    return hits >= 2


def _strip_df2_headers(df2: pd.DataFrame) -> pd.DataFrame:
    """Supprime toute ligne d’en-tête résiduelle de df2 (ex. 'index, référence, partenaire, url')."""
    if df2 is None or df2.empty:
        return df2
    cols = [0, 1, 2, 3]
    cols = [c for c in cols if c in df2.columns]
    if not cols:
        return df2
    mask_header = df2[cols].apply(lambda row: _looks_like_header_row(row.values.tolist()), axis=1)
    if mask_header.any():
        df2 = df2.loc[~mask_header].reset_index(drop=True)
    return df2


def _make_df2(df: pd.DataFrame, partco_set: set) -> pd.DataFrame:
    """
    Construit df2 strictement en positions :
      - ≥ 4 colonnes (0..3)
      - filtre: 3e col (index 2, partenaire) != 'pe' (insensible casse)
               ET partenaire pas dans partco_set
    Renvoie colonnes 0..3 (index, référence, partenaire, url) sans en-tête résiduel.
    """
    if df is None or df.empty:
        return pd.DataFrame()
    if df.shape[1] < 4:
        messagebox.showwarning("Colonnes CSV",
                               "Le CSV doit comporter au moins 4 colonnes (0..3 : index, référence, partenaire, url).")
        return pd.DataFrame()

    partners = df.iloc[:, 2].astype(str).str.strip()
    mask_not_pe = partners.str.lower().ne("pe")
    mask_not_seen = ~partners.isin(partco_set) if partco_set else pd.Series(True, index=df.index)
    mask = (mask_not_pe & mask_not_seen)

    if not mask.any():
        return pd.DataFrame()

    df2 = df.loc[mask, [0, 1, 2, 3]].copy().reset_index(drop=True)
    for c in [0, 1, 2, 3]:
        df2[c] = df2[c].astype(str).str.strip()

    # Supprimer les éventuelles lignes d'en-tête résiduelles
    df2 = _strip_df2_headers(df2)
    return df2


# ============================ JSON list utils =================================

def _ensure_json_list(path: str) -> list:
    """
    Garantit l'existence d'un fichier JSON liste à `path`.
    - Crée le parent si besoin, crée le fichier avec [] s'il n'existe pas.
    - Si invalide, réinitialise à [].
    Retourne toujours une liste Python.
    """
    try:
        parent = os.path.dirname(path)
        if parent and not os.path.isdir(parent):
            os.makedirs(parent, exist_ok=True)
        if not os.path.exists(path):
            with open(path, "w", encoding="utf-8") as f:
                json.dump([], f, ensure_ascii=False, indent=2)
                f.flush(); os.fsync(f.fileno())
            return []
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        return data if isinstance(data, list) else []
    except Exception:
        try:
            with open(path, "w", encoding="utf-8") as f:
                json.dump([], f, ensure_ascii=False, indent=2)
                f.flush(); os.fsync(f.fileno())
        except Exception as e2:
            messagebox.showwarning("partco/partpb/robot", f"Échec création {os.path.abspath(path)} : {e2}")
        return []


def _save_json_list(path: str, items: list):
    items = list(dict.fromkeys([str(x).strip() for x in items if str(x).strip()]))
    try:
        with open(path, "w", encoding="utf-8") as f:
            json.dump(items, f, ensure_ascii=False, indent=2)
            f.flush(); os.fsync(f.fileno())
    except Exception as e:
        messagebox.showwarning("Sauvegarde", f"Impossible d’écrire {os.path.abspath(path)} : {e}")


# =============================== Selenium =====================================

try:
    from selenium import webdriver
    _selenium_import_error = None
except Exception as _e:
    webdriver = None
    _selenium_import_error = _e


def _is_driver_alive(driver) -> bool:
    try:
        _ = driver.window_handles
        return True
    except Exception:
        return False


def _get_existing_driver(app):
    for attr in ("browser", "_driver", "driver"):
        drv = getattr(app, attr, None)
        if drv is not None and _is_driver_alive(drv):
            return drv
    return None


def _make_new_driver(app, params: Params):
    """
    Crée un Chrome avec EXACTEMENT les mêmes dossiers que l'interface :
    - profil persistant dans <params.output_dir>/chrome-profile
    - téléchargements dans <params.output_dir>/pdfpar
    Utilise app._build_chrome_options(download_dir, profile_dir) si dispo.
    """
    if webdriver is None:
        messagebox.showerror("Selenium", f"Selenium indisponible : {_selenium_import_error}")
        return None
    try:
        outdir = (params.output_dir or "").strip() or os.getcwd()
        pdf_dir = os.path.join(outdir, "pdfpar")
        prof_dir = os.path.join(outdir, "chrome-profile")
        os.makedirs(pdf_dir, exist_ok=True)
        os.makedirs(prof_dir, exist_ok=True)

        if hasattr(app, "_build_chrome_options"):
            try:
                options = app._build_chrome_options(download_dir=pdf_dir, profile_dir=prof_dir)
            except Exception:
                options = None
        else:
            options = None

        if options is None:
            # Secours: garantir au minimum le profil persistant et le dossier de téléchargement
            options = webdriver.ChromeOptions()
            options.add_argument(f"--user-data-dir={prof_dir}")
            prefs = {
                "download.default_directory": pdf_dir,
                "savefile.default_directory": pdf_dir,
                "download.prompt_for_download": False,
                "download.directory_upgrade": True,
            }
            options.add_experimental_option("prefs", prefs)

        driver = webdriver.Chrome(options=options)
        app.browser = driver
        app._driver = driver
        try:
            if hasattr(app, "log"):
                app.log.write(f"Chrome initialisé ✓ (profil: {prof_dir}, téléchargements: {pdf_dir})")
        except Exception:
            pass
        return driver
    except Exception as e:
        messagebox.showerror("Selenium", f"Impossible de démarrer Chrome : {e}")
        return None


# =============================== UI helpers ===================================

def _ensure_info_vars(app):
    if not hasattr(app, "_info_vars"):
        app._info_vars = {
            "rang": tk.StringVar(value="Rang : —"),
            "ref": tk.StringVar(value="Référence : —"),
            "part": tk.StringVar(value="Partenaire : —"),
        }


def _refresh_info_vars(app):
    _ensure_info_vars(app)
    df2 = getattr(app, "_df2", None)
    i = getattr(app, "_row_idx", 0)
    if df2 is None or df2.empty or not (0 <= i < len(df2)):
        app._info_vars["rang"].set("Rang : —")
        app._info_vars["ref"].set("Référence : —")
        app._info_vars["part"].set("Partenaire : —")
        return
    n = len(df2)
    app._info_vars["rang"].set(f"Rang : {i+1}/{n}")
    app._info_vars["ref"].set(f"Référence : {df2.iloc[i, 1]}")
    app._info_vars["part"].set(f"Partenaire : {df2.iloc[i, 2]}")


def _hide_main_window(app):
    try:
        if app.state() != "withdrawn":
            app.withdraw()
            app._main_hidden = True
    except Exception:
        pass


def _restore_main_window(app):
    try:
        if getattr(app, "_main_hidden", False):
            app.deiconify()
            app.lift()
            app._main_hidden = False
    except Exception:
        pass


# =========================== Navigation / marquage ============================

def _navigate_to_current(app, params: Params):
    """Ouvre l'URL (col 3) de la ligne courante dans Selenium."""
    df2 = getattr(app, "_df2", None)
    i = getattr(app, "_row_idx", 0)
    if df2 is None or df2.empty or not (0 <= i < len(df2)):
        return

    url = str(df2.iloc[i, 3]).strip()
    if not url or url.lower() == "nan":
        messagebox.showinfo("URL manquante", "Pas d'URL exploitable pour cette ligne.")
        return
    if not (url.startswith("http://") or url.startswith("https://")):
        url = "http://" + url

    driver = _get_existing_driver(app) or _make_new_driver(app, params)
    if not driver:
        return
    try:
        driver.get(url)
    except Exception as e:
        messagebox.showwarning("Ouverture URL", f"Impossible d'ouvrir : {url}\n{e}")


def _advance_to_next_not_in_partco(app):
    """Avance tant que le partenaire courant est déjà dans partco."""
    df2 = getattr(app, "_df2", None)
    if df2 is None or df2.empty:
        return
    partco_set = set(getattr(app, "_partco_list", []))
    n = len(df2)
    while app._row_idx < n:
        partner = str(df2.iloc[app._row_idx, 2]).strip()
        if partner and partner in partco_set:
            app._row_idx += 1
        else:
            break


def _append_current_to_partco(app):
    """Ajoute le partenaire courant (col 2) à partco.txt si absent, puis sauvegarde."""
    df2 = getattr(app, "_df2", None)
    i = getattr(app, "_row_idx", 0)
    if df2 is None or df2.empty or not (0 <= i < len(df2)):
        return
    partner = str(df2.iloc[i, 2]).strip()
    if not partner or partner.lower() == "pe":
        return
    if partner not in app._partco_list:
        app._partco_list.append(partner)
        _save_json_list(app._partco_path, app._partco_list)


def _append_current_to_partpb(app):
    """Ajoute le partenaire courant (col 2) à partpb.txt si absent, puis sauvegarde."""
    df2 = getattr(app, "_df2", None)
    i = getattr(app, "_row_idx", 0)
    if df2 is None or df2.empty or not (0 <= i < len(df2)):
        return
    partner = str(df2.iloc[i, 2]).strip()
    if not partner or partner.lower() == "pe":
        return
    if partner not in app._partpb_list:
        app._partpb_list.append(partner)
        _save_json_list(app._partpb_path, app._partpb_list)


def _append_current_to_robot(app):
    """Ajoute le partenaire courant (col 2) à robot.txt si absent, puis sauvegarde."""
    df2 = getattr(app, "_df2", None)
    i = getattr(app, "_row_idx", 0)
    if df2 is None or df2.empty or not (0 <= i < len(df2)):
        return
    partner = str(df2.iloc[i, 2]).strip()
    if not partner or partner.lower() == "pe":
        return
    if partner not in app._robot_list:
        app._robot_list.append(partner)
        _save_json_list(app._robot_path, app._robot_list)


# ================================ Callbacks ===================================

def _open_or_focus_suivant_window(app, params: Params):
    win = getattr(app, "_suivant_win", None)
    if win and win.winfo_exists():
        try:
            _refresh_info_vars(app)
            win.lift()
            win.focus_force()
            return
        except Exception:
            pass
    win = tk.Toplevel(app)
    app._suivant_win = win
    win.title("Suite de l’étude")
    try:
        win.configure(bg=BG)
    except Exception:
        pass
    win.resizable(False, False)

    _ensure_info_vars(app)
    _refresh_info_vars(app)

    container = ttk.Frame(win, padding=12)
    container.pack(fill="both", expand=True)

    info = ttk.Frame(container)
    info.pack(fill="x", pady=(0, 10))
    ttk.Label(info, textvariable=app._info_vars["rang"]).pack(anchor="w")
    ttk.Label(info, textvariable=app._info_vars["ref"]).pack(anchor="w")
    ttk.Label(info, textvariable=app._info_vars["part"]).pack(anchor="w")

    ttk.Separator(container, orient="horizontal").pack(fill="x", pady=6)

    ttk.Button(container, text="➡  Site suivant",
               command=lambda: _handle_site_suivant(app, params)).pack(fill="x", pady=(0, 8))
    ttk.Button(container, text="⚠️  Problématique",
               command=lambda: _handle_problem(app, params)).pack(fill="x", pady=(0, 8))
    ttk.Button(container, text="🤖  Robot",
               command=lambda: _handle_robot(app, params)).pack(fill="x")

    def _on_close():
        _restore_main_window(app)
        try:
            win.destroy()
        except Exception:
            pass

    win.protocol("WM_DELETE_WINDOW", _on_close)
    win.update_idletasks()
    req_w, req_h = win.winfo_reqwidth(), win.winfo_reqheight()
    win.geometry(f"{req_w}x{req_h}+0+0")


def _handle_site_suivant(app, params: Params):
    """Ajoute partenaire courant à partco, avance au prochain non traité, ouvre son site."""
    df2 = getattr(app, "_df2", None)
    if df2 is None or df2.empty:
        messagebox.showinfo("Site suivant", "Aucune donnée.")
        return
    _append_current_to_partco(app)
    app._row_idx += 1
    _advance_to_next_not_in_partco(app)
    if app._row_idx >= len(df2):
        messagebox.showinfo("Fin", "Vous avez atteint la fin du fichier.")
        _start_partpb(app, params)
        return
        _start_partpb(app, params)
        return
    _refresh_info_vars(app)
    _navigate_to_current(app, params)


def _handle_problem(app, params: Params):
    """Ajoute partenaire courant à partco et partpb, avance au prochain non traité, ouvre son site."""
    df2 = getattr(app, "_df2", None)
    if df2 is None or df2.empty:
        messagebox.showinfo("Problématique", "Aucune donnée.")
        return
    _append_current_to_partco(app)
    _append_current_to_partpb(app)
    app._row_idx += 1
    _advance_to_next_not_in_partco(app)
    if app._row_idx >= len(df2):
        messagebox.showinfo("Fin", "Vous avez atteint la fin du fichier.")
        _start_partpb(app, params)
        return
        _start_partpb(app, params)
        return
    _refresh_info_vars(app)
    _navigate_to_current(app, params)


def _handle_robot(app, params: Params):
    """Ajoute partenaire courant à partco et robot, avance au prochain non traité, ouvre son site."""
    df2 = getattr(app, "_df2", None)
    if df2 is None or df2.empty:
        messagebox.showinfo("Robot", "Aucune donnée.")
        return
    _append_current_to_partco(app)
    _append_current_to_robot(app)
    app._row_idx += 1
    _advance_to_next_not_in_partco(app)
    if app._row_idx >= len(df2):
        messagebox.showinfo("Fin", "Vous avez atteint la fin du fichier.")
        _start_partpb(app, params)
        return
        _start_partpb(app, params)
        return
    _refresh_info_vars(app)
    _navigate_to_current(app, params)




def _start_partpb(app, params: Params):
    """Démarre l'étape 'partenaires problématiques' puis revient si erreur."""
    try:
        _hide_main_window(app)
    except Exception:
        pass
    try:
        import importlib
        partpb_stage = importlib.import_module("partpb_stage")
        partpb_stage.begin(app, params)
    except Exception as e:
        try:
            from tkinter import messagebox as _mb
            _mb.showerror("Étape suivante", f"Impossible de démarrer l'étape partenaires problématiques:\n{e}")
        except Exception:
            pass

# ================================ Entrée publique =============================

def handle_end_study(app, params: Params):
    """
    - Lit le CSV strictement en positions (sans en-tête).
    - Crée/charge partco.txt, partpb.txt et robot.txt **dans le même dossier que le CSV**
      (donc dossier défini par interface.py via params.output_dir).
    - Construit df2 en excluant :
        * partenaire == 'pe' (col 2)
        * partenaire déjà dans partco
      puis supprime toute ligne d’en-tête résiduelle.
    - Prépare UI + mini-fenêtre, et ouvre Selenium sur la 1re ligne éligible.
    """
    # 1) Localise le CSV
    csv_filename = f"etud-{params.lieux}-{params.emission}jours-{params.domaine}-rayon{params.rayon}km.csv"
    csv_path = os.path.normpath(os.path.join(params.output_dir, csv_filename))
    fichier_trouve = os.path.exists(csv_path)

    # 2) Reset UI + logo
    for w in app.winfo_children():
        w.destroy()
    if hasattr(app, "_load_and_place_logo"):
        try:
            holder = ttk.Frame(app, style="TFrame")
            holder.pack(pady=32)
            app._load_and_place_logo(parent=holder)
        except Exception:
            pass

    if not fichier_trouve:
        tk.Label(app, text="❌ Le fichier n'existe pas.", font=("Arial", 20, "bold"), fg=RED, bg=BG).pack(pady=20)
        ttk.Button(app, text="⬅ Retour à l'accueil", command=app.reload_main_interface).pack(pady=12)
        return

    # 3) Dossier cible = même dossier que le CSV (=> params.output_dir)
    target_dir = os.path.dirname(csv_path) or os.getcwd()
    app._partco_path = os.path.join(target_dir, "partco.txt")
    app._partpb_path = os.path.join(target_dir, "partpb.txt")
    app._robot_path  = os.path.join(target_dir, "robot.txt")

    # 4) Lecture CSV (positions)
    df = _load_df(csv_path)
    if df.empty or df.shape[1] < 4:
        tk.Label(app, text="⚠️ CSV vide ou incomplet (≥ 4 colonnes requises).",
                 font=("Arial", 16, "bold"), fg="#996600", bg=BG).pack(pady=18)
        ttk.Button(app, text="⬅ Retour à l'accueil", command=app.reload_main_interface).pack(pady=12)
        return

    # 5) Création/chargement des fichiers partenaires **dans target_dir**
    app._partco_list = _ensure_json_list(app._partco_path)
    app._partpb_list = _ensure_json_list(app._partpb_path)
    app._robot_list  = _ensure_json_list(app._robot_path)

    # 6) Filtre df -> df2 (et suppression d’un éventuel en-tête dans df2)
    app._df = df
    app._df2 = _make_df2(df, set(app._partco_list))
    app._row_idx = 0
    app._suivant_win = None

    if app._df2 is None or app._df2.empty:
        _start_partpb(app, params)
        return

    # 7) Réutilise driver existant si déjà lancé par interface.py
    drv = _get_existing_driver(app)
    if drv:
        app.browser = drv
        app._driver = drv

    # Affiche le dossier effectif pour contrôle visuel
    tk.Label(app,
             text=f"Dossier de travail :\n{target_dir}\n(partco.txt / partpb.txt / robot.txt seront ici)\n"
                  f"chrome-profile/pdfpar utilisés pour Selenium.",
             font=("Arial", 9), fg="#444", bg=BG).pack(pady=(0, 6))

    tk.Label(app, text="✅ Le fichier existe.", font=("Arial", 20, "bold"), fg="#008000", bg=BG).pack(pady=10)
    ttk.Button(
        app,
        text="Suivant ▶",
        command=lambda: (_refresh_info_vars(app), _navigate_to_current(app, params), _open_or_focus_suivant_window(app, params), _hide_main_window(app))
    ).pack(pady=10)
