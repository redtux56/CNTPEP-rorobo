#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Étape "partenaires problématiques" (partpb) — DevTools via WebDriver (CDP)
- Filtre CSV: 3e col (partenaire) != 'pe' ET ∈ partpb.txt (comparaison en minuscules)
- Affiche l’URL (4e col) dans le même Chrome (profil identique)
- Bouton "OK (imprimer + suivant)" : impression via CDP Page.printToPDF
  dans le dossier PDF *défini comme dans interface.py* :
  "pdf-{lieux}-{emission}jours-{domaine}-rayon{rayon}km"
  nommage: "{idx:04d}-par-{ref}.pdf"
- SKIP automatique (avant affichage et juste avant impression) si un fichier "{idx:04d}-par-*.pdf"
  existe déjà dans CE répertoire (pré-skip par préfixe).
- Si aucun partenaire problématique: on informe puis on enchaîne automatiquement l’étape "partenaires connus".
- Fin : fenêtre "Fin des partenaires problématiques" → OK lance l’étape "partenaires connus" (Tk.after).
"""

import os
import base64
import time
import re
import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd

from interface import Params, BG
import finalisation as fin


# ---------------------------------------------------------------------------
# Helpers: listes / chemins / noms
# ---------------------------------------------------------------------------

def _read_list_any(path: str) -> list[str]:
    """Lit un fichier de liste en JSON (liste) OU une valeur par ligne."""
    # 1) tentative JSON
    try:
        if os.path.exists(path):
            import json
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
                if isinstance(data, list):
                    return [str(x) for x in data]
    except Exception:
        pass
    # 2) fallback: une valeur par ligne
    vals = []
    try:
        with open(path, "r", encoding="utf-8") as f:
            for line in f:
                s = line.strip()
                if s:
                    vals.append(s)
    except Exception:
        pass
    return vals


def _compute_csv_path(params: Params) -> str:
    return os.path.join(
        params.output_dir or os.getcwd(),
        f"etud-{params.lieux}-{params.emission}jours-{params.domaine}-rayon{params.rayon}km.csv",
    )


def _compute_final_pdf_dir(params: Params) -> str:
    """Répertoire PDF calculé *exactement* comme dans interface.py."""
    base = params.output_dir or os.getcwd()
    pdf_dir = os.path.join(
        base,
        f"pdf-{params.lieux}-{params.emission}jours-{params.domaine}-rayon{params.rayon}km",
    )
    os.makedirs(pdf_dir, exist_ok=True)
    return pdf_dir


def _sanitize_filename(name: str) -> str:
    # Remplace caractères interdits Windows
    return re.sub(r'[<>:"/\\|?*]', "_", name)


def _has_existing_par_pdf_for_idx(pdf_dir: str, idx_etud: int) -> bool:
    """
    Skip robuste: si n'importe quel fichier correspondant au motif
    ^{idx:04d}-par-.*\.pdf (insensible à la casse) existe dans le dossier PDF, on considère que c'est déjà fait.
    """
    prefix = f"{idx_etud:04d}-par-"
    try:
        for fn in os.listdir(pdf_dir):
            if fn.startswith(prefix) and fn.lower().endswith(".pdf"):
                return True
    except FileNotFoundError:
        return False
    return False


# ---------------------------------------------------------------------------
# Dataset
# ---------------------------------------------------------------------------

def _build_target_df(csv_path: str, partpb_set: set) -> pd.DataFrame:
    """
    Garde colonnes [0,1,2,3] (idx, ref, partenaire, url) où partenaire != 'pe' et ∈ partpb_set (comparé en minuscules).
    """
    df = fin._load_df(csv_path)
    if df.empty or df.shape[1] < 4:
        return pd.DataFrame()

    partners_l = df.iloc[:, 2].astype(str).str.strip().str.lower()
    mask = partners_l.ne("pe") & partners_l.isin(partpb_set)
    if not mask.any():
        return pd.DataFrame()

    out = df.loc[mask, [0, 1, 2, 3]].copy().reset_index(drop=True)

    # Éventuel en-tête résiduel
    try:
        def _looks_like_header_row(values):
            kw = (
                "index", "indice", "ref", "réf", "référence", "reference",
                "partenaire", "nom_partenaire", "url", "url_partenaire", "type", "contrat"
            )
            vals = [str(v or "").strip().lower() for v in values]
            hits = sum(1 for t in kw if any(t in v for v in vals))
            return hits >= 2
        if len(out) and _looks_like_header_row(out.loc[0, [0, 1, 2, 3]].tolist()):
            out = out.iloc[1:].reset_index(drop=True)
    except Exception:
        pass

    for c in [0, 1, 2, 3]:
        out[c] = out[c].astype(str).str.strip()
    return out


# ---------------------------------------------------------------------------
# Impression DevTools (CDP)
# ---------------------------------------------------------------------------

def _wait_doc_ready(driver, timeout: float = 30.0) -> None:
    t0 = time.monotonic()
    while time.monotonic() - t0 < timeout:
        try:
            state = driver.execute_script("return document.readyState")
            if state == "complete":
                return
        except Exception:
            pass
        time.sleep(0.2)


def _wait_for_file(path: str, timeout: float = 10.0, stable_checks: int = 2) -> bool:
    t0 = time.monotonic()
    last = -1
    stable = 0
    while time.monotonic() - t0 < timeout:
        if os.path.exists(path):
            try:
                size = os.path.getsize(path)
            except Exception:
                size = -1
            if size == last and size > 0:
                stable += 1
                if stable >= stable_checks:
                    return True
            else:
                stable = 0
                last = size
        time.sleep(0.3)
    return os.path.exists(path) and os.path.getsize(path) > 0


def _devtools_print_to_file(driver, out_pdf: str, print_opts=None) -> bool:
    """
    Imprime la page courante via CDP "Page.printToPDF".
    Écrit d'abord un .tmp, puis remplace → pas de lecture partielle.
    Filet de sécurité : si `out_pdf` existe déjà, on ne (ré)imprime pas.
    """
    try:
        if os.path.exists(out_pdf):
            return True
    except Exception:
        if os.path.exists(out_pdf):
            return True

    opts = {
        "printBackground": True,
        "landscape": False,
        "scale": 1.0,
        "paperWidth": 8.27,   # A4 (inches)
        "paperHeight": 11.69, # A4 (inches)
        "marginTop": 0.4,
        "marginBottom": 0.4,
        "marginLeft": 0.4,
        "marginRight": 0.4,
    }
    if print_opts:
        opts.update(print_opts)
    try:
        res = driver.execute_cdp_cmd("Page.printToPDF", opts)
        data_b64 = res.get("data")
        if not data_b64:
            return False
        raw = base64.b64decode(data_b64)
        tmp = out_pdf + ".tmp"
        with open(tmp, "wb") as f:
            f.write(raw)
        try:
            os.replace(tmp, out_pdf)
        except Exception:
            import shutil
            shutil.move(tmp, out_pdf)
        return _wait_for_file(out_pdf, timeout=10.0)
    except Exception:
        return False


# ---------------------------------------------------------------------------
# UI + boucle
# ---------------------------------------------------------------------------

def begin(app, params: Params):
    """
    Étape 'partenaires problématiques' (manuel: 1 clic OK par ligne).
    Impression DevTools-only, fichiers dans le répertoire PDF *calculé comme interface.py*.
    """
    app._partpb_params = params  # mémoriser pour l’étape suivante

    # CSV + listes
    csv_path = _compute_csv_path(params)
    if not os.path.exists(csv_path):
        messagebox.showerror(
            "CSV introuvable",
            f"Impossible de trouver le CSV :\n{csv_path}\n\n"
            f"Vérifie les paramètres (lieux / émission / domaine / rayon / dossier de sortie)."
        )
        return

    target_dir = os.path.dirname(csv_path) or os.getcwd()
    partpb_path = os.path.join(target_dir, "partpb.txt")
    partpb_list = _read_list_any(partpb_path)
    partpb_set = {str(x).strip().lower() for x in (partpb_list or []) if str(x).strip()}

    df = _build_target_df(csv_path, partpb_set)

    # ⬇️ Si aucun partenaire problématique : informer + enchaîner automatiquement l'étape "connus"
    if df.empty:
        try:
            messagebox.showinfo(
                "Aucun partenaire problématique",
                "Aucune ligne du CSV n'appartient à partpb.txt (et ≠ 'pe').\n"
                "On passe à l'étape 'partenaires connus'."
            )
        except Exception:
            pass

        def _launch_next():
            try:
                import importlib
                partco_stage = importlib.import_module("partco_stage")
                partco_stage.begin(app, params)
            except Exception as e:
                try:
                    messagebox.showerror(
                        "Étape suivante",
                        f"Impossible de démarrer l'étape 'partenaires connus':\n{e}"
                    )
                except Exception:
                    pass

        app.after(0, _launch_next)
        return

    # Selenium (mêmes paramètres/profils que l’interface)
    driver = fin._get_existing_driver(app) or fin._make_new_driver(app, params)
    if not driver:
        return

    # Répertoire PDF EXACT (style interface.py)
    final_pdf_dir = _compute_final_pdf_dir(params)

    # État
    app._partpb_df = df
    app._partpb_idx = 0
    app._partpb_pdfdir = final_pdf_dir

    # === Pré-skip (avant affichage) : avancer jusqu'à une ligne sans PDF existant
    def _advance_to_next_without_existing_pdf():
        n = len(app._partpb_df)
        while app._partpb_idx < n:
            idx_val = int(str(app._partpb_df.iloc[app._partpb_idx, 0]).strip())
            if _has_existing_par_pdf_for_idx(app._partpb_pdfdir, idx_val):
                app._partpb_idx += 1
            else:
                break
        if app._partpb_idx >= n:
            _finish_window(app)
            return False
        return True

    # Fenêtre minimale
    win = tk.Toplevel(app)
    win.title("Partenaires problématiques")
    win.configure(bg=BG)
    win.geometry("440x210+0+0")
    win.resizable(False, False)

    rang_var = tk.StringVar(value="")
    ref_var = tk.StringVar(value="")
    part_var = tk.StringVar(value="")
    status_var = tk.StringVar(value="Prêt.")

    def _refresh_labels():
        i = app._partpb_idx
        n = len(app._partpb_df)
        q = app._partpb_df.iloc[i]
        rang_var.set(f"Rang : {i+1}/{n}")
        ref_var.set(f"Référence : {q[1]}")
        part_var.set(f"Partenaire : {q[2]}")

    def _open_current():
        i = app._partpb_idx
        q = app._partpb_df.iloc[i]
        url = str(q[3]).strip()
        if not url or not (url.startswith("http://") or url.startswith("https://")):
            url = "http://" + url
        try:
            driver.get(url)
            _wait_doc_ready(driver, timeout=30.0)
        except Exception as e:
            messagebox.showwarning("Ouverture URL", f"Impossible d'ouvrir : {url}\n{e}")

    def _advance_or_finish():
        app._partpb_idx += 1
        if not _advance_to_next_without_existing_pdf():
            try:
                win.destroy()
            except Exception:
                pass
            return
        _refresh_labels()
        _open_current()
        status_var.set("Prêt.")

    def _on_ok():
        # garde-fou: si entre-temps un PDF a été créé, resauter
        if not _advance_to_next_without_existing_pdf():
            try:
                win.destroy()
            except Exception:
                pass
            return

        i = app._partpb_idx
        q = app._partpb_df.iloc[i]
        idx_etud = int(str(q[0]).strip())
        ref_raw = str(q[1]).strip()
        ref_sane = _sanitize_filename(ref_raw)
        url = str(q[3]).strip()
        if not url.startswith(("http://", "https://")):
            url = "http://" + url

        # Pré-skip local : si un PDF {idx:04d}-par-*.pdf existe déjà, on passe
        if _has_existing_par_pdf_for_idx(app._partpb_pdfdir, idx_etud):
            _advance_or_finish()
            return

        out_final = os.path.join(app._partpb_pdfdir, f"{idx_etud:04d}-par-{ref_sane}.pdf")

        # Filet local: si la cible exacte existe déjà, skip
        if os.path.exists(out_final):
            _advance_or_finish()
            return

        status_var.set("Impression (DevTools)…")
        win.update_idletasks()

        ok = _devtools_print_to_file(driver, out_final)
        if not ok:
            status_var.set("Échec DevTools. Réessaye ou passe.")
            messagebox.showwarning("Impression", "Échec d'impression via DevTools (CDP).")
            return

        status_var.set("OK. Passage au suivant…")
        win.update_idletasks()
        _advance_or_finish()

    # Démarrage : *avant* d'afficher la première, sauter celles déjà traitées
    if not _advance_to_next_without_existing_pdf():
        try:
            win.destroy()
        except Exception:
            pass
        return

    # Layout
    frm = ttk.Frame(win, padding=12)
    frm.pack(fill="both", expand=True)

    ttk.Label(frm, textvariable=rang_var).pack(anchor="w")
    ttk.Label(frm, textvariable=ref_var).pack(anchor="w")
    ttk.Label(frm, textvariable=part_var).pack(anchor="w")
    ttk.Separator(frm, orient="horizontal").pack(fill="x", pady=8)
    ttk.Label(frm, textvariable=status_var).pack(anchor="w", pady=(0, 8))
    ttk.Button(frm, text="OK (imprimer + suivant)", command=_on_ok).pack(fill="x")

    _refresh_labels()
    _open_current()
    status_var.set("Prêt.")


def _finish_window(app):
    """
    Fin de l'étape 'partenaires problématiques' → enchaîne sur 'partenaires connus'.
    """
    top = tk.Toplevel(app)
    top.title("Fin des partenaires problématiques")
    top.configure(bg=BG)
    top.geometry("360x160+0+0")
    ttk.Label(top, text="Fin des partenaires problématiques").pack(pady=16)

    def _on_ok():
        try:
            top.destroy()
        except Exception:
            pass

        def _launch():
            try:
                import importlib
                partco_stage = importlib.import_module("partco_stage")
                params = getattr(app, "_partpb_params", None)
                partco_stage.begin(app, params)
            except Exception as e:
                try:
                    from tkinter import messagebox as _mb
                    _mb.showerror("Étape suivante",
                                  f"Impossible de démarrer l'étape partenaires connus:\n{e}")
                except Exception:
                    pass

        # important: planifier l'appel après la destruction de la fenêtre
        app.after(0, _launch)

    ttk.Button(top, text="OK", command=_on_ok).pack(pady=8)
