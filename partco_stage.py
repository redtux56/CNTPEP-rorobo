#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Étape "partenaires connus" (auto) — DevTools only, avec pré-skip par préfixe
- Déclenchée depuis partpb_stage._finish_window(app) (OK)
- Filtre: 3e col (partenaire) != 'pe' AND ∈ partco.txt AND ∉ robot.txt
- Impression silencieuse via CDP Page.printToPDF (même driver Chrome)
- Nommage: "{idx:04d}-par-{ref}.pdf"
- Fenêtre compacte avec barre de progression X/Y + libellés (réf, partenaire, statut)
- Timeout 30s + 1 retry; si encore échec, propose une "capture manuelle" (Reprendre / Passer)
- Pré-skip: avant d'afficher/traiter, saute toutes les lignes dont un PDF "{idx:04d}-par-*.pdf" est déjà présent
  dans le répertoire PDF calculé comme dans interface.py
- À la fin: fenêtre "Fin des partenaires connus" → OK ⇒ lance merge_stage.begin(app, params)
"""

import os
import re
import time
import base64
import tkinter as tk
from tkinter import ttk, messagebox

import pandas as pd

from interface import Params, BG
import finalisation as fin

def _read_list_any(path: str) -> list[str]:
    """Lit un fichier de liste en JSON (liste) OU une valeur par ligne."""
    try:
        if os.path.exists(path):
            import json
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
                if isinstance(data, list):
                    return [str(x) for x in data]
    except Exception:
        pass
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
    return os.path.join(params.output_dir or os.getcwd(),
                        f"etud-{params.lieux}-{params.emission}jours-{params.domaine}-rayon{params.rayon}km.csv")

def _compute_final_pdf_dir(params: Params) -> str:
    base = params.output_dir or os.getcwd()
    pdf_dir = os.path.join(base, f"pdf-{params.lieux}-{params.emission}jours-{params.domaine}-rayon{params.rayon}km")
    os.makedirs(pdf_dir, exist_ok=True)
    return pdf_dir

def _sanitize_filename(name: str) -> str:
    return re.sub(r'[<>:"/\\|?*]', '_', name)

def _has_existing_par_pdf_for_idx(pdf_dir: str, idx_etud: int) -> bool:
    prefix = f"{idx_etud:04d}-par-"
    try:
        for fn in os.listdir(pdf_dir):
            if fn.startswith(prefix) and fn.lower().endswith(".pdf"):
                return True
    except FileNotFoundError:
        return False
    return False

def _build_target_df(csv_path: str, partco_set: set, robot_set: set) -> pd.DataFrame:
    df = fin._load_df(csv_path)
    if df.empty or df.shape[1] < 4:
        return pd.DataFrame()
    partners_l = df.iloc[:, 2].astype(str).str.strip().str.lower()
    mask = (partners_l.ne("pe")
            & partners_l.isin(partco_set)
            & ~partners_l.isin(robot_set))
    if not mask.any():
        return pd.DataFrame()
    out = df.loc[mask, [0, 1, 2, 3]].copy().reset_index(drop=True)
    try:
        def _looks_like_header_row(values):
            kw = ("index","indice","ref","réf","référence","reference","partenaire",
                  "nom_partenaire","url","url_partenaire","type","contrat")
            vals = [str(v or "").strip().lower() for v in values]
            hits = sum(1 for t in kw if any(t in v for v in vals))
            return hits >= 2
        if len(out) and _looks_like_header_row(out.loc[0, [0,1,2,3]].tolist()):
            out = out.iloc[1:].reset_index(drop=True)
    except Exception:
        pass
    for c in [0, 1, 2, 3]:
        out[c] = out[c].astype(str).str.strip()
    return out

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

def _devtools_print_to_file(driver, out_pdf: str, print_opts: dict | None = None) -> bool:
    opts = {
        "printBackground": True,
        "landscape": False,
        "scale": 1.0,
        "paperWidth": 8.27,
        "paperHeight": 11.69,
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
        import base64
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

def begin(app, params: Params):
    """Entrée de l'étape 'partenaires connus'."""
    csv_path = _compute_csv_path(params)
    if not os.path.exists(csv_path):
        messagebox.showerror("CSV introuvable",
                             f"Impossible de trouver le CSV :\n{csv_path}\n\n"
                             f"Vérifie les paramètres (lieux / émission / domaine / rayon / dossier de sortie).")
        return

    target_dir = os.path.dirname(csv_path) or os.getcwd()
    partco_list = _read_list_any(os.path.join(target_dir, "partco.txt"))
    robot_list  = _read_list_any(os.path.join(target_dir, "robot.txt"))
    partco_set = {s.strip().lower() for s in partco_list if str(s).strip()}
    robot_set  = {s.strip().lower() for s in robot_list if str(s).strip()}

    df = _build_target_df(csv_path, partco_set, robot_set)
    if df.empty:
        _finish_window(app, params)
        return

    driver = fin._get_existing_driver(app) or fin._make_new_driver(app, params)
    if not driver:
        return

    final_pdf_dir = _compute_final_pdf_dir(params)

    app._partco_df = df
    app._partco_idx = 0
    app._partco_total = len(df)
    app._partco_pdfdir = final_pdf_dir
    app._partco_params = params

    win = tk.Toplevel(app)
    win.title("Partenaires connus — traitement automatique")
    win.configure(bg=BG)
    win.geometry("540x210+0+210")
    win.resizable(False, False)

    lbl_head = tk.Label(win, text="Impression silencieuse (DevTools)", bg=BG, font=("Arial", 11, "bold"))
    lbl_head.pack(anchor="w", padx=12, pady=(10,4))

    frm = ttk.Frame(win, padding=12)
    frm.pack(fill="both", expand=True)

    var_prog = tk.StringVar(value=f"0 / {app._partco_total}")
    var_ref  = tk.StringVar(value="Réf : ")
    var_part = tk.StringVar(value="Partenaire : ")
    var_stat = tk.StringVar(value="Statut : prêt.")

    ttk.Label(frm, textvariable=var_prog).pack(anchor="w")
    ttk.Label(frm, textvariable=var_ref).pack(anchor="w")
    ttk.Label(frm, textvariable=var_part).pack(anchor="w")
    ttk.Label(frm, textvariable=var_stat).pack(anchor="w", pady=(4,6))

    pbar = ttk.Progressbar(frm, orient="horizontal", mode="determinate", length=480)
    pbar.pack(anchor="w", pady=(0,4))
    pbar["maximum"] = app._partco_total
    pbar["value"] = 0

    def _advance_to_next_without_existing_pdf():
        n = app._partco_total
        while app._partco_idx < n:
            row = app._partco_df.iloc[app._partco_idx]
            idx_val = int(str(row[0]).strip())
            if _has_existing_par_pdf_for_idx(app._partco_pdfdir, idx_val):
                app._partco_idx += 1
                pbar["value"] = app._partco_idx
                var_prog.set(f"{app._partco_idx} / {n}")
            else:
                break
        if app._partco_idx >= n:
            try:
                win.destroy()
            except Exception:
                pass
            _finish_window(app, params)
            return False
        return True

    def _update_labels(i):
        n = app._partco_total
        var_prog.set(f"{i} / {n}")
        if i < n:
            row = app._partco_df.iloc[i]
            var_ref.set(f"Réf : {row[1]}")
            var_part.set(f"Partenaire : {row[2]}")

    def _process_current():
        if not _advance_to_next_without_existing_pdf():
            return
        i = app._partco_idx
        n = app._partco_total
        _update_labels(i)

        row = app._partco_df.iloc[i]
        idx_etud = int(str(row[0]).strip())
        ref = _sanitize_filename(str(row[1]).strip())
        url = str(row[3]).strip()
        if not url.startswith(("http://", "https://")):
            url = "http://" + url

        out_final = os.path.join(app._partco_pdfdir, f"{idx_etud:04d}-par-{ref}.pdf")

        if _has_existing_par_pdf_for_idx(app._partco_pdfdir, idx_etud) or os.path.exists(out_final):
            app._partco_idx += 1
            pbar["value"] = app._partco_idx
            win.after(400, _process_current)
            return

        def _try_print_once() -> bool:
            try:
                var_stat.set("Statut : chargement page…")
                win.update_idletasks()
                driver.get(url)
                _wait_doc_ready(driver, timeout=30.0)
                var_stat.set("Statut : impression (DevTools)…")
                win.update_idletasks()
                ok = _devtools_print_to_file(driver, out_final)
                return ok
            except Exception:
                return False

        ok = _try_print_once()
        if not ok:
            var_stat.set("Statut : 2ᵉ essai…")
            win.update_idletasks()
            ok = _try_print_once()

        if not ok:
            def _manual_popup():
                top = tk.Toplevel(win)
                top.title("Capture manuelle")
                top.geometry("560x160+560+210")
                tk.Label(top, text="Échec d'impression après 2 essais.\nTu peux capturer la page manuellement si besoin, puis cliquer Reprendre.", justify="left").pack(pady=10)
                btns = ttk.Frame(top); btns.pack(pady=8)
                def _on_resume():
                    try:
                        var_stat.set("Statut : reprise (DevTools)…")
                        win.update_idletasks()
                        ok2 = _devtools_print_to_file(driver, out_final)
                    except Exception:
                        ok2 = False
                    top.destroy()
                    _finish_manual(ok2)
                def _on_skip():
                    top.destroy()
                    _finish_manual(False)
                ttk.Button(btns, text="Reprendre", command=_on_resume).pack(side="left", padx=6)
                ttk.Button(btns, text="Passer", command=_on_skip).pack(side="left", padx=6)
            def _finish_manual(success: bool):
                app._partco_idx += 1
                pbar["value"] = app._partco_idx
                win.after(400, _process_current)
            _manual_popup()
            return

        var_stat.set("Statut : OK.")
        app._partco_idx += 1
        pbar["value"] = app._partco_idx
        win.after(400, _process_current)

    if not _advance_to_next_without_existing_pdf():
        return
    _update_labels(app._partco_idx)
    win.after(300, _process_current)

def _finish_window(app, params: Params):
    top = tk.Toplevel(app)
    top.title("Fin des partenaires connus")
    top.configure(bg=BG)
    top.geometry("360x160+0+420")
    ttk.Label(top, text="Fin des partenaires connus").pack(pady=16)

    def _on_ok():
        try:
            top.destroy()
        except Exception:
            pass
        # ➜ Lancer la fusion
        def _launch_merge():
            try:
                import importlib
                merge_stage = importlib.import_module("merge_stage")
                merge_stage.begin(app, params)
            except Exception as e:
                try:
                    messagebox.showerror("Fusion PDF", f"Impossible de démarrer la fusion :\n{e}")
                except Exception:
                    pass
        app.after(0, _launch_merge)

    ttk.Button(top, text="OK", command=_on_ok).pack(pady=8)
