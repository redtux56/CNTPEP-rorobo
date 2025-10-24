#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Étape de fusion des PDF après 'partenaires connus', puis export Excel basé sur le modèle.

Déclenchement attendu : depuis partco_stage._finish_window(app), bouton OK => merge_stage.begin(app, params)

Règles de fusion (ordre = CSV, fichiers dans le dossier PDF calculé comme interface.py) :
- partner == 'pe'                     -> ajouter: "{idx:04d}-{ref_sane}.pdf"
- partner ∈ partco ET ∈ robot        -> ajouter: "{idx:04d}-{ref_sane}.pdf"
- partner ∈ partco ET ∉ robot        -> ajouter: "{idx:04d}-{ref_sane}.pdf" puis "{idx:04d}-par-{ref_sane}.pdf"
- sinon (ni 'pe', ni partco)         -> SKIP

Excel (mapping, mise en page et mise en forme) :
- Onglet unique, écriture à partir de la ligne 2 (ignorer la 1ʳᵉ ligne de df = en-têtes)
- A ← df col 1 (nombre), B ← df col 2 (texte), C ← df col 3 (texte), D ← df col 5 (type libre)
- E:R ← cellules vides
- S ← formule par ligne : =COUNTIF(F{n}:R{n},1)>0
- Ligne "total" finale sous les données :
    A="total", B-D vides,
    E:R = SUM(E2:E{n}) … SUM(R2:R{n}),
    S = COUNTIF(S2:S{n},TRUE)
- Mise en forme :
    - Zébrage (A:S) sur les lignes de données : ligne 2 grise (gris 25 %), alternance gris/blanc
    - Bordures : grille fine sur toutes les cellules de données (A2:S{dernière_donnée})
    - Ligne "total" : grille fine + bordure supérieure épaisse
- Mise en page impression :
    - Répéter la ligne 1 en haut de chaque page
    - Zone d’impression : A1:S{dernière_ligne_totaux}
    - Orientation paysage, Ajuster à 1 page en largeur (hauteur libre)
- Fichier de sortie : fusion-{lieux}-{emission}jours-{domaine}-rayon{rayon}km.xlsx (écraser si existe)
"""

import os
import re
import tkinter as tk
from tkinter import ttk, messagebox

from interface import Params, BG
import finalisation as fin

# PyPDF (pypdf ou PyPDF2)
HAS_PYPDF = True
try:
    from pypdf import PdfReader, PdfWriter
except Exception:
    try:
        from PyPDF2 import PdfReader, PdfWriter  # type: ignore
    except Exception:
        HAS_PYPDF = False

# Excel (openpyxl)
HAS_OXL = True
try:
    import openpyxl
    from openpyxl import load_workbook
    from openpyxl.utils import get_column_letter
    from openpyxl.styles import PatternFill, Border, Side
except Exception:
    HAS_OXL = False


def _read_list_any(path: str) -> list[str]:
    """Lit JSON (liste) ou une valeur par ligne."""
    # JSON d'abord
    try:
        if os.path.exists(path):
            import json
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
                if isinstance(data, list):
                    return [str(x) for x in data]
    except Exception:
        pass
    # Lignes
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


def _compute_pdf_dir(params: Params) -> str:
    base = params.output_dir or os.getcwd()
    pdf_dir = os.path.join(base, f"pdf-{params.lieux}-{params.emission}jours-{params.domaine}-rayon{params.rayon}km")
    os.makedirs(pdf_dir, exist_ok=True)
    return pdf_dir


def _sanitize_filename(name: str) -> str:
    return re.sub(r'[<>:"/\\|?*]', "_", name)


def _add_bookmark(writer, title: str, page_index: int):
    """Compat pypdf / PyPDF2."""
    try:
        add = getattr(writer, "add_outline_item", None)
        if add:  # pypdf
            add(title, page_index)
            return
        add2 = getattr(writer, "addBookmark", None)
        if add2:  # PyPDF2
            add2(title, page_index)
    except Exception:
        pass


def _export_excel_from_df(params: Params) -> str:
    """
    Construit l'Excel à partir du CSV (df) et du modèle 'grilledebase.xlsx'.
    Retourne le chemin du fichier Excel écrit. Affiche une erreur et renvoie "" en cas d'échec.
    """
    if not HAS_OXL:
        messagebox.showerror("Export Excel", "openpyxl n'est pas disponible. Installe 'openpyxl'.")
        return ""

    csv_path = _compute_csv_path(params)
    pdf_dir = _compute_pdf_dir(params)
    out_xlsx = os.path.join(pdf_dir, f"fusion-{params.lieux}-{params.emission}jours-{params.domaine}-rayon{params.rayon}km.xlsx")

    # Charger df
    df = fin._load_df(csv_path)
    if df.empty or df.shape[1] < 5:
        messagebox.showerror("Export Excel", "Le CSV est vide ou invalide (moins de 5 colonnes).")
        return ""

    # Ignorer la 1ʳᵉ ligne (en-têtes)
    if len(df) >= 1:
        df_data = df.iloc[1:].reset_index(drop=True)
    else:
        df_data = df.copy()

    # Chercher le modèle
    template_candidates = [
        os.path.join(os.getcwd(), "grilledebase.xlsx"),
        os.path.join(os.path.dirname(__file__) if '__file__' in globals() else os.getcwd(), "grilledebase.xlsx"),
        os.path.join(os.path.dirname(csv_path), "grilledebase.xlsx"),
    ]
    template_path = ""
    for p in template_candidates:
        if os.path.exists(p):
            template_path = p
            break
    if not template_path:
        messagebox.showerror("Export Excel", "Modèle 'grilledebase.xlsx' introuvable. Place-le à la racine du projet.")
        return ""

    # Charger le modèle
    try:
        wb = load_workbook(template_path)
    except Exception as e:
        messagebox.showerror("Export Excel", f"Impossible d'ouvrir le modèle Excel :\n{e}")
        return ""
    ws = wb.active  # un seul onglet

    # Écriture à partir de la ligne 2
    start_row = 2
    n = len(df_data)

    for i in range(n):
        row_excel = start_row + i
        row = df_data.iloc[i]

        # A (nombre si possible)
        try:
            a_val = int(str(row.iloc[0]).strip())
        except Exception:
            try:
                a_val = float(str(row.iloc[0]).strip())
            except Exception:
                a_val = row.iloc[0]
        ws.cell(row=row_excel, column=1, value=a_val)  # A

        ws.cell(row=row_excel, column=2, value=str(row.iloc[1]))  # B
        ws.cell(row=row_excel, column=3, value=str(row.iloc[2]))  # C
        ws.cell(row=row_excel, column=4, value=row.iloc[4])       # D (df col 5)

        # E:R vides
        for col in range(5, 19):  # E..R
            ws.cell(row=row_excel, column=col, value=None)

        # S : =COUNTIF(F{n}:R{n},1)>0
        formula = f"=COUNTIF(F{row_excel}:R{row_excel},1)>0"
        ws.cell(row=row_excel, column=19, value=formula)  # S

    # ----- Mise en forme des lignes de données (zébrage + bordures) -----
    if n > 0:
        data_first = start_row
        data_last = start_row + n - 1
        # Styles
        zebra_fill = PatternFill(fill_type='solid', start_color='FFD9D9D9', end_color='FFD9D9D9')  # Gris 25 %
        thin = Side(style='thin')
        thin_border = Border(left=thin, right=thin, top=thin, bottom=thin)

        # Zébrage A:S (ligne 2 grise puis alternance) + grille fine
        for r in range(data_first, data_last + 1):
            is_grey = ((r - data_first) % 2 == 0)  # ligne 2 grise
            for c in range(1, 19 + 1):  # A=1 ... S=19
                cell = ws.cell(row=r, column=c)
                if is_grey:
                    cell.fill = zebra_fill
                # grille fine
                cell.border = thin_border

    # Ligne des totaux (si n > 0)
    if n > 0:
        total_row = start_row + n
        ws.cell(row=total_row, column=1, value="total")  # A
        # B,C,D vides
        ws.cell(row=total_row, column=2, value=None)
        ws.cell(row=total_row, column=3, value=None)
        ws.cell(row=total_row, column=4, value=None)
        # E:R = SUM
        for col in range(5, 19):  # E..R
            col_letter = get_column_letter(col)
            ws.cell(row=total_row, column=col, value=f"=SUM({col_letter}{start_row}:{col_letter}{total_row-1})")
        # S = COUNTIF(..., TRUE)
        ws.cell(row=total_row, column=19, value=f"=COUNTIF(S{start_row}:S{total_row-1},TRUE)")

        # Bordures de la ligne total : grille fine + bordure supérieure épaisse
        thin = Side(style='thin')
        medium = Side(style='medium')
        for c in range(1, 19 + 1):  # A..S
            cell = ws.cell(row=total_row, column=c)
            cell.border = Border(
                left=thin,
                right=thin,
                top=medium,   # démarcation
                bottom=thin
            )
        last_row = total_row
    else:
        last_row = start_row  # pas de ligne total si aucune donnée

    # Mise en page impression (+ nettoyage Print_Titles pour éviter réparation)
    try:
        # Supprimer d'éventuels Print_Titles existants (évite conflits definedNames)
        try:
            dn_keep = []
            for dn in getattr(wb.defined_names, "definedName", []):
                if getattr(dn, "name", "") != "_xlnm.Print_Titles":
                    dn_keep.append(dn)
            if hasattr(wb.defined_names, "definedName"):
                wb.defined_names.definedName = dn_keep
        except Exception:
            pass

        # Répéter la ligne 1
        ws.print_title_rows = "1:1"
        # Zone d'impression
        ws.print_area = f"A1:S{last_row}"
        # Orientation paysage + ajuster à 1 page en largeur
        ws.page_setup.orientation = 'landscape'
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
    except Exception:
        pass

    # Recalcul complet à l'ouverture (évite incohérences initiales)
    try:
        wb.calc_properties.fullCalcOnLoad = True
    except Exception:
        pass

    # Écriture du fichier (écrasement autorisé)
    try:
        if os.path.exists(out_xlsx):
            os.remove(out_xlsx)
    except Exception:
        pass

    try:
        wb.save(out_xlsx)
    except Exception as e:
        messagebox.showerror("Export Excel", f"Erreur lors de l'enregistrement du fichier Excel :\n{e}")
        return ""

    return out_xlsx


def begin(app, params: Params):
    """Démarre l’étape de fusion (et l’export Excel) avec UI et progression.
       Ajout: si le PDF fusionné existe déjà, proposer Écraser/Refusionner ou Passer à l’Excel.
    """
    if not HAS_PYPDF:
        messagebox.showerror("Fusion PDF", "Impossible d'importer pypdf/PyPDF2.\nInstalle 'pypdf' ou 'PyPDF2'.")
        return

    csv_path = _compute_csv_path(params)
    if not os.path.exists(csv_path):
        messagebox.showerror("CSV introuvable",
                             f"Impossible de trouver le CSV :\n{csv_path}\n\n"
                             f"Vérifie les paramètres (lieux / émission / domaine / rayon / dossier de sortie).")
        return

    pdf_dir = _compute_pdf_dir(params)
    out_pdf = os.path.join(
        pdf_dir,
        f"fusion-{params.lieux}-{params.emission}jours-{params.domaine}-rayon{params.rayon}km.pdf"
    )

    # Si un fichier fusionné existe déjà, proposer l’action (deux boutons)
    if os.path.exists(out_pdf):
        overwrite = messagebox.askyesno(
            "Fichier fusionné déjà présent",
            "Le fichier de fusion existe déjà.\n\n"
            "Voulez-vous l'écraser et relancer la fusion ?\n\n"
            "Oui = Écraser et refusionner\n"
            "Non = Passer directement à l'export Excel"
        )
        if not overwrite:
            # Passer directement à l'export Excel (sans toucher au PDF existant)
            out_xlsx = _export_excel_from_df(params)
            if out_xlsx:
                messagebox.showinfo("Export Excel terminé", f"Fichier créé :\n{out_xlsx}")
            try:
                app.destroy()
            except Exception:
                pass
            return
        else:
            # Écraser = supprimer l'ancien fichier avant de refusionner
            try:
                os.remove(out_pdf)
            except Exception as e:
                messagebox.showerror("Fusion PDF", f"Impossible d'écraser le fichier existant :\n{out_pdf}\n{e}")
                return

    # Lire listes
    target_dir = os.path.dirname(csv_path) or os.getcwd()
    partco_set = {s.strip().lower() for s in _read_list_any(os.path.join(target_dir, "partco.txt")) if str(s).strip()}
    robot_set  = {s.strip().lower() for s in _read_list_any(os.path.join(target_dir, "robot.txt")) if str(s).strip()}

    # Charger CSV et construire l'ordre
    df = fin._load_df(csv_path)
    if df.empty or df.shape[1] < 4:
        messagebox.showerror("Fusion PDF", "Le CSV est vide ou invalide (moins de 4 colonnes).")
        return

    merge_items: list[tuple[str, int, str, str, bool]] = []  # (path, idx, ref, partner, is_par)

    for _, row in df.iterrows():
        try:
            idx = int(str(row.iloc[0]).strip())
        except Exception:
            continue
        ref = str(row.iloc[1]).strip()
        partner = str(row.iloc[2]).strip()
        partner_l = partner.lower()
        ref_sane = _sanitize_filename(ref)

        # Règles de fusion
        if partner_l == "pe":
            p1 = os.path.join(pdf_dir, f"{idx:04d}-{ref_sane}.pdf")
            merge_items.append((p1, idx, ref, partner, False))
            continue

        if partner_l not in partco_set:
            continue  # SKIP

        if partner_l in robot_set:
            p1 = os.path.join(pdf_dir, f"{idx:04d}-{ref_sane}.pdf")
            merge_items.append((p1, idx, ref, partner, False))
        else:
            p1 = os.path.join(pdf_dir, f"{idx:04d}-{ref_sane}.pdf")
            p2 = os.path.join(pdf_dir, f"{idx:04d}-par-{ref_sane}.pdf")
            merge_items.append((p1, idx, ref, partner, False))
            merge_items.append((p2, idx, ref, partner, True))

    # Si rien à fusionner, tenter quand même l'export Excel puis fermer l'appli
    if not merge_items:
        out_xlsx = _export_excel_from_df(params)
        if out_xlsx:
            messagebox.showinfo("Export Excel terminé", f"Fichier créé :\n{out_xlsx}")
        try:
            app.destroy()
        except Exception:
            pass
        return

    # Vérifier que tous les fichiers existent
    missing = [p for (p, *_rest) in merge_items if not os.path.exists(p)]
    if missing:
        msg = "PDF manquant — fusion interrompue :\n" + "\n".join(missing[:20])
        if len(missing) > 20:
            msg += f"\n(+ {len(missing)-20} autres…)"
        messagebox.showerror("Fusion PDF", msg)
        return

    # UI de fusion
    win = tk.Toplevel(app)
    win.title("Fusion des PDF")
    win.configure(bg=BG)
    win.geometry("560x240+0+630")
    win.resizable(False, False)

    tk.Label(win, text="Fusion des PDF (ordre CSV)", bg=BG, font=("Arial", 11, "bold")).pack(anchor="w", padx=12, pady=(10,4))

    frm = ttk.Frame(win, padding=12)
    frm.pack(fill="both", expand=True)

    total = len(merge_items)
    var_prog = tk.StringVar(value=f"0 / {total}")
    var_file = tk.StringVar(value="Fichier : ")
    var_stat = tk.StringVar(value="Statut : prêt.")

    ttk.Label(frm, textvariable=var_prog).pack(anchor="w")
    ttk.Label(frm, textvariable=var_file).pack(anchor="w")
    ttk.Label(frm, textvariable=var_stat).pack(anchor="w", pady=(4,6))

    pbar = ttk.Progressbar(frm, orient="horizontal", mode="determinate", length=520)
    pbar.pack(anchor="w", pady=(0,6))
    pbar["maximum"] = total
    pbar["value"] = 0

    writer = PdfWriter()
    page_offset = 0  # index de la première page du doc courant dans le writer

    state = {"i": 0}

    def _finish_all():
        # Fermer fenêtre de fusion si ouverte
        try:
            win.destroy()
        except Exception:
            pass
        # Export Excel
        out_xlsx = _export_excel_from_df(params)
        if out_xlsx:
            messagebox.showinfo("Export Excel terminé", f"Fichier créé :\n{out_xlsx}")
        # Fermer l'application
        try:
            app.destroy()
        except Exception:
            pass

    def _step():
        nonlocal page_offset
        i = state["i"]
        if i >= total:
            # Écrire le fichier fusionné
            try:
                with open(out_pdf, "wb") as f:
                    writer.write(f)
            except Exception as e:
                messagebox.showerror("Fusion PDF", f"Erreur lors de l'écriture du PDF fusionné :\n{e}")
                _finish_all()
                return
            # Enchaîner export Excel puis fin
            _finish_all()
            return

        path, idx, ref, partner, is_par = merge_items[i]
        var_file.set(f"Fichier : {os.path.basename(path)}")
        var_stat.set("Statut : ajout…")
        win.update_idletasks()

        # Append
        try:
            reader = PdfReader(path)
            nb = len(reader.pages)
            # Signet (titre selon -par- ou non)
            if is_par:
                title = f"{idx:04d} – {ref} – {partner}"
            else:
                title = f"{idx:04d} – {ref}"
            _add_bookmark(writer, title, page_offset)
            # Pages
            for p in range(nb):
                writer.add_page(reader.pages[p])
            page_offset += nb
        except Exception as e:
            messagebox.showerror("Fusion PDF", f"Erreur lecture/ajout :\n{path}\n{e}")
            _finish_all()
            return

        # Avance
        pbar["value"] = i + 1
        var_prog.set(f"{i+1} / {total}")
        state["i"] = i + 1
        win.after(200, _step)

    # Démarrer
    win.after(300, _step)
