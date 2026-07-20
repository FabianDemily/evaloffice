"""Generation de rapports PDF individuels, charte graphique HERS."""

import zipfile
from pathlib import Path

from correcteurs.correcteur_excel import resumer_echecs

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

BLEU_HERS = colors.HexColor("#5BC4E8")
MAGENTA_HERS = colors.HexColor("#E0007A")
GRIS_HERS = colors.HexColor("#4A4A4A")

CHEMINS_CALIBRI = [
    "/Library/Fonts/Microsoft/Calibri.ttf",
    "/Library/Fonts/Calibri.ttf",
    str(Path.home() / "Library/Fonts/Calibri.ttf"),
    "C:/Windows/Fonts/calibri.ttf",
]


def _police_par_defaut():
    """Enregistre Calibri si disponible sur la machine, sinon retombe sur Helvetica."""
    for chemin in CHEMINS_CALIBRI:
        if Path(chemin).exists():
            pdfmetrics.registerFont(TTFont("Calibri", chemin))
            return "Calibri"
    return "Helvetica"


POLICE = _police_par_defaut()


def _styles():
    base = getSampleStyleSheet()
    return {
        "titre": ParagraphStyle("titre", parent=base["Title"], fontName=POLICE, textColor=MAGENTA_HERS),
        "sous_titre": ParagraphStyle("sous_titre", parent=base["Heading2"], fontName=POLICE, textColor=BLEU_HERS),
        "normal": ParagraphStyle("normal", parent=base["Normal"], fontName=POLICE, textColor=GRIS_HERS),
    }


ORANGE_MANUEL = colors.HexColor("#F59E0B")


def _table_criteres(exercice):
    styles = _styles()
    style_motif = ParagraphStyle(
        "motif", parent=styles["normal"],
        fontSize=8, textColor=colors.HexColor("#888888"), fontName=POLICE,
    )
    lignes = [["Compétence", "Statut", "Points"]]
    style_tableau = [
        ("BACKGROUND", (0, 0), (-1, 0), BLEU_HERS),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, -1), POLICE),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.lightgrey),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
    ]

    for i, critere in enumerate(exercice["criteres"], start=1):
        corrige = critere.get("corrige_manuellement", False)
        motif = critere.get("motif_manuel", "")
        label = critere.get("competence", critere["description"])

        if corrige:
            statut_txt = "(corr. manuelle)"
            contenu_comp = Paragraph(f"{label}<br/><i>{motif}</i>" if motif else label, styles["normal"])
            style_tableau.append(("BACKGROUND", (0, i), (-1, i), colors.HexColor("#FEF3C7")))
            style_tableau.append(("TEXTCOLOR", (1, i), (1, i), ORANGE_MANUEL))
        else:
            statut_txt = "Acquis" if critere["statut"] == "reussi" else "Non acquis"
            contenu_comp = Paragraph(label, styles["normal"])

        lignes.append([
            contenu_comp,
            statut_txt,
            f"{critere['points_obtenus']:.2f} / {critere['points_max']:.2f}",
        ])

    table = Table(lignes, colWidths=[9 * cm, 3 * cm, 3 * cm])
    table.setStyle(TableStyle(style_tableau))
    return table


def _table_comparaison(resultat, historique):
    styles = _styles()
    lignes = [["Session", "Score", "Bareme", "Pourcentage"]]
    for ancien in historique + [resultat]:
        bareme = ancien["bareme_total"] or 1
        pourcentage = round(100 * ancien["points_obtenus"] / bareme, 1)
        lignes.append([ancien["session"], f"{ancien['points_obtenus']:.2f}", str(bareme), f"{pourcentage}%"])

    table = Table(lignes, colWidths=[4 * cm, 3 * cm, 3 * cm, 3 * cm])
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), MAGENTA_HERS),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, -1), POLICE),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.lightgrey),
    ]))
    return table


def generer_rapport_pdf(resultat, chemin_sortie, historique=None, nom_cours="", code_cours=""):
    """Genere le rapport PDF individuel d'un etudiant.

    `historique` : liste optionnelle des resultats des sessions precedentes
    (meme etudiant, meme module), triee par ordre chronologique.
    `nom_cours` / `code_cours` : identifient le cours dans le titre du rapport
    (repli sur "EvalOffice" si non renseignes, cf onglet Configuration).
    """
    styles = _styles()
    identite = resultat["identite"]
    bareme = resultat["bareme_total"] or 1
    pourcentage = round(100 * resultat["points_obtenus"] / bareme, 1)

    if nom_cours or code_cours:
        titre_cours = " — ".join(t for t in (nom_cours, code_cours) if t)
    else:
        titre_cours = "EvalOffice"

    elements = [
        Paragraph(f"{titre_cours} — Rapport", styles["titre"]),
        Spacer(1, 0.3 * cm),
        Paragraph(f"{identite['nom']} {identite['prenom']}", styles["sous_titre"]),
        Paragraph(
            f"Session {resultat['session']} {resultat['annee']} — module {resultat['module']}",
            styles["normal"],
        ),
        Spacer(1, 0.4 * cm),
        Paragraph(
            f"Score global : {resultat['points_obtenus']:.2f} / {bareme} ({pourcentage}%)",
            styles["sous_titre"],
        ),
        Spacer(1, 0.4 * cm),
    ]

    for exercice in resultat["exercices"]:
        elements.append(Paragraph(
            f"{exercice['titre']} — {exercice['points_obtenus']:.2f} / {exercice['points_max']:.2f}",
            styles["sous_titre"],
        ))
        elements.append(_table_criteres(exercice))
        elements.append(Spacer(1, 0.3 * cm))

        criteres_manques = [c for c in exercice["criteres"] if c["statut"] != "reussi"]
        if criteres_manques:
            elements.append(Paragraph("Feedback :", styles["normal"]))
            for critere in criteres_manques:
                label = critere.get("competence", critere["description"])
                resume = resumer_echecs(critere.get("details", []))
                texte = f"- {label} : {resume}" if resume else f"- A revoir : {label}"
                elements.append(Paragraph(texte, styles["normal"]))
            elements.append(Spacer(1, 0.3 * cm))

    if historique:
        elements.append(Paragraph("Progression entre sessions", styles["sous_titre"]))
        elements.append(_table_comparaison(resultat, historique))

    doc = SimpleDocTemplate(str(chemin_sortie), pagesize=A4)
    doc.build(elements)


def _generer_histogramme_scores(pourcentages):
    """Retourne un BytesIO PNG de l'histogramme des scores (matplotlib)."""
    import io
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    import matplotlib.patches as mpatches

    fig, ax = plt.subplots(figsize=(7, 3))
    bins = list(range(0, 101, 10))
    n, _, patches = ax.hist(pourcentages, bins=bins, edgecolor="white", linewidth=0.8)
    for patch, left in zip(patches, bins[:-1]):
        if left >= 70:
            patch.set_facecolor("#22c55e")
        elif left >= 50:
            patch.set_facecolor("#f59e0b")
        else:
            patch.set_facecolor("#ef4444")
    ax.axvline(50, color="#555", linestyle="--", linewidth=1, label="Seuil 50 %")
    ax.set_xlabel("Score (%)")
    ax.set_ylabel("Nb étudiants")
    ax.set_xlim(0, 100)
    ax.set_xticks(bins)
    ax.yaxis.get_major_locator().set_params(integer=True)
    ax.legend(fontsize=8)
    ax.spines[["top", "right"]].set_visible(False)
    fig.tight_layout()
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=130)
    plt.close(fig)
    buf.seek(0)
    return buf


def _generer_graphique_competences(comp_stats):
    """Retourne un BytesIO PNG du taux de réussite par compétence (barres horizontales)."""
    import io
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt

    labels = list(comp_stats.keys())
    taux = [round(100 * v["reussi"] / v["total"], 1) for v in comp_stats.values()]
    # Tri du moins bon au meilleur
    paires = sorted(zip(taux, labels))
    taux_tries = [p[0] for p in paires]
    labels_tries = [p[1] for p in paires]
    couleurs = ["#22c55e" if t >= 75 else "#f59e0b" if t >= 50 else "#ef4444" for t in taux_tries]

    hauteur = max(2.5, len(labels) * 0.45)
    fig, ax = plt.subplots(figsize=(7, hauteur))
    bars = ax.barh(labels_tries, taux_tries, color=couleurs, edgecolor="white", linewidth=0.5)
    ax.axvline(50, color="#555", linestyle="--", linewidth=1)
    ax.set_xlim(0, 100)
    ax.set_xlabel("Taux de réussite (%)")
    for bar, t in zip(bars, taux_tries):
        ax.text(bar.get_width() + 1, bar.get_y() + bar.get_height() / 2,
                f"{t:.0f} %", va="center", fontsize=8)
    ax.spines[["top", "right"]].set_visible(False)
    fig.tight_layout()
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=130)
    plt.close(fig)
    buf.seek(0)
    return buf


def generer_rapport_classe_pdf(resultats, chemin_sortie, nom_cours="", code_cours=""):
    """Génère un rapport PDF de synthèse pour l'enseignant (toute la classe)."""
    from reportlab.platypus import Image as RLImage
    from datetime import date

    styles = _styles()
    identifiant_cours = " — ".join(t for t in (nom_cours, code_cours) if t) or "EvalOffice"
    session_info = f"{resultats[0]['session']} {resultats[0]['annee']} — module {resultats[0]['module']}" if resultats else ""

    # Métriques
    bareme = resultats[0]["bareme_total"] or 1
    pourcentages = [round(100 * r["points_obtenus"] / bareme, 1) for r in resultats]
    moyenne = round(sum(pourcentages) / len(pourcentages), 1)
    mediane = round(sorted(pourcentages)[len(pourcentages) // 2], 1)
    nb_reussi = sum(1 for p in pourcentages if p >= 50)

    # Statistiques par compétence
    comp_stats = {}
    for r in resultats:
        for ex in r["exercices"]:
            for c in ex["criteres"]:
                label = c.get("competence", c["description"])[:45]
                if label not in comp_stats:
                    comp_stats[label] = {"reussi": 0, "total": 0}
                comp_stats[label]["total"] += 1
                if c["statut"] == "reussi":
                    comp_stats[label]["reussi"] += 1

    elements = [
        Paragraph(f"{identifiant_cours} — Rapport de classe", styles["titre"]),
        Spacer(1, 0.2 * cm),
        Paragraph(session_info, styles["sous_titre"]),
        Paragraph(f"Généré le {date.today().strftime('%d/%m/%Y')} — {len(resultats)} copie(s) corrigée(s)", styles["normal"]),
        Spacer(1, 0.4 * cm),
    ]

    # Tableau métriques
    lignes_metrics = [
        ["Moyenne", "Médiane", "Taux de réussite", "Min / Max"],
        [f"{moyenne} %", f"{mediane} %", f"{nb_reussi}/{len(resultats)}", f"{min(pourcentages):.0f} % / {max(pourcentages):.0f} %"],
    ]
    tbl_metrics = Table(lignes_metrics, colWidths=[3.75 * cm] * 4)
    tbl_metrics.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), BLEU_HERS),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, -1), POLICE),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.lightgrey),
    ]))
    elements += [tbl_metrics, Spacer(1, 0.5 * cm)]

    # Histogramme des scores
    elements.append(Paragraph("Distribution des scores", styles["sous_titre"]))
    buf_hist = _generer_histogramme_scores(pourcentages)
    elements += [RLImage(buf_hist, width=14 * cm, height=6 * cm), Spacer(1, 0.4 * cm)]

    # Graphique compétences
    elements.append(Paragraph("Taux de réussite par compétence", styles["sous_titre"]))
    buf_comp = _generer_graphique_competences(comp_stats)
    hauteur_comp = min(12, max(4, len(comp_stats) * 0.9))
    elements += [RLImage(buf_comp, width=14 * cm, height=hauteur_comp * cm), Spacer(1, 0.4 * cm)]

    # Tableau détail compétences
    elements.append(Paragraph("Détail par compétence", styles["sous_titre"]))
    lignes_comp = [["Compétence", "Réussis", "Total", "Taux"]]
    for label, v in sorted(comp_stats.items(), key=lambda x: x[1]["reussi"] / x[1]["total"]):
        taux = round(100 * v["reussi"] / v["total"], 1)
        lignes_comp.append([Paragraph(label, styles["normal"]), str(v["reussi"]), str(v["total"]), f"{taux} %"])
    tbl_comp = Table(lignes_comp, colWidths=[9 * cm, 2 * cm, 2 * cm, 2 * cm])
    tbl_comp.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), BLEU_HERS),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, -1), POLICE),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.lightgrey),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
    ]))
    elements += [tbl_comp, Spacer(1, 0.5 * cm)]

    # Tableau des résultats individuels
    elements.append(Paragraph("Résultats individuels", styles["sous_titre"]))
    lignes_ind = [["Nom", "Prénom", "Points", "Barème", "Score"]]
    for r in sorted(resultats, key=lambda x: x["points_obtenus"], reverse=True):
        idf = r["identite"]
        pct = round(100 * r["points_obtenus"] / bareme, 1)
        lignes_ind.append([idf["nom"], idf["prenom"], f"{r['points_obtenus']:.2f}", str(bareme), f"{pct} %"])
    tbl_ind = Table(lignes_ind, colWidths=[4 * cm, 4 * cm, 2.5 * cm, 2.5 * cm, 2 * cm])
    tbl_ind.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), MAGENTA_HERS),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, -1), POLICE),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.lightgrey),
        # Colorier les lignes < 50 %
        *[("BACKGROUND", (0, i), (-1, i), colors.HexColor("#FEE2E2"))
          for i, r in enumerate(resultats, start=1)
          if round(100 * r["points_obtenus"] / bareme, 1) < 50],
    ]))
    elements.append(tbl_ind)

    doc = SimpleDocTemplate(str(chemin_sortie), pagesize=A4)
    doc.build(elements)


def generer_zip_rapports(resultats, chemin_zip, dossier_temp, historiques=None, nom_cours="", code_cours=""):
    """Genere un PDF par resultat puis les regroupe dans un ZIP.

    `historiques` : dict optionnel {(nom, prenom): [resultats precedents]}.
    """
    historiques = historiques or {}
    chemins_pdf = []
    for resultat in resultats:
        identite = resultat["identite"]
        cle = (identite["nom"], identite["prenom"])
        chemin_pdf = Path(dossier_temp) / f"{identite['nom']}_{identite['prenom']}_{resultat['session']}.pdf"
        generer_rapport_pdf(
            resultat, chemin_pdf, historique=historiques.get(cle),
            nom_cours=nom_cours, code_cours=code_cours,
        )
        chemins_pdf.append(chemin_pdf)

    with zipfile.ZipFile(chemin_zip, "w") as archive:
        for chemin_pdf in chemins_pdf:
            archive.write(chemin_pdf, arcname=chemin_pdf.name)
