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


def _table_criteres(exercice):
    styles = _styles()
    lignes = [["Compétence", "Statut", "Points"]]
    for critere in exercice["criteres"]:
        icone = "Acquis" if critere["statut"] == "reussi" else "Non acquis"
        lignes.append([
            Paragraph(critere.get("competence", critere["description"]), styles["normal"]),
            icone,
            f"{critere['points_obtenus']:.2f} / {critere['points_max']:.2f}",
        ])

    table = Table(lignes, colWidths=[9 * cm, 3 * cm, 3 * cm])
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), BLEU_HERS),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, -1), POLICE),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.lightgrey),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
    ]))
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
