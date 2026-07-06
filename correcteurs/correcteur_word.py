"""Correcteur Word générique, piloté par config JSON.

Chaque critère est identifié par son type et ses ancres textuelles.
La correction est tolérante aux variantes linguistiques (FR/EN) via styles_multilingues.
"""

import unicodedata
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn

from .styles_multilingues import canonique, est_style

STATUT_REUSSI = "reussi"
STATUT_ECHOUE = "echoue"

FEUILLE_GARDE = None  # Word n'a pas d'onglets


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _normaliser_texte(s):
    """Minuscules sans accents pour comparaisons tolérantes."""
    s = unicodedata.normalize("NFD", s)
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    return s.lower()


def _trouver_paragraphe(doc, ancre):
    """Retourne le paragraphe correspondant à l'ancre (insensible à la casse/accents).

    Privilégie la correspondance exacte pour éviter de confondre l'ancre réelle
    avec les paragraphes de consignes qui la citent entre guillemets.
    """
    ancre_n = _normaliser_texte(ancre)
    # 1. Correspondance exacte
    for p in doc.paragraphs:
        if _normaliser_texte(p.text) == ancre_n:
            return p
    # 2. Repli : contient l'ancre
    for p in doc.paragraphs:
        if ancre_n in _normaliser_texte(p.text):
            return p
    return None


def _texte_document(doc):
    """Retourne tout le texte du document en minuscules sans accents."""
    return _normaliser_texte(" ".join(p.text for p in doc.paragraphs))


def _texte_entete(doc):
    try:
        return _normaliser_texte(
            " ".join(p.text for p in doc.sections[0].header.paragraphs)
        )
    except Exception:
        return ""


def _texte_pied(doc):
    try:
        return _normaliser_texte(
            " ".join(p.text for p in doc.sections[0].footer.paragraphs)
        )
    except Exception:
        return ""


def _compter_notes_de_bas_de_page(doc):
    """Compte les notes de bas de page utilisateur (hors séparateurs système)."""
    try:
        footnotes_part = doc.part.footnotes
        if footnotes_part is None:
            return 0
        notes = footnotes_part._element.findall(qn("w:footnote"))
        # Les id -1 et 0 sont les séparateurs système
        user_notes = [n for n in notes if n.get(qn("w:id")) not in ("-1", "0")]
        return len(user_notes)
    except Exception:
        return 0


def _a_saut_de_page(doc, ancre):
    """Vérifie si un saut de page (manuel ou par propriété) précède le paragraphe ancré."""
    # Utilise _trouver_paragraphe pour éviter de matcher les paragraphes de consignes
    # qui citent l'ancre entre guillemets plutôt que l'ancre réelle dans le corps.
    para = _trouver_paragraphe(doc, ancre)
    if para is None:
        return False

    # Retrouver l'index en comparant les éléments XML (doc.paragraphs retourne
    # de nouveaux objets Python à chaque appel, donc .index() échoue).
    paragraphes = doc.paragraphs
    idx = next((i for i, p in enumerate(paragraphes) if p._p is para._p), None)
    if idx is None:
        return False

    # Vérifie un saut dans le paragraphe lui-même (w:br type="page")
    for br in para._p.iter(qn("w:br")):
        if br.get(qn("w:type")) == "page":
            return True

    # Vérifie la propriété pageBreakBefore
    pPr = para._p.find(qn("w:pPr"))
    if pPr is not None:
        pbr = pPr.find(qn("w:pageBreakBefore"))
        if pbr is not None and pbr.get(qn("w:val"), "1") != "0":
            return True

    # Vérifie si les 1 ou 2 paragraphes précédents contiennent un saut de page
    # (Word insère souvent le saut dans un paragraphe vide juste avant)
    for offset in (1, 2):
        if idx >= offset:
            prev = paragraphes[idx - offset]
            for br in prev._p.iter(qn("w:br")):
                if br.get(qn("w:type")) == "page":
                    return True

    return False


def _a_champ_toc(doc):
    """Vérifie si le document contient un champ TOC ou des paragraphes de style TDM."""
    # Recherche dans le XML du corps
    corps_xml = doc.element.body.xml
    if "TOC" in corps_xml or "\\o" in corps_xml:
        return True
    # Repli : cherche des paragraphes de style TDM
    for p in doc.paragraphs:
        if canonique(p.style.style_id) in ("toc1", "toc2", "toc3"):
            return True
        if canonique(p.style.name) in ("toc1", "toc2", "toc3"):
            return True
    return False


# ---------------------------------------------------------------------------
# Fonctions de correction par type
# ---------------------------------------------------------------------------

def _corriger_mise_en_page(doc, critere):
    points_max = critere.get("points", 1)
    orientation_attendue = critere.get("orientation", "paysage")
    marge_attendue_cm = critere.get("marge_cm", 2.0)
    tolerance_cm = 0.3  # tolérance de ±3 mm

    section = doc.sections[0]
    from docx.enum.section import WD_ORIENT
    from docx.shared import Cm

    details = []
    pts = 0
    pts_par = points_max / 2

    # Orientation
    est_paysage = section.page_width > section.page_height
    if orientation_attendue == "paysage" and est_paysage:
        pts += pts_par
        details.append(f"Orientation paysage : +{pts_par:.2f} pt")
    elif orientation_attendue == "portrait" and not est_paysage:
        pts += pts_par
        details.append(f"Orientation portrait : +{pts_par:.2f} pt")
    else:
        orientation_trouvee = "paysage" if est_paysage else "portrait"
        details.append(
            f"Orientation incorrecte : {orientation_trouvee} trouvée, "
            f"{orientation_attendue} attendue (0 pt)"
        )

    # Marges (on vérifie la marge gauche comme indicateur)
    from docx.shared import Cm as _Cm
    marge_gauche_cm = section.left_margin.cm if section.left_margin else 999
    if abs(marge_gauche_cm - marge_attendue_cm) <= tolerance_cm:
        pts += pts_par
        details.append(f"Marges {marge_gauche_cm:.1f} cm : +{pts_par:.2f} pt")
    else:
        details.append(
            f"Marges incorrectes : {marge_gauche_cm:.1f} cm trouvés, "
            f"{marge_attendue_cm} cm attendus (0 pt)"
        )

    return points_max, pts, details


def _corriger_structure_hierarchique(doc, critere):
    points_max = critere.get("points", 1)
    paragraphes = critere.get("paragraphes", [])
    if not paragraphes:
        return points_max, 0, ["Aucun paragraphe à vérifier"]

    pts_par = points_max / len(paragraphes)
    pts = 0.0
    details = []

    for spec in paragraphes:
        ancre = spec["ancre"]
        style_attendu = spec["style"]  # clé canonique ex: "heading1"
        para = _trouver_paragraphe(doc, ancre)
        if para is None:
            details.append(f"Paragraphe « {ancre[:30]} » : introuvable (0 pt)")
            continue
        if est_style(para, style_attendu):
            pts += pts_par
            details.append(f"« {ancre[:30]} » : style correct (+{pts_par:.2f} pt)")
        else:
            style_trouve = para.style.name
            details.append(
                f"« {ancre[:30]} » : style incorrect ({style_trouve} trouvé, "
                f"{style_attendu} attendu) (0 pt)"
            )

    return points_max, pts, details


def _corriger_styles_paragraphe(doc, critere):
    points_max = critere.get("points", 1)
    style_attendu = critere.get("style_attendu", "body_text")
    ancres = critere.get("ancres", [])
    if not ancres:
        return points_max, 0, ["Aucun paragraphe à vérifier"]

    pts_par = points_max / len(ancres)
    pts = 0.0
    details = []

    for ancre in ancres:
        para = _trouver_paragraphe(doc, ancre)
        if para is None:
            details.append(f"Paragraphe « {ancre[:30]}... » : introuvable (0 pt)")
            continue
        if est_style(para, style_attendu):
            pts += pts_par
            details.append(f"Style correct sur « {ancre[:30]}... » : +{pts_par:.2f} pt")
        else:
            details.append(
                f"Style incorrect sur « {ancre[:30]}... » : "
                f"{para.style.name} trouvé (0 pt)"
            )

    return points_max, pts, details


def _corriger_mise_en_forme_caracteres(doc, critere):
    points_max = critere.get("points", 1)
    ancre = critere.get("ancre_paragraphe", "")
    mot_cible = critere.get("mot_cible", "")
    gras_attendu = critere.get("gras", False)
    taille_attendue = critere.get("taille_pt")
    couleur_attendue = critere.get("couleur_hex", "").upper().replace("#", "")
    tolerance_pt = 1

    para = _trouver_paragraphe(doc, ancre)
    if para is None:
        return points_max, 0, [f"Paragraphe ancre introuvable (0 pt)"]

    # Cherche le run contenant le mot cible
    mot_n = _normaliser_texte(mot_cible)
    run_cible = None
    for run in para.runs:
        if mot_n in _normaliser_texte(run.text):
            run_cible = run
            break

    if run_cible is None:
        return points_max, 0, [f"Mot « {mot_cible} » introuvable dans le paragraphe (0 pt)"]

    nb_criteres = sum([1, gras_attendu, bool(taille_attendue), bool(couleur_attendue)])
    pts_par = points_max / nb_criteres
    pts = pts_par  # point pour avoir trouvé le mot
    details = [f"Mot « {mot_cible} » trouvé : +{pts_par:.2f} pt"]

    if gras_attendu:
        if run_cible.bold:
            pts += pts_par
            details.append(f"Gras présent : +{pts_par:.2f} pt")
        else:
            details.append("Gras manquant (0 pt)")

    if taille_attendue:
        taille_trouvee = run_cible.font.size.pt if run_cible.font.size else None
        if taille_trouvee and abs(taille_trouvee - taille_attendue) <= tolerance_pt:
            pts += pts_par
            details.append(f"Taille {taille_trouvee:.0f} pt : +{pts_par:.2f} pt")
        else:
            details.append(
                f"Taille incorrecte : {taille_trouvee} pt trouvé, "
                f"{taille_attendue} pt attendu (0 pt)"
            )

    if couleur_attendue:
        try:
            couleur_trouvee = str(run_cible.font.color.rgb).upper()
        except Exception:
            couleur_trouvee = ""
        if couleur_trouvee == couleur_attendue:
            pts += pts_par
            details.append(f"Couleur #{couleur_trouvee} : +{pts_par:.2f} pt")
        else:
            details.append(
                f"Couleur incorrecte : #{couleur_trouvee} trouvée, "
                f"#{couleur_attendue} attendue (0 pt)"
            )

    return points_max, pts, details


def _corriger_rechercher_remplacer(doc, critere):
    points_max = critere.get("points", 1)
    terme_ancien = _normaliser_texte(critere.get("terme_ancien", ""))
    terme_nouveau = _normaliser_texte(critere.get("terme_nouveau", ""))
    texte = _texte_document(doc)
    details = []
    pts = 0

    ancien_present = terme_ancien in texte
    nouveau_present = terme_nouveau in texte

    if not ancien_present and nouveau_present:
        pts = points_max
        details.append(
            f"Remplacement effectué : « {critere['terme_nouveau']} » présent, "
            f"« {critere['terme_ancien']} » absent (+{points_max:.2f} pt)"
        )
    elif ancien_present and nouveau_present:
        pts = points_max / 2
        details.append(
            f"Remplacement partiel : « {critere['terme_ancien']} » encore présent "
            f"dans le document (+{pts:.2f} pt)"
        )
    elif not nouveau_present:
        details.append(
            f"Remplacement non effectué : "
            f"« {critere['terme_nouveau']} » absent (0 pt)"
        )

    return points_max, pts, details


def _corriger_saisie_texte(doc, critere):
    points_max = critere.get("points", 1)
    mots_cles = [_normaliser_texte(m) for m in critere.get("mots_cles", [])]
    texte = _texte_document(doc)
    details = []

    if not mots_cles:
        return points_max, 0, ["Aucun mot-clé défini"]

    pts_par = points_max / len(mots_cles)
    pts = 0.0

    for mot in mots_cles:
        if mot in texte:
            pts += pts_par
            details.append(f"Mot-clé « {mot} » trouvé : +{pts_par:.2f} pt")
        else:
            details.append(f"Mot-clé « {mot} » absent (0 pt)")

    return points_max, pts, details


def _corriger_stabilite_mise_en_page(doc, critere):
    points_max = critere.get("points", 1)
    ancre = critere.get("ancre_saut", "")
    ok = _a_saut_de_page(doc, ancre)
    if ok:
        details = [f"Saut de page détecté avant « {ancre[:40]} » : +{points_max:.2f} pt"]
    else:
        details = [f"Saut de page absent avant « {ancre[:40]} » (0 pt)"]
    return points_max, (points_max if ok else 0), details


def _corriger_image(doc, critere):
    points_max = critere.get("points", 1)
    largeur_min = critere.get("largeur_min_cm", 0)
    largeur_max = critere.get("largeur_max_cm", 999)

    shapes = doc.inline_shapes
    if not shapes:
        return points_max, 0, ["Aucune image détectée (0 pt)"]

    from docx.shared import Cm
    details = [f"{len(shapes)} image(s) détectée(s)"]
    pts_par = points_max / 2
    pts = pts_par  # présence

    img = shapes[0]
    largeur_cm = img.width.cm if img.width else 0
    if largeur_min <= largeur_cm <= largeur_max:
        pts += pts_par
        details.append(f"Largeur {largeur_cm:.1f} cm (attendu {largeur_min}–{largeur_max} cm) : +{pts_par:.2f} pt")
    else:
        details.append(
            f"Largeur {largeur_cm:.1f} cm hors plage {largeur_min}–{largeur_max} cm (0 pt)"
        )

    return points_max, pts, details


def _corriger_table_des_matieres(doc, critere):
    points_max = critere.get("points", 1)
    ok = _a_champ_toc(doc)
    if ok:
        details = [f"Table des matières détectée : +{points_max:.2f} pt"]
    else:
        details = ["Table des matières absente (0 pt)"]
    return points_max, (points_max if ok else 0), details


def _corriger_entete_pied_page(doc, critere):
    points_max = critere.get("points", 1)
    mots_entete = [_normaliser_texte(m) for m in critere.get("mots_cles_entete", [])]
    mots_pied = [_normaliser_texte(m) for m in critere.get("mots_cles_pied", [])]

    pts_par = points_max / 2
    pts = 0.0
    details = []

    texte_entete = _texte_entete(doc)
    entete_ok = all(m in texte_entete for m in mots_entete) if mots_entete else bool(texte_entete.strip())
    if entete_ok:
        pts += pts_par
        details.append(f"En-tête renseigné : +{pts_par:.2f} pt")
    else:
        manquants = [m for m in mots_entete if m not in texte_entete]
        details.append(f"En-tête incomplet — éléments manquants : {', '.join(manquants)} (0 pt)")

    texte_pied = _texte_pied(doc)
    pied_ok = all(m in texte_pied for m in mots_pied) if mots_pied else bool(texte_pied.strip())
    if pied_ok:
        pts += pts_par
        details.append(f"Pied de page renseigné : +{pts_par:.2f} pt")
    else:
        manquants = [m for m in mots_pied if m not in texte_pied]
        details.append(f"Pied de page incomplet — éléments manquants : {', '.join(manquants)} (0 pt)")

    return points_max, pts, details


def _corriger_note_de_bas_de_page(doc, critere):
    points_max = critere.get("points", 1)
    nb_min = critere.get("nb_notes_min", 1)
    nb = _compter_notes_de_bas_de_page(doc)
    if nb >= nb_min:
        details = [f"{nb} note(s) de bas de page détectée(s) : +{points_max:.2f} pt"]
        return points_max, points_max, details
    else:
        details = [f"Aucune note de bas de page détectée (0 pt)"]
        return points_max, 0, details


# ---------------------------------------------------------------------------
# Dispatcher
# ---------------------------------------------------------------------------

_DISPATCH = {
    "mise_en_page":              lambda doc, critere: _corriger_mise_en_page(doc, critere),
    "structure_hierarchique":    lambda doc, critere: _corriger_structure_hierarchique(doc, critere),
    "styles_paragraphe":         lambda doc, critere: _corriger_styles_paragraphe(doc, critere),
    "mise_en_forme_caracteres":  lambda doc, critere: _corriger_mise_en_forme_caracteres(doc, critere),
    "rechercher_remplacer":      lambda doc, critere: _corriger_rechercher_remplacer(doc, critere),
    "saisie_texte":              lambda doc, critere: _corriger_saisie_texte(doc, critere),
    "stabilite_mise_en_page":    lambda doc, critere: _corriger_stabilite_mise_en_page(doc, critere),
    "image":                     lambda doc, critere: _corriger_image(doc, critere),
    "table_des_matieres":        lambda doc, critere: _corriger_table_des_matieres(doc, critere),
    "entete_pied_page":          lambda doc, critere: _corriger_entete_pied_page(doc, critere),
    "note_de_bas_de_page":       lambda doc, critere: _corriger_note_de_bas_de_page(doc, critere),
}


def corriger_critere_word(doc, critere):
    type_critere = critere["type"]
    if type_critere not in _DISPATCH:
        raise ValueError(f"Type de critère Word inconnu : {type_critere}")
    return _DISPATCH[type_critere](doc, critere)


def corriger_copie_word(chemin_fichier, config):
    """Corrige une copie Word selon la config et retourne le résultat structuré."""
    doc = Document(str(chemin_fichier))

    # Identité depuis le nom de fichier
    parts = Path(chemin_fichier).stem.split("_")
    if len(parts) >= 2:
        identite = {"nom": parts[0], "prenom": parts[1]}
    else:
        identite = {"nom": Path(chemin_fichier).stem, "prenom": ""}

    resultat = {
        "identite": identite,
        "session": config.get("session"),
        "annee": config.get("annee"),
        "module": config.get("module"),
        "variante": config.get("variante"),
        "watermark": None,
        "exercices": [],
        "points_obtenus": 0.0,
        "points_max": 0.0,
    }

    for exercice in config["exercices"]:
        ex_res = {
            "id": exercice["id"],
            "titre": exercice["titre"],
            "criteres": [],
            "points_obtenus": 0.0,
            "points_max": 0.0,
        }
        for critere in exercice["criteres"]:
            pts_max, pts_obt, details = corriger_critere_word(doc, critere)
            statut = STATUT_REUSSI if pts_obt >= pts_max - 1e-6 else STATUT_ECHOUE
            ex_res["criteres"].append({
                "id": critere["id"],
                "competence": critere.get("competence", critere.get("description", "")),
                "description": critere.get("description", ""),
                "points_max": pts_max,
                "points_obtenus": pts_obt,
                "statut": statut,
                "details": details,
            })
            ex_res["points_obtenus"] += pts_obt
            ex_res["points_max"] += pts_max

        resultat["exercices"].append(ex_res)
        resultat["points_obtenus"] += ex_res["points_obtenus"]
        resultat["points_max"] += ex_res["points_max"]

    resultat["bareme_total"] = config.get("bareme_total", resultat["points_max"])
    return resultat
