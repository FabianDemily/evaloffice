"""Correspondances multilingues des noms de styles Word.

python-docx expose deux identifiants pour chaque style :
  - style.style_id  : identifiant XML interne (ex: "Heading1", "Titre1")
  - style.name      : nom affiché localisé   (ex: "Heading 1", "Titre 1")

Les deux varient selon la langue de Word. Ce module normalise les deux
vers une clé canonique anglaise minuscule (ex: "heading1", "toc1").
"""

import unicodedata

# Clé canonique → noms localisés connus
SYNONYMES_STYLES = {
    "heading1":      ["Heading 1", "Titre 1", "Überschrift 1", "Encabezado 1", "Kop 1"],
    "heading2":      ["Heading 2", "Titre 2", "Überschrift 2", "Encabezado 2", "Kop 2"],
    "heading3":      ["Heading 3", "Titre 3", "Überschrift 3", "Encabezado 3", "Kop 3"],
    "normal":        ["Normal", "Standard", "Normale", "Estándar"],
    "body_text":     ["Body Text", "Corps de texte", "Textkörper", "Cuerpo de texto"],
    "toc1":          ["TOC 1", "Table des matières 1", "Inhaltsverzeichnis 1", "TDM 1"],
    "toc2":          ["TOC 2", "Table des matières 2", "Inhaltsverzeichnis 2", "TDM 2"],
    "toc3":          ["TOC 3", "Table des matières 3"],
    "header":        ["Header", "En-tête", "Kopfzeile", "Encabezado"],
    "footer":        ["Footer", "Pied de page", "Fußzeile", "Pie de página"],
    "footnote_text": ["Footnote Text", "Note de bas de page", "Fußnotentext"],
    "caption":       ["Caption", "Légende", "Beschriftung"],
}

# styleId XML → clé canonique (couvre les variantes les plus fréquentes)
_ID_MAP = {
    "heading1": "heading1", "heading2": "heading2", "heading3": "heading3",
    "titre1": "heading1",   "titre2": "heading2",   "titre3": "heading3",
    "berschrift1": "heading1", "berschrift2": "heading2", "berschrift3": "heading3",
    "ttulo1": "heading1",   "ttulo2": "heading2",
    "normal": "normal",
    "bodytext": "body_text", "corpsdetexte": "body_text",
    "toc1": "toc1", "toc2": "toc2", "toc3": "toc3",
    "tabledesmatires1": "toc1", "tabledesmatires2": "toc2",
    "inhaltsverzeichnis1": "toc1",
    "header": "header", "entte": "header",
    "footer": "footer", "piedtdepage": "footer",
    "footnotereference": "footnote_text",
    "notedebassdepage": "footnote_text",
}


def _norm(s):
    """Supprime espaces, tirets, accents pour comparaison robuste."""
    s = s.replace(" ", "").replace("-", "").replace("_", "")
    s = unicodedata.normalize("NFD", s)
    return "".join(c for c in s if unicodedata.category(c) != "Mn").lower()


def canonique(style_id_ou_nom):
    """Retourne la clé canonique depuis un styleId ou un nom localisé."""
    n = _norm(style_id_ou_nom)
    if n in _ID_MAP:
        return _ID_MAP[n]
    for cle, noms in SYNONYMES_STYLES.items():
        if any(_norm(nom) == n for nom in noms):
            return cle
    return style_id_ou_nom.lower()


def est_style(para, cle_canonique):
    """Retourne True si le paragraphe a le style correspondant à la clé canonique."""
    style = para.style
    if canonique(style.style_id) == cle_canonique:
        return True
    if canonique(style.name) == cle_canonique:
        return True
    return False
