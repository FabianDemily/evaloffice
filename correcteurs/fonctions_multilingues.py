"""Table de correspondance des noms de fonctions Excel selon la langue."""

FONCTIONS_MULTILINGUES = {
    "SOMME": ["SOMME", "SUM", "SUMME", "SUMA", "SOMMA", "SOM"],
    "MOYENNE": ["MOYENNE", "AVERAGE", "MITTELWERT", "PROMEDIO", "MEDIA", "GEMIDDELDE"],
    "MAX": ["MAX"],
    "MIN": ["MIN"],
    "NB": ["NB", "COUNT", "ANZAHL", "CONTAR", "CONTA", "AANTAL"],
    "SI": ["SI", "IF", "WENN", "SE", "ALS"],
    "PRODUIT": ["PRODUIT", "PRODUCT", "PRODUKT", "PRODUCTO", "PRODOTTO"],
}


def synonymes(fonction_canonique: str) -> list[str]:
    """Retourne la liste des synonymes connus pour une fonction canonique (ex: SOMME)."""
    return FONCTIONS_MULTILINGUES.get(fonction_canonique.upper(), [fonction_canonique.upper()])
