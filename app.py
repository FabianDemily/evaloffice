"""EvalOffice - application Streamlit pour la HERS.

Onglet 2 (Correction) est pleinement fonctionnel. Les onglets 1, 3 et 4
sont des emplacements a completer dans les etapes suivantes du projet.
"""

import io
import json
import tempfile
import zipfile
from pathlib import Path

import pandas as pd
import streamlit as st

from configuration import charger_configuration, sauver_configuration
from correcteurs.correcteur_excel import corriger_copie, resultats_vers_lignes_csv, resumer_echecs
from correcteurs.correcteur_word import corriger_copie_word
from generateur.gen_excel import charger_competences, charger_themes, generer_epreuve_excel, sauver_competences
from generateur.gen_word import charger_competences_word, generer_epreuve_word
from rapports.pdf_report import generer_zip_rapports

RACINE = Path(__file__).parent
DOSSIER_EPREUVES = RACINE / "data" / "epreuves"
DOSSIER_RESULTATS = RACINE / "data" / "resultats"


def _detecter_variante(nom_fichier):
    """Repere un segment lettre unique majuscule dans le nom de fichier (ex: '..._A_demily.xlsx' -> 'A')."""
    for segment in Path(nom_fichier).stem.split("_"):
        if len(segment) == 1 and segment.isalpha() and segment.isupper():
            return segment
    return None

st.set_page_config(page_title="EvalOffice", layout="wide")
st.title("EvalOffice — HERS")

configuration_cours = charger_configuration()

onglet_configuration, onglet_generateur, onglet_correction, onglet_rapports, onglet_historique = st.tabs(
    ["Configuration", "Generateur d'epreuves", "Correction", "Rapports PDF", "Historique"]
)

# --- Onglet 0 : Configuration ------------------------------------------------
with onglet_configuration:
    st.subheader("Configuration du cours")
    st.caption(
        "Le nom et le code du cours sont utilises comme prefixe des fichiers generes "
        "et comme titre des rapports PDF (a la place de 'EvalOffice')."
    )
    nom_cours_saisi = st.text_input("Nom du cours", value=configuration_cours.get("nom_cours", ""))
    code_cours_saisi = st.text_input("Code du cours", value=configuration_cours.get("code_cours", ""))
    if st.button("Enregistrer la configuration"):
        sauver_configuration({"nom_cours": nom_cours_saisi, "code_cours": code_cours_saisi})
        configuration_cours = {"nom_cours": nom_cours_saisi, "code_cours": code_cours_saisi}
        st.success("Configuration enregistree.")

# --- Onglet 1 : Generateur d'epreuves ---------------------------------------
TYPES_CRITERES_SUPPORTES = [
    "formule_fonction", "formule_multiplication", "formule_operateurs_mixtes", "tri", "filtre_auto",
    "mise_en_forme_conditionnelle", "graphique", "tcd", "texte_motcles",
    "recopie_formule", "mise_en_forme_cellule", "format_nombre",
]

with onglet_generateur:
    st.subheader("Generateur d'epreuves")

    themes = charger_themes()
    competences_excel = charger_competences()
    competences_word = charger_competences_word()

    if not themes:
        st.warning("Aucun theme defini dans data/themes.json")
    else:
        col_a, col_b, col_c = st.columns(3)
        with col_a:
            annee = st.number_input("Annee", min_value=2020, max_value=2100, value=2025, step=1)
        with col_b:
            session = st.text_input("Nom de l'epreuve", value="janvier", help="Ex: janvier, juin, aout, TP1, rattrapage...")
        with col_c:
            module = st.selectbox("Module", ["excel", "word"])

        competences = competences_excel if module == "excel" else competences_word

        contexte = st.selectbox("Contexte thematique", list(themes.keys()))
        nb_lignes = None
        if module == "excel":
            nb_lignes = st.slider(
                "Nombre de lignes de donnees par exercice", 5, 40, 20,
                help="Un nombre eleve decourage la saisie manuelle des formules ligne par ligne.",
            )

        st.markdown("**Competences a integrer**")
        col_tout, col_rien, col_vide = st.columns([1, 1, 4])
        with col_tout:
            if st.button("Tout selectionner", width="stretch"):
                for comp in competences:
                    st.session_state[f"actif_{comp['id']}"] = True
                st.rerun()
        with col_rien:
            if st.button("Tout deselectionner", width="stretch"):
                for comp in competences:
                    st.session_state[f"actif_{comp['id']}"] = False
                st.rerun()

        if module == "excel":
            groupes_labels = {
                "formules": "Formules de base", "tri": "Tri et filtres",
                "visuel": "Mise en forme et graphiques", "tcd": "Tableau croise dynamique",
                "synthese": "Question ouverte", "recopie": "Recopie de formule",
                "format_cellule": "Mise en forme d'une feuille de calcul",
                "format_nombre": "Formats de nombres",
            }
            groupes_couleurs = {
                "formules":      "#5BC4E8",
                "tri":           "#E0007A",
                "visuel":        "#7B5EA7",
                "tcd":           "#E8873A",
                "synthese":      "#4A9A6F",
                "recopie":       "#C0392B",
                "format_cellule":"#2980B9",
                "format_nombre": "#16A085",
            }
        else:
            groupes_labels = {
                "mise_en_page":  "Mise en page",
                "structure":     "Structure et styles",
                "mise_en_forme": "Mise en forme des caracteres",
                "outils":        "Outils Word",
                "objets":        "Objets inseres",
            }
            groupes_couleurs = {
                "mise_en_page":  "#5BC4E8",
                "structure":     "#7B5EA7",
                "mise_en_forme": "#E0007A",
                "outils":        "#E8873A",
                "objets":        "#4A9A6F",
            }
        actives = []
        points_par_competence = {}
        options_graphique = {}
        for groupe_id, groupe_label in groupes_labels.items():
            comp_du_groupe = [c for c in competences if c.get("groupe") == groupe_id]
            if not comp_du_groupe:
                continue
            couleur = groupes_couleurs.get(groupe_id, "#4A4A4A")
            st.markdown(
                f'<div style="background:{couleur};color:white;padding:4px 10px;'
                f'border-radius:4px;margin:10px 0 4px 0;font-weight:600;">'
                f'{groupe_label}</div>',
                unsafe_allow_html=True,
            )
            for comp in comp_du_groupe:
                col_check, col_points = st.columns([3, 1])
                with col_check:
                    coche = st.checkbox(
                        f"{comp['label']} — {comp['description']}",
                        value=comp.get("actif_defaut", True),
                        key=f"actif_{comp['id']}",
                    )
                with col_points:
                    points = st.number_input(
                        "Points", value=float(comp.get("points_defaut", 1)), min_value=0.0, step=0.5,
                        key=f"points_{comp['id']}", label_visibility="collapsed",
                    )
                if coche:
                    actives.append(comp["id"])
                points_par_competence[comp["id"]] = points

                if comp["id"] == "graphique" and coche:
                    with st.container():
                        gc1, gc2, gc3, gc4 = st.columns([1.5, 1, 1, 1])
                        with gc1:
                            type_graphique = st.selectbox(
                                "Type impose", ["libre", "histogramme", "courbes", "secteurs"],
                                key="graphique_type",
                                help="'libre' = l'etudiant choisit le type. 'histogramme' = colonnes verticales."
                            )
                        with gc2:
                            titre_oblige = st.checkbox("Titre oblige", key="graphique_titre", value=False)
                        with gc3:
                            tendance_disabled = (type_graphique == "secteurs")
                            tendance = st.checkbox(
                                "Courbe de tendance", key="graphique_tendance", value=False,
                                disabled=tendance_disabled,
                                help="Non disponible pour les graphiques en secteurs"
                            )
                        with gc4:
                            st.caption(
                                f"Points divises entre {1 + (1 if type_graphique != 'libre' else 0) + (1 if titre_oblige else 0) + (1 if tendance and not tendance_disabled else 0)} sous-critere(s)"
                            )
                    options_graphique = {
                        "type_impose": type_graphique if type_graphique != "libre" else None,
                        "titre_oblige": titre_oblige,
                        "tendance": tendance and not tendance_disabled,
                    }


        st.markdown("**Variantes**")
        nb_variantes = st.number_input("Nombre de variantes a generer simultanement", min_value=1, max_value=6, value=1, step=1)
        lettres_variantes = [chr(ord("A") + i) for i in range(nb_variantes)]

        if module == "excel":
            st.markdown("**Tracabilite**")
            activer_watermark = st.checkbox(
                "Activer le watermark de tracabilite (identifiant cache + hash des donnees source)",
                value=False,
                help=(
                    "Permet de detecter une substitution complete du fichier ou une falsification "
                    "des donnees source (prix, quantites...). Ne detecte PAS l'usage d'une IA — "
                    "c'est un indicateur de vigilance a verifier manuellement, jamais une preuve."
                ),
            )
        else:
            activer_watermark = False

        if st.button("Generer l'epreuve", disabled=not actives):
            DOSSIER_EPREUVES.mkdir(parents=True, exist_ok=True)
            noms_enregistres = []
            code_cours = configuration_cours.get("code_cours", "").strip()
            base_nom = f"{code_cours}_{annee}_{session}_{module}" if code_cours else f"{annee}_{session}_{module}"
            with tempfile.TemporaryDirectory() as tmp:
                chemin_zip = Path(tmp) / f"{base_nom}.zip"
                with zipfile.ZipFile(chemin_zip, "w") as archive:
                    for variante in lettres_variantes:
                        prefixe = f"{base_nom}_{variante}"
                        chemin_json = Path(tmp) / f"{prefixe}.json"

                        if module == "excel":
                            wb_etudiant, wb_corrige, config = generer_epreuve_excel(
                                contexte=contexte, session=session, annee=int(annee), variante=variante,
                                competences_actives=actives, points_par_competence=points_par_competence,
                                nb_lignes=nb_lignes, seed=f"{annee}-{session}-{variante}", competences=competences,
                                activer_watermark=activer_watermark,
                                options_par_competence={"graphique": options_graphique} if options_graphique else None,
                            )
                            chemin_etudiant = Path(tmp) / f"{prefixe}_etudiant.xlsx"
                            chemin_corrige  = Path(tmp) / f"{prefixe}_corrige.xlsx"
                            contenu_json = json.dumps(config, indent=2, ensure_ascii=False)
                            wb_etudiant.save(chemin_etudiant)
                            wb_corrige.save(chemin_corrige)
                            chemin_json.write_text(contenu_json, encoding="utf-8")
                            archive.write(chemin_etudiant, arcname=chemin_etudiant.name)
                            archive.write(chemin_corrige,  arcname=chemin_corrige.name)

                        else:  # word
                            doc_etudiant, doc_corrige, config = generer_epreuve_word(
                                contexte=contexte, session=session, annee=int(annee), variante=variante,
                                competences_actives=actives, points_par_competence=points_par_competence,
                                seed=f"{annee}-{session}-{variante}", competences=competences,
                            )
                            chemin_etudiant = Path(tmp) / f"{prefixe}_etudiant.docx"
                            chemin_corrige  = Path(tmp) / f"{prefixe}_corrige.docx"
                            contenu_json = json.dumps(config, indent=2, ensure_ascii=False)
                            doc_etudiant.save(chemin_etudiant)
                            doc_corrige.save(chemin_corrige)
                            chemin_json.write_text(contenu_json, encoding="utf-8")
                            archive.write(chemin_etudiant, arcname=chemin_etudiant.name)
                            archive.write(chemin_corrige,  arcname=chemin_corrige.name)

                        archive.write(chemin_json, arcname=chemin_json.name)

                        chemin_app = DOSSIER_EPREUVES / f"{prefixe}.json"
                        if chemin_app.exists():
                            st.warning(f"{chemin_app.name} existait deja et a ete remplace par cette nouvelle generation.")
                        chemin_app.write_text(contenu_json, encoding="utf-8")
                        noms_enregistres.append(chemin_app.name)

                ext = "xlsx" if module == "excel" else "docx"
                st.success(f"{nb_variantes} variante(s) generee(s) ({ext.upper()}).")
                st.info(
                    "Config(s) enregistree(s) dans data/epreuves, disponible(s) dans l'onglet "
                    f"Correction : {', '.join(noms_enregistres)}"
                )
                st.download_button(
                    "Telecharger le ZIP (etudiant + corrige + config JSON)",
                    data=chemin_zip.read_bytes(),
                    file_name=chemin_zip.name,
                    mime="application/zip",
                )

        with st.expander("Gerer le catalogue de competences"):
            st.caption(
                "Ajouter une competence reutilise un type de correction existant "
                "(formule, tri, filtre, mise en forme, graphique, TCD, mots-cles) avec tes propres parametres. "
                "Une logique de correction entierement nouvelle necessite du developpement."
            )

            st.markdown("**Retirer une competence**")
            id_a_retirer = st.selectbox(
                "Competence a retirer", options=[c["id"] for c in competences], key="id_a_retirer"
            )
            if st.button("Retirer du catalogue"):
                competences_maj = [c for c in competences if c["id"] != id_a_retirer]
                sauver_competences(competences_maj)
                st.success(f"Competence '{id_a_retirer}' retiree. Rechargez la page pour voir le changement.")

            st.markdown("**Ajouter une competence**")
            with st.form("ajout_competence"):
                nouvel_id = st.text_input("Identifiant unique (ex: formule_produit)")
                nouveau_label = st.text_input("Libelle")
                nouvelle_description = st.text_input("Description")
                nouveau_groupe = st.selectbox("Groupe / exercice", list(groupes_labels.keys()), format_func=lambda g: groupes_labels[g])
                nouveau_type = st.selectbox("Type de critere (logique de correction)", TYPES_CRITERES_SUPPORTES)
                nouveau_points = st.number_input("Points par defaut", min_value=0.0, value=1.0, step=0.5)
                nouvelle_fonction = ""
                nouveaux_mots_cles = ""
                if nouveau_type == "formule_fonction":
                    nouvelle_fonction = st.text_input("Fonction Excel attendue (ex: SOMME, MAX, MIN)")
                if nouveau_type == "texte_motcles":
                    nouveaux_mots_cles = st.text_input("Mots-cles attendus (separes par des virgules)")

                if st.form_submit_button("Ajouter au catalogue"):
                    if not nouvel_id or any(c["id"] == nouvel_id for c in competences):
                        st.error("Identifiant manquant ou deja utilise.")
                    else:
                        nouvelle_competence = {
                            "id": nouvel_id, "label": nouveau_label or nouvel_id,
                            "description": nouvelle_description, "groupe": nouveau_groupe,
                            "type": nouveau_type, "actif_defaut": False, "points_defaut": nouveau_points,
                        }
                        if nouveau_type == "formule_fonction":
                            nouvelle_competence["fonction"] = nouvelle_fonction.strip().upper()
                        if nouveau_type == "texte_motcles":
                            nouvelle_competence["mots_cles"] = [m.strip() for m in nouveaux_mots_cles.split(",") if m.strip()]

                        competences_maj = competences + [nouvelle_competence]
                        sauver_competences(competences_maj)
                        st.success(f"Competence '{nouvel_id}' ajoutee. Rechargez la page pour la voir apparaitre.")

# --- Onglet 2 : Correction automatique --------------------------------------
with onglet_correction:
    st.subheader("Correction automatique")

    configs_disponibles = sorted(DOSSIER_EPREUVES.glob("*.json"))
    if not configs_disponibles:
        st.warning(f"Aucun fichier de config trouve dans {DOSSIER_EPREUVES}")
    else:
        col_select, col_supprimer = st.columns([4, 1])
        with col_select:
            nom_config = st.selectbox(
                "Fichier de configuration de l'epreuve",
                options=[f.name for f in configs_disponibles],
            )
        chemin_config = DOSSIER_EPREUVES / nom_config
        config = json.loads(chemin_config.read_text(encoding="utf-8"))
        st.caption(
            f"Session {config.get('session')} {config.get('annee')} — "
            f"module {config.get('module')} — variante {config.get('variante')} — "
            f"contexte {config.get('contexte')}"
        )
        with col_supprimer:
            st.write("")
            if st.button("Supprimer cette config", key="supprimer_config"):
                chemin_config.unlink()
                st.success(f"{nom_config} supprime.")
                st.rerun()

        module_config = config.get("module", "excel")
        type_fichier = "docx" if module_config == "word" else "xlsx"
        fichiers = st.file_uploader(
            f"Deposer les copies etudiantes (.{type_fichier})",
            type=[type_fichier],
            accept_multiple_files=True,
        )

        if st.button("Lancer la correction", disabled=not fichiers):
            resultats = []
            erreurs = []
            avertissements = []
            variante_config = config.get("variante")
            with tempfile.TemporaryDirectory() as tmp:
                for fichier in fichiers:
                    variante_detectee = _detecter_variante(fichier.name)
                    if variante_detectee and variante_config and variante_detectee != variante_config:
                        avertissements.append(
                            f"{fichier.name} : le nom du fichier suggere la variante '{variante_detectee}' "
                            f"mais la config selectionnee est la variante '{variante_config}'. "
                            "Verifiez que c'est bien le bon fichier de config."
                        )

                    chemin_tmp = Path(tmp) / fichier.name
                    chemin_tmp.write_bytes(fichier.getvalue())
                    try:
                        if module_config == "word":
                            resultat = corriger_copie_word(chemin_tmp, config)
                        else:
                            resultat = corriger_copie(chemin_tmp, config)
                        resultats.append(resultat)
                    except Exception as exc:
                        erreurs.append(f"{fichier.name} : {exc}")

            st.session_state["resultats_correction"] = resultats
            if avertissements:
                st.warning("\n\n".join(avertissements))
            if erreurs:
                st.error("Erreurs lors de la correction :\n" + "\n".join(erreurs))

        resultats = st.session_state.get("resultats_correction", [])
        if resultats:
            lignes_synthese = []
            for resultat in resultats:
                identite = resultat["identite"]
                bareme = resultat["bareme_total"] or 1
                pourcentage = round(100 * resultat["points_obtenus"] / bareme, 1)
                watermark = resultat.get("watermark")
                vigilance = "—" if watermark is None else ("✅" if watermark["ok"] else "⚠️ A verifier")
                lignes_synthese.append({
                    "Nom": identite["nom"],
                    "Prenom": identite["prenom"],
                    "Points": round(resultat["points_obtenus"], 2),
                    "Bareme": bareme,
                    "Pourcentage": pourcentage,
                    "Statut": "Reussi" if pourcentage >= 50 else "Echoue",
                    "Tracabilite": vigilance,
                })
            df_synthese = pd.DataFrame(lignes_synthese)
            if any(r.get("watermark") for r in resultats):
                st.caption(
                    "Colonne 'Tracabilite' : indicateur de vigilance (substitution de fichier ou "
                    "donnees source modifiees), pas une preuve. A verifier manuellement avant toute decision."
                )

            col1, col2 = st.columns(2)
            with col1:
                seuil_score = st.slider("Score minimum (%)", 0, 100, 0)
            with col2:
                statuts = st.multiselect("Statut", ["Reussi", "Echoue"], default=["Reussi", "Echoue"])

            df_filtre = df_synthese[
                (df_synthese["Pourcentage"] >= seuil_score) & (df_synthese["Statut"].isin(statuts))
            ]
            st.dataframe(df_filtre)

            lignes_detail = []
            for resultat in resultats:
                lignes_detail.extend(resultats_vers_lignes_csv(resultat))
            df_detail = pd.DataFrame(lignes_detail)

            csv_buffer = io.StringIO()
            df_detail.to_csv(csv_buffer, index=False)
            st.download_button(
                "Exporter le detail en CSV",
                data=csv_buffer.getvalue(),
                file_name=f"{chemin_config.stem}_resultats.csv",
                mime="text/csv",
            )

            with st.expander("Detail par etudiant et par critere"):
                for resultat in resultats:
                    identite = resultat["identite"]
                    bareme = resultat["bareme_total"] or 1
                    pourcentage = round(100 * resultat["points_obtenus"] / bareme, 1)
                    st.markdown(
                        f"**{identite['nom']} {identite['prenom']}** — "
                        f"Score : {resultat['points_obtenus']:.2f} / {bareme} ({pourcentage}%)"
                    )
                    watermark = resultat.get("watermark")
                    if watermark and not watermark["ok"]:
                        for message in watermark["messages"]:
                            st.warning(message)
                    for exercice in resultat["exercices"]:
                        st.markdown(
                            f"*{exercice['titre']}* — "
                            f"{exercice['points_obtenus']:.2f} / {exercice['points_max']:.2f}"
                        )
                        # Tableau compétences
                        lignes_tableau = []
                        for critere in exercice["criteres"]:
                            icone = "✅" if critere["statut"] == "reussi" else "❌"
                            lignes_tableau.append({
                                "": icone,
                                "Compétence": critere.get("competence", critere["description"]),
                                "Points": f"{critere['points_obtenus']:.2f} / {critere['points_max']:.2f}",
                            })
                        st.dataframe(
                            pd.DataFrame(lignes_tableau),
                            hide_index=True,
                            column_config={"": st.column_config.TextColumn(width="small")},
                        )
                        # Feedback détaillé pour les critères non acquis
                        criteres_manques = [c for c in exercice["criteres"] if c["statut"] != "reussi"]
                        if criteres_manques:
                            lignes_feedback = []
                            for critere in criteres_manques:
                                label = critere.get("competence", critere["description"])
                                resume = resumer_echecs(critere.get("details", []))
                                lignes_feedback.append(
                                    f"- **{label}** : {resume}" if resume else f"- À revoir : **{label}**"
                                )
                            st.markdown("**Feedback :**\n" + "\n".join(lignes_feedback))
                    st.divider()

# --- Onglet 3 : Rapports PDF ------------------------------------------------
with onglet_rapports:
    st.subheader("Rapports PDF")
    resultats_rapports = st.session_state.get("resultats_correction", [])

    if not resultats_rapports:
        st.info("Lancez d'abord une correction dans l'onglet 'Correction'.")
    else:
        st.write(f"{len(resultats_rapports)} copie(s) corrigee(s) disponible(s).")
        if st.button("Generer les rapports PDF (ZIP)"):
            with tempfile.TemporaryDirectory() as tmp:
                chemin_zip = Path(tmp) / "rapports.zip"
                generer_zip_rapports(
                    resultats_rapports, chemin_zip, tmp,
                    nom_cours=configuration_cours.get("nom_cours", ""),
                    code_cours=configuration_cours.get("code_cours", ""),
                )
                st.download_button(
                    "Telecharger le ZIP des rapports",
                    data=chemin_zip.read_bytes(),
                    file_name="rapports_pdf.zip",
                    mime="application/zip",
                )

# --- Onglet 4 : Historique et statistiques (a developper) ------------------
with onglet_historique:
    st.info("Onglet en cours de developpement (etape 6 du brief).")
