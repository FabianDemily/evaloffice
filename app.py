"""EvalOffice - application Streamlit pour la HERS.

Onglet 2 (Correction) est pleinement fonctionnel. Les onglets 1, 3 et 4
sont des emplacements a completer dans les etapes suivantes du projet.
"""

import io
import json
import tempfile
import uuid
import zipfile
from pathlib import Path

import pandas as pd
import streamlit as st

from configuration import charger_configuration, sauver_configuration
from correcteurs.correcteur_excel import corriger_copie, resultats_vers_lignes_csv, resumer_echecs
from correcteurs.correcteur_word import corriger_copie_word
from generateur.gen_excel import charger_competences, charger_themes, generer_epreuve_excel, sauver_competences
from generateur.gen_word import charger_competences_word, generer_epreuve_word
from rapports.pdf_report import generer_zip_rapports, generer_rapport_classe_pdf

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
    ["Configuration", "Generateur d'epreuves", "Correction", "Rapports", "Historique"]
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
                "recopie": "Recopie de formule",
                "recherche": "Fonctions de recherche",
                "format_cellule": "Mise en forme d'une feuille de calcul",
                "format_nombre": "Formats de nombres",
            }
            groupes_couleurs = {
                "formules":      "#5BC4E8",
                "tri":           "#E0007A",
                "visuel":        "#7B5EA7",
                "tcd":           "#E8873A",
                "recopie":       "#C0392B",
                "recherche":     "#8E44AD",
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
        options_styles_paragraphe = {}
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

                if comp["id"] == "styles_paragraphe" and coche and module == "word":
                    with st.container():
                        sp1, sp2, sp3 = st.columns([1.5, 1, 1])
                        with sp1:
                            taille_imposee = st.checkbox(
                                "Taille imposee (varie par variante)",
                                key="styles_taille", value=True,
                                help="Cycle automatique 12/13/14 pt selon la variante — empêche la copie entre etudiants"
                            )
                        with sp2:
                            justifie = st.checkbox("Texte justifie", key="styles_justifie", value=True)
                        with sp3:
                            nb_sous = 1 + (1 if taille_imposee else 0) + (1 if justifie else 0)
                            st.caption(f"{nb_sous} sous-critere(s) — {points/nb_sous:.2f} pt chacun")
                    options_styles_paragraphe = {
                        "taille_imposee": taille_imposee,
                        "justifie": justifie,
                    }


        # Avertissement combinaison risquée styles_paragraphe + mise_en_forme_caracteres
        if module == "word":
            comps_cochees = {comp["id"] for comp, coche in zip(competences, [
                st.session_state.get(f"comp_{comp['id']}", False) for comp in competences
            ]) if coche}
            # Reconstruction simple : relire les checkboxes depuis session_state
            sp_active = st.session_state.get("comp_styles_paragraphe", False)
            mfc_active = st.session_state.get("comp_mise_en_forme_caracteres", False)
            tdm_active = st.session_state.get("comp_table_des_matieres", False)
            sh_active = st.session_state.get("comp_structure_hierarchique", False)
            if sp_active and mfc_active:
                st.warning(
                    "⚠️ Combinaison « Styles de paragraphe » + « Mise en forme des caractères » : "
                    "la consigne précisera d'appliquer le style Corps de texte EN PREMIER, "
                    "puis la mise en forme du mot. L'ordre a été ajusté automatiquement."
                )
            if tdm_active and not sh_active:
                st.warning(
                    "⚠️ « Table des matières » sans « Structure hiérarchique » : "
                    "la table des matières nécessite que les styles de titres soient appliqués. "
                    "Pensez à ajouter la compétence Structure hiérarchique."
                )

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
                                nb_lignes=nb_lignes, seed=f"{annee}-{session}-{variante}-{uuid.uuid4().hex[:8]}", competences=competences,
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
                                options_par_competence={"styles_paragraphe": options_styles_paragraphe} if options_styles_paragraphe else None,
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
            # Sauvegarde automatique dans data/resultats/
            DOSSIER_RESULTATS.mkdir(parents=True, exist_ok=True)
            for r in resultats:
                idf = r["identite"]
                nom_fichier = f"{idf['nom']}_{idf['prenom']}_{r['session']}_{r['annee']}_{r['module']}.json"
                chemin_json = DOSSIER_RESULTATS / nom_fichier
                r_avec_config = dict(r)
                r_avec_config["config_epreuve"] = config
                chemin_json.write_text(json.dumps(r_avec_config, indent=2, ensure_ascii=False), encoding="utf-8")

            if avertissements:
                st.warning("\n\n".join(avertissements))
            if erreurs:
                st.error("Erreurs lors de la correction :\n" + "\n".join(erreurs))

        resultats = st.session_state.get("resultats_correction", [])
        if resultats:
            # Détection croisée : last_modified_by présent sur plusieurs copies
            compteur_lmb = {}
            for r in resultats:
                lmb = r.get("metadonnees", {}).get("last_modified_by", "—")
                if lmb and lmb != "—":
                    compteur_lmb.setdefault(lmb, []).append(
                        f"{r['identite']['nom']} {r['identite']['prenom']}"
                    )
            lmb_suspects = {lmb: noms for lmb, noms in compteur_lmb.items() if len(noms) > 1}

            lignes_synthese = []
            for resultat in resultats:
                identite = resultat["identite"]
                bareme = resultat["bareme_total"] or 1
                pourcentage = round(100 * resultat["points_obtenus"] / bareme, 1)
                watermark = resultat.get("watermark")
                vigilance_wm = "—" if watermark is None else ("✅" if watermark["ok"] else "⚠️")
                meta = resultat.get("metadonnees", {})
                alertes_meta = list(meta.get("alertes", []))
                lmb = meta.get("last_modified_by", "—")
                if lmb in lmb_suspects:
                    alertes_meta.append(f"Même compte '{lmb}' sur {len(lmb_suspects[lmb])} copies")
                vigilance_meta = "⚠️" if alertes_meta else "✅"
                lignes_synthese.append({
                    "Nom": identite["nom"],
                    "Prenom": identite["prenom"],
                    "Points": round(resultat["points_obtenus"], 2),
                    "Bareme": bareme,
                    "Pourcentage": pourcentage,
                    "Statut": "Reussi" if pourcentage >= 50 else "Echoue",
                    "Watermark": vigilance_wm,
                    "Métadonnées": vigilance_meta,
                })
            df_synthese = pd.DataFrame(lignes_synthese)
            if any(r.get("watermark") for r in resultats):
                st.caption(
                    "Colonne 'Watermark' : substitution de fichier ou données modifiées. "
                    "Colonne 'Métadonnées' : auteur/modificateur/durée suspects. "
                    "⚠️ = à vérifier manuellement, jamais une preuve."
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

            if len(resultats) >= 2:
                st.info("Consultez l'onglet **Rapports** pour les graphiques et le rapport de classe.")

            st.markdown("---")
            st.markdown("**Détail par étudiant**")
            fc1, fc2, fc3 = st.columns([2, 1, 1])
            with fc1:
                recherche = st.text_input("🔍 Rechercher (nom ou prénom)", placeholder="ex: Dupont", label_visibility="collapsed")
            with fc2:
                tri_choix = st.selectbox("Trier par", ["Nom (A→Z)", "Score (↑)", "Score (↓)"], label_visibility="collapsed")
            with fc3:
                filtre_statut = st.multiselect("Statut", ["✅ Réussi", "❌ Échoué"], default=["✅ Réussi", "❌ Échoué"], label_visibility="collapsed")

            # Indexer les résultats avec leur position d'origine (pour les clés Streamlit)
            resultats_indexes = list(enumerate(resultats))

            # Filtrage
            if recherche.strip():
                terme = recherche.strip().lower()
                resultats_indexes = [
                    (i, r) for i, r in resultats_indexes
                    if terme in r["identite"]["nom"].lower() or terme in r["identite"]["prenom"].lower()
                ]
            if filtre_statut and len(filtre_statut) < 2:
                vouloir_reussi = "✅ Réussi" in filtre_statut
                resultats_indexes = [
                    (i, r) for i, r in resultats_indexes
                    if (r["points_obtenus"] / (r["bareme_total"] or 1) >= 0.5) == vouloir_reussi
                ]

            # Tri
            if tri_choix == "Nom (A→Z)":
                resultats_indexes.sort(key=lambda x: (x[1]["identite"]["nom"], x[1]["identite"]["prenom"]))
            elif tri_choix == "Score (↑)":
                resultats_indexes.sort(key=lambda x: x[1]["points_obtenus"])
            elif tri_choix == "Score (↓)":
                resultats_indexes.sort(key=lambda x: x[1]["points_obtenus"], reverse=True)

            st.caption(f"{len(resultats_indexes)} copie(s) affichée(s) sur {len(resultats)}")

            for idx_r, resultat in resultats_indexes:
                identite = resultat["identite"]
                bareme = resultat["bareme_total"] or 1
                pourcentage = round(100 * resultat["points_obtenus"] / bareme, 1)
                icone_statut = "✅" if pourcentage >= 50 else "❌"
                label_expander = (
                    f"{icone_statut} {identite['nom']} {identite['prenom']} — "
                    f"{resultat['points_obtenus']:.2f} / {bareme} ({pourcentage}%)"
                )
                with st.expander(label_expander):
                    watermark = resultat.get("watermark")
                    if watermark and not watermark["ok"]:
                        for message in watermark["messages"]:
                            st.warning(message)
                    meta = resultat.get("metadonnees")
                    if meta:
                        alertes_meta = list(meta.get("alertes", []))
                        lmb = meta.get("last_modified_by", "—")
                        if lmb in lmb_suspects:
                            alertes_meta.append(
                                f"Compte '{lmb}' identique sur : {', '.join(lmb_suspects[lmb])}"
                            )
                        with st.expander("🔍 Métadonnées du fichier"):
                            c1, c2 = st.columns(2)
                            c1.markdown(f"**Auteur original** : {meta['creator']}")
                            c1.markdown(f"**Dernière modif. par** : {lmb}")
                            c2.markdown(f"**Créé le** : {meta['created']}")
                            c2.markdown(f"**Modifié le** : {meta['modified']}")
                            duree = meta.get("duree_minutes")
                            if duree is not None:
                                c1.markdown(f"**Durée de travail** : {duree} min")
                            for alerte in alertes_meta:
                                st.warning(f"⚠️ {alerte}")
                    for idx_e, exercice in enumerate(resultat["exercices"]):
                        st.markdown(
                            f"*{exercice['titre']}* — "
                            f"{exercice['points_obtenus']:.2f} / {exercice['points_max']:.2f}"
                        )
                        # Tableau compétences
                        lignes_tableau = []
                        for critere in exercice["criteres"]:
                            if critere.get("corrige_manuellement"):
                                icone = "✏️"
                            elif critere["statut"] == "reussi":
                                icone = "✅"
                            else:
                                icone = "❌"
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
                        criteres_manques = [c for c in exercice["criteres"] if c["statut"] != "reussi" and not c.get("corrige_manuellement")]
                        if criteres_manques:
                            lignes_feedback = []
                            for critere in criteres_manques:
                                label = critere.get("competence", critere["description"])
                                resume = resumer_echecs(critere.get("details", []))
                                lignes_feedback.append(
                                    f"- **{label}** : {resume}" if resume else f"- À revoir : **{label}**"
                                )
                            st.markdown("**Feedback :**\n" + "\n".join(lignes_feedback))

                        # --- Correction manuelle ---
                        with st.expander("✏️ Correction manuelle"):
                            corrections_en_attente = {}
                            for idx_c, critere in enumerate(exercice["criteres"]):
                                label = critere.get("competence", critere["description"])
                                key_base = f"manuel_{idx_r}_{idx_e}_{idx_c}"
                                col1, col2 = st.columns([2, 3])
                                with col1:
                                    pts_actuels = critere["points_obtenus"]
                                    pts_max = critere["points_max"]
                                    nouveau_pts = st.number_input(
                                        f"{label}",
                                        min_value=0.0,
                                        max_value=float(pts_max),
                                        value=float(round(pts_actuels, 2)),
                                        step=0.5,
                                        key=key_base + "_pts",
                                    )
                                with col2:
                                    motif = st.text_input(
                                        "Motif",
                                        value=critere.get("motif_manuel", ""),
                                        placeholder="Ex : travail partiel accepté",
                                        key=key_base + "_motif",
                                    )
                                corrections_en_attente[idx_c] = (nouveau_pts, motif)

                            if st.button("Appliquer les corrections", key=f"btn_manuel_{idx_r}_{idx_e}"):
                                for idx_c, (pts, motif) in corrections_en_attente.items():
                                    critere = exercice["criteres"][idx_c]
                                    pts_anciens = critere["points_obtenus"]
                                    delta = pts - pts_anciens
                                    critere["points_obtenus"] = pts
                                    critere["statut"] = "reussi" if pts >= critere["points_max"] - 1e-6 else "echoue"
                                    critere["corrige_manuellement"] = (pts != pts_anciens or bool(motif))
                                    if motif:
                                        critere["motif_manuel"] = motif
                                    # Recalculer les totaux
                                    exercice["points_obtenus"] += delta
                                    resultat["points_obtenus"] += delta
                                st.success("Corrections appliquées.")
                                st.rerun()

# --- Onglet 3 : Rapports ------------------------------------------------
with onglet_rapports:
    st.subheader("Rapports")
    resultats_rapports = st.session_state.get("resultats_correction", [])

    if not resultats_rapports:
        st.info("Lancez d'abord une correction dans l'onglet 'Correction'.")
    else:
        st.write(f"{len(resultats_rapports)} copie(s) corrigee(s) disponible(s).")

        col_btn1, col_btn2 = st.columns(2)

        with col_btn1:
            st.markdown("**Rapports individuels**")
            st.caption("Un PDF par étudiant (score, compétences, feedback), regroupés dans un ZIP.")
            if st.button("Générer les rapports étudiants (ZIP)"):
                # Charger l'historique pour enrichir les PDF individuels
                historiques_pdf = {}
                if DOSSIER_RESULTATS.exists():
                    for f in DOSSIER_RESULTATS.glob("*.json"):
                        try:
                            r_hist = json.loads(f.read_text(encoding="utf-8"))
                            idf_h = r_hist["identite"]
                            cle_h = (idf_h["nom"], idf_h["prenom"])
                            historiques_pdf.setdefault(cle_h, []).append(r_hist)
                        except Exception:
                            pass
                    # Trier par session/année et exclure la session courante
                    sessions_courantes = {(r["session"], r["annee"]) for r in resultats_rapports}
                    for cle_h in historiques_pdf:
                        historiques_pdf[cle_h] = sorted(
                            [r for r in historiques_pdf[cle_h] if (r["session"], r["annee"]) not in sessions_courantes],
                            key=lambda x: (x["annee"], x["session"])
                        )

                with tempfile.TemporaryDirectory() as tmp:
                    chemin_zip = Path(tmp) / "rapports.zip"
                    generer_zip_rapports(
                        resultats_rapports, chemin_zip, tmp,
                        nom_cours=configuration_cours.get("nom_cours", ""),
                        code_cours=configuration_cours.get("code_cours", ""),
                        historiques=historiques_pdf,
                    )
                    st.download_button(
                        "Télécharger le ZIP des rapports",
                        data=chemin_zip.read_bytes(),
                        file_name="rapports_etudiants.zip",
                        mime="application/zip",
                    )

        with col_btn2:
            st.markdown("**Rapport de classe**")
            st.caption("Un seul PDF avec statistiques, graphiques et détail des compétences.")
            if st.button("Générer le rapport de classe (PDF)", disabled=len(resultats_rapports) < 2):
                with tempfile.TemporaryDirectory() as tmp:
                    chemin_pdf = Path(tmp) / "rapport_classe.pdf"
                    generer_rapport_classe_pdf(
                        resultats_rapports, chemin_pdf,
                        nom_cours=configuration_cours.get("nom_cours", ""),
                        code_cours=configuration_cours.get("code_cours", ""),
                    )
                    st.download_button(
                        "Télécharger le rapport de classe",
                        data=chemin_pdf.read_bytes(),
                        file_name="rapport_classe.pdf",
                        mime="application/pdf",
                    )
            if len(resultats_rapports) < 2:
                st.caption("(nécessite au moins 2 copies corrigées)")

        # Aperçu des graphiques dans l'onglet
        if len(resultats_rapports) >= 2:
            st.markdown("---")
            st.markdown("**Aperçu — Analyse de la classe**")

            import altair as alt
            bareme_r = resultats_rapports[0]["bareme_total"] or 1
            pourcentages_r = [round(100 * r["points_obtenus"] / bareme_r, 1) for r in resultats_rapports]
            moyenne_r = round(sum(pourcentages_r) / len(pourcentages_r), 1)
            mediane_r = round(sorted(pourcentages_r)[len(pourcentages_r) // 2], 1)
            nb_reussi_r = sum(1 for p in pourcentages_r if p >= 50)

            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Moyenne", f"{moyenne_r} %")
            m2.metric("Médiane", f"{mediane_r} %")
            m3.metric("Réussite", f"{nb_reussi_r}/{len(pourcentages_r)}")
            m4.metric("Min / Max", f"{min(pourcentages_r):.0f} % / {max(pourcentages_r):.0f} %")

            df_hist_r = pd.DataFrame({"Score (%)": pourcentages_r})
            hist_r = (
                alt.Chart(df_hist_r)
                .mark_bar(color="#5BC4E8", stroke="#2a8caa", strokeWidth=1)
                .encode(
                    alt.X("Score (%):Q", bin=alt.Bin(step=10), title="Score (%)", scale=alt.Scale(domain=[0, 100])),
                    alt.Y("count():Q", title="Nb étudiants"),
                    tooltip=[alt.Tooltip("Score (%):Q", bin=alt.Bin(step=10)), "count():Q"],
                )
                .properties(title="Distribution des scores", height=220)
            )
            st.altair_chart(hist_r, use_container_width=True)

            comp_stats_r = {}
            for r in resultats_rapports:
                for ex in r["exercices"]:
                    for c in ex["criteres"]:
                        label = c.get("competence", c["description"])[:40]
                        if label not in comp_stats_r:
                            comp_stats_r[label] = {"reussi": 0, "total": 0}
                        comp_stats_r[label]["total"] += 1
                        if c["statut"] == "reussi":
                            comp_stats_r[label]["reussi"] += 1

            lignes_comp_r = [
                {"Compétence": label, "Taux (%)": round(100 * v["reussi"] / v["total"], 1),
                 "Réussis": v["reussi"], "Total": v["total"],
                 "Couleur": "#22c55e" if round(100 * v["reussi"] / v["total"], 1) >= 75
                             else "#f59e0b" if round(100 * v["reussi"] / v["total"], 1) >= 50 else "#ef4444"}
                for label, v in comp_stats_r.items()
            ]
            df_comp_r = pd.DataFrame(lignes_comp_r).sort_values("Taux (%)")
            COULEURS_D = ["#22c55e", "#f59e0b", "#ef4444"]
            bars_r = (
                alt.Chart(df_comp_r)
                .mark_bar()
                .encode(
                    alt.Y("Compétence:N", sort=None, title=None),
                    alt.X("Taux (%):Q", scale=alt.Scale(domain=[0, 100]), title="Taux de réussite (%)"),
                    color=alt.Color("Couleur:N", scale=alt.Scale(domain=COULEURS_D, range=COULEURS_D), legend=None),
                    tooltip=["Compétence:N", "Taux (%):Q", "Réussis:Q", "Total:Q"],
                )
                .properties(title="Taux de réussite par compétence", height=max(180, len(df_comp_r) * 35))
            )
            ligne_50_r = alt.Chart(pd.DataFrame({"x": [50]})).mark_rule(color="#888", strokeDash=[4, 3]).encode(x="x:Q")
            st.altair_chart(bars_r + ligne_50_r, use_container_width=True)

            meilleures_r = df_comp_r.nlargest(3, "Taux (%)")
            difficiles_r = df_comp_r.nsmallest(3, "Taux (%)")
            c_ok2, c_ko2 = st.columns(2)
            with c_ok2:
                st.success("**Mieux réussies**\n\n" + "\n\n".join(
                    f"- {row['Compétence']} ({row['Taux (%)']:.0f} %)" for _, row in meilleures_r.iterrows()
                ))
            with c_ko2:
                st.error("**Plus difficiles**\n\n" + "\n\n".join(
                    f"- {row['Compétence']} ({row['Taux (%)']:.0f} %)" for _, row in difficiles_r.iterrows()
                ))

# --- Onglet 4 : Historique ---------------------------------------------------
with onglet_historique:
    st.subheader("Historique des corrections")

    # Charger tous les JSON sauvegardés
    fichiers_hist = sorted(DOSSIER_RESULTATS.glob("*.json")) if DOSSIER_RESULTATS.exists() else []
    if not fichiers_hist:
        st.info("Aucune correction sauvegardée. Lancez une correction pour alimenter l'historique.")
    else:
        tous_resultats_hist = []
        for f in fichiers_hist:
            try:
                tous_resultats_hist.append(json.loads(f.read_text(encoding="utf-8")))
            except Exception:
                pass

        # --- Vue globale de la classe ---
        st.markdown("### Vue globale — progression de la classe")
        sessions_stats = {}
        for r in tous_resultats_hist:
            cle = f"{r['session']} {r['annee']} ({r['module']})"
            bareme_h = r["bareme_total"] or 1
            pct = round(100 * r["points_obtenus"] / bareme_h, 1)
            if cle not in sessions_stats:
                sessions_stats[cle] = []
            sessions_stats[cle].append(pct)

        # Index session → config_epreuve (première copie qui en possède une)
        configs_par_session = {}
        for r in tous_resultats_hist:
            cle = f"{r['session']} {r['annee']} ({r['module']})"
            if cle not in configs_par_session and r.get("config_epreuve"):
                configs_par_session[cle] = r["config_epreuve"]

        if len(sessions_stats) >= 1:
            lignes_glob = []
            for session_label, pcts in sorted(sessions_stats.items()):
                a_config = "✅" if session_label in configs_par_session else "—"
                lignes_glob.append({
                    "Session": session_label,
                    "Nb copies": len(pcts),
                    "Moyenne (%)": round(sum(pcts) / len(pcts), 1),
                    "Min (%)": min(pcts),
                    "Max (%)": max(pcts),
                    "Réussis": sum(1 for p in pcts if p >= 50),
                    "Épreuve": a_config,
                })
            df_glob = pd.DataFrame(lignes_glob)
            st.dataframe(df_glob, hide_index=True)

            # Détail de l'épreuve par session
            sessions_avec_config = [s for s in sorted(sessions_stats.keys()) if s in configs_par_session]
            if sessions_avec_config:
                session_detail = st.selectbox(
                    "Consulter l'épreuve d'une session", ["— Choisir —"] + sessions_avec_config,
                    key="hist_session_detail"
                )
                if session_detail != "— Choisir —":
                    cfg = configs_par_session[session_detail]
                    with st.expander(f"Épreuve : {session_detail}", expanded=True):
                        st.caption(
                            f"Module : {cfg.get('module', '?')} · "
                            f"Variante : {cfg.get('variante', '?')} · "
                            f"Contexte : {cfg.get('contexte', '?')}"
                        )
                        # Tableau des compétences évaluées
                        exercices_cfg = cfg.get("exercices", [])
                        lignes_ep = []
                        for ex in exercices_cfg:
                            for c in ex.get("criteres", []):
                                lignes_ep.append({
                                    "Exercice": ex.get("titre", "?"),
                                    "Compétence": c.get("competence", c.get("description", "?")),
                                    "Barème": c.get("points_max", "?"),
                                })
                        if lignes_ep:
                            df_ep = pd.DataFrame(lignes_ep)
                            total_ep = sum(
                                c.get("points_max", 0)
                                for ex in exercices_cfg
                                for c in ex.get("criteres", [])
                                if isinstance(c.get("points_max"), (int, float))
                            )
                            st.dataframe(df_ep, hide_index=True)
                            st.caption(f"Barème total : {total_ep} pt(s)")

            if len(sessions_stats) >= 2:
                import altair as alt
                courbe = (
                    alt.Chart(df_glob)
                    .mark_line(point=True, color="#5BC4E8")
                    .encode(
                        alt.X("Session:N", sort=None, title="Session"),
                        alt.Y("Moyenne (%):Q", scale=alt.Scale(domain=[0, 100]), title="Moyenne classe (%)"),
                        tooltip=["Session:N", "Moyenne (%):Q", "Nb copies:Q"],
                    )
                    .properties(title="Évolution de la moyenne de la classe", height=220)
                )
                st.altair_chart(courbe, use_container_width=True)

        st.markdown("---")

        # --- Vue par étudiant ---
        st.markdown("### Progression par étudiant")

        # Construire index {(nom, prenom): [résultats triés par session]}
        par_etudiant = {}
        for r in tous_resultats_hist:
            idf = r["identite"]
            cle_e = (idf["nom"], idf["prenom"])
            par_etudiant.setdefault(cle_e, []).append(r)
        for cle_e in par_etudiant:
            par_etudiant[cle_e].sort(key=lambda x: (x["annee"], x["session"]))

        # Sélecteur étudiant — label → clé pour éviter le split fragile sur les noms avec espaces
        labels_etudiants = {
            f"{nom} {prenom}".strip(): (nom, prenom)
            for nom, prenom in sorted(par_etudiant.keys())
        }
        etudiant_choisi = st.selectbox("Choisir un étudiant", list(labels_etudiants.keys()))
        if etudiant_choisi:
            cle_ch = labels_etudiants[etudiant_choisi]
            historique_etudiant = par_etudiant.get(cle_ch, [])

            if historique_etudiant:
                bareme_e = historique_etudiant[0]["bareme_total"] or 1

                # Métriques rapides
                pcts_e = [round(100 * r["points_obtenus"] / (r["bareme_total"] or 1), 1) for r in historique_etudiant]
                me1, me2, me3 = st.columns(3)
                me1.metric("Sessions", len(historique_etudiant))
                me2.metric("Dernier score", f"{pcts_e[-1]} %")
                if len(pcts_e) >= 2:
                    delta = round(pcts_e[-1] - pcts_e[-2], 1)
                    me3.metric("Évolution", f"{pcts_e[-1]} %", delta=f"{delta:+.1f} %")

                # Courbe de progression
                lignes_prog = [
                    {
                        "Session": f"{r['session']} {r['annee']}",
                        "Score (%)": round(100 * r["points_obtenus"] / (r["bareme_total"] or 1), 1),
                    }
                    for r in historique_etudiant
                ]
                df_prog = pd.DataFrame(lignes_prog)

                import altair as alt
                courbe_e = (
                    alt.Chart(df_prog)
                    .mark_line(point=True, color="#E0007A")
                    .encode(
                        alt.X("Session:N", sort=None),
                        alt.Y("Score (%):Q", scale=alt.Scale(domain=[0, 100])),
                        tooltip=["Session:N", "Score (%):Q"],
                    )
                    .properties(title=f"Progression de {etudiant_choisi}", height=200)
                )
                ligne_50_e = alt.Chart(pd.DataFrame({"y": [50]})).mark_rule(
                    color="#888", strokeDash=[4, 3]
                ).encode(y="y:Q")
                st.altair_chart(courbe_e + ligne_50_e, use_container_width=True)

                # Tableau comparaison compétences inter-sessions
                st.markdown("**Détail des compétences par session**")
                # Collecter toutes les compétences rencontrées
                toutes_comps = []
                for r in historique_etudiant:
                    for ex in r["exercices"]:
                        for c in ex["criteres"]:
                            label = c.get("competence", c["description"])
                            if label not in toutes_comps:
                                toutes_comps.append(label)

                lignes_comp_hist = []
                for comp in toutes_comps:
                    ligne = {"Compétence": comp}
                    for r in historique_etudiant:
                        session_label = f"{r['session']} {r['annee']}"
                        val = "—"
                        for ex in r["exercices"]:
                            for c in ex["criteres"]:
                                if c.get("competence", c["description"]) == comp:
                                    pts = c["points_obtenus"]
                                    pts_max = c["points_max"]
                                    icone = "✅" if c["statut"] == "reussi" else "❌"
                                    val = f"{icone} {pts:.1f}/{pts_max:.1f}"
                        ligne[session_label] = val
                    lignes_comp_hist.append(ligne)

                st.dataframe(pd.DataFrame(lignes_comp_hist), hide_index=True)

        st.markdown("---")
        st.caption(f"{len(fichiers_hist)} fichier(s) dans l'historique — dossier : data/resultats/")

        with st.expander("🗑️ Gestion de l'historique"):
            st.warning("La suppression est irréversible.")
            c_del1, c_del2 = st.columns(2)
            with c_del1:
                # Supprimer uniquement l'étudiant sélectionné
                if etudiant_choisi:
                    nom_s, prenom_s = labels_etudiants[etudiant_choisi]
                    fichiers_etudiant = [
                        f for f in fichiers_hist
                        if f.name.startswith(f"{nom_s}_{prenom_s}_")
                    ]
                    if st.button(f"Supprimer l'historique de {etudiant_choisi} ({len(fichiers_etudiant)} fichier(s))"):
                        for f in fichiers_etudiant:
                            f.unlink()
                        st.success(f"Historique de {etudiant_choisi} supprimé.")
                        st.rerun()
            with c_del2:
                if st.button("Supprimer tout l'historique", type="primary"):
                    for f in fichiers_hist:
                        f.unlink()
                    st.success("Historique entièrement effacé.")
                    st.rerun()
