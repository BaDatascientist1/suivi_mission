import streamlit as st
st.set_page_config(page_title="Dashboard Suivi de Mission", layout="wide")
st.title("📊 Dashboard de Suivi des Missions – Clientisgroup")
st.markdown("""
<style>
h1, h2, h3 {
    color: #003366;
}
div.stButton > button {
    background-color: #0059b3;
    color: white;
    font-weight: bold;
    border-radius: 8px;
}
div.stButton > button:hover {
    background-color: #003d73;
}
thead tr th {
    background-color: #e6f0ff !important;
    color: #003366;
    font-weight: bold;
}
</style>
""", unsafe_allow_html=True)

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import matplotlib.pyplot as plt
import seaborn as sns
import os
from openpyxl import load_workbook
from datetime import date
from datetime import timedelta
import unicodedata
from collections import Counter

# Chargement des données
# Chargement du fichier principal
# Titre principal

df_mission = pd.read_excel("dataset.xlsx", sheet_name="Sheet1")
df_mission = df_mission[df_mission["Services"].isin(["Conformité ISO", "Formation"])]
import re
import unicodedata

def clean_phase(texte):
    if pd.isna(texte):
        return ""
    texte = str(texte)
    texte = unicodedata.normalize("NFKD", texte).encode("ascii", "ignore").decode("utf-8")
    texte = re.sub(r"\s+", " ", texte)
    texte = texte.strip().lower().capitalize()
    return texte

df_mission["Phases"] = df_mission["Phases"].apply(clean_phase)


# Nettoyage minimal
def clean_phase(phase):
    if pd.isna(phase):
        return ""
    # Convertir en str, enlever accents, strip et mettre la 1re lettre en majuscule
    phase = str(phase).strip()
    phase = unicodedata.normalize("NFKD", phase).encode("ascii", "ignore").decode("utf-8")
    return phase.capitalize()

df_mission["Phases"] = df_mission["Phases"].apply(clean_phase)

df_mission = df_mission.dropna(subset=["Statut Avancement"])
df_mission["Statut Avancement"] = df_mission["Statut Avancement"].astype(str).str.strip().str.lower()

# Parsing des dates
for col in ["Date Début", "Date Elaboration Prévisionnelle","Date Elaboration Effective","Date CTCQ Prévisionnelle","Date CTCQ Effective","Date Approbation Prévisionnelle","Date Approbation Effective","Date Finalisation Prévisionnelle","Date Finalisation Effective","Date Facturation","Date Règlement"]:
    df_mission[col] = pd.to_datetime(df_mission[col], errors='coerce')

# Ajout des jours de retard
if "Retard (jours)" not in df_mission.columns:
    df_mission["Retard (jours)"] = (
        pd.to_datetime(df_mission["Date Finalisation Effective"], errors="coerce") -
        pd.to_datetime(df_mission["Date Finalisation Prévisionnelle"], errors="coerce")
    ).dt.days
    df_mission["Retard (jours)"] = df_mission["Retard (jours)"].apply(lambda x: x if x and x > 0 else 0)



# Tabs
tabs = st.tabs(["Vue d’ensemble","Suivi Responsabilités", "Suivi Opérationnel", "Suivi des Missions", "➕ Ajouter une mission"])

# Onglet 1 – Vue d'ensemble avec KPI
with tabs[0]:
    st.subheader("Vue d'ensemble des missions")

    # Chargement du fichier
    if "reload_df" in st.session_state and st.session_state["reload_df"]:
        df_mission = pd.read_excel("dataset.xlsx", sheet_name="Sheet1")
        st.session_state["reload_df"] = False
    else:
        df_mission = pd.read_excel("dataset.xlsx", sheet_name="Sheet1")

    # Fonction de carte KPI stylisée
    def styled_kpi(title, value, background_color="#FFFFFF", value_color="#0D47A1"):
        html = f"""
        <div style="background-color:{background_color}; padding:15px; border-radius:15px;
                    box-shadow:2px 2px 8px rgba(0,0,0,0.05); text-align:center; margin-bottom:10px;">
            <div style="font-size:16px; font-weight:500; color:#333;">{title}</div>
            <div style="font-size:26px; font-weight:bold; margin-top:5px; color:{value_color};">{value}</div>
        </div>
        """
        st.markdown(html, unsafe_allow_html=True)

    # KPI initiaux
    col1, col2, col3 = st.columns(3)
    with col1:
        styled_kpi("📌 Nombre de missions", df_mission["ID_Mission"].nunique())
    with col2:
        styled_kpi("📄 Nombre Activités", df_mission["Activités"].shape[0])
    with col3:
        styled_kpi("📦 Nombre de livrables", df_mission["Livrables"].nunique())

    # Ajout de colonnes manquantes
    if "Commentaires" not in df_mission.columns:
        df_mission["Commentaires"] = ""
    if "Ref" not in df_mission.columns:
        st.error("❌ La colonne 'Ref' est nécessaire.")
    else:
        # Réorganisation
        colonnes_affichage = ["ID_Mission","Missions", "Services", "Porteurs", "Phases", "Activités",
                              "Livrables", "Date Début", "Date Elaboration Prévisionnelle", "Date Elaboration Effective",
                              "Responsable Elaboration","Satisfaction Elaboration",
                              "Date CTCQ Prévisionnelle", "Date CTCQ Effective","Responsable CTCQ","Satisfaction CTCQ", "Conformité",
                              "Date Approbation Prévisionnelle", "Date Approbation Effective","Responsable Approbation","Satisfaction Approbation",
                              "Date Finalisation Prévisionnelle", "Date Finalisation Effective","Satisfaction Globale","Date Satisfaction Client","Statut Avancement","Nom Client", "Commentaires","Date Facturation","Statut Règlement","Date Règlement","Code Projet Client","Zone Géographique"]
        df_vue = df_mission[colonnes_affichage + ["Ref"]].copy()
        df_vue = df_vue[df_vue["Services"].isin(["Conformité ISO", "Formation"])]
        #df_vue_affiche = df_vue[ [col for col in colonnes_affichage if col in df_vue.columns] ]

        #df_vue = df_vue[[col for col in df_vue.columns if col != "Ref"]]
        # Filtres
        st.write("### Filtres")
        col1, col2, col3, col4,col5,col6 = st.columns(6)
        Ref_missions = df_vue["ID_Mission"].fillna("(Inconnu)").unique().tolist()
        missions = df_vue["Missions"].fillna("(Inconnu)").unique().tolist()
        type_phases = df_vue["Phases"].fillna("(Inconnu)").unique().tolist()
        type_services = df_vue["Services"].fillna("(Inconnu)").unique().tolist()
        type_activites = df_vue["Activités"].fillna("(Inconnu)").unique().tolist()
        livrables = df_vue["Livrables"].fillna("(Inconnu)").unique().tolist()
        
        with col1:
            selected_RefMission = st.selectbox("Choisir un numéro mission", ["Tous"] + sorted(Ref_missions))
        with col2:
            selected_mission = st.selectbox("Choisir une mission", ["Toutes"] + sorted(missions))
        with col3:
            selected_phase = st.selectbox("Choisir une phase", ["Toutes"] + sorted(type_phases))
        with col4:
            selected_service = st.selectbox("Choisir un service ", ["Tous"] + sorted(type_services))
        with col5:
            selected_activite = st.selectbox("Choisir une activité", ["Toutes"] + sorted(type_activites))       
        with col6:
            selected_livrable = st.selectbox("Choisir un livrable", ["Tous"] + sorted(livrables))

        # Application des filtres
        
        filtered_df = df_vue.copy()
        if selected_RefMission != "Tous":
            filtered_df = filtered_df[filtered_df["ID_Mission"].fillna("(Inconnu)") == selected_RefMission]
        if selected_mission != "Toutes":
            filtered_df = filtered_df[filtered_df["Missions"].fillna("(Inconnu)") == selected_mission]
        if selected_phase != "Toutes":
            filtered_df = filtered_df[filtered_df["Phases"].fillna("(Inconnu)") == selected_phase]
        if selected_service != "Tous":
            filtered_df = filtered_df[filtered_df["Services"].fillna("(Inconnu)") == selected_service]
        if selected_activite != "Toutes":
            filtered_df = filtered_df[filtered_df["Activités"].fillna("(Inconnu)") == selected_activite]
        if selected_livrable != "Tous":
            filtered_df = filtered_df[filtered_df["Livrables"].fillna("(Inconnu)") == selected_livrable]

        # KPI dynamiques
        st.markdown("### 📊 Indicateurs de performance")

        nb_total = len(filtered_df)
        nb_realisees = filtered_df["Date Finalisation Effective"].notna().sum()
        nb_conformes = filtered_df["Conformité"].str.upper().eq("OUI").sum()
        nb_nonconformes = filtered_df["Conformité"].str.upper().eq("NON").sum()
        nb_nonApplicables = filtered_df["Conformité"].str.upper().eq("NON APPLICABLE").sum()

        taux_action = (nb_realisees / nb_total) * 100 if nb_total > 0 else 0
        taux_conformite = (nb_conformes / nb_total) * 100 if nb_total > 0 else 0
        taux_nonconformite = (nb_nonconformes / nb_total) * 100 if nb_total > 0 else 0
        taux_nonApplicable = (nb_nonApplicables / nb_total) * 100 if nb_total > 0 else 0

        def kpi_card(title, value, delta, background_color, value_color):
            html = f"""
            <div style="background-color:{background_color}; padding:15px; border-radius:15px;
                        box-shadow:2px 2px 8px rgba(0,0,0,0.05); text-align:center;">
                <div style="font-size:16px; font-weight:500; color:#333;">{title}</div>
                <div style="font-size:26px; font-weight:bold; margin-top:5px; color:{value_color};">{value}</div>
                <div style="font-size:14px; opacity:0.85; color:{value_color};">{delta}</div>
            </div>
            """
            st.markdown(html, unsafe_allow_html=True)

        white = "#FFFFFF"
        blue_dark = "#0D47A1"
        col1, col2, col3, col4 = st.columns(4)

        with col1:
            kpi_card("✅ Actions réalisées", f"{nb_realisees}/{nb_total}", f"{taux_action:.1f} %",blue_dark, white)
        with col2:
            kpi_card("📋 Conformes", f"{nb_conformes}", f"{taux_conformite:.1f} %", white, blue_dark)
        with col3:
            kpi_card("⚠️ Non conformes", f"{nb_nonconformes}", f"{taux_nonconformite:.1f} %", white, blue_dark)
        with col4:
            kpi_card("❔ Non applicables", f"{nb_nonApplicables}", f"{taux_nonApplicable:.1f} %", white, blue_dark)

        # Masquer colonne Ref à l'affichage
        colonnes_affichees = [col for col in filtered_df.columns if col != "Ref"]

        # Suivi des missions
        st.write("### Tableau de suivi des missions")
        filtered_df["Date Approbation Effective"] = pd.to_datetime(filtered_df["Date Approbation Effective"], errors="coerce")
        filtered_df["Date Satisfaction Client"] = pd.to_datetime(filtered_df["Date Satisfaction Client"], errors="coerce")
        filtered_df["Date Facturation"] = pd.to_datetime(filtered_df["Date Facturation"], errors="coerce")
        filtered_df["Date Règlement"] = pd.to_datetime(filtered_df["Date Règlement"], errors="coerce")
        filtered_df["Responsable CTCQ"] = filtered_df["Responsable CTCQ"].astype(str)
        filtered_df["Responsable Elaboration"] = filtered_df["Responsable Elaboration"].astype(str)
        filtered_df["Responsable Approbation"] = filtered_df["Responsable Approbation"].astype(str)
        filtered_df["Code Projet Client"] = filtered_df["Code Projet Client"].astype(str)
        filtered_df["Zone Géographique"] = filtered_df["Zone Géographique"].astype(str)
        filtered_df["Nom Client"] = filtered_df["Nom Client"].astype(str)



        edited_df = st.data_editor(
            filtered_df,
            use_container_width=True,
            num_rows="dynamic",
        column_config={
            "Ref": st.column_config.TextColumn("Ref", disabled=True, width="small"),
            "Missions": st.column_config.SelectboxColumn(
                "Missions", options=["CO", "GO", "Inspection", "Évaluation", "Autre"]
            ),
            "Type de Missions": st.column_config.SelectboxColumn(
                "Type de Missions", options=["CO", "GO", "Inspection", "Évaluation", "Autre"]
            ),
            "Conformité": st.column_config.SelectboxColumn(
                "Conformité", options=["OUI", "NON", "Non Applicable"]
            ),
            "Commentaires": st.column_config.TextColumn("Commentaires"),
            "Date Elaboration Effective": st.column_config.DateColumn(
                label="Date Elaboration Effective", format="YYYY-MM-DD"
            ),
            "Date CTCQ Effective": st.column_config.DateColumn(
                label="Date CTCQ Effective", format="YYYY-MM-DD"
            ),
            "Date Approbation Effective": st.column_config.DateColumn(
                label="Date Approbation Effective", format="YYYY-MM-DD"
            ),
            "Date Finalisation Effective": st.column_config.DateColumn(
                label="Date Finalisation Effective", format="YYYY-MM-DD"
            ),
            "Date Début": st.column_config.DateColumn(
                label="Date Début", format="YYYY-MM-DD"
            ),
            "Date Satisfaction Client": st.column_config.DateColumn(
                label="Date Satisfaction Client", format="YYYY-MM-DD"
            ),
            "Date Facturation": st.column_config.DateColumn(
                label="Date Facturation", format="YYYY-MM-DD"
            ),
            "Date Règlement": st.column_config.DateColumn(
                label="Date Règlement", format="YYYY-MM-DD"
            ),
            "Statut Avancement": st.column_config.SelectboxColumn(
                "Statut Avancement", options=["Non entamé", "En cours", "Bloqué", "Clôturé", "Clôturé avec retard"]
            ),
            "Statut Règlement": st.column_config.SelectboxColumn(
                "Statut Règlement", options=["Réglé", "En attente", "Partiellement réglé", "Non réglé"]
            ),
            "Satisfaction Globale": st.column_config.SelectboxColumn(
                "Satisfaction Globale", options=["1", "2", "3", "4","5"]
            ),
            "Satisfaction Elaboration": st.column_config.SelectboxColumn(
                "Satisfaction Elaboration", options=["1", "2", "3", "4","5"]
            ),
            "Satisfaction CTCQ": st.column_config.SelectboxColumn(
                "Satisfaction CTCQ", options=["1", "2", "3", "4","5"]
            ),
            "Satisfaction Approbation": st.column_config.SelectboxColumn(
                "Satisfaction Approbation",  options=["1", "2", "3", "4","5"]
            ),
            "Code Projet Client": st.column_config.TextColumn("Code Projet Client"),
            "Zone Géographique": st.column_config.TextColumn("Zone Géographique"),
            "Responsable Elaboration": st.column_config.TextColumn("Responsable Elaboration"),
            "Responsable CTCQ": st.column_config.TextColumn("Responsable CTCQ"),
            "Responsable Approbation": st.column_config.TextColumn("Responsable Approbation"),
            "Porteurs": st.column_config.TextColumn("Porteurs"),
            "Services": st.column_config.TextColumn("Services"),
            "Nom Client": st.column_config.TextColumn("Nom Client")
        }
    )
        
        import io

        # Exportation en Excel via un bouton
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            filtered_df.to_excel(writer, index=False, sheet_name="Missions")
            writer.close()  # <-- ou juste ne rien mettre, le with le fait déjà
        
        st.download_button(
            label="📥 Télécharger le tableau filtré (Excel)",
            data=output.getvalue(),
            file_name="missions_filtrees.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )


        # Mise à jour des données si Ref existe
        df_original = pd.read_excel("dataset.xlsx")
        if "Ref" in edited_df.columns and "Ref" in df_original.columns:
            edited_df_clean = edited_df[edited_df["Ref"].notna()].copy()
            edited_df_clean.set_index("Ref", inplace=True)
            df_original.set_index("Ref", inplace=True)
            edited_df_clean = edited_df_clean[edited_df_clean.index.isin(df_original.index)]
            for col in edited_df_clean.columns:
                df_original.loc[edited_df_clean.index, col] = edited_df_clean[col]
            df_original.reset_index(inplace=True)
            try:
                df_original.to_excel("dataset.xlsx", index=False)
                st.success("✅ Modifications enregistrées avec succès.")
            except PermissionError:
                st.error("❌ Veuillez fermer 'dataset.xlsx' puis réessayer.")
        else:
            st.error("❌ La colonne 'Ref' est requise dans les deux tables.")
    #-----------------------------------------------------------------------
    #-----------------------------------------------------------------------
    # Ensemnles des indicateurs de performances Cles
    # Calculs réels à partir du df_vue (supposé filtré sur ISO/Formation)
    #-----------------------------------------------------------------------
    df_kpi = df_vue.copy()
    
    # 🔧 Style CSS moderne
    st.markdown("""
    <style>
    .kpi-card {
        background-color: #ffffff;
        padding: 16px;
        margin-bottom: 10px;
        border-left: 6px solid #1f77b4;
        border-radius: 12px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.07);
    }
    .kpi-title {
        font-size: 15px;
        font-weight: 600;
        color: #34495e;
        margin-bottom: 4px;
    }
    .kpi-value {
        font-size: 24px;
        font-weight: bold;
        color: #1f77b4;
    }
    .kpi-sub {
        font-size: 13px;
        color: #7f8c8d;
    }
    </style>
    """, unsafe_allow_html=True)
    
    # 🔧 Fonction pour afficher un KPI modernisé
    def afficher_kpi(col, titre, valeur, sous_valeur=None, icone="📊"):
        col.markdown(f"""
        <div class="kpi-card">
            <div class="kpi-title">{icone} {titre}</div>
            <div class="kpi-value">{valeur}</div>
            {'<div class="kpi-sub">' + sous_valeur + '</div>' if sous_valeur else ''}
        </div>
        """, unsafe_allow_html=True)

    
    st.markdown("#### 🔵 Suivi des livrables et jalons")

    total_livrable = df_kpi.shape[0]
    finalises = df_kpi["Date Finalisation Effective"].notna().sum()
    controles = df_kpi["Date CTCQ Effective"].notna().sum()
    approbations = df_kpi["Date Approbation Effective"].notna().sum()
    retard_livrable = df_kpi[df_kpi["Date Finalisation Effective"] > df_kpi["Date Finalisation Prévisionnelle"]].shape[0]
    
    delai_production = (pd.to_datetime(df_kpi["Date Finalisation Effective"]) - pd.to_datetime(df_kpi["Date Début"])).dt.days.mean()
    delai_elab_ctcq = (pd.to_datetime(df_kpi["Date CTCQ Effective"]) - pd.to_datetime(df_kpi["Date Elaboration Effective"])).dt.days.mean()
    
    col1, col2, col3 = st.columns(3)
    afficher_kpi(col1, "% Livrables finalisés", f"{finalises / total_livrable:.0%}", f"{finalises} / {total_livrable}", "✅")
    afficher_kpi(col2, "% Validés (CTCQ)", f"{controles / total_livrable:.0%}", f"{controles} / {total_livrable}", "📄")
    afficher_kpi(col3, "% Approuvés", f"{approbations / total_livrable:.0%}", f"{approbations} / {total_livrable}", "🗂️")
    
    col1, col2, col3 = st.columns(3)
    afficher_kpi(col1, "% Retard", f"{retard_livrable / total_livrable:.0%}", f"{retard_livrable} en retard", "⏰")
    afficher_kpi(col2, "Délai moyen (prod.)", f"{delai_production:.1f} j", "Début → Finalisation", "🕓")
    afficher_kpi(col3, "Délai élab. → CTCQ", f"{delai_elab_ctcq:.1f} j", "Élaboration → Contrôle", "📈")


    st.markdown("#### 🔵 Suivi budgétaire et facturation")

    factures = df_kpi["Date Facturation"].notna().sum()
    total_mission = df_kpi["ID_Mission"].shape[0]
    regles = df_kpi[df_kpi["Statut Règlement"].astype(str).str.lower() == "réglé"].shape[0]
    delais_reglement = (pd.to_datetime(df_kpi["Date Règlement"]) - pd.to_datetime(df_kpi["Date Facturation"])).dt.days.mean()
    montant_attente = df_kpi[df_kpi["Statut Règlement"].astype(str).str.lower() != "réglé"].shape[0]
    
    col1, col2, col3 = st.columns(3)
    afficher_kpi(col1, "% Facturées", f"{factures / total_mission:.0%}", f"{factures} / {total_mission}", "🧾")
    afficher_kpi(col2, "Délai moyen règlement", f"{delais_reglement:.1f} j", "Facture → Règlement", "⏳")
    afficher_kpi(col3, "% Règlements reçus", f"{regles / factures:.0%}" if factures else "0%", f"{regles} / {factures}", "💰")
    st.markdown(f"**📌 Missions non réglées :** `{montant_attente}`")


    st.markdown("#### 🔵 Satisfaction client")

    total_satisfaction = df_kpi.shape[0]
    retour = df_kpi["Date Satisfaction Client"].notna().sum()
    notes = df_kpi["Satisfaction Globale"].dropna()
    satisfaits = notes[notes >= 7].count()
    
    col1, col2, col3 = st.columns(3)
    afficher_kpi(col1, "% de retours", f"{retour / total_satisfaction:.0%}", f"{retour} / {total_satisfaction}", "📥")
    afficher_kpi(col2, "Note moyenne", f"{notes.mean():.1f}" if not notes.empty else "-", "", "⭐")
    afficher_kpi(col3, "% Clients satisfaits", f"{satisfaits / total_satisfaction:.0%}", f"{satisfaits} / {total_satisfaction}", "👍")


    


# Onglet 2 - Responsabilites
with tabs[1]:
    st.subheader("🎯 Suivi des responsabilités")      
    
    # Nettoyer et récupérer les collaborateurs uniques
    collab_col = "Porteurs"
    collaborateurs = df_vue[collab_col].dropna().unique().tolist()
    
    # Générer les colonnes dynamiquement (max 6 pour lisibilité)
    colonnes = st.columns(min(len(collaborateurs), 6))
    
    # Affichage des KPI par collaborateur
    for i, collab in enumerate(collaborateurs):
        df_c = df_vue[df_vue[collab_col] == collab]
        total = df_c.shape[0]
        finalises = df_c["Date Finalisation Effective"].notna().sum()
        avance = df_c[df_c["Date Finalisation Effective"] < df_c["Date Finalisation Prévisionnelle"]].shape[0]
        retard = df_c[df_c["Date Finalisation Effective"] > df_c["Date Finalisation Prévisionnelle"]].shape[0]
        note =  df_c["Satisfaction Globale"].dropna()
        commentaires = df_c["Commentaires"].dropna().astype(str).str.lower()
        # Nettoyer et découper les mots
        mots = []
        
        for ligne in commentaires:
            mots += re.findall(r'\b[a-zéèàêîôûç]+\b', ligne)  # mots simples uniquement
        
        # Compter les occurrences
        compteur = Counter(mots)
        mots_frequents = compteur.most_common(3)  # Top 5 mots

        
        with colonnes[i % len(colonnes)]:
            st.markdown(f"### 👤 {collab}")
            st.metric("🧑‍💼 Livrables confiés", total)
            st.metric("✔️ Finalisés", f"{finalises / total:.0%}" if total else "0%")
            st.metric("✅ Livrés en avance", f"{avance / total:.0%}" if total else "0%")
            st.metric("📉 Retard", f"{retard / total:.0%}" if total else "0%")
            st.metric("⭐ Mean satisfaction", f"{notes.mean():.1f}" if not notes.empty else "-")
            if mots_frequents:
                st.markdown("**🔤 Mots fréquents :**")
                for mot, freq in mots_frequents:
                    st.markdown(f"- {mot} ({freq})")
            else:
                st.markdown("_Aucun commentaire exploitable_")
    # ⬅️ Création des filtres horizontaux
    colf1, colf2 = st.columns(2)
    
    with colf1:
        filtre_elab = st.selectbox("✍️ Responsable Elaboration", ["Tous"] + sorted(df_vue["Responsable Elaboration"].dropna().unique().tolist()))
    with colf2:
        filtre_ctcq = st.selectbox("🧪 Responsable CTCQ", ["Tous"] + sorted(df_vue["Responsable CTCQ"].dropna().unique().tolist()))
    df_kpi = df_vue.copy()
    

    # Application des filtres
    
    if filtre_elab != "Tous":
        df_kpi = df_kpi[df_kpi["Responsable Elaboration"].astype(str).str.strip() == filtre_elab]
    
    if filtre_ctcq != "Tous":
        df_kpi = df_kpi[df_kpi["Responsable CTCQ"].astype(str).str.strip() == filtre_ctcq]
        
    st.markdown("#### 🔵 Niveau d’avancement par personne")
    total_personnel = df_kpi.shape[0]
    finalises = df_kpi["Date Finalisation Effective"].notna().sum()
    controles = df_kpi["Date CTCQ Effective"].notna().sum()
    en_cours = df_kpi[df_kpi["Statut Avancement"].str.lower() == "en cours"].shape[0]
    notes = df_kpi["Satisfaction Globale"].dropna()

    col1, col2 = st.columns(2)
    afficher_kpi(col1, "% livrables Finalisés", f"{finalises / total_personnel:.0%}", f"{finalises} / {total_personnel}", "✔️")
    afficher_kpi(col2, "% livrables Validés", f"{controles / total_personnel:.0%}", f"{controles} / {total_personnel}", "📋")
    afficher_kpi(col1, "Nombre de livrables en cours", en_cours, "", "🔄")
    afficher_kpi(col2, "Total gérés", total_personnel, "", "🧑‍💼")
    afficher_kpi(col1, "Moyenne Satisfaction", f"{notes.mean():.1f}" if not notes.empty else "-", "", "⭐")

    
    st.markdown("#### 🕓 Ponctualité (Retards ou Avances)")
    
    retards_ponct = df_kpi[df_kpi["Date Finalisation Effective"] > df_kpi["Date Finalisation Prévisionnelle"]].shape[0]
    avances = df_kpi[df_kpi["Date Finalisation Effective"] < df_kpi["Date Finalisation Prévisionnelle"]].shape[0]
    total_ponctualite = df_kpi.shape[0]
    
    delta_retard = (pd.to_datetime(df_kpi["Date Finalisation Effective"]) - pd.to_datetime(df_kpi["Date Finalisation Prévisionnelle"]))
    moyenne_retard = delta_retard[delta_retard.dt.days > 0].dt.days.mean()
    moyenne_avance = delta_retard[delta_retard.dt.days < 0].dt.days.abs().mean()
    
    col1, col2 = st.columns(2)
    afficher_kpi(col1, "% de livrables livrés en retard", f"{retards_ponct / total_ponctualite:.0%}", f"{retards_ponct} en retard", "📉")
    afficher_kpi(col2, "% de livrables livrés en avance", f"{avances / total_ponctualite:.0%}", f"{avances} en avance", "📈")
    afficher_kpi(col1, "Délai moyen dépassement", f"{moyenne_retard:.1f} j" if not pd.isna(moyenne_retard) else "-", "", "⏰")
    afficher_kpi(col2, "Délai moyen anticipation", f"{moyenne_avance:.1f} j" if not pd.isna(moyenne_avance) else "-", "", "🚀")


    st.markdown("#### ✅ Indice de fiabilité individuelle")
    total_fiabilite = df_kpi.shape[0]
    respect_jalons = df_kpi[df_kpi["Date Finalisation Effective"] <= df_kpi["Date Finalisation Prévisionnelle"]].shape[0]
    commentaires = df_kpi["Commentaires"].astype(str).str.lower()
    positifs = commentaires.str.count("livr|ok|fait").sum()
    negatifs = commentaires.str.count("bloqu|retard|en attente").sum()
    blocages = commentaires.str.contains("bloqu").sum()
    
    notes = df_kpi["Satisfaction Globale"].dropna()
    
    col1, col2 = st.columns(2)
    afficher_kpi(col1, "% Respect des jalons", f"{respect_jalons / total_fiabilite:.0%}", f"{respect_jalons} / {total_fiabilite}", "🎯")
    afficher_kpi(col2, "Blocages signalés", blocages, "", "⚠️")
    afficher_kpi(col1, "Nombre de Commentaires + / −", f"{int(positifs)} / {int(negatifs)}", "", "💬")
    afficher_kpi(col2, "Note moy.satisfaction livrables", f"{notes.mean():.1f}" if not notes.empty else "-", "", "🌟")

# Onglet 3 – Visualisations
with tabs[2]:
    st.subheader("Suivi opérationnel")
         # Analyse des retards intermédiaires (élaboration → CTCQ → approbation)
    st.subheader("⏱️ Retards par étape intermédiaire")

    # Conversion des dates au cas où
    for col in ["Date Début", "Date Elaboration Prévisionnelle","Date Elaboration Effective", "Date CTCQ Prévisionnelle","Date CTCQ Effective", "Date Approbation Prévisionnelle","Date Approbation Effective"]:
        df_mission[col] = pd.to_datetime(df_mission[col], errors='coerce')

    # Calcul des durées
    df_mission["Duree_Elaboration"] = (df_mission["Date Elaboration Effective"] - df_mission["Date Début"]).dt.days
    df_mission["Duree_CTCQ"] = (df_mission["Date CTCQ Effective"] - df_mission["Date Elaboration Effective"]).dt.days
    df_mission["Duree_Approbation"] = (df_mission["Date Approbation Effective"] - df_mission["Date CTCQ Effective"]).dt.days

    # Comparaison aux seuils
    df_mission["Retard_Elaboration"] = df_mission["Date Elaboration Effective"] > df_mission["Date Elaboration Prévisionnelle"]
    df_mission["Retard_CTCQ"] = df_mission["Date CTCQ Effective"] > df_mission["Date CTCQ Prévisionnelle"]
    df_mission["Retard_Approbation"] = df_mission["Date Approbation Effective"] > df_mission["Date Approbation Prévisionnelle"]

    # Comptage des retards
    retard_intermediaire = {
        "Élaboration": df_mission["Retard_Elaboration"].sum(),
        "CT/CQ": df_mission["Retard_CTCQ"].sum(),
        "Approbation": df_mission["Retard_Approbation"].sum()
    }

    total = len(df_mission)

     # 🔸 Pourcentages de réalisations avec ou sans retard intermédiaire
    total_valides = df_mission[["Retard_Elaboration", "Retard_CTCQ", "Retard_Approbation"]].notna().all(axis=1).sum()

    nb_sans_retard_inter = (
        (~df_mission["Retard_Elaboration"] & 
         ~df_mission["Retard_CTCQ"] & 
         ~df_mission["Retard_Approbation"]).sum()
    )
    nb_avec_retard_inter = total_valides - nb_sans_retard_inter

    pct_sans_retard_inter = (nb_sans_retard_inter / total_valides) * 100 if total_valides > 0 else 0
    pct_avec_retard_inter = (nb_avec_retard_inter / total_valides) * 100 if total_valides > 0 else 0

    # 🔸 Pourcentages de réalisations avec ou sans retard global
    df_dates = df_mission[["Date Finalisation Prévisionnelle", "Date Finalisation Effective"]].dropna()
    total_realisees = len(df_dates)

    nb_sans_retard_global = (df_dates["Date Finalisation Effective"] <= df_dates["Date Finalisation Prévisionnelle"]).sum()
    nb_avec_retard_global = total_realisees - nb_sans_retard_global

    pct_sans_retard_global = (nb_sans_retard_global / total_realisees) * 100 if total_realisees > 0 else 0
    pct_avec_retard_global = (nb_avec_retard_global / total_realisees) * 100 if total_realisees > 0 else 0

    # 🌟 Affichage des KPI sous forme d'étiquettes stylisées
    st.markdown("### 📊 Indicateurs de performance des actions")

    col1, col2, col3, col4 = st.columns(4)

    with col1:
        st.markdown(f"""
        <div style="background-color:#28a745;padding:10px;border-radius:10px;text-align:center;color:white;">
            ✅<br><b>{pct_sans_retard_inter:.1f}%</b><br>Taux de réalisations des actions sans retard intermédiaire
        </div>
        """, unsafe_allow_html=True)

    with col2:
        st.markdown(f"""
        <div style="background-color:#dc3545;padding:10px;border-radius:10px;text-align:center;color:white;">
            ⚠️<br><b>{pct_avec_retard_inter:.1f}%</b><br> Taux de réalisations des actions avec retard intermédiaire
        </div>
        """, unsafe_allow_html=True)

    with col3:
        st.markdown(f"""
        <div style="background-color:#17a2b8;padding:10px;border-radius:10px;text-align:center;color:white;">
            ⏱️<br><b>{pct_sans_retard_global:.1f}%</b><br>Taux de réalisations des actions dans les délais
        </div>
        """, unsafe_allow_html=True)

    with col4:
        st.markdown(f"""
        <div style="background-color:#ffc107;padding:10px;border-radius:10px;text-align:center;color:black;">
            🕒<br><b>{pct_avec_retard_global:.1f}%</b><br>Taux de réalisations des actions hors délais
        </div>
        """, unsafe_allow_html=True)

    col1, col2 = st.columns(2)

    # Calcul des effectifs et des pourcentages
    statut_counts = df_mission["Statut Avancement"].value_counts().reset_index()
    statut_counts.columns = ["Statut Avancement", "Nombre"]
    statut_counts["Pourcentage"] = 100 * statut_counts["Nombre"] / statut_counts["Nombre"].sum()

    # Fonction de mappage couleur selon le statut
    def get_statut_color(statut):
        statut = str(statut).lower()
        if "non entamé" in statut:
            return "#D3D3D3"       # gris 
        elif "clôturé" in statut and "retard" not in statut:
            return "#90EE90"       # vert
        elif "bloqué" in statut:
            return "#FF0000"       # rouge
        elif "en cours" in statut:
            return "#FFA500"       # orange
        elif "retard" in statut:
            return "#FFFF00"       # jaune
        else:
            return "#87CEEB"       # bleu ciel par défaut

    # Appliquer la couleur à chaque ligne
    statut_counts["Couleur"] = statut_counts["Statut Avancement"].apply(get_statut_color)

    # Création du graphique Plotly avec couleurs personnalisées
    fig_statut = px.bar(
        statut_counts,
        x="Statut Avancement",
        y="Nombre",
        color="Statut Avancement",
        color_discrete_map={row["Statut Avancement"]: row["Couleur"] for _, row in statut_counts.iterrows()},
        title="Répartition par statut"
    )
    col1.plotly_chart(fig_statut, use_container_width=True)

    # Graphique circulaire des phases
    phase_counts = df_mission["Phases"].value_counts().reset_index()
    phase_counts.columns = ["Phases", "Nombre"]
    fig_phase = px.pie(phase_counts, names="Phases", values="Nombre", title="Répartition par phase")
    col2.plotly_chart(fig_phase, use_container_width=True)


    st.subheader("Répartition des statuts par phase en pourcentage")
    
    # Filtrer les données comme plus haut
    pivot = df_mission.pivot_table(index='Phases', columns='Statut Avancement', aggfunc='size', fill_value=0)
    
    # Calcul des pourcentages
    pivot_percent = pivot.div(pivot.sum(axis=1), axis=0) * 100
    pivot_percent = pivot_percent.reset_index().melt(id_vars='Phases', var_name='Statut Avancement', value_name='Pourcentage')
    
    # Fonction de couleur selon le statut
    def get_statut_color(statut):
        statut = str(statut).lower()
        if "non entamé" in statut:
            return "#D3D3D3"       # gris 
        elif "clôturé" in statut and "retard" not in statut:
            return "#90EE90"       # vert
        elif "bloqué" in statut:
            return "#FF0000"       # rouge
        elif "en cours" in statut:
            return "#FFA500"       # orange
        elif "retard" in statut:
            return "#FFFF00"       # jaune
        else:
            return "#87CEEB"       # bleu ciel par défaut
    
    # Générer la map de couleurs pour les statuts présents dans les données
    unique_statuts = pivot_percent["Statut Avancement"].unique()
    color_map = {statut: get_statut_color(statut) for statut in unique_statuts}
    
    # Création du graphique avec couleurs personnalisées
    fig = px.bar(
        pivot_percent,
        x="Phases",
        y="Pourcentage",
        color="Statut Avancement",
        color_discrete_map=color_map,
        title="Répartition en % des statuts par phase",
        text_auto='.1f',
    )
    
    # Mise en forme du graphique
    fig.update_layout(
        barmode='stack',
        xaxis_title="Phase",
        yaxis_title="Pourcentage (%)",
        yaxis=dict(ticksuffix="%")
    )
    
    st.plotly_chart(fig, use_container_width=True)


  
    fig_retard_inter = go.Figure(go.Bar(
        x=list(retard_intermediaire.values()),
        y=list(retard_intermediaire.keys()),
        orientation='h',
        marker_color=["#FFA500", "#FF0000", "#FFD700"],
        text=[f"{v} ({v/total:.1%})" for v in retard_intermediaire.values()],
        textposition="outside"
    ))

    fig_retard_inter.update_layout(
        title="Nombre de retards par étape intermédiaire",
        xaxis_title="Nombre de retards",
        yaxis_title="Étape",
        height=300,
        margin=dict(t=40, b=20)
    )

    st.plotly_chart(fig_retard_inter, use_container_width=True)

       


    # GANTT Chart
    st.subheader("📅 Diagramme de Gantt")
   
# Onglet 4 – Suivi des missions

with tabs[3]:
    st.markdown(
    """
    <style>
        section.main {
            background-color: #f0f0f0; /* gris clair */
        }
    </style>
    """,
    unsafe_allow_html=True
)

    st.subheader("Suivi des missions")

    # Réorganisation des colonnes
    colonnes_ord = ["Missions", "Services", "Porteurs", "Phases", "Activités",
                              "Livrables", "Date Début", "Date Elaboration Prévisionnelle", "Date Elaboration Effective",
                              "Responsable Elaboration","Satisfaction Elaboration",
                              "Date CTCQ Prévisionnelle", "Date CTCQ Effective","Responsable CTCQ","Satisfaction CTCQ", "Conformité",
                              "Date Approbation Prévisionnelle", "Date Approbation Effective","Responsable Approbation","Satisfaction Approbation",
                              "Date Finalisation Prévisionnelle", "Date Finalisation Effective","Satisfaction Globale","Date Satisfaction Client","Statut Avancement","Nom Client", "Commentaires","Date Facturation","Statut Règlement","Code Projet Client","Zone Géographique"]
    colonnes_sel = ["Missions", "Services", "Porteurs", "Phases", "Activités",
                              "Livrables", "Date Début", "Date Elaboration Prévisionnelle", "Date Elaboration Effective",
                              "Responsable Elaboration","Satisfaction Elaboration",
                              "Date CTCQ Prévisionnelle", "Date CTCQ Effective","Responsable CTCQ","Satisfaction CTCQ", "Conformité",
                              "Date Approbation Prévisionnelle", "Date Approbation Effective","Responsable Approbation","Satisfaction Approbation",
                              "Date Finalisation Prévisionnelle", "Date Finalisation Effective","Satisfaction Globale","Date Satisfaction Client","Statut Avancement","Nom Client", "Commentaires","Date Facturation","Statut Règlement","Code Projet Client","Zone Géographique"]
    
    df_obs = df_mission[colonnes_ord].copy()
    df_obs = df_obs[df_obs["Services"].isin(["Conformité ISO", "Formation"])]


    # Filtres
    col1, col2 = st.columns(2)

    with col1:
        missions = df_obs["Missions"].dropna().unique().tolist()
        missions.insert(0, "Toutes")
        selected_mission = st.radio("Choisir une mission", missions)

    with col2:
        if selected_mission != "Toutes":
            collaborateurs = df_obs[df_obs["Missions"] == selected_mission]["Porteurs"].dropna().unique().tolist()
        else:
            collaborateurs = df_obs["Porteurs"].dropna().unique().tolist()

        collaborateurs.insert(0, "Tous")
        selected_collab = st.radio("Choisir un collaborateur", collaborateurs)

    # Filtrage du DataFrame
    mission_data = df_obs.copy()
    if selected_mission != "Toutes":
        mission_data = mission_data[mission_data["Missions"] == selected_mission]
    if selected_collab != "Tous":
        mission_data = mission_data[mission_data["Porteurs"] == selected_collab]


    st.markdown("### 📊 Répartition des missions")
    
    # ⬜ Style CSS pour mini-sections avec fond blanc et ombre légère
    custom_box_style = """
    <style>
    .graph-box {
        background-color: white;
        padding: 10px;
        border-radius: 10px;
        box-shadow: 1px 1px 5px rgba(0,0,0,0.1);
        margin-bottom: 10px;
    }
    .graph-title {
        font-size: 14px;
        font-weight: bold;
        color: #003366;
        margin-bottom: 5px;
    }
    </style>
    """
    st.markdown(custom_box_style, unsafe_allow_html=True)
    
    # 📊 Graphiques côte à côte
    col3, col4 = st.columns(2)
    
    with col3:
        st.markdown('<div class="graph-box"><div class="graph-title">Répartition par statut</div>', unsafe_allow_html=True)
    
        statut_counts = mission_data["Statut Avancement"].value_counts().reset_index()
        statut_counts.columns = ["Statut Avancement", "Nombre"]
        statut_counts["Pourcentage"] = 100 * statut_counts["Nombre"] / statut_counts["Nombre"].sum()
    
        def get_statut_color(statut):
            statut = str(statut).lower()
            if "non entamé" in statut:
                return "#4F4F4F"
            elif "clôturé" in statut and "retard" not in statut:
                return "#90EE90"
            elif "bloqué" in statut:
                return "#FF0000"
            elif "en cours" in statut:
                return "#FFA500"
            elif "retard" in statut:
                return "#FFFF00"
            else:
                return "#87CEEB"
    
        statut_counts["Couleur"] = statut_counts["Statut Avancement"].apply(get_statut_color)
    
        fig1, ax1 = plt.subplots(figsize=(3, 2))
        bars = ax1.bar(
            statut_counts["Statut Avancement"],
            statut_counts["Nombre"],
            color=statut_counts["Couleur"]
        )
    
        for bar, pct in zip(bars, statut_counts["Pourcentage"]):
            ax1.text(
                bar.get_x() + bar.get_width() / 2,
                bar.get_height(),
                f"{pct:.0f}%",
                ha='center',
                va='bottom',
                fontsize=8
            )
    
        ax1.set_ylabel("")
        ax1.set_xlabel("")
        ax1.set_title("", fontsize=10)
        ax1.tick_params(axis='x', labelrotation=45, labelsize=8)
        ax1.spines[['right', 'top']].set_visible(False)
        st.pyplot(fig1)
    
        st.markdown('</div>', unsafe_allow_html=True)
    
    with col4:
        st.markdown('<div class="graph-box"><div class="graph-title">Répartition par phase</div>', unsafe_allow_html=True)
    
        phase_counts = mission_data["Phases"].value_counts()
    
        fig2, ax2 = plt.subplots(figsize=(3, 2))
        ax2.pie(
            phase_counts,
            labels=phase_counts.index,
            autopct='%1.1f%%',
            startangle=90,
            textprops={'fontsize': 8}
        )
        ax2.axis('equal')  # Cercle parfait
        st.pyplot(fig2)
    
        st.markdown('</div>', unsafe_allow_html=True)
    
        # 🔸 Affichage du tableau stylisé
            
        
        today = pd.Timestamp.today().normalize()
        
        def color_previsionnelle(row):
            styles = {}
            # Étapes à traiter
            etapes = ["Date Elaboration", "Date CTCQ", "Date Approbation"]
            for etape in etapes:
                prev_col = f"{etape} Prévisionnelle"
                eff_col = f"{etape} Effective"
                prev_date = row.get(prev_col)
                eff_date = row.get(eff_col)
        
                if pd.isna(prev_date):
                    styles[prev_col] = ''
                elif prev_date.date() == (today + timedelta(days=1)).date():
                    styles[prev_col] = 'background-color: orange; color: black'
                elif pd.notna(eff_date) and eff_date > prev_date:
                    styles[prev_col] = 'background-color: red; color: white'
                else:
                    styles[prev_col] = 'background-color: lightgreen; color: black'
            return pd.Series(styles)
        # 🔸 Coloration conditionnelle du statut
        def color_statut(val):
            val = str(val).lower()
            if "non entamé" in val:
                return 'background-color: #4F4F4F; color: white'  # gris foncé
            elif "clôturé" in val and "retard" not in val:
                return 'background-color: #90EE90; color: black'  # vert clair
            elif "bloqué" in val:
                return 'background-color: #FF0000; color: white'  # rouge
            elif "en cours" in val:
                return 'background-color: #FFA500; color: black'  # orange
            elif "retard" in val:
                return 'background-color: #FFFF00; color: black'  # jaune
            else:
                return ''
    
        styled_df = mission_data[colonnes_sel].style\
            .applymap(color_statut, subset=["Statut Avancement"])\
            .apply(color_previsionnelle, axis=1)
    
        #styled_df = mission_data[colonnes_sel].style.applymap(color_statut, subset=["Statut"])
            # 🔹 KPI : Retards intermédiaires (filtrés)
    
     # Recalcul des durées et retards sur le sous-ensemble filtré
    df_temp = df_mission.copy()
    df_temp["Date Début"] = pd.to_datetime(df_temp["Date Début"], errors='coerce')
    df_temp["Date Elaboration Prévisionnelle"] = pd.to_datetime(df_temp["Date Elaboration Prévisionnelle"], errors='coerce')
    df_temp["Date Elaboration Effective"] = pd.to_datetime(df_temp["Date Elaboration Effective"], errors='coerce')
    
    df_temp["Date CTCQ Prévisionnelle"] = pd.to_datetime(df_temp["Date CTCQ Prévisionnelle"], errors='coerce')
    df_temp["Date CTCQ Effective"] = pd.to_datetime(df_temp["Date CTCQ Effective"], errors='coerce')
    df_temp["Date Approbation Prévisionnelle"] = pd.to_datetime(df_temp["Date Approbation Prévisionnelle"], errors='coerce')
    df_temp["Date Approbation Effective"] = pd.to_datetime(df_temp["Date Approbation Effective"], errors='coerce')
    
    df_temp["Duree_Elaboration_Eff"] = (df_temp["Date Elaboration Effective"] - df_temp["Date Début"]).dt.days
    df_temp["Duree_Elaboration_Prév"] = (df_temp["Date Elaboration Prévisionnelle"] - df_temp["Date Début"]).dt.days
    
    df_temp["Duree_CTCQ_Eff"] = (df_temp["Date CTCQ Effective"] - df_temp["Date Elaboration Effective"]).dt.days
    df_temp["Duree_CTCQ_Prév"] = (df_temp["Date CTCQ Prévisionnelle"] - df_temp["Date Elaboration Effective"]).dt.days
    
    df_temp["Duree_Approbation_Eff"] = (df_temp["Date Approbation Effective"] - df_temp["Date CTCQ Effective"]).dt.days
    df_temp["Duree_Approbation_Prév"] = (df_temp["Date Approbation Prévisionnelle"] - df_temp["Date Approbation Effective"]).dt.days
    
    
    df_temp["Retard_Elaboration"] = df_temp["Duree_Elaboration_Eff"] > df_temp["Duree_Elaboration_Prév"]
    df_temp["Retard_CTCQ"] = df_temp["Duree_CTCQ_Eff"] > df_temp["Duree_CTCQ_Eff"]
    df_temp["Retard_Approbation"] = df_temp["Duree_Approbation_Eff"] > df_temp["Duree_Approbation_Prév"]
    
        # Application des mêmes filtres sur df_temp
    df_retards = df_temp.copy()
    if selected_mission != "Toutes":
        df_retards = df_retards[df_retards["Missions"] == selected_mission]
    if selected_collab != "Tous":
        df_retards = df_retards[df_retards["Porteurs"] == selected_collab]
    
        # Comptage
    nb_elab = df_retards["Retard_Elaboration"].sum()
    nb_ctcq = df_retards["Retard_CTCQ"].sum()
    nb_appro = df_retards["Retard_Approbation"].sum()
    
    total_missions = len(df_retards)
    
    st.markdown("### 📊 Indicateurs de retards intermédiaires")
    
        # Fonction pour formatage
    def format_pct(n, total):
        return f"{n / total:.0%}" if total else "0%"
        
        # CSS compact et discret
    compact_kpi_style = """
    <style>
    .kpi-label {
        background-color: #ffffff;
        border-left: 5px solid #003366;
        padding: 10px 15px;
        margin: 5px 0;
        font-size: 14px;
        color: #003366;
        font-weight: bold;
        box-shadow: 0 1px 2px rgba(0,0,0,0.05);
        border-radius: 4px;
    }
    .kpi-label .value {
        font-size: 18px;
        font-weight: bold;
        color: #003366;
        margin-left: 5px;
    }
    .kpi-label .pct {
        font-size: 14px;
        font-weight: normal;
        color: #0077b6;
        margin-left: 10px;
    }
    </style>
        """
    st.markdown(compact_kpi_style, unsafe_allow_html=True)
        
        # Affichage sur une ligne en 3 colonnes
    col1, col2, col3 = st.columns(3)
        
    with col1:
        st.markdown(f"""
        <div class="kpi-label">⏱️ Élaboration : 
            <span class="value">{nb_elab}</span> 
            <span class="pct">({format_pct(nb_elab, total_missions)})</span>
        </div>
        """, unsafe_allow_html=True)
        
    with col2:
        st.markdown(f"""
        <div class="kpi-label">📄 CT/CQ : 
                <span class="value">{nb_ctcq}</span> 
                <span class="pct">({format_pct(nb_ctcq, total_missions)})</span>
            </div>
            """, unsafe_allow_html=True)
        
        with col3:
            st.markdown(f"""
            <div class="kpi-label">✅ Approbation : 
                <span class="value">{nb_appro}</span> 
                <span class="pct">({format_pct(nb_appro, total_missions)})</span>
            </div>
            """, unsafe_allow_html=True)
    
    
    
    st.dataframe(styled_df, use_container_width=True)
    
with tabs[4]:
    # ⬇️ FORMULAIRE de saisie (dans le st.form)
    with st.form("ajout_mission_form", clear_on_submit=False):
        col1, col2 = st.columns(2)
        with col1:
            mission_id_mode = st.radio("🔗 Choix du mode", ["Créer une nouvelle mission", "Ajouter à une mission existante"])
            if mission_id_mode == "Ajouter à une mission existante" and missions_existantes:
                mission_id = st.selectbox("🆔 Sélectionner une mission existante", missions_existantes)
            else:
                mission_id = st.text_input("🆕 Créer un nouvel ID de mission")
    
            mission = st.selectbox("📂 Mission", ["CO", "GO", "Inspection", "Évaluation", "Autre"])
            service = st.selectbox("🏢 Services concernés", ["Formation", "Conformité ISO"])
            porteur = st.text_input("👤 Nom du porteur")
            phase = st.selectbox("📍 Phase", ["Préparation", "Déroulement", "Clôture"])
            activite = st.text_input("🧭 Activité")
            livrable = st.text_input("📄 Livrable attendu")
    
        with col2:
            date_debut = st.date_input("📅 Date de début")
            date_elab_prev = st.date_input("📅 Élaboration prévisionnelle")
            date_ctcq_prev = st.date_input("📅 CTCQ prévisionnelle")
            date_appro_prev = st.date_input("📅 Approbation prévisionnelle")
            date_fin_prev = st.date_input("📅 Fin prévisionnelle")
            responsable_elab = st.text_input("👤 Responsable Élaboration")
            responsable_ctcq = st.text_input("👤 Responsable CTCQ")
            responsable_appro = st.text_input("👤 Responsable Approbation")
            nom_clt = st.text_input("👤 Nom Client")
            zone_geo = st.text_input("Zone Géographique")
    
        commentaires = st.text_area("🗒️ Commentaires", "")
    
        # Prévisualisation uniquement
        submitted = st.form_submit_button("🔍 Prévisualiser")
    
    # ⬇️ Si on a cliqué sur Prévisualiser
    if submitted:
        if not mission_id.strip():
            st.error("❌ Veuillez renseigner un identifiant de mission.")
        else:
            unique_ref = f"AUTO-{pd.Timestamp.now().strftime('%Y%m%d%H%M%S%f')}"
            new_row = {
                "ID_Mission": mission_id.strip(),
                "Missions": mission,
                "Services": service,
                "Porteurs": porteur,
                "Phases": phase,
                "Activités": activite,
                "Livrables": livrable,
                "Date Début": pd.to_datetime(date_debut),
                "Date Elaboration Prévisionnelle": pd.to_datetime(date_elab_prev),
                "Date CTCQ Prévisionnelle": pd.to_datetime(date_ctcq_prev),
                "Date Approbation Prévisionnelle": pd.to_datetime(date_appro_prev),
                "Date Finalisation Prévisionnelle": pd.to_datetime(date_fin_prev),
                "Responsable Elaboration": responsable_elab,
                "Responsable CTCQ": responsable_ctcq,
                "Responsable Approbation": responsable_appro,
                "Nom Client": nom_clt,
                "Zone Géographique": zone_geo,
                "Commentaires": commentaires,
                "Ref": unique_ref
            }
    
            st.session_state["new_row_preview"] = new_row
            st.markdown("### 📋 Aperçu de la ligne à enregistrer")
            st.dataframe(pd.DataFrame([new_row]))
    
    # ⬇️ BOUTON HORS FORMULAIRE pour confirmer l'ajout
    path_excel = "dataset.xlsx"
    if "new_row_preview" in st.session_state:
        if st.button("✅ Enregistrer la mission"):
            try:
                df_exist = pd.read_excel(path_excel)
                df_new = pd.concat([df_exist, pd.DataFrame([st.session_state["new_row_preview"]])], ignore_index=True)
                df_new.to_excel(path_excel, index=False)
                st.success("✅ Mission ajoutée avec succès.")
                del st.session_state["new_row_preview"]
                st.rerun()
            except Exception as e:
                st.error(f"❌ Erreur lors de l'enregistrement : {e}")
   
