import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import matplotlib.pyplot as plt
import seaborn as sns
import os
from openpyxl import load_workbook
from datetime import date
from datetime import timedelta

# Chargement des données
df_mission = pd.read_excel("dataset.xlsx", sheet_name="Sheet1")

# Nettoyage minimal
df_mission = df_mission.dropna(subset=["Statut"])
df_mission["Statut"] = df_mission["Statut"].astype(str).str.strip().str.lower()

# Parsing des dates
for col in ["Début", "Elaboration Prévisionnelle","Elaboration Effective", "Fin Prévisionnelle", "Fin Effective"]:
    df_mission[col] = pd.to_datetime(df_mission[col], errors='coerce')

# Ajout des jours de retard
if "Retard (jours)" not in df_mission.columns:
    df_mission["Retard (jours)"] = (
        pd.to_datetime(df_mission["Fin Effective"], errors="coerce") -
        pd.to_datetime(df_mission["Fin Prévisionnelle"], errors="coerce")
    ).dt.days
    df_mission["Retard (jours)"] = df_mission["Retard (jours)"].apply(lambda x: x if x and x > 0 else 0)

# Titre principal
st.set_page_config(page_title="Dashboard Suivi de Mission", layout="wide")
st.title("📊 Dashboard de Suivi des Missions – Clientisgroup")

# Tabs
tabs = st.tabs(["Vue d’ensemble", "Suivi Opérationnel", "Suivi des Missions"])

# Onglet 1 – Vue d'ensemble avec KPI
with tabs[0]:
    st.subheader("Vue d'ensemble des missions")
    
    # KPI
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("📌 Nombre de missions", df_mission["Type de Missions"].nunique())
    with col2:
        st.metric("📄 Nombre d'actions", df_mission["Etapes"].shape[0])
    with col3:
        st.metric("📦 Nombre de livrables", df_mission["Livrables"].nunique())

    # Ajouter la colonne "Commentaires" si elle n'existe pas
    if "Commentaires" not in df_mission.columns:
        df_mission["Commentaires"] = ""

    # S'assurer que "Ref" existe
    if "Ref" not in df_mission.columns:
        st.error("❌ La colonne 'Ref' est nécessaire dans df_mission pour l'enregistrement des modifications.")
    else:
        # Réorganisation des colonnes visibles (on cache Ref à l’affichage)
        colonnes_ordre_affichage = ["Missions", "Type de Missions", "Porteurs", "Phases", "Etapes",
                        "Livrables","Début", "Elaboration Prévisionnelle","Elaboration Effective", "CTCQ Prévisionnelle","CTCQ Effective", "Conformité", 
                                    "Approbation Prévisionnelle","Approbation Effective","Fin Prévisionnelle", "Fin Effective","Statut", "Commentaires"]
        
        colonnes_ordre_totale = colonnes_ordre_affichage + ["Ref"]  # Ref est gardé en interne
        df_vue = df_mission[colonnes_ordre_totale].copy()
        
        # Filtres
        st.write("### Filtres")
        col1, col2, col3 = st.columns(3)
        # Nettoyage des valeurs manquantes pour les menus déroulants
        type_missions = df_vue["Type de Missions"].fillna("(Inconnu)").unique().tolist()
        missions = df_vue["Missions"].fillna("(Inconnu)").unique().tolist()
        livrables = df_vue["Livrables"].fillna("(Inconnu)").unique().tolist()
        
        with col1:
            selected_typeMission = st.selectbox("Choisir un type de mission", ["Tous"] + sorted(type_missions))
        with col2:
            selected_mission = st.selectbox("Choisir une mission", ["Toutes"] + sorted(missions))
        with col3:
            selected_livrable = st.selectbox("Choisir un livrable", ["Tous"] + sorted(livrables))
        
        # Application des filtres avec prise en compte des "(Inconnu)"
        filtered_df = df_vue.copy()
        
        if selected_typeMission != "Tous":
            if selected_typeMission == "(Inconnu)":
                filtered_df = filtered_df[filtered_df["Type de Missions"].isna()]
            else:
                filtered_df = filtered_df[filtered_df["Type de Missions"] == selected_typeMission]
        
        if selected_mission != "Toutes":
            if selected_mission == "(Inconnu)":
                filtered_df = filtered_df[filtered_df["Missions"].isna()]
            else:
                filtered_df = filtered_df[filtered_df["Missions"] == selected_mission]
        
        if selected_livrable != "Tous":
            if selected_livrable == "(Inconnu)":
                filtered_df = filtered_df[filtered_df["Livrables"].isna()]
            else:
                filtered_df = filtered_df[filtered_df["Livrables"] == selected_livrable]


        # Application des filtres
        filtered_df = df_vue.copy()
        if selected_mission != "Toutes":
            filtered_df = filtered_df[filtered_df["Missions"] == selected_mission]
        if selected_typeMission != "Tous":
            filtered_df = filtered_df[filtered_df["Type de Missions"] == selected_typeMission]
        if selected_livrable != "Tous":
            filtered_df = filtered_df[filtered_df["Livrables"] == selected_livrable]

        # On masque "Ref" uniquement à l’affichage
        colonnes_affichees = [col for col in filtered_df.columns if col != "Ref"]

        

           # KPI calculés sur les données filtrées
    total_missions = len(filtered_df)
    missions_realisees = filtered_df["Fin Effective"].notna().sum()
    missions_conformes = filtered_df["Conformité"].eq("OUI").sum()
    missions_nonConformes = filtered_df["Conformité"].eq("NON").sum()
    missions_nonApplicables = filtered_df["Conformité"].eq("Non Applicable").sum()


    
    taux_realisation = missions_realisees / total_missions * 100 if total_missions > 0 else 0
    taux_conformite = missions_conformes / total_missions * 100 if total_missions > 0 else 0
    taux_nonConformite = missions_nonConformes / total_missions * 100 if total_missions > 0 else 0
    taux_nonApplicable = missions_nonApplicables / total_missions * 100 if total_missions > 0 else 0


    
    # Édition directe
    st.write("### Tableau de suivi des missions")
    
    edited_df = st.data_editor(
        filtered_df,
        use_container_width=True,
        num_rows="dynamic",
        column_config={
            "Missions": st.column_config.SelectboxColumn(
                "Missions",
                options=["CO", "GO", "Inspection", "Évaluation", "Autre"],
                required=False
            ),
            "Conformité": st.column_config.SelectboxColumn(
                "Conformité",
                options=["OUI", "NON", "Non Applicable"],
                required=False
            ),
            "Commentaires": st.column_config.TextColumn(
                "Commentaires"
            )
        }
    )
    
    # Chargement du fichier complet
    dataset_path = "dataset.xlsx"
    df_original = pd.read_excel(dataset_path)
    
    key_column = "Ref"
    
    if key_column in edited_df.columns and key_column in df_original.columns:
        # Nettoyage de l’index
        edited_df_clean = edited_df[edited_df["Ref"].notna()].copy()
        edited_df_clean.set_index("Ref", inplace=True)
        df_original.set_index("Ref", inplace=True)
    
        edited_df_clean = edited_df_clean[edited_df_clean.index.isin(df_original.index)]
    
        for column in edited_df_clean.columns:
            df_original.loc[edited_df_clean.index, column] = edited_df_clean[column]
    
        df_original.reset_index(inplace=True)
        edited_df.reset_index(inplace=True)
    
        try:
            df_original.to_excel(dataset_path, index=False)
            st.success("✅ Modifications insérées dans 'dataset.xlsx' sans perte des autres colonnes.")
        except PermissionError:
            st.error("❌ Fichier ouvert ailleurs. Fermez 'dataset.xlsx' puis réessayez.")
    else:
        st.error(f"❌ La colonne '{key_column}' doit exister dans les deux tables.")
    
        # ✅ KPI dynamiques selon le filtre
    st.markdown("### 📊 Indicateurs de performance")
    
    nb_total = len(filtered_df)
    nb_realisees = filtered_df["Fin Effective"].notna().sum()
    taux_action = (nb_realisees / nb_total) * 100 if nb_total > 0 else 0
    
    nb_conformes = filtered_df["Conformité"].str.upper().eq("OUI").sum()
    taux_conformite = (nb_conformes / nb_total) * 100 if nb_total > 0 else 0

    nb_nonconformes = filtered_df["Conformité"].str.upper().eq("NON").sum()
    taux_nonConformites = (nb_nonconformes / nb_total) * 100 if nb_total > 0 else 0

    nb_nonApplicables = filtered_df["Conformité"].str.upper().eq("NON APPLICABLE").sum()
    taux_nonApplicabilite = (nb_nonApplicables / nb_total) * 100 if nb_total > 0 else 0
    
    col_kpi1, col_kpi2, col_kpi3,col_kpi4 = st.columns(4)
    with col_kpi1:
        st.metric(label="✅ Actions réalisées", value=f"{nb_realisees}/{nb_total}", delta=f"{taux_action:.1f}%")
    with col_kpi2:
        st.metric(label="📋 Conformité (OUI)", value=f"{nb_conformes}/{nb_total}", delta=f"{taux_conformite:.1f}%")
    with col_kpi3:
        st.metric(label="📋 Conformité (NON)", value=f"{nb_nonconformes}/{nb_total}", delta=f"{taux_nonConformites:.1f}%")
    with col_kpi4:
        st.metric(label="📋 Conformité (Non Applicable)", value=f"{nb_nonApplicables}/{nb_total}", delta=f"{taux_nonApplicabilite:.1f}%")

   

# Onglet 2 – Visualisations
with tabs[1]:
    st.subheader("Suivi opérationnel")
         # Analyse des retards intermédiaires (élaboration → CTCQ → approbation)
    st.subheader("⏱️ Retards par étape intermédiaire")

    # Conversion des dates au cas où
    for col in ["Début", "Elaboration Prévisionnelle","Elaboration Effective", "CTCQ Prévisionnelle","CTCQ Effective", "Approbation Prévisionnelle","Approbation Effective"]:
        df_mission[col] = pd.to_datetime(df_mission[col], errors='coerce')

    # Calcul des durées
    df_mission["Duree_Elaboration"] = (df_mission["Elaboration Effective"] - df_mission["Début"]).dt.days
    df_mission["Duree_CTCQ"] = (df_mission["CTCQ Effective"] - df_mission["Elaboration Effective"]).dt.days
    df_mission["Duree_Approbation"] = (df_mission["Approbation Effective"] - df_mission["CTCQ Effective"]).dt.days

    # Comparaison aux seuils
    df_mission["Retard_Elaboration"] = df_mission["Elaboration Effective"] > df_mission["Elaboration Prévisionnelle"]
    df_mission["Retard_CTCQ"] = df_mission["CTCQ Effective"] > df_mission["CTCQ Prévisionnelle"]
    df_mission["Retard_Approbation"] = df_mission["Approbation Effective"] > df_mission["Approbation Prévisionnelle"]

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
    df_dates = df_mission[["Fin Prévisionnelle", "Fin Effective"]].dropna()
    total_realisees = len(df_dates)

    nb_sans_retard_global = (df_dates["Fin Effective"] <= df_dates["Fin Prévisionnelle"]).sum()
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
    statut_counts = df_mission["Statut"].value_counts().reset_index()
    statut_counts.columns = ["Statut", "Nombre"]
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
    statut_counts["Couleur"] = statut_counts["Statut"].apply(get_statut_color)

    # Création du graphique Plotly avec couleurs personnalisées
    fig_statut = px.bar(
        statut_counts,
        x="Statut",
        y="Nombre",
        color="Statut",
        color_discrete_map={row["Statut"]: row["Couleur"] for _, row in statut_counts.iterrows()},
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
    pivot = df_mission.pivot_table(index='Phases', columns='Statut', aggfunc='size', fill_value=0)
    
    # Calcul des pourcentages
    pivot_percent = pivot.div(pivot.sum(axis=1), axis=0) * 100
    pivot_percent = pivot_percent.reset_index().melt(id_vars='Phases', var_name='Statut', value_name='Pourcentage')
    
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
    unique_statuts = pivot_percent["Statut"].unique()
    color_map = {statut: get_statut_color(statut) for statut in unique_statuts}
    
    # Création du graphique avec couleurs personnalisées
    fig = px.bar(
        pivot_percent,
        x="Phases",
        y="Pourcentage",
        color="Statut",
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
   
# Onglet 3 – Suivi des missions

with tabs[2]:
    st.subheader("Suivi des missions")

    # Réorganisation des colonnes
    colonnes_ord = ["Missions", "Type de Missions", "Porteurs", "Phases", "Etapes",
                    "Livrables","Début","Elaboration Prévisionnelle","Elaboration Effective","CTCQ Prévisionnelle","CTCQ Effective","Conformité","Approbation Prévisionnelle","Approbation Effective","Fin Prévisionnelle", "Fin Effective","Statut", "Commentaires"]
    colonnes_sel = ["Missions", "Type de Missions", "Porteurs", "Phases", "Etapes",
                    "Livrables","Début","Elaboration Prévisionnelle","Elaboration Effective","CTCQ Prévisionnelle","CTCQ Effective","Conformité","Approbation Prévisionnelle","Approbation Effective","Fin Prévisionnelle", "Fin Effective","Statut", "Commentaires"]
    
    df_obs = df_mission[colonnes_ord].copy()

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

    # 🔹 Graphiques : Répartition par statut et par phase
 

# Container pour les deux graphes côte à côte
    col3, col4 = st.columns(2)
    
    with col3:
        # Répartition par statut
        statut_counts = mission_data["Statut"].value_counts().reset_index()
        statut_counts.columns = ["Statut", "Nombre"]
        statut_counts["Pourcentage"] = 100 * statut_counts["Nombre"] / statut_counts["Nombre"].sum()
    
        # Définir les couleurs en fonction du statut
        def get_statut_color(statut):
            statut = str(statut).lower()
            if "non entamé" in statut:
                return "#D3D3D3"
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
    
        statut_counts["Couleur"] = statut_counts["Statut"].apply(get_statut_color)
    
        # Bar chart compact
        fig1, ax1 = plt.subplots(figsize=(2, 1))  # Taille réduite
        bars = ax1.bar(
            statut_counts["Statut"],
            statut_counts["Nombre"],
            color=statut_counts["Couleur"]
        )
    
        # Ajout des pourcentages
        for bar, pct in zip(bars, statut_counts["Pourcentage"]):
            height = bar.get_height()
            ax1.text(
                bar.get_x() + bar.get_width() / 2,
                height,
                f"{pct:.0f}%",
                ha='center',
                va='bottom',
                fontsize=6
            )
    
        ax1.set_title("Répartition Statut", fontsize=5)
        ax1.set_ylabel("")
        ax1.set_xlabel("")
        ax1.tick_params(axis='x', labelrotation=45, labelsize=5)
        ax1.spines[['right', 'top']].set_visible(False)
        st.pyplot(fig1)
                                                                                                                                                                                                                                                                                    
    with col4:
        # Répartition circulaire par Phase
        phase_counts = mission_data["Phases"].value_counts()
        fig2, ax2 = plt.subplots(figsize=(2, 1))  # Taille réduite
        ax2.pie(
            phase_counts,
            labels=phase_counts.index,
            autopct='%1.1f%%',
            startangle=140,
            textprops={'fontsize': 4}
        )
        ax2.set_title("Répartition par phases", fontsize=5)
        st.pyplot(fig2)

    # 🔸 Coloration conditionnelle du statut
    def color_statut(val):
        val = str(val).lower()
        if "non entamé" in val:
            return 'background-color: #4F4F4F; color: white'  # gris foncé
        elif "clôturé" in val and "retard" not in val:
            return 'background-color: #D3D3D3; color: black'  # gris clair
        elif "bloqué" in val:
            return 'background-color: #FF0000; color: white'  # rouge
        elif "en cours" in val:
            return 'background-color: #FFA500; color: black'  # orange
        elif "retard" in val:
            return 'background-color: #FFFF00; color: black'  # jaune
        else:
            return ''

    # 🔸 Affichage du tableau stylisé
        
    
    today = pd.Timestamp.today().normalize()
    
    def color_previsionnelle(row):
        styles = {}
        # Étapes à traiter
        etapes = ["Elaboration", "CTCQ", "Approbation"]
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
    
    styled_df = mission_data[colonnes_sel].style\
        .applymap(color_statut, subset=["Statut"])\
        .apply(color_previsionnelle, axis=1)

    #styled_df = mission_data[colonnes_sel].style.applymap(color_statut, subset=["Statut"])
        # 🔹 KPI : Retards intermédiaires (filtrés)

    # Recalcul des durées et retards sur le sous-ensemble filtré
    df_temp = df_mission.copy()
    df_temp["Début"] = pd.to_datetime(df_temp["Début"], errors='coerce')
    df_temp["Elaboration Prévisionnelle"] = pd.to_datetime(df_temp["Elaboration Prévisionnelle"], errors='coerce')
    df_temp["Elaboration Effective"] = pd.to_datetime(df_temp["Elaboration Effective"], errors='coerce')

    df_temp["CTCQ Prévisionnelle"] = pd.to_datetime(df_temp["CTCQ Prévisionnelle"], errors='coerce')
    df_temp["CTCQ Effective"] = pd.to_datetime(df_temp["CTCQ Effective"], errors='coerce')
    df_temp["Approbation Prévisionnelle"] = pd.to_datetime(df_temp["Approbation Prévisionnelle"], errors='coerce')
    df_temp["Approbation Effective"] = pd.to_datetime(df_temp["Approbation Effective"], errors='coerce')

    df_temp["Duree_Elaboration_Eff"] = (df_temp["Elaboration Effective"] - df_temp["Début"]).dt.days
    df_temp["Duree_Elaboration_Prév"] = (df_temp["Elaboration Prévisionnelle"] - df_temp["Début"]).dt.days

    df_temp["Duree_CTCQ_Eff"] = (df_temp["CTCQ Effective"] - df_temp["Elaboration Effective"]).dt.days
    df_temp["Duree_CTCQ_Prév"] = (df_temp["CTCQ Prévisionnelle"] - df_temp["Elaboration Effective"]).dt.days

    df_temp["Duree_Approbation_Eff"] = (df_temp["Approbation Effective"] - df_temp["CTCQ Effective"]).dt.days
    df_temp["Duree_Approbation_Prév"] = (df_temp["Approbation Prévisionnelle"] - df_temp["Approbation Effective"]).dt.days


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

    # Affichage des KPI
    st.markdown("### 📊 Indicateurs de retards intermédiaires")
    col_kpi1, col_kpi2, col_kpi3 = st.columns(3)

    col_kpi1.metric("⏱️ Retards_Élaboration", f"{nb_elab}", f"{nb_elab/total_missions:.0%}" if total_missions else "0%")
    col_kpi2.metric("📄 Retards_CT/CQ", f"{nb_ctcq}", f"{nb_ctcq/total_missions:.0%}" if total_missions else "0%")
    col_kpi3.metric("✅ Retards_Approbation", f"{nb_appro}", f"{nb_appro/total_missions:.0%}" if total_missions else "0%")

    st.dataframe(styled_df, use_container_width=True)

  
    