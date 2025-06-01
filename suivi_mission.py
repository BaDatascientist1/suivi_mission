import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import matplotlib.pyplot as plt
import seaborn as sns
import os
from openpyxl import load_workbook
from datetime import date

# Chargement des données
df_mission = pd.read_excel("dataset.xlsx", sheet_name="Sheet1")

# Nettoyage minimal
df_mission = df_mission.dropna(subset=["Statut"])
df_mission["Statut"] = df_mission["Statut"].astype(str).str.strip().str.lower()

# Parsing des dates
for col in ["Début", "Elaboration", "Fin Prévisionnelle", "Fin Effective"]:
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
    st.write("### Indicateurs clés")
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("📌 Nombre de missions", df_mission["Missions"].nunique())
    with col2:
        st.metric("📄 Nombre d'étapes", df_mission["Etapes"].shape[0])
    with col3:
        st.metric("📦 Nombre de livrables", df_mission["Liste des livrables"].nunique())
        
    # Ajouter la colonne "Commentaires" si elle n'existe pas
    if "Commentaires" not in df_mission.columns:
        df["Commentaires"] = ""

    # Réorganisation des colonnes
    colonnes_ordre = ["Ref","N°", "Missions", "Type de Missions", "Porteurs", "Phases", "Etapes",
                      "Liste des livrables", "Début", "Commentaires"]
    df_vue = df_mission[colonnes_ordre].copy()
    
    # Filtres
    st.write("### Filtres")

    col1, col2, col3 = st.columns(3)
    with col1:
        selected_mission = st.selectbox("Choisir une mission", ["Toutes"] + sorted(df_vue["Missions"].dropna().unique()))
    with col2:
        selected_typeMission = st.selectbox("Choisir un type de mission", ["Tous"] + sorted(df_vue["Type de Missions"].dropna().unique()))
    with col3:
        selected_livrable = st.selectbox("Choisir un livrable", ["Tous"] + sorted(df_vue["Liste des livrables"].dropna().unique()))

    # Application des filtres
    filtered_df = df_vue.copy()
    if selected_mission != "Toutes":
        filtered_df = filtered_df[filtered_df["Missions"] == selected_mission]
    if selected_typeMission != "Tous":
        filtered_df = filtered_df[filtered_df["Type de Missions"] == selected_typeMission]
    if selected_livrable != "Tous":
        filtered_df = filtered_df[filtered_df["Liste des livrables"] == selected_livrable]

    # Édition directe
    st.write("### Tableau de suivi des missions")

    edited_df = st.data_editor(
        filtered_df,
        use_container_width=True,
        num_rows="dynamic",
        column_config={
            "Type de Missions": st.column_config.SelectboxColumn(
                "Type de Missions",
                options=["CO", "GO", "Inspection", "Évaluation", "Autre"],
                required=False
            ),
            "Commentaires": st.column_config.TextColumn(
                "Commentaires"
            )
        }
    )
    # Chargement du fichier complet (avec toutes les colonnes)
    dataset_path = "dataset.xlsx"
    df_original = pd.read_excel(dataset_path)
    
    # `edited_df` est le DataFrame modifié dans Streamlit (avec seulement quelques colonnes affichées)
    # On suppose que vous avez une clé unique, comme "ID" ou "Nom de la mission", pour faire la jointure
    # Remplace "ID" par le bon nom de ta colonne identifiante
    key_column = "Ref"  # ou un autre identifiant unique
    
    # Vérifier que la clé existe dans les deux DataFrames
    if key_column in edited_df.columns and key_column in df_original.columns:
        # Mettre à jour les lignes correspondantes dans df_original avec les valeurs de edited_df
        df_original.set_index(key_column, inplace=True)
        edited_df.set_index(key_column, inplace=True)
    
        for column in edited_df.columns:
            df_original.loc[edited_df.index, column] = edited_df[column]
    
        df_original.reset_index(inplace=True)
        edited_df.reset_index(inplace=True)
    
        # Sauvegarde dans le même fichier Excel
        try:
            df_original.to_excel(dataset_path, index=False)
            st.success("✅ Modifications insérées dans 'dataset.xlsx' sans perte des autres colonnes.")
        except PermissionError:
            st.error("❌ Fichier ouvert ailleurs. Fermez 'dataset.xlsx' puis réessayez.")
    else:
        st.error(f"❌ La colonne '{key_column}' doit exister dans les deux tables.")

   
    


# Onglet 2 – Visualisations
with tabs[1]:
    st.subheader("Suivi opérationnel")

    col1, col2 = st.columns(2)
    
    statut_counts = df_mission["Statut"].value_counts().reset_index()
    statut_counts.columns = ["Statut", "count"]
    fig_statut = px.bar(statut_counts, x="Statut", y="count", color="Statut", title="Répartition par statut")
    col1.plotly_chart(fig_statut, use_container_width=True)

    phase_counts = df_mission["Phases"].value_counts().reset_index()
    phase_counts.columns = ["Phases", "count"]
    fig_phase = px.pie(phase_counts, names="Phases", values="count", title="Répartition par phase")
    col2.plotly_chart(fig_phase, use_container_width=True)

    st.subheader("Répartition des statuts par phase en pourcentage")

    # Filtrer les données comme plus haut
    pivot = df_mission.pivot_table(index='Phases', columns='Statut', aggfunc='size', fill_value=0)

    # Calcul des pourcentages
    pivot_percent = pivot.div(pivot.sum(axis=1), axis=0) * 100
    pivot_percent = pivot_percent.reset_index().melt(id_vars='Phases', var_name='Statut', value_name='Pourcentage')

    # Création du graphique
    fig = px.bar(
        pivot_percent,
        x="Phases",
        y="Pourcentage",
        color="Statut",
        title="Répartition en % des statuts par phase",
        text_auto='.1f',
    )

    fig.update_layout(barmode='stack', xaxis_title="Phase", yaxis_title="Pourcentage (%)", yaxis=dict(ticksuffix="%"))
    st.plotly_chart(fig, use_container_width=True)

       # Analyse des retards intermédiaires (élaboration → CTCQ → approbation)
    st.subheader("⏱️ Retards par étape intermédiaire")

    # Conversion des dates au cas où
    for col in ["Début", "Elaboration", "CTCQ", "Approbation"]:
        df_mission[col] = pd.to_datetime(df_mission[col], errors='coerce')

    # Calcul des durées
    df_mission["Duree_Elaboration"] = (df_mission["Elaboration"] - df_mission["Début"]).dt.days
    df_mission["Duree_CTCQ"] = (df_mission["CTCQ"] - df_mission["Elaboration"]).dt.days
    df_mission["Duree_Approbation"] = (df_mission["Approbation"] - df_mission["CTCQ"]).dt.days

    # Comparaison aux seuils
    df_mission["Retard_Elaboration"] = df_mission["Duree_Elaboration"] > df_mission["Seuil1"]
    df_mission["Retard_CTCQ"] = df_mission["Duree_CTCQ"] > df_mission["Seuil2"]
    df_mission["Retard_Approbation"] = df_mission["Duree_Approbation"] > df_mission["Seuil3"]

    # Comptage des retards
    retard_intermediaire = {
        "Élaboration": df_mission["Retard_Elaboration"].sum(),
        "CT/CQ": df_mission["Retard_CTCQ"].sum(),
        "Approbation": df_mission["Retard_Approbation"].sum()
    }

    total = len(df_mission)

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
            ✅<br><b>{pct_sans_retard_inter:.1f}%</b><br>Sans retard intermédiaire
        </div>
        """, unsafe_allow_html=True)

    with col2:
        st.markdown(f"""
        <div style="background-color:#dc3545;padding:10px;border-radius:10px;text-align:center;color:white;">
            ⚠️<br><b>{pct_avec_retard_inter:.1f}%</b><br>Avec retard intermédiaire
        </div>
        """, unsafe_allow_html=True)

    with col3:
        st.markdown(f"""
        <div style="background-color:#17a2b8;padding:10px;border-radius:10px;text-align:center;color:white;">
            ⏱️<br><b>{pct_sans_retard_global:.1f}%</b><br>Sans retard global
        </div>
        """, unsafe_allow_html=True)

    with col4:
        st.markdown(f"""
        <div style="background-color:#ffc107;padding:10px;border-radius:10px;text-align:center;color:black;">
            🕒<br><b>{pct_avec_retard_global:.1f}%</b><br>Avec retard global
        </div>
        """, unsafe_allow_html=True)


    # GANTT Chart
    st.subheader("📅 Diagramme de Gantt")
   
# Onglet 3 – Suivi des missions
with tabs[2]:
    st.subheader("Suivi des missions")

    # Réorganisation des colonnes
    colonnes_ord = ["N°", "Missions", "Type de Missions", "Porteurs", "Phases", "Etapes",
                    "Liste des livrables", "Statut", "Début", "Fin Prévisionnelle", "Fin Effective", "Commentaires"]
    colonnes_sel = ["Type de Missions", "Phases", "Etapes",
                    "Liste des livrables", "Statut", "Début", "Fin Prévisionnelle", "Fin Effective", "Commentaires"]
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
                return "#4F4F4F"
            elif "clôturé" in statut and "retard" not in statut:
                return "#D3D3D3"
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
    styled_df = mission_data[colonnes_sel].style.applymap(color_statut, subset=["Statut"])
        # 🔹 KPI : Retards intermédiaires (filtrés)

    # Recalcul des durées et retards sur le sous-ensemble filtré
    df_temp = df_mission.copy()
    df_temp["Début"] = pd.to_datetime(df_temp["Début"], errors='coerce')
    df_temp["Elaboration"] = pd.to_datetime(df_temp["Elaboration"], errors='coerce')
    df_temp["CTCQ"] = pd.to_datetime(df_temp["CTCQ"], errors='coerce')
    df_temp["Approbation"] = pd.to_datetime(df_temp["Approbation"], errors='coerce')

    df_temp["Duree_Elaboration"] = (df_temp["Elaboration"] - df_temp["Début"]).dt.days
    df_temp["Duree_CTCQ"] = (df_temp["CTCQ"] - df_temp["Elaboration"]).dt.days
    df_temp["Duree_Approbation"] = (df_temp["Approbation"] - df_temp["CTCQ"]).dt.days

    df_temp["Retard_Elaboration"] = df_temp["Duree_Elaboration"] > df_temp["Seuil1"]
    df_temp["Retard_CTCQ"] = df_temp["Duree_CTCQ"] > df_temp["Seuil2"]
    df_temp["Retard_Approbation"] = df_temp["Duree_Approbation"] > df_temp["Seuil3"]

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
