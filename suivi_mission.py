import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import date

# Chargement des données
df_mission = pd.read_excel("dataset.xlsx", sheet_name="Maquette mission")

# Nettoyage minimal
df_mission = df_mission.dropna(subset=["Statut"])

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
st.title("📊 Dashboard de Suivi des Missions – Clientis Group")

# Tabs
tabs = st.tabs(["Vue d’ensemble", "Visualisations", "Suivi Collaborateurs", "Étapes longues"])

# Onglet 1 – Vue d'ensemble avec KPI
with tabs[0]:
    st.subheader("Vue d'ensemble des missions")

    col1, col2, col3, col4 = st.columns(4)

    total_missions = df_mission["Missions"].nunique()
    total_etapes = df_mission.shape[0]
    etapes_en_cours = df_mission[df_mission["Statut"] == "En cours"].shape[0]
    etapes_cloturees = df_mission[df_mission["Statut"].str.contains("Clôturé", na=False)].shape[0]

    # Étapes à traiter aujourd'hui
    etapes_today = df_mission[df_mission["Focus sur les actions du jour"] > 0]
    nb_etapes_a_traiter = etapes_today.shape[0]

    with col1:
        st.metric("Nombre de missions", total_missions)
    with col2:
        st.metric("Étapes totales", total_etapes)
    with col3:
        st.metric("À traiter aujourd'hui", nb_etapes_a_traiter)
    with col4:
        st.metric("Étapes clôturées", etapes_cloturees)

    # Filtres stylisés
    st.markdown("<style>.stMultiSelect>div>div{border-radius:10px !important; background-color:#f0f2f6;}</style>", unsafe_allow_html=True)
    st.markdown("---")
    phase_filter = st.multiselect("Filtrer par phase :", df_mission["Phases"].dropna().unique())
    statut_filter = st.multiselect("Filtrer par statut :", df_mission["Statut"].dropna().unique())

    filtered_df = df_mission.copy()
    if statut_filter:
        filtered_df = filtered_df[filtered_df["Statut"].isin(statut_filter)]
    if phase_filter:
        filtered_df = filtered_df[filtered_df["Phases"].isin(phase_filter)]

    st.dataframe(filtered_df[["Missions", "Phases", "Etapes", "Statut", "Porteurs", "Fin Prévisionnelle", "Fin Effective"]])

# Onglet 2 – Visualisations
with tabs[1]:
    st.subheader("Visualisations globales")

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

    # Analyse des retards par phase
    st.subheader("📌 Où perd-on le plus de temps ?")
    retard_par_phase = df_mission.groupby("Phases")["Retard (jours)"].sum().reset_index()
    fig_retard = px.bar(retard_par_phase.sort_values("Retard (jours)", ascending=False),
                        x="Phases", y="Retard (jours)", color="Phases",
                        title="Total des jours de retard par phase")
    st.plotly_chart(fig_retard, use_container_width=True)

    # GANTT Chart
    st.subheader("📅 Diagramme de Gantt")
    gantt_df = df_mission.dropna(subset=["Début", "Fin Prévisionnelle"])
    fig_gantt = px.timeline(gantt_df, x_start="Début", x_end="Fin Prévisionnelle", y="Etapes", color="Phases")
    fig_gantt.update_yaxes(autorange="reversed")
    st.plotly_chart(fig_gantt, use_container_width=True)

# Onglet 3 – Suivi collaborateurs
with tabs[2]:
    st.subheader("Suivi des collaborateurs")

    collaborateurs = df_mission["Porteurs"].dropna().unique()
    selected_collab = st.selectbox("Choisir un collaborateur", collaborateurs)

    collab_data = df_mission[df_mission["Porteurs"] == selected_collab]

    col1, col2, col3, col4 = st.columns(4)

    with col1:
        st.metric("🧮 Total d'étapes", collab_data.shape[0])
    with col2:
        st.metric("✅ Étapes clôturées", collab_data[collab_data["Statut"].str.contains("Clôturé", na=False)].shape[0])
    with col3:
        st.metric("🚧 Étapes en cours", collab_data[collab_data["Statut"] == "En cours"].shape[0])
    with col4:
        st.metric("⏰ Étapes en retard", collab_data[collab_data["Statut"].str.contains("retard", na=False)].shape[0])

    st.dataframe(collab_data[["Missions", "Phases", "Etapes", "Statut", "Fin Prévisionnelle", "Fin Effective"]])

    st.subheader("Visualisations spécifiques au collaborateur sélectionné")

    col1, col2 = st.columns(2)

    with col1:
        statut_counts = collab_data["Statut"].value_counts().reset_index()
        statut_counts.columns = ["Statut", "count"]
        fig_statut = px.bar(statut_counts, x="Statut", y="count", color="Statut", title="Répartition par statut")
        st.plotly_chart(fig_statut, use_container_width=True, key="statut_plot")

    with col2:
        phase_counts = collab_data["Phases"].value_counts().reset_index()
        phase_counts.columns = ["Phases", "count"]
        fig_phase = px.pie(phase_counts, names="Phases", values="count", title="Répartition par phase")
        st.plotly_chart(fig_phase, use_container_width=True, key="phase_plot")

# Onglet 4 – Étapes longues
with tabs[3]:
    st.subheader("Étapes les plus longues")

    df_temp = df_mission.copy()

    for col in ["Début", "Elaboration", "Fin Effective"]:
        df_temp[col] = pd.to_datetime(df_temp[col], errors="coerce")

    df_temp["Date Départ"] = df_temp["Elaboration"].combine_first(df_temp["Début"])
    df_temp["Duree (jours)"] = (df_temp["Fin Effective"] - df_temp["Date Départ"]).dt.days
    df_temp = df_temp.dropna(subset=["Duree (jours)"])

    plus_lentes = df_temp.sort_values("Duree (jours)", ascending=False).head(10)
    st.write("Top 10 des étapes ayant duré le plus longtemps")
    st.dataframe(plus_lentes[["Missions", "Etapes", "Porteurs", "Duree (jours)", "Statut"]])
