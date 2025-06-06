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

# Chargement des données
# Chargement du fichier principal
# Titre principal

df_mission = pd.read_excel("dataset.xlsx", sheet_name="Sheet1")

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



# Tabs
tabs = st.tabs(["Vue d’ensemble", "Suivi Opérationnel", "Suivi des Missions", "➕ Ajouter une mission"])

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
        styled_kpi("📄 Nombre d'actions", df_mission["Etapes"].shape[0])
    with col3:
        styled_kpi("📦 Nombre de livrables", df_mission["Livrables"].nunique())

    # Ajout de colonnes manquantes
    if "Commentaires" not in df_mission.columns:
        df_mission["Commentaires"] = ""
    if "Ref" not in df_mission.columns:
        st.error("❌ La colonne 'Ref' est nécessaire.")
    else:
        # Réorganisation
        colonnes_affichage = ["ID_Mission","Missions", "Type de Missions", "Porteurs", "Phases", "Etapes",
                              "Livrables", "Début", "Elaboration Prévisionnelle", "Elaboration Effective",
                              "CTCQ Prévisionnelle", "CTCQ Effective", "Conformité",
                              "Approbation Prévisionnelle", "Approbation Effective",
                              "Fin Prévisionnelle", "Fin Effective", "Statut", "Commentaires"]
        df_vue = df_mission[colonnes_affichage + ["Ref"]].copy()

        # Filtres
        st.write("### Filtres")
        col1, col2, col3, col4 = st.columns(4)
        Ref_missions = df_vue["ID_Mission"].fillna("(Inconnu)").unique().tolist()
        type_missions = df_vue["Type de Missions"].fillna("(Inconnu)").unique().tolist()
        missions = df_vue["Missions"].fillna("(Inconnu)").unique().tolist()
        livrables = df_vue["Livrables"].fillna("(Inconnu)").unique().tolist()
        
        with col1:
            selected_RefMission = st.selectbox("Choisir un numéro mission", ["Tous"] + sorted(Ref_missions))
        with col2:
            selected_typeMission = st.selectbox("Choisir un type de mission", ["Tous"] + sorted(type_missions))
        with col3:
            selected_mission = st.selectbox("Choisir une mission", ["Toutes"] + sorted(missions))
        with col4:
            selected_livrable = st.selectbox("Choisir un livrable", ["Tous"] + sorted(livrables))

        # Application des filtres
        
        filtered_df = df_vue.copy()
        if selected_RefMission != "Tous":
            filtered_df = filtered_df[filtered_df["ID_Mission"].fillna("(Inconnu)") == selected_RefMission]
        if selected_typeMission != "Tous":
            filtered_df = filtered_df[filtered_df["Type de Missions"].fillna("(Inconnu)") == selected_typeMission]
        if selected_mission != "Toutes":
            filtered_df = filtered_df[filtered_df["Missions"].fillna("(Inconnu)") == selected_mission]
        if selected_livrable != "Tous":
            filtered_df = filtered_df[filtered_df["Livrables"].fillna("(Inconnu)") == selected_livrable]

        # KPI dynamiques
        st.markdown("### 📊 Indicateurs de performance")

        nb_total = len(filtered_df)
        nb_realisees = filtered_df["Fin Effective"].notna().sum()
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
        filtered_df["Approbation Effective"] = pd.to_datetime(filtered_df["Approbation Effective"], errors="coerce")
        edited_df = st.data_editor(
            filtered_df,
            use_container_width=True,
            num_rows="dynamic",
            column_config={
                "Missions": st.column_config.SelectboxColumn("Missions", options=["CO", "GO", "Inspection", "Évaluation", "Autre"]),
                "Conformité": st.column_config.SelectboxColumn("Conformité", options=["OUI", "NON", "Non Applicable"]),
                "Commentaires": st.column_config.TextColumn("Commentaires"),
                "Elaboraion Effective": st.column_config.DateColumn(label="Elaboraion Effective", format="YYYY-MM-DD"),
                "CTCQ Effective": st.column_config.DateColumn(label="CTCQ Effective", format="YYYY-MM-DD"),
                "Approbation Effective": st.column_config.DateColumn(label="Approbation Effective", format="YYYY-MM-DD"),
                "Fin Effective": st.column_config.DateColumn(label="Fin Effective", format="YYYY-MM-DD")
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
    
        statut_counts = mission_data["Statut"].value_counts().reset_index()
        statut_counts.columns = ["Statut", "Nombre"]
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
    
        statut_counts["Couleur"] = statut_counts["Statut"].apply(get_statut_color)
    
        fig1, ax1 = plt.subplots(figsize=(3, 2))
        bars = ax1.bar(
            statut_counts["Statut"],
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

with tabs[3]:
    st.subheader("📝 Formulaire d'ajout d'une nouvelle mission ou d'une phase")
    st.markdown(
        "Remplissez les informations ci-dessous pour créer une nouvelle mission ou ajouter une phase à une mission existante."
    )

    # Charger les missions existantes pour liste déroulante
    path_excel = "dataset.xlsx"
    try:
        df_exist = pd.read_excel(path_excel)
        missions_existantes = df_exist["ID_Mission"].dropna().unique().tolist()
    except Exception:
        df_exist = pd.DataFrame()
        missions_existantes = []

    with st.form("ajout_mission_form", clear_on_submit=False):
        col1, col2 = st.columns(2)

        with col1:
            # Mission_ID à saisir ou choisir
            mission_id_mode = st.radio("🔗 Choix du mode", ["Créer une nouvelle mission", "Ajouter à une mission existante"])
            if mission_id_mode == "Ajouter à une mission existante" and missions_existantes:
                mission_id = st.selectbox("🆔 Sélectionner une mission existante", missions_existantes)
            else:
                mission_id = st.text_input("🆕 Créer un nouvel ID de mission")

            type_mission = st.text_input("📌 Type de mission")
            mission = st.selectbox("📂 Mission", ["CO", "GO", "Inspection", "Évaluation", "Autre"])
            porteur = st.text_input("👤 Nom du porteur")
            phase = st.selectbox("📍 Phase", ["Préparation", "Déroulement", "Clôture"])
            etape = st.text_input("🧩 Étape")

        with col2:
            livrable = st.text_input("📄 Livrable attendu")
            date_debut = st.date_input("📅 Date de début")
            date_elab_prev = st.date_input("📅 Élaboration prévisionnelle")
            date_ctcq_prev = st.date_input("📅 CTCQ prévisionnelle")
            date_appro_prev = st.date_input("📅 Approbation prévisionnelle")
            date_fin_prev = st.date_input("📅 Fin prévisionnelle")
            #conformite = st.selectbox("✅ Conformité", ["OUI", "NON", "Non Applicable"])
            #statut = st.selectbox("📊 Statut", ["Non entamé", "En cours", "Bloqué", "Clôturé", "Clôturé avec retard"])

        commentaires = st.text_area("🗒️ Commentaires", "")

        submitted = st.form_submit_button("🔍 Prévisualiser")

    if submitted:
        if not mission_id.strip():
            st.error("❌ Veuillez renseigner un identifiant de mission (Mission_ID).")
        else:
            st.markdown("### 📋 Aperçu de la ligne à ajouter")
            import time
            unique_ref = f"AUTO-{pd.Timestamp.now().strftime('%Y%m%d%H%M%S%f')}"
            new_row = {
                "ID_Mission": mission_id.strip(),
                "Missions": mission,
                "Type de Missions": type_mission,
                "Porteurs": porteur,
                "Phases": phase,
                "Etapes": etape,
                "Livrables": livrable,
                "Début": pd.to_datetime(date_debut),
                "Elaboration Prévisionnelle": pd.to_datetime(date_elab_prev),
                "CTCQ Prévisionnelle": pd.to_datetime(date_ctcq_prev),
                "Approbation Prévisionnelle": pd.to_datetime(date_appro_prev),
                "Fin Prévisionnelle": pd.to_datetime(date_fin_prev),
               # "Conformité": conformite,
                #"Statut": statut,
                "Commentaires": commentaires,
                "Ref": unique_ref
            }

            st.dataframe(pd.DataFrame([new_row]))

            try:
                df_exist = pd.read_excel(path_excel)
                df_new = pd.concat([df_exist, pd.DataFrame([new_row])], ignore_index=True)
            
                # Vérifier que la ligne est bien ajoutée
                st.write("Nombre de lignes avant :", df_exist.shape[0])
                st.write("Nombre de lignes après :", df_new.shape[0])
            
                df_new.to_excel(path_excel, index=False)
                 # Recharge le fichier Excel après insertion pour que l'onglet 0 affiche la version à jour
                st.session_state["reload_df"] = True
                st.success("🎉 La mission a bien été ajoutée à la base de données.")
                st.rerun()
            except Exception as e:
                st.error(f"❌ Erreur lors de l'enregistrement : {e}")


  
    
