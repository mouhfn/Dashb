import sys
import streamlit as st
import pandas as pd
import pprint
from datetime import datetime
from pathlib import Path
import base64
from openpyxl import load_workbook
import numpy as np
import openpyxl
import pandas as pd
import re
from datetime import datetime, time, timedelta



def normalize_product_name(name):
    if not name:
        return ""
    # Mise en majuscules et suppression des caractères spéciaux
    name = name.upper()
    
    # Supprime les espaces entre les lettres : "D A P" → "DAP"
    name = re.sub(r'\b([A-Z])\s+([A-Z])\b', r'\1\2', name)

    # Supprime tous les caractères non alphanumériques et remplace-les par un espace
    name = re.sub(r'[^A-Z0-9]', ' ', name)

    # Supprime les espaces multiples
    name = re.sub(r'\s+', ' ', name).strip()

    # Dictionnaire des règles de normalisation
    patterns = {
        'MAP 11 52 SPECIAL': [
            r'MAP\s*11\s*52',
            r'MAP\s*SPECIAL',
            r'MAP\s*11\s*52\s*SPECIAL',
            r'MAP\s*11\s*52\s*Special\s*Low\s*Cd',

        ],
        'NPK 14 18 18 6S 1B2O3': [
            r'NPK\s*14\s*18\s*18\s*6S\s*1B2O3\s*AFRIQUE',
            r'NPK\s*14\s*18\s*18\s*6S\s*1B2O3'],
         

        'DAP SPECIAL': [
            r'DAP\s*SPC',
            r'DAP\s*SPECIAL',
        ],
        'DAP EURO': [
            r'DAP\s*EURO',
            r'DAP\s*EU',
            r'DAP\s*Euro\s*Low\s*Cd'
            
        ],
        'DAP STANDARD': [
            r'DAP\s*STANDARD',
            r'DAP\s*STD',
        ],
        'TSP SPECIAL JORF': [
            r'TSP\s*JORF',
            r'TSP-JORF',
            r'TSP\s*SPECIAL\s*JORF',
            r'TSP\s*SPC\s*JARF',

        ],
        'NPS 3 30 9S': [
            r'NPS\s*3\s*30\s*9S\s*OFAS',
            r'NPS\s*3\s*30\s*9S',
        ],
        'UREE': [r'\bUREE\b', r'\bURÉE\b']
    }

    for standard, variants in patterns.items():
        for pattern in variants:
            if re.search(pattern, name):
                return standard

    return name# retourne le nom nettoyé mais non reconnu


# Configurer la mise en page
st.set_page_config(
    layout="wide",
    page_title="Suivi des chargements",
    page_icon="🌐",
)

# ✅ CSS : Barre verte + placement heure à droite
st.markdown("""
    <style>
        .top-bar {
            background-color: #00a65a; /* Vert */
            height: 4px;
            width: 100%;
            position: fixed;
            top: 0;
            left: 0;
            z-index: 100;
        }
            

        .custom-title {
            display: flex;
            justify-content: space-between;
            align-items: center;
            padding-top: 20px;
            padding-bottom: 10px;
            
        }

        .title-text {
            font-size: 1em;
            font-weight: bold;

        }

        .time-box {
            background-color: #85A98F;
            padding: 10px 20px;
            border-radius: 10px;
            font-size: 1em;
            color: #FFFFFF;
        }
    
        .block-container {
            padding-top: 30px !important;
        }
    </style>
    <div class="top-bar"></div>
""", unsafe_allow_html=True)

# ✅ Afficher Titre + Heure
now = datetime.now().strftime("%d/%m/%Y %H:%M:%S")

logo_path = "logo-white.png"  # Vérifie que le fichier est bien dans le même dossier

# Lecture du fichier et encodage en base64
logo_base64 = base64.b64encode(Path(logo_path).read_bytes()).decode()

st.markdown(f"""
    <div style="display: flex; align-items: center; justify-content: space-between; background-color: #328E6E; padding: 7px; border-radius: 8px;">
        <div style="flex: 1; display: flex; align-items: center;">
            <img src="data:image/png;base64,{logo_base64}" alt="logo" width="70"  style="margin-top : 10px; margin-right: 18px;"/>
        </div>
        <div style="flex: 2;color:#FFFFFF; text-align: center;margin-top : 10px; font-size: 24px; font-weight: bold;">
            🌐 Axe Chargement digital
        </div>
        <div style="flex: 1;color : White ; text-align: right; font-size: 18px;">
            🕒 {now}
        </div>
    </div>
""", unsafe_allow_html=True)

# Titre centré
# Charger les données depuis le fichier Excel

now = datetime.now()

# Si l'heure est avant 7h, utiliser la feuille du jour précédent
if now.time() < time(7, 0):
    effective_date = now - timedelta(days=1)
else:
    effective_date = now

# Formater la date au format "dd-mm-YYYY"
sheet_name = effective_date.strftime("%d-%m-%Y")

# Nouvelle version utilisant le lien en ligne
excel_url = "https://eocp-my.sharepoint.com/personal/mouad_elhafiani_ocpgroup_ma/_layouts/15/download.aspx?share=EW83yQHE4YdElMWYtV52Ii0BNrLL5-3KGGclpOgKaRA7UA"

# Bouton pour rafraîchir les données (relance l'application)
if st.button("🔄 Rafraîchir les données"):
    st.rerun()

# Charger le fichier Excel via le lien en ligne
xls = pd.ExcelFile(excel_url)
if sheet_name not in xls.sheet_names:
    sheet_name = xls.sheet_names[0]  # ou une autre feuille par défaut
df = pd.read_excel(xls, sheet_name=sheet_name, header=None)

# Extraire les données des quais
loading_data = {}

# Parcourir les colonnes de B à AD (index 1 à 30)
for col in range(1, 31):
    quai = df.iloc[6, col]  # Ligne 7 = index 6
    if pd.isna(quai):
        continue

    ship = df.iloc[1, col]             # Ligne 2 = index 1
    quantity_requested = df.iloc[4, col]  # Ligne 5 = index 4
    product_type = df.iloc[5, col]     # Ligne 6 = index 5
    origin = df.iloc[7, col]  
    total_charge = df.iloc[37, 0]  # Ligne 30, colonne B (0-indexed)

    # Cumul de chargement : lignes 13 à 36 (index 12 à 35)
    cumul_data = df.iloc[12:36, col].dropna()

    max_loaded = 0
    tonnage_last_hour = 0

    if not cumul_data.empty:
        max_index = cumul_data.idxmax()  # Ligne où le cumul est max
        max_loaded = cumul_data[max_index]

        # Vérifier que la colonne col+1 existe
        if col + 1 < df.shape[1]:
            tonnage_last_hour = df.iloc[max_index, col + 1]
    
    # Construction de la structure
    if quai not in loading_data:
        loading_data[quai] = {
            "ship": ship,
            "products": {}
        }

    loading_data[quai]["products"][product_type]= {
        "loaded": max_loaded,
        "target": quantity_requested,
        "last_hour": tonnage_last_hour,
        "Source" : origin,
        
    }
print(loading_data)
# Afficher les données dans le tableau de bord
page = st.sidebar.selectbox(
    "Navigation",
    ["Suivi de chargement","CTE"]
)

# Vue 1 : Suivi de chargement
# Vue 1 : Suivi de chargement
# Vue 1 : Suivi de chargement
if page == "Suivi de chargement":
    st.markdown(
    f"""
    <div style='display:flex; justify-content:space-between; align-items:center;'>
        <h3 style="font-size:24px; margin:0;">🚢 Suivi de chargement en temps réel</h3>
        <span style="font-size:20px;color:#FFFFFF; background-color:#F7374F; padding:6px 12px; border-radius:8px;">
            Total Chargé : <strong>{total_charge} t</strong>
        </span>
    </div>
    """,
    unsafe_allow_html=True
)



    # Diviser la page en deux colonnes : Quai 1 (gauche) et Quai 2 (droite)
 # Diviser la page en trois colonnes : col1, separation (sep_col) et col2
cols = st.columns([1, 0.05, 1])

# Affichage pour la colonne 1 (Quais 1)
with cols[0]:
    st.markdown('<h4 style="margin-bottom:10px;">🧭 Quais - 1</h4>', unsafe_allow_html=True)
    for quai, info in loading_data.items():
        if quai.startswith("1"):  # Filtrer les quais de gauche
            st.markdown("<hr style='border-top:3px solid #000; margin:15px 0;'>", unsafe_allow_html=True)
            if not info["ship"] or pd.isna(info["ship"]):
                st.markdown(
                    f'<div style="font-size:20px; font-weight:bold;">Quai {quai} – 🚩 Quai Libre</div>',
                    unsafe_allow_html=True,
                )
            else:
                st.markdown(
                    f'<div style="font-size:20px; font-weight:bold;">Quai {quai} – Navire : {info["ship"]}</div>',
                    unsafe_allow_html=True,
                )
                for product, stats in info["products"].items():
                    progress = stats["loaded"] / stats["target"] if stats["target"] > 0 else 0
                    percentage = progress * 100

                    # Display product name with green color
                    product_html = f'<span style="color:green;">Qualité 🌾 : {product}</span>'

                    # Set color for the source cube based on the source name
                    source = stats["Source"]
                    if source == "HE2":
                        cube_color = "#FF5722"  # Orange, for example
                    elif source == "JFC1":
                        cube_color = "#FFEB3B"  # Yellow
                    elif source == "JFC2":
                        cube_color = "#B0BEC5"  # Gray
                    elif source == "JLN":
                        cube_color = "#03A9F4"  # Blue
                    else:
                        cube_color = "#4CAF50"  # Default green

                    # Create the colored cube for the source
                    source_html = f'<span style="background-color:{cube_color}; display:inline-block; width:20px; height:20px; margin-right:5px;"></span>{source}'

                    st.markdown(
                        f'<div style="font-size:18px; font-weight:bold;">🔸 {product_html} | Chargé ✅ : {stats["loaded"]} / {stats["target"]} t | Dernière heure 🕒 : {stats["last_hour"]} t </div>',
                        unsafe_allow_html=True,
                    )
                    st.progress(progress)
                    st.markdown(
                        f'<div style="font-size:18px; font-weight:bold; text-align:center;">{percentage:.2f}% Chargé | Source de chargement actuel 🏗️ : {source_html}</div>',
                        unsafe_allow_html=True,
                    )

# Colonne de séparation
with cols[1]:
    st.markdown("<div style='border-left: 3px solid blue; height:100vh;'></div>", unsafe_allow_html=True)

# Affichage pour la colonne 2 (Quais 2)
with cols[2]:
    st.markdown('<h4 style="margin-bottom:10px;">🧭 Quais - 2</h4>', unsafe_allow_html=True)
    for quai, info in loading_data.items():
        if quai.startswith("2"):  # Filtrer les quais pour la colonne de droite
            st.markdown("<hr style='border-top:3px solid #000; margin:15px 0;'>", unsafe_allow_html=True)
            if not info["ship"] or pd.isna(info["ship"]):
                st.markdown(
                    f'<div style="font-size:18px; font-weight:bold;">Quai {quai} – 🚩 Quai Libre</div>',
                    unsafe_allow_html=True,
                )
            else:
                st.markdown(
                    f'<div style="font-size:20px; font-weight:bold;">Quai {quai} – Navire : {info["ship"]}</div>',
                    unsafe_allow_html=True,
                )
                for product, stats in info["products"].items():
                    progress = stats["loaded"] / stats["target"] if stats["target"] > 0 else 0
                    percentage = progress * 100

                    # Display product name with green color
                    product_html = f'<span style="color:green;">Qualité 🌾 : {product}</span>'

                    # Set color for the source cube based on the source name
                    source = stats["Source"]
                    if source == "HE2":
                        cube_color = "#FF5722"  # Orange, for example
                    elif source == "JFC1":
                        cube_color = "#FFEB3B"  # Yellow
                    elif source == "JFC2":
                        cube_color = "#B0BEC5"  # Gray
                    elif source == "JLN":
                        cube_color = "#03A9F4"  # Blue
                    else:
                        cube_color = "#4CAF50"  # Default green

                    # Create the colored cube for the source
                    source_html = f'<span style="background-color:{cube_color}; display:inline-block; width:20px; height:20px; margin-right:5px;"></span>{source}'

                    st.markdown(
                        f'<div style="font-size:18px; font-weight:bold;">🔸 {product_html} | Chargé ✅ : {stats["loaded"]} / {stats["target"]} t | Dernière heure 🕒 : {stats["last_hour"]} t </div>',
                        unsafe_allow_html=True,
                    )
                    st.progress(progress)
                    st.markdown(
                        f'<div style="font-size:18px; font-weight:bold; text-align:center;">{percentage:.2f}% Chargé | Source de chargement actuel 🏗️ : {source_html}</div>',
                        unsafe_allow_html=True,
                    )
