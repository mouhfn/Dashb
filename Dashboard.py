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

# Charger le fichier Excel et extraire la feuille souhaitée
excel_file = "SituationHFN.xlsx"
df = pd.read_excel(excel_file, sheet_name=sheet_name, header=None)

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
    ["Suivi de chargement", "Planification", "Stock","Navires en Rade","CTE"]
)

# Vue 1 : Suivi de chargement
# Vue 1 : Suivi de chargement
# Vue 1 : Suivi de chargement