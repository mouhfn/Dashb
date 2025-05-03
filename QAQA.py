import sys
import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import datetime
import re

# --------------------- UI CONFIGURATION ---------------------
st.set_page_config(page_title="Dashboard: Planifié vs Chargé", layout="wide")
st.markdown("<h1 style='text-align: center;'>📊 Dashboard Chargement: Planifié vs Réel</h1>", unsafe_allow_html=True)

# --------------------- FILE UPLOAD ---------------------
uploaded_file = st.sidebar.file_uploader("📥 Charger le fichier Excel", type=["xlsx"])
if not uploaded_file:
    st.sidebar.warning("⚠️ Veuillez uploader le fichier Excel.")
    st.stop()

df = pd.read_excel(uploaded_file, header=None)

# --------------------- DATA EXTRACTION ---------------------
columns = list(df.loc[7, 5:30].index)  # Ligne 8 (index 7)
start_row = 8
col_e = 4  # Column E is index 4

# Determine end_row: iterate from the start_row until an empty cell is found in column E
for row in range(start_row, len(df)):
    if pd.isna(df.iat[row, col_e]):
        end_row = row    # Stop here if cell in column E is empty
        break
else:
    end_row = len(df) 
hours = []
hourly_planned = {}
cumulative_planned = []
cum_plan = 0

for col_idx in range(6, 31):
    total_hour = 0
    for row in range(start_row, end_row):
        current = pd.to_numeric(df.iat[row, col_idx], errors='coerce')
        previous = pd.to_numeric(df.iat[row, col_idx - 1], errors='coerce')
        if pd.notna(current) and pd.notna(previous) and current > 0 and previous >= 0:
            delta = current - previous
            if delta > 0:
                total_hour += delta
    hour_label = df.iat[7, col_idx]
    hourly_planned[hour_label] = total_hour
    cum_plan += total_hour
    cumulative_planned.append(cum_plan)
    hours.append(hour_label)

# --------------------- REAL INPUTS ---------------------
st.sidebar.header("✍️ Saisie réelle")
hourly_real = {}
real_cumulative = []

for i, hour in enumerate(hours):
    val = st.sidebar.number_input(f"{hour}", min_value=0.0, value=0.0, step=0.1)
    if val == 0.0 and i > 0:
        val = hourly_real[hours[i - 1]]
    hourly_real[hour] = val
    real_cumulative.append(val if i == 0 else val)

# --------------------- DIFFERENCE ---------------------
hourly_diff = {hour: hourly_real[hour] - hourly_planned[hour] for hour in hours}

# --------------------- PLOT ---------------------
fig = go.Figure()
fig.add_trace(go.Scatter(x=hours, y=cumulative_planned, mode="lines+markers",
                         name="Planifié", line=dict(color="royalblue", width=3)))
fig.add_trace(go.Scatter(x=hours, y=real_cumulative, mode="lines+markers",
                         name="Réel", line=dict(color="green", width=3)))
fig.update_layout(title="📈 Charge Cumulée Planifiée vs Réelle",
                  xaxis_title="Heure", yaxis_title="Tonnes",
                  legend=dict(orientation="h", y=-0.2))

# --------------------- CURRENT HOUR INFO ---------------------
now = datetime.datetime.now()
current_hour = now.hour
matching_index = next((i for i, h in enumerate(hours)
                       if re.search(r'\d+', h) and int(re.search(r'\d+', h).group()) == current_hour), -1)
if matching_index == -1:
    matching_index = len(hours) - 1

last_real = hourly_real[hours[matching_index]]
last_plan = cumulative_planned[matching_index]
last_gap = last_real - last_plan

# --------------------- DISPLAY ---------------------
col1, col2 = st.columns([3, 1])

with col1:
    st.plotly_chart(fig, use_container_width=True)

with col2:
    def info_box(title, value, color):
        return f"""
        <div style="
            background-color: {color};
            padding: 15px;
            border-radius: 10px;
            margin: 10px 0;
            color: white;
            font-size: 20px;
            font-weight: bold;
            text-align: center;">
            {title}<br>{value:.2f} T
        </div>
        """
    st.markdown(info_box("🔄 Chargé", last_real, "#2ecc71"), unsafe_allow_html=True)
    st.markdown(info_box("📋 Planifié", last_plan, "#2980b9"), unsafe_allow_html=True)
    st.markdown(info_box("⚖️ Écart", last_gap, "#e74c3c" if last_gap < 0 else "#27ae60"), unsafe_allow_html=True)

# --------------------- OBJECTIFS ---------------------
def get_idx_for_hour(h_target):
    for i, h in enumerate(hours):
        try:
            if int(re.sub("[^0-9]", "", h)) == h_target:
                return i
        except:
            continue
    return len(hours) - 1

idx14 = get_idx_for_hour(14)
idx22 = get_idx_for_hour(22)
idx07 = get_idx_for_hour(7)

obj14 = cumulative_planned[idx14]
obj22 =(cumulative_planned[idx22] - cumulative_planned[idx14])
obj07 = cumulative_planned[idx07] - cumulative_planned[idx22]
st.markdown("<hr>", unsafe_allow_html=True)
st.subheader("🎯 Objectifs de la journée")
col_o1, col_o2, col_o3 = st.columns(3)
with col_o1:
    st.metric("1er Objectif (14h)", f"{obj14:.2f} T")
with col_o2:
    st.metric("2ème Objectif (22h)", f"{obj22:.2f} T")
with col_o3:
    st.metric("3ème Objectif (7h)", f"{obj07:.2f} T")

# --------------------- FINAL TABLE ---------------------
# Construct a summary table with only rows until the last hour
# ...existing code...

# Compute differential real values:
reel_values = [hourly_real[hours[0]]] + [
    hourly_real[hours[i]] - hourly_real[hours[i-1]] for i in range(1, matching_index+1)
]

# Construct the final table using the computed reel_values
df_table = pd.DataFrame({
    "Heure": hours[:matching_index+1],
    "Planifié (T)": [hourly_planned[h] for h in hours[:matching_index+1]],
    "Réel (T)": reel_values,
    "Écart (T)": [-(hourly_planned[h] - reel) for h, reel in zip(hours[:matching_index+1], reel_values)]
})
st.markdown("---")
st.subheader("📋 Détail horaire (jusqu'à la dernière heure)")
st.dataframe(
    df_table.style.format({
        "Planifié (T)": "{:.2f}",
        "Réel (T)": "{:.2f}",
        "Écart (T)": "{:.2f}"
    }),
    use_container_width=True
)