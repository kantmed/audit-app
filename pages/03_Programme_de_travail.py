import streamlit as st
import pandas as pd

path = "rapport audit.xlsx"
sheet_name = "Programme de travail"

programmes = pd.read_excel(path, sheet_name, skiprows=7)


st.header('Annexe N° 03- Programme de travail')

st.table(programmes)