import streamlit as st
import pandas as pd

path = "rapport audit.xlsx"
sheet_name = "Referenciel audit"

referenciel = pd.read_excel(path, sheet_name, skiprows=7)


st.header('Annexe N° 02-Referenciel audit')

st.table(referenciel)
