import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from streamlit_option_menu import option_menu

path = "rapport audit.xlsx"
sheet_name="FO000"
book = load_workbook(path)
activeSheet = book[sheet_name]
df=pd.read_excel(path,sheet_name)

st.header(activeSheet["B1"].value)

selected = option_menu(
    menu_title=None,
    options=["Saisie des données","Afficahge des données" ],
    orientation="horizontal")

if option_menu == "Afficahge des données":
    st.table(df)
else:
    with st.form("observation_form",clear_on_submit=True):
        st.subheader("PARTIE RESERVEE AUX AUDITEURS")
        anomalie=st.text_area("Anomalie",activeSheet["C12"].value)
        cretere=st.text_area("Cretere",activeSheet["C13"].value)
        constat=st.text_area("Constat",activeSheet["C14"].value)
        cause=st.text_area("Causes",activeSheet["C15"].value)
        consequence=st.text_area("Consequences",activeSheet["C20"].value)
        recommandation=st.text_area("Recommandations",activeSheet["C24"].value)
        cause_option=st.radio("Cause options",options=["M1","M2","M3","M4","M5"],horizontal=True)
        consequence_option=st.radio("Consequence options",options=["R1","R2","R3","R4"],horizontal=True)
        recommandations=st.multiselect("Recomandations options",options=["S1","S2","S3","S4","S5","S6","S7","S8"])

        submited=st.form_submit_button("Sauvegarder")
    if submited:
        for row in range (15,32):
             activeSheet[f"G{row}"].value=""
        for row in range(15,20):
            if activeSheet[f"F{row}"].value == cause_option:
                activeSheet[f"G{row}"].value="X"
        for row in range(20,24):
            if activeSheet[f"F{row}"].value == consequence_option:
                activeSheet[f"G{row}"].value="X" 
        for row in range(24,32):
            for value in recommandations:
                if activeSheet[f"F{row}"].value ==value:
                    activeSheet[f"G{row}"].value="X"               
    book.save(path)
    with st.form("observation_auditee_form",clear_on_submit=True):
        st.subheader("PARTIE RESERVÉE AUX AUDITÉS")
        reponse=st.text_area("Reponse",activeSheet["C37"].value)
        action=st.text_area("Action",activeSheet["C41"].value)
        reponse_option=st.radio("Reponse options",options=["Acceptee","Retenue","Rejetee","A I'etude"],horizontal=True)
        action_option=st.radio("Action options",options=["Realisee","En cours","A realiser"],horizontal=True)

        submited=st.form_submit_button("Sauvegarder")
	
	
		
	
	

	
