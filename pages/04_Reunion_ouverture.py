import streamlit as st
import pandas as pd
from streamlit_option_menu import option_menu
from openpyxl import load_workbook
from openpyxl.styles.borders import Border, Side
path = "rapport audit.xlsx"
sheet_name = "PV ouverture"

df_ouverture= pd.read_excel(path, sheet_name)


book = load_workbook(path)
workSheet = book.worksheets[6]
border = Border(left=Side(style='thin'), 
                     right=Side(style='thin'), 
                     top=Side(style='thin'), 
                     bottom=Side(style='thin'))
st.header("Annexe N° 04- PV de la reunion d'ouverture")
st.write(st.session_state)
selected = option_menu(
    menu_title=None,
    options=["Saisie des données", "Afficahge des données"],
    icons=["pencil-fill", "bar-chart-fill"],
    orientation="horizontal"
)
if selected == "Afficahge des données":
    st.table(df_ouverture)
else:    
    with st.form("participant_form",clear_on_submit=True):
        st.subheader("Prticipants a la reunion d'ouverture")
        participant = st.text_input("Participant")
        fonction = st.text_input("Fonction")
        submited = st.form_submit_button("Sauvegarder")
        for i in range(16,20):
            cellParticipant=workSheet["C{}".format(i)]
            cellFonction=workSheet["D{}".format(i)]
            if cellParticipant.value==None:
                cellParticipant.value=participant
                cellFonction.value=fonction
                break
        book.save(path)
    with st.form("presnetion_form",clear_on_submit=True):
        st.subheader("1-Présentation de la mission et ses objectifs ")
        presentation=st.text_area('Presentation',value=workSheet["A33"].value,height=150)
        submited = st.form_submit_button("Sauvegarder")
        if submited:
            workSheet["A33"].value=presentation
        book.save(path)
    with st.form("organisation_form",clear_on_submit=True):
        st.subheader("2-Organisation des aspects pratiques de son déroulement")
        organisation=st.text_area('Organisation',value=workSheet["A50"].value,height=150)
        submited = st.form_submit_button("Sauvegarder")
        if submited:
            workSheet["A50"].value=organisation
        book.save(path)
    with st.form("interlocuteur_form",clear_on_submit=True):
        st.subheader("3-Principaux interlocuteurs de la Mission")
        interlocuteur = st.text_input("Interlocuteur")
        fonction = st.text_input("Fonction")
        submited = st.form_submit_button("Sauvegarder")
        for i in range(36,42):
            cellIntelocuteur=workSheet["A{}".format(i)]
            cellFonction=workSheet["C{}".format(i)]
            if cellIntelocuteur.value==None:
                cellIntelocuteur.value=interlocuteur
                cellFonction.value=fonction
                break
        book.save(path)
    with st.form("attente_form",clear_on_submit=True):
        st.subheader("4-Attentes des audités")
        attentes=st.text_area('Attente',value=workSheet["A44"].value,height=150)
        submited = st.form_submit_button("Sauvegarder")
        if submited:
            workSheet["A44"].value=attentes
        book.save(path)
    with st.form("documents_form",clear_on_submit=True):
        st.subheader("5-Liste des documents à mettre à la disposition de la mission.")
        documents=st.text_area('Documents',value= workSheet["A47"].value,height=200)
        submited = st.form_submit_button("Sauvegarder")
        if submited:
            workSheet["A47"].value=documents
        book.save(path)
    with st.form("6-Divers_form",clear_on_submit=True):
        st.subheader("Divers")
        divers=st.text_area('Divers',value=workSheet["A50"].value,height=150)
        submited = st.form_submit_button("Sauvegarder")
        if submited:
            workSheet["A50"].value=divers
        book.save(path)


