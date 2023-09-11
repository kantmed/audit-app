import streamlit as st
import pandas as pd
from streamlit_option_menu import option_menu
from openpyxl import load_workbook

path = "rapport audit.xlsx"
sheet_name = "Parametres"

book = load_workbook(path)


parametres_sheet = book.worksheets[0]

parametres = pd.read_excel(path, sheet_name)
auditeurs = pd.read_excel(path, sheet_name='Auditeurs')
structures = pd.read_excel(path, sheet_name='Structures')


def get_index(auditeur):
    auditeurs['auditeur'].loc[lambda x: x == auditeur].index[0]

def authenticate():
    return True


def main():
    st.set_page_config(
        page_title="IRA Ghardaia 174",
        page_icon="🧊",
        layout="wide",
    )
    st.header('IRA Ghardaia 174')
    st.subheader("Rapport mission d'audit")
    selected = option_menu(
        menu_title=None,
        options=["Saisie des données", "Afficahge des données"],
        orientation="horizontal"
    )

    if selected == "Afficahge des données":
        st.table(parametres)

    else:
        with st.form("parametre_form"):
            left, rigth = st.columns(2)
            with left:
                ira = st.text_input('IRA', parametres_sheet["B2"].value)
                date_depart = st.date_input(
                    'Date depart', parametres_sheet['B5'].value)
                structure = st.selectbox('Structure auditée', options=structures)
                numero_ordre = st.text_input(
                    'N° Ordre', parametres_sheet['B9'].value)
                auditeur = st.selectbox('Auditeur', options=auditeurs, index=5)
                submited = st.form_submit_button('Sauvegarder')
            with rigth:
                objet = st.selectbox(
                    'Objet de la mission',
                    ("Mission d'audit evaluation", "Mission d'inspection", "Mission d'audit retour"))
                date_retour = st.date_input(
                    'Date retour', parametres_sheet['B6'].value)
                code_theme = st.selectbox(
                    'Choisir',
                    ("Unite", "Processus", "Structure", "Risque"))
                chef_mission = st.selectbox(
                    'Chef de mission', options=auditeurs, index=4)

                if submited:
                    typeAudit = ''
                    codeTheme = ''
                    month = ''
                    if objet == "Mission d'audit evaluation":
                        typeAudit = "MAE"
                    elif objet == "Mission d'inspection":
                        typeAudit = "MAI"
                    else:
                        typeAudit = "MAR"

                    if code_theme == "Unite":
                        codeTheme = "U"
                    elif objet == "Structure":
                        codeTheme = "S"
                    elif objet == "Risque":
                        codeTheme = "P"
                    else:
                        typeAudit = "MAR"

                    if date_depart.month < 10:
                        month = '0{}'.format(date_depart.month)
                    else:
                        month = date_depart.month

                    parametres_sheet['B2'] = ira
                    parametres_sheet['B3'] = objet
                    parametres_sheet['B4'] = structure[-3:]+'-' + \
                        typeAudit+'-'+codeTheme+'-'+numero_ordre + '-'+month \
                        + '-{}'.format(date_depart.year)
                    parametres_sheet['B5'] = date_depart
                    parametres_sheet['B6'] = date_retour
                    parametres_sheet['B7'] = structure
                    parametres_sheet['B9'] = numero_ordre
                    parametres_sheet['B10'] = chef_mission
                    parametres_sheet['B11'] = auditeur
                    # PARAMETRES['B8'] = structures.filter(like=structure, axis=0)['gre'].values[0]
                    book.save('rapport audit.xlsx')

if authenticate():
    main()
else:
    st.set_page_config(initial_sidebar_state='collapsed')
    st.markdown(
    """
<style>
    [data-testid="collapsedControl"] {
        display: none
    }
</style>
""",
    unsafe_allow_html=True,
)