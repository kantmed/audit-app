import streamlit as st
import pandas as pd

path = "rapport audit.xlsx"
sheet_name = "Referenciel risques"

referenciel = pd.read_excel(path, sheet_name, skiprows=7)


st.header('Annexe N° 01-Referenciel risques ')
st.table(referenciel)

# gd=GridOptionsBuilder.from_dataframe(referenciel)
# gd.configure_pagination(enabled=True)
# gd.configure_columns(cellRenderer='agGroupCellRenderer')

# gd.configure_default_column(editable=True,groupable=True)
# sel_mode=st.radio('Selection Mode',['single','multiple'])
# gd.configure_selection(selection_mode=sel_mode, use_checkbox=True)
# gdOptions=gd.build()
# AgGrid(referenciel,gridOptions=gdOptions,height=500,
#         width='100%')


