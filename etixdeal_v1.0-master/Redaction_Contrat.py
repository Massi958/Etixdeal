import streamlit as st
from multiapp import MultiApp
from apps import Contrat_Avenant_CDI,Contrat_Sous_Traitant, Contrat_CDI, Contrat_Avenant # import your app modules here
st.set_page_config(layout="wide")
app = MultiApp()

def local_css(file_name):
    with open(file_name) as f:
        st.markdown('<style>{}</style>'.format(f.read()), unsafe_allow_html=True)

local_css("style.css")
# Add all your application here
app.add_app("Contrat Sous Traitant", Contrat_Sous_Traitant.app)
app.add_app("Contrat_CDI", Contrat_CDI.app)
app.add_app("Contrat_Avenant", Contrat_Avenant.app)
app.add_app("Contrat_Avenant_CDI", Contrat_Avenant_CDI.app)
# The main app
app.run()