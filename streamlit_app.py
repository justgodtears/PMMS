import streamlit as st

st.logo("logo2.png",size="large", icon_image="logo2.png")

#definiowanie stron
home = st.Page("home.py", title="Home", icon=":material/home:")
converter = st.Page("converter.py", title="Converter", icon=":material/table_convert:")


pg = st.navigation([home, converter])
pg.run()


