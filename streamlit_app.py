import streamlit as st

st.logo("logo2.png",size="large", icon_image="logo2.png")

#definiowanie stron
home = st.Page("home.py", title="Home", icon=":material/home:")
converter = st.Page("converter.py", title="Converter", icon=":material/table_convert:")
directs_converter = st.Page("convert_directs.py", title="Convert Directs", icon=":material/transform:")

pg = st.navigation([home, converter, directs_converter])
pg.run()


