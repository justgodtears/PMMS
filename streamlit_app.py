import streamlit as st

st.logo("logo2.png",size="large", icon_image="logo2.png")

#definiowanie stron
home = st.Page("home.py", title="Home", icon=":material/home:")
dashboard = st.Page("dashboard.py", title="Dashboard", icon=":material/dashboard:")
route_calculator = st.Page("routcalc.py", title="Route Calculator", icon=":material/calculate:")

pg = st.navigation([home, dashboard, route_calculator])
pg.run()


