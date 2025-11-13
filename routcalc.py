import streamlit as st

st.set_page_config(layout="wide")

st.title("Route Calculator")

carrier = st.selectbox(
    "Please choose the carrier from the list",
    ("Deus 24T","Logitec 12T", "Logitec 24T", "Ihro 12T", "Ihro 24T", "BGM Van", "BGM 12T"),
    index=None,
    placeholder="Select carrier...",
)

st.write("You selected:", carrier)

number = st.number_input(
    "Insert a distance", value=None, placeholder="Type a amount of kms..."
)
st.write("The current amount is ", number, "km")



if carrier and number is not None:
 ### Tutaj jest Deus
 if carrier == "Deus 24T":
    stawka = (number * 0.75) + 500
    st.metric(label="Price", value=(f"{stawka}€"))
 ### Tutaj jest logitec
 elif carrier == "Logitec 24T":
    if number <= 250:
        stawka = 650
        st.metric(label="Price", value=(f"{stawka}€"))
    elif number <= 350:
        stawka = 750
        st.metric(label="Price", value=(f"{stawka}€"))
    elif number <= 450:
        stawka = 900
        st.metric(label="Price", value=(f"{stawka}€"))
    else:
        stawka = 1030
        st.metric(label="Price", value=(f"{stawka}€"))
 elif carrier == "Logitec 12T":
    if number <= 250:
        stawka = 550
        st.metric(label="Price", value=(f"{stawka}€"))
    elif number <= 350:
        stawka = 600
        st.metric(label="Price", value=(f"{stawka}€"))
    elif number <= 450:
        stawka = 650
        st.metric(label="Price", value=(f"{stawka}€"))
    elif number <= 550:
        stawka = 750
        st.metric(label="Price", value=(f"{stawka}€"))
    else:
        stawka = 900
        st.metric(label="Price", value=(f"{stawka}€"))
 ### Tutaj jest Ihro
 elif carrier == "Ihro 24T":
    stawka = (number * 0.94) + 465
    st.metric(label="Price", value=(f"{stawka}€"))
 elif carrier == "Ihro 12T":
     if number <= 250:
         stawka = 620
         st.metric(label="Price", value=(f"{stawka}€"))
     else:
         nadmiar_km = number - 250
         stawka = 620 + (nadmiar_km * 0.85)
         st.metric(label="Price", value=(f"{stawka}€"))
 ### Tutaj jest BGM
 elif carrier == "BGM Van":
     if number <= 350:
         stawka = 290
         st.metric(label="Price", value=(f"{stawka}€"))
     else:
         nadmiar_km = number - 350
         stawka = 290 + (nadmiar_km * 0.75)
         st.metric(label="Price", value=(f"{stawka}€"))
 elif carrier == "BGM 12T":
     if number <= 350:
         stawka = 600
         st.metric(label="Price", value=(f"{stawka}€"))
     else:
         nadmiar_km = number - 350
         stawka = 600 + (nadmiar_km * 0.75)
         st.metric(label="Price", value=(f"{stawka}€"))
 ### Tutaj jak brak wyboru
 elif carrier == "None":
     st.metric(label="Price", value="-")



