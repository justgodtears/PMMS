import streamlit as st

st.title("Route Calculator")

option = st.selectbox(
    "Please choose the carrier from the list",
    ("Deus", "Logitec", "Ihro"),
    index=None,
    placeholder="Select carrier...",
)

st.write("You selected:", option)


number = st.number_input(
    "Insert a distance", value=None, placeholder="Type a amount of kms..."
)
st.write("The current amount is ", number, "km")



if option and number is not None:
 if option == "Deus":
    stawka = (number * 0.75) + 500
    st.metric(label="Price", value=(f"{stawka}€"))
 elif option == "Logitec":
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
 elif option == "Ihro":
    stawka = (number * 0.94) + 465
    st.metric(label="Price", value=(f"{stawka}€"))
 elif option == "None":
     st.metric(label="Price", value="-")


