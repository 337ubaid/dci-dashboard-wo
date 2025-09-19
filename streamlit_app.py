import streamlit as st

pages = [
    st.Page(
        "home.py",
        title="Home",
        icon=":material/home:"
    ),
]

page = st.navigation(pages)
page.run()


st.sidebar.caption(
    "Developed by Digistar Class Intern 2025"
)
