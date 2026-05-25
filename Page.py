import streamlit as st

#SetupPage
Summary = st.Page(
    page="Summary.py",
    title="Summary",
    default=True,
)
Input_Data = st.Page(
    page="Input Data.py",
    title="Input Data",
)
Pivot = st.Page(
    page="Pivot.py",
    title="Pivot",
)
pg = st.navigation({"Choose": [Pivot, Summary, Input_Data]})
pg.run()