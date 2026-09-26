import streamlit as st

st.set_page_config(
    page_title="Architecture CEU Tracker",
    page_icon="📚",
    layout="wide"
)

st.title("Architecture CEU Tracker")
st.write("AIA + Virginia + Washington Continuing Education")

st.divider()

st.subheader("Credits from Previous Year")

col1, col2, col3 = st.columns(3)

with col1:
    st.metric("AIA", "0")

with col2:
    st.metric("Virginia", "0")

with col3:
    st.metric("Washington", "0")

st.divider()

col1, col2, col3 = st.columns(3)

with col1:
    st.subheader("AIA")
    st.write("**Cycle:** Annual")
    st.write("**Credits Required:** 18")
    st.write("**Credits Achieved:** 0")
    st.write("**HSW Required:** 12")
    st.write("**HSW Achieved:** 0")

with col2:
    st.subheader("Virginia License")
    st.write("**Cycle:** Biennial")
    st.write("**Credits Required:** TBD")
    st.write("**Credits Achieved:** 0")
    st.write("**HSW Required:** TBD")
    st.write("**HSW Achieved:** 0")

with col3:
    st.subheader("Washington License")
    st.write("**Cycle:** Biennial")
    st.write("**Credits Required:** TBD")
    st.write("**Credits Achieved:** 0")
    st.write("**HSW Required:** TBD")
    st.write("**HSW Achieved:** 0")