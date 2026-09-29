import streamlit as st
import requests
from urllib.parse import quote

# ---------------------------------------------------------
# PAGE SETUP
# ---------------------------------------------------------

st.set_page_config(
    page_title="Architecture CEU Tracker",
    page_icon="📚",
    layout="wide"
)

# ---------------------------------------------------------
# MICROSOFT LOGIN
# ---------------------------------------------------------

if not st.user.is_logged_in:
    st.title("Architecture CEU Tracker")
    st.write("Sign in with Microsoft to access your continuing education records.")

    if st.button("Sign in with Microsoft"):
        st.login()

    st.stop()

else:
    login_col1, login_col2 = st.columns([6, 1])

    with login_col1:
        st.caption(f"Signed in as {st.user.name}")

    with login_col2:
        if st.button("Sign out"):
            st.logout()

# ---------------------------------------------------------
# MICROSOFT ACCESS TEST
# ---------------------------------------------------------

if "access" in st.user.tokens:
    st.success("Microsoft access token available.")
else:
    st.error("Microsoft access token not available.")

# ---------------------------------------------------------
# ONEDRIVE FOLDER TEST
# ---------------------------------------------------------

url = (
    "https://graph.microsoft.com/v1.0/"
    "me/drive/root/search(q='08_Continuing Education')"
)

headers = {
    "Authorization": f"Bearer {st.user.tokens['access']}"
}

response = requests.get(
    url,
    headers=headers,
    timeout=30
)

if response.status_code == 200:
    items = response.json().get("value", [])

    st.success("OneDrive CE folder found.")

    for item in items:
        st.write(item["name"])

else:
    st.error("Could not read the OneDrive CE folder.")
    st.write(response.status_code)
    st.write(response.text)

# ---------------------------------------------------------
# BASIC STYLING
# ---------------------------------------------------------

st.markdown(
    """
    <style>

    .block-container {
        max-width: 1400px;
        padding-top: 2rem;
    }

    .credential-title {
        font-size: 1.6rem;
        font-weight: 700;
        margin-bottom: 0.2rem;
    }

    .credential-subtitle {
        color: #777;
        font-size: 0.9rem;
        margin-bottom: 1.5rem;
    }

    .section-heading {
        font-size: 0.78rem;
        font-weight: 700;
        text-transform: uppercase;
        letter-spacing: 0.06rem;
        color: #777;
        margin-top: 1rem;
        margin-bottom: 0.25rem;
    }

    .large-value {
        font-size: 1.4rem;
        font-weight: 600;
        margin-bottom: 0.5rem;
    }

    .requirement-line {
        font-size: 1rem;
        margin-bottom: 0.25rem;
    }

    </style>
    """,
    unsafe_allow_html=True
)


# ---------------------------------------------------------
# CREDENTIAL DATA
#
# For now these are manual values.
# We will connect them to a database later.
# ---------------------------------------------------------

credentials = {

    "AIA": {
        "full_name": "American Institute of Architects",
        "cycle_length": "1-Year",
        "current_cycle": "2026",
        "total_required": 18,
        "hsw_required": 12,
        "credit_type": "LU",
        "rollover_credits": 0.0,
    },

    "Virginia": {
        "full_name": "Virginia Architect License",
        "cycle_length": "2-Year",
        "current_cycle": "2025–2026",
        "total_required": "TBD",
        "hsw_required": "TBD",
        "credit_type": "CE",
        "rollover_credits": 0.0,
    },

    "Washington": {
        "full_name": "Washington Architect License",
        "cycle_length": "2-Year",
        "current_cycle": "2025–2026",
        "total_required": "TBD",
        "hsw_required": "TBD",
        "credit_type": "PDH",
        "rollover_credits": 0.0,
    },
}


# ---------------------------------------------------------
# CREDENTIAL CARD FUNCTION
# ---------------------------------------------------------

def credential_card(name, data):

    with st.container(border=True):

        st.markdown(
            f'<div class="credential-title">{name}</div>',
            unsafe_allow_html=True
        )

        st.markdown(
            f'<div class="credential-subtitle">'
            f'{data["full_name"]}'
            f'</div>',
            unsafe_allow_html=True
        )

        # -------------------------
        # CYCLE
        # -------------------------

        st.markdown(
            '<div class="section-heading">Cycle</div>',
            unsafe_allow_html=True
        )

        st.markdown(
            f'<div class="large-value">'
            f'{data["cycle_length"]}'
            f'</div>',
            unsafe_allow_html=True
        )

        # -------------------------
        # CURRENT CYCLE
        # -------------------------

        st.markdown(
            '<div class="section-heading">Current Cycle</div>',
            unsafe_allow_html=True
        )

        st.markdown(
            f'<div class="large-value">'
            f'{data["current_cycle"]}'
            f'</div>',
            unsafe_allow_html=True
        )

        st.divider()

        # -------------------------
        # CREDIT REQUIREMENTS
        # -------------------------

        st.markdown(
            '<div class="section-heading">'
            'Credit Requirements'
            '</div>',
            unsafe_allow_html=True
        )

        st.markdown(
            f'<div class="requirement-line">'
            f'<strong>Total:</strong> '
            f'{data["total_required"]} {data["credit_type"]}'
            f'</div>',
            unsafe_allow_html=True
        )

        st.markdown(
            f'<div class="requirement-line">'
            f'<strong>HSW:</strong> '
            f'{data["hsw_required"]}'
            f'</div>',
            unsafe_allow_html=True
        )

        st.divider()

        # -------------------------
        # ROLLOVER
        # -------------------------

        st.markdown(
            '<div class="section-heading">'
            'Rollover Credits'
            '</div>',
            unsafe_allow_html=True
        )

        st.metric(
            label="From Previous Cycle",
            value=f'{data["rollover_credits"]:g}'
        )


# ---------------------------------------------------------
# PAGE HEADER
# ---------------------------------------------------------

st.title("Architecture CEU Tracker")

st.write(
    "AIA + Virginia + Washington Continuing Education"
)

st.divider()


# ---------------------------------------------------------
# CREDENTIAL CARDS
# ---------------------------------------------------------

col1, col2, col3 = st.columns(3)

with col1:
    credential_card(
        "AIA",
        credentials["AIA"]
    )

with col2:
    credential_card(
        "Virginia",
        credentials["Virginia"]
    )

with col3:
    credential_card(
        "Washington",
        credentials["Washington"]
    )
    # ---------------------------------------------------------
# CERTIFICATES BY YEAR
# ---------------------------------------------------------

st.divider()

st.subheader("Certificates by Year")

st.caption(
    "Continuing education certificates organized by completion year."
)


# ---------------------------------------------------------
# TEMPORARY SAMPLE DATA
#
# This is ONLY to develop the visual layout.
# Eventually this information will come automatically
# from the OneDrive certificate folders.
# ---------------------------------------------------------

certificate_data = {

    2026: [
        {
            "course": "Accessible Means of Egress",
            "date": "03/12/2026",
            "aia": 1.0,
            "hsw": 1.0,
        },
        {
            "course": "Building Envelope Design",
            "date": "05/07/2026",
            "aia": 1.0,
            "hsw": 1.0,
        },
        {
            "course": "Revit: Advanced Documentation",
            "date": "07/16/2026",
            "aia": 1.0,
            "hsw": 0.0,
        },
    ],

    2025: [
        {
            "course": "Energy Code Fundamentals",
            "date": "02/20/2025",
            "aia": 1.0,
            "hsw": 1.0,
        },
        {
            "course": "Mass Timber Construction",
            "date": "06/11/2025",
            "aia": 2.0,
            "hsw": 2.0,
        },
    ],
}


# ---------------------------------------------------------
# DISPLAY YEARS
# ---------------------------------------------------------

for year in range(2026, 2014, -1):

    certificates = certificate_data.get(year, [])

    total_aia = sum(
        certificate["aia"]
        for certificate in certificates
    )

    total_hsw = sum(
        certificate["hsw"]
        for certificate in certificates
    )

    expander_title = (
        f"{year} — "
        f"{total_aia:g} AIA / "
        f"{total_hsw:g} HSW"
    )

    with st.expander(expander_title):

        if not certificates:

            st.caption(
                "No certificates loaded yet."
            )

        else:

            # Column headings

            course_col, date_col, aia_col, hsw_col = st.columns(
                [5, 2, 1, 1]
            )

            with course_col:
                st.markdown("**Course / Certificate**")

            with date_col:
                st.markdown("**Date**")

            with aia_col:
                st.markdown("**AIA**")

            with hsw_col:
                st.markdown("**HSW**")

            st.divider()

            # Certificate rows

            for certificate in certificates:

                course_col, date_col, aia_col, hsw_col = st.columns(
                    [5, 2, 1, 1]
                )

                with course_col:
                    st.write(certificate["course"])

                with date_col:
                    st.write(certificate["date"])

                with aia_col:
                    st.write(
                        f'{certificate["aia"]:.1f}'
                    )

                with hsw_col:
                    st.write(
                        f'{certificate["hsw"]:.1f}'
                    )

            st.divider()

            # Year totals

            total_col, blank_col, aia_col, hsw_col = st.columns(
                [5, 2, 1, 1]
            )

            with total_col:
                st.markdown("**TOTAL**")

            with aia_col:
                st.markdown(
                    f"**{total_aia:.1f}**"
                )

            with hsw_col:
                st.markdown(
                    f"**{total_hsw:.1f}**"
                )