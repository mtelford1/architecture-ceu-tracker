import streamlit as st
import requests
import re

from io import BytesIO
from urllib.parse import quote
from pypdf import PdfReader

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
# ONEDRIVE DATA LOAD
# ---------------------------------------------------------

year_folders = []

headers = {
    "Authorization": f"Bearer {st.user.tokens['access']}"
}

# Search OneDrive for the Continuing Education folder
search_url = (
    "https://graph.microsoft.com/v1.0/"
    "me/drive/root/search(q='08_Continuing Education')"
)

search_response = requests.get(
    search_url,
    headers=headers,
    timeout=30
)

if search_response.status_code == 200:

    search_items = search_response.json().get("value", [])

    ce_folder = next(
        (
            item for item in search_items
            if item.get("name") == "08_Continuing Education"
            and "folder" in item
        ),
        None
    )

    if ce_folder:

        ce_folder_id = ce_folder["id"]

        ce_children_url = (
            "https://graph.microsoft.com/v1.0/"
            f"me/drive/items/{ce_folder_id}/children"
        )

        ce_children_response = requests.get(
            ce_children_url,
            headers=headers,
            timeout=30
        )

        if ce_children_response.status_code == 200:

            ce_children = ce_children_response.json().get("value", [])

            # Find the AIA folder
            aia_folder = next(
                (
                    item for item in ce_children
                    if item.get("name") == "AIA"
                    and "folder" in item
                ),
                None
            )

            if aia_folder:

                aia_folder_id = aia_folder["id"]

                aia_children_url = (
                    "https://graph.microsoft.com/v1.0/"
                    f"me/drive/items/{aia_folder_id}/children"
                )

                aia_children_response = requests.get(
                    aia_children_url,
                    headers=headers,
                    timeout=30
                )

                if aia_children_response.status_code == 200:

                    aia_children = aia_children_response.json().get(
                        "value",
                        []
                    )

                    # Keep only 4-digit year folders
                    for item in aia_children:

                        name = item.get("name", "")

                        if (
                            "folder" in item
                            and len(name) == 4
                            and name.isdigit()
                        ):
                            year_folders.append(
                                {
                                    "year": int(name),
                                    "name": name,
                                    "id": item["id"]
                                }
                            )

                    year_folders = sorted(
                        year_folders,
                        key=lambda x: x["year"],
                        reverse=True
                    )

                else:
                    st.error("Could not read the AIA folder.")

            else:
                st.error("AIA folder not found.")

        else:
            st.error("Could not read the Continuing Education folder.")

    else:
        st.error("Could not find the Continuing Education folder.")

else:
    st.error("Could not search OneDrive.")

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
# CERTIFICATE READING
# ---------------------------------------------------------

@st.cache_data(
    ttl=86400,
    show_spinner=False
)
def read_pdf_text(
    file_id,
    etag,
    _access_token
):

    url = (
        "https://graph.microsoft.com/v1.0/"
        f"me/drive/items/{file_id}/content"
    )

    response = requests.get(
        url,
        headers={
            "Authorization": f"Bearer {_access_token}"
        },
        timeout=60
    )

    response.raise_for_status()

    pdf_file = BytesIO(response.content)

    reader = PdfReader(pdf_file)

    pages = []

    for page in reader.pages:

        page_text = page.extract_text()

        if page_text:
            pages.append(page_text)

    return "\n".join(pages)


# ---------------------------------------------------------
# CERTIFICATE DATA EXTRACTION
# ---------------------------------------------------------

def extract_certificate_data(
    text,
    filename
):

    # Default course title is the filename
    title = re.sub(
        r"\.pdf$",
        "",
        filename,
        flags=re.IGNORECASE
    )

    completion_date = ""

    aia_credit = None
    hsw_credit = 0.0

    if not text.strip():

        return {
            "title": title,
            "date": "",
            "aia": None,
            "hsw": None,
            "status": "No readable PDF text"
        }

    # ---------------------------------------------
    # CLEAN TEXT
    # ---------------------------------------------

    lines = []

    for line in text.splitlines():

        clean_line = re.sub(
            r"\s+",
            " ",
            line
        ).strip()

        if clean_line:
            lines.append(clean_line)

    compact_text = re.sub(
        r"\s+",
        " ",
        text
    )

    # ---------------------------------------------
    # COURSE TITLE
    # ---------------------------------------------

    title_patterns = [
        r"Course Title\s*[:\-]\s*(.+)",
        r"Course Name\s*[:\-]\s*(.+)",
        r"Program Title\s*[:\-]\s*(.+)",
        r"Program Name\s*[:\-]\s*(.+)",
    ]

    for pattern in title_patterns:

        match = re.search(
            pattern,
            text,
            re.IGNORECASE
        )

        if match:

            candidate = match.group(1).strip()

            candidate = re.sub(
                r"\s+",
                " ",
                candidate
            )

            if candidate:
                title = candidate

            break

    # ---------------------------------------------
    # COMPLETION DATE
    # ---------------------------------------------

    date_patterns = [
        r"Completion Date\s*[:\-]\s*([A-Za-z0-9,\/\-\s]+)",
        r"Date Completed\s*[:\-]\s*([A-Za-z0-9,\/\-\s]+)",
        r"Completed\s*[:\-]\s*([A-Za-z0-9,\/\-\s]+)",
    ]

    for pattern in date_patterns:

        match = re.search(
            pattern,
            text,
            re.IGNORECASE
        )

        if match:

            completion_date = re.sub(
                r"\s+",
                " ",
                match.group(1)
            ).strip()

            break

    # ---------------------------------------------
    # AIA LU / HSW
    # ---------------------------------------------

    aia_candidates = []
    hsw_candidates = []

    lu_pattern = re.compile(
        r"(?<!\d)"
        r"(\d+(?:\.\d+)?)"
        r"\s*"
        r"(?:AIA\s*)?"
        r"LU(?:s)?\b",
        re.IGNORECASE
    )

    for match in lu_pattern.finditer(
        compact_text
    ):

        value = float(
            match.group(1)
        )

        # Prevent things such as years or IDs
        # from being interpreted as credits.
        if 0 < value <= 20:

            aia_candidates.append(
                value
            )

            context_start = max(
                0,
                match.start() - 50
            )

            context_end = min(
                len(compact_text),
                match.end() + 50
            )

            context = compact_text[
                context_start:context_end
            ]

            if re.search(
                r"\bHSW\b",
                context,
                re.IGNORECASE
            ):
                hsw_candidates.append(
                    value
                )

    # Some certificates explicitly state
    # the HSW value separately.

    explicit_hsw_patterns = [
        r"(\d+(?:\.\d+)?)\s*LU\s*[/|]\s*HSW",
        r"(\d+(?:\.\d+)?)\s*HSW\b",
        r"\bHSW\s*(?:LU|Credit|Credits)?"
        r"\s*[:\-]?\s*"
        r"(\d+(?:\.\d+)?)",
    ]

    for pattern in explicit_hsw_patterns:

        matches = re.findall(
            pattern,
            compact_text,
            re.IGNORECASE
        )

        for value in matches:

            try:
                number = float(value)

                if 0 < number <= 20:
                    hsw_candidates.append(
                        number
                    )

            except ValueError:
                pass

    # ---------------------------------------------
    # FINAL CREDIT VALUES
    # ---------------------------------------------

    if aia_candidates:

        aia_credit = max(
            aia_candidates
        )

    if hsw_candidates:

        hsw_credit = max(
            hsw_candidates
        )

    # If the document clearly identifies
    # HSW credit but the LU parser missed
    # the general AIA amount, HSW necessarily
    # counts as LU as well.

    if (
        aia_credit is None
        and hsw_credit > 0
    ):
        aia_credit = hsw_credit

    if aia_credit is None:

        status = "Credits need review"

    else:

        status = "Parsed"

    return {
        "title": title,
        "date": completion_date,
        "aia": aia_credit,
        "hsw": hsw_credit,
        "status": status
    }

# ---------------------------------------------------------
# CERTIFICATES BY YEAR
# ---------------------------------------------------------

st.divider()

st.subheader(
    "Certificates by Year"
)

st.caption(
    "Certificate information is read automatically "
    "from the PDFs stored in OneDrive."
)

access_token = st.user.tokens[
    "access"
]

for year_folder in year_folders:

    year = year_folder["year"]

    folder_id = year_folder["id"]

    files_url = (
        "https://graph.microsoft.com/v1.0/"
        f"me/drive/items/{folder_id}/children"
    )

    files_response = requests.get(
        files_url,
        headers=headers,
        timeout=30
    )

    if files_response.status_code != 200:

        with st.expander(
            str(year)
        ):

            st.error(
                "Could not read this "
                "year's OneDrive folder."
            )

        continue

    items = files_response.json().get(
        "value",
        []
    )

    # Only process PDF files.
    certificate_files = [
        item
        for item in items
        if (
            "file" in item
            and item.get(
                "name",
                ""
            ).lower().endswith(".pdf")
        )
    ]

    certificate_files = sorted(
        certificate_files,
        key=lambda x: x.get(
            "name",
            ""
        ).lower()
    )

    certificate_records = []

    with st.spinner(
        f"Reading {year} certificates..."
    ):

        for file in certificate_files:

            filename = file.get(
                "name",
                "Unnamed certificate"
            )

            file_id = file["id"]

            etag = file.get(
                "eTag",
                ""
            )

            try:

                pdf_text = read_pdf_text(
                    file_id,
                    etag,
                    access_token
                )

                certificate = (
                    extract_certificate_data(
                        pdf_text,
                        filename
                    )
                )

            except Exception:

                certificate = {
                    "title": re.sub(
                        r"\.pdf$",
                        "",
                        filename,
                        flags=re.IGNORECASE
                    ),
                    "date": "",
                    "aia": None,
                    "hsw": None,
                    "status": "Could not read PDF"
                }

            certificate[
                "filename"
            ] = filename

            certificate_records.append(
                certificate
            )

    # ---------------------------------------------
    # YEAR TOTALS
    # ---------------------------------------------

    total_aia = sum(
        record["aia"] or 0
        for record in certificate_records
    )

    total_hsw = sum(
        record["hsw"] or 0
        for record in certificate_records
    )

    needs_review = sum(
        1
        for record in certificate_records
        if record["status"] != "Parsed"
    )

    expander_title = (
        f"{year} — "
        f"{total_aia:.1f} AIA / "
        f"{total_hsw:.1f} HSW"
    )

    if needs_review:

        expander_title += (
            f" — {needs_review} needs review"
        )

    # ---------------------------------------------
    # YEAR EXPANDER
    # ---------------------------------------------

    with st.expander(
        expander_title
    ):

        if not certificate_records:

            st.caption(
                "No PDF certificates found."
            )

            continue

        heading_course, \
        heading_date, \
        heading_aia, \
        heading_hsw = st.columns(
            [5, 2, 1, 1]
        )

        with heading_course:
            st.markdown(
                "**Course / Certificate**"
            )

        with heading_date:
            st.markdown(
                "**Date**"
            )

        with heading_aia:
            st.markdown(
                "**AIA**"
            )

        with heading_hsw:
            st.markdown(
                "**HSW**"
            )

        st.divider()

        for record in certificate_records:

            course_col, \
            date_col, \
            aia_col, \
            hsw_col = st.columns(
                [5, 2, 1, 1]
            )

            with course_col:

                st.write(
                    record["title"]
                )

                if (
                    record["status"]
                    != "Parsed"
                ):

                    st.caption(
                        f"⚠ {record['status']}"
                    )

            with date_col:

                st.write(
                    record["date"]
                    or "—"
                )

            with aia_col:

                if record["aia"] is None:
                    st.write("—")
                else:
                    st.write(
                        f'{record["aia"]:.1f}'
                    )

            with hsw_col:

                if record["hsw"] is None:
                    st.write("—")
                else:
                    st.write(
                        f'{record["hsw"]:.1f}'
                    )

        # ---------------------------------------------
        # YEAR TOTAL
        # ---------------------------------------------

        st.divider()

        total_col, \
        blank_col, \
        aia_total_col, \
        hsw_total_col = st.columns(
            [5, 2, 1, 1]
        )

        with total_col:
            st.markdown(
                "**YEAR TOTAL**"
            )

        with aia_total_col:
            st.markdown(
                f"**{total_aia:.1f}**"
            )

        with hsw_total_col:
            st.markdown(
                f"**{total_hsw:.1f}**"
            )