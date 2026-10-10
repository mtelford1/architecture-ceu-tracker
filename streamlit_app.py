import streamlit as st
import requests
import base64
import json

from datetime import datetime, timezone
from io import BytesIO
from urllib.parse import quote
from pypdf import PdfReader

# ---------------------------------------------------------
# PDF READER
# ---------------------------------------------------------

def read_onedrive_pdf(file_id, access_token):

    url = (
        "https://graph.microsoft.com/v1.0/"
        f"me/drive/items/{file_id}/content"
    )

    response = requests.get(
        url,
        headers={
            "Authorization": f"Bearer {access_token}"
        },
        timeout=60
    )

    response.raise_for_status()

    pdf_file = BytesIO(response.content)

    reader = PdfReader(pdf_file)

    text = []

    for page in reader.pages:

        page_text = page.extract_text()

        if page_text:
            text.append(page_text)

    return "\n".join(text)

# ---------------------------------------------------------
# CERTIFICATE CREDIT PARSER
# ---------------------------------------------------------

import re


def extract_aia_credits(text):

    clean_text = re.sub(
        r"\s+",
        " ",
        text
    )

    aia_credit = None
    hsw_credit = 0.0

    # Look for values such as:
    # 1.0 AIA LU
    # 1 AIA LU
    # 1.0 LU
    # 1 LU/HSW
    # 1.0 HSW

    aia_patterns = [
        r"(\d+(?:\.\d+)?)\s*AIA\s*LU",
        r"(\d+(?:\.\d+)?)\s*LU\b",
    ]

    hsw_patterns = [
        r"(\d+(?:\.\d+)?)\s*(?:AIA\s*)?LU(?:s)?\s*(?:[/|\-]|\(|\[)?\s*HSW\b",
        r"(\d+(?:\.\d+)?)\s*HSW\b",
        r"\bHSW\b\s*(?:LU|Credit|Credits|Hour|Hours)?\s*[:\-]?\s*(\d+(?:\.\d+)?)",
        r"(\d+(?:\.\d+)?)\s*(?:AIA\s*)?LU(?:s)?\s*\(?\s*Health\s*,?\s*Safety\s*(?:and|&)\s*Welfare",
    ]

    # Find AIA/LU credit
    for pattern in aia_patterns:

        match = re.search(
            pattern,
            clean_text,
            re.IGNORECASE
        )

        if match:

            value = float(
                match.group(1)
            )

            if 0 < value <= 20:
                aia_credit = value
                break

    # Find HSW credit
    for pattern in hsw_patterns:

        match = re.search(
            pattern,
            clean_text,
            re.IGNORECASE
        )

        if match:

            value = float(
                match.group(1)
            )

            if 0 < value <= 20:
                hsw_credit = value
                break

    # If certificate says something like
    # "1.0 LU/HSW", that credit counts as
    # both AIA LU and HSW.
    if aia_credit is None and hsw_credit > 0:
        aia_credit = hsw_credit

    return aia_credit, hsw_credit

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
# MICROSOFT ACCESS TOKEN CHECK
# ---------------------------------------------------------

if "access" in st.user.tokens:

    access_token = st.user.tokens["access"]

    try:
        payload_part = access_token.split(".")[1]
        payload_part += "=" * (-len(payload_part) % 4)

        payload = json.loads(
            base64.urlsafe_b64decode(
                payload_part
            ).decode("utf-8")
        )

        token_expiration = payload.get("exp", 0)

        current_timestamp = datetime.now(
            timezone.utc
        ).timestamp()

        if token_expiration <= current_timestamp:

            st.warning(
                "Your Microsoft OneDrive connection has expired."
            )

            if st.button("Reconnect Microsoft"):
                st.logout()

            st.stop()

    except Exception:

        st.error(
            "Microsoft connection could not be verified."
        )

        st.stop()

else:

    st.error(
        "Microsoft access token not available."
    )

    st.stop()
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
    st.write(f"Status code: {search_response.status_code}")
    st.write(search_response.text)

# ---------------------------------------------------------
# CERTIFICATE DATA LOAD
# ---------------------------------------------------------

certificate_data_by_year = {}

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

        certificate_data_by_year[year] = {
            "records": [],
            "total_aia": 0.0,
            "total_hsw": 0.0,
            "needs_review": 0,
            "folder_error": True,
        }

        continue

    items = files_response.json().get(
        "value",
        []
    )

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

    for certificate_file in certificate_files:

        file_name = certificate_file.get(
            "name",
            "Unnamed Certificate"
        )

        file_id = certificate_file["id"]

        try:

            certificate_text = read_onedrive_pdf(
                file_id,
                access_token
            )

            aia_credit, hsw_credit = extract_aia_credits(
                certificate_text
            )

            certificate_records.append(
                {
                    "name": file_name,
                    "aia": aia_credit,
                    "hsw": hsw_credit,
                    "text": certificate_text,
                    "status": (
                        "Read"
                        if aia_credit is not None
                        else "Needs review"
                    )
                }
            )

        except Exception:

            certificate_records.append(
                {
                    "name": file_name,
                    "aia": None,
                    "hsw": None,
                    "text": "",
                    "status": "Could not read"
                }
            )

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
        if record["status"] != "Read"
    )

    certificate_data_by_year[year] = {
        "records": certificate_records,
        "total_aia": total_aia,
        "total_hsw": total_hsw,
        "needs_review": needs_review,
        "folder_error": False,
    }


year_credit_totals = {
    year: data["total_aia"]
    for year, data in certificate_data_by_year.items()
}


# ---------------------------------------------------------
# ROLLOVER CALCULATIONS
# ---------------------------------------------------------

current_year = datetime.now(
    timezone.utc
).year


def calculate_annual_rollover(
    year_totals,
    required_credits,
    start_year,
    current_year
):

    rollover = 0.0

    current_previous = 0.0
    current_earned = 0.0

    for year in range(
        start_year,
        current_year + 1
    ):

        previous_rollover = rollover
        earned = year_totals.get(
            year,
            0.0
        )

        rollover = max(
            0.0,
            previous_rollover
            + earned
            - required_credits
        )

        if year == current_year:
            current_previous = previous_rollover
            current_earned = earned

    return {
        "previous_rollover": current_previous,
        "earned": current_earned,
        "required": required_credits,
        "rollover": rollover,
    }


def get_biennial_cycle(
    anchor_year,
    year
):

    cycle_start = (
        anchor_year
        + ((year - anchor_year) // 2) * 2
    )

    return (
        cycle_start,
        cycle_start + 1
    )


def calculate_biennial_rollover(
    year_totals,
    required_credits,
    anchor_year,
    current_year
):

    current_cycle_start, current_cycle_end = get_biennial_cycle(
        anchor_year,
        current_year
    )

    rollover = 0.0

    current_previous = 0.0
    current_earned = 0.0

    cycle_start = anchor_year

    while cycle_start <= current_cycle_start:

        cycle_end = cycle_start + 1

        previous_rollover = rollover

        earned = (
            year_totals.get(
                cycle_start,
                0.0
            )
            + year_totals.get(
                cycle_end,
                0.0
            )
        )

        rollover = max(
            0.0,
            previous_rollover
            + earned
            - required_credits
        )

        if cycle_start == current_cycle_start:
            current_previous = previous_rollover
            current_earned = earned

        cycle_start += 2

    return {
        "cycle_start": current_cycle_start,
        "cycle_end": current_cycle_end,
        "previous_rollover": current_previous,
        "earned": current_earned,
        "required": required_credits,
        "rollover": rollover,
    }


aia_start_year = (
    min(year_credit_totals)
    if year_credit_totals
    else current_year
)

aia_rollover = calculate_annual_rollover(
    year_credit_totals,
    18,
    aia_start_year,
    current_year
)

virginia_cycle_start, virginia_cycle_end = get_biennial_cycle(
    2015,
    current_year
)

washington_rollover = calculate_biennial_rollover(
    year_credit_totals,
    24,
    2025,
    current_year
)

aia_current_credits = year_credit_totals.get(
    current_year,
    0.0
)

virginia_current_credits = (
    year_credit_totals.get(
        virginia_cycle_start,
        0.0
    )
    + year_credit_totals.get(
        virginia_cycle_end,
        0.0
    )
)

washington_current_credits = (
    year_credit_totals.get(
        washington_rollover["cycle_start"],
        0.0
    )
    + year_credit_totals.get(
        washington_rollover["cycle_end"],
        0.0
    )
)


# ---------------------------------------------------------
# CREDENTIAL DATA
# ---------------------------------------------------------

credentials = {

    "AIA": {
        "full_name": "American Institute of Architects",
        "cycle_length": "1-Year",
        "current_cycle": str(current_year),
        "total_required": 18,
        "hsw_required": 12,
        "credit_type": "LU",
        "current_credits": aia_current_credits,
        "rollover_credits": aia_rollover["previous_rollover"],
        "previous_rollover": aia_rollover["previous_rollover"],
        "current_credits": aia_rollover["earned"],
        "rollover_period": "year",
        "rollover_allowed": True,
    },

    "Virginia": {
        "full_name": "Virginia Architect License",
        "cycle_length": "2-Year",
        "current_cycle": (
            f'{virginia_cycle_start}'
            f'–{virginia_cycle_end}'
        ),
        "total_required": 16,
        "hsw_required": "No separate minimum",
        "credit_type": "CE",
        "current_credits": virginia_current_credits,
        "rollover_credits": 0.0,
        "rollover_allowed": False,
    },

    "Washington": {
        "full_name": "Washington Architect License",
        "cycle_length": "2-Year",
        "current_cycle": (
            f'{washington_rollover["cycle_start"]}'
            f'–{washington_rollover["cycle_end"]}'
        ),
        "total_required": 24,
        "hsw_required": 16,
        "credit_type": "PDH",
        "current_credits": washington_current_credits,
        "rollover_credits": washington_rollover["previous_rollover"],
        "previous_rollover": washington_rollover["previous_rollover"],
        "current_credits": washington_rollover["earned"],
        "rollover_period": "cycle",
        "rollover_allowed": True,
    },
}


# ---------------------------------------------------------
# CREDENTIAL CARD FUNCTION
# ---------------------------------------------------------

def credential_card(name, data):

    with st.container(border=True):

        st.subheader(name)

        st.caption(data["full_name"])

        st.markdown("**Cycle**")
        st.write(data["cycle_length"])

        st.markdown("**Current Cycle**")
        st.write(data["current_cycle"])

        st.markdown("**Current Credits to Date**")

        current_credit_text = (
            f'{data["current_credits"]:.2f}'
        )

        if (
            data["current_credits"]
            >= data["total_required"]
        ):
            current_credit_text += " — COMPLETE"

        st.write(current_credit_text)

        st.divider()

        st.markdown("**Credit Requirements**")
        st.write(
            f'Total: {data["total_required"]} {data["credit_type"]}'
        )
        st.write(
            f'HSW: {data["hsw_required"]}'
        )

        st.divider()

        st.markdown("**Rollover Credits**")

        if data.get("rollover_allowed", True):

            rollover_source = (
                "Previous Year"
                if data["cycle_length"] == "1-Year"
                else "Previous Cycle"
            )

            st.write(
                f'{rollover_source}: {data["rollover_credits"]:.2f}'
            )

        else:

            st.write("Not permitted")


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
    "Continuing education certificates from OneDrive."
)

for year_folder in year_folders:

    year = year_folder["year"]

    year_data = certificate_data_by_year.get(
        year,
        {
            "records": [],
            "total_aia": 0.0,
            "total_hsw": 0.0,
            "needs_review": 0,
            "folder_error": True,
        }
    )

    if year_data["folder_error"]:

        with st.expander(str(year)):

            st.error(
                "Could not read this year's OneDrive folder."
            )

        continue

    certificate_records = year_data["records"]
    total_aia = year_data["total_aia"]
    total_hsw = year_data["total_hsw"]
    needs_review = year_data["needs_review"]

    expander_title = (
        f"{year} — "
        f"{total_aia:.1f} AIA / "
        f"{total_hsw:.1f} HSW"
    )

    if needs_review > 0:

        expander_title += (
            f" — {needs_review} needs review"
        )

    with st.expander(expander_title):

        if not certificate_records:

            st.caption(
                "No PDF certificates found."
            )

        else:

            name_col, aia_col, hsw_col = st.columns(
                [6, 1, 1]
            )

            with name_col:
                st.markdown(
                    "**Certificate**"
                )

            with aia_col:
                st.markdown(
                    "**AIA**"
                )

            with hsw_col:
                st.markdown(
                    "**HSW**"
                )

            st.divider()

            for index, record in enumerate(
                certificate_records
            ):

                name_col, aia_col, hsw_col = st.columns(
                    [6, 1, 1]
                )

                with name_col:

                    st.write(
                        record["name"]
                    )

                    if record["status"] != "Read":

                        st.caption(
                            f"⚠ {record['status']}"
                        )

                        show_text = st.checkbox(
                            "Show extracted PDF text",
                            key=(
                                f"show_text_"
                                f"{year}_"
                                f"{index}"
                            )
                        )

                        if show_text:

                            extracted_text = record.get(
                                "text",
                                ""
                            )

                            if extracted_text.strip():

                                st.text_area(
                                    "Extracted text",
                                    value=extracted_text[:5000],
                                    height=200,
                                    disabled=True,
                                    key=(
                                        f"pdf_text_"
                                        f"{year}_"
                                        f"{index}"
                                    )
                                )

                            else:

                                st.write(
                                    "No readable text was "
                                    "extracted from this PDF."
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

            st.divider()

            total_col, aia_total_col, hsw_total_col = st.columns(
                [6, 1, 1]
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
