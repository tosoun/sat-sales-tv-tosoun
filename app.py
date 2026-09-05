import base64
import datetime
import glob
import json
import os
import unicodedata
from io import BytesIO

import pandas as pd
import requests
import streamlit as st
import streamlit.components.v1 as components


st.set_page_config(
    page_title="Πωλήσεις 2 Προϊόντων ανά Κατάστημα 2026",
    layout="wide"
)


st.markdown(
    """
    <style>
    .stApp { background-color: #2c3e50 !important; }
    #MainMenu {visibility: hidden;}
    header {visibility: hidden;}
    footer {visibility: hidden;}
    [data-testid="stToolbar"] {visibility: hidden; display: none;}
    [data-testid="stDecoration"] {visibility: hidden; display: none;}

    div[data-baseweb="select"] > div,
    .stRadio label p {
        color: white !important;
    }

    .block-container {
        padding: 0rem 0.5rem !important;
        max-width: 100% !important;
    }

    div[data-testid="stHorizontalBlock"] {
        align-items: center !important;
    }

    /* ==================================================
       MOBILE COMPACT v3
       Admin + Περιφέρεια μένουν στην ΙΔΙΑ γραμμή.
       Τα δύο πλαίσια μικραίνουν.
       Το XLSX μπαίνει ακριβώς δίπλα στην Περιφέρεια.
       ================================================== */

    @media (max-width: 900px) {

        .block-container {
            padding: 0rem 0.30rem !important;
            max-width: 100% !important;
            width: 100% !important;
        }

        /* ΚΥΡΙΑ ΣΕΙΡΑ: Admin αριστερά / Περιφέρεια δεξιά */
        div[data-testid="stHorizontalBlock"]:has(div[data-testid="stExpander"]) {
            flex-direction: row !important;
            align-items: flex-start !important;
            gap: 0.35rem !important;
            width: 100% !important;
        }

        /* Admin = πιο μικρό */
        div[data-testid="stHorizontalBlock"]:has(div[data-testid="stExpander"])
        > div[data-testid="column"]:first-child {
            flex: 0 0 42% !important;
            width: 42% !important;
            min-width: 0 !important;
            max-width: 42% !important;
        }

        /* Περιφέρεια = λίγο μεγαλύτερο */
        div[data-testid="stHorizontalBlock"]:has(div[data-testid="stExpander"])
        > div[data-testid="column"]:nth-child(2) {
            flex: 0 0 58% !important;
            width: 58% !important;
            min-width: 0 !important;
            max-width: 58% !important;
        }

        /* Μικρότερο Admin expander */
        div[data-testid="stExpander"] {
            width: 100% !important;
        }

        div[data-testid="stExpander"] details summary {
            padding-left: 0.45rem !important;
            padding-right: 0.45rem !important;
            min-height: 42px !important;
        }

        div[data-testid="stExpander"] details summary p {
            font-size: 12px !important;
            white-space: nowrap !important;
            overflow: hidden !important;
            text-overflow: ellipsis !important;
        }

        /* Εσωτερική σειρά: Περιφέρεια + Excel δίπλα-δίπλα */
        div[data-testid="stHorizontalBlock"]
        div[data-testid="stHorizontalBlock"] {
            flex-direction: row !important;
            align-items: flex-end !important;
            gap: 0.25rem !important;
            width: 100% !important;
        }

        /* Select Περιφέρειας */
        div[data-testid="stHorizontalBlock"]
        div[data-testid="stHorizontalBlock"]
        > div[data-testid="column"]:first-child {
            flex: 1 1 auto !important;
            width: auto !important;
            min-width: 0 !important;
        }
    }
    </style>

    <script>
    function removeManageButton() {

        const doc = window.parent.document;
        const buttons = doc.querySelectorAll('button');

        buttons.forEach(btn => {

            if (
                btn.innerText.includes('Manage app') ||
                btn.innerHTML.includes('Manage')
            ) {
                btn.style.display = 'none';
            }

        });

        const pwdInputs =
            doc.querySelectorAll('input[type="password"]');

        pwdInputs.forEach(input => {

            input.setAttribute(
                'autocomplete',
                'username'
            );

            input.setAttribute(
                'data-form-type',
                'other'
            );

            input.removeAttribute('name');

        });
    }

    setInterval(removeManageButton, 300);
    </script>
    """,
    unsafe_allow_html=True,
)


excel_path_1 = "Πωλήσεις Ειδών Προσφορών από 29082026 - 29082026.xlsx"
excel_path_2 = "S3 - Πωλήσεις Ειδών-6.xlsx"

time_path = "upload_time.txt"
date_path = "upload_date.txt"

confetti_path = "confetti_status.txt"
cheer_path = "cheer_status.txt"


def upload_to_github(
    file_path,
    repo_name,
    token,
    commit_message="Update sales file"
):

    if not token or not repo_name:
        return False

    try:

        url = (
            f"https://api.github.com/repos/"
            f"{repo_name}/contents/{file_path}"
        )

        headers = {
            "Authorization": f"token {token}",
            "Accept": "application/vnd.github.v3+json",
        }

        r = requests.get(
            url,
            headers=headers
        )

        sha = None

        if r.status_code == 200:
            sha = r.json().get("sha")

        if not os.path.exists(file_path):
            return False

        with open(file_path, "rb") as f:
            content_bytes = f.read()

        content_encoded = (
            base64
            .b64encode(content_bytes)
            .decode("utf-8")
        )

        data = {
            "message": commit_message,
            "content": content_encoded
        }

        if sha:
            data["sha"] = sha

        put_r = requests.put(
            url,
            headers=headers,
            data=json.dumps(data)
        )

        return put_r.status_code in [200, 201]

    except Exception:
        return False


# ==================================================
# ΚΟΜΦΕΤΙ
# ==================================================

confetti_enabled = True

if os.path.exists(confetti_path):

    try:

        with open(
            confetti_path,
            "r",
            encoding="utf-8"
        ) as cf:

            confetti_enabled = (
                cf.read().strip() == "True"
            )

    except Exception:
        pass


# ==================================================
# ΧΕΙΡΟΚΡΟΤΗΜΑ
# ==================================================

cheer_enabled = True

if os.path.exists(cheer_path):

    try:

        with open(
            cheer_path,
            "r",
            encoding="utf-8"
        ) as ch:

            cheer_enabled = (
                ch.read().strip() == "True"
            )

    except Exception:
        pass


if "selected_region" not in st.session_state:

    st.session_state.selected_region = "τομεας 3"


st.markdown(
    """
    <style>

    .row-widget.stSelectbox {
        margin-bottom: 0px !important;
    }

    div[data-testid="stSelectbox"] {
        display: flex !important;
        flex-direction: row !important;
        align-items: center !important;
        gap: 15px !important;
    }

    div[data-testid="stSelectbox"] label {
        font-size: 13px !important;
        font-weight: 700 !important;
        color: #e74c3c !important;
        margin-bottom: 0px !important;
        white-space: nowrap !important;
        min-width: fit-content !important;
    }

    div[data-testid="stSelectbox"]
    div[data-baseweb="select"] {
        flex-grow: 1 !important;
    }

    div[data-testid="stDownloadButton"] > button {
        min-height: 38px !important;
        height: 38px !important;
        padding: 0 10px !important;
        font-size: 11px !important;
        font-weight: 800 !important;
        white-space: nowrap !important;
    }

    </style>
    """,
    unsafe_allow_html=True,
)


col_admin, col_input_space = st.columns([3, 5], gap="small")


# ==================================================
# ADMIN
# ==================================================

with col_admin:

    with st.expander(
        "⚙️ Διαχείριση Αρχείων (Admin 2026)"
    ):

        password = st.text_input(
            "Εισάγετε κωδικό διαχειριστή:",
            key="admin_pass",
            placeholder="Κωδικός",
            type="default",
        )

        components.html(
            """
            <script>

            const doc = window.parent.document;
            const inputs = doc.querySelectorAll('input');

            inputs.forEach(input => {

                if (
                    input.getAttribute('aria-label') &&
                    input
                    .getAttribute('aria-label')
                    .includes('κωδικό')
                ) {

                    input.setAttribute(
                        'autocomplete',
                        'off'
                    );

                    input.setAttribute(
                        'data-form-type',
                        'other'
                    );

                    input.removeAttribute('name');

                }

            });

            </script>
            """,
            height=0,
        )

        if password == "2845":

            st.markdown("---")

            col_up1, col_up2 = st.columns(2)

            with col_up1:

                uploaded_file_1 = st.file_uploader(
                    "Αρχείο 1 (Πωλήσεις Ειδών Προσφορών):",
                    type=["xlsx"],
                    key="up1",
                )

            with col_up2:

                uploaded_file_2 = st.file_uploader(
                    "Αρχείο 2 (S3 - Πωλήσεις Ειδών-6):",
                    type=["xlsx"],
                    key="up2",
                )

            st.markdown("---")

            time_options = []

            for hour in range(8, 23):

                for minute in (0, 30):

                    time_options.append(
                        datetime.time(
                            hour,
                            minute
                        )
                    )

            time_options.append(
                datetime.time(22, 0)
            )

            time_options = sorted(
                list(set(time_options))
            )

            now = (
                datetime.datetime.now()
                -
                datetime.timedelta(hours=1)
            )

            default_minute = (
                0
                if now.minute < 30
                else 30
            )

            default_hour = max(
                8,
                min(
                    22,
                    now.hour
                )
            )

            default_time = datetime.time(
                default_hour,
                default_minute
            )

            if (
                "selected_half_hour"
                not in st.session_state
            ):

                st.session_state[
                    "selected_half_hour"
                ] = default_time

            if (
                "selected_report_date"
                not in st.session_state
            ):

                st.session_state[
                    "selected_report_date"
                ] = datetime.date.today()

            selected_date = st.date_input(
                "Ημερομηνία αναφοράς:",
                value=st.session_state[
                    "selected_report_date"
                ]
            )

            st.session_state[
                "selected_report_date"
            ] = selected_date

            selected_time = st.selectbox(
                "Ώρα αναφοράς:",
                options=time_options,
                index=(
                    time_options.index(
                        st.session_state[
                            "selected_half_hour"
                        ]
                    )
                    if
                    st.session_state[
                        "selected_half_hour"
                    ]
                    in time_options
                    else 0
                ),
                format_func=lambda x: x.strftime("%H:%M"),
            )

            st.session_state[
                "selected_half_hour"
            ] = selected_time

            col_confetti, col_cheer = st.columns(2)

            with col_confetti:

                confetti_choice = st.radio(
                    "Κομφετί:",
                    ["ΝΑΙ", "ΟΧΙ"],
                    index=(
                        0
                        if confetti_enabled
                        else 1
                    ),
                    horizontal=True,
                    key="conf_radio",
                )

            with col_cheer:

                cheer_choice = st.radio(
                    "Χειροκρότημα:",
                    ["ΝΑΙ", "ΟΧΙ"],
                    index=(
                        0
                        if cheer_enabled
                        else 1
                    ),
                    horizontal=True,
                    key="cheer_radio",
                )

            if (
                uploaded_file_1 is not None
                and
                uploaded_file_2 is not None
            ):

                upload_signature = (
                    f"{uploaded_file_1.name}_"
                    f"{uploaded_file_2.name}_"
                    f"{uploaded_file_1.size}_"
                    f"{uploaded_file_2.size}"
                )

                gh_token = None
                repo_name = None

                try:

                    if (
                        hasattr(st, "secrets")
                        and
                        "GITHUB_TOKEN"
                        in st.secrets
                    ):

                        gh_token = (
                            st.secrets[
                                "GITHUB_TOKEN"
                            ]
                        )

                    if (
                        hasattr(st, "secrets")
                        and
                        "REPO_NAME"
                        in st.secrets
                    ):

                        repo_name = (
                            st.secrets[
                                "REPO_NAME"
                            ]
                        )

                except Exception:
                    pass

                current_time_str = (
                    selected_time.strftime(
                        "%H:%M"
                    )
                )

                current_date_str = (
                    selected_date.strftime(
                        "%d/%m/%Y"
                    )
                )

                with open(
                    time_path,
                    "w",
                    encoding="utf-8"
                ) as tf:

                    tf.write(
                        current_time_str
                    )

                with open(
                    date_path,
                    "w",
                    encoding="utf-8"
                ) as df_file:

                    df_file.write(
                        current_date_str
                    )

                with open(
                    confetti_path,
                    "w",
                    encoding="utf-8"
                ) as cf:

                    cf.write(
                        str(
                            confetti_choice
                            == "ΝΑΙ"
                        )
                    )

                with open(
                    cheer_path,
                    "w",
                    encoding="utf-8"
                ) as ch:

                    ch.write(
                        str(
                            cheer_choice
                            == "ΝΑΙ"
                        )
                    )

                with open(
                    excel_path_1,
                    "wb"
                ) as f:

                    f.write(
                        uploaded_file_1
                        .getbuffer()
                    )

                if gh_token and repo_name:

                    upload_to_github(
                        excel_path_1,
                        repo_name,
                        gh_token,
                        "Auto-update sales file 1",
                    )

                with open(
                    excel_path_2,
                    "wb"
                ) as f:

                    f.write(
                        uploaded_file_2
                        .getbuffer()
                    )

                if gh_token and repo_name:

                    upload_to_github(
                        excel_path_2,
                        repo_name,
                        gh_token,
                        "Auto-update sales file 2",
                    )

                if gh_token and repo_name:

                    upload_to_github(
                        time_path,
                        repo_name,
                        gh_token,
                        "Auto-update upload time",
                    )

                    upload_to_github(
                        date_path,
                        repo_name,
                        gh_token,
                        "Auto-update upload date",
                    )

                st.session_state[
                    "last_uploaded_sig"
                ] = upload_signature

                st.success(
                    "Και τα δύο αρχεία ανέβηκαν "
                    "αυτόματα και συγχρονίστηκαν "
                    "επιτυχώς!"
                )

                components.html(
                    """
                    <script>

                    setTimeout(
                        function() {

                            window.parent
                            .location
                            .reload();

                        },
                        1500
                    );

                    </script>
                    """,
                    height=0,
                )

        elif password:

            st.error("Λάθος κωδικός!")


# ==================================================
# ΠΕΡΙΦΕΡΕΙΑ
# ==================================================

with col_input_space:

    region_options = [
        "τομεας 3",
        "χαραλαμπιδης",
        "μητρουλης",
        "παππας",
        "πατσης",
        "πονοπουλος",
        "σκιαδοπουλος",
        "σαμογλου",
        "μπουτσκος",
        "πουτογλιδης",
    ]

    selected_region = st.selectbox(
        "📍 ΠΕΡΙΦΕΡΕΙΑ",
        options=region_options,
        index=0,
        format_func=lambda x: x.upper(),
    )

    active_filter = (
        selected_region.lower()
    )


# ==================================================
# ΩΡΑ
# ==================================================

file_time_str = "--:--"

if os.path.exists(time_path):

    try:

        with open(
            time_path,
            "r",
            encoding="utf-8"
        ) as tf:

            file_time_str = (
                tf.read().strip()
            )

    except Exception:
        pass


# ==================================================
# ΗΜΕΡΟΜΗΝΙΑ
# ==================================================

file_date_str = (
    datetime.date.today()
    .strftime("%d/%m/%Y")
)

if os.path.exists(date_path):

    try:

        with open(
            date_path,
            "r",
            encoding="utf-8"
        ) as df_file:

            file_date_str = (
                df_file.read().strip()
            )

    except Exception:
        pass


# ==================================================
# DATA
# ==================================================

def load_data(path):

    if os.path.exists(path):

        try:

            df = pd.read_excel(
                path,
                header=None
            )

            return df

        except Exception:

            return pd.DataFrame()

    return pd.DataFrame()


def clean_quantity_value(val):

    if pd.isna(val):
        return 0.0

    if isinstance(
        val,
        (int, float)
    ):
        return float(val)

    s_val = str(val).strip()

    if (
        "," in s_val
        and
        "." in s_val
    ):

        s_val = (
            s_val
            .replace(".", "")
            .replace(",", ".")
        )

    elif "," in s_val:

        s_val = (
            s_val
            .replace(",", ".")
        )

    try:

        return float(s_val)

    except Exception:

        return 0.0


def format_smart_num(num):

    if num == int(num):

        return (
            f"{int(num):,}"
            .replace(",", ".")
        )

    else:

        parts = (
            f"{num:.3f}"
            .split(".")
        )

        int_part = int(
            parts[0]
        )

        dec_part = (
            parts[1]
            .rstrip("0")
        )

        formatted_int = (
            f"{int_part:,}"
            .replace(",", ".")
        )

        return (
            f"{formatted_int},"
            f"{dec_part}"
        )


# ==================================================
# ΚΑΝΟΝΙΚΟΠΟΙΗΣΗ ΚΕΙΜΕΝΟΥ
# ΒΓΑΖΕΙ ΤΟΝΟΥΣ / ΠΕΖΑ-ΚΕΦΑΛΑΙΑ
# ==================================================

def normalize_text(text):

    text = str(text).strip().lower()

    text = unicodedata.normalize(
        "NFD",
        text
    )

    text = "".join(
        char
        for char in text
        if unicodedata.category(char) != "Mn"
    )

    return text


# ==================================================
# ΕΠΕΞΕΡΓΑΣΙΑ EXCEL
#
# Η ΣΤΗΛΗ ΠΟΣΟΤΗΤΑ ΕΝΤΟΠΙΖΕΤΑΙ
# ΑΠΟ ΤΟ ΟΝΟΜΑ ΤΗΣ ΚΑΙ ΟΧΙ ΑΠΟ ΤΗ ΘΕΣΗ ΤΗΣ
# ==================================================

def process_sales_df(
    df,
    file_name=""
):

    if df.empty:

        return (
            "ΕΙΔΟΣ",
            pd.DataFrame(),
            0.0,
            1.0
        )

    custom_title = "ΕΙΔΟΣ"

    # ==================================================
    # ΕΝΤΟΠΙΣΜΟΣ ΤΙΤΛΟΥ
    # ==================================================

    for i in range(
        min(
            5,
            len(df)
        )
    ):

        for j in range(
            len(df.columns)
        ):

            val = str(
                df.iloc[i, j]
            ).strip()

            normalized_val = (
                normalize_text(val)
            )

            if (
                val
                and
                normalized_val != "nan"
                and
                "καταστημα"
                not in normalized_val
                and
                "πληρωτ"
                not in normalized_val
                and
                "ποσοτ"
                not in normalized_val
                and
                "αξια"
                not in normalized_val
                and
                "κοστος"
                not in normalized_val
            ):

                custom_title = val
                break

        if custom_title != "ΕΙΔΟΣ":
            break

    # ==================================================
    # ΒΡΙΣΚΕΙ ΤΗ ΓΡΑΜΜΗ ΕΠΙΚΕΦΑΛΙΔΩΝ
    # ΨΑΧΝΕΙ ΣΕ ΟΛΟ ΤΟ EXCEL
    # ==================================================

    header_row_idx = None

    for i in range(len(df)):

        row_values = [
            normalize_text(value)
            for value in df.iloc[i].values
        ]

        has_store = any(
            "καταστημα" in value
            for value in row_values
        )

        has_quantity = any(
            (
                "ποσοτητα" in value
                or
                "ποσοτ" in value
            )
            for value in row_values
        )

        if (
            has_store
            and
            has_quantity
        ):

            header_row_idx = i
            break

    if header_row_idx is None:

        st.error(
            f"Δεν βρέθηκε γραμμή με "
            f"ΚΑΤΑΣΤΗΜΑ και ΠΟΣΟΤΗΤΑ "
            f"στο αρχείο {file_name}"
        )

        return (
            custom_title,
            pd.DataFrame(),
            0.0,
            1.0
        )

    # ==================================================
    # ΟΡΙΣΜΟΣ ΕΠΙΚΕΦΑΛΙΔΩΝ
    # ==================================================

    df.columns = (
        df.iloc[
            header_row_idx
        ]
        .astype(str)
        .str.strip()
    )

    df = (
        df.iloc[
            header_row_idx + 1:
        ]
        .reset_index(
            drop=True
        )
    )

    # ==================================================
    # ΒΡΙΣΚΕΙ ΤΗ ΣΤΗΛΗ ΚΑΤΑΣΤΗΜΑ
    # ΟΠΟΥ ΚΑΙ ΑΝ ΒΡΙΣΚΕΤΑΙ
    # ==================================================

    store_col = None

    for col in df.columns:

        normalized_col = (
            normalize_text(col)
        )

        if "καταστημα" in normalized_col:

            store_col = col
            break

    # ==================================================
    # ΒΡΙΣΚΕΙ ΤΗ ΣΤΗΛΗ ΠΟΣΟΤΗΤΑ
    # ΟΠΟΥ ΚΑΙ ΑΝ ΒΡΙΣΚΕΤΑΙ
    # ==================================================

    quantity_col = None

    # Πρώτα ψάχνουμε ακριβώς "ΠΟΣΟΤΗΤΑ"
    for col in df.columns:

        normalized_col = (
            normalize_text(col)
        )

        if normalized_col == "ποσοτητα":

            quantity_col = col
            break

    # Μετά παραλλαγές όπως
    # "ΠΟΣΟΤΗΤΑ ΠΩΛΗΣΕΩΝ"
    if quantity_col is None:

        for col in df.columns:

            normalized_col = (
                normalize_text(col)
            )

            if "ποσοτητα" in normalized_col:

                quantity_col = col
                break

    # Και τέλος πιθανό "ΠΟΣΟΤ."
    if quantity_col is None:

        for col in df.columns:

            normalized_col = (
                normalize_text(col)
            )

            if "ποσοτ" in normalized_col:

                quantity_col = col
                break

    if store_col is None:

        st.error(
            f"Δεν βρέθηκε η στήλη "
            f"ΚΑΤΑΣΤΗΜΑ "
            f"στο αρχείο {file_name}"
        )

        return (
            custom_title,
            pd.DataFrame(),
            0.0,
            1.0
        )

    if quantity_col is None:

        st.error(
            f"Δεν βρέθηκε η στήλη "
            f"ΠΟΣΟΤΗΤΑ "
            f"στο αρχείο {file_name}"
        )

        return (
            custom_title,
            pd.DataFrame(),
            0.0,
            1.0
        )

    # ==================================================
    # ΚΡΑΤΑΜΕ ΜΟΝΟ:
    # ΚΑΤΑΣΤΗΜΑ + ΠΟΣΟΤΗΤΑ
    # ==================================================

    df_selected = (
        df[
            [
                store_col,
                quantity_col
            ]
        ]
        .copy()
    )

    df_selected.columns = [
        "Κατάστημα",
        "Ποσότητα"
    ]

    df_selected = (
        df_selected
        .dropna(
            subset=[
                "Κατάστημα",
                "Ποσότητα"
            ]
        )
    )

    df_selected[
        "Κατάστημα"
    ] = (
        df_selected[
            "Κατάστημα"
        ]
        .astype(str)
        .str.strip()
    )

    df_selected = (
        df_selected[
            ~df_selected[
                "Κατάστημα"
            ]
            .str.contains(
                "Κατάστημα|ΠΟΣΟΤ|ΠΑΡΑΔΕΙΓΜΑ|NaN",
                case=False,
                na=False
            )
        ]
    )

    df_clean = (
        df_selected[
            ~df_selected[
                "Κατάστημα"
            ]
            .str.contains(
                "Total|Συνολο|ΣΥΝΟΛΟ",
                case=False,
                na=False
            )
        ]
        .copy()
    )

    # ==================================================
    # ΜΕΤΑΤΡΟΠΗ ΠΟΣΟΤΗΤΑΣ ΣΕ ΑΡΙΘΜΟ
    # ==================================================

    df_clean[
        "Num_Sales"
    ] = (
        df_clean[
            "Ποσότητα"
        ]
        .apply(
            clean_quantity_value
        )
    )

    # ==================================================
    # ΤΑΞΙΝΟΜΗΣΗ
    # ==================================================

    df_stores = (
        df_clean
        .sort_values(
            by="Num_Sales",
            ascending=False
        )
        .reset_index(
            drop=True
        )
    )

    total_sum = (
        df_stores[
            "Num_Sales"
        ]
        .sum()
    )

    max_sales = (
        df_stores[
            "Num_Sales"
        ]
        .max()
        if
        not df_stores.empty
        else
        1.0
    )

    return (
        custom_title,
        df_stores,
        total_sum,
        max_sales
    )


# ==================================================
# ΟΜΑΔΕΣ
# ==================================================

MITROULIS_KEYWORDS = [
    "301", "302", "309", "486", "304",
    "352", "308", "353", "354", "355",
    "356", "366", "374",
    "αιανη",
    "πλ.ελευθεριας",
    "πλ.λασσανη",
    "σιατιστα",
    "25ης μαρτιου",
    "ελ.βενιζελου",
    "ιωαννη αρθη",
    "κοζανης και γρεβενων",
    "χλοη",
    "δισπυλο",
    "αθ. διακου",
    "γραμμου",
    "μανιακοι"
]


PONOPOULOS_KEYWORDS = [
    "231", "232", "233", "373",
    "237", "235", "372", "236",
    "483", "161", "166", "384",
    "234",
    "σβορωνου",
    "υψηλαντου",
    "περδικα",
    "εγνατιας",
    "πλαταμωνας",
    "λεπτοκαρυα",
    "κορινος",
    "αγ.νικολαιου",
    "λιτοχωρο",
    "αντιγονου",
    "κατερινη",
    "αριστοτελους",
    "π.τσαλδαρη",
    "19ης μαΐου",
    "χατζογλου"
]


CHARALAMPIDIS_KEYWORDS = [
    "211", "201", "347", "212",
    "239", "219", "220", "222",
    "223", "240", "241", "493",
    "δ.γεωργιαδου",
    "λαρισα",
    "νικηταρα",
    "ιωαννινων",
    "23ης οκτωβριου",
    "χατζημιχαλη",
    "φιλιππουπολη",
    "θυατειρων",
    "βενιζελου",
    "ν.ιωνια",
    "βολος",
    "αχιλλοπουλου",
    "κασσαβετη",
    "28ης οκτωβριου",
    "κουμουνδουρου",
    "μεταμορφωσεως",
    "αλεξανδρας",
    "σκιαθος"
]


PAPPAS_KEYWORDS = [
    "210", "346", "202", "204",
    "206", "209", "205", "207",
    "208", "215",
    "ελασσονα",
    "βυζαντιου",
    "λαρισης",
    "φαρσαλα",
    "αβερωφ",
    "καρδιτσα",
    "καραϊσκακη",
    "σοφαδες",
    "κονδυλη",
    "τρικαλα",
    "δεληγιωργη",
    "ελευθεριος"
]


PATSIS_KEYWORDS = [
    "198", "225", "226", "316",
    "317", "381", "228", "229",
    "224", "315", "359", "378",
    "399", "444",
    "γκουρας",
    "νικοπολεως",
    "ιωαννινα",
    "γ.παπανδρεου",
    "κατω νεοχωροπουλο",
    "μαρικας κοτοπουλη",
    "ριζαριο",
    "λεωφ. δημοκρατιας",
    "καρδαμιτσια",
    "κοραη",
    "κ.παλαιολογου",
    "ανατολη",
    "καρυωτακη",
    "λεωφ. ειρηνης",
    "πρεβεζα",
    "πλ. κιλκις",
    "ανεξαρτησιας",
    "αρτα",
    "26ο χλμ",
    "λουρος",
    "ηγουμενιτσα",
    "θεσπρωτιας",
    "παραμυθια",
    "αγ.μαρινας"
]


SKIADOPOULOS_KEYWORDS = [
    "531", "539", "567", "537",
    "533", "525", "535", "534",
    "566", "540", "565", "530",
    "528", "529", "532", "536",
    "538", "549",
    "κερκυρα",
    "αχαραβη",
    "κασσιωπη",
    "σιδαρι",
    "καρουσαδες",
    "μαρκατο",
    "μαντουκι",
    "αλυκες",
    "υπερ εθνικη οδος λευκιμμης",
    "λευκιμμη",
    "μωραιτικα",
    "κομβος βρυωνη",
    "καστελλοι",
    "αλεπου",
    "σαροκο",
    "ιωαννου θεοτοκη",
    "παλλαδα",
    "λαικη αγορα",
    "γερασιμου",
    "πινια",
    "νοσοκομειο",
    "σπηλια",
    "μητροπολιτου μεθοδιου"
]


SAMOGLOU_KEYWORDS = [
    "8907", "8912", "8913",
    "8918", "8920", "8926",
    "8932", "8933", "8938",
    "8939", "8941", "8943",
    "καστορια χιλιοδενδρο",
    "καστορια αγ μηνα",
    "καστορια μ αλεξανδρου",
    "κοζανη παυλου χαριση",
    "γρεβενα μακεδονομαχων",
    "σερβια κοζανης κων καρπου",
    "βελβεντος κοζανης",
    "κοζανη οσε",
    "γρεβενα θεωδ ζιακα",
    "γρεβενα καβαφη",
    "κοζανη φιλιππου"
]


BOUTSKOS_KEYWORDS = [
    "8267", "8273", "8317",
    "8318", "8334", "8402",
    "8408", "8448", "8568",
    "8569", "8591",
    "δομοκος παπαφλεσσα",
    "πλαταμωνας κων καραμανλη",
    "αλμυρος βολου ν μιχοπουλου",
    "θεσ σινδος παλαιολογου",
    "θεσ ολυμπου",
    "λαρισα στρατηγου φραγκου",
    "σκοπελος 2 χλμ επο σκοπελου",
    "λαρισα ιωαννινων",
    "βολος γιαννη δημου συ",
    "στεφανοβικειο βελεστινο συ",
    "αγια λαρισης"
]


POUTOGLIDIS_KEYWORDS = [
    "8333", "8916", "8917",
    "8923", "8925", "8928",
    "8931", "8936", "8937",
    "8940", "8945",
    "πτολεμαιδα 25 μαρτιου",
    "φλωρινα κρεσνας",
    "φλωρινα καστρισιανακη",
    "φλωρινα δημ παπαθανασιου",
    "κροκος κοζανης ιοακ λιουλια",
    "φιλιωτας μ αλεξανδρου",
    "αμυνταιο 28 συνταγμπ πεζικου",
    "πτολεμαιδα χρυσ σμυρνης",
    "πτολεμαιδα θεολογιδη",
    "φλωρινα cash & carry",
    "κοζανη κροκος cash & carry"
]


TOMEAS_3_KEYWORDS = (
    MITROULIS_KEYWORDS
    + PONOPOULOS_KEYWORDS
    + CHARALAMPIDIS_KEYWORDS
    + PAPPAS_KEYWORDS
    + PATSIS_KEYWORDS
    + SKIADOPOULOS_KEYWORDS
    + SAMOGLOU_KEYWORDS
    + BOUTSKOS_KEYWORDS
    + POUTOGLIDIS_KEYWORDS
)


GROUPS_MAPPING = {

    "τομεας 3":
        TOMEAS_3_KEYWORDS,

    "τομέας 3":
        TOMEAS_3_KEYWORDS,

    "μητρουλης":
        MITROULIS_KEYWORDS,

    "πονοπουλος":
        PONOPOULOS_KEYWORDS,

    "χαραλαμπιδης":
        CHARALAMPIDIS_KEYWORDS,

    "παππας":
        PAPPAS_KEYWORDS,

    "πατσης":
        PATSIS_KEYWORDS,

    "σκιαδοπουλος":
        SKIADOPOULOS_KEYWORDS,

    "σαμογλου":
        SAMOGLOU_KEYWORDS,

    "μπουτσκος":
        BOUTSKOS_KEYWORDS,

    "πουτογλιδης":
        POUTOGLIDIS_KEYWORDS,
}


# ==================================================
# ΦΟΡΤΩΣΗ
# ==================================================

title_1, df_stores_1, _, max_sales_1 = (
    process_sales_df(
        load_data(excel_path_1),
        file_name=excel_path_1
    )
)


title_2, df_stores_2, _, max_sales_2 = (
    process_sales_df(
        load_data(excel_path_2),
        file_name=excel_path_2
    )
)


def filter_dataframe(df_stores):

    if df_stores.empty:

        return (
            df_stores,
            0.0
        )

    filtered_df = (
        df_stores.copy()
    )

    if active_filter:

        if active_filter in GROUPS_MAPPING:

            keywords = (
                GROUPS_MAPPING[
                    active_filter
                ]
            )

            pattern = "|".join(
                [
                    r"\b"
                    + kw
                    + r"\b"
                    for kw
                    in keywords
                ]
            )

            filtered_df = (
                filtered_df[
                    filtered_df[
                        "Κατάστημα"
                    ]
                    .str.lower()
                    .str.contains(
                        pattern,
                        case=False,
                        na=False,
                        regex=True
                    )
                ]
            )

        else:

            filtered_df = (
                filtered_df[
                    filtered_df[
                        "Κατάστημα"
                    ]
                    .str.lower()
                    .str.contains(
                        active_filter,
                        na=False
                    )
                ]
            )

    total_sum = (
        filtered_df[
            "Num_Sales"
        ]
        .sum()
    )

    return (
        filtered_df
        .reset_index(drop=True),
        total_sum
    )


df_stores_1, total_sum_1 = (
    filter_dataframe(
        df_stores_1
    )
)


df_stores_2, total_sum_2 = (
    filter_dataframe(
        df_stores_2
    )
)


# ==================================================
# ΕΞΑΓΩΓΗ ΤΡΕΧΟΥΣΑΣ ΠΛΗΡΟΦΟΡΙΑΣ ΣΕ EXCEL
# ΜΙΚΡΟ ΚΟΥΜΠΙ ΔΙΠΛΑ ΣΤΗΝ ΠΕΡΙΦΕΡΕΙΑ
# ==================================================

def prepare_export_df(df_stores, total_sum):

    if df_stores.empty:

        export_df = pd.DataFrame(
            columns=[
                "Κατάστημα",
                "Ποσότητα"
            ]
        )

    else:

        export_df = (
            df_stores[
                [
                    "Κατάστημα",
                    "Num_Sales"
                ]
            ]
            .copy()
        )

        export_df.columns = [
            "Κατάστημα",
            "Ποσότητα"
        ]

    total_row = pd.DataFrame(
        [
            {
                "Κατάστημα": "TOTAL",
                "Ποσότητα": total_sum,
            }
        ]
    )

    return pd.concat(
        [
            export_df,
            total_row
        ],
        ignore_index=True
    )


def safe_excel_sheet_name(name, fallback):

    invalid_chars = ['\\', '/', '*', '?', ':', '[', ']']

    safe_name = str(name).strip()

    for char in invalid_chars:
        safe_name = safe_name.replace(char, " ")

    safe_name = " ".join(safe_name.split())

    if not safe_name:
        safe_name = fallback

    return safe_name[:31]


def add_report_info_after_total(export_df):

    info_rows = pd.DataFrame(
        [
            {
                "Κατάστημα": "",
                "Ποσότητα": ""
            },
            {
                "Κατάστημα": "ΠΕΡΙΦΕΡΕΙΑ",
                "Ποσότητα": selected_region.upper()
            },
            {
                "Κατάστημα": "ΗΜΕΡΟΜΗΝΙΑ",
                "Ποσότητα": file_date_str
            },
            {
                "Κατάστημα": "ΩΡΑ",
                "Ποσότητα": file_time_str
            },
        ]
    )

    return pd.concat(
        [
            export_df,
            info_rows
        ],
        ignore_index=True
    )


def build_excel_file():

    output = BytesIO()

    export_1 = prepare_export_df(
        df_stores_1,
        total_sum_1
    )

    export_2 = prepare_export_df(
        df_stores_2,
        total_sum_2
    )

    # Μετά το TOTAL:
    # ΠΕΡΙΦΕΡΕΙΑ - ΗΜΕΡΟΜΗΝΙΑ - ΩΡΑ
    export_1 = add_report_info_after_total(export_1)
    export_2 = add_report_info_after_total(export_2)

    # Τα φύλλα παίρνουν τα πραγματικά ονόματα των δύο ειδών.
    sheet_1 = safe_excel_sheet_name(
        title_1,
        "Είδος 1"
    )

    sheet_2 = safe_excel_sheet_name(
        title_2,
        "Είδος 2"
    )

    if sheet_2 == sheet_1:
        base = sheet_2[:27]
        sheet_2 = f"{base} (2)"

    with pd.ExcelWriter(
        output,
        engine="openpyxl"
    ) as writer:

        # ΜΟΝΟ δύο φύλλα:
        # [Όνομα 1ου είδους] | [Όνομα 2ου είδους]
        export_1.to_excel(
            writer,
            sheet_name=sheet_1,
            index=False
        )

        export_2.to_excel(
            writer,
            sheet_name=sheet_2,
            index=False
        )

        for sheet_name in [
            sheet_1,
            sheet_2
        ]:

            ws = writer.book[sheet_name]

            for column_cells in ws.columns:

                max_length = 0
                column_letter = column_cells[0].column_letter

                for cell in column_cells:

                    cell_value = (
                        ""
                        if cell.value is None
                        else str(cell.value)
                    )

                    max_length = max(
                        max_length,
                        len(cell_value)
                    )

                ws.column_dimensions[
                    column_letter
                ].width = min(
                    max_length + 3,
                    45
                )

    output.seek(0)

    return output.getvalue()


excel_data = None
excel_download_filename = "sales.xlsx"

try:

    excel_data = build_excel_file()

    safe_region_name = (
        normalize_text(selected_region)
        .replace(" ", "_")
    )

    excel_download_filename = (
        f"sales_{safe_region_name}_"
        f"{datetime.date.today().strftime('%Y%m%d')}.xlsx"
    )

except Exception:
    excel_data = None


# ==================================================
# BANNER - ΝΕΟ / 5 ΣΕΠΤΕΜΒΡΙΟΥ
# ==================================================

img_src = ""

# Ψάχνει εικόνες/GIF που μπορούν να χρησιμοποιηθούν ως banner.
# Αν υπάρχουν περισσότερα από ένα, επιλέγει αυτό που τροποποιήθηκε πιο πρόσφατα.
banner_files = (
    glob.glob("ChatGPT Image*.png")
    + glob.glob("ChatGPT Image*.jpg")
    + glob.glob("ChatGPT Image*.jpeg")
    + glob.glob("ChatGPT Image*.gif")
    + glob.glob("*banner*.jpg")
    + glob.glob("*banner*.jpeg")
    + glob.glob("*banner*.png")
    + glob.glob("*banner*.gif")
    + glob.glob("*ΣΦΥΡΙ*.gif")
    + glob.glob("*σφυρι*.gif")
    + glob.glob("*sfyri*.gif")
    + glob.glob("*hammer*.gif")
)

# Αφαιρούμε τυχόν διπλοεγγραφές
banner_files = list(dict.fromkeys(banner_files))

if banner_files:

    # Πάντα το νεότερο αρχείο
    banner_filename = max(
        banner_files,
        key=os.path.getmtime
    )

    extension = os.path.splitext(
        banner_filename
    )[1].lower()

    mime_types = {
        ".png": "image/png",
        ".jpg": "image/jpeg",
        ".jpeg": "image/jpeg",
        ".gif": "image/gif",
    }

    mime_type = mime_types.get(
        extension,
        "image/png"
    )

    with open(
        banner_filename,
        "rb"
    ) as image_file:

        img_src = (
            f"data:{mime_type};base64,"
            f"{base64.b64encode(image_file.read()).decode()}"
        )


# ==================================================
# HTML
# ==================================================

try:

    html_content = f"""

    <script
        src="https://cdn.jsdelivr.net/npm/canvas-confetti@1.5.1/dist/confetti.browser.min.js">
    </script>

    <link
        href="https://fonts.googleapis.com/css2?family=Montserrat:wght@500;600;700;800&display=swap"
        rel="stylesheet"
    >

    <style>


    @keyframes blink-number-slow {{

        0% {{
            opacity: 1;
            color: #2ecc71;
            text-shadow:
                0 0 12px
                rgba(46, 204, 113, 0.7);
        }}

        50% {{
            opacity: 0.25;
            color: #27ae60;
            text-shadow: none;
        }}

        100% {{
            opacity: 1;
            color: #2ecc71;
            text-shadow:
                0 0 12px
                rgba(46, 204, 113, 0.7);
        }}

    }}


    @keyframes rotate-phone-smooth {{

        0% {{
            transform:
                rotate(0deg)
                scale(1);
        }}

        35% {{
            transform:
                rotate(-90deg)
                scale(1.15);
        }}

        65% {{
            transform:
                rotate(-90deg)
                scale(1.15);
        }}

        100% {{
            transform:
                rotate(0deg)
                scale(1);
        }}

    }}


    body {{

        font-family:
            'Montserrat',
            sans-serif;

        margin: 0;
        padding: 0;
        background: transparent;
        width: 100%;
        overflow-x: hidden;

    }}


    .main-container {{

        position: relative;

        background:
            rgba(0, 0, 0, 0.6);

        padding: 0;

        border-radius: 0;

        box-shadow: none;

        backdrop-filter:
            blur(8px);

        -webkit-backdrop-filter:
            blur(8px);

        width: 100%;

        max-width: 100%;

        margin: 0 auto;

        text-align: center;

        overflow: hidden;

    }}


    .banner-container {{

        position: relative;
        width: 100%;

    }}


    .banner-img {{

        width: 100%;
        height: auto;
        display: block;
        border-radius: 0;
        margin: 0;
        padding: 0;

    }}


    .rotate-hint-overlay {{

        position: absolute;

        top: 4px;
        right: 12px;

        display: flex;

        align-items: center;

        gap: 5px;

        background: transparent;

        padding: 0;

        margin: 0;

        z-index: 5;

    }}


    .phone-icon-wrap {{

        display: inline-block;

        font-size: 20px;

        transform-origin: center;

        animation:
            rotate-phone-smooth
            3.5s
            infinite
            ease-in-out;

        filter:
            drop-shadow(
                0 2px 4px
                rgba(0, 0, 0, 0.8)
            );

    }}


    .turn-mobile-text {{

        font-size: 10px;

        color: #ffffff;

        text-transform: uppercase;

        font-weight: 800;

        letter-spacing: 0.5px;

        white-space: nowrap;

        text-shadow:
            0 2px 4px
            rgba(0, 0, 0, 0.9);

    }}


    @media
    (orientation: landscape) {{

        .rotate-hint-overlay {{

            display:
                none !important;

        }}

    }}


    .content-wrapper {{

        position: relative;

        padding: 25px;

    }}


    .header-area {{

        display: flex;

        justify-content:
            space-between;

        align-items: center;

        margin-bottom: 20px;

    }}


    .top-left-area {{

        text-align: left;

    }}


    .top-left-text {{

        color: #3498db;

        font-size: 13px;

        font-weight: 700;

        letter-spacing: 1px;

        text-transform: uppercase;

        margin-bottom: 2px;

    }}


    .top-left-subtext {{

        color: #2ecc71;

        font-size: 11px;

        font-weight: 800;

        letter-spacing: 1px;

        text-transform: uppercase;

        margin-bottom: 3px;

    }}


    .top-left-date {{

        color: #bdc3c7;

        font-size: 11px;

        font-weight: 600;

        letter-spacing: 0.5px;

        margin-top: 2px;

    }}


    .top-left-time {{

        color: #95a5a6;

        font-size: 11px;

        font-weight: 600;

        letter-spacing: 0.5px;

        margin-top: 2px;

    }}


    .columns-container {{

        display: grid;

        grid-template-columns:
            repeat(
                auto-fit,
                minmax(320px, 1fr)
            );

        gap: 20px;

        width: 100%;

    }}


    .product-column {{

        width: 100%;

    }}


    .sub-title {{

        color: #3498db;

        font-size: 18px;

        margin-bottom: 15px;

        font-weight: 700;

        text-transform: uppercase;

        letter-spacing: 1px;

        text-align: center;

    }}


    .poll-item {{

        background:
            rgba(255, 255, 255, 0.08);

        padding:
            12px 18px;

        border-radius:
            12px;

        margin-bottom:
            12px;

        text-align:
            left;

        border:
            1px solid
            rgba(255, 255, 255, 0.1);

    }}


    .poll-info {{

        display: flex;

        justify-content:
            space-between;

        align-items: center;

        color: white;

        font-size: 14px;

        font-weight: 600;

        margin-bottom: 8px;

        gap: 8px;

    }}


    .poll-info
    span:first-child {{

        overflow: hidden;

        text-overflow:
            ellipsis;

        flex: 1;

        min-width: 0;

    }}


    .poll-info
    span:last-child {{

        white-space:
            nowrap;

        text-align: right;

        flex-shrink: 0;

        min-width: 100px;

    }}


    .win-number-first {{

        color: #2ecc71;

        animation:
            blink-number-slow
            2.5s
            infinite
            ease-in-out;

        font-weight: 700;

    }}


    .progress-bar-bg {{

        background:
            rgba(255, 255, 255, 0.15);

        border-radius: 10px;

        height: 12px;

        width: 100%;

        overflow: hidden;

    }}


    .progress-fill {{

        background: #3498db;

        height: 100%;

        border-radius: 10px;

    }}


    .total-item {{

        background:
            rgba(52, 152, 219, 0.25);

        border:
            1px solid #3498db;

    }}


    .footer-tools {{

        display: flex;
        align-items: center;
        justify-content: flex-start;
        gap: 10px;
        margin-top: 15px;
        width: 100%;

    }}


    .watermark {{

        text-align: left;

        color:
            rgba(255, 255, 255, 0.16);

        font-size: 12px;

        font-weight: 600;

        letter-spacing: 1px;

        margin: 0;

        text-transform: none;

        user-select: none;

    }}


    .excel-bottom-btn {{

        display: inline-flex;
        align-items: center;
        justify-content: center;
        min-height: 30px;
        padding: 0 10px;
        border-radius: 8px;
        border: 1px solid rgba(255,255,255,0.22);
        background: rgba(255,255,255,0.10);
        color: #ffffff;
        font-family: 'Montserrat', sans-serif;
        font-size: 10px;
        font-weight: 800;
        text-decoration: none;
        white-space: nowrap;

    }}


    .excel-bottom-btn:hover {{

        background: rgba(255,255,255,0.18);

    }}


    </style>


    <div class="main-container">


        <div class="banner-container">


            <img
                src="{img_src}"
                class="banner-img"
                alt="banner"
            >


        </div>


        <div class="content-wrapper">


            <div
                class="rotate-hint-overlay"
            >

                <span
                    class="phone-icon-wrap"
                >
                    📱
                </span>

                <span
                    class="turn-mobile-text"
                >
                    TURN MOBILE
                </span>

            </div>


            <audio
                id="cheerAudio"
                preload="auto"
            >

                <source
                    src="https://www.myinstants.com/media/sounds/applause.mp3"
                    type="audio/mpeg"
                >

            </audio>


            <div
                class="header-area"
            >


                <div
                    class="top-left-area"
                >


                    <div
                        class="top-left-text"
                    >
                        ΤΟΜΕΑΣ 3
                    </div>


                    <div
                        class="top-left-subtext"
                    >
                        UPDATE SALES
                    </div>


                    <div
                        class="top-left-date"
                    >
                        {file_date_str}
                    </div>


                    <div
                        class="top-left-time"
                    >
                        εως: {file_time_str}
                    </div>


                </div>


            </div>


            <div
                class="columns-container"
            >

    """


    # ==================================================
    # ΣΤΗΛΗ 1
    # ==================================================

    html_content += (
        '<div class="product-column">'
    )


    html_content += (
        f'<div class="sub-title">'
        f'{title_1}'
        f'</div>'
    )


    if not df_stores_1.empty:


        for index, row in (
            df_stores_1.iterrows()
        ):


            katastima = str(
                row["Κατάστημα"]
            )


            if (
                katastima.lower() == "nan"
                or
                not katastima.strip()
            ):

                continue


            num = row["Num_Sales"]


            formatted_num = (
                format_smart_num(num)
            )


            bar_width = (

                round(
                    (
                        num
                        /
                        max_sales_1
                    )
                    *
                    100
                )

                if max_sales_1 > 0

                else 0

            )


            if bar_width > 100:
                bar_width = 100


            if index == 0:


                html_content += f"""

                <div class="poll-item">

                    <div class="poll-info">

                        <span>
                            <b>
                                {katastima}
                            </b>
                        </span>

                        <span
                            class="win-number-first"
                        >
                            {formatted_num} τμχ/κιλ
                        </span>

                    </div>


                    <div
                        class="progress-bar-bg"
                    >

                        <div
                            class="progress-fill"
                            style="width: {bar_width}%;"
                        >
                        </div>

                    </div>

                </div>

                """


            else:


                html_content += f"""

                <div class="poll-item">

                    <div class="poll-info">

                        <span>
                            <b>
                                {katastima}
                            </b>
                        </span>

                        <span>
                            <b>
                                {formatted_num} τμχ/κιλ
                            </b>
                        </span>

                    </div>


                    <div
                        class="progress-bar-bg"
                    >

                        <div
                            class="progress-fill"
                            style="width: {bar_width}%;"
                        >
                        </div>

                    </div>

                </div>

                """


        formatted_total_1 = (
            format_smart_num(
                total_sum_1
            )
        )


        html_content += f"""

        <div
            class="poll-item total-item"
        >

            <div
                class="poll-info"
            >

                <span>
                    <b>
                        TOTAL
                    </b>
                </span>

                <span>
                    <b>
                        {formatted_total_1} τμχ/κιλ
                    </b>
                </span>

            </div>


            <div
                class="progress-bar-bg"
            >

                <div
                    class="progress-fill"
                    style="width: 100%;"
                >
                </div>

            </div>

        </div>

        """


    else:


        html_content += (
            '<div '
            'style="color: white; padding: 20px;">'
            'Δεν βρέθηκαν δεδομένα.'
            '</div>'
        )


    html_content += "</div>"


    # ==================================================
    # ΣΤΗΛΗ 2
    # ==================================================

    html_content += (
        '<div class="product-column">'
    )


    html_content += (
        f'<div class="sub-title">'
        f'{title_2}'
        f'</div>'
    )


    if not df_stores_2.empty:


        for index, row in (
            df_stores_2.iterrows()
        ):


            katastima = str(
                row["Κατάστημα"]
            )


            if (
                katastima.lower() == "nan"
                or
                not katastima.strip()
            ):

                continue


            num = row["Num_Sales"]


            formatted_num = (
                format_smart_num(num)
            )


            bar_width = (

                round(
                    (
                        num
                        /
                        max_sales_2
                    )
                    *
                    100
                )

                if max_sales_2 > 0

                else 0

            )


            if bar_width > 100:
                bar_width = 100


            if index == 0:


                html_content += f"""

                <div class="poll-item">

                    <div class="poll-info">

                        <span>
                            <b>
                                {katastima}
                            </b>
                        </span>

                        <span
                            class="win-number-first"
                        >
                            {formatted_num} τμχ/κιλ
                        </span>

                    </div>


                    <div
                        class="progress-bar-bg"
                    >

                        <div
                            class="progress-fill"
                            style="width: {bar_width}%;"
                        >
                        </div>

                    </div>

                </div>

                """


            else:


                html_content += f"""

                <div class="poll-item">

                    <div class="poll-info">

                        <span>
                            <b>
                                {katastima}
                            </b>
                        </span>

                        <span>
                            <b>
                                {formatted_num} τμχ/κιλ
                            </b>
                        </span>

                    </div>


                    <div
                        class="progress-bar-bg"
                    >

                        <div
                            class="progress-fill"
                            style="width: {bar_width}%;"
                        >
                        </div>

                    </div>

                </div>

                """


        formatted_total_2 = (
            format_smart_num(
                total_sum_2
            )
        )


        html_content += f"""

        <div
            class="poll-item total-item"
        >

            <div
                class="poll-info"
            >

                <span>
                    <b>
                        TOTAL
                    </b>
                </span>

                <span>
                    <b>
                        {formatted_total_2} τμχ/κιλ
                    </b>
                </span>

            </div>


            <div
                class="progress-bar-bg"
            >

                <div
                    class="progress-fill"
                    style="width: 100%;"
                >
                </div>

            </div>

        </div>

        """


    else:


        html_content += (
            '<div '
            'style="color: white; padding: 20px;">'
            'Δεν βρέθηκαν δεδομένα.'
            '</div>'
        )


    html_content += "</div>"


    html_content += "</div>"


    # ==================================================
    # ΚΟΜΦΕΤΙ - ΜΟΝΟ ΤΟΜΕΑΣ 3
    # ==================================================

    if (
        confetti_enabled
        and
        active_filter
        in [
            "τομεας 3",
            "τομέας 3"
        ]
    ):

        html_content += """

        <script>

        setTimeout(
            function() {

                confetti({

                    particleCount: 90,

                    spread: 90,

                    origin: {
                        x: 0.5,
                        y: 0.25
                    }

                });


                setTimeout(
                    function() {

                        confetti({

                            particleCount: 110,

                            spread: 110,

                            origin: {
                                x: 0.5,
                                y: 0.25
                            }

                        });

                    },
                    3000
                );

            },
            300
        );

        </script>

        """


    # ==================================================
    # ΧΕΙΡΟΚΡΟΤΗΜΑ - ΜΟΝΟ ΤΟΜΕΑΣ 3
    # ==================================================

    if (
        cheer_enabled
        and
        active_filter
        in [
            "τομεας 3",
            "τομέας 3"
        ]
    ):

        html_content += """

        <script>

        function playCheer() {

            const audio =
                document.getElementById(
                    'cheerAudio'
                );


            if(audio) {

                audio.volume = 0.5;


                audio.play()

                .then(() => {

                    window.removeEventListener(
                        'click',
                        playCheer
                    );

                    window.removeEventListener(
                        'touchstart',
                        playCheer
                    );

                })

                .catch(
                    function(error) {

                        console.log(
                            "Audio play blocked"
                        );

                    }
                );

            }

        }


        window.addEventListener(
            'click',
            playCheer
        );


        window.addEventListener(
            'touchstart',
            playCheer
        );


        setTimeout(
            playCheer,
            1000
        );

        </script>

        """


    # ==================================================
    # WATERMARK
    # ==================================================

    html_content += """

        </div>

    </div>

    """


    components.html(
        html_content,
        height=1200,
        scrolling=True
    )

    # ==================================================
    # ΚΑΤΩ ΜΠΑΡΑ: ΥΔΑΤΟΓΡΑΦΗΜΑ + XLSX
    # Το κουμπί είναι κανονικό Streamlit download_button,
    # ώστε να διατηρείται πλήρως η δομή του Excel και τα φύλλα.
    # ==================================================

    bottom_left, bottom_excel, bottom_space = st.columns(
        [1.25, 0.85, 7.90],
        gap="small"
    )

    with bottom_left:
        st.markdown(
            """
            <div style="
                color: rgba(255,255,255,0.22);
                font-size: 12px;
                font-weight: 600;
                letter-spacing: 1px;
                padding-top: 8px;
                white-space: nowrap;
            ">
                tosounidis 2026
            </div>
            """,
            unsafe_allow_html=True,
        )

    with bottom_excel:

        if excel_data is not None:

            excel_b64 = base64.b64encode(
                excel_data
            ).decode("utf-8")

            st.markdown(
                f"""
                <a
                    href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{excel_b64}"
                    target="_blank"
                    rel="noopener noreferrer"
                    download="{excel_download_filename}"
                    style="
                        display:flex;
                        align-items:center;
                        justify-content:center;
                        width:100%;
                        min-height:38px;
                        padding:0 8px;
                        border-radius:8px;
                        border:1px solid rgba(255,255,255,0.25);
                        background:#ffffff;
                        color:#2c3e50;
                        font-size:11px;
                        font-weight:800;
                        text-decoration:none;
                        box-sizing:border-box;
                        white-space:nowrap;
                    "
                >
                    ⬇ XLSX
                </a>
                """,
                unsafe_allow_html=True,
            )

        else:
            st.caption("XLSX error")


except Exception as e:

    st.error(
        f"Σφάλμα κατά τη φόρτωση του dashboard: {e}"
    )
