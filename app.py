import base64
import datetime
import glob
import html
import json
import os
import re
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
product_count_path = "product_count.txt"

# Η τελευταία επιλογή (1 ή 2 προϊόντα) διατηρείται για όλους τους επισκέπτες.
def load_product_count():
    try:
        with open(product_count_path, "r", encoding="utf-8") as f:
            saved_count = int(f.read().strip())
        if saved_count in (1, 2):
            return saved_count
    except (OSError, ValueError):
        pass
    return 2


active_product_count = load_product_count()


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



# Έλεγχος αρχείου Excel πριν από την καταχώρηση.
# Δεν αλλάζει τον τρόπο παρουσίασης των πωλήσεων.
def validate_uploaded_sales_file(file_bytes, display_name):
    try:
        preview = pd.read_excel(BytesIO(file_bytes), header=None)
        if preview.empty:
            raise ValueError("Το αρχείο Excel είναι κενό.")

        def normalized(v):
            txt = unicodedata.normalize("NFD", str(v).strip().lower())
            return "".join(c for c in txt if unicodedata.category(c) != "Mn")

        header_idx = None
        store_idx = quantity_idx = None
        for i in range(len(preview)):
            labels = [normalized(x) for x in preview.iloc[i].values]
            stores = [j for j, x in enumerate(labels) if "καταστημα" in x]
            quantities = [j for j, x in enumerate(labels) if "ποσοτ" in x]
            if stores and quantities:
                header_idx = i
                store_idx = stores[0]
                quantity_idx = quantities[0]
                break
        if header_idx is None:
            raise ValueError("Δεν βρέθηκαν οι στήλες ΚΑΤΑΣΤΗΜΑ και ΠΟΣΟΤΗΤΑ/ΠΟΣΟΤΗΤΕΣ.")

        data = preview.iloc[header_idx + 1:]
        stores = data.iloc[:, store_idx].astype(str).str.strip()
        quantities = data.iloc[:, quantity_idx]
        valid = (stores.ne('') & stores.str.lower().ne('nan')
                 & ~stores.str.contains(r"total|σύνολο|συνολο|κατάστημα", case=False, regex=True, na=False))
        quantities = quantities.loc[valid].apply(
            lambda x: float(x) if isinstance(x, (float, int)) else
            float(str(x).replace('.', '').replace(',', '.')) if ',' in str(x) else float(str(x))
            if str(x).replace('.', '', 1).isdigit() else float('nan')
        )
        valid_sales = quantities.notna().sum()
        if valid_sales == 0:
            raise ValueError("Δεν βρέθηκαν αριθμητικές ποσότητες στα καταστήματα.")
        return int(valid_sales), float(quantities.sum())
    except Exception as exc:
        raise ValueError(f"{display_name}: {exc}") from exc




col_admin, col_input_space = st.columns([4, 4])


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

        if password == "2845":

            st.markdown("---")

            selected_product_count = st.radio(
                "Πόσα προϊόντα θα καταχωρήσετε;",
                options=[1, 2],
                index=0 if active_product_count == 1 else 1,
                horizontal=True,
                key="upload_product_count",
            )

            col_up1, col_up2 = st.columns(2)

            with col_up1:

                uploaded_file_1 = st.file_uploader(
                    "Αρχείο 1 (Πωλήσεις Ειδών Προσφορών):",
                    type=["xlsx"],
                    key="up1",
                )

            uploaded_file_2 = None
            if selected_product_count == 2:
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



            # Δεν γράφουμε Excel αν λείπουν οι αναγκαίες στήλες.
            upload_sales_checks = None
            if (uploaded_file_1 is not None and
                    (selected_product_count == 1 or uploaded_file_2 is not None)):
                try:
                    bytes_1 = uploaded_file_1.getvalue()
                    check_1 = validate_uploaded_sales_file(bytes_1, uploaded_file_1.name)
                    bytes_2 = None
                    check_2 = None
                    if selected_product_count == 2:
                        bytes_2 = uploaded_file_2.getvalue()
                        check_2 = validate_uploaded_sales_file(bytes_2, uploaded_file_2.name)
                    upload_sales_checks = (check_1, check_2)
                except ValueError as error:
                    st.error(f"Δεν ενημερώθηκαν οι πωλήσεις: {error}")
                    uploaded_file_1 = None

            if (
                uploaded_file_1 is not None
                and (
                    selected_product_count == 1
                    or uploaded_file_2 is not None
                )
            ):

                upload_signature = (
                    f"{selected_product_count}_"
                    f"{uploaded_file_1.name}_"
                    f"{uploaded_file_2.name if uploaded_file_2 is not None else ''}_"
                    f"{uploaded_file_1.size}_"
                    f"{uploaded_file_2.size if uploaded_file_2 is not None else 0}"
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

                    f.write(bytes_1)

                # Τα νέα δεδομένα διαβάζονται και από τη μνήμη της τρέχουσας
                # συνεδρίας, ακόμη κι αν το προσωρινό αρχείο διαγραφεί.
                st.session_state["sales_data_excel_1"] = bytes_1

                if gh_token and repo_name:

                    upload_to_github(
                        excel_path_1,
                        repo_name,
                        gh_token,
                        "Auto-update sales file 1",
                    )

                if selected_product_count == 2:
                    with open(
                        excel_path_2,
                        "wb"
                    ) as f:

                        f.write(bytes_2)

                    st.session_state["sales_data_excel_2"] = bytes_2

                    if gh_token and repo_name:

                        upload_to_github(
                            excel_path_2,
                            repo_name,
                            gh_token,
                            "Auto-update sales file 2",
                        )

                with open(product_count_path, "w", encoding="utf-8") as count_file:
                    count_file.write(str(selected_product_count))
                active_product_count = selected_product_count
                st.session_state["sales_data_product_count"] = selected_product_count

                if gh_token and repo_name:
                    upload_to_github(
                        product_count_path,
                        repo_name,
                        gh_token,
                        "Update number of products",
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

                if selected_product_count == 1:
                    st.success(
                        f"Το Excel διαβάστηκε: {upload_sales_checks[0][0]} γραμμές καταστημάτων, "
                        f"ποσότητα {upload_sales_checks[0][1]:,.2f}. Οι πωλήσεις ανανεώθηκαν στην παρούσα συνεδρία."
                    )
                else:
                    st.success(
                        f"Τα δύο Excel διαβάστηκαν: {upload_sales_checks[0][0]} και "
                        f"{upload_sales_checks[1][0]} γραμμές καταστημάτων. "
                        "Οι πωλήσεις ανανεώθηκαν στην παρούσα συνεδρία."
                    )

                # Επιστροφή στην οθόνη πωλήσεων ΜΟΝΟ στο πρώτο
                # upload νέου περιεχομένου (όχι σε κάθε Streamlit rerun).
                import hashlib
                upload_content_id = hashlib.sha256(
                    bytes_1 + b"\x00" + (bytes_2 or b"") +
                    str(selected_product_count).encode("ascii")
                ).hexdigest()
                if st.session_state.get("_last_dashboard_upload") != upload_content_id:
                    st.session_state["_last_dashboard_upload"] = upload_content_id
                    st.session_state["_return_to_dashboard_once"] = True


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

def load_data(path, state_key=None):
    if state_key and st.session_state.get(state_key) is not None:
        try:
            return pd.read_excel(BytesIO(st.session_state[state_key]), header=None)
        except Exception as exc:
            st.error(f"Δεν μπορεί να διαβαστεί το Excel της συνεδρίας: {exc}")
            return pd.DataFrame()

    if os.path.exists(path):
        try:
            return pd.read_excel(path, header=None)
        except Exception as exc:
            st.error(f"Δεν μπορεί να διαβαστεί το Excel {path}: {exc}")
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


def clean_sales_value(val):
    # Διαβάζει την αξία πωλήσεων όπως δίνεται στο αρχείο Excel.
    if pd.isna(val):
        return float("nan")
    if isinstance(val, (int, float)):
        return float(val)

    value = (
        str(val)
        .replace("€", "")
        .replace("\u00a0", "")
        .replace(" ", "")
        .strip()
    )
    if not value:
        return float("nan")
    if "," in value and "." in value:
        value = value.replace(".", "").replace(",", ".")
    elif "," in value:
        value = value.replace(",", ".")
    elif value.count(".") == 1 and len(value.rsplit(".", 1)[1]) == 3:
        value = value.replace(".", "")
    try:
        return float(value)
    except ValueError:
        return float("nan")


def format_sales_euros(value):
    number = f"{value:,.2f}"
    return number.replace(",", "_").replace(".", ",").replace("_", ".") + " €"


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

    # Εντοπίζουμε μόνο τη συνολική αξία πωλήσεων (όχι τιμή μονάδας ή κόστος).
    value_candidates = [
        col for col in df.columns
        if "αξια" in normalize_text(col)
        and "μοναδ" not in normalize_text(col)
        and "τιμη" not in normalize_text(col)
        and "κοστ" not in normalize_text(col)
    ]
    value_col = next(
        (col for col in value_candidates if normalize_text(col) == "αξια"),
        None
    )
    if value_col is None:
        value_col = next(
            (col for col in value_candidates if "πωλησ" in normalize_text(col)),
            None
        )
    if value_col is None and value_candidates:
        value_col = value_candidates[0]

    # ==================================================
    # ΚΡΑΤΑΜΕ ΜΟΝΟ:
    # ΚΑΤΑΣΤΗΜΑ + ΠΟΣΟΤΗΤΑ
    # (ΕΠΙΠΛΕΟΝ ΑΞΙΑ ΓΙΑ ΤΟ ΤΕΛΙΚΟ TOTAL)
    # ==================================================

    df_selected = (
        df[[store_col, quantity_col] + ([value_col] if value_col is not None else [])]
        .copy()
    )

    df_selected.columns = [
        "Κατάστημα",
        "Ποσότητα"
    ] + (["Αξία"] if value_col is not None else [])

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

    # Η αξία ακολουθεί ακριβώς τα ίδια καταστήματα/φίλτρα με την ποσότητα.
    df_clean["Num_Value"] = (
        df_clean["Αξία"].apply(clean_sales_value)
        if value_col is not None
        else float("nan")
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
        load_data(excel_path_1, "sales_data_excel_1"),
        file_name=excel_path_1
    )
)


title_2, df_stores_2, _, max_sales_2 = (
    process_sales_df(
        load_data(excel_path_2, "sales_data_excel_2"),
        file_name=excel_path_2
    )
)


def _build_store_code_owners():
    """Χαρτογραφεί κάθε κωδικό καταστήματος σε μία και μόνο ομάδα."""
    owners = {}
    ambiguous = set()

    for group_name, keywords in GROUPS_MAPPING.items():
        if group_name in ("τομεας 3", "τομέας 3"):
            continue
        for keyword in keywords:
            value = str(keyword).strip()
            if re.fullmatch(r"\d{3,4}", value):
                if value in owners and owners[value] != group_name:
                    ambiguous.add(value)
                else:
                    owners[value] = group_name

    for code in ambiguous:
        owners.pop(code, None)
    return owners


STORE_CODE_OWNERS = _build_store_code_owners()


def _build_text_keyword_owners():
    """Κρατά μόνο φράσεις που ανήκουν αποκλειστικά σε μία ομάδα."""
    owners = {}
    ambiguous = set()

    for group_name, keywords in GROUPS_MAPPING.items():
        if group_name in ("τομεας 3", "τομέας 3"):
            continue
        for keyword in keywords:
            value = str(keyword).strip()
            if not value or re.fullmatch(r"\d{3,4}", value):
                continue
            normalized = normalize_text(value)
            if normalized in owners and owners[normalized] != group_name:
                ambiguous.add(normalized)
            else:
                owners[normalized] = group_name

    for phrase in ambiguous:
        owners.pop(phrase, None)
    return owners


TEXT_KEYWORD_OWNERS = _build_text_keyword_owners()


def _store_belongs_to_group(store_name, group_name):
    """Αντιστοίχιση πρώτα με ακριβή κωδικό, ποτέ με αριθμούς διεύθυνσης."""
    normalized_name = normalize_text(store_name)
    code_match = re.match(r"^\s*(\d{3,4})(?=\D|$)", normalized_name)

    if code_match:
        # Όταν υπάρχει κωδικός στην αρχή, αυτός είναι η μοναδική πηγή αλήθειας.
        return STORE_CODE_OWNERS.get(code_match.group(1)) == group_name

    # Εφεδρική αντιστοίχιση μόνο με μοναδικές, ολόκληρες φράσεις διεύθυνσης.
    for phrase, owner in TEXT_KEYWORD_OWNERS.items():
        if owner != group_name:
            continue
        pattern = r"(?<!\w)" + re.escape(phrase) + r"(?!\w)"
        if re.search(pattern, normalized_name):
            return True

    return False


def filter_dataframe(df_stores):

    if df_stores.empty:
        return df_stores.copy(), 0.0

    filtered_df = df_stores.copy()

    if active_filter and active_filter not in ("τομεας 3", "τομέας 3"):
        if active_filter in GROUPS_MAPPING:
            filtered_df = filtered_df[
                filtered_df["Κατάστημα"].apply(
                    lambda store: _store_belongs_to_group(
                        store, active_filter
                    )
                )
            ]
        else:
            filtered_df = filtered_df[
                filtered_df["Κατάστημα"].astype(str).str.contains(
                    re.escape(active_filter), case=False, na=False, regex=True
                )
            ]

    total_sum = filtered_df["Num_Sales"].sum()

    return filtered_df.reset_index(drop=True), total_sum


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

# Υπολογισμός της συνολικής αξίας στα ήδη φιλτραρισμένα καταστήματα.
total_value_1 = (
    df_stores_1["Num_Value"].sum(min_count=1)
    if "Num_Value" in df_stores_1.columns else float("nan")
)
total_value_2 = (
    df_stores_2["Num_Value"].sum(min_count=1)
    if "Num_Value" in df_stores_2.columns else float("nan")
)


# ==================================================
# ΚΥΛΙΟΜΕΝΟ ΜΗΝΥΜΑ: ΚΑΤΑΣΤΗΜΑΤΑ ΜΕ 0 Ή 1 ΠΩΛΗΣΗ
# Ελέγχει μόνο καταστήματα που υπάρχουν στο Excel.
# ==================================================

def stores_with_zero_or_one_sale(df_stores):
    if df_stores.empty or not {"Κατάστημα", "Num_Sales"}.issubset(df_stores.columns):
        return []

    # Ο κωδικός είναι τα 3 ή 4 πρώτα ψηφία (όπως στο φίλτρο περιφερειών).
    store_codes = df_stores["Κατάστημα"].astype(str).str.extract(
        r"^\s*(\d{3,4})(?=\D|$)", expand=False
    )
    eligible = df_stores.assign(_store_code=store_codes).dropna(subset=["_store_code"])
    if eligible.empty:
        return []

    # Συγκεντρώνει τυχόν περισσότερες από μία γραμμές του ίδιου καταστήματος.
    sales_by_store = eligible.groupby("_store_code")["Num_Sales"].sum()
    low_sales = sales_by_store[sales_by_store.isin([0.0, 1.0])]
    return [
        code
        for code, qty in sorted(low_sales.items(), key=lambda pair: (pair[1], int(pair[0])))
    ]


# Ξεχωριστό κυλιόμενο μήνυμα για κάθε ενεργό προϊόν.
def build_product_ticker(df_stores, product_title):
    codes = stores_with_zero_or_one_sale(df_stores)
    if df_stores.empty:
        codes_text = "Δεν υπάρχουν διαθέσιμα δεδομένα"
    elif not codes:
        codes_text = "Κανένα κατάστημα"
    else:
        codes_text = ", ".join(sorted(set(codes), key=int))
    safe_title = html.escape(str(product_title))
    ticker_mode = "moving" if codes else "static"
    return (
        f'<div class="sales-warning-ticker" data-mode="{ticker_mode}">'
        '<div class="sales-warning-viewport">'
        f'<div class="sales-warning-ticker-text">⚠️ {safe_title} — Καταστήματα με 0 ή 1 πώληση: {codes_text}</div>'
        '</div></div>'
    )


ticker_html = build_product_ticker(df_stores_1, title_1)
if active_product_count == 2:
    ticker_html += build_product_ticker(df_stores_2, title_2)


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

    summary_rows = [
        ["Περιφέρεια", selected_region.upper()],
        ["Ημερομηνία", file_date_str],
        ["Ώρα", file_time_str],
        ["Προϊόν 1", title_1],
    ]
    if active_product_count == 2:
        summary_rows.append(["Προϊόν 2", title_2])

    summary_df = pd.DataFrame(
        summary_rows,
        columns=[
            "Πεδίο",
            "Τιμή"
        ]
    )

    with pd.ExcelWriter(
        output,
        engine="openpyxl"
    ) as writer:

        summary_df.to_excel(
            writer,
            sheet_name="Σύνοψη",
            index=False
        )

        export_1.to_excel(
            writer,
            sheet_name="Προϊόν 1",
            index=False
        )

        if active_product_count == 2:
            export_2.to_excel(
                writer,
                sheet_name="Προϊόν 2",
                index=False
            )

        for sheet_name in (["Σύνοψη", "Προϊόν 1"] +
                           (["Προϊόν 2"] if active_product_count == 2 else [])):

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




# ==================================================
# BANNER - ΑΝΑΖΗΤΗΣΗ ΣΤΟΝ ΦΑΚΕΛΟ ΤΗΣ ΕΦΑΡΜΟΓΗΣ
# ==================================================

img_src = ""

# Το Streamlit Cloud δεν εκτελεί πάντα την εφαρμογή από τον ίδιο
# working directory. Γι' αυτό ψάχνουμε πρώτα δίπλα στο app.py
# και έπειτα στους συνηθισμένους φακέλους εικόνων.
app_dir = os.path.dirname(os.path.abspath(__file__))
search_dirs = [
    app_dir,
    os.getcwd(),
    os.path.join(app_dir, "assets"),
    os.path.join(app_dir, "static"),
    os.path.join(app_dir, "images"),
]

valid_extensions = {".png", ".jpg", ".jpeg", ".gif", ".webp"}
banner_files = []
seen_paths = set()

for folder in search_dirs:
    if not os.path.isdir(folder):
        continue
    try:
        for filename in os.listdir(folder):
            lower_name = filename.casefold()
            extension = os.path.splitext(filename)[1].lower()
            looks_like_banner = (
                "banner" in lower_name
                or "image 21 σεπ 2026, 11_24_53 π.μ." in lower_name
                or "chatgpt image" in lower_name
                or "σφυρι" in lower_name
                or "σφυρί" in lower_name
                or "sfyri" in lower_name
                or "hammer" in lower_name
            )
            if extension in valid_extensions and looks_like_banner:
                full_path = os.path.abspath(os.path.join(folder, filename))
                if full_path not in seen_paths:
                    seen_paths.add(full_path)
                    banner_files.append(full_path)
    except OSError:
        continue

if banner_files:
    # Επιλέγουμε το πιο πρόσφατο αρχείο banner.
    banner_filename = max(banner_files, key=os.path.getmtime)
    extension = os.path.splitext(banner_filename)[1].lower()
    mime_types = {
        ".png": "image/png",
        ".jpg": "image/jpeg",
        ".jpeg": "image/jpeg",
        ".gif": "image/gif",
        ".webp": "image/webp",
    }
    mime_type = mime_types.get(extension, "image/png")

    try:
        with open(banner_filename, "rb") as image_file:
            encoded_image = base64.b64encode(image_file.read()).decode("ascii")
        img_src = f"data:{mime_type};base64,{encoded_image}"
    except OSError:
        img_src = ""

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


    /* Στατικό μήνυμα χωρίς κωδικούς, αλλιώς κύλιση ολόκληρου του κειμένου. */
    @keyframes sales-warning-scroll {{
        from {{ transform: translateX(var(--ticker-start)); }}
        to {{ transform: translateX(var(--ticker-end)); }}
    }}
    .sales-warning-ticker {{
        width: 100%;
        box-sizing: border-box;
        padding: 8px 12px;
        color: #f8e7ad;
        font-size: 12px;
        font-weight: 700;
        text-align: left;
        border-bottom: 1px solid rgba(255,255,255,0.08);
    }}
    .sales-warning-viewport {{ width: 100%; overflow: hidden; }}
    .sales-warning-ticker-text {{ width: max-content; white-space: nowrap; }}
    .sales-warning-ticker[data-mode="static"] .sales-warning-ticker-text {{
        width: auto; white-space: normal;
    }}
    .sales-warning-ticker-text.ready {{
        animation: sales-warning-scroll var(--ticker-duration, 30s) linear infinite;
    }}
    .sales-warning-ticker:hover .sales-warning-ticker-text {{ animation-play-state: paused; }}
    @media (prefers-reduced-motion: reduce) {{
        .sales-warning-ticker-text.ready {{ animation: none; }}
        .sales-warning-viewport {{ overflow-x: auto; }}
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


    .watermark {{

        text-align: left;

        color:
            rgba(255, 255, 255, 0.16);

        font-size: 12px;

        font-weight: 600;

        letter-spacing: 1px;

        margin-top: 15px;

        margin-left: 0;

        margin-right: 0;

        margin-bottom: 0;

        text-transform: none;

        user-select: none;

    }}


    </style>


    <div class="main-container">


        {f'<div class="banner-container"><img src="{img_src}" class="banner-img" alt="banner"></div>' if img_src else ""}

        {ticker_html}
        <script>
        (() => {{
            function fillTickers() {{
                document.querySelectorAll('.sales-warning-ticker').forEach(row => {{
                    const viewport = row.querySelector('.sales-warning-viewport');
                    const text = row.querySelector('.sales-warning-ticker-text');
                    text.classList.remove('ready');
                    if (row.dataset.mode !== "moving") return;
                    const start = viewport.clientWidth;
                    const end = -text.scrollWidth;
                    text.style.setProperty('--ticker-start', `${{start}}px`);
                    text.style.setProperty('--ticker-end', `${{end}}px`);
                    text.style.setProperty('--ticker-duration', `${{Math.max(12, (start - end) / 65)}}s`);
                    text.classList.add('ready');
                }});
            }}
            fillTickers();
            window.addEventListener('resize', fillTickers);
            if (document.fonts) document.fonts.ready.then(fillTickers);
        }})();
        </script>


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
        formatted_value_1 = (
            format_sales_euros(total_value_1)
            if pd.notna(total_value_1) else "Μη διαθέσιμη"
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
                    <small style="display:block; color:#ffd166; font-size:12px; font-weight:700;">
                        Αξία: {formatted_value_1}
                    </small>
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


    if active_product_count == 2:
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
            formatted_value_2 = (
                format_sales_euros(total_value_2)
                if pd.notna(total_value_2) else "Μη διαθέσιμη"
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
                        <small style="display:block; color:#ffd166; font-size:12px; font-weight:700;">
                            Αξία: {formatted_value_2}
                        </small>
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

            <div class="watermark">
                tosounidis 2026
            </div>

        </div>

    </div>

    """


    # Αόρατος στόχος για επιστροφή στο banner μετά την καταχώρηση.
    st.markdown('<div id="sales-display-anchor"></div>', unsafe_allow_html=True)

    components.html(
        html_content,
        height=1200,
        scrolling=True
    )


except Exception as e:

    st.error(
        f"Σφάλμα κατά τη φόρτωση του dashboard: {e}"
    )


# ==================================================
# ΚΟΥΜΠΙ XLSX - ΤΕΡΜΑ ΚΑΤΩ ΣΤΗΝ ΕΦΑΡΜΟΓΗ
# ==================================================

try:

    excel_data = build_excel_file()

    safe_region_name = (
        normalize_text(selected_region)
        .replace(" ", "_")
    )

    st.download_button(
        label="⬇ XLSX",
        data=excel_data,
        file_name=(
            f"sales_{safe_region_name}_"
            f"{datetime.date.today().strftime('%Y%m%d')}.xlsx"
        ),
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
        key="download_current_sales_xlsx",
    )

except Exception as excel_error:

    st.caption(
        "XLSX error"
    )


# ==================================================
# ΕΠΙΣΤΡΟΦΗ ΣΤΗΝ ΠΑΡΟΥΣΙΑΣΗ ΜΕΤΑ ΤΟ UPLOAD
# Δεν γίνεται πλήρης ανανέωση.
# ==================================================
if st.session_state.pop("_return_to_dashboard_once", False):
    components.html(
        """
        <script>
        (() => {
          const goBack = () => {
            try {
              const doc = window.parent.document;
              const admin = Array.from(doc.querySelectorAll('details'))
                .find(el => (el.querySelector('summary')?.textContent || '')
                  .includes('Διαχείριση Αρχείων'));
              if (admin && admin.open) admin.querySelector('summary')?.click();
              setTimeout(() => {
                const banner = doc.getElementById('sales-display-anchor');
                banner?.scrollIntoView({behavior: 'smooth', block: 'start'});
              }, 350);
            } catch (err) {
              /* Δεν επηρεάζουμε ποτέ την καταχώρηση Excel. */
            }
          };
          setTimeout(goBack, 450);
        })();
        </script>
        """,
        height=0,
    )
