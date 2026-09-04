import base64
import datetime
import glob
import json
import os
import pandas as pd
import requests
import streamlit as st
import streamlit.components.v1 as components

st.set_page_config(page_title="Πωλήσεις 2 Προϊόντων ανά Κατάστημα", layout="wide")

st.markdown(
    """
    <style>
    .stApp { background-color: #2c3e50 !important; }
    #MainMenu {visibility: hidden;} 
    header {visibility: hidden;} 
    footer {visibility: hidden;}
    [data-testid="stToolbar"] {visibility: hidden; display: none;}
    [data-testid="stDecoration"] {visibility: hidden; display: none;}
    div[data-baseweb="select"] > div, .stRadio label p { color: white !important; }
    .block-container { padding: 0rem 0.5rem !important; max-width: 100% !important; }
    
    div[data-testid="stHorizontalBlock"] {
        align-items: center !important;
    }
    </style>
    
    <script>
    function removeManageButton() {
        const doc = window.parent.document;
        const buttons = doc.querySelectorAll('button');
        buttons.forEach(btn => {
            if (btn.innerText.includes('Manage app') || btn.innerHTML.includes('Manage')) {
                btn.style.display = 'none';
            }
        });
        
        const pwdInputs = doc.querySelectorAll('input[type="password"]');
        pwdInputs.forEach(input => {
            input.setAttribute('autocomplete', 'username');
            input.setAttribute('data-form-type', 'other');
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

# --- ΠΛΗΡΗΣ ΚΑΙ ΕΝΗΜΕΡΩΜΕΝΟΣ ΚΑΤΑΛΟΓΟΣ ΔΙΚΤΥΩΝ (DICTIONARY) ---
catalog_data = {
    "ΜΗΤΡΟΥΛΗΣ": [
        {"code": "301", "address": "ΑΙΑΝΗ - ΚΟΖΑΝΗ"},
        {"code": "302", "address": "ΠΛ. ΕΛΕΥΘΕΡΙΑΣ - ΚΟΖΑΝΗ"},
        {"code": "309", "address": "ΠΛ. ΛΑΣΣΑΝΗ 13 - ΚΟΖΑΝΗ"},
        {"code": "486", "address": "Μ. ΑΛΕΞΑΝΔΡΟΥ 23 - ΣΙΑΤΙΣΤΑ"},
        {"code": "304", "address": "25ης ΜΑΡΤΙΟΥ 74 - ΠΤΟΛΕΜΑΪΔΑ"},
        {"code": "352", "address": "ΓΡΕΒΕΝΑ"},
        {"code": "308", "address": "ΦΙΛΙΠΠΟΥ 1 - ΠΤΟΛΕΜΑΪΔΑ"},
        {"code": "353", "address": "ΕΛ. ΒΕΝΙΖΕΛΟΥ & ΙΩΑΝΝΗ ΑΡΤΗ - ΦΛΩΡΙΝΑ"},
        {"code": "354", "address": "ΚΟΖΑΝΗΣ ΚΑΙ ΓΡΕΒΕΝΩΝ - ΦΛΩΡΙΝΑ"},
        {"code": "355", "address": "Χλόη Καστοριάς - ΚΑΣΤΟΡΙΑ"},
        {"code": "356", "address": "ΔΙΣΠΥΛΟ ΚΑΣΤΟΡΙΑΣ"},
        {"code": "366", "address": "ΑΘ. ΔΙΑΚΟΥ 30 & ΓΡΑΜΜΟΥ 59"},
        {"code": "374", "address": "ΜΑΝΙΑΚΟΙ ΚΑΣΤΟΡΙΑΣ"}
    ],
    "ΠΟΝΟΠΟΥΛΟΣ": [
        {"code": "231", "address": "ΣΒΟΡΩΝΟΥ"},
        {"code": "232", "address": "ΥΨΗΛΑΝΤΟΥ 3 - ΚΑΤΕΡΙΝΗ"},
        {"code": "233", "address": "ΠΕΡΔΙΚΑ 8 – ΚΑΤΕΡΙΝΗ"},
        {"code": "373", "address": "25ης ΜΑΡΤΙΟΥ & ΕΓΝΑΤΙΑΣ (ΠΛΑΤΑΜΩΝΑΣ)"},
        {"code": "237", "address": "ΛΕΠΤΟΚΑΡΥΑ"},
        {"code": "235", "address": "ΚΟΡΙΝΟΣ"},
        {"code": "372", "address": "ΑΓ. ΝΙΚΟΛΑΟΥ - ΛΙΤΟΧΩΡΟ"},
        {"code": "236", "address": "ΑΝΤΙΓΟΝΟΥ 2-6 - ΚΑΤΕΡΙΝΗ"},
        {"code": "483", "address": "ΑΡΙΣΤΟΤΕΛΟΥΣ 2-4 - ΚΑΤΕΡΙΝΗ"},
        {"code": "161", "address": "Μ. ΑΛΕΞΑΝΔΡΟΥ 74 - ΚΑΤΕΡΙΝΗ"},
        {"code": "166", "address": "Π. ΤΣΑΛΔΑΡΗ - ΚΑΤΕΡΙΝΗ"},
        {"code": "384", "address": "ΚΑΤΕΡΙΝΗ - 19ης Μαΐου 24"},
        {"code": "234", "address": "ΧΑΤΖΟΓΛΟΥ 2 – ΚΑΤΕΡΙΝΗ"}
    ],
    "ΧΑΡΑΛΑΜΠΙΔΗΣ": [
        {"code": "211", "address": "Δ. ΓΕΩΡΓΙΑΔΟΥ 24 - ΛΑΡΙΣΑ"},
        {"code": "201", "address": "ΝΙΚΗΤΑΡΑ 13 - ΛΑΡΙΣΑ"},
        {"code": "347", "address": "ΙΩΑΝΝΙΝΩΝ 80 - ΛΑΡΙΣΑ"},
        {"code": "212", "address": "23ης ΟΚΤΩΒΡΙΟΥ 102-104 - ΛΑΡΙΣΑ"},
        {"code": "239", "address": "Χατζημιχάλη 49 - Φιλιππούπολη ΛΑΡΙΣΑ"},
        {"code": "219", "address": "ΘΥΑΤΕΙΡΩΝ & Βενιζέλου Ν. Ιωνία - ΒΟΛΟΣ"},
        {"code": "220", "address": "ΑΧΙΛΛΟΠΟΥΛΟΥ 171 - ΒΟΛΟΣ"},
        {"code": "222", "address": "ΚΑΣΣΑΒΕΤΗ 14 & 28ης ΟΚΤΩΒΡΙΟΥ - ΒΟΛΟΣ"},
        {"code": "223", "address": "Ν. Ιωνία - ΒΟΛΟΣ"},
        {"code": "240", "address": "ΚΟΥΜΟΥΝΔΟΥΡΟΥ 150 - ΒΟΛΟΣ"},
        {"code": "241", "address": "ΜΕΤΑΜΟΡΦΩΣΕΩΣ 21 & ΑΛΕΞΑΝΔΡΑΣ - ΒΟΛΟΣ"},
        {"code": "493", "address": "ΣΚΙΑΘΟΣ"}
    ],
    "ΠΑΠΠΑΣ": [
        {"code": "210", "address": "ΕΛΑΣΣΟΝΑ ΒΥΖΑΝΤΙΟΥ"},
        {"code": "202", "address": "ΑΒΕΡΩΦ 22 - ΚΑΡΔΙΤΣΑ"},
        {"code": "204", "address": "ΚΑΡΑΪΣΚΑΚΗ 95 - ΚΑΡΔΙΤΣΑ"},
        {"code": "206", "address": "ΣΟΦΑΔΕΣ ΚΑΡΔΙΤΣΑΣ"},
        {"code": "209", "address": "ΚΟΝΔΥΛΗ - ΤΡΙΚΑΛΑ"},
        {"code": "205", "address": "ΚΟΝΔΥΛΗ 15 - ΤΡΙΚΑΛΑ"},
        {"code": "207", "address": "ΔΕΛΗΓΙΩΡΓΗ - ΤΡΙΚΑΛΑ"},
        {"code": "208", "address": "ΕΛΕΥΘΕΡΙΟΣ - ΤΡΙΚΑΛΑ"},
        {"code": "215", "address": "ΑΒΕΡΩΦ ΚΑΙ ΛΑΡΙΣΗΣ"}
    ],
    "ΠΑΤΣΗΣ": [
        {"code": "198", "address": "ΓΚΟΥΡΑΣ & ΝΙΚΟΠΟΛΕΩΣ - ΙΩΑΝΝΙΝΑ"},
        {"code": "225", "address": "Γ. ΠΑΠΑΝΔΡΕΟΥ 26-28 - ΙΩΑΝΝΙΝΑ"},
        {"code": "226", "address": "ΚΑΤΩ ΝΕΟΧΩΡΟΠΟΥΛΟ - ΙΩΑΝΝΙΝΑ"},
        {"code": "316", "address": "ΜΑΡΙΚΑΣ ΚΟΤΟΠΟΥΛΗ 66–68 ΙΩΑΝΝΙΝΑ ΡΙΖΑΡΙΟ"},
        {"code": "317", "address": "ΛΕΩΦ. ΔΗΜΟΚΡΑΤΙΑΣ ΚΑΡΔΑΜΙΤΣΙΑ - ΙΩΑΝΝΙΝΑ"},
        {"code": "381", "address": "ΚΟΡΑΗ - ΙΩΑΝΝΙΝΑ"},
        {"code": "228", "address": "ΚΑΡΥΩΤΑΚΗ 15 & ΛΕΩΦ. ΕΙΡΗΝΗΣ - ΠΡΕΒΕΖΑ"},
        {"code": "229", "address": "ΙΩΑΝΝΙΝΩΝ 199 - ΠΡΕΒΕΖΑ"},
        {"code": "224", "address": "ΠΛ. ΚΙΛΚΙΣ & ΑΝΕΞΑΡΤΗΣΙΑΣ - ΑΡΤΑ"},
        {"code": "315", "address": "26ο χλμ ΕΘΝΙΚΗΣ ΟΔΟΥ ΠΡΕΒΕΖΗΣ - ΙΩΑΝΝΙΝΩΝ ΛΟΥΡΟΣ"},
        {"code": "359", "address": "ΗΓΟΥΜΕΝΙΤΣΑ ΘΕΣΠΡΩΤΙΑΣ"},
        {"code": "378", "address": "ΠΑΡΑΜΥΘΙΑ ΘΕΣΠΡΩΤΙΑΣ"},
        {"code": "399", "address": "ΗΓΟΥΜΕΝΙΤΣΑ ΘΕΣΠΡΩΤΙΑΣ"},
        {"code": "444", "address": "ΑΓ. ΜΑΡΙΝΑΣ - ΙΩΑΝΝΙΝΑ"}
    ],
    "ΣΚΙΑΔΟΠΟΥΛΟΣ": [
        {"code": "531", "address": "ΚΕΡΚΥΡΑ ΑΧΑΡΑΒΗ"},
        {"code": "539", "address": "ΚΕΡΚΥΡΑ ΚΑΣΣΙΩΠΗ"},
        {"code": "567", "address": "ΚΕΡΚΥΡΑ ΣΙΔΑΡΙ ΚΑΡΟΥΣΑΔΕΣ ΜΑΡΚΑΤΟ"},
        {"code": "537", "address": "ΚΕΡΚΥΡΑ ΜΑΝΤΟΥΚΙ"},
        {"code": "533", "address": "ΚΕΡΚΥΡΑ ΑΛΥΚΕΣ"},
        {"code": "525", "address": "ΚΕΡΚΥΡΑ ΥΠΕΡ. ΕΘΝΙΚΗ ΟΔΟΣ ΛΕΥΚΙΜΜΗΣ ΚΕΡΚΥΡΑΣ"},
        {"code": "535", "address": "ΚΕΡΚΥΡΑ ΛΕΥΚΙΜΜΗ"},
        {"code": "534", "address": "ΚΕΡΚΥΡΑ ΜΩΡΑΙΤΙΚΑ"},
        {"code": "566", "address": "ΚΕΡΚΥΡΑ ΚΟΜΒΟΣ ΒΡΥΩΝΗ ΚΑΣΤΕΛΛΑΝΟΙ ΜΑΡΚΑΤΟ"},
        {"code": "540", "address": "ΚΕΡΚΥΡΑ ΑΛΕΠΟΥ ΜΑΡΚΑΤΟ"},
        {"code": "565", "address": "ΚΕΡΚΥΡΑ ΣΑΡΟΚΟ ΜΑΡΚΑΤΟ"},
        {"code": "530", "address": "ΚΕΡΚΥΡΑ ΙΩΑΝΝΟΥ ΘΕΟΤΟΚΗ ΠΑΛΛΑΔΑ"},
        {"code": "528", "address": "ΚΕΡΚΥΡΑ ΛΑΙΚΗ ΑΓΟΡΑ ΓΕΡΑΣΙΜΟΥ ΜΑΡΚΟΡΑ"},
        {"code": "529", "address": "ΚΕΡΚΥΡΑ ΠΙΝΙΑ"},
        {"code": "532", "address": "ΚΕΡΚΥΡΑ ΝΟΣΟΚΟΜΕΙΟ"},
        {"code": "536", "address": "ΚΕΡΚΥΡΑ ΣΠΗΛΙΑ"},
        {"code": "538", "address": "ΚΕΡΚΥΡΑ ΜΗΤΡΟΠΟΛΙΤΟΥ ΜΕΘΟΔΙΟΥ"},
        {"code": "549", "address": "ΜΩΡΑΙΤΙΚΑ ΝΈΟ"}
    ],
    "ΣΑΜΟΓΛΟΥ": [
        {"code": "8907", "address": "ΚΑΣΤΟΡΙΑ ΧΙΛΙΟΔΕΝΔΡΟ ΓΡ"},
        {"code": "8912", "address": "ΚΑΣΤΟΡΙΑ ΑΓ. ΜΗΝΑ ΓΡ"},
        {"code": "8913", "address": "ΚΑΣΤΟΡΙΑ Μ. ΑΛΕΞΑΝΔΡΟΥ ΓΡ"},
        {"code": "8918", "address": "ΚΟΖΑΝΗ ΠΑΥΛΟΥ ΧΑΡΙΣΗ ΓΡ"},
        {"code": "8920", "address": "ΓΡΕΒΕΝΑ ΜΑΚΕΔΟΝΟΜΑΧΩΝ ΓΡ"},
        {"code": "8926", "address": "ΣΕΡΒΙΑ ΚΟΖΑΝΗΣ ΚΩΝ. ΚΑΡΠΟΥ ΓΡ"},
        {"code": "8932", "address": "ΣΙΑΤΙΣΤΑ ΓΡ"},
        {"code": "8933", "address": "ΒΕΛΒΕΝΤΟΣ ΚΟΖΑΝΗΣ ΓΡ"},
        {"code": "8938", "address": "ΚΟΖΑΝΗ ΟΣΕ ΓΡ"},
        {"code": "8939", "address": "ΓΡΕΒΕΝΑ ΘΕΩΔ. ΖΙΑΚΑ ΓΡ"},
        {"code": "8941", "address": "ΓΡΕΒΕΝΑ ΚΑΒΑΦΗ ΓΡ"},
        {"code": "8943", "address": "ΚΟΖΑΝΗ ΦΙΛΙΠΠΟΥ Β΄ ΓΡ"}
    ],
    "ΜΠΟΥΤΣΚΟΣ": [
        {"code": "8267", "address": "ΔΟΜΟΚΟΣ ΠΑΠΑΦΛΕΣΣΑ"},
        {"code": "8273", "address": "ΠΛΑΤΑΜΩΝΑΣ ΚΩΝ. ΚΑΡΑΜΑΝΛΗ"},
        {"code": "8317", "address": "ΑΛΜΥΡΟΣ ΒΟΛΟΥ Ν. ΜΙΧΟΠΟΥΛΟΥ"},
        {"code": "8318", "address": "ΘΕΣΣ. ΣΙΝΔΟΣ ΠΑΛΑΙΟΛΟΓΟΥ"},
        {"code": "8334", "address": "ΘΕΣΣΑΛΟΝΙΚΗ ΟΛΥΜΠΟΥ 93"},
        {"code": "8402", "address": "ΛΑΡΙΣΑ ΣΤΡΑΤΗΓΟΥ ΦΡΑΓΚΟΥ"},
        {"code": "8408", "address": "ΣΚΟΠΕΛΟΣ 2 ΧΛΜ ΕΠΟ. ΣΚΟΠΕΛΟΥ"},
        {"code": "8448", "address": "ΛΑΡΙΣΑ ΙΩΑΝΝΙΝΩΝ"},
        {"code": "8568", "address": "ΒΟΛΟΣ ΓΙΑΝΝΗ ΔΗΜΟΥ ΣΥ."},
        {"code": "8569", "address": "ΣΤΕΦΑΝΟΒΙΚΕΙΟ - ΒΕΛΕΣΤΙΝΟ ΣΥ."},
        {"code": "8591", "address": "ΑΓΙΑ ΛΑΡΙΣΗΣ"}
    ],
    "ΠΟΥΤΟΓΛΙΔΗΣ": [
        {"code": "8333", "address": "ΠΤΟΛΕΜΑΪΔΑ 25 ΜΑΡΤΙΟΥ"},
        {"code": "8916", "address": "ΦΛΩΡΙΝΑ ΚΡΕΣΝΑΣ ΓΡ"},
        {"code": "8917", "address": "ΦΛΩΡΙΝΑ ΚΑΣΤΡΙΣΙΑΝΑΚΗ ΓΡ"},
        {"code": "8923", "address": "ΦΛΩΡΙΝΑ ΔΗΜ. ΠΑΠΑΘΑΝΑΣΙΟΥ ΓΡ"},
        {"code": "8925", "address": "ΚΡΟΚΟΣ ΚΟΖΑΝΗΣ ΙΩΑΚ. ΛΙΟΥΛΙΑ ΓΡ"},
        {"code": "8928", "address": "ΦΙΛΩΤΑΣ Μ. ΑΛΕΞΑΝΔΡΟΥ ΓΡ"},
        {"code": "8931", "address": "ΑΜΥΝΤΑΙΟ 28 ΣΥΝΤΑΓΜ. ΠΕΖΙΚΟΥ ΓΡ"},
        {"code": "8936", "address": "ΠΤΟΛΕΜΑΪΔΑ ΧΡΥΣ. ΣΜΥΡΝΗΣ ΓΡ"},
        {"code": "8937", "address": "ΠΤΟΛΕΜΑΪΔΑ ΘΕΟΛΟΓΙΔΗ ΓΡ"},
        {"code": "8940", "address": "ΦΛΩΡΙΝΑ CASH & CARRY ΓΡ"},
        {"code": "8945", "address": "ΚΟΖΑΝΗ ΚΡΟΚΟΣ CASH & CARRY ΓΡ"}
    ]
}

# Δυναμική παραγωγή των GROUPS_MAPPING με βάση το νέο catalog_data
GROUPS_MAPPING = {}
tomeas_3_list = []

for manager, stores in catalog_data.items():
    mgr_key = manager.lower()
    mgr_keywords = []
    for store in stores:
        code = str(store["code"]).strip().lower()
        address = str(store["address"]).strip().lower()
        mgr_keywords.append(code)
        mgr_keywords.append(address)
        for word in address.replace("-", " ").replace("/", " ").replace("&", " ").split():
            if len(word) > 2:
                mgr_keywords.append(word)
                
    mgr_keywords = list(set(mgr_keywords))
    GROUPS_MAPPING[mgr_key] = mgr_keywords
    tomeas_3_list.extend(mgr_keywords)

GROUPS_MAPPING["τομεας 3"] = list(set(tomeas_3_list))
GROUPS_MAPPING["τομέας 3"] = list(set(tomeas_3_list))
