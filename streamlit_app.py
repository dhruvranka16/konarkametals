import re
import pandas as pd
import streamlit as st
from io import BytesIO
import json

# --- MODIFIED: Robust secrets handling ---
# This try...except block allows the app to run locally without a secrets.toml file
try:
    # Check if we have secrets defined (on Streamlit Cloud or in a local file)
    if "firestore_service_account" in st.secrets:
        from google.cloud import firestore
        from google.oauth2 import service_account
    else:
        # Secrets file exists, but the key is missing
        st.warning("Firebase secrets not found. App will run in local-only mode. Rule changes will not be saved.")
        firestore = None
except st.errors.StreamlitSecretNotFoundError:
    # This block runs ONLY if the secrets.toml file is completely missing
    st.warning("Running in local-only mode (no secrets.toml file found). Rule changes will not be saved.")
    firestore = None
# --- END MODIFICATION ---


# --- Default Rule Definitions ---
# These are used to "seed" the database the first time the app runs.
DEFAULT_P1_RULES = {
    "Equal Angle / Unequal Angle": {"prod": (180, "Production rate below 180 Prod/hour"), "rec": (80, "Recovery below 80%"), "spd": (4, "Speed less than 4 mm")},
    "Rectangular Tube": {"prod": (220, "Production rate below 220 Prod/hour"), "rec": (80, "Recovery below 80%"), "spd": (5.5, "Speed less than 5.5 mm")},
    "Square Tube": {"prod": (280, "Production rate below 280 Prod/hour"), "rec": (82, "Recovery below 82%"), "spd": (5.5, "Speed less than 5.5 mm")},
    "Round Pipe": {"prod": (230, "Production rate below 230 Prod/hour"), "rec": (78, "Recovery below 78%"), "spd": (4.5, "Speed less than 4.5 mm")},
    "Single Track Top / Single Track Bottom": {"prod": (240, "Production rate below 240 Prod/hour"), "rec": (81, "Recovery below 81%"), "spd": (5.0, "Speed less than 5.0 mm")},
    "Mini Dumal": {"prod": (250, "Production rate below 250 Prod/hour"), "rec": (80, "Recovery below 80%"), "spd": (5.0, "Speed less than 5.0 mm")},
    "40 mm Clip": {"prod": (230, "Production rate below 230 Prod/hour"), "rec": (79, "Recovery below 79%"), "spd": (4.8, "Speed less than 4.8 mm")},
    "40 mm Outer": {"prod": (260, "Production rate below 260 Prod/hour"), "rec": (80, "Recovery below 80%"), "spd": (5.0, "Speed less than 5.0 mm")},
    "40 mm Frame": {"prod": (260, "Production rate below 260 Prod/hour"), "rec": (80, "Recovery below 80%"), "spd": (5.0, "Speed less than 5.0 mm")},
    "Two Track Top / Two Track Bottom": {"prod": (240, "Production rate below 240 Prod/hour"), "rec": (80, "Recovery below 80%"), "spd": (4.5, "Speed less than 4.5 mm")},
    "Handle / Interlock / Top Bottom": {"prod": (220, "Production rate below 220 Prod/hour"), "rec": (80, "Recovery below 80%"), "spd": (4.7, "Speed less than 4.7 mm")},
    "Three Track Top": {"prod": (250, "Production rate below 250 Prod/hour"), "rec": (80, "Recovery below 80%"), "spd": (4.5, "Speed less than 4.5 mm")},
    "Three Track Bottom": {"prod": (330, "Production rate below 330 Prod/hour"), "rec": (82, "Recovery below 82%"), "spd": (5.0, "Speed less than 5.0 mm")},
}

DEFAULT_P2_RULES = {
    "Equal Angle / Unequal Angle": {"prod": (550, "Production rate below 550 Prod/hour"), "rec": (80, "Recovery below 80%"), "spd": (5, "Speed less than 5 mm")},
    "Rectangular Tube": {"prod": (300, "Production rate below 300 Prod/hour"), "rec": (80, "Recovery below 80%"), "spd": (3.5, "Speed less than 3.5 mm")},
    "Square Tube": {"prod": (400, "Production rate below 400 Prod/hour"), "rec": (80, "Recovery below 80%"), "spd": (3.5, "Speed less than 3.5 mm")},
    "Round Pipe": {"prod": (450, "Production rate below 450 Prod/hour"), "rec": (80, "Recovery below 80%"), "spd": (4, "Speed less than 4 mm")},
    "Handle": {"prod": (450, "Production rate below 450 Prod/hour"), "rec": (79, "Recovery below 79%"), "spd": (3.8, "Speed less than 3.8 mm")},
    "Interlock": {"prod": (450, "Production rate below 450 Prod/hour"), "rec": (79, "Recovery below 79%"), "spd": (3.7, "Speed less than 3.7 mm")},
    "Top Bottom": {"prod": (450, "Production rate below 450 Prod/hour"), "rec": (79, "Recovery below 79%"), "spd": (3.8, "Speed less than 3.8 mm")},
    "Glass Meeting": {"prod": (450, "Production rate below 450 Prod/hour"), "rec": (79, "Recovery below 79%"), "spd": (3.5, "Speed less than 3.5 mm")},
    "Bearing Bottom": {"prod": (450, "Production rate below 450 Prod/hour"), "rec": (79, "Recovery below 79%"), "spd": (3.4, "Speed less than 3.4 mm")},
    "Single Track Top": {"prod": (525, "Production rate below 525 Prod/hour"), "rec": (80, "Recovery below 80%"), "spd": (5.5, "Speed less than 5.5 mm")},
    "Single Track Bottom": {"prod": (525, "Production rate below 525 Prod/hour"), "rec": (80, "Recovery below 80%"), "spd": (5.25, "Speed less than 5.25 mm")},
    "Two Track Top": {"prod": (525, "Production rate below 525 Prod/hour"), "rec": (80, "Recovery below 80%"), "spd": (5.0, "Speed less than 5.0 mm")},
    "Two Track Bottom": {"prod": (525, "Production rate below 525 Prod/hour"), "rec": (80, "Recovery below 80%"), "spd": (4.9, "Speed less than 4.9 mm")},
    "Three Track Top": {"prod": (525, "Production rate below 525 Prod/hour"), "rec": (80, "Recovery below 80%"), "spd": (4.5, "Speed less than 4.5 mm")},
    "Three Track Bottom": {"prod": (525, "Production rate below 525 Prod/hour"), "rec": (80, "Recovery below 80%"), "spd": (4.5, "Speed less than 4.5 mm")},
    "Four Track Top": {"prod": (525, "Production rate below 525 Prod/hour"), "rec": (80, "Recovery below 80%"), "spd": (4.5, "Speed less than 4.5 mm")},
    "Four Track Bottom": {"prod": (525, "Production rate below 525 Prod/hour"), "rec": (80, "Recovery below 80%"), "spd": (4.5, "Speed less than 4.5 mm")},
    "Dumal / Dumal 2 Track / Dumal 3 Track / Dumal 4 Track": {"prod": (470, "Production rate below 470 Prod/hour"), "rec": (80, "Recovery below 80%"), "spd": (3.5, "Speed less than 3.5 mm")},
    "Curtain Wall": {"prod": (500, "Production rate below 500 Prod/hour"), "rec": (80, "Recovery below 80%"), "spd": (4.0, "Speed less than 4.0 mm")},
    "52 MM / 42 MM": {"prod": (520, "Production rate below 520 Prod/hour"), "rec": (80, "Recovery below 80%"), "spd": (4.0, "Speed less than 4.0 mm")},
    "Mini Dumal": {"prod": (400, "Production rate below 400 Prod/hour"), "rec": (80, "Recovery below 80%"), "spd": (3.0, "Speed less than 3.0 mm")},
    "Dumal Shutter": {"prod": (380, "Production rate below 380 Prod/hour"), "rec": (80, "Recovery below 80%"), "spd": (3.0, "Speed less than 3.0 mm")},
    "40 MM Outer Clip Mullion": {"prod": (280, "Production rate below 280 Prod/hour"), "rec": (80, "Recovery below 80%"), "spd": (5.7, "Speed less than 5.7 mm")},
}

DEFAULT_FAMILY_KEYWORDS = {
    # P2 Specific
    "40 mm outer clip mullion": "40 MM Outer Clip Mullion",
    "dumal shutter": "Dumal Shutter",
    "dumal 2 track": "Dumal / Dumal 2 Track / Dumal 3 Track / Dumal 4 Track",
    "dumal 3 track": "Dumal / Dumal 2 Track / Dumal 3 Track / Dumal 4 Track",
    "dumal 4 track": "Dumal / Dumal 2 Track / Dumal 3 Track / Dumal 4 Track",
    "dumal": "Dumal / Dumal 2 Track / Dumal 3 Track / Dumal 4 Track",
    "glass meeting": "Glass Meeting",
    "bearing bottom": "Bearing Bottom",
    "four track top": "Four Track Top",
    "four track bottom": "Four Track Bottom",
    "curtain wall": "Curtain Wall",
    "52 mm": "52 MM / 42 MM",
    "42 mm": "52 MM / 42 MM",
    # P1 and P2
    "mini dumal": "Mini Dumal",
    "three track top": "Three Track Top",
    "three track bottom": "Three Track Bottom",
    "two track top": "Two Track Top",
    "two track bottom": "Two Track Bottom",
    "single track top": "Single Track Top",
    "single track bottom": "Single Track Bottom",
    "handle": "Handle",
    "interlock": "Interlock",
    "top bottom": "Top Bottom",
    "equal angle": "Equal Angle / Unequal Angle",
    "unequal angle": "Equal Angle / Unequal Angle",
    "r.t": "Rectangular Tube", # Abbreviation
    "rectangular tube": "Rectangular Tube",
    "s.t": "Square Tube", # Abbreviation
    "square tube": "Square Tube",
    "round tube": "Round Pipe", # Alias
    "round pipe": "Round Pipe",
    # P1 Specific
    "40 mm clip": "40 mm Clip",
    "40 mm outer": "40 mm Outer",
    "40 mm frame": "40 mm Frame",
}

DEFAULT_DEPARTMENT_NUMBER_MAP = {
    1: "Tool Room",
    2: "Production Department",
    3: "Tool Room and Production Department",
    4: "Foundry",
    5: "Maintainance",
    0: "No Department"
}

# --- Firestore Initialization ---

@st.cache_resource
def get_firestore_db():
    """Initializes and returns a Firestore client, or None if setup fails."""
    if firestore is None:
        return None
    try:
        # Get credentials from Streamlit Secrets
        creds_json = {
            "type": st.secrets["firestore_service_account"]["type"],
            "project_id": st.secrets["firestore_service_account"]["project_id"],
            "private_key_id": st.secrets["firestore_service_account"]["private_key_id"],
            "private_key": st.secrets["firestore_service_account"]["private_key"],
            "client_email": st.secrets["firestore_service_account"]["client_email"],
            "client_id": st.secrets["firestore_service_account"]["client_id"],
            "auth_uri": st.secrets["firestore_service_account"]["auth_uri"],
            "token_uri": st.secrets["firestore_service_account"]["token_uri"],
            "auth_provider_x509_cert_url": st.secrets["firestore_service_account"]["auth_provider_x509_cert_url"],
            "client_x509_cert_url": st.secrets["firestore_service_account"]["client_x509_cert_url"]
        }
        
        creds = service_account.Credentials.from_service_account_info(creds_json)
        db = firestore.Client(credentials=creds, project=creds_json["project_id"])
        return db
    except Exception as e:
        st.error(f"Failed to connect to Firebase. Please check your st.secrets setup. Error: {e}")
        return None

def load_rules(db):
    """Loads rules from Firestore into session_state, seeding if necessary."""
    if 'rules_loaded' in st.session_state:
        return

    if db is None:
        # No database connection, use local defaults
        st.session_state.p1_rules = DEFAULT_P1_RULES.copy()
        st.session_state.p2_rules = DEFAULT_P2_RULES.copy()
        st.session_state.family_keywords = DEFAULT_FAMILY_KEYWORDS.copy()
        
        # Convert department map keys to str for JSON compatibility if needed
        st.session_state.department_map = {str(k): v for k, v in DEFAULT_DEPARTMENT_NUMBER_MAP.items()}
        st.session_state.rules_loaded = True
        return

    # Use a single document to store all rules
    doc_ref = db.collection("press_analyzer_rules").document("all_rules")
    try:
        doc = doc_ref.get()
        if doc.exists:
            # Load rules from DB
            rules = doc.to_dict()
            st.session_state.p1_rules = rules.get("p1_rules", DEFAULT_P1_RULES)
            st.session_state.p2_rules = rules.get("p2_rules", DEFAULT_P2_RULES)
            st.session_state.family_keywords = rules.get("family_keywords", DEFAULT_FAMILY_KEYWORDS)
            
            # Firestore stores keys as strings, so we ensure keys are str
            db_dept_map = rules.get("department_map", DEFAULT_DEPARTMENT_NUMBER_MAP)
            st.session_state.department_map = {str(k): v for k, v in db_dept_map.items()}
        else:
            # No rules in DB, seed with defaults
            st.session_state.p1_rules = DEFAULT_P1_RULES.copy()
            st.session_state.p2_rules = DEFAULT_P2_RULES.copy()
            st.session_state.family_keywords = DEFAULT_FAMILY_KEYWORDS.copy()
            st.session_state.department_map = {str(k): v for k, v in DEFAULT_DEPARTMENT_NUMBER_MAP.items()}
            
            # Save defaults to Firestore
            doc_ref.set({
                "p1_rules": st.session_state.p1_rules,
                "p2_rules": st.session_state.p2_rules,
                "family_keywords": st.session_state.family_keywords,
                "department_map": st.session_state.department_map
            })
        
        st.session_state.rules_loaded = True
    except Exception as e:
        st.error(f"Error loading rules from Firestore: {e}. Using default rules.")
        # Fallback to defaults
        if 'p1_rules' not in st.session_state:
            st.session_state.p1_rules = DEFAULT_P1_RULES.copy()
            st.session_state.p2_rules = DEFAULT_P2_RULES.copy()
            st.session_state.family_keywords = DEFAULT_FAMILY_KEYWORDS.copy()
            st.session_state.department_map = {str(k): v for k, v in DEFAULT_DEPARTMENT_NUMBER_MAP.items()}
        st.session_state.rules_loaded = True

def save_rules(db):
    """Saves the current session_state rules to Firestore."""
    if db is None:
        st.warning("Running in local mode. Rule changes will not be saved.")
        return

    try:
        doc_ref = db.collection("press_analyzer_rules").document("all_rules")
        doc_ref.set({
            "p1_rules": st.session_state.p1_rules,
            "p2_rules": st.session_state.p2_rules,
            "family_keywords": st.session_state.family_keywords,
            "department_map": st.session_state.department_map
        })
        st.sidebar.success("Rules saved to database!")
    except Exception as e:
        st.sidebar.error(f"Failed to save rules: {e}")

# --- Helper Functions (Now use st.session_state) ---

def get_die_family(die_name):
    """Maps a die name string to its standardized family name using session_state rules."""
    if not isinstance(die_name, str):
        return None
    name_low = die_name.lower()
    # Iterate through keywords from session_state
    for keyword, family in st.session_state.family_keywords.items():
        if keyword in name_low:
            return family
    return None # No match found

def apply_flagging_rules(row, press_val):
    """Applies the flagging logic to a single DataFrame row using session_state rules."""
    family = row['Die Family']
    prod = row['PROD/HOUR']
    rec = row['RECOVERY %']
    spd = row['Speed(mm)']
    reasons = []
    ruleset = None
    
    press_val_std = "N/A"
    if pd.notna(press_val):
        press_val_std = str(press_val).upper().strip()

    if "P1" in press_val_std:
        if family in ["Handle", "Interlock", "Top Bottom"]:
            family = "Handle / Interlock / Top Bottom"
        if family in ["Single Track Top", "Single Track Bottom"]:
            family = "Single Track Top / Single Track Bottom"
        if family in ["Two Track Top", "Two Track Bottom"]:
             family = "Two Track Top / Two Track Bottom"
        ruleset = st.session_state.p1_rules.get(family)
        
    elif "P2" in press_val_std:
        ruleset = st.session_state.p2_rules.get(family)
        
    else:
        return f"Press '{press_val}' not identified as P1 or P2"

    if ruleset:
        # Helper to safely get rule, as DB structure might be list
        def get_rule_val(rule_tuple_or_list):
            return rule_tuple_or_list[0]

        if pd.notna(prod) and 'prod' in ruleset and prod < get_rule_val(ruleset['prod']):
            reasons.append(ruleset['prod'][1])
        if pd.notna(rec) and 'rec' in ruleset and rec < get_rule_val(ruleset['rec']):
            reasons.append(ruleset['rec'][1])
        if pd.notna(spd) and 'spd' in ruleset and spd < get_rule_val(ruleset['spd']):
            reasons.append(ruleset['spd'][1])
    else:
        if family: 
            return f"No rules defined for family: {family}"
        else: 
             return "Die family not found"

    return " and ".join(reasons) if reasons else None

@st.cache_data
def to_excel(df):
    """Caches the conversion of a DataFrame to an Excel file in memory."""
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Flagged_Dies')
    processed_data = output.getvalue()
    return processed_data

def normalize_sheet_name(name: str) -> str:
    """Converts sheet names to a standard format for comparison."""
    return re.sub(r"\s+", "", str(name).strip().lower())

# --- UI Functions for Sidebar ---

def build_sidebar(db):
    """Creates the sidebar UI for rule management."""
    st.sidebar.title("Rule Management")
    st.sidebar.info("Rule changes are saved to a central database.")
    
    # Expander for Family Keywords
    with st.sidebar.expander("Family Keyword Mapping", expanded=False):
        st.write("Map keywords (e.g., 'r.t') to a family (e.g., 'Rectangular Tube').")
        
        # Display current keywords in a table
        kw_df = pd.DataFrame(
            st.session_state.family_keywords.items(), 
            columns=['Keyword', 'Family']
        ).sort_values(by="Keyword")
        st.dataframe(kw_df, use_container_width=True)

        # Form to add/update a keyword
        with st.form("add_keyword_form"):
            st.write("**Add / Update Keyword**")
            kw = st.text_input("Keyword (all lowercase)").strip().lower()
            fam = st.text_input("Family Name (exact match to Flagging Rule Family)")
            add_kw = st.form_submit_button("Save Keyword")
            
            if add_kw and kw and fam:
                st.session_state.family_keywords[kw] = fam
                save_rules(db)
                st.success(f"Saved keyword '{kw}' -> '{fam}'")
                st.rerun() # Use rerun to update the UI

        # Form to remove a keyword
        with st.form("remove_keyword_form"):
            st.write("**Remove Keyword**")
            kw_to_remove = st.selectbox("Select keyword to remove", options=kw_df['Keyword'])
            remove_kw = st.form_submit_button("Remove Keyword", type="primary")

            if remove_kw and kw_to_remove:
                if kw_to_remove in st.session_state.family_keywords:
                    del st.session_state.family_keywords[kw_to_remove]
                    save_rules(db)
                    st.success(f"Removed keyword '{kw_to_remove}'")
                    st.rerun()
                else:
                    st.error("Keyword not found.")

    # Expander for P1 Rules
    with st.sidebar.expander("P1 Flagging Rules", expanded=False):
        build_flagging_rule_ui("P1", db)
        
    # Expander for P2 Rules
    with st.sidebar.expander("P2 Flagging Rules", expanded=False):
        build_flagging_rule_ui("P2", db)

def build_flagging_rule_ui(press_type, db):
    """Reusable UI for editing P1 or P2 rules."""
    rules_key = f"{press_type.lower()}_rules" # e.g., "p1_rules"
    
    st.write(f"Edit flagging thresholds for Press {press_type}.")
    
    # Display current rules
    st.json(st.session_state[rules_key], expanded=False)
    
    st.write(f"**Add / Edit {press_type} Family Rule**")
    
    # --- BUG FIX: Selectbox is now OUTSIDE the form ---
    # This allows the app to rerun and show the conditional text_input
    family_list = sorted(st.session_state[rules_key].keys())
    family_to_edit = st.selectbox(
        "Select Family to Edit", 
        options=family_list + ["-- ADD NEW FAMILY --"], 
        key=f"{rules_key}_family_select"
    )
    # --- END BUG FIX ---
    
    with st.form(f"edit_{rules_key}_form"):
        
        # --- BUG FIX: Conditional text_input is INSIDE the form ---
        new_family_name = ""
        if family_to_edit == "-- ADD NEW FAMILY --":
            new_family_name = st.text_input("New Family Name")
        # --- END BUG FIX ---
        
        # Get current values if editing
        current_vals = {"prod": (0, ""), "rec": (0, ""), "spd": (0.0, "")}
        if family_to_edit != "-- ADD NEW FAMILY --" and family_to_edit in st.session_state[rules_key]:
            current_vals = st.session_state[rules_key][family_to_edit]

        # Ensure current_vals has the right structure
        prod_val = current_vals.get('prod', [0, ""])[0]
        rec_val = current_vals.get('rec', [0, ""])[0]
        spd_val = current_vals.get('spd', [0.0, ""])[0]

        # Inputs for thresholds
        prod_thresh = st.number_input("PROD/HOUR <", value=prod_val, key=f"{rules_key}_prod")
        rec_thresh = st.number_input("RECOVERY % <", value=rec_val, key=f"{rules_key}_rec")
        spd_thresh = st.number_input("Speed (mm) <", value=float(spd_val), format="%.2f", key=f"{rules_key}_spd")
        
        save_rule = st.form_submit_button("Save Rule")
        
        if save_rule:
            final_family_name = new_family_name.strip() if family_to_edit == "-- ADD NEW FAMILY --" else family_to_edit
            
            if not final_family_name:
                st.error("Family Name cannot be empty.")
                return # Changed from st.stop() to return

            # Create the new rule entry
            new_rule = {
                "prod": (prod_thresh, f"Production rate below {prod_thresh} Prod/hour"),
                "rec": (rec_thresh, f"Recovery below {rec_thresh}%"),
                "spd": (spd_thresh, f"Speed less than {spd_thresh} mm")
            }
            
            # Update session state and save
            st.session_state[rules_key][final_family_name] = new_rule
            save_rules(db)
            st.success(f"Saved rules for family: {final_family_name}")
            st.rerun()

# --- Streamlit App ---

def main():
    st.set_page_config(page_title="PRESS PROD SHEET Viewer", layout="wide")
    st.title("📘 PRESS PROD SHEET Viewer & Analyst")

    # --- Initialize App ---
    db = get_firestore_db()
    load_rules(db)
    
    # --- Build Sidebar UI ---
    build_sidebar(db)

    # --- Main Page UI ---
    st.markdown("""
    Upload an **.xlsx** file.  
    This app will analyze **'PRESS PROD SHEET'** based on customizable rules
    and display a report of **flagged dies**.
    """)

    uploaded = st.file_uploader("Upload Excel file (.xlsx only)", type=["xlsx"])

    if uploaded:
        # Analysis logic only runs if a file is uploaded
        try:
            xls = pd.ExcelFile(uploaded, engine="openpyxl")
            target_sheet = None
            for s in xls.sheet_names:
                if normalize_sheet_name(s) == "pressprod":
                    target_sheet = s
                    break

            if target_sheet is None:
                st.error("❌ No sheet named 'PRESS PROD SHEET' found (case/space-insensitive).")
                return

            st.success(f"✅ Found sheet: **{target_sheet}**. Extracting and analyzing data...")
            
            # --- Load Remark Mapping ---
            remark_map_sheet_name = None
            for s in xls.sheet_names:
                norm_s = normalize_sheet_name(s)
                if norm_s == "sheet1" or norm_s == "mapping": 
                    remark_map_sheet_name = s
                    break
            
            remark_to_number_map = {}
            if remark_map_sheet_name:
                try:
                    df_map = pd.read_excel(uploaded, sheet_name=remark_map_sheet_name, engine="openpyxl", header=None)
                    start_row = -1
                    # Try to find the header row "Remarks"
                    for i, row in df_map.iterrows():
                        # Check if 'Remarks' is in the row values
                        if 'Remarks' in row.astype(str).values:
                            start_row = i + 1
                            break
                    
                    if start_row != -1:
                        # Found header, select data from start_row
                        df_map = df_map.iloc[start_row:, [0, 1]]
                    else:
                        # Fallback: assume data starts from row index 4 (row 5 in Excel)
                        st.warning(f"Could not find 'Remarks' header in {remark_map_sheet_name}, falling back to default row 5.")
                        # Reread and slice
                        df_map = pd.read_excel(uploaded, sheet_name=remark_map_sheet_name, engine="openpyxl", header=None)
                        df_map = df_map.iloc[4:, [0, 1]]

                    df_map = df_map.dropna(subset=[0])
                    df_map = df_map.drop_duplicates(subset=[0], keep='first')
                    df_map[0] = df_map[0].astype(str).str.strip().str.lower()
                    df_map[1] = pd.to_numeric(df_map[1], errors='coerce')
                    # Convert number to string for lookup, as map keys are strings
                    remark_to_number_map = pd.Series(df_map[1].values, index=df_map[0]).dropna().astype(int).astype(str).to_dict()
                    st.success(f"✅ Found and loaded remark mapping from sheet: **{remark_map_sheet_name}**")
                except Exception as e:
                    st.error(f"❌ Error loading '{remark_map_sheet_name}' for remark mapping. Error: {e}")
            else:
                st.error("❌ Could not find 'Sheet1' or 'Mapping' sheet for remark-to-department.")

            # --- Read Main Sheet Data ---
            df_full = pd.read_excel(uploaded, sheet_name=target_sheet, engine="openpyxl", header=None)

            # --- Extract Header Info ---
            date_val, press_val, operator_val, supervisor_val = "N/A", "N/A", "N/A", "N/A"
            try:
                date_val = df_full.iloc[4, 1]
                press_val = df_full.iloc[4, 2]
                operator_val = df_full.iloc[4, 3]
                supervisor_val = df_full.iloc[4, 11]
                
                date_val = str(date_val).split(" ")[0] if pd.notna(date_val) else "N/A"
                press_val = str(press_val) if pd.notna(press_val) else "N/A"
                operator_val = str(operator_val) if pd.notna(operator_val) else "N/A"
                supervisor_val = str(supervisor_val) if pd.notna(supervisor_val) else "N/A"

                st.subheader("Sheet Information")
                col1, col2, col3, col4 = st.columns(4)
                col1.metric("Date", date_val)
                col2.metric("Press", press_val)
                col3.metric("Operator", operator_val)
                col4.metric("Supervisor", supervisor_val)
                st.divider()
            except Exception as e:
                st.warning(f"Could not extract header info. Error: {e}")

            # --- Parse Main Data Table ---
            try:
                header_row_6 = df_full.iloc[5]
                header_row_7 = df_full.iloc[6]
                new_cols = []
                for i in range(len(header_row_6)):
                    col_6_val = str(header_row_6[i])
                    col_7_val = str(header_row_7[i])
                    if pd.notna(header_row_7[i]) and "unnamed" not in col_7_val.lower():
                        new_cols.append(col_7_val.strip())
                    elif pd.notna(header_row_6[i]) and "unnamed" not in col_6_val.lower():
                        new_cols.append(col_6_val.strip())
                    else:
                        new_cols.append(f"col_{i}")
                
                df_data = df_full.iloc[8:].copy()
                df_data.columns = new_cols
                df_data = df_data.dropna(how='all')

                # --- 6. Filter and Pre-process Data ---
                if 'DIE NO.' not in df_data.columns or 'DIE NAME' not in df_data.columns:
                    st.error("Error: Missing 'DIE NO.' or 'DIE NAME' column. Cannot proceed.")
                    return
                
                original_row_count = len(df_data)
                df_data['DIE NO.'] = pd.to_numeric(df_data['DIE NO.'], errors='coerce')
                df_data = df_data[pd.notna(df_data['DIE NO.'])]
                df_data = df_data[pd.notna(df_data['DIE NAME'])]
                df_data = df_data[df_data['DIE NAME'].astype(str).str.strip() != '']
                
                if len(df_data) < original_row_count:
                    st.info(f"Filtered out {original_row_count - len(df_data)} rows with invalid 'DIE NO.' or 'DIE NAME'.")

                # --- 7. Perform Analysis ---
                cols_to_convert = ['PROD/HOUR', 'RECOVERY %', 'Speed(mm)']
                for col in cols_to_convert:
                    if col in df_data.columns:
                        df_data[col] = pd.to_numeric(df_data[col], errors='coerce')
                    else:
                        st.error(f"Error: Missing required column '{col}'.")
                        return
                
                df_data['Die Family'] = df_data['DIE NAME'].apply(get_die_family)
                
                # Map remark string to number (which is now a string)
                df_data['Department_Number'] = df_data['REMARK'].astype(str).str.strip().str.lower().map(remark_to_number_map)
                
                # Use department map from session_state (which has string keys)
                df_data['Department'] = df_data['Department_Number'].map(st.session_state.department_map)
                
                df_data['Flagging Reason'] = df_data.apply(apply_flagging_rules, axis=1, press_val=press_val)
                
                # --- 8. Create and Display Flagged Report ---
                flagged_df = df_data[df_data['Flagging Reason'].notna()].copy()
                st.subheader("🚩 Flagged Dies Report")
                
                if flagged_df.empty:
                    st.success("✅ No dies were flagged based on the provided rules.")
                else:
                    flagged_df['Date'] = date_val
                    flagged_df['Press'] = press_val
                    flagged_df['Operator Name'] = operator_val
                    flagged_df['Supervisor Name'] = supervisor_val
                    
                    output_cols = ['Date', 'Press', 'DIE NO.', 'DIE NAME', 'Flagging Reason', 
                                   'Operator Name', 'Supervisor Name', 'REMARK', 'Department']
                    final_report_cols = [col for col in output_cols if col in flagged_df.columns]
                    final_df = flagged_df[final_report_cols].rename(columns={
                        'DIE NO.': 'Die Number',
                        'DIE NAME': 'Profile Name',
                        'REMARK': 'Remark'
                    })
                    
                    st.metric("Total Flagged Dies", len(final_df))
                    st.dataframe(final_df, use_container_width=True, hide_index=True)
                    
                    excel_data = to_excel(final_df)
                    st.download_button(
                        label="📥 Download Flagged Report as Excel",
                        data=excel_data,
                        file_name=f"flagged_dies_report_{date_val.replace('/', '-')}.xlsx",
                        mime="application/vnd.ms-excel"
                    )
            except Exception as e:
                st.error(f"Error processing the data table: {e}")
                st.exception(e)

        except Exception as e:
            st.exception(e)
    else:
        st.info("Please upload an Excel (.xlsx) file to begin analysis.")

if __name__ == "__main__":
    main()
