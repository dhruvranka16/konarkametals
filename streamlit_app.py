import re
import pandas as pd
import streamlit as st
from io import BytesIO

# --- Rule Definitions ---
# These dictionaries define the performance thresholds for flagging dies.
# Each key is a 'Die Family', and the value is a dictionary of metrics ('prod', 'rec', 'spd').
# For each metric, it stores a tuple: (threshold, flag_message)

P1_RULES = {
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

P2_RULES = {
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

# Keywords for family mapping, ordered from most specific to most general.
# This dictionary maps keywords found in 'DIE NAME' to a standardized 'Die Family'.
# The order is important: more specific names must come before general ones.
FAMILY_KEYWORDS = {
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

# --- Department Number to Name Mapping ---
# Maps the numeric code from 'Sheet1'/'Mapping' to a readable department name.
DEPARTMENT_NUMBER_MAP = {
    1: "Tool Room",
    2: "Production Department",
    3: "Tool Room and Production Department",
    4: "Foundry",
    5: "Maintainance",
    0: "No Department" # Handle '0' as a specific category
}
# --- END NEW MAPPING ---

# --- Helper Functions ---

def get_die_family(die_name):
    """Maps a die name string to its standardized family name."""
    if not isinstance(die_name, str):
        return None
    name_low = die_name.lower()
    # Iterate through keywords and return the first match
    for keyword, family in FAMILY_KEYWORDS.items():
        if keyword in name_low:
            return family
    return None # No match found

def apply_flagging_rules(row, press_val):
    """Applies the flagging logic to a single DataFrame row."""
    family = row['Die Family']
    prod = row['PROD/HOUR']
    rec = row['RECOVERY %']
    spd = row['Speed(mm)']
    reasons = []
    ruleset = None
    
    # Standardize press value to handle variations (e.g., "P1", "P1 A", "p1")
    press_val_std = "N/A"
    if pd.notna(press_val):
        press_val_std = str(press_val).upper().strip()

    if "P1" in press_val_std:
        # P1 has combined rules for some families
        if family in ["Handle", "Interlock", "Top Bottom"]:
            family = "Handle / Interlock / Top Bottom"
        if family in ["Single Track Top", "Single Track Bottom"]:
            family = "Single Track Top / Single Track Bottom"
        if family in ["Two Track Top", "Two Track Bottom"]:
             family = "Two Track Top / Two Track Bottom"
        ruleset = P1_RULES.get(family)
        
    elif "P2" in press_val_std:
        # P2 rules are more direct
        ruleset = P2_RULES.get(family)
        
    else:
        # Flag if press value is missing or not 'P1'/'P2'
        return f"Press '{press_val}' not identified as P1 or P2"

    if ruleset:
        # Check each rule for the matched family
        if pd.notna(prod) and 'prod' in ruleset and prod < ruleset['prod'][0]:
            reasons.append(ruleset['prod'][1])
        if pd.notna(rec) and 'rec' in ruleset and rec < ruleset['rec'][0]:
            reasons.append(ruleset['rec'][1])
        if pd.notna(spd) and 'spd' in ruleset and spd < ruleset['spd'][0]:
            reasons.append(ruleset['spd'][1])
    else:
        # Flag if a family was identified but has no rules
        if family: 
            return f"No rules defined for family: {family}"
        else: 
            # Flag if the die name didn't map to any family
             return "Die family not found"

    # Join all triggered flag reasons, or return None if no flags
    return " and ".join(reasons) if reasons else None

@st.cache_data
def to_excel(df):
    """
    Caches the conversion of a DataFrame to an Excel file in memory.
    This prevents re-computing the Excel file on every app rerun.
    """
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Flagged_Dies')
    processed_data = output.getvalue()
    return processed_data

def normalize_sheet_name(name: str) -> str:
    """
    Converts sheet names to a standard format for comparison.
    (e.g., "PRESS PROD SHEET" -> "pressprod")
    """
    # Replaces all whitespace with empty string, converts to lower
    return re.sub(r"\s+", "", str(name).strip().lower())

# --- Streamlit App ---

# Page configuration (must be the first st command)
st.set_page_config(page_title="PRESS PROD SHEET Viewer", layout="wide")
st.title("📘 PRESS PROD SHEET Viewer & Analyst")

st.markdown("""
Upload an **.xlsx** file.  
This app will automatically find the sheet named **'PRESS PROD SHEET'**,
analyze it based on performance rules, and display a report of **flagged dies**.
""")

# File uploader widget
uploaded = st.file_uploader("Upload Excel file (.xlsx only)", type=["xlsx"])

if uploaded:
    try:
        # Load the entire Excel file into memory
        xls = pd.ExcelFile(uploaded, engine="openpyxl")

        # --- 1. Find the main 'PRESS PROD SHEET' ---
        target_sheet = None
        for s in xls.sheet_names:
            if normalize_sheet_name(s) == "pressprod":
                target_sheet = s
                break

        if target_sheet is None:
            st.error("❌ No sheet named 'PRESS PROD SHEET' found (case/space-insensitive). Please rename and upload again.")
        else:
            st.success(f"✅ Found sheet: **{target_sheet}**. Extracting and analyzing data...")
            
            # --- 2. Load Remark Mapping from 'Sheet1' or 'Mapping' ---
            remark_map_sheet_name = None
            for s in xls.sheet_names:
                norm_s = normalize_sheet_name(s)
                if norm_s == "mapping": 
                    remark_map_sheet_name = s
                    break
            
            remark_to_number_map = {}
            if remark_map_sheet_name:
                try:
                    # Read mapping sheet, no header assumed
                    df_map = pd.read_excel(uploaded, sheet_name=remark_map_sheet_name, engine="openpyxl", header=None)
                    
                    # Find the header "Remarks" to dynamically locate data start
                    start_row = -1
                    for i, row in df_map.iterrows():
                        if 'Remarks' in row.astype(str).values:
                            start_row = i + 1 # Data starts on the next row
                            break
                    
                    if start_row != -1:
                        df_map = df_map.iloc[start_row:, [0, 1]] # Data is in col 0 (remark) and col 1 (number)
                    else:
                        # Fallback if "Remarks" header not found, use original assumption
                        st.warning("Could not find 'Remarks' header in mapping sheet, falling back to row 5.")
                        df_map = pd.read_excel(uploaded, sheet_name=remark_map_sheet_name, engine="openpyxl", header=None)
                        df_map = df_map.iloc[4:, [0, 1]] # Data is in col 0 (remark) and col 1 (number)

                    df_map = df_map.dropna(subset=[0]) # Drop rows where remark is empty
                    df_map = df_map.drop_duplicates(subset=[0], keep='first') # Handle duplicate remarks
                    
                    # Standardize remarks for matching (lower, stripped)
                    df_map[0] = df_map[0].astype(str).str.strip().str.lower()
                    # Convert number, coercing errors (non-numbers become NaN)
                    df_map[1] = pd.to_numeric(df_map[1], errors='coerce')
                    
                    # Create the dictionary, dropping any NaN values
                    remark_to_number_map = pd.Series(df_map[1].values, index=df_map[0]).dropna().to_dict()
                    st.success(f"✅ Found and loaded remark mapping from sheet: **{remark_map_sheet_name}**")
                except Exception as e:
                    st.error(f"❌ Error loading '{remark_map_sheet_name}' for remark mapping. Department analysis will be incomplete. Error: {e}")
            else:
                st.error("❌ Could not find 'Sheet1' or 'Mapping' sheet for remark-to-department. Department analysis will be incomplete.")

            # --- 3. Read the main sheet data ---
            # Read without a header to parse it manually
            df_full = pd.read_excel(uploaded, sheet_name=target_sheet, engine="openpyxl", header=None)

            # --- 4. Extract single-value header info ---
            date_val, press_val, operator_val, supervisor_val = "N/A", "N/A", "N/A", "N/A"
            try:
                # Values are at fixed locations (based on file structure)
                # Row 5 (index 4), Cols B, C, D, L (index 1, 2, 3, 11)
                date_val = df_full.iloc[4, 1]
                press_val = df_full.iloc[4, 2]
                operator_val = df_full.iloc[4, 3]
                supervisor_val = df_full.iloc[4, 11]
                
                # Clean up values for display
                date_val = str(date_val) if pd.notna(date_val) else "N/A"
                if " " in date_val: # Handle Excel datetime format
                   date_val = date_val.split(" ")[0]
                
                press_val = str(press_val) if pd.notna(press_val) else "N/A"
                operator_val = str(operator_val) if pd.notna(operator_val) else "N/A"
                supervisor_val = str(supervisor_val) if pd.notna(supervisor_val) else "N/A"

                # Display header info in metric cards
                st.subheader("Sheet Information")
                col1, col2, col3, col4 = st.columns(4)
                col1.metric("Date", date_val)
                col2.metric("Press", press_val)
                col3.metric("Operator", operator_val)
                col4.metric("Supervisor", supervisor_val)
                st.divider()

            except Exception as e:
                st.warning(f"Could not extract header info (Date, Press, etc.). The sheet structure might be different. Error: {e}")

            # --- 5. Parse the main data table (with multi-row headers) ---
            try:
                # Headers are on Row 6 (index 5) and Row 7 (index 6)
                header_row_6 = df_full.iloc[5]
                header_row_7 = df_full.iloc[6]

                # Combine the two header rows into a single list
                new_cols = []
                for i in range(len(header_row_6)):
                    col_6_val = str(header_row_6[i])
                    col_7_val = str(header_row_7[i])
                    
                    # Prioritize Row 7, then Row 6, then a default name
                    if pd.notna(header_row_7[i]) and "unnamed" not in col_7_val.lower():
                        new_cols.append(col_7_val.strip())
                    elif pd.notna(header_row_6[i]) and "unnamed" not in col_6_val.lower():
                        new_cols.append(col_6_val.strip())
                    else:
                        new_cols.append(f"col_{i}")
                
                # Data starts from Row 9 (index 8)
                df_data = df_full.iloc[8:].copy()
                df_data.columns = new_cols
                df_data = df_data.dropna(how='all') # Drop fully empty rows

                # --- 6. Filter and Pre-process Data ---
                
                # Check for essential columns for analysis
                if 'DIE NO.' not in df_data.columns:
                    st.error("Error: Missing required column 'DIE NO.' for analysis. Cannot proceed.")
                    st.stop()
                if 'DIE NAME' not in df_data.columns:
                    st.error("Error: Missing required column 'DIE NAME' for analysis. Cannot proceed.")
                    st.stop()
                
                original_row_count = len(df_data)

                # Force 'DIE NO.' to numeric; non-numeric values become NaN
                df_data['DIE NO.'] = pd.to_numeric(df_data['DIE NO.'], errors='coerce')

                # Filter out rows where 'DIE NO.' is missing (NaN)
                df_data = df_data[pd.notna(df_data['DIE NO.'])]
                
                # Filter out rows where 'DIE NAME' is missing or blank
                df_data = df_data[pd.notna(df_data['DIE NAME'])]
                df_data = df_data[df_data['DIE NAME'].astype(str).str.strip() != '']

                filtered_row_count = len(df_data)
                
                if original_row_count > filtered_row_count:
                    st.info(f"Filtered out {original_row_count - filtered_row_count} rows with missing or invalid 'DIE NO.' or 'DIE NAME' before analysis.")

                # --- 7. Perform Analysis ---
                
                # A. Clean numeric data for analysis
                cols_to_convert = ['PROD/HOUR', 'RECOVERY %', 'Speed(mm)']
                for col in cols_to_convert:
                    if col in df_data.columns:
                        df_data[col] = pd.to_numeric(df_data[col], errors='coerce')
                    else:
                        # Stop if a required metric column is missing
                        st.error(f"Error: Missing required column '{col}' for analysis.")
                        st.stop()
                
                # B. Apply mappings
                df_data['Die Family'] = df_data['DIE NAME'].apply(get_die_family)
                
                # Apply new department mapping
                # Map remark string (lower, stripped) to number
                df_data['Department_Number'] = df_data['REMARK'].astype(str).str.strip().str.lower().map(remark_to_number_map)
                # Map number to department name
                df_data['Department'] = df_data['Department_Number'].map(DEPARTMENT_NUMBER_MAP)
                
                # C. Apply flagging rules row by row
                df_data['Flagging Reason'] = df_data.apply(apply_flagging_rules, axis=1, press_val=press_val)
                
                # --- 8. Create and Display Flagged Report ---
                
                # Filter for rows that have a flagging reason
                flagged_df = df_data[df_data['Flagging Reason'].notna()].copy()
                
                st.subheader("🚩 Flagged Dies Report")
                
                if flagged_df.empty:
                    st.success("✅ No dies were flagged based on the provided rules.")
                else:
                    # Add the single-value header info to the report
                    flagged_df['Date'] = date_val
                    flagged_df['Press'] = press_val
                    flagged_df['Operator Name'] = operator_val
                    flagged_df['Supervisor Name'] = supervisor_val
                    
                    # Select and rename columns for the final report
                    output_cols = [
                        'Date', 'Press', 'DIE NO.', 'DIE NAME', 'Flagging Reason', 
                        'Operator Name', 'Supervisor Name', 'REMARK', 'Department'
                    ]
                    
                    # Ensure all desired columns exist before selecting
                    final_report_cols = [col for col in output_cols if col in flagged_df.columns]
                    final_df = flagged_df[final_report_cols]
                    
                    # Rename columns for a cleaner final report
                    final_df = final_df.rename(columns={
                        'DIE NO.': 'Die Number',
                        'DIE NAME': 'Profile Name',
                        'REMARK': 'Remark'
                    })
                    
                    # Display the total count and the report table
                    st.metric("Total Flagged Dies", len(final_df))
                    st.dataframe(final_df, use_container_width=True, hide_index=True)
                    
                    # Add download button
                    excel_data = to_excel(final_df)
                    st.download_button(
                        label="📥 Download Flagged Report as Excel",
                        data=excel_data,
                        file_name=f"flagged_dies_report_{date_val.replace('/', '-')}.xlsx",
                        mime="application/vnd.ms-excel"
                    )

            except Exception as e:
                st.error(f"Error processing the data table. The sheet structure might be non-standard.")
                st.exception(e)

    except Exception as e:
        # Catch-all for file loading or parsing errors
        st.exception(e)
else:
    # Initial state before any file is uploaded
    st.info("Please upload an Excel (.xlsx) file. The app will analyze the 'PRESS PROD SHEET' and report flagged dies.")

