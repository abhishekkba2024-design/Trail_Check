import re
import pandas as pd
import streamlit as st

# --------------------- Page ---------------------
st.set_page_config(page_title="Excel Header Checker (Multi-sheet, 4 schemas)", page_icon="✅", layout="centered")
st.title("✅ Excel Header Checker (Multi-sheet, 4 schemas)")

st.caption(
    "Upload an Excel file (.xlsx/.xlsm). Each sheet must match one of the 4 allowed schemas. "
    "If all sheets match, you'll get the Forms link; else you'll see precise mismatches."
)

FORMS_LINK = "https://forms.office.com/Pages/ResponsePage.aspx?id=GIKK1zVBJkCjqBzdciO01bNCosI6R5VPnjnHYY4WuWxUNzNKNURZWlJJSVdKSDVITlE2WEwzUExKRy4u"

# --------------------- Your 4 schemas ---------------------
SCHEMAS = {
    "Schema 1": [
        "S. No.", "Plant", "Process", "Shop Type", "Shop\nCode", "Shop Description", "Cost Centre",
        "Line", "Line type", "Iot Status", "Capacity per shift ( As per prod)", "Logical PSL", "Physical PSL **",
        "Operational Shifts", "Details of Deviation in Shift *", "No of Machines", "Category\n(4W/2W)",
        "Part No/\nModel No", "Part Name/\nModel Description", "Gross RM Weight\n(in kg/ Part)", "Broad Model",
        "Engine/TM", "Financial Year", "Month", "OK Volume", "Rejection Volume",
        "OEE/Line Efficiency As per Production", "Remarks ( If Any)",
    ],
    "Schema 2": [
        "S. No.", "Plant", "Maint Dept Code", "Maint Dept Description", "Cost Centre", "Type", "Shop Coverage*", "Remarks",
    ],
    "Schema 3": [
        "S. No.", "Plant", "Process", "Shop Type", "Shop\nCode/Maint Dept", "Shop Description", "Cost Centre",
        "Month", "Activity Category", "Sub Activity*", "Manpower Grade", "Headcount",
    ],
    "Schema 4": [
        "S. No.", "Plant", "Process", "Shop", "Line", "Line Type", "Model", "Fuel Type",
        "Part No/\nModel No", "Part Name/\nModel Description", "Annual Capacity as per OEE*",
        "SMM\nPer Part**", "Takt Time", "Cycle Time", "OEE Factor (%)",
    ],
}

# --------------------- Options ---------------------
match_mode = st.radio(
    "Match mode",
    ["Exact (case & spaces must match)",
     "Lenient (ignores case/extra spaces/underscores/newlines/quotes/*)"],
    index=1
)
single_schema_for_workbook = st.checkbox(
    "Require that all sheets match the **same** schema", value=False,
    help="If ON, every sheet must conform to the same schema (1, 2, 3, or 4)."
)
max_header_scan_rows = st.slider(
    "Try header row at lines (0 = first row)", min_value=0, max_value=5, value=3,
    help="The app will try these many rows as header rows and pick the best."
)

# --------------------- Utils ---------------------
def normalize(s: str) -> str:
    """
    Lenient normalization:
      - Lowercase, trim
      - Convert NBSP to space
      - Replace any whitespace or underscores with a single space
      - Remove straight & curly quotes
      - Remove asterisks (*)
      - Collapse multiple spaces
    """
    s = str(s)
    s = s.replace("\u00a0", " ")  # NBSP to normal space
    s = s.strip().lower()
    s = re.sub(r"[\s_]+", " ", s)  # spaces, tabs, newlines, underscores -> single space
    s = s.replace('"', '').replace("'", '').replace("“", "").replace("”", "").replace("*", "")
    s = re.sub(r"\s+", " ", s).strip()
    return s

def read_headers(xls: pd.ExcelFile, sheet_name: str, header_row: int):
    """Read only headers for a given sheet & header row. Returns (list[str] | None, error | None)."""
    try:
        df = pd.read_excel(xls, sheet_name=sheet_name, header=header_row, nrows=0, engine="openpyxl")
        return list(map(str, df.columns)), None
    except Exception as e:
        return None, e

def compare_header_lists(expected, actual, mode: str):
    """
    Compare expected vs actual header lists.
    Returns:
      (is_match, diffs, missing, unexpected, pos_mismatch_count)
    """
    if actual is None:
        return False, ["❌ Unable to read headers."], set(), set(), 10**9

    if mode.startswith("Exact"):
        eq = lambda a, b: a == b
        exp_for_sets, act_for_sets = expected, actual
    else:
        eq = lambda a, b: normalize(a) == normalize(b)
        exp_for_sets = [normalize(h) for h in expected]
        act_for_sets = [normalize(h) for h in actual]

    diffs = []
    pos_mismatch = 0
    max_len = max(len(expected), len(actual))
    for i in range(max_len):
        exp = expected[i] if i < len(expected) else "(none)"
        act = actual[i] if i < len(actual) else "(none)"
        ok = eq(exp, act)
        pos_mismatch += 0 if ok else 1
        diffs.append(f"{'✅' if ok else '❌'} Index {i}: expected **{exp}**, got **{act}**")

    missing = set(exp_for_sets) - set(act_for_sets)
    unexpected = set(act_for_sets) - set(exp_for_sets)
    is_match = (pos_mismatch == 0) and (len(missing) == 0) and (len(unexpected) == 0)
    return is_match, diffs, missing, unexpected, pos_mismatch

def score_mismatch(missing, unexpected, pos_mismatch_count):
    return len(missing) + len(unexpected) + pos_mismatch_count

def evaluate_sheet_against_schemas(xls: pd.ExcelFile, sheet_name: str, mode: str, try_rows: int):
    result = {
        "sheet_name": sheet_name,
        "matched": False,
        "matched_schema": None,
        "header_row": None,
        "actual_headers": None,
        "diffs": [],
        "missing_cols": set(),
        "unexpected_cols": set(),
        "read_error": None,
        "best_schema": None,
        "best_header_row": None,
        "best_score": None
    }

    best = {"schema": None, "row": None, "score": 10**9, "diffs": [], "missing": set(), "unexpected": set(), "headers": None}

    for hdr_row in range(0, try_rows + 1):
        actual, err = read_headers(xls, sheet_name, header_row=hdr_row)
        if err is not None:
            if result["read_error"] is None:
                result["read_error"] = str(err)
            continue
        if actual is None or len(actual) == 0:
            continue

        for schema_name, expected in SCHEMAS.items():
            is_match, diffs, missing, unexpected, pos_mis = compare_header_lists(expected, actual, mode)
            if is_match:
                result.update({
                    "matched": True, "matched_schema": schema_name, "header_row": hdr_row,
                    "actual_headers": actual, "diffs": diffs,
                    "missing_cols": missing, "unexpected_cols": unexpected, "read_error": None
                })
                return result
            else:
                score = score_mismatch(missing, unexpected, pos_mis)
                if score < best["score"]:
                    best.update({
                        "schema": schema_name, "row": hdr_row, "score": score,
                        "diffs": diffs, "missing": missing, "unexpected": unexpected, "headers": actual
                    })

    if best["schema"] is not None:
        result.update({
            "matched": False, "matched_schema": None,
            "actual_headers": best["headers"], "diffs": best["diffs"],
            "missing_cols": best["missing"], "unexpected_cols": best["unexpected"],
            "best_schema": best["schema"], "best_header_row": best["row"], "best_score": best["score"]
        })
    return result

# --------------------- Upload & Validate ---------------------
uploaded = st.file_uploader("Upload an Excel file (.xlsx/.xlsm)", type=["xlsx", "xlsm"])

if uploaded:
    st.subheader("Uploaded file")
    st.write(f"**Name:** `{uploaded.name}`")

    try:
        xls = pd.ExcelFile(uploaded, engine="openpyxl")  # stable path
        st.write("**Sheets found:** ", ", ".join(xls.sheet_names))
    except Exception as e:
        st.error("❌ Failed to open Excel with openpyxl. Confirm the file is .xlsx/.xlsm and not .xls.")
        st.exception(e)
        st.stop()

    sheet_results = []
    all_sheets_pass = True
    schemas_used = set()

    with st.spinner("Validating each sheet against the 4 schemas..."):
        for sheet in xls.sheet_names:
            res = evaluate_sheet_against_schemas(xls, sheet, match_mode, max_header_scan_rows)
            sheet_results.append(res)
            if res["matched"]:
                schemas_used.add(res["matched_schema"])
            else:
                all_sheets_pass = False

    if all_sheets_pass and single_schema_for_workbook and len(schemas_used) > 1:
        all_sheets_pass = False

    st.subheader("Result")
    if all_sheets_pass:
        if single_schema_for_workbook:
            only_schema = list(schemas_used)[0] if schemas_used else "Unknown"
            st.success(f"✅ All sheets match the same schema: **{only_schema}**")
        else:
            st.success("✅ Every sheet matches one of the allowed schemas.")
        st.markdown(FORMS_LINK)
    else:
        st.error("❌ Some sheets did not match any schema. See details below.")

    for res in sheet_results:
        if res["matched"]:
            with st.expander(f"Sheet: {res['sheet_name']} — ✅ MATCH ({res['matched_schema']}, header_row={res['header_row']})", expanded=False):
                st.write("**Headers found:**")
                st.code(res["actual_headers"])
                st.write("**Position-wise check:**")
                for d in res["diffs"]:
                    st.write(d)
        else:
            with st.expander(f"Sheet: {res['sheet_name']} — ❌ NO MATCH", expanded=True):
                if res["actual_headers"]:
                    st.write("**Headers found (best-try row):**")
                    st.code(res["actual_headers"])
                if res["read_error"]:
                    st.error(f"Read error (first encountered): {res['read_error']}")
                if res["best_schema"]:
                    st.warning(f"Closest schema: **{res['best_schema']}** (tried header_row={res['best_header_row']}, score={res['best_score']})")
                    st.write("**Position-wise differences:**")
                    for d in res["diffs"]:
                        st.write(d)
                    if res["missing_cols"]:
                        st.warning("Missing columns (expected but not present):")
                        for m in sorted(res["missing_cols"]):
                            st.write(f"- {m}")
                    if res["unexpected_cols"]:
                        st.info("Unexpected columns (present but not expected):")
                            st.write(f"- {u}")
                else:
                    st.write("Could not find a close schema match.")

    with st.expander("Diagnostics", expanded=False):
        st.write("**pandas version**:", pd.__version__)
        try:
            import openpyxl
            st.write("**openpyxl version**:", openpyxl.__version__)
        except Exception as e:
            st.write("openpyxl not importable:", e)
else:
    st.info("Upload an Excel file to begin.")
