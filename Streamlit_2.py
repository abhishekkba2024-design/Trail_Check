import re
import pandas as pd
import streamlit as st

# -------------------------- UI & Page --------------------------
st.set_page_config(page_title="Excel Header Checker (Multi-sheet, 4 schemas)", page_icon="✅", layout="centered")
st.title("✅ Excel Header Checker (Multi-sheet, 4 schemas)")

st.caption(
    "Upload an Excel file (.xlsx/.xls). The app checks if **every sheet** matches **one** of the 4 allowed schemas. "
    "If all sheets match, you'll get the Forms link. Otherwise, you'll see precise mismatches."
)

FORMS_LINK = "https://forms.office.com/Pages/ResponsePage.aspx?id=GIKK1zVBJkCjqBzdciO01bNCosI6R5VPnjnHYY4WuWxUNzNKNURZWlJJSVdKSDVITlE2WEwzUExKRy4u"

# -------------------------- YOUR 4 SCHEMAS --------------------------
# These are taken directly from your message.
# Newlines inside headers are kept where indicated (e.g., "Shop\nCode").
# Lenient mode will ignore case, extra spaces/underscores, and newlines are normalized to a single space.

SCHEMAS = {
    "Schema 1": [
        "S. No.",
        "Plant",
        "Process",
        "Shop Type",
        "Shop\nCode",
        "Shop Description",
        "Cost Centre",
        "Line",
        "Line type",
        "Iot Status",
        "Capacity per shift ( As per prod)",
        "Logical PSL",
        "Physical PSL **",
        "Operational Shifts",
        "Details of Deviation in Shift *",
        "No of Machines",
        "Category\n(4W/2W)",
        "Part No/\nModel No",
        "Part Name/\nModel Description",
        "Gross RM Weight\n(in kg/ Part)",
        "Broad Model",
        "Engine/TM",
        "Financial Year",
        "Month",
        "OK Volume",
        "Rejection Volume",
        "OEE/Line Efficiency As per Production",
        "Remarks ( If Any)",
    ],
    "Schema 2": [
        "S. No.",
        "Plant",
        "Maint Dept Code",
        "Maint Dept Description",
        "Cost Centre",
        "Type",
        "Shop Coverage*",
        "Remarks",
    ],
    "Schema 3": [
        "S. No.",
        "Plant",
        "Process",
        "Shop Type",
        "Shop\nCode/Maint Dept",
        "Shop Description",
        "Cost Centre",
        "Month",
        "Activity Category",
        "Sub Activity*",
        "Manpower Grade",
        "Headcount",
    ],
    "Schema 4": [
        "S. No.",
        "Plant",
        "Process",
        "Shop",
        "Line",
        "Line Type",
        "Model",
        "Fuel Type",
        "Part No/\nModel No",
        "Part Name/\nModel Description",
        "Annual Capacity as per OEE*",
        "SMM\nPer Part**",
        "Takt Time",
        "Cycle Time",
        "OEE Factor (%)",
    ],
}

# -------------------------- MATCHING OPTIONS --------------------------
match_mode = st.radio(
    "Match mode",
    ["Exact (case & spaces must match)", "Lenient (ignores case/extra spaces/underscores/newlines/quotes/*)"],
    index=1
)
single_schema_for_workbook = st.checkbox(
    "Require that all sheets match the **same** schema", value=False,
    help="If checked, every sheet must conform to the same schema (Schema 1, or Schema 2, etc.)."
)

max_header_scan_rows = st.slider(
    "Try header row at lines (0 = first row)",
    min_value=0, max_value=5, value=3,
    help="The app will try these many rows as header rows and pick the best."
)

# -------------------------- UTILITIES --------------------------
def normalize(s: str) -> str:
    """
    Lenient normalization:
      - Lowercase
      - Strip
      - Replace any whitespace (space, tabs, newlines) & underscores with a single space
      - Remove straight & curly quotes
      - Remove asterisks (*)
      - Collapse multiple spaces
    """
    s = str(s)
    s = s.strip().lower()
    # Replace any whitespace or underscores with single space
    s = re.sub(r"[\s_]+", " ", s)
    # Remove quotes and asterisks
    s = s.replace('"', '').replace("'", '').replace("“", "").replace("”", "").replace("*", "")
    # Collapse spaces again after removals
    s = re.sub(r"\s+", " ", s).strip()
    return s

def get_engine_from_name(filename: str) -> str:
    # .xlsx -> openpyxl; .xls -> xlrd
    if filename.lower().endswith(".xls"):
        return "xlrd"
    return "openpyxl"

def read_headers(xls: pd.ExcelFile, sheet_name: str, header_row: int):
    """Read only headers for a given sheet & header row. Returns (list[str] | None, error | None)."""
    try:
        df = pd.read_excel(xls, sheet_name=sheet_name, header=header_row, nrows=0)
        return list(map(str, df.columns)), None
    except Exception as e:
        return None, e

def compare_header_lists(expected, actual, mode: str):
    """
    Compare expected vs actual header lists.
    Returns:
      - is_match: bool
      - diffs: list[str] position-wise differences
      - missing: set[str]
      - unexpected: set[str]
      - pos_mismatch_count: int
    """
    if actual is None:
        return False, ["❌ Unable to read headers."], set(), set(), 10**9

    if mode.startswith("Exact"):
        eq = lambda a, b: a == b
        show = lambda x: x
        exp_for_sets = expected
        act_for_sets = actual
    else:
        eq = lambda a, b: normalize(a) == normalize(b)
        show = lambda x: normalize(x)
        exp_for_sets = [normalize(h) for h in expected]
        act_for_sets = [normalize(h) for h in actual]

    diffs = []
    pos_mismatch_count = 0
    max_len = max(len(expected), len(actual))
    for i in range(max_len):
        exp = expected[i] if i < len(expected) else "(none)"
        act = actual[i] if i < len(actual) else "(none)"
        ok = eq(exp, act)
        status = "✅" if ok else "❌"
        if not ok:
            pos_mismatch_count += 1
        diffs.append(f"{status} Index {i}: expected **{exp}**, got **{act}**")

    missing = set(exp_for_sets) - set(act_for_sets)
    unexpected = set(act_for_sets) - set(exp_for_sets)

    is_match = (pos_mismatch_count == 0) and (len(missing) == 0) and (len(unexpected) == 0)
    return is_match, diffs, missing, unexpected, pos_mismatch_count

def score_mismatch(missing, unexpected, pos_mismatch_count):
    """Lower is better."""
    return len(missing) + len(unexpected) + pos_mismatch_count

def evaluate_sheet_against_schemas(xls: pd.ExcelFile, sheet_name: str, mode: str, try_rows: int):
    """
    Try multiple header rows (0..try_rows) and all schemas.
    Returns:
      {
        "sheet_name": str,
        "matched": bool,
        "matched_schema": Optional[str],
        "header_row": Optional[int],
        "actual_headers": Optional[list[str]],
        "diffs": list[str],
        "missing_cols": set[str],
        "unexpected_cols": set[str],
        "read_error": Optional[str],
        "best_schema": Optional[str],        # If not matched, which schema is closest
        "best_header_row": Optional[int],    # Header row where it was closest
        "best_score": Optional[int]
      }
    """
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
            # Track read error only if all rows fail
            if result["read_error"] is None:
                result["read_error"] = str(err)
            continue

        # If the sheet is blank or no columns
        if actual is None or len(actual) == 0:
            continue

        # Try each schema
        for schema_name, expected in SCHEMAS.items():
            is_match, diffs, missing, unexpected, pos_mis = compare_header_lists(expected, actual, mode)
            if is_match:
                # Found a perfect match—return early
                result.update({
                    "matched": True,
                    "matched_schema": schema_name,
                    "header_row": hdr_row,
                    "actual_headers": actual,
                    "diffs": diffs,
                    "missing_cols": missing,
                    "unexpected_cols": unexpected,
                    "read_error": None
                })
                return result
            else:
                score = score_mismatch(missing, unexpected, pos_mis)
                if score < best["score"]:
                    best.update({
                        "schema": schema_name,
                        "row": hdr_row,
                        "score": score,
                        "diffs": diffs,
                        "missing": missing,
                        "unexpected": unexpected,
                        "headers": actual
                    })

    # No exact match; return best attempt (for guidance)
    if best["schema"] is not None:
        result.update({
            "matched": False,
            "matched_schema": None,
            "header_row": None,
            "actual_headers": best["headers"],
            "diffs": best["diffs"],
            "missing_cols": best["missing"],
            "unexpected_cols": best["unexpected"],
            "best_schema": best["schema"],
            "best_header_row": best["row"],
            "best_score": best["score"]
        })
    return result

# -------------------------- FILE UPLOAD --------------------------
uploaded = st.file_uploader("Upload an Excel file (.xlsx/.xls)", type=["xlsx", "xls"])

if uploaded:
    st.subheader("Uploaded file")
    st.write(f"**Name:** `{uploaded.name}`")

    engine = get_engine_from_name(uploaded.name)
    try:
        xls = pd.ExcelFile(uploaded, engine=engine)
    except Exception as e:
        st.error(f"❌ Failed to open Excel: {e}")
        st.stop()

    st.write("**Sheets found:**", ", ".join(xls.sheet_names))

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

    # If we require a single schema across the workbook, enforce it
    if all_sheets_pass and single_schema_for_workbook and len(schemas_used) > 1:
        all_sheets_pass = False

    # -------------------------- SUMMARY --------------------------
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

    # -------------------------- DETAILS PER SHEET --------------------------
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
                        for u in sorted(res["unexpected_cols"]):
                            st.write(f"- {u}")
                else:
                    st.write("Could not find a close schema match.")
else:
    st.info("Upload an Excel file to begin.")
