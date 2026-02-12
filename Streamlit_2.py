import re
import pandas as pd
import streamlit as st

# -------------------------- Page --------------------------
st.set_page_config(page_title="Standard & ABC Costing Inputs", page_icon="✅", layout="centered")
st.title("Standard & ABC Costing Inputs")

st.caption(
    "Upload an Excel file (.xlsx/.xlsm). Each sheet must match one of the 4 allowed schemas. "
    "If all sheets match, you will receive the Forms link."
)

FORMS_LINK = "https://forms.office.com/Pages/ResponsePage.aspx?id=GIKK1zVBJkCjqBzdciO01bNCosI6R5VPnjnHYY4WuWxUNzNKNURZWlJJSVdKSDVITlE2WEwzUExKRy4u"

# -------------------------- SCHEMAS --------------------------
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

# -------------------------- OPTIONS --------------------------
# Force Lenient mode only (no UI control)
MATCH_MODE = "Lenient"

# -------------------------- Utility Functions --------------------------
def normalize(s: str) -> str:
    s = str(s).replace("\u00a0", " ")
    s = s.strip().lower()
    s = re.sub(r"[\s_]+", " ", s)
    s = s.replace('"', '').replace("'", '').replace("“", "").replace("”", "").replace("*", "")
    s = re.sub(r"\s+", " ", s).strip()
    return s

def read_headers(xls: pd.ExcelFile, sheet_name: str):
    try:
        df = pd.read_excel(xls, sheet_name=sheet_name, header=0, nrows=0, engine="openpyxl")
        return list(map(str, df.columns)), None
    except Exception as e:
        return None, e

def compare_headers(expected, actual, mode: str):
    if actual is None:
        return False, ["❌ Unable to read headers."], set(), set()

    # Only Lenient path will be used
    if mode.startswith("Exact"):
        exp_set = expected
        act_set = actual
        eq = lambda a, b: a == b
    else:
        exp_set = [normalize(h) for h in expected]
        act_set = [normalize(h) for h in actual]
        eq = lambda a, b: normalize(a) == normalize(b)

    diffs = []
    max_len = max(len(expected), len(actual))
    for i in range(max_len):
        exp = expected[i] if i < len(expected) else "(none)"
        act = actual[i] if i < len(actual) else "(none)"
        ok = eq(exp, act)
        diffs.append(f"{'✅' if ok else '❌'} Index {i}: expected **{exp}**, got **{act}**")

    missing = set(exp_set) - set(act_set)
    unexpected = set(act_set) - set(exp_set)

    is_match = len(missing) == 0 and len(unexpected) == 0 and all(
        diffs[i].startswith("✅") for i in range(min(len(expected), len(actual)))
    )

    return is_match, diffs, missing, unexpected

def evaluate_sheet(xls, sheet_name, mode):
    actual, err = read_headers(xls, sheet_name)

    if err:
        return {
            "matched": False,
            "schema": None,
            "actual_headers": None,
            "diffs": [f"❌ Could not read sheet: {err}"],
            "missing": set(),
            "unexpected": set()
        }

    best_schema = None
    best_score = 999999
    best_details = None

    for schema_name, expected in SCHEMAS.items():
        is_match, diffs, missing, unexpected = compare_headers(expected, actual, mode)
        score = len(missing) + len(unexpected)

        if is_match:
            return {
                "matched": True,
                "schema": schema_name,
                "actual_headers": actual,
                "diffs": diffs,
                "missing": missing,
                "unexpected": unexpected
            }

        if score < best_score:
            best_schema = schema_name
            best_score = score
            best_details = (diffs, missing, unexpected)

    diffs, missing, unexpected = best_details

    return {
        "matched": False,
        "schema": best_schema,
        "actual_headers": actual,
        "diffs": diffs,
        "missing": missing,
        "unexpected": unexpected
    }

# -------------------------- UPLOAD --------------------------
uploaded = st.file_uploader("Upload an Excel file (.xlsx/.xlsm)", type=["xlsx", "xlsm"])

if uploaded:
    st.subheader("Uploaded file")
    st.write(f"**Name:** `{uploaded.name}`")

    try:
        xls = pd.ExcelFile(uploaded, engine="openpyxl")
    except Exception as e:
        st.error("❌ Could not open the Excel file.")
        st.exception(e)
        st.stop()

    sheet_results = []

    for sheet in xls.sheet_names:
        res = evaluate_sheet(xls, sheet, MATCH_MODE)
        sheet_results.append((sheet, res))

    all_pass = all(res["matched"] for _, res in sheet_results)

    st.subheader("Result")
    if all_pass:
        st.success("✅ All sheets match the allowed schemas.")
        st.markdown(FORMS_LINK)
    else:
        st.error("❌ One or more sheets did not match any schema.")

    for sheet, res in sheet_results:
        with st.expander(f"Sheet: {sheet} — {'✅ MATCH' if res['matched'] else '❌ NO MATCH'}", expanded=not res["matched"]):
            if res["actual_headers"]:
                st.write("**Headers found:**")
                st.code(res["actual_headers"])
            st.write("**Schema matched/closest:**", res["schema"])
            for d in res["diffs"]:
                st.write(d)
            if res["missing"]:
                st.warning("Missing:")
                st.code(list(res["missing"]))
            if res["unexpected"]:
                st.info("Unexpected:")
                st.code(list(res["unexpected"]))
else:
    st.info("Upload an Excel file to begin.")
