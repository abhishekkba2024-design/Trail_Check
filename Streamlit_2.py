import re
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Excel Header Checker", page_icon="✅", layout="centered")
st.title("✅ Excel Header Checker")

# ---- Configure your expected headers here ----
EXPECTED_HEADERS = ["Employ_id", "Name", "Department"]

st.write("**Expected headers (in order):**")
st.code(EXPECTED_HEADERS)

# Option: choose match mode
match_mode = st.radio(
    "Match mode",
    ["Exact (case & spaces must match)", "Lenient (ignore case/extra spaces/underscores)"],
    index=1
)

def normalize(h: str) -> str:
    """Lowercase, trim, and collapse underscores/spaces to one space for lenient match."""
    s = str(h).strip().lower()
    s = re.sub(r"[_\\s]+", " ", s)
    return s

uploaded = st.file_uploader("Upload an Excel file (.xlsx/.xls)", type=["xlsx", "xls"])

if uploaded:
    try:
        # Read first sheet by default (you can expose a sheet selector if needed)
        df = pd.read_excel(uploaded, engine="openpyxl")
        actual_headers = list(df.columns.astype(str))

        st.subheader("Headers found in file")
        st.code(actual_headers)

        if match_mode.startswith("Exact"):
            is_match = (actual_headers == EXPECTED_HEADERS)
        else:
            is_match = (
                [normalize(h) for h in actual_headers] ==
                [normalize(h) for h in EXPECTED_HEADERS]
            )

        if is_match:
            st.success("✅ Headers match!")
            # Show a clickable link to GOOGLE
            st.markdown("https://forms.office.com/Pages/ResponsePage.aspx?id=GIKK1zVBJkCjqBzdciO01bNCosI6R5VPnjnHYY4WuWxUNzNKNURZWlJJSVdKSDVITlE2WEwzUExKRy4u")
        else:
            st.error("❌ Correct your header")
            # Optional: show a quick diff to help fix
            st.markdown("**Differences (position-wise):**")
            max_len = max(len(EXPECTED_HEADERS), len(actual_headers))
            for i in range(max_len):
                exp = EXPECTED_HEADERS[i] if i < len(EXPECTED_HEADERS) else "(none)"
                act = actual_headers[i] if i < len(actual_headers) else "(none)"
                status = "✅" if (
                    (exp == act) if match_mode.startswith("Exact")
                    else (normalize(exp) == normalize(act))
                ) else "❌"
                st.write(f"{status} Index {i}: expected **{exp}**, got **{act}**")

            # Also show missing / unexpected (set-wise, ignoring order)
            if match_mode.startswith("Exact"):
                exp_set = set(EXPECTED_HEADERS)
                act_set = set(actual_headers)
            else:
                # map normalized → original to report clearly
                norm_exp_map = {normalize(h): h for h in EXPECTED_HEADERS}
                norm_act_map = {normalize(h): h for h in actual_headers}
                exp_set = set(norm_exp_map.keys())
                act_set = set(norm_act_map.keys())

            missing = exp_set - act_set
            unexpected = act_set - exp_set

            if missing:
                st.warning("Missing columns:")
                for m in missing:
                    st.write(f"- {m}")
            if unexpected:
                st.info("Unexpected columns:")
                for u in unexpected:
                    st.write(f"- {u}")

    except Exception as e:
        st.error(f"Failed to read Excel: {e}")
else:
    st.info("Upload an Excel file to begin.")

