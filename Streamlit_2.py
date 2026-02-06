import re
import pandas as pd
import streamlit as st
from datetime import datetime, timedelta

# Try zoneinfo; fall back to manual IST offset if not available
try:
    from zoneinfo import ZoneInfo
    def now_ist() -> datetime:
        return datetime.now(ZoneInfo("Asia/Kolkata"))
except Exception:
    # Robust fallback: use UTC and add +5:30 to avoid server-local time surprises
    def now_ist() -> datetime:
        return datetime.utcnow() + timedelta(hours=5, minutes=30)

st.set_page_config(page_title="Excel Header Checker", page_icon="✅", layout="centered")
st.title("✅ Excel Header Checker")

EXPECTED_HEADERS = ["Employ_id", "Name", "Department"]

st.write("**Expected headers (in order):**")
st.code(EXPECTED_HEADERS)

match_mode = st.radio(
    "Match mode",
    ["Exact (case & spaces must match)", "Lenient (ignore case/extra spaces/underscores)"],
    index=1
)

def normalize(h: str) -> str:
    """Lowercase, trim, and collapse underscores/spaces to one space for lenient match."""
    s = str(h).strip().lower()
    s = re.sub(r"[_\s]+", " ", s)
    return s

def get_submission_link() -> tuple[str, str]:
    """
    If today's IST date is after the 5th (i.e., 6..31), return Google.
    Else, return Microsoft Forms link.
    Returns: (url, label)
    """
    now = now_ist()
    if now.day > 5:
        return "https://www.google.com", "Google"
    else:
        return ("https://forms.office.com/Pages/ResponsePage.aspx"
                "?id=GIKK1zVBJkCjqBzdciO01bNCosI6R5VPnjnHYY4WuWxUNzNKNURZWlJJSVdKSDVITlE2WEwzUExKRy4u", 
                "Microsoft Forms")

uploaded = st.file_uploader("Upload an Excel file (.xlsx/.xls)", type=["xlsx", "xls"])

if uploaded:
    try:
        # You can optionally detect extension and switch engine if you truly need .xls
        actual_headers = list((pd.read_excel(uploaded, engine="openpyxl")).columns.astype(str))

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
            link_url, link_label = get_submission_link()
            st.markdown(link_url)
            # Debug/info caption so you can verify the rule that was applied
            now = now_ist()
            st.caption(
                f"Rule applied for {now.strftime('%d-%b-%Y %H:%M')} IST "
                f"→ Showing **{link_label}** (day={now.day}, condition: day > 5)."
            )
        else:
            st.error("❌ Correct your header")
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

            if match_mode.startswith("Exact"):
                exp_set = set(EXPECTED_HEADERS)
                act_set = set(actual_headers)
            else:
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
``
