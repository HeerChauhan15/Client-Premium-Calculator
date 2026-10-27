import streamlit as st
import pandas as pd
import numpy as np
import os
import re
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter

st.set_page_config(
    page_title="MURUDE SAHAKARI PATSANSTHA MARYADIT	",
    page_icon="🏦",
    layout="wide"
)

st.title("🏦 MURUDE SAHAKARI PATSANSTHA MARYADIT")
st.subheader("Client Premium Calculator")
st.markdown("Select plan details below")

# ============================================
# FIXED SETTINGS — these are no longer user-selectable.
# Life Type is always "Single Life" and Cover Type is always "Reducing".
# ============================================
LIFE_TYPE_FIXED = "Single Life"
COVER_TYPE_FIXED = "Reducing"

# Only two rate files now — one per Loan Type. Life Type and Cover Type are
# fixed above, so the file only needs to vary by loan_type.
# NOTE: File names must match exactly (including the space after the hyphen).
FILE_MAP = {
    "Home Loan": "Homeloan- Reducing.xlsx",
    "Business Loan": "Business Loan- Reducing.xlsx",
}

# GST is fixed and always applied on top of the rate.
GST_RATE_FIXED = 18.0

# Loader has been removed — kept at 0 so the formula below still works
# unchanged (Rate After Loader = Base Rate / (1 - 0/100) = Base Rate).
LOADER_PCT_FIXED = 0.0

# ============================================
# GST FORMULA (applied to every rate before premium is computed):
#   Rate After Loader = Base Rate / (1 - Loader% / 100)   [Loader is fixed at 0]
#   Final Rate         = Rate After Loader x (1 + GST% / 100)
# ============================================
def apply_loader_and_gst(base_rate, loader_pct, gst_pct=GST_RATE_FIXED):
    after_loader = base_rate / (1 - (loader_pct / 100.0))
    final_rate = after_loader * (1 + (gst_pct / 100.0))
    return final_rate


def load_rate_table(cover_type, life_type, loan_type):
    fname = FILE_MAP[loan_type]
    if not os.path.exists(fname):
        raise FileNotFoundError(
            f"File not found: '{fname}' — Please make sure this file is in the GitHub repo."
        )

    raw = pd.read_excel(fname, sheet_name="Sheet1", header=None)

    header_row = None
    for i, row in raw.iterrows():
        for val in row.values:
            if isinstance(val, str) and "AGE" in val.upper():
                header_row = i
                break
        if header_row is not None:
            break

    if header_row is None:
        raise ValueError("Could not find AGE/TERM header row in the file.")

    df = pd.read_excel(fname, sheet_name="Sheet1", header=header_row)
    df.columns = [str(c).strip() for c in df.columns]

    age_col = df.columns[0]
    df = df.dropna(subset=[age_col])
    df[age_col] = pd.to_numeric(df[age_col], errors='coerce')
    df = df.dropna(subset=[age_col])
    df[age_col] = df[age_col].astype(int)
    df = df.set_index(age_col)

    tenure_map = {}
    for col in df.columns:
        try:
            tenure_map[int(float(col))] = col
        except Exception:
            pass

    return df, tenure_map


def get_rate(df, tenure_map, age, tenure):
    if age not in df.index:
        raise ValueError(f"Age {age} not found in rate table.")
    if tenure not in tenure_map:
        raise ValueError(f"Tenure {tenure} yrs not found in rate table.")
    return float(df.loc[age, tenure_map[tenure]])


def find_column(df, target):
    """Exact (case/space-insensitive) match."""
    target_norm = target.strip().lower().replace(" ", "")
    for col in df.columns:
        col_norm = str(col).strip().lower().replace(" ", "")
        if col_norm == target_norm:
            return col
    return None


def find_sum_assured_columns(df):
    """
    Detects the Sum Assured/Sum Insured column AND the Loan Outstanding column
    (if both are present). Per-row logic then uses Sum Assured when that row
    has a value, and falls back to Loan Outstanding for that row only when
    Sum Assured is blank/missing.
    """
    sa_col = None
    lo_col = None
    for col in df.columns:
        norm = re.sub(r'[\s_\-/]+', '', str(col).lower())
        if sa_col is None and ('sumassured' in norm or 'suminsured' in norm):
            sa_col = col
        if lo_col is None and ('loanoutstanding' in norm or 'outstandingamount' in norm or 'outstandingloan' in norm
                or norm == 'outstanding' or 'loanos' in norm or norm == 'os'):
            lo_col = col
    return sa_col, lo_col


# ============================================
# DROPDOWNS
# Life Type and Type of Cover are fixed (Single Life / Reducing) and no
# longer shown as selectable inputs.
# ============================================
life_type = LIFE_TYPE_FIXED
cover_type = COVER_TYPE_FIXED

loan_type = st.selectbox("Select Loan Type", ["Home Loan", "Business Loan"])

# Loader input removed — always 0. GST @ 18% is still applied automatically.
loader_pct = LOADER_PCT_FIXED

# ============================================
# LIMITS — defined ONCE here and used by BOTH the Individual and Group
# Premium Calculators (rates in the backend files are per ₹1,00,000)
# Home Loan:     Sum Assured ₹1,00,000 – ₹60,00,000 | Tenure 5 – 25 yrs
# Business Loan: Sum Assured ₹50,000   – ₹50,00,000 | Tenure 2 – 10 yrs
# ============================================
AGE_MIN, AGE_MAX = 18, 65

if loan_type == "Home Loan":
    sa_min, sa_max = 50000, 6000000
    min_tenure, max_tenure = 5, 25
else:  # Business Loan
    sa_min, sa_max = 50000, 5000000
    min_tenure, max_tenure = 2, 10

st.divider()

# ============================================
# MANUAL SECTION
# ============================================

st.subheader("🔢 Individual Premium Calculator")

# (keys include loan_type so limits reset cleanly when the loan type changes)
col3, col4 = st.columns(2)
with col3:
    age = st.number_input(
        "Enter Age", min_value=AGE_MIN, max_value=AGE_MAX, value=30, step=1,
        key=f"age_manual_{loan_type}"
    )
with col4:
    tenure = st.number_input(
        "Enter Tenure (Years)",
        min_value=min_tenure,
        max_value=max_tenure,
        value=min_tenure,
        step=1,
        help=f"Tenure in years ({min_tenure} – {max_tenure}).",
        key=f"tenure_manual_{loan_type}"
    )

sum_assured_manual = st.number_input(
    "Select Sum Assured (₹)",
    min_value=sa_min,
    max_value=sa_max,
    value=sa_min,
    step=1,
    help=f"Enter the exact Sum Assured for this member (₹{sa_min:,} – ₹{sa_max:,}).",
    key=f"sa_manual_{loan_type}"
)

if st.button("Get Rate", type="primary"):
    try:
        df_rates, tenure_map = load_rate_table(cover_type, life_type, loan_type)
        base_rate = get_rate(df_rates, tenure_map, age, tenure)
        # Net premium = premium before GST; GST = 18% of net; Gross = Net + GST
        net_rate = base_rate / (1 - (loader_pct / 100.0))
        net_premium = net_rate * (sum_assured_manual / 100000)
        gst_amount = net_premium * (GST_RATE_FIXED / 100.0)
        gross_premium = net_premium + gst_amount
        st.success(
            f"✅ {life_type} | {loan_type} | {cover_type} Cover | Age {age} | Tenure {tenure} yrs | "
            f"Sum Assured ₹{sum_assured_manual:,} | GST {GST_RATE_FIXED}%"
        )
        m1, m2, m3 = st.columns(3)
        m1.metric("Net Premium", f"₹ {net_premium:,.2f}")
        m2.metric(f"GST Amount ({GST_RATE_FIXED:.0f}%)", f"₹ {gst_amount:,.2f}")
        m3.metric("Gross Premium", f"₹ {gross_premium:,.2f}")
    except Exception as e:
        st.error(f"Error: {e}")

st.divider()

# ============================================
# EXCEL UPLOAD SECTION
# ============================================

head_col, info_col = st.columns([4, 1])
with head_col:
    st.subheader("📂 Group Premium Calculator")
with info_col:
    # Popover = the "popup" for Excel requirements; keeps the page clean.
    with st.popover("ℹ️ Excel format", use_container_width=True):
        st.markdown("**Required columns**")
        st.markdown(
            "- **Name**\n"
            "- **Age**\n"
            "- **Sum Assured** (or **Loan Outstanding** if blank)\n"
            "- **Tenure** (in years)"
        )
        st.markdown("**Limits for current selection**")
        st.markdown(
            f"- Loan Type: **{loan_type}**\n"
            f"- Sum Assured: ₹{sa_min:,} – ₹{sa_max:,}\n"
            f"- Tenure: {min_tenure} – {max_tenure} yrs\n"
            f"- Age: {AGE_MIN} – {AGE_MAX}"
        )
        st.caption("Each row uses its own Sum Assured from the Excel.")

uploaded_file = st.file_uploader(
    f"Upload Excel File for {loan_type}",
    type=["xlsx"],
    help="Rates are picked based on the Loan Type selected above."
)

if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file)
        df.columns = [str(c).strip() for c in df.columns]

        st.subheader("Uploaded Data Preview")
        st.dataframe(df.head())

        # Same limits as the Individual Premium Calculator
        min_t, max_t = min_tenure, max_tenure

        df_rates, tenure_map = load_rate_table(cover_type, life_type, loan_type)

        name_col = find_column(df, "Name")
        age_col = find_column(df, "Age")
        tenure_col = find_column(df, "Tenure")
        sa_col, lo_col = find_sum_assured_columns(df)

        if not name_col or not age_col or not tenure_col:
            raise ValueError(
                "Excel must contain mandatory columns: Name, Age, and Tenure."
            )
        if not sa_col and not lo_col:
            raise ValueError("Excel must contain a Sum Assured column (e.g. 'Sum Assured', 'Sum Insured') or, failing that, a 'Loan Outstanding' column.")

        df[age_col] = pd.to_numeric(df[age_col], errors='coerce')
        df[tenure_col] = pd.to_numeric(df[tenure_col], errors='coerce')
        if sa_col:
            df[sa_col] = pd.to_numeric(df[sa_col], errors='coerce')
        if lo_col:
            df[lo_col] = pd.to_numeric(df[lo_col], errors='coerce')

        if df[tenure_col].dropna().median() > 30:
            st.info("ℹ️ Tenure values look like months — auto-converting to years.")
            df[tenure_col] = (df[tenure_col] / 12).round(0).astype('Int64')
        else:
            df[tenure_col] = df[tenure_col].round(0).astype('Int64')

        df[age_col] = df[age_col].round(0).astype('Int64')

        premiums = []
        statuses = []
        sa_used_list = []
        for idx, row in df.iterrows():
            try:
                r_age = int(row[age_col]) if pd.notna(row[age_col]) else None
                r_tenure = int(row[tenure_col]) if pd.notna(row[tenure_col]) else None

                # Per-row: use Sum Assured if this row has it; otherwise fall back
                # to Loan Outstanding for this row.
                r_sa = None
                if sa_col and pd.notna(row[sa_col]):
                    r_sa = float(row[sa_col])
                elif lo_col and pd.notna(row[lo_col]):
                    r_sa = float(row[lo_col])

                if r_sa is None:
                    raise ValueError("Sum Assured / Loan Outstanding value missing")

                if r_age is None or r_age < AGE_MIN or r_age > AGE_MAX:
                    raise ValueError(f"Age must be between {AGE_MIN} and {AGE_MAX}")
                if r_tenure is None or r_tenure < min_t or r_tenure > max_t:
                    raise ValueError(f"Tenure must be between {min_t} and {max_t} yrs")
                if r_sa < sa_min or r_sa > sa_max:
                    raise ValueError(f"Sum Assured must be between ₹{sa_min:,} and ₹{sa_max:,}")

                r_base = get_rate(df_rates, tenure_map, r_age, r_tenure)
                r_final = apply_loader_and_gst(r_base, loader_pct)
                premium = round(r_final * (r_sa / 100000), 2)
                premiums.append(premium)
                statuses.append("✅")
                sa_used_list.append(r_sa)
            except Exception as e:
                premiums.append(None)
                statuses.append(f"❌ {e}")
                sa_used_list.append(None)

        df["Sum Assured Used"] = sa_used_list
        df["Premium"] = premiums
        df["Status"] = statuses

        display_sa_col = sa_col if sa_col else lo_col
        core_cols = [name_col, age_col, tenure_col, display_sa_col, "Sum Assured Used", "Premium"]
        extra_cols = [c for c in df.columns if c not in core_cols]
        df = df[core_cols + extra_cols]

        total_premium = pd.to_numeric(pd.Series(premiums), errors='coerce').sum()
        st.metric("💰 Grand Total Premium", f"₹ {total_premium:,.2f}")

        st.subheader("Rate Lookup Output")
        st.dataframe(df, use_container_width=True)

        total_row = {c: "" for c in df.columns}
        total_row[name_col] = "TOTAL PREMIUM"
        total_row["Premium"] = round(total_premium, 2)
        df_out = pd.concat([df, pd.DataFrame([total_row])], ignore_index=True)

        output_file = "Rate_Output.xlsx"
        df_out.to_excel(output_file, index=False)

        with open(output_file, "rb") as file:
            st.download_button(
                label="⬇ Download Output Excel",
                data=file,
                file_name=output_file,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Error: {e}")
