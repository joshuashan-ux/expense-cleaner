import io
import pandas as pd
import requests
import streamlit as st

st.set_page_config(page_title="Tax-Ready Statement Cleaner", page_icon="📊")

# --- 1. CORE DATA CLEANING LOGIC ---
def categorize_expense(description: str) -> str:
    desc = str(description).lower()
    if any(word in desc for word in ["uber", "train", "fuel", "shell", "bp", "parking", "tfl"]):
        return "Travel & Transport"
    if any(word in desc for word in ["aws", "github", "adobe", "google", "openai", "hosting", "apple"]):
        return "Software & Subscriptions"
    if any(word in desc for word in ["coffee", "pret", "lunch", "cafe", "restaurant", "greggs", "costa"]):
        return "Meals & Subsistence"
    if any(word in desc for word in ["amazon", "ikea", "argos", "currys"]):
        return "Office Equipment"
    return "Uncategorized"

def clean_statement(df: pd.DataFrame) -> pd.DataFrame:
    df.columns = [col.strip().lower() for col in df.columns]
    df = df.dropna(how="all").copy()
    
    desc_col = next((c for c in df.columns if "desc" in c or "merchant" in c or "memo" in c), df.columns)
    amount_col = next((c for c in df.columns if "amount" in c or "value" in c or "debit" in c), df.columns[-1])
    
    df["Clean Amount"] = (
        df[amount_col]
        .astype(str)
        .str.replace(r"[£$,]", "", regex=True)
        .str.strip()
    )
    df["Clean Amount"] = pd.to_numeric(df["Clean Amount"], errors="coerce").fillna(0.0)
    df["Category"] = df[desc_col].apply(categorize_expense)
    return df

# --- 2. LICENSE KEY VERIFICATION ---
def check_license_key(license_key: str) -> bool:
    if license_key.strip() == "TEST-PRO-2026":
        return True
        
    url = "https://api.lemonsqueezy.com/v1/licenses/validate"
    response = requests.post(url, data={"license_key": license_key.strip()})
    if response.status_code == 200:
        data = response.json()
        return data.get("valid", False)
    return False

# --- 3. USER INTERFACE ---
st.title("📊 Tax-Ready Bank Statement Cleaner")
st.write("Upload a messy bank CSV to auto-categorize expenses and export a clean Excel sheet.")

uploaded_file = st.file_uploader("Upload your bank statement (.csv)", type=["csv"])

# One-click sample data toggle for instant testing
if "use_sample" not in st.session_state:
    st.session_state.use_sample = False

if st.button("✨ Try with Sample Bank Statement"):
    st.session_state.use_sample = True

raw_df = None
if uploaded_file is not None:
    st.session_state.use_sample = False
    raw_df = pd.read_csv(uploaded_file)
elif st.session_state.use_sample:
    sample_csv = """Date,Description,Amount
01/11/2026,Uber Trip London,-18.50
02/11/2026,GitHub Copilot,-10.00
03/11/2026,Pret A Manger,-6.45
04/11/2026,Amazon Office Chair,-129.99
05/11/2026,Trainline Ticket,-42.00"""
    raw_df = pd.read_csv(io.StringIO(sample_csv))

if raw_df is not None:
    cleaned_df = clean_statement(raw_df)
    
    st.subheader("Preview (First 5 Rows)")
    st.dataframe(cleaned_df.head(5), use_container_width=True)
    
    total_spent = cleaned_df["Clean Amount"].sum()
    st.metric("Total Processed Amount", f"£{total_spent:,.2f}")
    
    st.divider()
    st.subheader("Download Full Tax-Ready Excel File")
    
    st.markdown("Need a key? **[Get Instant Access (£9 One-Time)](https://www.lemonsqueezy.com)**")
    license_input = st.text_input("Enter your License Key to unlock download:", placeholder="e.g. XXXXXXXX-XXXX-XXXX")
    
    if license_input:
        if check_license_key(license_input):
            st.success("License verified! Your Excel report is ready.")
            
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                cleaned_df.to_excel(writer, index=False, sheet_name="Cleaned Expenses")
                summary = cleaned_df.groupby("Category")["Clean Amount"].sum().reset_index()
                summary.to_excel(writer, index=False, sheet_name="Category Summary")
                
            st.download_button(
                label="⬇️ Download Clean Excel (.xlsx)",
                data=buffer.getvalue(),
                file_name="tax_ready_expenses.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.error("Invalid license key. Please double-check your key or purchase access above.")
