import streamlit as st
import pandas as pd
import plotly.express as px
import io
from datetime import datetime
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import time
from streamlit_autorefresh import st_autorefresh
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

st.set_page_config(page_title="Glass Rejection Dashboard", layout="wide")

# === Hide Streamlit UI ===
st.markdown("""
    <style>
    #MainMenu, footer, header {visibility: hidden;}
    a[href^="https://github.com"],
    button[title="View app source"],
    button[title="Open app menu"],
    svg[data-testid="icon-pencil"],
    [data-testid="stActionButtonIcon"] svg[data-testid="icon-pencil"] {
        display: none !important;
    }
    </style>
""", unsafe_allow_html=True)

# === Auto Refresh ===
st_autorefresh(interval=300000, key="auto_refresh")  # every 5 minutes

# === Logo ===
st.markdown("<div style='text-align: center;'>", unsafe_allow_html=True)
st.image("KV-Logo-1.png", width=150)
st.markdown("</div>", unsafe_allow_html=True)

# === Load Google Sheet ===
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds_dict = st.secrets["google_service_account"]
creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
client = gspread.authorize(creds)

SHEET_ID = "1nYqbCDifAqllvVvNksw0xMD2BIlo3nwCKeCLN-hAgL0"
sheet = client.open_by_key(SHEET_ID).worksheet("AllData")
data = sheet.get_all_records()
df = pd.DataFrame(data)

st.success(f"✅ Loaded {len(df)} rows from Google Sheets at {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

# === Preprocess ===
df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
df["Year"] = df["Date"].dt.year
df["Quarter"] = df["Date"].dt.to_period("Q").astype(str)
df["Week#"] = df["Date"].dt.isocalendar().week
df["Reason"] = df["Reason"].astype(str).str.strip().str.lower()
df["Type"] = df["Type"].astype(str)

tab1, tab2, tab3 = st.tabs(["📊 Dashboard", "📄 Data Table", "📝 New Entry Form"])

# === DASHBOARD ===
with tab1:
    st.title("📊 Glass Rejection Intelligence Dashboard")

    st.markdown("### 📅 Weekly Rejections")
    selected_year = st.radio("Choose Year", sorted(df["Year"].dropna().unique()), horizontal=True)
    df_week = df[df["Year"] == selected_year]
    weekly = df_week.groupby("Week#")["Qty"].sum().reset_index()
    fig1 = px.line(weekly, x="Week#", y="Qty", markers=True)
    fig1.update_layout(
        xaxis=dict(tickmode="linear", dtick=3),
        shapes=[dict(type="line", x0=w, x1=w, yref="paper", y0=0, y1=1,
                     line=dict(color="cyan", width=2, dash="dot")) for w in [13, 26, 39, 52]]
    )
    st.plotly_chart(fig1, use_container_width=True)

    st.markdown("### 🔍 Rejections by Reason")
    reason_year = st.radio("Year", sorted(df["Year"].unique()), horizontal=True, key="reason_year")
    df_reason = df[df["Year"] == reason_year]
    fig2 = px.bar(df_reason.groupby("Reason")["Qty"].sum().reset_index(), x="Reason", y="Qty", color="Reason")
    st.plotly_chart(fig2, use_container_width=True)

    st.markdown("### 🧊 Rejections by Glass Type")
    type_year = st.radio("Year", sorted(df["Year"].unique()), horizontal=True, key="glass_type")
    top_types = df[df["Year"] == type_year]["Type"].value_counts().nlargest(5).index.tolist()
    df_type = df[(df["Year"] == type_year) & (df["Type"].isin(top_types))]
    fig3 = px.bar(df_type.groupby("Type")["Qty"].sum().reset_index(), x="Type", y="Qty", color="Type")
    st.plotly_chart(fig3, use_container_width=True)

    st.markdown("### 🏭 Rejections by Department")
    valid_quarters = [f"{y}Q{i}" for y in [2024, 2025] for i in range(1, 5) if not (y == 2025 and i > 2)]
    selected_q = st.radio("Select Quarter", valid_quarters, horizontal=True)
    df_q = df[df["Quarter"] == selected_q]
    if not df_q.empty:
        fig4 = px.pie(df_q.groupby("Dept.")["Qty"].sum().reset_index(), names="Dept.", values="Qty", hole=0.4)
        st.plotly_chart(fig4, use_container_width=True)
    else:
        st.warning("No data found for the selected quarter.")

# === DATA TABLE ===
with tab2:
    tab_data1, tab_data2 = st.tabs(["🟦 Scratched Glass Records", "🟥 Production Issue Records"])

    with tab_data1:
        st.markdown("### 🟦 Scratched Glass Records")
        selected_year = st.radio("Select Year", sorted(df["Year"].unique(), reverse=True), horizontal=True, key="scratch_year")
        df_scratch = df[(df["Reason"] == "scratched") & (df["Year"] == selected_year)]
        st.metric("Total Qty", int(df_scratch["Qty"].sum()))
        st.dataframe(df_scratch.sort_values(by="Date", ascending=False), use_container_width=True, height=500)

    with tab_data2:
        st.markdown("### 🟥 Production Issue Records")
        selected_year = st.radio("Select Year", sorted(df["Year"].unique(), reverse=True), horizontal=True, key="prod_year")
        df_prod = df[(df["Reason"] == "production issue") & (df["Year"] == selected_year)]
        st.metric("Total Qty", int(df_prod["Qty"].sum()))
        st.dataframe(df_prod.sort_values(by="Date", ascending=False), use_container_width=True, height=500)

# === ENTRY FORM ===
with tab3:
    st.title("📝 New Glass Rejection Entry")
    date = st.date_input("Date")
    size = st.text_input("Size")
    thickness = st.radio("Thickness (mm)", ["3mm", "4mm", "5mm", "6mm", "Other"], horizontal=True)
    glass_type = st.radio("Glass Type", ["Clear", "Lowe", "Tempered", "Tinted"], horizontal=True)
    reason = st.radio("Reason", ["Broken", "Missing", "Defective", "Production Issue", "Scratched", "Wrong Size", "Other"], horizontal=True)
    qty = st.number_input("Qty", step=1, min_value=1)
    vendor = st.radio("Vendor", ["Cardinal", "Woodbridge"], horizontal=True)
    so = st.text_input("SO")
    dept = st.radio("Department", ["Patio Door", "Other"], horizontal=True)

    month = date.strftime("%B")
    year = date.year
    week = float(date.isocalendar().week)
    formatted_date = date.strftime("%d-%m-%y")

    if st.button("Submit Entry"):
        new_row = [week, formatted_date, month, str(year), size, thickness, glass_type, reason, str(qty), vendor, so, dept]
        try:
            sheet.append_row(new_row)

            msg = MIMEMultipart()
            msg['From'] = st.secrets["email"]["sender"]
            msg['To'] = ", ".join([
                "ragavan.ramachandran@kvcustomwd.com",
                "ning.ma@kvcustomwd.com",
                "jonathan.bozanin@kvcustomwd.com"
            ])
            msg['Subject'] = "New Glass Rejection Submitted"

            body = f"""
            <h4>New Glass Rejection Entry Submitted</h4>
            <table border='1' cellpadding='5'>
                <tr><th>Field</th><th>Value</th></tr>
                {''.join([f"<tr><td>{col}</td><td>{val}</td></tr>" for col, val in zip(df.columns, new_row)])}
            </table>
            """
            msg.attach(MIMEText(body, 'html'))

            server = smtplib.SMTP(st.secrets["email"]["smtp_server"], st.secrets["email"]["port"])
            server.starttls()
            server.login(st.secrets["email"]["sender"], st.secrets["email"]["password"])
            server.sendmail(st.secrets["email"]["sender"], msg['To'].split(', '), msg.as_string())
            server.quit()

            st.success("✅ Submitted and emailed successfully!")
        except Exception as e:
            st.error(f"❌ Submission failed: {e}")
