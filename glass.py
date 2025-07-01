import streamlit as st
import pandas as pd
import plotly.express as px
import io
from datetime import datetime
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import json

st.set_page_config(page_title="Glass Rejection Dashboard", layout="wide")

# === Logo ===
st.markdown("<div style='text-align: center;'>", unsafe_allow_html=True)
st.image("KV-Logo-1.png", width=150)
st.markdown("</div>", unsafe_allow_html=True)

# === Load Data from Google Sheet ===
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds_dict = st.secrets["google_service_account"]
creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
client = gspread.authorize(creds)


SHEET_ID = "SHEET_ID = "1IoXN2Mk2MyNdp_8F8sT82l61TFuQWTk0""  # from your sheet URL
sheet = client.open_by_key(SHEET_ID).worksheet("AllData")
data = sheet.get_all_records()
df = pd.DataFrame(data)

st.success(f"✅ Loaded {len(df)} rows from Google Sheets at {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

# === Preprocess ===
df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
df["Year"] = df["Date"].dt.year
df["Quarter"] = df["Date"].dt.to_period("Q").astype(str)
df["Week#"] = df["Date"].dt.isocalendar().week
df["Reason"] = df["Reason"].astype(str)
df["Type"] = df["Type"].astype(str)

tab1, tab2, tab3 = st.tabs(["📊 Dashboard", "📝 Data Entry", "📄 Data Table"])

# === DASHBOARD TAB ===
with tab1:
    st.title("📊 Glass Rejection Intelligence Dashboard")

    # Weekly Rejections
    st.markdown("### 📅 Weekly Rejections")
    selected_year = st.radio("Choose Year", sorted(df["Year"].dropna().unique()), horizontal=True)
    df_week = df[df["Year"] == selected_year]
    weekly = df_week.groupby("Week#")["Qty"].sum().reset_index()
    fig1 = px.line(weekly, x="Week#", y="Qty", markers=True)
    fig1.update_layout(
        xaxis=dict(tickmode="linear", tick0=1, dtick=3, tickvals=list(range(1, 53))),
        shapes=[dict(type="line", x0=w, x1=w, yref="paper", y0=0, y1=1,
                     line=dict(color="cyan", width=2, dash="dot")) for w in [13, 26, 39, 52]]
    )
    st.plotly_chart(fig1, use_container_width=True)

    # Rejections by Reason
    st.markdown("### 🔍 Rejections by Reason")
    reason_year = st.radio("Year", sorted(df["Year"].unique()), horizontal=True, key="reason_year")
    df_reason = df[df["Year"] == reason_year]
    reason_data = df_reason.groupby("Reason")["Qty"].sum().reset_index()
    fig2 = px.bar(reason_data, x="Reason", y="Qty", color="Reason")
    st.plotly_chart(fig2, use_container_width=True)

    # Rejections by Glass Type
    st.markdown("### 🧊 Rejections by Glass Type")
    type_year = st.radio("Year", sorted(df["Year"].unique()), horizontal=True, key="glass_type")
    top_types = df[df["Year"] == type_year]["Type"].value_counts().nlargest(5).index.tolist()
    df_type = df[(df["Year"] == type_year) & (df["Type"].isin(top_types))]
    type_data = df_type.groupby("Type")["Qty"].sum().reset_index()
    fig3 = px.bar(type_data, x="Type", y="Qty", color="Type")
    st.plotly_chart(fig3, use_container_width=True)

    # Rejections by Department
    st.markdown("### 🏭 Rejections by Department")
    valid_quarters = [f"{y}Q{i}" for y in [2024, 2025] for i in range(1, 5) if not (y == 2025 and i > 2)]
    selected_q = st.radio("Select Quarter", valid_quarters, horizontal=True)
    df_q = df[df["Quarter"] == selected_q]
    if not df_q.empty:
        dept_data = df_q.groupby("Dept.")["Qty"].sum().reset_index()
        fig4 = px.pie(dept_data, names="Dept.", values="Qty", hole=0.4)
        st.plotly_chart(fig4, use_container_width=True)
    else:
        st.warning("No data found for the selected quarter.")

    # === Excel Export with Charts ===
    st.markdown("### 📤 Download Excel Report (with charts)")
    if st.button("📥 Generate Excel Report"):
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            # Write full data
            df.to_excel(writer, sheet_name="AllData", index=False)
            workbook = writer.book
            ws_data = writer.sheets["AllData"]
            ws_charts = workbook.add_worksheet("Charts")

            # Helper to create chart
            def create_chart(chart_type, title, category_col, value_col, position):
                chart = workbook.add_chart({'type': chart_type})
                chart.add_series({
                    'name': title,
                    'categories': ['AllData', 1, category_col, len(df), category_col],
                    'values': ['AllData', 1, value_col, len(df), value_col],
                })
                chart.set_title({'name': title})
                chart.set_style(10)
                ws_charts.insert_chart(position, chart)

            # Get column indexes
            col_map = {col: i for i, col in enumerate(df.columns)}

            # Recreate charts from dashboard
            df["Week#"] = df["Date"].dt.isocalendar().week
            create_chart('column', 'Weekly Rejections', col_map["Week#"], col_map["Qty"], "A1")
            create_chart('bar', 'Rejections by Reason', col_map["Reason"], col_map["Qty"], "A20")
            create_chart('bar', 'Rejections by Glass Type', col_map["Type"], col_map["Qty"], "A39")
            create_chart('pie', 'Rejections by Department', col_map["Dept."], col_map["Qty"], "A58")

        st.download_button(
            label="📥 Download Excel with Charts",
            data=output.getvalue(),
            file_name="Rejection_Report_With_Charts.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# === DATA ENTRY TAB ===
with tab2:
    st.info("📥 This version does not allow manual entry. Please update the Google Sheet instead.")

# === DATA TABLE TAB ===
with tab3:
    st.title("📄 All Rejection Records")
    df_table = df.sort_values(by="Date", ascending=False)
    st.dataframe(df_table, use_container_width=True, height=600)
