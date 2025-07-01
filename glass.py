import streamlit as st
import pandas as pd
import plotly.express as px
import io
from datetime import datetime

st.set_page_config(page_title="Glass Rejection Dashboard", layout="wide")
st.markdown("✅ App is running on Streamlit Cloud.")

# === Custom Theme ===
st.markdown("""
<style>
html, body, [class*="css"] {
    background-color: #121212;
    color: white;
}
.stButton>button, .stDownloadButton>button {
    background-color: #00c8c8;
    color: #121212;
    padding: 6px 14px;
    border-radius: 6px;
}
</style>
""", unsafe_allow_html=True)

# === Logo at Top ===
st.markdown("<div style='text-align: center;'>", unsafe_allow_html=True)
st.image("KV-Logo-1.png", width=150)
st.markdown("</div>", unsafe_allow_html=True)

# === File Upload ===
df=pd.read_excel("Glassline Rejection Data.xlsx" \
"")

df = pd.read_excel(uploaded_file)
df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
df["Year"] = df["Date"].dt.year
df["Month"] = df["Date"].dt.month
df["Quarter"] = df["Date"].dt.to_period("Q").astype(str)
df["Week#"] = df["Date"].dt.isocalendar().week
df["MonthYear"] = df["Date"].dt.to_period("M").astype(str)
df["MonthYearSort"] = df["Date"].dt.strftime("%Y%m").astype(int)
df["Reason"] = df["Reason"].astype(str)
df["Type"] = df["Type"].astype(str)

# === Tabs ===
tab1, tab2, tab3 = st.tabs(["📊 Dashboard", "📝 Data Entry", "📄 Data Table"])

# === DASHBOARD TAB ===
with tab1:
    st.title("📊 Glass Rejection Intelligence Dashboard")

    st.markdown("### 📅 Weekly Rejections")
    selected_year = st.radio("Choose Year", sorted(df["Year"].dropna().unique()), horizontal=True)
    df_week = df[df["Year"] == selected_year]
    weekly = df_week.groupby("Week#")["Qty"].sum().reset_index()
    fig1 = px.line(weekly, x="Week#", y="Qty", markers=True, template="plotly_dark")
    fig1.update_layout(
        xaxis=dict(tickmode="linear", tick0=1, dtick=3, tickvals=list(range(1, 53)), title="Week Number"),
        shapes=[dict(type="line", x0=w, x1=w, yref="paper", y0=0, y1=1,
                     line=dict(color="cyan", width=2, dash="dot")) for w in [13, 26, 39, 52]]
    )
    st.plotly_chart(fig1, use_container_width=True)

    st.markdown("### 🔍 Rejections by Reason")
    reason_year = st.radio("Year", sorted(df["Year"].unique()), horizontal=True, key="reason_year")
    df_reason = df[df["Year"] == reason_year]
    reason_data = df_reason.groupby("Reason")["Qty"].sum().reset_index()
    fig2 = px.bar(reason_data, x="Reason", y="Qty", color="Reason", template="plotly_dark")
    st.plotly_chart(fig2, use_container_width=True)

    st.markdown("### 🧊 Rejections by Glass Type")
    type_year = st.radio("Year", sorted(df["Year"].unique()), horizontal=True, key="glass_type")
    top_types = df[df["Year"] == type_year]["Type"].value_counts().nlargest(5).index.tolist()
    df_type = df[(df["Year"] == type_year) & (df["Type"].isin(top_types))]
    type_data = df_type.groupby("Type")["Qty"].sum().reset_index()
    fig3 = px.bar(type_data, x="Type", y="Qty", color="Type", template="plotly_dark")
    st.plotly_chart(fig3, use_container_width=True)

    st.markdown("### 🏭 Rejections by Department")
    valid_quarters = [f"{y}Q{i}" for y in [2024, 2025] for i in range(1, 5) if not (y == 2025 and i > 2)]
    selected_q = st.radio("Select Quarter", valid_quarters, horizontal=True)
    df_q = df[df["Quarter"] == selected_q]
    if not df_q.empty:
        dept_data = df_q.groupby("Dept.")["Qty"].sum().reset_index()
        fig4 = px.pie(dept_data, names="Dept.", values="Qty", template="plotly_dark", hole=0.4)
        st.plotly_chart(fig4, use_container_width=True)
    else:
        st.warning("No data found for the selected quarter.")

    # Excel Export
    st.markdown("### 📤 Download Excel Report (with charts)")
    if st.button("📥 Generate Excel Report"):
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            df.to_excel(writer, sheet_name="AllData", index=False)
            wb = writer.book
            ws = writer.sheets["AllData"]

            reason_summary = df.groupby("Reason")["Qty"].sum().reset_index()
            reason_summary.to_excel(writer, sheet_name="ChartData", startrow=0, index=False)

            chart = wb.add_chart({'type': 'column'})
            chart.add_series({
                'name': 'Qty by Reason',
                'categories': ['ChartData', 1, 0, len(reason_summary), 0],
                'values': ['ChartData', 1, 1, len(reason_summary), 1],
            })
            chart.set_title({'name': 'Qty by Reason'})
            chart.set_x_axis({'name': 'Reason'})
            chart.set_y_axis({'name': 'Qty'})
            ws.insert_chart('L2', chart)

        st.download_button(
            label="📥 Download Excel",
            data=output.getvalue(),
            file_name="Rejection_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# === DATA ENTRY TAB (Disabled since we're only uploading) ===
with tab2:
    st.info("📥 Please upload an Excel file on the top. Manual entry is disabled in this version.")

# === DATA TABLE TAB ===
with tab3:
    st.title("📄 All Rejection Records")
    df_table = df.sort_values(by="Date", ascending=False)
    st.dataframe(df_table, use_container_width=True, height=600)
