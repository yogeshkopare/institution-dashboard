import streamlit as st
import pandas as pd
import plotly.express as px
import os

# -------------------------------
# File path
# -------------------------------
file_path = r'ownership_data.xlsx'  # Excel file must be in the same folder as this script

if not os.path.exists(file_path):
    st.error(f"Excel file not found at {file_path}. Please check the file name!")
    st.stop()  # Stops Streamlit execution if file missing

# -------------------------------
# Load Data (Main Ownership Data)
# -------------------------------
@st.cache_data
def load_data(file_path):
    df = pd.read_excel(file_path, sheet_name="Sheet1")

    # Ensure quarter_date is datetime
    df["quarter_date"] = pd.to_datetime(df["quarter_date"], errors="coerce")
    df["quarter_date_str"] = df["quarter_date"].dt.strftime("%b %Y")  

    # Calculated Fields
    df["NetSharesChange"] = df["InstitutionSharesBought"] - df["InstitutionSharesSold"]
    df["PercentChangeHeld"] = (df["NetSharesChange"] / df["Total_SharesOutstanding"]) * 100
    df["InstitutionOwnershipValue"] = (df["InstitutionPercentHeld"] / 100) * df["Total_SharesOutstanding"]

    if "InstitutionShareHeld" not in df.columns:
        df["InstitutionShareHeld"] = (df["InstitutionPercentHeld"] / 100) * df["Total_SharesOutstanding"]

    return df

# -------------------------------
# Load Institution Holdings Data (New Sheet)
# -------------------------------
@st.cache_data
def load_institution_data(file_path):
    inst_df = pd.read_excel(file_path, sheet_name="Institution_Holdings")

    numeric_cols = ["total_market_value", "total_shares", "share_change", "share_change_percentage"]
    for col in numeric_cols:
        if col in inst_df.columns:
            inst_df[col] = pd.to_numeric(inst_df[col], errors="coerce").fillna(0)

    return inst_df

# -------------------------------
# Helper
# -------------------------------
def to_million(x):
    return round(x / 1_000_000, 2)

# -------------------------------
# Main App
# -------------------------------
def main():
    st.set_page_config(page_title="Institutional Ownership Dashboard", layout="wide")
    st.title("📊 Institutional Ownership Dashboard")
    st.markdown("A modern dashboard to analyze institutional holdings, peers, and top shareholders.")

    # -------------------------------
    # Sidebar Filters
    # -------------------------------
    st.sidebar.header("🔎 Filters")
    df = load_data(file_path)
    inst_df = load_institution_data(file_path)

    company_list = df["Company_symbol"].unique()
    company = st.sidebar.selectbox("Select Company", company_list)

    filtered_df = df[df["Company_symbol"] == company].sort_values("quarter_date")
    quarter_list = filtered_df["quarter_date_str"].unique()
    selected_quarters = st.sidebar.multiselect("Select Quarters", quarter_list, default=quarter_list)
    filtered_df = filtered_df[filtered_df["quarter_date_str"].isin(selected_quarters)]

    if filtered_df.empty:
        st.warning("⚠️ No data available for selected filters.")
        st.stop()

    # -------------------------------
    # Summary Metrics
    # -------------------------------
    latest = filtered_df.iloc[-1]
    first = filtered_df.iloc[0]

    st.markdown("### 📌 Key Metrics (Latest Quarter)")
    c1, c2, c3, c4 = st.columns(4)
    c5, c6, c7, c8 = st.columns(4)

    c1.metric("Institution % Held", f"{latest['InstitutionPercentHeld']:.2f}%")
    c2.metric("Institution Shares Held (M)", f"{to_million(latest['InstitutionShareHeld']):,}M")
    c3.metric("Net Share Change (M)", f"{to_million(latest['InstitutionShareHeld'] - first['InstitutionShareHeld']):+,.2f}M")
    c4.metric("Institution Holders", f"{latest['Institutionholdernumber']:,}")

    c5.metric("Shares Bought (Q) (M)", f"{to_million(latest['InstitutionSharesBought']):,}M")
    c6.metric("Shares Sold (Q) (M)", f"{to_million(latest['InstitutionSharesSold']):,}M")
    c7.metric("Share Float (M)", f"{to_million(latest['Sharefloat']):,}M")
    c8.metric("Shares Outstanding (M)", f"{to_million(latest['Total_SharesOutstanding']):,}M")

    st.markdown("### 📈 Change Between Selected Quarters")
    d1, d2, d3, d4 = st.columns(4)
    d5, d6, d7 = st.columns(3)

    d1.metric("Change in % Held", f"{(latest['InstitutionPercentHeld'] - first['InstitutionPercentHeld']):+.2f}%")
    d2.metric("Change in Shares Held (M)", f"{to_million(latest['InstitutionShareHeld'] - first['InstitutionShareHeld']):+,.2f}M")
    d3.metric("Change in Holders", f"{latest['Institutionholdernumber'] - first['Institutionholdernumber']:+,}")
    d4.metric("Change in Shares Bought (M)", f"{to_million(latest['InstitutionSharesBought'] - first['InstitutionSharesBought']):+,.2f}M")

    d5.metric("Change in Shares Sold (M)", f"{to_million(latest['InstitutionSharesSold'] - first['InstitutionSharesSold']):+,.2f}M")
    d6.metric("Change in Float (M)", f"{to_million(latest['Sharefloat'] - first['Sharefloat']):+,.2f}M")
    d7.metric("Change in Shares Outstanding (M)", f"{to_million(latest['Total_SharesOutstanding'] - first['Total_SharesOutstanding']):+,.2f}M")

    st.markdown("---")

    # -------------------------------
    # Tabs for Charts and Tables
    # -------------------------------
    tab1, tab2, tab3, tab4 = st.tabs([
        "📉 Ownership Trends",
        "📊 Secondary Graphs",
        "🏦 Institutions",
        "🏆 Peer Comparisons"
    ])

    # Ownership Trend
    with tab1:
        filtered_df["OwnershipCombined"] = (
            filtered_df["InstitutionPercentHeld"].map(lambda x: f"{x:.2f}%") + " | " +
            filtered_df["InstitutionShareHeld"].map(lambda x: f"{to_million(x):,}M Shares")
        )
        fig_main = px.line(
            filtered_df, x="quarter_date_str", y="InstitutionPercentHeld",
            markers=True, title="Institution Ownership Over Time",
            labels={"quarter_date_str": "Quarter", "InstitutionPercentHeld": "% Held"},
            line_shape="spline", template="plotly_dark", height=500
        )
        st.plotly_chart(fig_main, use_container_width=True)

    # Secondary Graphs
    with tab2:
        col1, col2 = st.columns([2, 1])
        with col1:
            filtered_df['NetChangeDiff'] = filtered_df['NetSharesChange'].diff().fillna(0)
            fig_net = px.bar(
                filtered_df, x="quarter_date_str", y="NetSharesChange",
                title="Net Shares Change per Quarter", color="NetSharesChange",
                color_continuous_scale=px.colors.sequential.Viridis,
                template="plotly_dark", height=400
            )
            st.plotly_chart(fig_net, use_container_width=True)
        with col2:
            fig_holders = px.line(
                filtered_df, x="quarter_date_str", y="Institutionholdernumber",
                markers=True, title="Institution Holder Count Over Time",
                labels={"quarter_date_str": "Quarter", "Institutionholdernumber": "Count"},
                line_shape="spline", template="plotly_dark", height=400
            )
            st.plotly_chart(fig_holders, use_container_width=True)

    # Top Institutions
    with tab3:
        st.subheader("Top 10 Institutions Holding Shares")
        company_inst = inst_df[inst_df["Company_symbol"] == company]
        if not company_inst.empty:
            top10_institutions = company_inst.sort_values("total_shares", ascending=False).head(10)
            chart_type = st.radio("Select Chart Type", ["Bar", "Pie"], horizontal=True)
            if chart_type == "Bar":
                fig_inst = px.bar(
                    top10_institutions, x="owner_name", y="total_shares",
                    color="total_market_value",
                    text=top10_institutions["total_shares"].map(lambda x: f"{x:,.0f}"),
                    title=f"Top 10 Institutions Holding Shares of {company}",
                    labels={"owner_name": "Institution", "total_shares": "Shares Held"},
                    template="plotly_dark", height=500
                )
                fig_inst.update_traces(textposition="outside")
                fig_inst.update_layout(xaxis_tickangle=-45)
            else:
                fig_inst = px.pie(
                    top10_institutions, names="owner_name", values="total_shares",
                    title=f"Top 10 Institutions Holding Shares of {company}",
                    template="plotly_dark", height=500, hole=0.4
                )
            st.plotly_chart(fig_inst, use_container_width=True)
            st.dataframe(top10_institutions[[
                "owner_name", "total_market_value", "total_shares",
                "share_change", "share_change_percentage"
            ]])
        else:
            st.warning("No institutional data available for this company.")

    # Peer Comparisons
    with tab4:
        st.subheader("Peers by Industry")
        industry_peers = df[(df["industry"] == latest["industry"]) & (df["Company_symbol"] != company)]
        peers_df = industry_peers.groupby("Company_symbol").apply(
            lambda x: x.loc[x['quarter_date'].idxmax()]
        ).reset_index(drop=True)
        peers_df = peers_df.sort_values("InstitutionPercentHeld", ascending=False).head(10)
        st.dataframe(peers_df[[
            "Company_symbol", "Company_name", "InstitutionPercentHeld",
            "InstitutionShareHeld", "Institutionholdernumber",
            "Sharefloat", "industry", "quarter_date_str"
        ]])

        st.subheader("Top 10 Companies by Institution % Held (Latest Quarter)")
        latest_quarter_df = df.loc[df.groupby("Company_symbol")["quarter_date"].idxmax()]
        top10_df = latest_quarter_df.sort_values("InstitutionPercentHeld", ascending=False).head(10)
        st.dataframe(top10_df[[
            "Company_symbol", "Company_name", "InstitutionPercentHeld",
            "InstitutionShareHeld", "Institutionholdernumber",
            "Sharefloat", "industry", "quarter_date_str"
        ]])

    # -------------------------------
    # Download Processed Data
    # -------------------------------
    st.sidebar.markdown("---")
    st.sidebar.subheader("⬇ Download Data")
    output_file = "processed_ownership_data.xlsx"
    df.to_excel(output_file, index=False)
    with open(output_file, "rb") as f:
        st.sidebar.download_button("Download Excel", f, file_name=output_file, mime="application/vnd.ms-excel")


# -------------------------------
if __name__ == "__main__":
    main()
