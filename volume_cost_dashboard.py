import pandas as pd
import streamlit as st
import plotly.express as px

# Load Data
@st.cache_data
def load_data():
    forecast_df = pd.read_excel("Outsourced Forecast .xlsx", sheet_name="FY26 PLAN", header=4)
    cost_df1 = pd.read_excel("Copy of OUTSIDE SERVICES_ cost per contact_FY26_April 2025.xlsx", sheet_name="Global English  & MT")
    cost_df2 = pd.read_excel("Copy of OUTSIDE SERVICES_ cost per contact_FY26_April 2025.xlsx", sheet_name="EMEA_LAM TPG")
    cost_df3 = pd.read_excel("Copy of OUTSIDE SERVICES_ cost per contact_FY26_April 2025.xlsx", sheet_name="Japanese")
    cost_df = pd.concat([cost_df1, cost_df2, cost_df3], ignore_index=True)
    return forecast_df, cost_df

forecast_df, cost_df = load_data()

# Clean Forecast Data
forecast_df = forecast_df.rename(columns={forecast_df.columns[1]: "Month"})
forecast_df = forecast_df.dropna(subset=["Month"])
forecast_df = forecast_df[forecast_df['Month'].apply(lambda x: isinstance(x, str))]

melted = forecast_df.melt(id_vars=["Month"], var_name="Channel_LOB", value_name="Volume")
melted = melted.dropna(subset=["Volume"])
melted[['Channel', 'LOB']] = melted['Channel_LOB'].str.extract(r'(Chat|Phone|Email|Remote)?\s*(.*)', expand=True)
melted['Channel'] = melted['Channel'].fillna("Unknown")

# Merge with Cost Data
cost_df = cost_df.dropna(subset=['Channel', 'Cost per Contact'])
combined = pd.merge(melted, cost_df, how='left', left_on=['Channel', 'LOB'], right_on=['Channel', 'Line Of Business'])
combined['Cost per Contact'] = pd.to_numeric(combined['Cost per Contact'], errors='coerce')
combined['Total Cost'] = combined['Volume'] * combined['Cost per Contact']
combined['Volume'] = pd.to_numeric(combined['Volume'], errors='coerce')

# Sidebar Filters
st.sidebar.header("Filter by")
channels = st.sidebar.multiselect("Channel", options=combined['Channel'].dropna().unique(), default=combined['Channel'].dropna().unique())
lobs = st.sidebar.multiselect("Line of Business", options=combined['LOB'].dropna().unique(), default=combined['LOB'].dropna().unique())
languages = st.sidebar.multiselect("Language", options=combined['Language'].dropna().unique(), default=combined['Language'].dropna().unique())

filtered = combined[
    (combined['Channel'].isin(channels)) &
    (combined['LOB'].isin(lobs)) &
    (combined['Language'].isin(languages))
]

# KPIs
st.title("📊 Volume & Cost per Contact Dashboard")
st.metric("Total Volume", f"{int(filtered['Volume'].sum()):,}")
st.metric("Avg Cost per Contact", f"${filtered['Cost per Contact'].mean():.2f}")
st.metric("Total Cost", f"${filtered['Total Cost'].sum():,.2f}")

# Charts
st.subheader("📈 Monthly Volume by Channel")
monthly_volume = filtered.groupby(['Month', 'Channel'])['Volume'].sum().reset_index()
fig_volume = px.line(monthly_volume, x='Month', y='Volume', color='Channel', markers=True)
st.plotly_chart(fig_volume)

st.subheader("💰 Monthly Total Cost by Channel")
monthly_cost = filtered.groupby(['Month', 'Channel'])['Total Cost'].sum().reset_index()
fig_cost = px.line(monthly_cost, x='Month', y='Total Cost', color='Channel', markers=True)
st.plotly_chart(fig_cost)

st.subheader("📦 Volume by Line of Business")
lob_vol = filtered.groupby('LOB')['Volume'].sum().sort_values(ascending=False).reset_index()
fig_lob = px.bar(lob_vol, x='LOB', y='Volume', title='Total Volume by LOB')
st.plotly_chart(fig_lob)

st.subheader("🌐 Volume by Language")
lang_vol = filtered.groupby('Language')['Volume'].sum().sort_values(ascending=False).reset_index()
fig_lang = px.bar(lang_vol, x='Language', y='Volume', title='Total Volume by Language')
st.plotly_chart(fig_lang)
