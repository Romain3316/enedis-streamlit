# 13. Courbe sur l’ensemble des données (version lissée et stylisée)
df_plot = df_final.copy()
df_plot["Datetime"] = pd.to_datetime(df_plot["Date"] + " " + df_plot["Heure"], dayfirst=True)

# Option : lisser avec une moyenne mobile sur 24h
df_plot["Conso_Smooth"] = df_plot["Moyenne_Conso"].rolling(window=24, min_periods=1).mean()

fig_full = px.area(
    df_plot,
    x="Datetime",
    y="Conso_Smooth",
    title="📈 Évolution de la consommation (ensemble des données)",
)

fig_full.update_traces(line=dict(width=2, color="crimson"), fill="tozeroy", opacity=0.7)
fig_full.update_layout(
    xaxis_title="Date et heure",
    yaxis_title="Consommation moyenne (lissée)",
    template="plotly_dark",
    hovermode="x unified"
)

st.plotly_chart(fig_full, use_container_width=True)
