                template="plotly_white",
                height=920,
                xaxis_title="Jour de la semaine",
                yaxis_title="Semaine",
            )

            calendar_heatmap.update_yaxes(
                autorange="reversed"
            )

            st.plotly_chart(
                calendar_heatmap,
                use_container_width=True,
            )

        st.subheader("Tableau détaillé")

        daily_display = daily_df[
            [
                "Date",
                "Jour",
                "Consommation_kWh",
                "Puissance_moyenne_kW",
                "Puissance_max_kW",
            ]
        ].copy()

        daily_display["Date"] = (
            daily_display["Date"]
            .dt.strftime("%d/%m/%Y")
        )

        daily_style = (
            daily_display.style
            .format(
                {
                    "Consommation_kWh": "{:.2f}",
                    "Puissance_moyenne_kW": "{:.2f}",
                    "Puissance_max_kW": "{:.2f}",
                }
            )
            .background_gradient(
                subset=["Consommation_kWh"],
                cmap="RdYlGn_r",
            )
        )

        st.dataframe(
            daily_style,
            use_container_width=True,
            height=520,
        )


# ============================================================
# QUALITÉ DES DONNÉES
# ============================================================

if workspace_page == "Analyse":
    with tab_quality:
        quality1, quality2, quality3, quality4 = st.columns(4)

        quality1.metric(
            "Couverture des relevés",
            f"{coverage_percent:.2f} %".replace(".", ","),
        )
        quality2.metric(
            "Relevés manquants",
            f"{missing_points_count:,}".replace(",", " "),
        )
        quality3.metric(
            "Jours à contrôler",
            len(atypical_quality_df),
        )
        quality4.metric(
            "Doublons hors heure d'hiver",
            duplicate_count,
        )

        st.caption(
            f"Référence : {expected_points_per_day} points pour une journée normale "
            f"avec un pas de {int(time_step.total_seconds() / 60)} minutes. "
            "Les journées de changement d'heure sont contrôlées spécifiquement "
            "sur 23 h ou 25 h. Les valeurs absentes ne sont pas interpolées."
        )

        quality_display = quality_report_df.copy()
        quality_display["Date"] = quality_display["Date"].dt.strftime("%d/%m/%Y")
        st.dataframe(
            quality_display,
            use_container_width=True,
            height=520,
            hide_index=True,
        )

        if atypical_quality_df.empty and duplicate_count == 0:
            st.success(
                "Aucune anomalie de complétude détectée. Les journées de passage "
                "à l'heure d'été et à l'heure d'hiver sont traitées selon leur "
                "durée réelle."
            )
        else:
            st.warning(
                "Des données sont à contrôler. Les totaux sont calculés uniquement "
                "à partir des relevés présents : aucune consommation manquante "
                "n'est reconstituée automatiquement."
            )

            if not atypical_quality_df.empty:
                atypical_display = atypical_quality_df.copy()
                atypical_display["Date"] = atypical_display["Date"].dt.strftime("%d/%m/%Y")
                st.dataframe(
                    atypical_display,
                    use_container_width=True,
                    hide_index=True,
                )


# ============================================================
# EXPORT
# ============================================================



# Render after all widgets so this download captures their latest values.
render_dossier_download(header_save)
