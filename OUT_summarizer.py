import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill, Alignment, Border, Side, Font

st.title("📊 OUT Summarizer — Country Report Generator")

uploaded_file = st.file_uploader("📂 Upload Excel file", type=["xlsx", "xls"])

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)
    df.columns = df.columns.str.strip()

    # Fill merged cells
    for col in ["Start", "End", "Channel", "ID"]:
        if col in df.columns:
            df[col] = df[col].ffill()

    if 'Country' not in df.columns:
        st.error("❌ Column 'Country' not found!")
        st.stop()
    df['Country'] = df['Country'].astype(str).str.strip().str.upper()

    if 'Start' not in df.columns:
        st.error("❌ Column 'Start' missing!")
        st.stop()
    df['Start'] = pd.to_datetime(
        df['Start'].astype(str).str.strip().str.replace(r"[-_.\\]", "/", regex=True),
        errors='coerce', infer_datetime_format=True, dayfirst=False
    ).dt.date

    parsed = df['Start'].notna().sum()
    st.info(f"📅 Parsed {parsed} / {len(df)} dates in 'Start' column")
    if parsed == 0:
        st.error("❌ No valid dates in 'Start'. Check Excel formatting.")
        st.stop()

    min_date = df['Start'].min()
    max_date = df['Start'].max()
    date_range = st.date_input(
        "📅 Select date range (based on 'Start')",
        value=(min_date, max_date), min_value=min_date, max_value=max_date
    )

    if 'Category' not in df.columns:
        st.error("❌ Column 'Category' missing!")
        st.stop()
    df['Category_norm'] = df['Category'].astype(str).str.strip().str.lower()
    categories = sorted(df['Category_norm'].unique())
    selected_category = st.selectbox("🏷️ Select category", categories)

    if st.button("🚀 Generate report"):
        start_date, end_date = date_range
        mask = (
            (df['Category_norm'] == selected_category) &
            (df['Start'] >= start_date) &
            (df['Start'] <= end_date)
        )
        filtered_df = df.loc[mask].copy()
        if filtered_df.empty:
            st.warning("⚠️ No data for these filters.")
            st.stop()

        countries = sorted(filtered_df["Country"].dropna().unique())
        keep_cols = [
            'Start', 'End', 'Channel', 'ID', 'Name', 'Description', 'Category',
            'Visits', 'Orders', 'Demand', 'CVR', 'AOV',
            'Expected Demand', 'Demand Diff to Expected', '% Expected Demand',
            'Country'
        ]
        existing_cols = [c for c in keep_cols if c in filtered_df.columns]
        filtered_df = filtered_df[existing_cols]

        numeric_cols = [
            'Visits', 'Orders', 'Demand', 'CVR', 'AOV',
            'Expected Demand', 'Demand Diff to Expected', '% Expected Demand'
        ]
        numeric_cols = [c for c in numeric_cols if c in filtered_df.columns]

        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            for country in countries:
                country_df = filtered_df[filtered_df['Country'] == country].copy()
                if not country_df.empty:
                    means = country_df[numeric_cols].mean().to_frame().T
                    means.index = ['Average']

                    # Podsumowanie — zaokrąglenia
                    if 'Orders' in means.columns:
                        means['Orders'] = round(means['Orders'])
                    for col in ['Demand', 'Expected Demand']:
                        if col in means.columns:
                            means[col] = round(means[col], 2)
                    for col in ['CVR', '% Expected Demand']:
                        if col in means.columns:
                            means[col] = round(means[col], 4)

                    country_df[numeric_cols] = country_df[numeric_cols].round(2)
                    final_df = pd.concat([country_df, means], ignore_index=False)
                    final_df.to_excel(writer, index=False, sheet_name=str(country))

                    ws = writer.sheets[str(country)]

                    # Style
                    thin = Side(border_style="thin", color="000000")
                    border = Border(left=thin, right=thin, top=thin, bottom=thin)
                    yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
                    blue_fill = PatternFill(start_color="ADD8E6", end_color="ADD8E6", fill_type="solid")
                    red_fill = PatternFill(start_color="FF9999", end_color="FF9999", fill_type="solid")
                    green_fill = PatternFill(start_color="99FF99", end_color="99FF99", fill_type="solid")
                    light_blue_fill = PatternFill(start_color="ADD8E6", end_color="ADD8E6", fill_type="solid")

                    # Nagłówki
                    for col_idx, col in enumerate(final_df.columns, 1):
                        ws.column_dimensions[get_column_letter(col_idx)].width = max(
                            12, max((len(str(cell.value)) if cell.value is not None else 0 for cell in ws[get_column_letter(col_idx)]), default=12)
                        )
                        cell = ws.cell(row=1, column=col_idx)
                        cell.fill = yellow_fill
                        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                        cell.font = Font(bold=True)
                        cell.border = border

                    # Wiersze danych i podsumowania
                    for row_idx in range(2, ws.max_row + 1):
                        for col_idx, col in enumerate(final_df.columns, 1):
                            cell = ws.cell(row=row_idx, column=col_idx)

                            if row_idx < ws.max_row:
                                # Dane
                                cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
                                cell.border = border
                                if col in ["Demand","Expected Demand","Demand Diff to Expected"]:
                                    cell.number_format = '€#,##0.00'
                                    if col in ["Demand Diff to Expected","% Expected Demand"] and isinstance(cell.value,(int,float)):
                                        cell.fill = green_fill if cell.value >=0 else red_fill
                                if col in ["CVR","% Expected Demand"]:
                                    cell.number_format = '0.00%'
                            else:
                                # PODSUMOWANIE (ostatni wiersz)
                                cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
                                cell.border = border

                                if col in ["Start","End","Channel","ID","Name","Description","Country"]:
                                    cell.fill = light_blue_fill

                                elif col == "Category":
                                    cell.value = "ARV"
                                    cell.fill = light_blue_fill

                                elif col in ["Visits","Orders","Demand","CVR","AOV",
                                             "Expected Demand","Demand Diff to Expected","% Expected Demand"]:

                                    # Jasnoniebieskie wypełnienie — poprawka tu!
                                    cell.fill = light_blue_fill

                                    if col == "Visits" and isinstance(cell.value,float):
                                        cell.value = round(cell.value)

                                    if col == "AOV" and isinstance(cell.value,float):
                                        cell.value = round(cell.value)

                                    if col in ["Demand","Expected Demand","Demand Diff to Expected"]:
                                        cell.number_format = '€#,##0.00'

                                    if col in ["CVR","% Expected Demand"]:
                                        cell.number_format = '0.00%'

                                else:
                                    # Pola z podsumowania bez wypełnienia
                                    cell.fill = PatternFill(fill_type=None)

                else:
                    pd.DataFrame({"Info": [f"No data for {country}"]}).to_excel(
                        writer, index=False, sheet_name=str(country)
                    )

            # Puste arkusze
            for empty_sheet in ["Brands", "Category", "Stock level", "Conclusions"]:
                pd.DataFrame().to_excel(writer, sheet_name=empty_sheet)

        output.seek(0)
        filename = f"{selected_category}_{start_date}_{end_date}.xlsx".replace(" ", "_")
        st.success("✅ Report ready!")
        st.download_button(
            "📥 Download report",
            data=output,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
