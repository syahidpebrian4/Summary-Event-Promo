import streamlit as st
import pandas as pd
import re

st.title("Event Summary")

uploaded_file = st.file_uploader(
    "Upload file Excel promo",
    type=["xlsx"]
)

if uploaded_file is not None:

    xls = pd.ExcelFile(uploaded_file)
    all_summaries = []

    for sheet_name in xls.sheet_names:

        try:
            df = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=None)

            # ===============================
            # AMBIL DATA DASAR
            # ===============================
            a3 = str(df.iloc[2, 0])
            a4 = str(df.iloc[3, 0])

            all_count = df.iloc[7, 3]
            all_claim = df.iloc[7, 4]

            sales_amount = df.iloc[7,12] if df.iloc[5,12] == "Amount" else "-"
            amount = df.iloc[7,13] if df.iloc[5,13] == "Amount" else "-"

            # ===============================
            # LEFT (AMAN DARI ERROR INDEX)
            # ===============================
            left = "-"
            if df.shape[1] > 16:
                if df.iloc[5,16] == "Left":
                    left = df.iloc[7,16]
                else:
                    left = df.iloc[7,15]

            # ===============================
            # NAMA PROMO (STOP DI KOMA / TANGGAL)
            # ===============================
            match = re.search(r"-\s*(.*?)(?=,|\s\d|$)", a3)
            nama_promo = match.group(1).strip() if match else "-"
            nama_promo = nama_promo.rstrip(",")

            # ===============================
            # MEKANISME & PERIODE
            # ===============================
            mekanisme = a4.split(" ", 1)[1].strip() if " " in a4 else "-"
            periode = a3.split(",")[-1].strip() if "," in a3 else "-"

            # ===============================
            # SIMPAN SUMMARY
            # ===============================
            all_summaries.append({
                "Sheet": sheet_name,
                "Nama Promo": nama_promo,
                "Mekanisme Promo": mekanisme,
                "Periode Promo": periode,
                "All Count": all_count,
                "All Claim": all_claim,
                "Sales Amount": sales_amount,
                "Amount": amount,
                "Left": left
            })

        except Exception:
            st.warning(f"Sheet {sheet_name} dilewati (format tidak sesuai)")

    # ===============================
    # TAMPILKAN & DOWNLOAD
    # ===============================
    if all_summaries:
        result_df = pd.DataFrame(all_summaries)

        st.subheader("Summary Semua Promo")
        st.dataframe(result_df)

        output_file = "promo_summary_all.xlsx"
        result_df.to_excel(output_file, index=False)

        with open(output_file, "rb") as f:
            st.download_button(
                label="Download Summary Excel",
                data=f,
                file_name=output_file,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
