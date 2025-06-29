import streamlit as st
import pandas as pd
import os
from portfolio_generator import load_data, create_portfolio

st.set_page_config(page_title="Child Portfolio Generator", layout="centered")

st.title("📚 Monthly Learning Portfolio Generator")

# Direct download URL from Google Drive
drive_url = "https://drive.google.com/uc?export=download&id=1V3dFqK5RHGe9DU1-PH1T32wO1GOydCh-"

if drive_url:
    # Load data
    may_df, june_df, desc_df = load_data(drive_url)

    # Enter child's unique id to select child
    child_id_input = st.text_input("Enter the Child ID to generate the portfolio")

    if child_id_input:
        if child_id_input in june_df['Child ID'].values:
            row_june = june_df[june_df['Child ID'] == child_id_input].iloc[0]
            row_may = may_df[may_df['Child ID'] == child_id_input].iloc[0]

            if st.button("Generate Portfolio"):
                output_path = "generated_docs"
                os.makedirs(output_path, exist_ok=True)
                filepath = create_portfolio(row_may, row_june, desc_df, output_path)
                with open(filepath, "rb") as f:
                    st.success(f"Portfolio generated for {row_june['Child Name']}")
                    st.download_button(
                        label="📥 Download Portfolio",
                        data=f,
                        file_name=os.path.basename(filepath),
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
    else:
        st.error("❌ Child ID not found. Please check and try again.")
