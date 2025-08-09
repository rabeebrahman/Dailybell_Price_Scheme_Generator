import os
import pandas as pd
from datetime import datetime

# === CONFIG ===
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PRICE_SLAB_FILE = os.path.join(BASE_DIR, "final_price_slab_for_dailybell_retail_slab_generator.xlsx")


def load_price_slab():
    """Load the price slab file."""
    if not os.path.exists(PRICE_SLAB_FILE):
        raise FileNotFoundError(f"Price slab file not found: {PRICE_SLAB_FILE}")
    return pd.read_excel(PRICE_SLAB_FILE)


def process_files(quotation_file_path, output_dir):
    """
    Process the uploaded quotation file using the fixed price slab,
    and save the processed file in the output directory.
    """
    try:
        # Load data
        slabs_df = load_price_slab()
        quotation_df = pd.read_excel(quotation_file_path)

        # Standardize column names
        quotation_df.columns = quotation_df.columns.str.strip()
        slabs_df.columns = slabs_df.columns.str.strip()

        required_columns = ["Product Name", "MRP", "Condition Type", "Min Qty", "Max Qty", "Price per Unit"]
        for col in required_columns:
            if col not in slabs_df.columns:
                raise ValueError(f"Price slab file missing column: {col}")
            if col not in quotation_df.columns and col != "Price per Unit":
                raise ValueError(f"Quotation file missing column: {col}")

        # Merge to get scheme prices
        merged_df = pd.merge(
            quotation_df,
            slabs_df,
            on=["Product Name", "MRP", "Condition Type", "Min Qty", "Max Qty"],
            how="left",
            suffixes=("", "_Scheme")
        )

        # Compare prices
        merged_df["Match"] = merged_df.apply(
            lambda row: "Match" if pd.isna(row["Price per Unit_Scheme"]) or row["Price per Unit"] == row["Price per Unit_Scheme"] else "Mismatch",
            axis=1
        )

        # Save processed file
        date_prefix = datetime.now().strftime("%Y-%m-%d")
        output_filename = f"{date_prefix}_processed_quotation.xlsx"
        output_path = os.path.join(output_dir, output_filename)
        merged_df.to_excel(output_path, index=False)

        return output_path, merged_df

    except Exception as e:
        raise RuntimeError(f"Error processing files: {str(e)}")
