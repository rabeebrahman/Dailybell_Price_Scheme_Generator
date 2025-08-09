import os
import pandas as pd

# Get absolute path to the current project directory
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Always point to the price slab file in the same folder as the app
PRICE_SLAB_FILE = os.path.join(BASE_DIR, "final_price_slab_for_dailybell_retail_slab_generator.xlsx")

def load_price_slab():
    """Load the fixed price slab file."""
    if not os.path.exists(PRICE_SLAB_FILE):
        raise FileNotFoundError(f"Price slab file not found: {PRICE_SLAB_FILE}")
    return pd.read_excel(PRICE_SLAB_FILE)

# Example usage in processing function
def process_quotation(quotation_file):
    price_slab_df = load_price_slab()
    quotation_df = pd.read_excel(quotation_file)

    # Your comparison logic here
    # ...
    return quotation_df
