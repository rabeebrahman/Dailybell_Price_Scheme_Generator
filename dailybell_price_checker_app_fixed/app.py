import os
import threading
import pandas as pd
from flask import Flask, render_template, request, send_file, jsonify
from datetime import datetime
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
OUTPUT_FOLDER = "outputs"
PRICE_SLAB_PATH = os.path.join("price_slabs", "final_price_slab_for_dailybell_retail_slab_generator.xlsx")

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)
os.makedirs("price_slabs", exist_ok=True)


# Auto-delete after delay
def delete_file_after_delay(filepath, delay=300):
    def delete_file():
        try:
            if os.path.exists(filepath):
                os.remove(filepath)
        except Exception as e:
            print(f"Error deleting file: {e}")

    threading.Timer(delay, delete_file).start()


# Price checking logic
def process_price_checker(quotation_file):
    try:
        df_quote = pd.read_excel(quotation_file)
        df_slab = pd.read_excel(PRICE_SLAB_PATH)

        merged_df = pd.merge(
            df_quote,
            df_slab,
            on=["Product Name", "MRP", "Condition Type", "Min Qty", "Max Qty", "Price per Unit"],
            how="left",
            suffixes=("", "_slab")
        )

        # Example calculation (adjust per your rules)
        merged_df["Scheme Price"] = merged_df["Price per Unit"]
        merged_df["Reason"] = merged_df.apply(
            lambda row: "No Match" if pd.isna(row["Price per Unit"]) else "",
            axis=1
        )

        # Save XLSX
        date_prefix = datetime.now().strftime("%Y-%m-%d")
        output_xlsx = os.path.join(OUTPUT_FOLDER, f"{date_prefix}_price_check_result.xlsx")
        merged_df.to_excel(output_xlsx, index=False)

        # Save PDF
        output_pdf = os.path.join(OUTPUT_FOLDER, f"{date_prefix}_price_check_result.pdf")
        generate_pdf(merged_df, output_pdf)

        delete_file_after_delay(output_xlsx)
        delete_file_after_delay(output_pdf)

        return output_xlsx, output_pdf
    except Exception as e:
        print(f"Error processing file: {e}")
        return None, None


# PDF generation with column name mapping
def generate_pdf(df, pdf_path):
    doc = SimpleDocTemplate(pdf_path, pagesize=A4)
    elements = []
    styles = getSampleStyleSheet()
    data = [df.columns.tolist()] + df.values.tolist()

    table_data = []
    for row in data:
        formatted = [Paragraph(str(cell), styles["BodyText"]) for cell in row]
        table_data.append(formatted)

    table = Table(table_data, repeatRows=1)
    style = TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.gray),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.whitesmoke),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 7),
        ("BOTTOMPADDING", (0, 0), (-1, 0), 6),
        ("GRID", (0, 0), (-1, -1), 0.25, colors.black),
    ])

    # Dynamic column lookup
    price_with_tax_col = df.columns.get_loc("Price with Tax")
    scheme_price_col = df.columns.get_loc("Scheme Price")
    reason_col = df.columns.get_loc("Reason")

    for idx, row in enumerate(df.itertuples(index=False), start=1):
        tax_price = row[price_with_tax_col]
        scheme_price = row[scheme_price_col]
        reason = row[reason_col]

        if (
            isinstance(tax_price, (int, float))
            and isinstance(scheme_price, (int, float))
            and round(tax_price, 2) != round(scheme_price, 2)
            and reason != "No Match"
        ):
            style.add("BACKGROUND", (0, idx), (-1, idx), colors.lightyellow)

    table.setStyle(style)
    elements.append(table)
    doc.build(elements)


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/process", methods=["POST"])
def process_file():
    if "quotation_file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400

    file = request.files["quotation_file"]
    if file.filename == "":
        return jsonify({"error": "No file selected"}), 400

    filepath = os.path.join(UPLOAD_FOLDER, file.filename)
    file.save(filepath)
    delete_file_after_delay(filepath)

    xlsx_path, pdf_path = process_price_checker(filepath)

    if not xlsx_path or not pdf_path:
        return jsonify({"error": "Processing failed"}), 500

    return jsonify({
        "xlsx_url": f"/download/{os.path.basename(xlsx_path)}",
        "pdf_url": f"/download/{os.path.basename(pdf_path)}"
    })


@app.route("/download/<filename>")
def download_file(filename):
    filepath = os.path.join(OUTPUT_FOLDER, filename)
    if os.path.exists(filepath):
        return send_file(filepath, as_attachment=True)
    return "File not found", 404


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
