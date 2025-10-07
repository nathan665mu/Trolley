import os
import secrets
from datetime import datetime
from pathlib import Path

from flask import Flask, render_template, request, redirect, url_for, send_file, flash
import pandas as pd
from werkzeug.utils import secure_filename

# Adjust sys.path to import from project root
import sys
BASE_DIR = Path(__file__).resolve().parents[1]
if str(BASE_DIR) not in sys.path:
    sys.path.insert(0, str(BASE_DIR))

# Reuse core logic from the existing script at project root
from trolley_interactive_enhanced_v2 import (
    parse_sku_name,
    get_search_results,
    find_best_match,
)


APP_DIR = Path(__file__).parent.resolve()
UPLOAD_DIR = BASE_DIR / "uploads"
RESULTS_DIR = BASE_DIR / "results"
UPLOAD_DIR.mkdir(exist_ok=True)
RESULTS_DIR.mkdir(exist_ok=True)


def allowed_file(filename: str) -> bool:
    return "." in filename and filename.rsplit(".", 1)[1].lower() in {"xlsx", "xls"}


def build_job_id() -> str:
    return secrets.token_hex(8)


def safe_int(value, default=None):
    try:
        return int(value)
    except Exception:
        return default


def process_dataframe(df: pd.DataFrame, column_name: str, limit: int | None = None, rate_limit: bool = False):
    results: list[dict] = []

    total_rows = len(df) if limit is None else min(len(df), limit)
    series = df[column_name].dropna().head(total_rows)

    for product_name in series:
        name = str(product_name).strip()
        if not name or name.lower() == "nan":
            continue

        brand, size, quantity = parse_sku_name(name)
        search_results, search_url = get_search_results(name)

        if not search_results:
            results.append({
                "SKU_Name": name,
                "Parsed_Brand": brand,
                "Parsed_Size": size,
                "Parsed_Quantity": quantity,
                "Search_URL": search_url,
                "Match_Status": "No Results",
                "Matched_Brand": "N/A",
                "Matched_Description": "N/A",
                "Matched_Size": "N/A",
                "Matched_Quantity": "N/A",
                "Matched_Price": "N/A",
                "Matched_URL": "N/A",
                "Match_Tier": "None",
            })
            continue

        best_match = find_best_match(brand, size, quantity, search_results)

        if best_match:
            brand_match = brand.lower() in best_match.brand.lower() if brand else True
            size_match = size.lower() in best_match.size.lower() if size else True
            qty_match = str(quantity) == best_match.quantity if quantity else True

            if brand_match and size_match and qty_match:
                match_tier = "Tier 1 (Perfect)"
            elif brand_match and size_match:
                match_tier = "Tier 2 (Brand+Size)"
            elif brand_match:
                match_tier = "Tier 3 (Brand Only)"
            else:
                match_tier = "Tier 4 (Fallback)"

            results.append({
                "SKU_Name": name,
                "Parsed_Brand": brand,
                "Parsed_Size": size,
                "Parsed_Quantity": quantity,
                "Search_URL": search_url,
                "Match_Status": "Matched",
                "Matched_Brand": best_match.brand,
                "Matched_Description": best_match.description,
                "Matched_Size": best_match.size,
                "Matched_Quantity": best_match.quantity,
                "Matched_Price": best_match.price,
                "Matched_URL": best_match.url,
                "Match_Tier": match_tier,
            })
        else:
            results.append({
                "SKU_Name": name,
                "Parsed_Brand": brand,
                "Parsed_Size": size,
                "Parsed_Quantity": quantity,
                "Search_URL": search_url,
                "Match_Status": "No Match",
                "Matched_Brand": "N/A",
                "Matched_Description": "N/A",
                "Matched_Size": "N/A",
                "Matched_Quantity": "N/A",
                "Matched_Price": "N/A",
                "Matched_URL": "N/A",
                "Match_Tier": "None",
            })

    # Save CSV to project-level results folder
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_path = RESULTS_DIR / f"trolley_results_{timestamp}.csv"
    pd.DataFrame(results).to_csv(output_path, index=False, encoding="utf-8-sig")
    return results, str(output_path)


def create_app():
    app = Flask(__name__, template_folder=str(APP_DIR / "templates"))
    app.config["SECRET_KEY"] = os.environ.get("FLASK_SECRET_KEY", secrets.token_hex(16))
    app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024

    @app.route("/")
    def index():
        return render_template("index.html")

    @app.route("/upload", methods=["POST"]) 
    def upload():
        if "file" not in request.files:
            flash("No file part in request.", "error")
            return redirect(url_for("index"))

        file = request.files["file"]
        if file.filename == "":
            flash("No file selected.", "error")
            return redirect(url_for("index"))

        if not allowed_file(file.filename):
            flash("Unsupported file type. Please upload an .xlsx or .xls file.", "error")
            return redirect(url_for("index"))

        job_id = build_job_id()
        fname = secure_filename(file.filename)
        dest_path = UPLOAD_DIR / f"{job_id}__{fname}"
        file.save(dest_path)

        try:
            df = pd.read_excel(dest_path)
        except Exception as e:
            flash(f"Failed to read Excel file: {e}", "error")
            return redirect(url_for("index"))

        samples = []
        for col in df.columns:
            vals = df[col].dropna().astype(str).head(3).tolist()
            samples.append({"name": col, "samples": vals})

        return render_template(
            "select_column.html",
            job_id=job_id,
            filename=fname,
            columns=df.columns.tolist(),
            samples=samples,
            total_rows=len(df),
        )

    @app.route("/process", methods=["POST"]) 
    def process():
        job_id = request.form.get("job_id")
        column_name = request.form.get("column")
        limit_mode = request.form.get("limit_mode", "5")
        custom_limit = safe_int(request.form.get("custom_limit"), None)

        files = list(UPLOAD_DIR.glob(f"{job_id}__*"))
        if not files:
            flash("Upload not found. Please upload the file again.", "error")
            return redirect(url_for("index"))
        file_path = files[0]

        try:
            df = pd.read_excel(file_path)
        except Exception as e:
            flash(f"Failed to read Excel file: {e}", "error")
            return redirect(url_for("index"))

        if column_name not in df.columns:
            flash("Selected column not found in file.", "error")
            return redirect(url_for("index"))

        limit_map = {"5": 5, "10": 10, "50": 50, "all": None, "custom": custom_limit}
        limit = limit_map.get(limit_mode, 5)
        if limit_mode == "custom" and (limit is None or limit <= 0):
            flash("Please enter a valid custom number of rows.", "error")
            return redirect(url_for("index"))

        if limit is None:
            limit = min(100, len(df))

        results, csv_path = process_dataframe(df, column_name, limit=limit, rate_limit=False)

        df_results = pd.DataFrame(results)
        tier_counts = df_results["Match_Tier"].value_counts().to_dict() if not df_results.empty else {}

        return render_template(
            "results.html",
            job_id=job_id,
            filename=os.path.basename(file_path),
            column_name=column_name,
            total_processed=len(results),
            results=results,
            csv_path=os.path.basename(csv_path),
            tier_counts=tier_counts,
        )

    @app.route("/download/<path:csv_name>")
    def download(csv_name: str):
        csv_path = RESULTS_DIR / csv_name
        if not csv_path.exists():
            flash("File not found.", "error")
            return redirect(url_for("index"))
        return send_file(csv_path, as_attachment=True, download_name=csv_path.name)

    return app


app = create_app()

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)), debug=True)
