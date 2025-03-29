from flask import Flask, render_template, request, send_file
import pandas as pd
import matplotlib.pyplot as plt
import os
import io
import base64

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
CHART_FOLDER = "static/charts"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(CHART_FOLDER, exist_ok=True)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

class RiskAssessmentApp:
    def __init__(self):
        self.df = None
        self.original_risk_names = None
        self.results = {}

    def get_required_columns(self):
        return [
            "Risk Name",
            "Probability (%)",
            "Impact (Severity)",
            "Detection",
            "Probability",
            "Impact",
            "Cause Likelihood",
            "Consequence Severity",
            "Barrier Effectiveness",
            "Task Affected"
        ]

    def get_result_columns(self, method):
        if method == "fmea":
            return ["Risk Name", "Probability (%)", "Impact (Severity)", "Detection", "Task Affected", "RPN"]
        elif method == "risk matrix":
            return ["Risk Name", "Probability", "Impact", "Task Affected", "Risk Score"]
        elif method == "bow-tie":
            return ["Risk Name", "Cause Likelihood", "Consequence Severity", "Barrier Effectiveness", "Task Affected", "Barrier Score"]

    def calculate_risk(self, method):
        valid_rows = []
        errors = []
        for idx, row in self.df.iterrows():
            try:
                risk_name = str(row["Risk Name"])
                task_affected = str(row["Task Affected"])
                if pd.isna(risk_name) or pd.isna(task_affected):
                    errors.append(f"Row {idx + 2}: Missing or invalid Risk Name or Task Affected ({method.upper()})")
                    continue

                if method == "fmea":
                    prob = float(row["Probability (%)"])
                    impact = float(row["Impact (Severity)"])
                    detection = float(row["Detection"])
                    if pd.isna(prob) or pd.isna(impact) or pd.isna(detection):
                        errors.append(f"Row {idx + 2}: Missing or invalid data in one of the columns (FMEA)")
                        continue
                    if not (0 <= prob <= 100):
                        errors.append(f"Row {idx + 2}: Probability (%) must be between 0 and 100, got {prob} (FMEA)")
                        continue
                    if not (1 <= impact <= 10):
                        errors.append(f"Row {idx + 2}: Impact (Severity) must be between 1 and 10, got {impact} (FMEA)")
                        continue
                    if not (1 <= detection <= 10):
                        errors.append(f"Row {idx + 2}: Detection must be between 1 and 10, got {detection} (FMEA)")
                        continue
                    score = (prob / 100) * impact * detection
                    row["RPN"] = score
                    valid_rows.append(row)

                elif method == "risk matrix":
                    prob = float(row["Probability"])
                    impact = float(row["Impact"])
                    if pd.isna(prob) or pd.isna(impact):
                        errors.append(f"Row {idx + 2}: Missing or invalid data in one of the columns (Risk Matrix)")
                        continue
                    if not (1 <= prob <= 5):
                        errors.append(f"Row {idx + 2}: Probability must be between 1 and 5, got {prob} (Risk Matrix)")
                        continue
                    if not (1 <= impact <= 5):
                        errors.append(f"Row {idx + 2}: Impact must be between 1 and 5, got {impact} (Risk Matrix)")
                        continue
                    score = prob * impact
                    row["Risk Score"] = score
                    valid_rows.append(row)

                elif method == "bow-tie":
                    cause_likelihood = float(row["Cause Likelihood"])
                    consequence_severity = float(row["Consequence Severity"])
                    barrier_effectiveness = float(row["Barrier Effectiveness"])
                    if pd.isna(cause_likelihood) or pd.isna(consequence_severity) or pd.isna(barrier_effectiveness):
                        errors.append(f"Row {idx + 2}: Missing or invalid data in one of the columns (Bow-Tie)")
                        continue
                    if not (1 <= cause_likelihood <= 5):
                        errors.append(f"Row {idx + 2}: Cause Likelihood must be between 1 and 5, got {cause_likelihood} (Bow-Tie)")
                        continue
                    if not (1 <= consequence_severity <= 5):
                        errors.append(f"Row {idx + 2}: Consequence Severity must be between 1 and 5, got {consequence_severity} (Bow-Tie)")
                        continue
                    if not (1 <= barrier_effectiveness <= 5):
                        errors.append(f"Row {idx + 2}: Barrier Effectiveness must be between 1 and 5, got {barrier_effectiveness} (Bow-Tie)")
                        continue
                    score = (cause_likelihood * consequence_severity) / barrier_effectiveness
                    row["Barrier Score"] = score
                    valid_rows.append(row)

            except (ValueError, TypeError) as e:
                errors.append(f"Row {idx + 2}: Invalid data - {str(e)} ({method.upper()})")

        if not valid_rows:
            return None, errors
        df = pd.DataFrame(valid_rows)
        score_column = "RPN" if method == "fmea" else "Risk Score" if method == "risk matrix" else "Barrier Score"
        df = df.sort_values(by=score_column, ascending=False)
        return df, errors

    def normalize_score(self, score, method):
        if method == "fmea":
            return score
        elif method == "risk matrix":
            return ((score - 1) / (25 - 1)) * 100
        elif method == "bow-tie":
            return ((score - 0.2) / (25 - 0.2)) * 100
        return score

    def plot_combined_scores(self):
        if not self.results:
            return None

        plt.figure(figsize=(12, 8))
        risk_names = self.original_risk_names
        y_positions = range(len(risk_names))

        scores = {"fmea": [], "risk matrix": [], "bow-tie": []}
        for risk in risk_names:
            for method in self.results:
                df, _ = self.results[method]
                score_column = "RPN" if method == "fmea" else "Risk Score" if method == "risk matrix" else "Barrier Score"
                score = df[df["Risk Name"] == risk][score_column].values
                normalized_score = self.normalize_score(score[0] if len(score) > 0 else 0, method)
                scores[method].append(normalized_score)

        plt.scatter(scores["fmea"], y_positions, color="red", label="FMEA (RPN)", s=100)
        plt.scatter(scores["risk matrix"], y_positions, color="green", label="Risk Matrix (Risk Score)", s=100)
        plt.scatter(scores["bow-tie"], y_positions, color="blue", label="Bow-Tie (Barrier Score)", s=100)

        for i in range(len(risk_names)):
            plt.plot([scores["fmea"][i], scores["risk matrix"][i]], [i, i], color="gray", linestyle="--", alpha=0.5)
            plt.plot([scores["risk matrix"][i], scores["bow-tie"][i]], [i, i], color="gray", linestyle="--", alpha=0.5)
            plt.plot([scores["bow-tie"][i], scores["fmea"][i]], [i, i], color="gray", linestyle="--", alpha=0.5)

        plt.yticks(y_positions, risk_names)
        plt.xlabel("Normalized Score (0 to 100)")
        plt.ylabel("Risk Name")
        plt.title("Comparison of Risk Scores Across Methods")
        plt.xlim(0, 100)
        plt.legend()
        plt.grid(True, which="both", linestyle="--", alpha=0.7)
        plt.tight_layout()

        buf = io.BytesIO()
        plt.savefig(buf, format="png")
        buf.seek(0)
        plt.close()
        return buf

risk_app = RiskAssessmentApp()

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        if "file" not in request.files or not request.form.get("method"):
            return render_template("index.html", error="Please upload a file and select a method.")

        file = request.files["file"]
        method = request.form.get("method").lower()
        valid_methods = ["fmea", "risk matrix", "bow-tie"]

        if method not in valid_methods:
            return render_template("index.html", error="Invalid method! Please select one of: FMEA, Risk Matrix, or Bow-Tie.")

        if file.filename == "":
            return render_template("index.html", error="No file selected.")

        if file and file.filename.endswith((".xlsx", ".xls")):
            file_path = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
            file.save(file_path)

            try:
                risk_app.df = pd.read_excel(file_path)
                actual_columns = risk_app.df.columns.tolist()
                required_columns = risk_app.get_required_columns()
                if not all(col in actual_columns for col in required_columns):
                    missing_cols = [col for col in required_columns if col not in actual_columns]
                    return render_template("index.html", error=f"Excel file must contain the following columns: {', '.join(required_columns)}\nMissing columns: {', '.join(missing_cols)}")
                else:
                    risk_app.original_risk_names = risk_app.df["Risk Name"].tolist()

                    risk_app.results["fmea"] = risk_app.calculate_risk("fmea")
                    risk_app.results["risk matrix"] = risk_app.calculate_risk("risk matrix")
                    risk_app.results["bow-tie"] = risk_app.calculate_risk("bow-tie")

                    all_errors = []
                    for m in risk_app.results:
                        df, errors = risk_app.results[m]
                        if df is None:
                            return render_template("index.html", error=f"No valid data to process for {m.upper()}. Please check your Excel file.\n" + "\n".join(errors))
                        all_errors.extend(errors)

                    selected_df = risk_app.results[method][0]
                    other_methods = [m for m in ["fmea", "risk matrix", "bow-tie"] if m != method]
                    other_df1 = risk_app.results[other_methods[0]][0].head(3)
                    other_df2 = risk_app.results[other_methods[1]][0].head(3)

                    selected_html = selected_df.to_html(index=False, classes="table table-striped")
                    other_html1 = other_df1.to_html(index=False, classes="table table-striped")
                    other_html2 = other_df2.to_html(index=False, classes="table table-striped")

                    with open("risk_assessment_results.txt", "w", encoding="utf-8") as f:
                        for m in risk_app.results:
                            df, _ = risk_app.results[m]
                            if m == method:
                                f.write(f"{m.upper()} Results (All Risks):\n")
                                f.write(df.to_string(index=False))
                            else:
                                f.write(f"\n\n{m.upper()} Top 3 Risks:\n")
                                f.write(df.head(3).to_string(index=False))
                            f.write("\n")

                    combined_plot = risk_app.plot_combined_scores()
                    if combined_plot:
                        combined_plot.seek(0)
                        plot_data = base64.b64encode(combined_plot.read()).decode("utf-8")
                        combined_plot.close()
                    else:
                        plot_data = None

                    return render_template(
                        "results.html",
                        selected_table=selected_html,
                        other_table1=other_html1,
                        other_table2=other_html2,
                        method=method.upper(),
                        other_method1=other_methods[0].upper(),
                        other_method2=other_methods[1].upper(),
                        errors=all_errors if all_errors else None,
                        combined_plot=plot_data
                    )

            except Exception as e:
                return render_template("index.html", error=f"Failed to load file: {str(e)}")

        return render_template("index.html", error="Invalid file format. Please upload an Excel file (.xlsx or .xls).")

    return render_template("index.html")

@app.route("/download_results")
def download_results():
    return send_file("risk_assessment_results.txt", as_attachment=True)

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=True)