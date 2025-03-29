import pandas as pd
import matplotlib.pyplot as plt
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os

class MethodSelectionWindow:
    def __init__(self, root):
        self.root = root
        self.root.title("Select Risk Assessment Method")
        self.root.geometry("400x200")

        self.label = tk.Label(root, text="Enter the risk assessment method you want to use:\n(FMEA, Risk Matrix, or Bow-Tie)")
        self.label.pack(pady=10)

        self.method_entry = tk.Entry(root)
        self.method_entry.pack(pady=5)

        self.confirm_button = tk.Button(root, text="Confirm", command=self.confirm_method)
        self.confirm_button.pack(pady=10)

    def confirm_method(self):
        method = self.method_entry.get().strip().lower()
        valid_methods = ["fmea", "risk matrix", "bow-tie"]
        if method not in valid_methods:
            messagebox.showerror("Error", "Invalid method! Please enter one of: FMEA, Risk Matrix, or Bow-Tie")
            return
        self.root.destroy()
        self.open_main_window(method)

    def open_main_window(self, method):
        main_root = tk.Tk()
        app = RiskAssessmentApp(main_root, method)
        main_root.mainloop()

class RiskAssessmentApp:
    def __init__(self, root, method):
        self.root = root
        self.method = method.lower()
        self.root.title(f"{self.method.upper()} Risk Assessment")
        self.root.geometry("800x800")

        self.show_method_instructions()

        self.label = tk.Label(root, text="Select an Excel file for risk assessment")
        self.label.pack(pady=10)

        self.select_button = tk.Button(root, text="Select Excel File", command=self.load_file)
        self.select_button.pack(pady=5)

        self.run_button = tk.Button(root, text="Run Analysis", command=self.run_analysis, state=tk.DISABLED)
        self.run_button.pack(pady=5)

        self.tree_frame = tk.Frame(root)
        self.tree_frame.pack(pady=10, fill=tk.BOTH, expand=True)

        self.tree_scroll_y = tk.Scrollbar(self.tree_frame, orient=tk.VERTICAL)
        self.tree_scroll_y.pack(side=tk.RIGHT, fill=tk.Y)

        self.tree_scroll_x = tk.Scrollbar(self.tree_frame, orient=tk.HORIZONTAL)
        self.tree_scroll_x.pack(side=tk.BOTTOM, fill=tk.X)

        self.tree = ttk.Treeview(
            self.tree_frame,
            yscrollcommand=self.tree_scroll_y.set,
            xscrollcommand=self.tree_scroll_x.set,
            height=10
        )
        self.tree.pack(fill=tk.BOTH, expand=True)

        self.tree_scroll_y.config(command=self.tree.yview)
        self.tree_scroll_x.config(command=self.tree.xview)

        self.other_methods = ["fmea", "risk matrix", "bow-tie"]
        self.other_methods.remove(self.method)

        self.top_risks_label1 = tk.Label(root, text=f"Top 3 Risks ({self.other_methods[0].upper()}):", font=("Arial", 12, "bold"))
        self.top_risks_label1.pack(pady=5)

        self.top_risks_frame1 = tk.Frame(root)
        self.top_risks_frame1.pack(fill=tk.BOTH, expand=False)

        self.top_risks_tree1 = ttk.Treeview(
            self.top_risks_frame1,
            height=3
        )
        self.top_risks_tree1.pack(fill=tk.BOTH, expand=False)

        self.top_risks_label2 = tk.Label(root, text=f"Top 3 Risks ({self.other_methods[1].upper()}):", font=("Arial", 12, "bold"))
        self.top_risks_label2.pack(pady=5)

        self.top_risks_frame2 = tk.Frame(root)
        self.top_risks_frame2.pack(fill=tk.BOTH, expand=False)

        self.top_risks_tree2 = ttk.Treeview(
            self.top_risks_frame2,
            height=3
        )
        self.top_risks_tree2.pack(fill=tk.BOTH, expand=False)

        self.status_label = tk.Label(root, text="", wraplength=700)
        self.status_label.pack(pady=5)

        self.df = None
        self.original_risk_names = None
        self.required_columns = self.get_required_columns()
        self.results = {}

    def show_method_instructions(self):
        messagebox.showinfo(
            "Instructions",
            "Your Excel file must have the following columns for all methods:\n"
            "- Risk Name (string)\n"
            "- Probability (%) (for FMEA, number between 0 and 100)\n"
            "- Impact (Severity) (for FMEA, number between 1 and 10)\n"
            "- Detection (for FMEA, number between 1 and 10)\n"
            "- Probability (for Risk Matrix, number between 1 and 5)\n"
            "- Impact (for Risk Matrix, number between 1 and 5)\n"
            "- Cause Likelihood (for Bow-Tie, number between 1 and 5)\n"
            "- Consequence Severity (for Bow-Tie, number between 1 and 5)\n"
            "- Barrier Effectiveness (for Bow-Tie, number between 1 and 5)\n"
            "- Task Affected (string)\n"
            "Example: A, 5, 5, 1, 3, 4, 3, 4, 2, AA"
        )

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

    def load_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if file_path:
            try:
                self.df = pd.read_excel(file_path)
                actual_columns = self.df.columns.tolist()
                if not all(col in actual_columns for col in self.required_columns):
                    missing_cols = [col for col in self.required_columns if col not in actual_columns]
                    messagebox.showerror("Error", f"Excel file must contain columns: {', '.join(self.required_columns)}\nMissing columns: {', '.join(missing_cols)}")
                    self.df = None
                    self.run_button.config(state=tk.DISABLED)
                else:
                    self.original_risk_names = self.df["Risk Name"].tolist()  # Store original order of risk names
                    self.status_label.config(text=f"Loaded file: {os.path.basename(file_path)}\nClick 'Run Analysis' to proceed.")
                    self.run_button.config(state=tk.NORMAL)
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load file: {str(e)}")
                self.df = None
                self.run_button.config(state=tk.DISABLED)

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

    def run_analysis(self):
        if self.df is None:
            messagebox.showerror("Error", "No data loaded. Please select an Excel file.")
            return

        self.results["fmea"] = self.calculate_risk("fmea")
        self.results["risk matrix"] = self.calculate_risk("risk matrix")
        self.results["bow-tie"] = self.calculate_risk("bow-tie")

        self.display_results()

    def normalize_score(self, score, method):
        if method == "fmea":
            return score  # Already in range 0 to 100
        elif method == "risk matrix":
            # Risk Score range: 1 to 25, map to 0 to 100
            return ((score - 1) / (25 - 1)) * 100
        elif method == "bow-tie":
            # Barrier Score range: 0.2 to 25, map to 0 to 100
            return ((score - 0.2) / (25 - 0.2)) * 100
        return score

    def plot_combined_scores(self):
        if not self.results:
            return

        plt.figure(figsize=(12, 8))
        risk_names = self.original_risk_names  # Use original order from Excel
        y_positions = range(len(risk_names))

        # Collect scores for each method
        scores = {"fmea": [], "risk matrix": [], "bow-tie": []}
        for risk in risk_names:
            for method in self.results:
                df, _ = self.results[method]
                score_column = "RPN" if method == "fmea" else "Risk Score" if method == "risk matrix" else "Barrier Score"
                score = df[df["Risk Name"] == risk][score_column].values
                normalized_score = self.normalize_score(score[0] if len(score) > 0 else 0, method)
                scores[method].append(normalized_score)

        # Plot scores for each method with different colors
        plt.scatter(scores["fmea"], y_positions, color="red", label="FMEA (RPN)", s=100)
        plt.scatter(scores["risk matrix"], y_positions, color="green", label="Risk Matrix (Risk Score)", s=100)
        plt.scatter(scores["bow-tie"], y_positions, color="blue", label="Bow-Tie (Barrier Score)", s=100)

        # Draw lines connecting scores for each risk
        for i in range(len(risk_names)):
            # FMEA to Risk Matrix
            plt.plot([scores["fmea"][i], scores["risk matrix"][i]], [i, i], color="gray", linestyle="--", alpha=0.5)
            # Risk Matrix to Bow-Tie
            plt.plot([scores["risk matrix"][i], scores["bow-tie"][i]], [i, i], color="gray", linestyle="--", alpha=0.5)
            # Bow-Tie to FMEA (optional, to close the loop)
            plt.plot([scores["bow-tie"][i], scores["fmea"][i]], [i, i], color="gray", linestyle="--", alpha=0.5)

        # Set y-axis labels (risk names)
        plt.yticks(y_positions, risk_names)
        plt.xlabel("Normalized Score (0 to 100)")
        plt.ylabel("Risk Name")
        plt.title("Comparison of Risk Scores Across Methods")
        plt.xlim(0, 100)  # Set x-axis range to 0 to 100
        plt.legend()
        plt.grid(True, which="both", linestyle="--", alpha=0.7)
        plt.tight_layout()
        plt.savefig("combined_risk_scores.png")
        plt.close()

    def display_results(self):
        all_errors = []
        for method in self.results:
            df, errors = self.results[method]
            if df is None:
                messagebox.showerror("Error", f"No valid data to process for {method.upper()}. Check your Excel file.\n" + "\n".join(errors))
                return
            all_errors.extend(errors)

        selected_df = self.results[self.method][0]
        self.tree.delete(*self.tree.get_children())
        self.tree["columns"] = self.get_result_columns(self.method)
        self.tree.heading("#0", text="")
        for col in self.tree["columns"]:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=120, anchor="center")

        for idx, row in selected_df.iterrows():
            values = [row[col] for col in self.get_result_columns(self.method)[:-1]]
            score_column = self.get_result_columns(self.method)[-1]
            values.append(f"{row[score_column]:.2f}")
            self.tree.insert("", tk.END, values=values)

        other_df1 = self.results[self.other_methods[0]][0].head(3)
        self.top_risks_tree1.delete(*self.top_risks_tree1.get_children())
        self.top_risks_tree1["columns"] = self.get_result_columns(self.other_methods[0])
        self.top_risks_tree1.heading("#0", text="")
        for col in self.top_risks_tree1["columns"]:
            self.top_risks_tree1.heading(col, text=col)
            self.top_risks_tree1.column(col, width=120, anchor="center")

        for idx, row in other_df1.iterrows():
            values = [row[col] for col in self.get_result_columns(self.other_methods[0])[:-1]]
            score_column = self.get_result_columns(self.other_methods[0])[-1]
            values.append(f"{row[score_column]:.2f}")
            self.top_risks_tree1.insert("", tk.END, values=values)

        other_df2 = self.results[self.other_methods[1]][0].head(3)
        self.top_risks_tree2.delete(*self.top_risks_tree2.get_children())
        self.top_risks_tree2["columns"] = self.get_result_columns(self.other_methods[1])
        self.top_risks_tree2.heading("#0", text="")
        for col in self.top_risks_tree2["columns"]:
            self.top_risks_tree2.heading(col, text=col)
            self.top_risks_tree2.column(col, width=120, anchor="center")

        for idx, row in other_df2.iterrows():
            values = [row[col] for col in self.get_result_columns(self.other_methods[1])[:-1]]
            score_column = self.get_result_columns(self.other_methods[1])[-1]
            values.append(f"{row[score_column]:.2f}")
            self.top_risks_tree2.insert("", tk.END, values=values)

        with open("risk_assessment_results.txt", "w", encoding="utf-8") as f:
            for method in self.results:
                df, _ = self.results[method]
                if method == self.method:
                    f.write(f"{method.upper()} Results (All Risks):\n")
                    f.write(df.to_string(index=False))
                else:
                    f.write(f"\n\n{method.upper()} Top 3 Risks:\n")
                    f.write(df.head(3).to_string(index=False))
                f.write("\n")

        if all_errors:
            self.status_label.config(text="Errors encountered:\n" + "\n".join(all_errors))
            with open("risk_assessment_results.txt", "a", encoding="utf-8") as f:
                f.write("\n\nErrors encountered:\n")
                f.write("\n".join(all_errors))
                f.write("\n")
        else:
            self.status_label.config(text="")

        for method in self.results:
            df, _ = self.results[method]
            if method == "fmea":
                x_axis, y_axis, score_column = "Probability (%)", "Impact (Severity)", "RPN"
            elif method == "risk matrix":
                x_axis, y_axis, score_column = "Probability", "Impact", "Risk Score"
            elif method == "bow-tie":
                x_axis, y_axis, score_column = "Cause Likelihood", "Consequence Severity", "Barrier Score"

            plt.figure(figsize=(10, 6))
            plt.scatter(df[x_axis], df[y_axis], s=100, c="blue", alpha=0.5)
            for i, txt in enumerate(df["Risk Name"]):
                plt.annotate(txt, (df[x_axis].iloc[i], df[y_axis].iloc[i]))
            plt.xlabel(x_axis)
            plt.ylabel(y_axis)
            plt.title(f"{method.upper()} Risk Matrix")
            plt.grid(True)
            plt.savefig(f"risk_matrix_{method}.png")
            plt.close()

            plt.figure(figsize=(10, 6))
            plt.bar(df["Risk Name"], df[score_column], color="orange")
            plt.xlabel("Risk Name")
            plt.ylabel(score_column)
            plt.title(f"{score_column} of Risks (Sorted) - {method.upper()}")
            plt.xticks(rotation=45, ha="right")
            plt.tight_layout()
            plt.savefig(f"score_bar_{method}.png")
            plt.close()

        # Plot the combined scores chart
        self.plot_combined_scores()

        self.status_label.config(text="Charts saved as 'risk_matrix_[method].png', 'score_bar_[method].png', and 'combined_risk_scores.png'\nResults also saved to 'risk_assessment_results.txt'")

root = tk.Tk()
app = MethodSelectionWindow(root)
root.mainloop()