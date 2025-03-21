import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os

def process_file(file_path):
    target_temps = [32, 250, 275]
    temp_tolerance = 0.3
    lookback_rows = 12

    df = pd.read_csv(file_path)
    sensor_columns = df.columns[2:]

    for col in sensor_columns:
        df[col] = pd.to_numeric(df[col], errors='coerce')

    def find_last_occurrence_auto(sensor_data, target):
        threshold = 0.4
        jump_threshold = 2
        candidates = []

        for i in range(len(sensor_data) - 1):
            temp_now = sensor_data.iloc[i]
            temp_next = sensor_data.iloc[i + 1]

            if pd.isna(temp_now) or pd.isna(temp_next):
                continue

            if abs(temp_now - target) < threshold:
                if target < 100:
                    if (temp_next - temp_now) > jump_threshold:
                        candidates.append(i)
                else:
                    if (temp_now - temp_next) > jump_threshold:
                        candidates.append(i)

        if not candidates:
            closest_idx = (sensor_data - target).abs().idxmin()
            return closest_idx

        return candidates[-1]

    results = []

    for col in sensor_columns:
        sensor_name = col.strip()
        avg_results = {'Sensor': sensor_name}

        for target in target_temps:
            sensor_data = df[col]
            idx = find_last_occurrence_auto(sensor_data, target)

            if idx is not None and idx - lookback_rows >= 0:
                window = sensor_data.iloc[idx - lookback_rows + 1:idx + 1]
                window_clean = window.dropna()

                if len(window_clean) >= 3:
                    window_sorted = window_clean.sort_values()
                    trimmed = window_sorted.iloc[1:-1]
                    avg_temp = trimmed.mean()
                    avg_results[f'Avg_{target}'] = round(avg_temp, 3)
                else:
                    avg_results[f'Avg_{target}'] = 'Not found'
            else:
                avg_results[f'Avg_{target}'] = 'Not found'

        results.append(avg_results)

    output_df = pd.DataFrame(results)
    output_csv = os.path.join(os.path.dirname(file_path), 'ellab_output_averages.csv')
    output_xlsx = os.path.join(os.path.dirname(file_path), 'ellab_output_averages.xlsx')

    output_df.to_csv(output_csv, index=False)
    output_df.to_excel(output_xlsx, index=False)

    return output_df


def browse_file():
    file_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
    if file_path:
        entry_file.delete(0, tk.END)
        entry_file.insert(0, file_path)


def run_processing():
    file_path = entry_file.get()
    if not file_path or not os.path.isfile(file_path):
        messagebox.showerror("Error", "Please select a valid CSV file.")
        return

    try:
        df_result = process_file(file_path)
        show_results(df_result)
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")


def show_results(df_result):
    for widget in frame_results.winfo_children():
        widget.destroy()

    tree = ttk.Treeview(frame_results)
    tree.pack(fill=tk.BOTH, expand=True)

    tree["columns"] = list(df_result.columns)
    tree["show"] = "headings"

    for col in df_result.columns:
        tree.heading(col, text=col)
        tree.column(col, width=100, anchor="center")

    for _, row in df_result.iterrows():
        tree.insert("", tk.END, values=list(row))


root = tk.Tk()
root.title("Ellab RTD Data Processor")
root.geometry("800x600")

frame_top = tk.Frame(root)
frame_top.pack(pady=10, padx=10, fill=tk.X)

entry_file = tk.Entry(frame_top, width=60)
entry_file.pack(side=tk.LEFT, padx=5)

btn_browse = tk.Button(frame_top, text="Browse", command=browse_file)
btn_browse.pack(side=tk.LEFT, padx=5)

btn_run = tk.Button(frame_top, text="Run", command=run_processing)
btn_run.pack(side=tk.LEFT, padx=5)

frame_results = tk.Frame(root)
frame_results.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)

root.mainloop()