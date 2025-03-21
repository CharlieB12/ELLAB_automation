import pandas as pd

# Load CSV (change path to your file)
file_path = 'pro02252025.csv'
df = pd.read_csv(file_path)

# Parameters
target_temps = [32, 250, 275]
temp_tolerance = 0.3
lookback_rows = 12

results = []

# Determine which columns are sensors
sensor_columns = df.columns[2:]

# Convert sensor columns to numeric
for col in sensor_columns:
    df[col] = pd.to_numeric(df[col], errors='coerce')

def find_last_occurrence_auto(sensor_data, target):
    threshold = 0.4
    jump_threshold = 2  # temperature jump difference to detect pulling out

    candidates = []
    for i in range(len(sensor_data)-1):
        temp_now = sensor_data.iloc[i]
        temp_next = sensor_data.iloc[i+1]

        if pd.isna(temp_now) or pd.isna(temp_next):
            continue

        if abs(temp_now - target) < threshold:
            if target < 100:  # low temp, look for rise
                if (temp_next - temp_now) > jump_threshold:
                    candidates.append(i)
            else:  # high temp, look for drop
                if (temp_now - temp_next) > jump_threshold:
                    candidates.append(i)

    # If no jump detected, fallback to closest occurrence
    if not candidates:
        closest_idx = (sensor_data - target).abs().idxmin()
        return closest_idx

    return candidates[-1]

for col in sensor_columns:
    sensor_name = col.strip()
    avg_results = {'Sensor': sensor_name}

    for target in target_temps:
        sensor_data = df[col]
        idx = find_last_occurrence_auto(sensor_data, target)

        if idx is not None and idx - lookback_rows >= 0:
            window = sensor_data.iloc[idx - lookback_rows + 1:idx + 1]  # Use correct indexing
            window_clean = window.dropna()

            print(f"Sensor: {sensor_name}, Target: {target}, Window values:")
            print(window_clean.values)

            if len(window_clean) >= 3:
                window_sorted = window_clean.sort_values()
                trimmed = window_sorted.iloc[1:-1]  # Remove only min and max by sorting
                avg_temp = trimmed.mean()
                avg_results[f'Avg_{target}'] = round(avg_temp, 3)
            else:
                avg_results[f'Avg_{target}'] = 'Not found'
        else:
            avg_results[f'Avg_{target}'] = 'Not found'

    results.append(avg_results)

output_df = pd.DataFrame(results)
output_df.to_csv('ellab_output_averages.csv', index=False)
output_df.to_excel('ellab_output_averages.xlsx', index=False)
print(output_df)