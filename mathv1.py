import pandas as pd
import numpy as np
from flask import Flask, request, render_template, send_file
import matplotlib.pyplot as plt
import os

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        # Load uploaded files
        reference_file = request.files["reference"]
        base_file = request.files["base"]
        read_file = request.files["read"]

        # Detect delimiters automatically
        reference_df = pd.read_csv(reference_file, sep=None, engine="python")
        base_df = pd.read_csv(base_file, sep=None, engine="python")
        read_df = pd.read_csv(read_file, sep=None, engine="python")

        # Extract 5th column (index 4)
        reference_col = reference_df.iloc[:, 4]
        base_col = base_df.iloc[:, 4]
        read_col = read_df.iloc[:, 4]

        # Step 1: Average of REFERENCE file
        REF = reference_col.mean()

        # Step 2: RB = READ - BASE (row by row)
        RB = read_col.values - base_col.values

        # Step 3: Final calc = ((RB - REF) / (0.0045 * REF))
        final_values = (RB - REF) / (0.0045 * REF)

        # Step 4: Rolling average with window 256
        rolling_avg = np.convolve(final_values, np.ones(256)/256, mode="valid")

        # Step 5: Use 4th column from READ/MEASUREMENT as x-axis
        x_values = read_df.iloc[:, 3].values[:len(rolling_avg)]

        # Step 6: Calculate statistics
        stats = {
            "Maximum": np.max(rolling_avg),
            "Minimum": np.min(rolling_avg),
            "Range": np.max(rolling_avg) - np.min(rolling_avg),
            "Average": np.mean(rolling_avg),
            "Standard Deviation": np.std(rolling_avg),
            "Q1 (25%)": np.percentile(rolling_avg, 25),
            "Median (50%)": np.percentile(rolling_avg, 50),
            "Q3 (75%)": np.percentile(rolling_avg, 75)
        }

        # Step 7: FFT (frequency domain)
        N = len(rolling_avg)
        fft_values = np.fft.fft(rolling_avg)
        fft_freqs = np.fft.fftfreq(N, d=1)  # assumes sample spacing = 1 unit
        fft_magnitude = np.abs(fft_values)[:N // 2]
        fft_freqs = fft_freqs[:N // 2]

        # Save main graph
        plt.figure(figsize=(8, 5))
        plt.plot(x_values, rolling_avg, label="Processed Data")
        plt.xlabel("4th Column of READ/MEASUREMENT")
        plt.ylabel("Final RB-REF Calculation")
        plt.title("Processed Graph")
        plt.legend()
        plt.grid(True)
        graph_path = os.path.join(UPLOAD_FOLDER, "graph.png")
        plt.savefig(graph_path, format="png")
        plt.close()

        # Save FFT graph
        plt.figure(figsize=(8, 5))
        plt.plot(fft_freqs, fft_magnitude, label="FFT Spectrum")
        plt.xlabel("Frequency (a.u.)")
        plt.ylabel("Magnitude")
        plt.title("FFT of Processed Data")
        plt.legend()
        plt.grid(True)
        fft_graph_path = os.path.join(UPLOAD_FOLDER, "fft_graph.png")
        plt.savefig(fft_graph_path, format="png")
        plt.close()

        # Save CSV (include data, stats, FFT)
        output_df = pd.DataFrame({
            "X (4th Column of READ/MEASUREMENT)": x_values,
            "Processed Values": rolling_avg
        })
        stats_df = pd.DataFrame(list(stats.items()), columns=["Statistic", "Value"])
        fft_df = pd.DataFrame({
            "Frequency": fft_freqs,
            "FFT Magnitude": fft_magnitude
        })

        csv_path = os.path.join(UPLOAD_FOLDER, "processed.csv")
        with open(csv_path, "w") as f:
            output_df.to_csv(f, index=False)
            f.write("\n\nStatistics\n")
            stats_df.to_csv(f, index=False)
            f.write("\n\nFFT Data\n")
            fft_df.to_csv(f, index=False)

        return render_template("index.html",
                               graph="graph.png",
                               fft_graph="fft_graph.png",
                               csv_file="processed.csv",
                               ref_value=REF,
                               stats=stats)

    return render_template("index.html")


@app.route("/uploads/<filename>")
def uploaded_file(filename):
    return send_file(os.path.join(UPLOAD_FOLDER, filename), as_attachment=True)


if __name__ == "__main__":
    app.run(debug=True)
