#!/usr/bin/env python3
"""
combined_app.py

Merges logic from app.py (Instrument Control) and app2.py (Math/Processing).

Workflow:
1. Connect Devices (Scope, Koolertron, Keithley).
2. Load a "Reference" CSV (Cached).
3. Run "Full Measurement Cycle":
    a. Measure BASE (Keithley OFF).
    b. Measure READ (Keithley ON).
    c. Perform App2 Math: (Read - Base - RefAvg) / (0.0045 * RefAvg).
    d. Display Graphs (Time Domain + FFT) in a popup.
    e. Save Data & Images to disk.
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime
import pyvisa
import numpy as np
import pandas as pd
import csv
import time
import os
import threading
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

# Try to import the driver; strict dependency
try:
    from koolertron import KoolertronSig
except ImportError:
    # Dummy class for testing if driver is missing, prevents immediate crash
    print("WARNING: 'koolertron.py' not found. Using dummy driver for UI testing.")
    class KoolertronSig:
        def __init__(self, port): self.connected = True
        def isConnected(self): return True
        def sinwave(self, *args, **kwargs): pass
        def squareWave(self, *args, **kwargs): pass
        def setDuty(self, *args, **kwargs): pass
        def setAmplitude(self, *args, **kwargs): pass

# ---------------- Configuration ----------------
NUM_AVERAGES = 10
READ_DELAY_SEC = 0.03
SCOPE_TIMEOUT_MS = 10000
DEFAULT_SCOPE_ADDR = "USB0::0x0699::0x0401::C010101::INSTR"
MEAS_DIR = "measurements"
KEITHLEY_OUTPUT_ON = ":OUTP ON"
KEITHLEY_OUTPUT_OFF = ":OUTP OFF"
# ------------------------------------------------

os.makedirs(MEAS_DIR, exist_ok=True)


class InstrumentApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Combined Measurement App (DPO + Kool + Keithley)")
        self.root.geometry("700x750")

        # Hardware Objects
        self.rm = None
        self.scope = None
        self.kool = None
        self.keith = None

        # Data Cache
        self.reference_data = None  # Stores the loaded reference numpy array
        self.reference_filename = "None"

        # UI Variables
        self.scope_addr = tk.StringVar(value=DEFAULT_SCOPE_ADDR)
        self.kool_port = tk.StringVar(value="COM3")
        self.keith_addr = tk.StringVar(value="")
        self.wave_type = tk.StringVar(value="SIN")
        self.freq = tk.DoubleVar(value=1000.0)
        self.amp = tk.DoubleVar(value=1.0)
        self.offset = tk.DoubleVar(value=0.0)
        self.duty_percent = tk.DoubleVar(value=50.0)
        self.channel = tk.StringVar(value="MATH")
        self.status_text = tk.StringVar(value="Not connected")
        self.ref_status = tk.StringVar(value="No Reference Loaded")

        self._build_gui()

    def _build_gui(self):
        pad = dict(padx=8, pady=6)
        frm = ttk.Frame(self.root, padding=12)
        frm.pack(fill="both", expand=True)

        # Title
        ttk.Label(frm, text="Instrument Panel & Analysis", font=("Segoe UI", 16, "bold")).grid(row=0, column=0, columnspan=3, **pad)

        # --- Connections ---
        r = 1
        ttk.Label(frm, text="Scope VISA Address:").grid(row=r, column=0, sticky="w", **pad)
        ttk.Entry(frm, textvariable=self.scope_addr, width=40).grid(row=r, column=1, columnspan=2, sticky="ew", **pad)
        r += 1

        ttk.Label(frm, text="Koolertron COM Port:").grid(row=r, column=0, sticky="w", **pad)
        ttk.Entry(frm, textvariable=self.kool_port, width=20).grid(row=r, column=1, sticky="w", **pad)
        r += 1

        ttk.Label(frm, text="Keithley VISA Address:").grid(row=r, column=0, sticky="w", **pad)
        ttk.Entry(frm, textvariable=self.keith_addr, width=40).grid(row=r, column=1, columnspan=2, sticky="ew", **pad)
        r += 1

        ttk.Button(frm, text="Connect Hardware", command=self.connect_devices).grid(row=r, column=0, columnspan=3, pady=(5, 15))
        r += 1

        ttk.Separator(frm).grid(row=r, column=0, columnspan=3, sticky="ew", pady=5)
        r += 1

        # --- Settings ---
        ttk.Label(frm, text="Wave Type:").grid(row=r, column=0, sticky="w", **pad)
        ttk.Combobox(frm, textvariable=self.wave_type, values=["SIN", "PULSE", "SQUARE"], state="readonly", width=15).grid(row=r, column=1, sticky="w", **pad)
        r += 1

        ttk.Label(frm, text="Frequency (Hz):").grid(row=r, column=0, sticky="w", **pad)
        ttk.Entry(frm, textvariable=self.freq, width=15).grid(row=r, column=1, sticky="w", **pad)
        r += 1

        ttk.Label(frm, text="Amplitude (Vpp):").grid(row=r, column=0, sticky="w", **pad)
        ttk.Entry(frm, textvariable=self.amp, width=15).grid(row=r, column=1, sticky="w", **pad)
        r += 1

        ttk.Label(frm, text="Offset (V):").grid(row=r, column=0, sticky="w", **pad)
        ttk.Entry(frm, textvariable=self.offset, width=15).grid(row=r, column=1, sticky="w", **pad)
        r += 1

        ttk.Label(frm, text="Duty Cycle (%):").grid(row=r, column=0, sticky="w", **pad)
        ttk.Entry(frm, textvariable=self.duty_percent, width=15).grid(row=r, column=1, sticky="w", **pad)
        r += 1

        ttk.Separator(frm).grid(row=r, column=0, columnspan=3, sticky="ew", pady=5)
        r += 1

        # --- Reference File Section ---
        ttk.Label(frm, text="Reference File:", font=("Segoe UI", 10, "bold")).grid(row=r, column=0, sticky="w", **pad)
        ttk.Label(frm, textvariable=self.ref_status, foreground="green").grid(row=r, column=1, sticky="w", **pad)
        ttk.Button(frm, text="Load Reference CSV", command=self.load_reference).grid(row=r, column=2, sticky="w", **pad)
        r += 1

        ttk.Separator(frm).grid(row=r, column=0, columnspan=3, sticky="ew", pady=5)
        r += 1

        # --- Actions ---
        btn_f = ttk.Frame(frm)
        btn_f.grid(row=r, column=0, columnspan=3, pady=10)
        
        # New Main Action Button
        self.btn_run = ttk.Button(btn_f, text="RUN FULL CYCLE\n(Base -> Read -> Process)", command=self.start_full_cycle_thread)
        self.btn_run.pack(side="left", padx=20, ipady=10)

        # Legacy individual buttons (Optional, kept for debugging)
        # ttk.Button(btn_f, text="Step 2 Only", command=lambda: self._start_thread(self.step2_action)).pack(side="left", padx=5)
        # ttk.Button(btn_f, text="Step 3 Only", command=lambda: self._start_thread(self.step3_action)).pack(side="left", padx=5)
        r += 1

        # --- Status & Log ---
        ttk.Label(frm, text="Status:").grid(row=r, column=0, sticky="w", **pad)
        ttk.Label(frm, textvariable=self.status_text, foreground="blue").grid(row=r, column=1, columnspan=2, sticky="w", **pad)
        r += 1

        self.logbox = tk.Text(frm, height=12, width=80, state="disabled")
        self.logbox.grid(row=r, column=0, columnspan=3, **pad)

    # ---------------- Logging ----------------
    def log(self, msg):
        ts = datetime.now().strftime("%H:%M:%S")
        line = f"[{ts}] {msg}\n"
        self.logbox.configure(state="normal")
        self.logbox.insert("end", line)
        self.logbox.see("end")
        self.logbox.configure(state="disabled")

    # ---------------- Hardware Connection ----------------
    def connect_devices(self):
        try:
            if not self.rm:
                self.rm = pyvisa.ResourceManager()
            
            # Scope
            if not self.scope:
                addr = self.scope_addr.get().strip()
                if not addr: raise ValueError("Scope address empty")
                self.scope = self.rm.open_resource(addr)
                self.scope.timeout = SCOPE_TIMEOUT_MS
                self.log(f"Scope: {self.scope.query('*IDN?').strip()}")

            # Koolertron
            if not self.kool:
                port = self.kool_port.get().strip()
                self.kool = KoolertronSig(port)
                if not self.kool.isConnected():
                    raise RuntimeError(f"Koolertron not responding on {port}")
                self.log(f"Koolertron connected on {port}")

            # Keithley
            kaddr = self.keith_addr.get().strip()
            if kaddr and not self.keith:
                try:
                    self.keith = self.rm.open_resource(kaddr)
                    self.keith.timeout = 5000
                    self.log(f"Keithley: {self.keith.query('*IDN?').strip()}")
                except Exception as e:
                    self.log(f"Keithley connection warning: {e}")

            self.status_text.set("Hardware Connected")
            messagebox.showinfo("Success", "Hardware Connected")
        except Exception as e:
            self.log(f"Connection Error: {e}")
            messagebox.showerror("Error", str(e))

    # ---------------- Reference File Loading ----------------
    def load_reference(self):
        path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv"), ("All Files", "*.*")])
        if not path:
            return

        try:
            # Logic adapted from app2.py
            df = pd.read_csv(path, sep=None, engine="python")
            
            # Heuristic: app2.py expects data in column index 4 (5th column)
            # app.py saves data in column index 1 (2nd column: "Voltage (V)")
            # We check shape to decide.
            if df.shape[1] >= 5:
                # Likely an app2-style reference file
                self.reference_data = df.iloc[:, 4].values
                self.log(f"Loaded Reference (Column 5): {len(self.reference_data)} points")
            elif df.shape[1] >= 2:
                # Likely an app.py raw save
                self.reference_data = df.iloc[:, 1].values
                self.log(f"Loaded Reference (Column 2): {len(self.reference_data)} points")
            else:
                # Fallback
                self.reference_data = df.iloc[:, 0].values
                self.log(f"Loaded Reference (Column 1): {len(self.reference_data)} points")

            self.reference_filename = os.path.basename(path)
            self.ref_status.set(f"Loaded: {self.reference_filename}")
        except Exception as e:
            self.log(f"Error loading reference: {e}")
            messagebox.showerror("Load Error", str(e))

    # ---------------- Threading Helpers ----------------
    def _start_thread(self, target):
        t = threading.Thread(target=target, daemon=True)
        t.start()

    def start_full_cycle_thread(self):
        if self.reference_data is None:
            messagebox.showwarning("Missing Reference", "Please load a Reference CSV file first.")
            return
        if not self.scope or not self.kool:
            messagebox.showwarning("Not Connected", "Connect hardware first.")
            return
        
        self.btn_run.config(state="disabled")
        self._start_thread(self.run_full_cycle_logic)

    # ---------------- CORE LOGIC: Measurement Cycle ----------------
    def run_full_cycle_logic(self):
        """
        1. Setup Generator.
        2. Keithley OFF -> Measure BASE.
        3. Keithley ON -> Measure READ.
        4. Keithley OFF.
        5. Process Data (Math).
        """
        try:
            # 1. Setup Generator
            if not self._configure_koolertron(on=True):
                raise RuntimeError("Failed to configure Koolertron")

            # 2. Measure BASE (Keithley OFF)
            self._keithley_set_output(False)
            time.sleep(0.5) # Settling time
            base_volts, base_time = self._capture_averaged("BASE")
            if base_volts is None: raise RuntimeError("Failed to capture BASE")

            # 3. Measure READ (Keithley ON)
            self._keithley_set_output(True)
            time.sleep(0.5) # Settling time
            read_volts, read_time = self._capture_averaged("READ")
            if read_volts is None: raise RuntimeError("Failed to capture READ")

            # 4. Cleanup
            self._keithley_set_output(False)

            # 5. Handover to Main Thread for Processing/Plotting
            # We must pass the data back to the UI thread because Tkinter hates threads touching UI
            self.root.after(0, lambda: self.process_and_display(base_volts, read_volts, read_time))

        except Exception as e:
            self.log(f"Cycle Error: {e}")
            self.status_text.set("Error during cycle")
        finally:
            self.root.after(0, lambda: self.btn_run.config(state="normal"))

    def _capture_averaged(self, label):
        self.status_text.set(f"Capturing {label} ({NUM_AVERAGES} avg)...")
        all_v = []
        t_axis = None
        
        for i in range(NUM_AVERAGES):
            v, t = self._read_scope_once()
            if v is None: return None, None
            all_v.append(v)
            if t_axis is None: t_axis = t
            time.sleep(READ_DELAY_SEC)

        # Crop to min length
        min_len = min(len(a) for a in all_v)
        all_v = [a[:min_len] for a in all_v]
        t_axis = t_axis[:min_len]

        avg = np.mean(np.vstack(all_v), axis=0)
        self.log(f"{label} captured ({min_len} pts).")
        return avg, t_axis

    # ---------------- Hardware Helpers ----------------
    def _configure_koolertron(self, on):
        try:
            freq = self.freq.get()
            amp = self.amp.get()
            off = self.offset.get()
            duty = self.duty_percent.get() / 100.0
            
            # Apply to both channels
            if self.wave_type.get() == "SIN":
                self.kool.sinwave(freq, amp, 0, off, chan=1)
                self.kool.sinwave(freq, amp, 0, off, chan=2)
            else:
                self.kool.squareWave(freq, amp, 0, off, chan=1)
                self.kool.squareWave(freq, amp, 0, off, chan=2)
            
            self.kool.setDuty(duty, channel=1)
            self.kool.setDuty(duty, channel=2)
            
            if on:
                self.kool.setAmplitude(amp, channel=1)
                self.kool.setAmplitude(amp, channel=2)
            else:
                self.kool.setAmplitude(0, channel=1)
                self.kool.setAmplitude(0, channel=2)
            return True
        except Exception as e:
            self.log(f"Koolertron Config Error: {e}")
            return False

    def _keithley_set_output(self, on):
        if not self.keith: return
        try:
            cmd = KEITHLEY_OUTPUT_ON if on else KEITHLEY_OUTPUT_OFF
            self.keith.write(cmd)
            self.log(f"Keithley: {cmd}")
        except Exception as e:
            self.log(f"Keithley Error: {e}")

    def _read_scope_once(self):
        try:
            src = self.channel.get()
            self.scope.write(f"DATA:SOURCE {src}")
            self.scope.write("DATA:WIDTH 1")
            self.scope.write("DATA:ENC RPB")
            
            xinc = float(self.scope.query("WFMPRE:XINCR?"))
            xorg = float(self.scope.query("WFMPRE:XZERO?"))
            ymult = float(self.scope.query("WFMPRE:YMULT?"))
            yoff = float(self.scope.query("WFMPRE:YOFF?"))
            yzero = float(self.scope.query("WFMPRE:YZERO?"))

            raw = self.scope.query_binary_values("CURVE?", datatype='B', container=np.array)
            volt = (raw - yoff) * ymult + yzero
            t = np.arange(len(volt)) * xinc + xorg
            return volt, t
        except Exception as e:
            self.log(f"Scope Read Error: {e}")
            return None, None

    # ---------------- MATH & PLOTTING (The app2.py logic) ----------------
    def process_and_display(self, base, read, time_axis):
        self.status_text.set("Processing Data...")
        
        # 1. Align Lengths
        # We have Base, Read, and Reference. They must be same length.
        n = min(len(base), len(read), len(self.reference_data))
        
        base_c = base[:n]
        read_c = read[:n]
        ref_c  = self.reference_data[:n]
        time_c = time_axis[:n]

        # 2. Math Logic from app2.py
        # Step 1: Average of REFERENCE
        REF_AVG = np.mean(ref_c)

        # Step 2: RB = READ - BASE
        RB = read_c - base_c

        # Step 3: Final calc
        # Avoid division by zero
        denom = (0.0045 * REF_AVG)
        if denom == 0: denom = 1e-9
        
        final_values = (RB - REF_AVG) / denom

        # Step 4: Rolling Average (Window 256)
        # Note: np.convolve 'valid' reduces the size of the array
        window = 256
        if len(final_values) > window:
            rolling_avg = np.convolve(final_values, np.ones(window)/window, mode="valid")
            # Adjust time axis to match valid convolution result
            x_vals = time_c[:len(rolling_avg)]
        else:
            self.log("Warning: Data length < 256, skipping smoothing.")
            rolling_avg = final_values
            x_vals = time_c

        # Step 5: Statistics
        stats = {
            "Max": np.max(rolling_avg),
            "Min": np.min(rolling_avg),
            "Avg": np.mean(rolling_avg),
            "Std": np.std(rolling_avg)
        }
        
        # Step 6: FFT
        N = len(rolling_avg)
        fft_values = np.fft.fft(rolling_avg)
        fft_freqs = np.fft.fftfreq(N, d=(x_vals[1] - x_vals[0])) # Use actual time delta
        
        # Take positive half
        half_n = N // 2
        fft_mag = np.abs(fft_values)[:half_n]
        fft_frq = fft_freqs[:half_n]

        # 3. Save to Disk (Auto-Save)
        ts_str = datetime.now().strftime("%Y%m%d_%H%M%S")
        base_fn = f"processed_{ts_str}"
        
        # Save CSV
        csv_path = os.path.join(MEAS_DIR, f"{base_fn}.csv")
        try:
            with open(csv_path, "w", newline='') as f:
                f.write("Statistics\n")
                for k,v in stats.items():
                    f.write(f"{k},{v}\n")
                f.write("\nData\n")
                # Columns: Time, RawRead, RawBase, Processed, Freq, FFTMag
                # We need to pad FFT columns with empty strings if lengths differ
                writer = csv.writer(f)
                writer.writerow(["Time", "Processed_Signal", "Frequency", "FFT_Magnitude"])
                
                rows = max(len(x_vals), len(fft_frq))
                for i in range(rows):
                    t_val = x_vals[i] if i < len(x_vals) else ""
                    p_val = rolling_avg[i] if i < len(rolling_avg) else ""
                    f_val = fft_frq[i] if i < len(fft_frq) else ""
                    m_val = fft_mag[i] if i < len(fft_mag) else ""
                    writer.writerow([t_val, p_val, f_val, m_val])
            
            self.log(f"Saved Data: {csv_path}")
        except Exception as e:
            self.log(f"CSV Save Error: {e}")

        # 4. Display Window
        self.show_results_window(x_vals, rolling_avg, fft_frq, fft_mag, stats, base_fn)
        self.status_text.set("Cycle Complete. Results shown.")

    def show_results_window(self, x, y, fx, fy, stats, filename_base):
        """Creates a popup Toplevel window with Matplotlib graphs."""
        top = tk.Toplevel(self.root)
        top.title(f"Results: {filename_base}")
        top.geometry("900x700")

        # Stats Label
        stat_str = " | ".join([f"{k}: {v:.4f}" for k,v in stats.items()])
        ttk.Label(top, text=stat_str, font=("Consolas", 10)).pack(pady=5)

        # Matplotlib Figure
        fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(8, 8))
        
        # Graph 1: Time Domain
        ax1.plot(x, y, 'b-')
        ax1.set_title("Processed Signal ((Read - Base - RefAvg) / ...)")
        ax1.set_xlabel("Time (s)")
        ax1.set_ylabel("Amplitude")
        ax1.grid(True)

        # Graph 2: FFT
        ax2.plot(fx, fy, 'r-')
        ax2.set_title("FFT Spectrum")
        ax2.set_xlabel("Frequency (Hz)")
        ax2.set_ylabel("Magnitude")
        ax2.grid(True)

        plt.tight_layout()

        canvas = FigureCanvasTkAgg(fig, master=top)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True)

        # Save Image Button (Manual trigger, though we can also auto-save)
        # Let's Auto-save the image too
        img_path = os.path.join(MEAS_DIR, f"{filename_base}.png")
        fig.savefig(img_path)
        self.log(f"Saved Graph: {img_path}")

        btn_close = ttk.Button(top, text="Close", command=top.destroy)
        btn_close.pack(pady=10)

if __name__ == "__main__":
    root = tk.Tk()
    app = InstrumentApp(root)
    root.mainloop()
