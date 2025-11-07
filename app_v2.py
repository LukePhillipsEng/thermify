import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime
import pyvisa
from koolertron import KoolertronSig
import numpy as np
import csv
import time


class InstrumentApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Tektronix DPO4034 + Koolertron Controller")
        self.root.geometry("480x420")

        # cached connections
        self.rm = None
        self.scope = None
        self.kool = None

        # ---------------- UI Layout ----------------
        container = ttk.Frame(root, padding=20)
        container.grid(row=0, column=0, sticky="nsew")
        root.grid_columnconfigure(0, weight=1)

        ttk.Label(container, text="Instrument Control Panel",
                  font=("Segoe UI", 15, "bold")).grid(row=0, column=0, columnspan=2, pady=(0, 10))

        # Inputs
        self.scope_addr = tk.StringVar(value="USB0::0x0699::0x0401::C010101::INSTR")
        self.kool_port = tk.StringVar(value="COM3")
        self.wave_type = tk.StringVar(value="SIN")
        self.freq = tk.DoubleVar(value=1000.0)
        self.amp = tk.DoubleVar(value=1.0)
        self.offset = tk.DoubleVar(value=0.0)

        row = 1
        for label, var in [
            ("Scope VISA Address:", self.scope_addr),
            ("Koolertron COM Port:", self.kool_port),
        ]:
            ttk.Label(container, text=label).grid(row=row, column=0, sticky="w", pady=3)
            ttk.Entry(container, textvariable=var, width=40).grid(row=row, column=1, pady=3)
            row += 1

        ttk.Label(container, text="Wave Type:").grid(row=row, column=0, sticky="w", pady=3)
        ttk.Combobox(container, textvariable=self.wave_type,
                     values=["SIN", "PULSE", "SQUARE"], state="readonly", width=15).grid(row=row, column=1, pady=3)
        row += 1

        for label, var, unit in [
            ("Frequency:", self.freq, "Hz"),
            ("Amplitude:", self.amp, "Vpp"),
            ("Offset:", self.offset, "V"),
        ]:
            ttk.Label(container, text=label).grid(row=row, column=0, sticky="w", pady=3)
            ttk.Entry(container, textvariable=var, width=20).grid(row=row, column=1, sticky="w", pady=3)
            row += 1

        # Buttons
        ttk.Button(container, text="Connect Instruments",
                   command=self.connect_instruments).grid(row=row, column=0, columnspan=2, pady=(12, 5))
        row += 1
        ttk.Button(container, text="Run Measurement Sequence",
                   command=self.run_measurement).grid(row=row, column=0, columnspan=2, pady=5)
        row += 1

        # Status labels
        self.status_label = ttk.Label(container, text="Status: Not Connected", foreground="red")
        self.status_label.grid(row=row, column=0, columnspan=2, pady=(10, 3))
        row += 1
        self.progress_label = ttk.Label(container, text="Ready.")
        self.progress_label.grid(row=row, column=0, columnspan=2, pady=(5, 0))

    # ---------------- Connection ----------------
    def connect_instruments(self):
        try:
            if not self.rm:
                self.rm = pyvisa.ResourceManager()
                print("Resources:", self.rm.list_resources())

            if not self.scope:
                self.scope = self.rm.open_resource(self.scope_addr.get())
                self.scope.timeout = 10000
                idn = self.scope.query("*IDN?")
                print("Connected to:", idn.strip())

            if not self.kool:
                self.kool = KoolertronSig(self.kool_port.get())
                if not self.kool.isConnected():
                    raise Exception("Koolertron not responding.")

            self.status_label.config(text="Status: Connected", foreground="green")
            self.progress_label.config(text="Tektronix & Koolertron ready.")
            messagebox.showinfo("Connected", "Instruments connected successfully.")

        except Exception as e:
            self.status_label.config(text="Status: Connection Failed", foreground="red")
            messagebox.showerror("Connection Error", str(e))

    # ---------------- Measurement ----------------
    def run_measurement(self):
        if not self.scope or not self.kool:
            messagebox.showwarning("Not Connected", "Please connect instruments first.")
            return

        wave = self.wave_type.get().upper()
        freq = self.freq.get()
        amp = self.amp.get()
        offset = self.offset.get()
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

        try:
            # Configure Koolertron for both channels
            if wave == "SIN":
                self.kool.sinwave(freq, amp, 0, offset, 1)
                self.kool.sinwave(freq, amp, 0, offset, 2)
            else:
                self.kool.squareWave(freq, amp, 0, offset, 1)
                self.kool.squareWave(freq, amp, 0, offset, 2)

            # ---- Step 1: Reference ----
            self.progress_label.config(text="Step 1: Take reference measurement.")
            messagebox.showinfo("Step 1", "Take reference (DDS OFF). Click OK when done.")
            ref_file = f"reference_{timestamp}.csv"
            self._read_math_waveform(ref_file)

            # ---- Step 2: Main Reading ----
            self.progress_label.config(text="Step 2: Capture reading with DDS ON.")
            messagebox.showinfo("Step 2", "Turn ON Koolertron and reduce noise, then click OK.")
            meas_file = f"CH3_{freq}_{offset}_Reading.csv"
            self._read_math_waveform(meas_file)

            # ---- Step 3: Base ----
            self.progress_label.config(text="Step 3: Capture base measurement.")
            messagebox.showinfo("Step 3", "Turn OFF DDS manually, then click OK to capture base.")
            base_file = f"{freq}_{offset}_base.csv"
            self._read_math_waveform(base_file)

            self.progress_label.config(text="✅ Sequence complete.")
            messagebox.showinfo("Complete", f"Measurements saved:\n{ref_file}\n{meas_file}\n{base_file}")

        except Exception as e:
            self.progress_label.config(text="❌ Error during measurement.")
            messagebox.showerror("Measurement Error", str(e))

    # ---------------- Helper: Read math waveform ----------------
    def _read_math_waveform(self, filename):
        """Read existing MATH waveform safely from DPO4034 and save to CSV."""
        try:
            ch = "MATH"
            self.scope.write(f"DATA:SOURCE {ch}")
            self.scope.write("DATA:WIDTH 1")
            self.scope.write("DATA:ENC RPB")

            xinc = float(self.scope.query("WFMPRE:XINCR?"))
            xorg = float(self.scope.query("WFMPRE:XZERO?"))
            ymult = float(self.scope.query("WFMPRE:YMULT?"))
            yoff = float(self.scope.query("WFMPRE:YOFF?"))
            yzero = float(self.scope.query("WFMPRE:YZERO?"))

            raw = self.scope.query_binary_values("CURVE?", datatype='B', container=np.array)
            volt = (raw - yoff) * ymult + yzero
            time_data = np.arange(len(volt)) * xinc + xorg

            with open(filename, "w", newline="") as f:
                w = csv.writer(f)
                w.writerow(["Time (s)", "Voltage (V)"])
                w.writerows(zip(time_data, volt))
            print(f"Saved waveform: {filename}")

        except Exception as e:
            raise Exception(f"Waveform read failed: {e}")


# ---------------- Run App ----------------
if __name__ == "__main__":
    root = tk.Tk()
    app = InstrumentApp(root)
    root.mainloop()
