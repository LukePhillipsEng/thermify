import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime
import pyvisa
from koolertron import KoolertronSig
import warnings


class InstrumentApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Tektronix DPO4034 + Koolertron Control")

        self.rm = None
        self.scope = None
        self.kool = None

        # --- Layout frame ---
        container = ttk.Frame(root, padding=20)
        container.grid(row=0, column=0, sticky="nsew")
        root.grid_columnconfigure(0, weight=1)

        ttk.Label(container, text="Instrument Control Panel", font=("Segoe UI", 14, "bold")).grid(
            row=0, column=0, columnspan=2, pady=(0, 10)
        )

        # -------------------------------
        # Inputs
        # -------------------------------
        self.scope_addr = tk.StringVar(value="USB0::0x0699::0x0401::C010101::INSTR")
        self.kool_port = tk.StringVar(value="COM3")
        self.wave_type = tk.StringVar(value="SIN")
        self.freq = tk.DoubleVar(value=1000.0)
        self.amplitude = tk.DoubleVar(value=1.0)
        self.offset = tk.DoubleVar(value=0.0)

        fields = [
            ("Scope VISA Address:", self.scope_addr),
            ("Koolertron COM Port:", self.kool_port),
        ]
        row = 1
        for label, var in fields:
            ttk.Label(container, text=label).grid(row=row, column=0, sticky="w", pady=3)
            ttk.Entry(container, textvariable=var, width=45).grid(row=row, column=1, pady=3)
            row += 1

        ttk.Label(container, text="Wave Type:").grid(row=row, column=0, sticky="w", pady=3)
        wave_menu = ttk.Combobox(
            container,
            textvariable=self.wave_type,
            values=["SIN", "PULSE", "SQUARE"],
            state="readonly",
            width=20,
        )
        wave_menu.grid(row=row, column=1, pady=3)
        row += 1

        ttk.Label(container, text="Frequency (Hz):").grid(row=row, column=0, sticky="w", pady=3)
        ttk.Entry(container, textvariable=self.freq).grid(row=row, column=1, pady=3)
        row += 1

        ttk.Label(container, text="Amplitude (V):").grid(row=row, column=0, sticky="w", pady=3)
        ttk.Entry(container, textvariable=self.amplitude).grid(row=row, column=1, pady=3)
        row += 1

        ttk.Label(container, text="Offset (V):").grid(row=row, column=0, sticky="w", pady=3)
        ttk.Entry(container, textvariable=self.offset).grid(row=row, column=1, pady=3)
        row += 1

        # -------------------------------
        # Buttons
        # -------------------------------
        ttk.Button(container, text="Connect Instruments", command=self.connect_instruments).grid(
            row=row, column=0, columnspan=2, pady=(10, 5)
        )
        row += 1
        ttk.Button(container, text="Run Measurement Sequence", command=self.run_measurement).grid(
            row=row, column=0, columnspan=2, pady=5
        )
        row += 1

        # -------------------------------
        # Status
        # -------------------------------
        self.status_label = ttk.Label(container, text="Status: Not Connected", foreground="red")
        self.status_label.grid(row=row, column=0, columnspan=2, pady=10)
        row += 1

        self.progress_label = ttk.Label(container, text="Ready.")
        self.progress_label.grid(row=row, column=0, columnspan=2, pady=(5, 0))

    # -------------------------------
    # Connection
    # -------------------------------
    def connect_instruments(self):
        try:
            # Ignore tm_devices warnings
            warnings.filterwarnings("ignore", message="The \"401\" model is not supported by tm_devices")

            # Connect to Tektronix scope via PyVISA
            if not self.rm:
                self.rm = pyvisa.ResourceManager()

            if not self.scope:
                addr = self.scope_addr.get()
                self.scope = self.rm.open_resource(addr)
                self.scope.timeout = 5000
                idn = self.scope.query("*IDN?")
                print(f"Connected to scope: {idn}")

            # Connect to Koolertron generator
            if not self.kool:
                self.kool = KoolertronSig(self.kool_port.get())
                if not self.kool.isConnected():
                    raise Exception("Koolertron not responding.")

            self.status_label.config(text="Status: Connected", foreground="green")
            self.progress_label.config(text="Instruments ready.")
            messagebox.showinfo("Connection", "Tektronix and Koolertron connected successfully!")

        except Exception as e:
            self.status_label.config(text="Status: Connection Failed", foreground="red")
            messagebox.showerror("Connection Error", f"Could not connect instruments:\n{e}")

    # -------------------------------
    # Measurement
    # -------------------------------
    def run_measurement(self):
        if not self.scope or not self.kool:
            messagebox.showwarning("Not Connected", "Please connect instruments first.")
            return

        wave = self.wave_type.get().upper()
        freq = self.freq.get()
        amp = self.amplitude.get()
        offset = self.offset.get()

        try:
            # Configure Koolertron both channels
            if wave == "SIN":
                self.kool.sinwave(freq, amp, 0, offset, 1)
                self.kool.sinwave(freq, amp, 0, offset, 2)
            elif wave in ("PULSE", "SQUARE"):
                self.kool.squareWave(freq, amp, 0, offset, 1)
                self.kool.squareWave(freq, amp, 0, offset, 2)
            else:
                messagebox.showwarning("Wave Error", f"Unsupported wave type: {wave}")
                return

            # ------------------------------
            # Step 1: Reference capture
            # ------------------------------
            self.progress_label.config(text="Step 1: Take reference measurement.")
            messagebox.showinfo("Step 1", "Take reference measurement (MATH CH3). Click OK when ready.")
            ref_name = f"reference_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
            self.scope.write(f"SAVE:WAVEFORM MATH,'{ref_name}'")

            # ------------------------------
            # Step 2: Reading with DDS ON
            # ------------------------------
            self.progress_label.config(text="Step 2: Capture reading with DDS ON.")
            messagebox.showinfo("Step 2", "Turn ON DDS & reduce noise. Click OK when ready.")
            meas_name = f"CH3_{freq}_{offset}_Reading.csv"
            self.scope.write(f"SAVE:WAVEFORM MATH,'{meas_name}'")

            # ------------------------------
            # Step 3: Base measurement
            # ------------------------------
            self.progress_label.config(text="Step 3: Capture base measurement.")
            messagebox.showinfo("Step 3", "Turn OFF DDS manually. Click OK to capture base measurement.")
            base_name = f"{freq}_{offset}_base.csv"
            self.scope.write(f"SAVE:WAVEFORM MATH,'{base_name}'")

            self.progress_label.config(text="Sequence complete.")
            messagebox.showinfo(
                "Complete",
                f"Measurements saved:\n{ref_name}\n{meas_name}\n{base_name}",
            )

        except Exception as e:
            self.progress_label.config(text="Error during measurement.")
            messagebox.showerror("Run Error", f"Error during measurement:\n{e}")


# -------------------------------
# Main
# -------------------------------
if __name__ == "__main__":
    root = tk.Tk()
    app = InstrumentApp(root)
    root.mainloop()
