#!/usr/bin/env python3
"""
koolertron_dpo_keithley_final.py

- Tektronix DPO4034 via PyVISA (read-only)
- KoolertronSig via serial (uses your provided driver)
- Keithley 2400 via PyVISA (simple ON/OFF control)
- Step 1: Connect devices (cached)
- Step 2: Keithley ON, take 10 reads from scope MATH channel, average, save CSV
- Step 3: Keithley OFF, take 10 reads, average, save CSV
- Duty cycle control via KoolertronSig.setDuty(0..1, channel)
- Saves CSVs to ./measurements/
"""

import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime
import pyvisa
import numpy as np
import csv
import time
import os
import threading
from koolertron import KoolertronSig  # your driver file

# ---------------- Configuration ----------------
NUM_AVERAGES = 10
READ_DELAY_SEC = 0.03  # 30 ms between rapid reads (tune if needed)
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
        self.root.title("DPO4034 + Koolertron + Keithley")
        self.root.geometry("640x560")

        # Cached connections
        self.rm = None
        self.scope = None
        self.kool = None
        self.keith = None

        # UI variables
        self.scope_addr = tk.StringVar(value=DEFAULT_SCOPE_ADDR)
        self.kool_port = tk.StringVar(value="COM3")
        self.keith_addr = tk.StringVar(value="")  # user-pasted VISA string
        self.wave_type = tk.StringVar(value="SIN")
        self.freq = tk.DoubleVar(value=1000.0)
        self.amp = tk.DoubleVar(value=1.0)
        self.offset = tk.DoubleVar(value=0.0)
        self.duty_percent = tk.DoubleVar(value=50.0)
        self.channel = tk.StringVar(value="MATH")  # channel to read; default MATH
        self.status_text = tk.StringVar(value="Not connected")

        self._build_gui()

    def _build_gui(self):
        pad = dict(padx=8, pady=6)
        frm = ttk.Frame(self.root, padding=12)
        frm.pack(fill="both", expand=True)

        ttk.Label(frm, text="Instrument Panel", font=("Segoe UI", 16, "bold")).grid(row=0, column=0, columnspan=2, **pad)

        r = 1
        ttk.Label(frm, text="Scope VISA Address:").grid(row=r, column=0, sticky="w", **pad)
        ttk.Entry(frm, textvariable=self.scope_addr, width=50).grid(row=r, column=1, sticky="w", **pad)
        r += 1

        ttk.Label(frm, text="Koolertron COM Port:").grid(row=r, column=0, sticky="w", **pad)
        ttk.Entry(frm, textvariable=self.kool_port, width=20).grid(row=r, column=1, sticky="w", **pad)
        r += 1

        ttk.Label(frm, text="Keithley VISA Address (paste):").grid(row=r, column=0, sticky="w", **pad)
        ttk.Entry(frm, textvariable=self.keith_addr, width=50).grid(row=r, column=1, sticky="w", **pad)
        r += 1

        ttk.Label(frm, text="Wave Type:").grid(row=r, column=0, sticky="w", **pad)
        ttk.Combobox(frm, textvariable=self.wave_type, values=["SIN", "PULSE", "SQUARE"], state="readonly", width=20).grid(row=r, column=1, sticky="w", **pad)
        r += 1

        ttk.Label(frm, text="Frequency (Hz):").grid(row=r, column=0, sticky="w", **pad)
        ttk.Entry(frm, textvariable=self.freq, width=20).grid(row=r, column=1, sticky="w", **pad)
        r += 1

        ttk.Label(frm, text="Amplitude (Vpp):").grid(row=r, column=0, sticky="w", **pad)
        ttk.Entry(frm, textvariable=self.amp, width=20).grid(row=r, column=1, sticky="w", **pad)
        r += 1

        ttk.Label(frm, text="Offset (V):").grid(row=r, column=0, sticky="w", **pad)
        ttk.Entry(frm, textvariable=self.offset, width=20).grid(row=r, column=1, sticky="w", **pad)
        r += 1

        ttk.Label(frm, text="Duty Cycle (%):").grid(row=r, column=0, sticky="w", **pad)
        ttk.Entry(frm, textvariable=self.duty_percent, width=20).grid(row=r, column=1, sticky="w", **pad)
        r += 1

        ttk.Label(frm, text="Scope Read Source:").grid(row=r, column=0, sticky="w", **pad)
        ttk.Combobox(frm, textvariable=self.channel, values=["CH1", "CH2", "CH3", "CH4", "MATH"], state="readonly", width=20).grid(row=r, column=1, sticky="w", **pad)
        r += 1

        ttk.Button(frm, text="Connect Devices (Step 1)", command=self.connect_devices).grid(row=r, column=0, columnspan=2, pady=(12,6))
        r += 1

        ttk.Button(frm, text="Step 2: Keithley ON -> Read (10x avg)", command=self._start_step_thread(self.step2_action)).grid(row=r, column=0, columnspan=2, pady=6)
        r += 1

        ttk.Button(frm, text="Step 3: Keithley OFF -> Read (10x avg)", command=self._start_step_thread(self.step3_action)).grid(row=r, column=0, columnspan=2, pady=6)
        r += 1

        ttk.Button(frm, text="Disconnect Devices", command=self.disconnect_devices).grid(row=r, column=0, columnspan=2, pady=(12,4))
        r += 1

        ttk.Separator(frm).grid(row=r, column=0, columnspan=2, sticky="ew", pady=(8,8))
        r += 1

        ttk.Label(frm, text="Status:").grid(row=r, column=0, sticky="w", **pad)
        ttk.Label(frm, textvariable=self.status_text, foreground="blue").grid(row=r, column=1, sticky="w", **pad)
        r += 1

        ttk.Label(frm, text="Log:").grid(row=r, column=0, sticky="nw", **pad)
        self.logbox = tk.Text(frm, height=10, width=70, state="disabled")
        self.logbox.grid(row=r, column=1, sticky="w", **pad)
        r += 1

    def _start_step_thread(self, target_func):
        # Return a function that starts target_func in a thread (to keep UI responsive)
        def start_thread():
            t = threading.Thread(target=target_func, daemon=True)
            t.start()
        return start_thread

    def log(self, msg):
        ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        line = f"[{ts}] {msg}\n"
        self.logbox.configure(state="normal")
        self.logbox.insert("end", line)
        self.logbox.see("end")
        self.logbox.configure(state="disabled")

    # ---------------- Connection (Step 1) ----------------
    def connect_devices(self):
        try:
            # Create ResourceManager once
            if not self.rm:
                self.rm = pyvisa.ResourceManager()
                self.log(f"ResourceManager created. Resources: {self.rm.list_resources()}")

            # Connect scope once
            if not self.scope:
                addr = self.scope_addr.get().strip()
                if not addr:
                    raise RuntimeError("Scope VISA address is empty")
                self.scope = self.rm.open_resource(addr)
                self.scope.timeout = SCOPE_TIMEOUT_MS
                idn = self.scope.query("*IDN?")
                self.log(f"Scope connected: {idn.strip()}")

            # Create Koolertron object (driver handles serial per-command)
            if not self.kool:
                port = self.kool_port.get().strip()
                if not port:
                    raise RuntimeError("Koolertron COM port is empty")
                self.kool = KoolertronSig(port)
                if not self.kool.isConnected():
                    raise RuntimeError(f"Koolertron not responding on {port}")
                self.log(f"Koolertron connected on {port}")

            # Keithley optional: don't fail if not provided, but try to open if user pasted address
            kaddr = self.keith_addr.get().strip()
            if kaddr and not self.keith:
                try:
                    self.keith = self.rm.open_resource(kaddr)
                    self.keith.timeout = 5000
                    idn_k = self.keith.query("*IDN?")
                    self.log(f"Keithley connected: {idn_k.strip()}")
                except Exception as e:
                    # Keep going; keithley operations will warn if needed
                    self.log(f"Keithley connection warning: {e}")

            self.status_text.set("Connected")
            self.log("Devices connected and cached.")
            messagebox.showinfo("Connected", "Devices connected and cached (Step 1 complete).")

        except Exception as e:
            self.status_text.set("Connection failed")
            self.log(f"Connection error: {e}")
            messagebox.showerror("Connection Error", str(e))

    def disconnect_devices(self):
        try:
            if self.scope:
                try:
                    self.scope.close()
                except Exception:
                    pass
                self.scope = None
            if self.keith:
                try:
                    self.keith.close()
                except Exception:
                    pass
                self.keith = None
            # KoolertronSig uses transient serial in sendCommand, so just drop the object
            self.kool = None
            self.status_text.set("Disconnected")
            self.log("Devices disconnected and cache cleared.")
            messagebox.showinfo("Disconnected", "Devices disconnected.")
        except Exception as e:
            self.log(f"Disconnect error: {e}")
            messagebox.showerror("Disconnect Error", str(e))

    # ---------------- Step 2: Keithley ON -> Read (10x avg) ----------------
    def step2_action(self):
        # turn Keithley ON, then acquire averaged waveform
        if not self.scope or not self.kool:
            messagebox.showwarning("Not connected", "Please connect devices first (Step 1).")
            return

        # Configure generator (wave + duty + amplitude ON)
        if not self._configure_koolertron(on=True):
            return

        # Keithley ON
        if not self._keithley_set_output(True):
            # if user doesn't have keithley connected, continue but warn
            self.log("Proceeding without Keithley ON (not connected).")

        # Capture and save averaged waveform
        self._capture_and_save(step_name="Step2_ON")

        # Turn Keithley OFF after reading
        self._keithley_set_output(False)

    # ---------------- Step 3: Keithley OFF -> Read (10x avg) ----------------
    def step3_action(self):
        # turn Koolertron output OFF (via amplitude 0), Keithley ON for measurement, capture averaged waveform, then Keithley OFF
        if not self.scope or not self.kool:
            messagebox.showwarning("Not connected", "Please connect devices first (Step 1).")
            return

        # Ensure generator is OFF (amplitude 0)
        if not self._configure_koolertron(on=False):
            return

        # Keithley ON
        if not self._keithley_set_output(True):
            self.log("Proceeding without Keithley ON (not connected).")

        # Capture and save averaged waveform
        self._capture_and_save(step_name="Step3_OFF")

        # Turn Keithley OFF
        self._keithley_set_output(False)

    # -------------- Helper: configure Koolertron --------------
    def _configure_koolertron(self, on: bool) -> bool:
        """Configure waveform, duty cycle, amplitude. If on==False, set amplitude=0 to disable output."""
        try:
            wave = self.wave_type.get().upper()
            freq = float(self.freq.get())
            amp = float(self.amp.get())
            offset = float(self.offset.get())
            duty_pct = float(self.duty_percent.get())
            duty_frac = max(0.0, min(1.0, duty_pct / 100.0))

            # choose waveform method for both channels
            if wave == "SIN":
                self.kool.sinwave(freq, amp, 0, offset, chan=1)
                self.kool.sinwave(freq, amp, 0, offset, chan=2)
            else:
                # PULSE & SQUARE use squareWave (as per driver)
                self.kool.squareWave(freq, amp, 0, offset, chan=1)
                self.kool.squareWave(freq, amp, 0, offset, chan=2)

            # set duty on both channels
            self.kool.setDuty(duty_frac, channel=1)
            self.kool.setDuty(duty_frac, channel=2)

            # amplitude on/off: driver uses volts -> setAmplitude accepts volts
            if on:
                self.kool.setAmplitude(amp, channel=1)
                self.kool.setAmplitude(amp, channel=2)
                self.log(f"Koolertron: configured and OUTPUT=ON (amp={amp} Vpp, duty={duty_pct}%)")
            else:
                self.kool.setAmplitude(0.0, channel=1)
                self.kool.setAmplitude(0.0, channel=2)
                self.log("Koolertron: OUTPUT=OFF (amplitude set to 0 V)")

            return True
        except Exception as e:
            self.log(f"Koolertron configuration error: {e}")
            messagebox.showerror("Koolertron Error", f"Configuration failed: {e}")
            return False

    # -------------- Helper: Keithley ON/OFF --------------
    def _keithley_set_output(self, on: bool) -> bool:
        """Turn Keithley output ON/OFF. Returns True if command sent, False if no keithley connected."""
        if not self.keith:
            kaddr = self.keith_addr.get().strip()
            if kaddr:
                try:
                    self.keith = self.rm.open_resource(kaddr)
                    self.keith.timeout = 5000
                except Exception as e:
                    self.log(f"Keithley open failed: {e}")
                    return False
            else:
                self.log("No Keithley address provided.")
                return False

        try:
            if on:
                self.keith.write(KEITHLEY_OUTPUT_ON)
                self.log("Keithley: OUTPUT ON")
            else:
                self.keith.write(KEITHLEY_OUTPUT_OFF)
                self.log("Keithley: OUTPUT OFF")
            return True
        except Exception as e:
            self.log(f"Keithley command error: {e}")
            return False

    # -------------- Capture & average --------------
    def _capture_and_save(self, step_name: str):
        """Perform NUM_AVERAGES rapid reads, average pointwise, save CSV (time, voltage)."""
        try:
            self.status_text.set(f"{step_name}: capturing {NUM_AVERAGES} acquisitions...")
            self.log(f"{step_name}: starting {NUM_AVERAGES} rapid acquisitions")
            all_volts = []
            time_axis = None

            for i in range(NUM_AVERAGES):
                volt, t = self._read_scope_once()
                if volt is None:
                    raise RuntimeError("Scope read failed during acquisition")
                all_volts.append(volt)
                if time_axis is None:
                    time_axis = t
                time.sleep(READ_DELAY_SEC)

            # Align lengths (crop to min length)
            lengths = [v.size for v in all_volts]
            min_len = min(lengths)
            if any(L != min_len for L in lengths):
                self.log(f"Captured lengths varied; cropping to {min_len} samples")
                all_volts = [v[:min_len] for v in all_volts]
                time_axis = time_axis[:min_len]

            stacked = np.vstack(all_volts)
            averaged = np.mean(stacked, axis=0)

            # Save CSV
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            freq_str = f"{int(self.freq.get())}Hz"
            offset_str = f"{self.offset.get()}V"
            filename = os.path.join(MEAS_DIR, f"{step_name}_{freq_str}_{offset_str}_avg_{ts}.csv")
            self._save_csv(filename, time_axis, averaged)
            self.log(f"{step_name}: averaged waveform saved to {filename}")
            self.status_text.set(f"{step_name}: saved {os.path.basename(filename)}")
            messagebox.showinfo("Capture Complete", f"{step_name}: saved {os.path.basename(filename)}")
        except Exception as e:
            self.log(f"{step_name} capture error: {e}")
            messagebox.showerror("Capture Error", str(e))

    # -------------- Single read from scope --------------
    def _read_scope_once(self):
        """Read the selected source (CHx or MATH) once and return (volt_array, time_array)."""
        try:
            src = self.channel.get().strip().upper()
            if src not in ("CH1", "CH2", "CH3", "CH4", "MATH"):
                src = "MATH"

            # Configure data format read-only
            self.scope.write(f"DATA:SOURCE {src}")
            self.scope.write("DATA:WIDTH 1")
            self.scope.write("DATA:ENC RPB")

            # Query waveform scaling parameters
            xinc = float(self.scope.query("WFMPRE:XINCR?"))
            xorg = float(self.scope.query("WFMPRE:XZERO?"))
            ymult = float(self.scope.query("WFMPRE:YMULT?"))
            yoff = float(self.scope.query("WFMPRE:YOFF?"))
            yzero = float(self.scope.query("WFMPRE:YZERO?"))

            # Read binary data (1-byte unsigned)
            raw = self.scope.query_binary_values("CURVE?", datatype='B', container=np.array)
            volt = (raw - yoff) * ymult + yzero
            t = np.arange(len(volt)) * xinc + xorg
            return volt, t
        except Exception as e:
            self.log(f"Scope read error: {e}")
            return None, None

    # -------------- Save CSV --------------
    def _save_csv(self, filename, time_axis, volt_axis):
        with open(filename, "w", newline="") as f:
            w = csv.writer(f)
            w.writerow(["Time (s)", "Voltage (V)"])
            w.writerows(zip(time_axis.tolist(), volt_axis.tolist()))

# -------------------- Run --------------------
def main():
    root = tk.Tk()
    app = InstrumentApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
