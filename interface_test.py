import pyvisa
import numpy as np
import csv
from datetime import datetime

def main():
    rm = pyvisa.ResourceManager()
    print("🔍 Available VISA resources:")
    for r in rm.list_resources():
        print("   ", r)

    # --- Connect to your Tektronix DPO4034 ---
    scope_addr = "USB0::0x0699::0x0401::C010101::INSTR"  # Update if needed
    print(f"\nConnecting to: {scope_addr}")
    scope = rm.open_resource(scope_addr)
    scope.timeout = 10000  # 10s timeout for large data

    # --- Identify the scope ---
    idn = scope.query("*IDN?")
    print(f"✅ Connected to: {idn.strip()}")

    # --- Choose channel to read (read-only) ---
    channel = "CH2"
    print(f"\n📡 Reading existing waveform from {channel}...")

    # Select channel as source (read-only setup)
    scope.write(f"DATA:SOURCE {channel}")
    scope.write("DATA:WIDTH 1")  # 1 byte per sample
    scope.write("DATA:ENC RPB")  # little-endian binary

    # --- Get scaling parameters (doesn't modify state) ---
    x_increment = float(scope.query("WFMPRE:XINCR?"))
    x_origin = float(scope.query("WFMPRE:XZERO?"))
    y_mult = float(scope.query("WFMPRE:YMULT?"))
    y_offset = float(scope.query("WFMPRE:YOFF?"))
    y_zero = float(scope.query("WFMPRE:YZERO?"))

    # --- Read waveform (safe) ---
    print("Reading waveform data (this may take a few seconds)...")
    raw = scope.query_binary_values("CURVE?", datatype='B', container=np.array)

    # --- Convert to real values ---
    voltage = (raw - y_offset) * y_mult + y_zero
    time = np.arange(len(voltage)) * x_increment + x_origin

    # --- Save waveform (optional) ---
    filename = f"DPO4034_{channel}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
    with open(filename, "w", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(["Time (s)", "Voltage (V)"])
        writer.writerows(zip(time, voltage))

    print(f"✅ Waveform from {channel} read successfully.")
    print(f"💾 Saved {len(voltage)} points to {filename}")

    # --- Cleanup ---
    scope.close()
    rm.close()

if __name__ == "__main__":
    main()
