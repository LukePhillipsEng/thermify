import pyvisa
import serial
import time
import numpy as np
import matplotlib.pyplot as plt

# ========================
# CONNECTION INITIALIZATION
# ========================

class InstrumentManager:
    """Manages and caches VISA + Serial connections."""
    def __init__(self):
        self.rm = pyvisa.ResourceManager()
        self.scope = None
        self.generator = None

    def connect_scope(self, visa_addr_hint="USB::0x0699::0x0408::INSTR"):
        """Connect to Tektronix DPO4034 via VISA."""
        if self.scope:
            print("✅ Oscilloscope already connected.")
            return self.scope
        resources = self.rm.list_resources()
        for r in resources:
            if "USB" in r or "GPIB" in r:
                try:
                    inst = self.rm.open_resource(r)
                    idn = inst.query("*IDN?")
                    if "TEKTRONIX" in idn.upper():
                        print(f"✅ Connected to Tektronix DPO4034: {idn.strip()}")
                        self.scope = inst
                        return inst
                except Exception as e:
                    print(f"Skipping {r}: {e}")
        raise ConnectionError("❌ DPO4034 not found via VISA.")

    def connect_generator(self, port_hint="COM3"):
        """Connect to Koolertron signal generator via serial."""
        if self.generator:
            print("✅ Generator already connected.")
            return self.generator
        try:
            gen = serial.Serial(port_hint, baudrate=115200, timeout=2)
            time.sleep(1)
            print(f"✅ Connected to Koolertron generator on {port_hint}")
            self.generator = gen
            return gen
        except Exception as e:
            raise ConnectionError(f"❌ Failed to connect to Koolertron: {e}")


# ========================
# SIGNAL GENERATOR CONTROL
# ========================

def set_waveform(gen, wave_type="SINE", freq=1000, amp=1.0):
    """Set Koolertron waveform parameters."""
    wave_dict = {"SINE": 0, "SQUARE": 1, "TRIANGLE": 2}
    if wave_type not in wave_dict:
        raise ValueError("Invalid wave type. Choose SINE, SQUARE, TRIANGLE.")
    try:
        cmd = f":WAVE {wave_dict[wave_type]},{freq},{amp}\n"
        gen.write(cmd.encode())
        time.sleep(0.5)
        print(f"📡 Koolertron set to {wave_type}, {freq} Hz, {amp} Vpp")
    except Exception as e:
        print(f"❌ Generator write failed: {e}")


# ========================
# OSCILLOSCOPE READING
# ========================

def read_waveform(scope, channel=1):
    """Read waveform data from Tektronix DPO4034."""
    try:
        scope.write(f"DATA:SOURCE CH{channel}")
        scope.write("DATA:ENCDG ASCII")
        scope.write("CURVE?")
        data_str = scope.read_raw().decode(errors="ignore").strip()
        data_points = [float(x) for x in data_str.split(",") if x.strip()]
        print(f"📈 Retrieved {len(data_points)} points from CH{channel}")
        return np.array(data_points)
    except Exception as e:
        print(f"❌ Error reading from scope CH{channel}: {e}")
        return np.array([])


# ========================
# EXPERIMENT SEQUENCE
# ========================

def run_experiment(scope, gen):
    """Run through multiple waveform types, read and plot data for each."""
    wave_types = ["SINE", "SQUARE", "TRIANGLE"]

    for wave in wave_types:
        print(f"\n=== Running {wave} Wave Measurement ===")
        set_waveform(gen, wave_type=wave, freq=1000, amp=1.0)
        time.sleep(2)

        data = read_waveform(scope, channel=1)
        if data.size == 0:
            print("⚠️ No data received. Skipping plot.")
            continue

        plt.figure()
        plt.plot(data)
        plt.title(f"{wave} Waveform from DPO4034")
        plt.xlabel("Sample #")
        plt.ylabel("Voltage (V)")
        plt.grid(True)
        plt.show()

        time.sleep(1)


# ========================
# MAIN
# ========================

if __name__ == "__main__":
    instruments = InstrumentManager()
    try:
        scope = instruments.connect_scope()
        gen = instruments.connect_generator(port_hint="COM3")  # change COM port as needed
        run_experiment(scope, gen)
    except Exception as e:
        print(e)
