"""Generate sample I-V data for a non-ohmic resistor (NTC-type, e.g. a thermistor
or semiconductor resistor) at three temperatures, in both supported layouts.

Model: I = (V / R(T)) * (1 + b * V^2)
  - R(T) falls as temperature rises (NTC: R = R25 * exp(B * (1/T - 1/T25)))
  - the b*V^2 term bends the curve upward, so I is not proportional to V
"""

import numpy as np
import pandas as pd

rng = np.random.default_rng(7)
R25, B, b = 100.0, 3500.0, 0.06          # ohms, kelvin, 1/V^2
temps_c = [25, 50, 75]


def current_ma(v, t_c):
    t_k, t25 = t_c + 273.15, 298.15
    r = R25 * np.exp(B * (1 / t_k - 1 / t25))
    i = (v / r) * (1 + b * v**2) * 1000                  # mA
    return i * (1 + rng.normal(0, 0.015, size=np.shape(v)))  # ~1.5% noise


rows = []
for t in temps_c:
    # Give the 75 °C trial different voltage steps to show per-series X values work.
    volts = np.arange(0, 5.01, 0.5) if t != 75 else np.arange(0, 4.81, 0.4)
    for v, i in zip(volts, current_ma(volts, t)):
        rows.append({"Temperature (°C)": t, "Voltage (V)": round(v, 2), "Current (mA)": round(i, 3)})
long_df = pd.DataFrame(rows)

volts = np.arange(0, 5.01, 0.5)
wide_df = pd.DataFrame({"Voltage (V)": volts})
for t in temps_c:
    wide_df[f"{t} °C"] = np.round(current_ma(volts, t), 3)

with pd.ExcelWriter("sample_iv_data.xlsx") as xw:
    long_df.to_excel(xw, sheet_name="Long format", index=False)
    wide_df.to_excel(xw, sheet_name="Wide format", index=False)

print(long_df.groupby("Temperature (°C)")["Current (mA)"].max())
print("wrote sample_iv_data.xlsx")
