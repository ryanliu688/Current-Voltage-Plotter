"""Streamlit web app: upload a spreadsheet of multi-series X/Y data and download
a PowerPoint slide with a labeled, multi-color, editable chart.

Run locally:   streamlit run app.py
"""

import matplotlib.pyplot as plt
import streamlit as st

import plotter

st.set_page_config(page_title="Spreadsheet → PowerPoint Plotter", page_icon="📈", layout="wide")
st.title("📈 Spreadsheet → PowerPoint Plotter")
st.caption(
    "Upload an Excel or CSV file with X/Y data for several trials (e.g. voltage vs. current "
    "at different temperatures) and download a slide with an editable chart."
)

with st.expander("What should my spreadsheet look like?"):
    st.markdown(
        "**Long format (recommended)**: one row per measurement, with a column that labels the trial.\n\n"
        "| Temperature (°C) | Voltage (V) | Current (mA) |\n|---|---|---|\n"
        "| 25 | 0.5 | 4.9 |\n| 25 | 1.0 | 9.8 |\n| 75 | 0.5 | 21.7 |\n\n"
        "**Wide format**: one shared X column, then one Y column per trial.\n\n"
        "| Voltage (V) | 25 °C | 50 °C | 75 °C |\n|---|---|---|---|\n"
        "| 0.5 | 4.9 | 11.1 | 21.7 |\n\n"
        "Trials in long format can each have different X values."
    )

uploaded = st.file_uploader("Spreadsheet", type=["xlsx", "xls", "csv"])
if uploaded is None:
    st.info("Upload a file to get started.")
    st.stop()

# ------------------------------------------------------------- read data ---
try:
    sheets = plotter.list_sheets(uploaded)
    sheet = st.selectbox("Sheet", sheets) if len(sheets) > 1 else sheets[0]
    uploaded.seek(0)
    df = plotter.read_table(uploaded, sheet_name=0 if sheet == "(csv)" else sheet)
except Exception as e:  # noqa: BLE001 - show any read problem to the user
    st.error(f"Couldn't read that file: {e}")
    st.stop()

st.dataframe(df.head(10), use_container_width=True)
cols = list(df.columns)

# --------------------------------------------------------- map columns ---
layouts = {"long": "Long: a column labels each trial", "wide": "Wide: one Y column per trial"}
layout = st.radio(
    "Data layout", list(layouts), format_func=layouts.get, horizontal=True,
    index=list(layouts).index(plotter.guess_layout(df)),
)

c1, c2, c3 = st.columns(3)
if layout == "long":
    if len(cols) < 3:
        st.error("Long format needs at least 3 columns (label, X, Y).")
        st.stop()
    g_guess, x_guess, y_guess = plotter.guess_long_columns(df)
    group_col = c1.selectbox("Trial label column (one line each)", cols, index=cols.index(g_guess))
    x_col = c2.selectbox("X column", cols, index=cols.index(x_guess))
    y_col = c3.selectbox("Y column", cols, index=cols.index(y_guess))
    y_cols = [y_col]
else:
    if len(cols) < 2:
        st.error("Wide format needs an X column and at least one Y column.")
        st.stop()
    group_col = None
    x_col = c1.selectbox("X column", cols, index=0)
    y_cols = c2.multiselect("Y columns (one line each)", [c for c in cols if c != x_col],
                            default=[c for c in cols if c != x_col])
    y_col = y_cols[0] if y_cols else ""

# -------------------------------------------------------------- labels ---
st.subheader("Labels")
l1, l2 = st.columns(2)
def _short(name):  # "Current (mA)" -> "Current"
    return str(name).split("(")[0].strip()


def _unit(name):  # "Temperature (°C)" -> " °C"
    s = str(name)
    return " " + s[s.find("(") + 1 : s.rfind(")")].strip() if "(" in s and ")" in s else ""


default_title = (f"{_short(y_col)} vs. {_short(x_col)}" if layout == "long"
                 else f"Measurements vs. {_short(x_col)}")
title = l1.text_input("Slide title", value=default_title)
legend_title = l2.text_input("Legend title", value=_short(group_col) if group_col else "")
l3, l4, l5 = st.columns(3)
x_title = l3.text_input("X-axis title", value=x_col)
y_title = l4.text_input("Y-axis title", value=y_col if layout == "long" else "Y value (edit me)")
suffix = l5.text_input("Add to each line label (e.g. ' °C')", value=_unit(group_col) if group_col else "")
o1, o2 = st.columns(2)
smooth = o1.checkbox("Smooth lines", value=True)
markers = o2.checkbox("Show data-point markers", value=True)

# ---------------------------------------------------- build the series ---
try:
    if layout == "long":
        series = plotter.series_from_long(df, group_col, x_col, y_col, suffix)
    else:
        if not y_cols:
            st.warning("Pick at least one Y column.")
            st.stop()
        series = plotter.series_from_wide(df, x_col, y_cols, suffix)
except Exception as e:  # noqa: BLE001
    st.error(f"Couldn't build the chart: {e}")
    st.stop()

if len(series) > len(plotter.PALETTE):
    st.warning(f"{len(series)} lines: colors repeat after {len(plotter.PALETTE)}.")

# ------------------------------------------------------------- preview ---
st.subheader("Preview")
fig, ax = plt.subplots(figsize=(10, 5.2))
mk = ["o", "s", "^", "D", "x", "+", "*", "_"]
for i, s in enumerate(series):
    ax.plot(s["x"], s["y"], color="#" + plotter.PALETTE[i % len(plotter.PALETTE)],
            marker=mk[i % len(mk)] if markers else None, linewidth=2, label=s["name"])
ax.set_title(title, fontsize=15, fontweight="bold")
ax.set_xlabel(x_title, fontweight="bold")
ax.set_ylabel(y_title, fontweight="bold")
ax.grid(color="#D9D9D9")
ax.legend(title=legend_title or None, loc="center left", bbox_to_anchor=(1.01, 0.5), frameon=False)
fig.tight_layout()
st.pyplot(fig)
st.caption(f"{len(series)} lines · " + ", ".join(f"{s['name']} ({len(s['x'])} pts)" for s in series))

# ------------------------------------------------------------ download ---
pptx_bytes = plotter.build_pptx(series, title, x_title, y_title, legend_title,
                                smooth=smooth, show_markers=markers)
st.download_button(
    "⬇️ Download PowerPoint slide", data=pptx_bytes,
    file_name=(title or "chart").replace("/", "-")[:60] + ".pptx",
    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
    type="primary",
)
st.caption("The chart in the slide is a native PowerPoint chart: right-click it → Edit Data to see or change the numbers.")
