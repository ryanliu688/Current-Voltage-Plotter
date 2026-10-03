# Spreadsheet → PowerPoint Plotter
[![Open in Streamlit](https://static.streamlit.io/badges/streamlit_badge_black_white.svg)](https://current-voltage-plotter.streamlit.app)

Upload an Excel/CSV file with X/Y data for several trials (e.g. voltage vs. current
at different temperatures) and download a slide with a labeled, multi-color,
**editable** PowerPoint chart.

## Run it on your computer
1. Install Python 3.10+ from python.org.
2. Open a terminal in this folder and run:
   ```
   pip install -r requirements.txt
   streamlit run app.py
   ```
3. Your browser opens the app. Upload a sample Excel spreadsheet to try it.

## Files
- `app.py`: the web page (upload, column pickers, labels, preview, download)
- `plotter.py`: reads the spreadsheet and builds the .pptx (no web code; reusable)
- `make_sample_data.py`: generates `sample_iv_data.xlsx` (non-ohmic resistor at 25/50/75 °C)
- `example_output.pptx`: the slide the app produced from the sample data

## Spreadsheet layouts
- **Long** (recommended): columns like `Temperature (°C) | Voltage (V) | Current (mA)`,
  one row per measurement. Each trial may use different X values.
- **Wide**: `Voltage (V) | 25 °C | 50 °C | 75 °C`, one shared X column.

Units in parentheses in the label column's header (e.g. `(°C)`) are added to the
legend labels automatically.
