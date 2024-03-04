from flask import Flask, render_template
import pandas as pd
import os

app = Flask(__name__)

def get_last_updated_time():
    file_path = "BSNL-UCs-LDH-RPR.xlsx"
    if os.path.exists(file_path):
        return os.path.getmtime(file_path)
    else:
        return None

def read_excel_data():
    file_path = "BSNL-UCs-LDH-RPR.xlsx"
    if os.path.exists(file_path):
        df = pd.read_excel(file_path)
        return df
    else:
        return None

@app.route('/')
def index():
    df = read_excel_data()
    row_count = len(df) if df is not None else 0
    ldh_count = len(df[df['SSA'] == 'LDH']) if df is not None else 0
    rpr_count = len(df[df['SSA'] == 'RPR']) if df is not None else 0

    last_updated_time = get_last_updated_time()
    if last_updated_time:
        last_updated_str = pd.to_datetime(last_updated_time, unit='s').strftime("%d-%B %Y %H:%M:%S IST")
    else:
        last_updated_str = "Unknown"

    header_text = f"{row_count} PENDING BSNL UCs - MCN, {ldh_count} - LUDHIANA, {rpr_count} - ROPAR"

    compact_df = df[['Phone', 'Customer Name', 'SSA', 'BBC Approved']] if df is not None else pd.DataFrame()
    detailed_df = df[['Phone', 'Customer Name', 'SSA', 'Mobile', 'InstallDate', 'Block', 'GP', 'Address', 'UC Received', 'BBC Approved']] if df is not None else pd.DataFrame()

    compact_data = compact_df.to_html(index=False, escape=False, classes="table")
    detailed_data = detailed_df.to_html(index=False, escape=False, classes="table")

    return render_template('BSNL-UC-index.html', last_updated_time=last_updated_str, header_text=header_text, compact_data=compact_data, detailed_data=detailed_data)

if __name__ == '__main__':
    app.run(port=4560)
