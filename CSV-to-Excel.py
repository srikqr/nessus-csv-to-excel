import pandas as pd
import numpy as np
import os
import sys
import glob

def install_missing_packages():
    import subprocess
    import pkg_resources
    required = {'pandas', 'openpyxl', 'xlsxwriter', 'numpy'}
    installed = {pkg.key for pkg in pkg_resources.working_set}
    missing = required - installed
    if missing:
        print(f"Installing missing packages: {missing}")
        subprocess.check_call([sys.executable, '-m', 'pip', 'install', *missing])

try:
    import pandas as pd
    import numpy as np
except ImportError:
    install_missing_packages()
    import pandas as pd
    import numpy as np

def get_column_map(df):
    col_map = {}
    cols_lower = {c.lower(): c for c in df.columns}
    mapping = {
        'plugin id': 'Plugin ID',
        'cvss score': 'CVSS Score',
        'risk': 'Risk',
        'host': 'Host',
        'protocol': 'Protocol',
        'port': 'Port',
        'name': 'Name',
        'synopsis': 'Synopsis',
        'description': 'Description',
        'solution': 'Solution',
        'plugin output': 'Plugin Output'
    }
    for lower_key, proper_key in mapping.items():
        col_map[proper_key] = cols_lower.get(lower_key, None)
    return col_map

def process_df(df):
    col_map = get_column_map(df)

    for key in col_map:
        if col_map[key] is None:
            df[key] = ''
            col_map[key] = key

    risk_col = df[col_map['Risk']].fillna('').astype(str).str.strip().str.lower()
    df_filtered = df[~risk_col.isin(['', 'none'])].copy()

    grouped_hosts = (
        df_filtered
        .groupby(col_map['Name'], dropna=False, observed=False, as_index=True)
        .agg({col_map['Host']: lambda x: x.astype(str).unique().tolist(),
              col_map['Protocol']: lambda x: x.astype(str).unique().tolist(),
              col_map['Port']: lambda x: x.astype(str).unique().tolist()})
    )

    def combine_hosts(row):
        hosts = row[col_map['Host']]
        protocols = row[col_map['Protocol']]
        ports = row[col_map['Port']]
        combined = set()
        for h in hosts:
            for p in protocols:
                for po in ports:
                    combined.add(f"{h} {p} {po}")
        return ", ".join(sorted(combined))

    grouped_hosts['Host Details'] = grouped_hosts.apply(combine_hosts, axis=1)

    unique_indices = df_filtered.drop_duplicates(subset=[col_map['Name']]).index

    df_final = pd.DataFrame(columns=[
        'Vulnerability Name', 'Host Details', 'Risk Rating', 'CVSS Base Score',
        'Status', 'Synopsis', 'Description', 'Recommendation', 'Vulnerability Result',
        'Remarks', 'Client Remarks'
    ])

    for i, idx in enumerate(unique_indices):
        vuln_name = df_filtered.at[idx, col_map['Name']]
        df_final.at[i, 'Vulnerability Name'] = vuln_name
        df_final.at[i, 'Host Details'] = grouped_hosts.at[vuln_name, 'Host Details'] if vuln_name in grouped_hosts.index else ''
        df_final.at[i, 'Risk Rating'] = df_filtered.at[idx, col_map['Risk']]
        df_final.at[i, 'CVSS Base Score'] = df_filtered.at[idx, col_map['CVSS Score']]
        df_final.at[i, 'Status'] = "Open"
        df_final.at[i, 'Synopsis'] = df_filtered.at[idx, col_map['Synopsis']]
        df_final.at[i, 'Description'] = df_filtered.at[idx, col_map['Description']]
        df_final.at[i, 'Recommendation'] = df_filtered.at[idx, col_map['Solution']]
        df_final.at[i, 'Vulnerability Result'] = df_filtered.at[idx, col_map['Plugin Output']]
        df_final.at[i, 'Remarks'] = ''
        df_final.at[i, 'Client Remarks'] = ''

    df_final.index += 1
    df_final.index.name = 'Serial No.'
    df_final.fillna('', inplace=True)

    return df_final

def format_worksheet(writer, sheet_name, df):
    workbook = writer.book
    worksheet = writer.sheets[sheet_name]

    header_format = workbook.add_format({
        'bold': True,
        'text_wrap': True,
        'valign': 'top',
        'fg_color': '#D9EAD3',
        'border': 1
    })
    cell_format = workbook.add_format({
        'text_wrap': True,
        'valign': 'top',
        'border': 1
    })

    for col_num, value in enumerate(df.columns):
        worksheet.write(0, col_num + 1, value, header_format)
    worksheet.write(0, 0, df.index.name or 'Index', header_format)

    # Set all column widths to 50
    worksheet.set_column(0, len(df.columns), 15)

    # Set row height to 50 for all rows including header
    worksheet.set_default_row(20, False)

    # Write data cells with wrap and border
    for row in range(1, len(df) + 1):
        worksheet.write(row, 0, df.index[row - 1], cell_format)
        for col in range(len(df.columns)):
            val = df.iloc[row - 1, col]
            worksheet.write(row, col + 1, val, cell_format)

    worksheet.autofilter(0, 0, len(df), len(df.columns))

def save_to_excel_all_in_one(df_final, output_path):
    writer = pd.ExcelWriter(output_path, engine='xlsxwriter')

    df_final.to_excel(writer, sheet_name='Full Report', index=True)
    format_worksheet(writer, 'Full Report', df_final)

    risk_order = ['Critical', 'High', 'Medium', 'Low']

    for risk in risk_order:
        df_risk = df_final[df_final['Risk Rating'].str.lower() == risk.lower()]
        if df_risk.empty:
            continue
        sheet_name = risk[:31]
        df_risk.to_excel(writer, sheet_name=sheet_name, index=True)
        format_worksheet(writer, sheet_name, df_risk)

    writer.close()
    print(f"Excel output saved to {output_path}")

def main(files):
    if not files:
        print("Usage: python script.py <file1.csv> [file2.csv ...] or *.csv")
        sys.exit(1)

    if not os.path.exists('final'):
        os.mkdir('final')

    for file in files:
        for filename in glob.glob(file):
            print(f"Processing {filename} ...")
            df = pd.read_csv(filename)
            df_final = process_df(df)
            base = os.path.splitext(os.path.basename(filename))[0]

            out_xlsx = os.path.join('final', f"{base}_final_report.xlsx")
            save_to_excel_all_in_one(df_final, out_xlsx)

if __name__ == "__main__":
    install_missing_packages()
    main(sys.argv[1:])
