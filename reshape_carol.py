import pandas as pd
import os

INPUT_FILE  = "Template_from_Carol.xlsx"
OUTPUT_FILE = "Carol_Clean.xlsx"

TOTAL_ROW_CATEGORIES = {
    'Total Personnel':       'Personnel',
    'Total Fringe benefits':  'Fringe Benefits',
    'Total Fringe Benefits':  'Fringe Benefits',
    'Total Travel':          'Travel',
    'Total Supplies':        'Supplies',
    'Total Contractor':      'Contractor',
    'Total Other':           'Other',
}

LINE_ITEM_CATEGORIES = {
    'Personnel': ['AOR Executive Director','Outreach Coordinator (contract)','Program Director'],
    'Fringe Benefits': ['Workers Comp','Workers Comp & FICA','FICA'],
    'Travel': ['Conference - 2 peo[le (CADCA?)','CADCA Conference','Mileage'],
    'Supplies': ['Coalition support materials','Social media w/printed materials',
                 'Youth classes & event supplies','Coping toolboxes',
                 'Forest Therapy supplies','General office supplies'],
    'Contractor': ['IBH','Alliance - Narcan Training','Synar Volunteers',
                   'Teambuilding & Leadership workshops','TBD Video/PSA'],
    'Other': ['CADCA membership','Prevention License','Waterford Chamber membership',
              'Chamber business profile','Mich SOS','T-Mobile','ADP',
              'Buffer, Mailchimp, Website, paddlenet (songer)','Chat GPT','Insurance',
              'Rent','Office Space','Venues','Deterra bags','Lock boxes',
              'Marketing Instigators','Campaign & event support','Meals & Entertainment',
              'Alliance student survey','Storage Unit','Gulla CPA','Volunteers & In-Kind'],
}

SECTION_HEADERS = {'Personnel','Fringe benefits','Fringe Benefits','Travel',
                   'Supplies','Contractor','Other','GRAND TOTALS'}
SUMMARY_LABELS  = {'Year total','Budgeted','Remaining','% of grant used','% of year gone'}
USE_TOTAL_ONLY  = {'Personnel','Fringe Benefits'}


def process_sheet(sheet_name, fiscal_year):
    xl       = pd.read_excel(INPUT_FILE, sheet_name=sheet_name, header=None)
    programs = xl.iloc[4, 1:].tolist()  # offset by 1 (col 0 = line item label)
    months   = xl.iloc[5, 1:].tolist()
    data     = xl.iloc[6:, :]

    # monthly_cols: (offset_index, program, month_str)
    # offset_index is the index into programs/months lists (0-based)
    # actual row.iloc index = offset_index + 1
    monthly_cols = []
    for i, (prog, month) in enumerate(zip(programs, months)):
        if prog not in ('DFC','ACHC') or pd.isna(month):
            continue
        if isinstance(month, str) and month in SUMMARY_LABELS:
            continue
        try:
            monthly_cols.append((i, prog, pd.to_datetime(month).strftime('%b-%y')))
        except Exception:
            continue

    # year_total_cols: program → offset_index
    year_total_cols = {}
    for i, (prog, month) in enumerate(zip(programs, months)):
        if isinstance(month, str) and 'Year total' in month and prog in ('DFC','ACHC'):
            year_total_cols[prog] = i

    months_per_prog = {}
    for _, prog, _ in monthly_cols:
        months_per_prog[prog] = months_per_prog.get(prog, 0) + 1

    rows = []
    for _, row in data.iterrows():
        label = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else None
        if not label or label == 'nan' or label in SECTION_HEADERS:
            continue

        # ── Handle Total rows (Personnel & Fringe Benefits) ──────────────
        if label in TOTAL_ROW_CATEGORIES:
            category = TOTAL_ROW_CATEGORIES[label]
            if category in USE_TOTAL_ONLY:
                for prog, offset_idx in year_total_cols.items():
                    # row.iloc[0] = line item label, so data starts at iloc[1]
                    # offset_idx is 0-based into programs list → row.iloc[offset_idx + 1]
                    yr_val = row.iloc[offset_idx + 1]
                    yr_val = float(yr_val) if pd.notna(yr_val) else 0.0
                    n = months_per_prog.get(prog, 12)
                    monthly_val = yr_val / n if n else 0.0
                    for _, mprog, month_str in monthly_cols:
                        if mprog == prog:
                            rows.append({
                                'Fiscal Year': fiscal_year,
                                'Category':    category,
                                'Line Item':   category,
                                'Program':     prog,
                                'Month':       month_str,
                                'Amount':      monthly_val,
                            })
            continue  # skip all other Total* rows

        if label.startswith('Total'):
            continue

        # ── Individual line items ─────────────────────────────────────────
        category = 'Other'
        for cat, items in LINE_ITEM_CATEGORIES.items():
            if label in items:
                category = cat
                break
        if category in USE_TOTAL_ONLY:
            continue

        for offset_idx, prog, month_str in monthly_cols:
            val = row.iloc[offset_idx + 1]
            rows.append({
                'Fiscal Year': fiscal_year,
                'Category':    category,
                'Line Item':   label,
                'Program':     prog,
                'Month':       month_str,
                'Amount':      float(val) if pd.notna(val) else 0.0,
            })

    return pd.DataFrame(rows)


if __name__ == '__main__':
    if not os.path.exists(INPUT_FILE):
        print(f"ERROR: Cannot find '{INPUT_FILE}'")
    else:
        print("Processing 2024-2025...")
        df1 = process_sheet('2024-2025', '2024-2025')
        print(f"  {len(df1)} rows")

        print("Processing 2025-2026...")
        df2 = process_sheet('2025-2026', '2025-2026')
        print(f"  {len(df2)} rows")

        combined = pd.concat([df1, df2], ignore_index=True)
        print(f"\nTotal: {len(combined)} rows")

        print("\nCategory totals (2024-2025, DFC):")
        check = combined[
            (combined['Fiscal Year'] == '2024-2025') & (combined['Program'] == 'DFC')
        ].groupby('Category')['Amount'].sum().sort_values(ascending=False)
        for cat, val in check.items():
            print(f"  {cat}: ${val:,.2f}")

        combined.to_excel(OUTPUT_FILE, index=False)
        print(f"\nSaved to '{OUTPUT_FILE}' — ready for Tableau!")
