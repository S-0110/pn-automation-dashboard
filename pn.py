#!/usr/bin/env python
# coding: utf-8

# In[1]:



# In[3]:


import pandas as pd
import numpy as np
import re
from openpyxl import load_workbook
from openpyxl.styles import Font

# =====================================================
# FILE PATHS
# =====================================================

input_file = "1_to_10.xlsx"
output_file = "Campaign_Analysis_Output.xlsx"

# =====================================================
# LOAD RAW DATA
# =====================================================

raw_df = pd.read_excel(input_file)

# =====================================================
# REQUIRED COLUMNS
# =====================================================

required_columns = [
    'Campaign Name',
    'Campaign ID',
    'Channel',
    'Variant',
    'Title',
    'Message',
    'Start Date',
    'Start Time',
    'Device',
    'Who query',
    'Conversion Event',
    'Run Date',
    'Total Sent(users)',
    'Total Viewed(users)',
    'Total Clicked(users)',
    'Total Delivered(users)',
    'Click through conversions',
    'Click through conversions %',
    'Total control group count',
    'Total control group conversions',
    'Total control group conversions %',
    'Total control group revenue',
    'Influenced Conversions',
    'Influenced Conversions %',
    'Influenced Revenue'
]

# =====================================================
# CHECK MISSING COLUMNS
# =====================================================

missing_cols = [col for col in required_columns if col not in raw_df.columns]

if missing_cols:
    raise ValueError(
        f"Missing columns in input file:\n{missing_cols}"
    )

analysis_df = raw_df[required_columns].copy()

# =====================================================
# HANDLE NULLS
# =====================================================

numeric_cols = [
    'Total Sent(users)',
    'Total Viewed(users)',
    'Total Clicked(users)',
    'Total Delivered(users)',
    'Click through conversions',
    'Total control group count',
    'Total control group conversions',
    'Total control group revenue',
    'Influenced Conversions',
    'Influenced Revenue'
]

for col in numeric_cols:
    analysis_df[col] = (
        pd.to_numeric(
            analysis_df[col],
            errors='coerce'
        )
        .fillna(0)
    )

# =====================================================
# METRIC CALCULATIONS
# =====================================================

analysis_df['CTR'] = np.where(
    analysis_df['Total Viewed(users)'] > 0,
    analysis_df['Total Clicked(users)'] /
    analysis_df['Total Viewed(users)'],
    0
)

analysis_df['Total Converted'] = (
    analysis_df['Click through conversions']
    +
    analysis_df['Influenced Conversions']
)

analysis_df['TC%'] = np.where(
    analysis_df['Total Viewed(users)'] > 0,
    (
        analysis_df['Total Converted']
        /
        analysis_df['Total Viewed(users)']
    ) * 100,
    0
)

analysis_df['CTR%'] = np.where(
    analysis_df['Total Viewed(users)'] > 0,
    (
        analysis_df['Total Clicked(users)']
        /
        analysis_df['Total Viewed(users)']
    ) * 100,
    0
)

analysis_df['Total CG Conv %'] = np.where(
    analysis_df['Total control group count'] > 0,
    (
        analysis_df['Total control group conversions']
        /
        analysis_df['Total control group count']
    ) * 100,
    0
)

analysis_df['Sent Rate'] = np.where(
    analysis_df['Total Sent(users)'] > 0,
    analysis_df['Total Converted']
    /
    analysis_df['Total Sent(users)'],
    0
)

analysis_df['Control Rate'] = np.where(
    analysis_df['Total control group count'] > 0,
    analysis_df['Total control group conversions']
    /
    analysis_df['Total control group count'],
    0
)

analysis_df['Campaign Conversion Rate (CR)'] = np.where(
    analysis_df['Total Viewed(users)'] > 0,
    analysis_df['Total Converted']
    /
    analysis_df['Total Viewed(users)'],
    0
)

analysis_df['Relative Lift'] = np.where(
    analysis_df['Control Rate'] > 0,
    (
        (
            analysis_df['Campaign Conversion Rate (CR)']
            -
            analysis_df['Control Rate']
        )
        /
        analysis_df['Control Rate']
    ) * 100,
    0
)

analysis_df['Incremental Conversions'] = (
    (
        analysis_df['Campaign Conversion Rate (CR)']
        -
        analysis_df['Control Rate']
    )
    *
    analysis_df['Total Viewed(users)']
)

# =====================================================
# CATEGORY
# =====================================================

analysis_df['Category'] = (
    analysis_df['Campaign Name']
    .astype(str)
    .str.split('_')
    .str[0]
)

# =====================================================
# SEGMENT
# =====================================================

analysis_df['Segment'] = (
    analysis_df['Campaign Name']
    .astype(str)
    .str.split('_')
    .str[-1]
)

# =====================================================
# EXPORT TO EXCEL
# =====================================================

with pd.ExcelWriter(output_file, engine='openpyxl') as writer:

    # Analysis Sheet
    analysis_df.to_excel(
        writer,
        sheet_name='Analysis',
        index=False
    )

    # =================================================
    # CATEGORY SHEETS
    # =================================================

    categories = (
        analysis_df['Category']
        .dropna()
        .unique()
    )

    used_sheet_names = set()

    for category in categories:

        cat_df = analysis_df[
            analysis_df['Category'] == category
        ]

        sheet_name = str(category)

        # Remove invalid Excel characters
        sheet_name = re.sub(
            r'[:\\/*?\[\]]',
            '_',
            sheet_name
        )

        # Remove leading/trailing spaces
        sheet_name = sheet_name.strip()

        # Default if blank
        if not sheet_name:
            sheet_name = "Unknown"

        # Excel max length
        sheet_name = sheet_name[:31]

        # Handle duplicates
        original_name = sheet_name
        counter = 1

        while sheet_name in used_sheet_names:
            suffix = f"_{counter}"
            sheet_name = (
                original_name[:31-len(suffix)]
                + suffix
            )
            counter += 1

        used_sheet_names.add(sheet_name)

        cat_df.to_excel(
            writer,
            sheet_name=sheet_name,
            index=False
        )

    # =================================================
    # SUMMARY SHEET
    # =================================================

    summary_sheet = "Summary"

    ctr_pivot = pd.pivot_table(
        analysis_df,
        index=[
            'Campaign ID',
            'Title',
            'Message',
            'Variant'
        ],
        values='CTR%',
        aggfunc='mean'
    ).reset_index()

    tc_pivot = pd.pivot_table(
        analysis_df,
        index=[
            'Campaign ID',
            'Title',
            'Message',
            'Variant'
        ],
        values='TC%',
        aggfunc='mean'
    ).reset_index()

    segment_ctr = pd.pivot_table(
        analysis_df,
        index='Segment',
        values='CTR%',
        aggfunc='mean'
    ).reset_index()

    segment_tc = pd.pivot_table(
        analysis_df,
        index='Segment',
        values='TC%',
        aggfunc='mean'
    ).reset_index()

    start_row = 0

    ctr_pivot.to_excel(
        writer,
        sheet_name=summary_sheet,
        startrow=start_row,
        index=False
    )

    start_row += len(ctr_pivot) + 4

    tc_pivot.to_excel(
        writer,
        sheet_name=summary_sheet,
        startrow=start_row,
        index=False
    )

    start_row += len(tc_pivot) + 4

    segment_ctr.to_excel(
        writer,
        sheet_name=summary_sheet,
        startrow=start_row,
        index=False
    )

    start_row += len(segment_ctr) + 4

    segment_tc.to_excel(
        writer,
        sheet_name=summary_sheet,
        startrow=start_row,
        index=False
    )

# =====================================================
# FORMAT EXCEL
# =====================================================

wb = load_workbook(output_file)

for sheet in wb.sheetnames:

    ws = wb[sheet]

    # Bold first row
    for cell in ws[1]:
        cell.font = Font(bold=True)

    # Auto column width
    for column_cells in ws.columns:

        max_length = max(
            len(str(cell.value))
            if cell.value is not None
            else 0
            for cell in column_cells
        )

        adjusted_width = min(max_length + 5, 50)

        ws.column_dimensions[
            column_cells[0].column_letter
        ].width = adjusted_width

wb.save(output_file)

print("=" * 50)
print("Automation completed successfully!")
print(f"Output saved as: {output_file}")
print("=" * 50)


# In[ ]:




