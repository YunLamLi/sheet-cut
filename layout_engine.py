import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.patches as patches
from openpyxl import Workbook
import os
import math
from collections import defaultdict

def generate_layout_and_summary(csv_path, output_folder, sheet_width=48.0, sheet_height=96.0):
    df = pd.read_csv(csv_path)

    # Drop rows missing width/height
    df = df.dropna(subset=['Width', 'Height'])
    
    # Normalize headers
    df.columns = [col.strip() for col in df.columns]
    
    # Default columns if missing
    if 'Quantity' not in df.columns:
        df['Quantity'] = 1
    if 'Part Name' not in df.columns:
        df['Part Name'] = ['Part-' + str(i) for i in range(len(df))]
    if 'Material' not in df.columns:
        df['Material'] = 'Default'

    # Expand by quantity
    parts = []
    for _, row in df.iterrows():
        for _ in range(int(row['Quantity'])):
            parts.append({
                'Part Name': row['Part Name'],
                'Width': float(row['Width']),
                'Height': float(row['Height']),
                'Thickness': float(row['Thickness']),
                'Material': row['Material']
            })

    # Group by thickness
    grouped_parts = defaultdict(list)
    for part in parts:
        grouped_parts[part['Thickness']].append(part)

    os.makedirs(output_folder, exist_ok=True)

    # Excel summary
    wb = Workbook()
    summary_ws = wb.active
    summary_ws.title = "Part Summary"
    summary_ws.append(["Part Name", "Width", "Height", "Thickness", "Material", "Quantity"])

    summary = df.groupby(['Part Name', 'Width', 'Height', 'Thickness', 'Material']).agg({'Quantity': 'sum'}).reset_index()
    for _, row in summary.iterrows():
        summary_ws.append(list(row))

    # Layout drawing
    for thickness, parts_list in grouped_parts.items():
        # Sort by area descending, then width
        parts_list.sort(key=lambda x: (-x['Height'] * x['Width'], -x['Width']))

        sheet_num = 1
        placed_parts = []
        unplaced = parts_list.copy()

        while unplaced:
            fig, ax = plt.subplots(figsize=(8.27, 11.69))  # A4 portrait
            ax.set_xlim(0, sheet_width)
            ax.set_ylim(0, sheet_height)
            ax.set_title(f"Sheet Layout - Thickness {thickness} - Sheet {sheet_num}")
            ax.invert_yaxis()

            x_cursor, y_cursor = 0, 0
            row_height = 0
            remaining_parts = []

            for part in unplaced:
                pw, ph = part['Width'], part['Height']

                # Fit in current row
                if x_cursor + pw <= sheet_width and y_cursor + ph <= sheet_height:
                    rect = patches.Rectangle((x_cursor, y_cursor), pw, ph, linewidth=1, edgecolor='black', facecolor='lightgray')
                    ax.add_patch(rect)

                    label = f"{part['Part Name']}\n{pw:.1f} x {ph:.1f}"
                    ax.text(x_cursor + pw/2, y_cursor + ph/2, label, ha='center', va='center', fontsize=6)

                    x_cursor += pw
                    row_height = max(row_height, ph)
                    placed_parts.append(part)
                else:
                    remaining_parts.append(part)

                # If row full, wrap
                if x_cursor + pw > sheet_width:
                    x_cursor = 0
                    y_cursor += row_height
                    row_height = 0

            filename = os.path.join(output_folder, f"layout_thickness_{thickness}_sheet_{sheet_num}.png")
            plt.tight_layout()
            plt.savefig(filename, dpi=300)
            plt.close()
            sheet_num += 1
            unplaced = remaining_parts

    excel_path = os.path.join(output_folder, "cutlist_summary.xlsx")
    wb.save(excel_path)
