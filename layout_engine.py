
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.patches as patches
from openpyxl import Workbook
import os
import math

def generate_layout_and_summary(csv_path, output_folder, sheet_width=48.0, sheet_height=96.0):
    df = pd.read_csv(csv_path)

    df = df.dropna(subset=['Width', 'Height'])
    df['Width'] = df['Width'].astype(float)
    df['Height'] = df['Height'].astype(float)
    df['Thickness'] = df['Thickness'].astype(float)
    df['Quantity'] = df['Quantity'].astype(int)

    grouped = df.groupby('Thickness')
    layout_files = []

    for thickness, group in grouped:
        parts = []
        for _, row in group.iterrows():
            for _ in range(row['Quantity']):
                parts.append({
                    'Part Name': row['Part Name'],
                    'Width': row['Width'],
                    'Height': row['Height'],
                    'Material': row['Material']
                })

        parts.sort(key=lambda x: max(x['Width'], x['Height']), reverse=True)

        layouts = []
        current_sheet = []
        remaining_height = sheet_height
        y_offset = 0
        x_offset = 0
        max_row_height = 0

        for part in parts:
            pw, ph = part['Width'], part['Height']
            if pw > sheet_width or ph > sheet_height:
                continue

            if x_offset + pw > sheet_width:
                x_offset = 0
                y_offset += max_row_height
                max_row_height = 0

            if y_offset + ph > sheet_height:
                layouts.append(current_sheet)
                current_sheet = []
                x_offset = 0
                y_offset = 0
                max_row_height = 0

            current_sheet.append((x_offset, y_offset, pw, ph, part))
            x_offset += pw
            max_row_height = max(max_row_height, ph)

        if current_sheet:
            layouts.append(current_sheet)

        for i, layout in enumerate(layouts):
            fig, ax = plt.subplots(figsize=(8, 4))
            ax.set_xlim(0, sheet_width)
            ax.set_ylim(0, sheet_height)
            ax.set_aspect('equal')
            ax.set_title(f'Sheet Layout - Thickness {thickness} - Sheet {i+1}')
            ax.invert_yaxis()

            for x, y, pw, ph, part in layout:
                rect = patches.Rectangle((x, y), pw, ph, linewidth=1, edgecolor='black', facecolor='lightgrey')
                ax.add_patch(rect)
                label = f"{part['Part Name']}
{pw:.1f} x {ph:.1f}"
                ax.text(x + pw / 2, y + ph / 2, label, ha='center', va='center', fontsize=7, wrap=True)

            os.makedirs(output_folder, exist_ok=True)
            layout_path = os.path.join(output_folder, f'layout_thickness_{thickness}_sheet_{i+1}.png')
            plt.savefig(layout_path, bbox_inches='tight')
            plt.close()
            layout_files.append(layout_path)

    wb = Workbook()
    ws = wb.active
    ws.title = 'Cut Summary'
    ws.append(['Part Name', 'Width', 'Height', 'Thickness', 'Material', 'Quantity'])

    for _, row in df.iterrows():
        ws.append([row['Part Name'], row['Width'], row['Height'], row['Thickness'], row['Material'], row['Quantity']])

    summary_path = os.path.join(output_folder, 'cut_summary.xlsx')
    wb.save(summary_path)

    return layout_files, summary_path

