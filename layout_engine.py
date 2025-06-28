
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.patches as patches
from openpyxl import Workbook
import os
import math

def generate_layout_and_summary(csv_path, output_folder, sheet_width=48.0, sheet_height=96.0):
    df = pd.read_csv(csv_path)

    # Drop rows missing essential size fields
    df = df.dropna(subset=['Width', 'Height', 'Thickness'])

    # Group by thickness
    grouped = df.groupby('Thickness')

    layout_images = []

    for thickness, group in grouped:
        parts = []
        for _, row in group.iterrows():
            width, height = row['Width'], row['Height']
            quantity = int(row['Quantity']) if 'Quantity' in row and not pd.isnull(row['Quantity']) else 1
            for _ in range(quantity):
                parts.append({
                    'Width': width,
                    'Height': height,
                    'Part Name': row.get('Part Name', ''),
                    'Material': row.get('Material', '')
                })

        layouts = []
        current_sheet = []
        cursor_x = cursor_y = max_height = 0

        for part in parts:
            pw, ph = part['Width'], part['Height']
            if cursor_x + pw > sheet_width:
                cursor_x = 0
                cursor_y += max_height
                max_height = 0
            if cursor_y + ph > sheet_height:
                layouts.append(current_sheet)
                current_sheet = []
                cursor_x = cursor_y = max_height = 0
            current_sheet.append((cursor_x, cursor_y, pw, ph, part))
            cursor_x += pw
            max_height = max(max_height, ph)
        if current_sheet:
            layouts.append(current_sheet)

        for i, layout in enumerate(layouts):
            fig, ax = plt.subplots(figsize=(8.27, 11.69))  # A4 portrait
            ax.set_xlim(0, sheet_width)
            ax.set_ylim(0, sheet_height)
            ax.set_aspect('equal')
            ax.set_title(f'Sheet Layout - Thickness {thickness} - Sheet {i+1}')
            ax.invert_yaxis()

            for x, y, pw, ph, part in layout:
                rect = patches.Rectangle((x, y), pw, ph, linewidth=1, edgecolor='black', facecolor='lightgrey')
                ax.add_patch(rect)
                label = f"{part['Part Name']}\n{pw:.1f} x {ph:.1f}"
                ax.text(x + pw / 2, y + ph / 2, label, ha='center', va='center', fontsize=6, wrap=True)

            os.makedirs(output_folder, exist_ok=True)
            layout_path = os.path.join(output_folder, f'layout_thickness_{thickness}_sheet_{i+1}.png')
            plt.savefig(layout_path, bbox_inches='tight', dpi=150)
            layout_images.append(layout_path)
            plt.close()

    # Excel Summary
    wb = Workbook()
    ws = wb.active
    ws.title = "Summary"
    ws.append(["Part Name", "Width", "Height", "Thickness", "Material", "Quantity"])
    for _, row in df.iterrows():
        ws.append([
            row.get('Part Name', ''),
            row['Width'],
            row['Height'],
            row['Thickness'],
            row.get('Material', ''),
            int(row['Quantity']) if 'Quantity' in row and not pd.isnull(row['Quantity']) else 1
        ])
    summary_path = os.path.join(output_folder, "summary.xlsx")
    wb.save(summary_path)

    return layout_images, summary_path
