import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.patches as patches
from openpyxl import Workbook
import os
import math

def generate_layout_and_summary(csv_path, output_folder, sheet_width=48.0, sheet_height=96.0):
    df = pd.read_csv(csv_path)

    # Drop rows with missing required dimensions
    df = df.dropna(subset=['Width', 'Height', 'Thickness'])

    # Group by thickness
    grouped = df.groupby('Thickness')

    summary_data = []

    for thickness, group in grouped:
        parts = group.to_dict(orient='records')
        sheet_num = 1
        part_index = 0
        max_parts = len(parts)

        while part_index < max_parts:
            for strategy in ['row', 'column']:
                fig, ax = plt.subplots(figsize=(8.27, 11.69))  # A4 portrait in inches

                ax.set_xlim(0, sheet_width)
                ax.set_ylim(0, sheet_height)
                ax.set_title(f'Thickness {thickness} - Sheet {sheet_num} ({strategy}-wise)')
                ax.set_aspect('equal')
                ax.axis('off')

                placed_parts = 0
                padding = 0.25  # 1/4 inch kerf margin

                x_cursor = y_cursor = row_height = col_width = 0
                layout_success = False

                for i in range(part_index, max_parts):
                    part = parts[i]
                    pw = part['Width']
                    ph = part['Height']

                    if strategy == 'row':
                        if x_cursor + pw > sheet_width:
                            x_cursor = 0
                            y_cursor += row_height + padding
                            row_height = 0
                        if y_cursor + ph > sheet_height:
                            break
                        rect_x, rect_y = x_cursor, y_cursor
                        x_cursor += pw + padding
                        row_height = max(row_height, ph)

                    else:  # column strategy
                        if y_cursor + ph > sheet_height:
                            y_cursor = 0
                            x_cursor += col_width + padding
                            col_width = 0
                        if x_cursor + pw > sheet_width:
                            break
                        rect_x, rect_y = x_cursor, y_cursor
                        y_cursor += ph + padding
                        col_width = max(col_width, pw)

                    rect = patches.Rectangle((rect_x, rect_y), pw, ph, edgecolor='black', facecolor='none')
                    ax.add_patch(rect)
                    label = f"{part['Part Name']}\n{pw}x{ph}"
                    ax.text(rect_x + pw/2, rect_y + ph/2, label,
                            fontsize=6, ha='center', va='center')
                    placed_parts += 1

                if placed_parts == 0:
                    continue

                filename = f"layout_thickness_{thickness}_sheet_{sheet_num}_{strategy}.png"
                plt.savefig(os.path.join(output_folder, filename), bbox_inches='tight', dpi=300)
                plt.close(fig)

            part_index += placed_parts
            sheet_num += 1

        for _, row in group.iterrows():
            summary_data.append({
                'Part Name': row['Part Name'],
                'Width': row['Width'],
                'Height': row['Height'],
                'Thickness': thickness,
                'Material': row.get('Material', 'Default'),
                'Quantity': row.get('Quantity', 1)
            })

    wb = Workbook()
    ws = wb.active
    ws.title = 'Cut Summary'
    headers = ['Part Name', 'Width', 'Height', 'Thickness', 'Material', 'Quantity']
    ws.append(headers)
    for row in summary_data:
        ws.append([row.get(h, '') for h in headers])

    wb.save(os.path.join(output_folder, 'cut_summary.xlsx'))
