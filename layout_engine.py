import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.patches as patches
from openpyxl import Workbook
import os
import math

def generate_layout_and_summary(csv_path, output_folder, sheet_width=48.0, sheet_height=96.0):
    df = pd.read_csv(csv_path)
    df = df.dropna(subset=['Width', 'Height'])

    df['Area'] = df['Width'] * df['Height']
    df = df.sort_values(by='Area', ascending=False)

    thickness_groups = df.groupby('Thickness')

    os.makedirs(output_folder, exist_ok=True)
    excel_path = os.path.join(output_folder, 'summary.xlsx')
    wb = Workbook()
    wb.remove(wb.active)

    for thickness, group in thickness_groups:
        group = group.reset_index(drop=True)
        fig, ax = plt.subplots(figsize=(8, 8 * sheet_height / sheet_width))
        ax.set_xlim(0, sheet_width)
        ax.set_ylim(0, sheet_height)
        ax.set_title(f"Layout for Thickness {thickness}\"")
        ax.set_xlabel("inches")
        ax.set_ylabel("inches")

        used_positions = []
        current_x = 0
        current_y = 0
        row_max_height = 0

        for i, row in group.iterrows():
            part_width = row['Width']
            part_height = row['Height']
            part_name = row['Part'] if 'Part' in row and pd.notna(row['Part']) else f"Part {i+1}"

            if current_x + part_width > sheet_width:
                current_x = 0
                current_y += row_max_height
                row_max_height = 0

            if current_y + part_height > sheet_height:
                continue

            rect = patches.Rectangle((current_x, current_y), part_width, part_height,
                                     linewidth=1, edgecolor='black', facecolor='lightgrey')
            ax.add_patch(rect)
            ax.text(current_x + part_width / 2, current_y + part_height / 2,
                    part_name, ha='center', va='center', fontsize=6)
            used_positions.append((current_x, current_y, part_width, part_height))

            current_x += part_width
            row_max_height = max(row_max_height, part_height)

        ax.set_aspect('equal')
        ax.invert_yaxis()
        png_filename = os.path.join(output_folder, f'layout_thickness_{thickness}.png')
        plt.savefig(png_filename, bbox_inches='tight')
        plt.close()

        ws = wb.create_sheet(title=f"{thickness}\"")
        ws.append(['Part', 'Width', 'Height', 'Quantity', 'Material'])
        for _, row in group.iterrows():
            ws.append([
                row['Part'] if 'Part' in row else f"Part {_ + 1}",
                row['Width'], row['Height'],
                row['Quantity'] if 'Quantity' in row else 1,
                row['Material'] if 'Material' in row else ''
            ])

    wb.save(excel_path)


