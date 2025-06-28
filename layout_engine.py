
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.patches as patches
from openpyxl import Workbook
import os

def generate_layout_and_summary(csv_path, output_folder, sheet_width=48.0, sheet_height=96.0):
    df = pd.read_csv(csv_path)

    # Normalize columns
    df.columns = [col.strip() for col in df.columns]
    df = df.rename(columns=lambda c: c.replace("Length", "Width") if "Length" in c else c)
    df = df.dropna(subset=["Width", "Height", "Thickness"])
    df["Width"] = df["Width"].astype(float)
    df["Height"] = df["Height"].astype(float)
    df["Thickness"] = df["Thickness"].astype(float)
    df["Quantity"] = df["Quantity"].fillna(1).astype(int)
    df["Part Name"] = df.get("Part Name", pd.Series(["Unnamed"] * len(df)))
    df["Material"] = df.get("Material", pd.Series(["Default"] * len(df)))

    df = df.sort_values(by=["Thickness", "Height", "Width", "Part Name"])

    os.makedirs(output_folder, exist_ok=True)
    summary_wb = Workbook()
    summary_ws = summary_wb.active
    summary_ws.title = "Cut Summary"
    summary_ws.append(["Part Name", "Width", "Height", "Thickness", "Material", "Quantity"])

    for thickness in sorted(df["Thickness"].unique()):
        parts = df[df["Thickness"] == thickness].copy()
        parts_list = []
        for _, row in parts.iterrows():
            for _ in range(row["Quantity"]):
                parts_list.append(row.to_dict())

        sheet_count = 0
        while parts_list:
            sheet_count += 1
            fig, ax = plt.subplots(figsize=(8.27, 11.69), dpi=300)
            ax.set_xlim(0, sheet_width)
            ax.set_ylim(0, sheet_height)
            ax.set_title(f"Thickness {thickness} - Sheet {sheet_count}")
            ax.set_aspect('equal')
            ax.axis("off")

            x_cursor, y_cursor = 0, 0
            row_height = 0
            placed_parts = []
            skipped_parts = []

            for part in parts_list:
                pw, ph = part["Width"], part["Height"]

                if x_cursor + pw > sheet_width:
                    x_cursor = 0
                    y_cursor += row_height
                    row_height = 0

                if y_cursor + ph > sheet_height:
                    skipped_parts.append(part)
                    continue

                rect = patches.Rectangle((x_cursor, y_cursor), pw, ph, linewidth=0.5, edgecolor='black', facecolor='lightgray')
                ax.add_patch(rect)
                label = f"{part['Part Name']}\n{pw}x{ph}""
                ax.text(x_cursor + pw / 2, y_cursor + ph / 2, label, ha='center', va='center', fontsize=4, wrap=True)
                x_cursor += pw
                row_height = max(row_height, ph)
                placed_parts.append(part)

            layout_filename = f"layout_thickness_{thickness}_sheet_{sheet_count}.png"
            layout_path = os.path.join(output_folder, layout_filename)
            fig.savefig(layout_path, bbox_inches='tight')
            plt.close(fig)

            for part in placed_parts:
                summary_ws.append([part["Part Name"], part["Width"], part["Height"],
                                   part["Thickness"], part["Material"], 1])

            parts_list = skipped_parts

    summary_path = os.path.join(output_folder, "cut_summary.xlsx")
    summary_wb.save(summary_path)
    print(f"Layout and summary saved to {output_folder}")
