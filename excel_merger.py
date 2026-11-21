# -*- coding: utf-8 -*-
"""
Excel 合并工具
把指定文件夹内所有 Excel 文件的第一个 sheet 合并到一个文件中，
每个 sheet 以原文件名命名。

双击运行：弹出文件夹选择框，选好后自动合并，输出 合并结果.xlsx。
也可命令行运行：excel_merger.exe 输入文件夹 输出.xlsx
"""

import sys
import os
import glob
import traceback

SUPPORTED_EXTS = (".xlsx", ".xlsm", ".xls")


def pick_folder():
    """弹出文件夹选择对话框（Windows 用 tkinter，无需额外依赖）。"""
    import tkinter as tk
    from tkinter import filedialog
    root = tk.Tk()
    root.withdraw()
    folder = filedialog.askdirectory(title="请选择存放 Excel 文件的文件夹")
    root.destroy()
    return folder


def copy_sheet(src_ws, dst_ws):
    for row in src_ws.iter_rows():
        for cell in row:
            if cell.value is not None:
                dst_ws.cell(row=cell.row, column=cell.column, value=cell.value)
    for merged in src_ws.merged_cells.ranges:
        dst_ws.merge_cells(str(merged))
    for letter, dim in src_ws.column_dimensions.items():
        if dim.width:
            dst_ws.column_dimensions[letter].width = dim.width
    for idx, dim in src_ws.row_dimensions.items():
        if dim.height:
            dst_ws.row_dimensions[idx].height = dim.height


def sanitize_sheet_name(name, used_names):
    for ch in r"\/?*[]:":
        name = name.replace(ch, "_")
    name = name[:31] or "Sheet"
    base, i = name, 1
    while name in used_names:
        suffix = f"_{i}"
        name = base[: 31 - len(suffix)] + suffix
        i += 1
    used_names.add(name)
    return name


def merge_excels(input_dir, output_path, log=print):
    from openpyxl import Workbook, load_workbook

    files = sorted(
        f for f in glob.glob(os.path.join(input_dir, "*"))
        if f.lower().endswith(SUPPORTED_EXTS) and os.path.abspath(f) != os.path.abspath(output_path)
    )
    if not files:
        raise RuntimeError(f"未在该文件夹中找到 Excel 文件: {input_dir}")

    wb_out = Workbook()
    wb_out.remove(wb_out.active)
    used_names = set()
    ok, fail = 0, 0

    for f in files:
        try:
            wb_in = load_workbook(f, read_only=False, data_only=False)
            src_ws = wb_in.worksheets[0]  # 只取第一个 tab
            sheet_name = sanitize_sheet_name(os.path.splitext(os.path.basename(f))[0], used_names)
            dst_ws = wb_out.create_sheet(title=sheet_name)
            copy_sheet(src_ws, dst_ws)
            wb_in.close()
            ok += 1
            log(f"已合并: {os.path.basename(f)} -> [{sheet_name}]")
        except Exception as e:
            fail += 1
            log(f"跳过 {os.path.basename(f)}: {e}")

    wb_out.save(output_path)
    log(f"\n完成：成功 {ok} 个，失败 {fail} 个")
    log(f"输出文件: {output_path}")
    return output_path


def main():
    if len(sys.argv) > 1:
        input_dir = os.path.abspath(sys.argv[1])
        output_path = os.path.abspath(sys.argv[2]) if len(sys.argv) > 2 else os.path.join(input_dir, "合并结果.xlsx")
    else:
        input_dir = pick_folder()
        if not input_dir:
            print("已取消")
            return
        output_path = os.path.join(input_dir, "合并结果.xlsx")

    try:
        merge_excels(input_dir, output_path)
    except Exception:
        print(traceback.format_exc())
    finally:
        # 双击运行时窗口不要立刻关闭
        if len(sys.argv) <= 1:
            try:
                input("\n按回车键退出...")
            except EOFError:
                pass


if __name__ == "__main__":
    main()
