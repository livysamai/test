import sys
import pandas as pd
from openpyxl import Workbook

# 将多个 Excel 文件的内容分别写入一个新的 Excel 文件的不同 Sheet 页中
def merge_files_to_sheets(input_files, output_file):
    wb = Workbook()
    wb.remove(wb.active)  # 删除默认的 Sheet 页

    for file in input_files:
        try:
            # 读取 Excel 文件
            xls = pd.ExcelFile(file)
            for sheet_name in xls.sheet_names:
                # 将每个 sheet 读取为 DataFrame
                df = pd.read_excel(file, sheet_name=sheet_name,header=None)
                sheet_title = f"{file.split('/')[-1].split('.')[0]}_{sheet_name}"[:31]  # 确保 Sheet 名称不超过 31 字符
                ws = wb.create_sheet(sheet_title)

                # 写入 DataFrame 数据到 Sheet
                for row in df.itertuples(index=False, name=None):
                    ws.append(row)
        except Exception as e:
            print(f"Error processing file {file}: {e}", file=sys.stderr)
            continue

    # 保存合并后的文件
    wb.save(output_file)
    print(f"合并完成，保存到: {output_file}")

if __name__ == "__main__":
    # 从命令行接收文件路径
    input_files = sys.argv[1:-1]
    output_file = sys.argv[-1]
    if not input_files or not output_file:
        print("请提供输入文件和输出文件路径", file=sys.stderr)
        sys.exit(1)

    merge_files_to_sheets(input_files, output_file)
