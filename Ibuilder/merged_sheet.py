import pandas as pd
import sys
import os

# 合并至一页
def merge_specific_sheets(file_list, sheet_name, output_file):
    try:
        # 检查是否有文件需要合并
        if not file_list or len(file_list) < 2:
            print("Error: At least two files are required for merging.", file=sys.stderr)
            return

        # 初始化一个空的 DataFrame
        merged_data = pd.DataFrame()

        # 遍历文件列表，将每个文件中指定 sheet 的数据加载并追加到 merged_data
        for file_path in file_list:
            if not os.path.exists(file_path):
                print(f"Error: File does not exist: {file_path}", file=sys.stderr)
                return

            print(f"Loading file: {file_path}, Sheet: {sheet_name}")
            try:
                # 读取指定的 sheet
                data = pd.read_excel(file_path, sheet_name=sheet_name)
                merged_data = pd.concat([merged_data, data], ignore_index=True)
            except ValueError as e:
                print(f"Error: Could not load sheet '{sheet_name}' from {file_path}: {e}", file=sys.stderr)
                return

        # 确保输出文件目录存在
        output_dir = os.path.dirname(output_file)
        if output_dir and not os.path.exists(output_dir):
            print(f"Error: Output directory does not exist: {output_dir}", file=sys.stderr)
            return

        # 保存合并后的数据到 Excel 文件
        merged_data.to_excel(output_file, index=False)
        print(f"Merge completed successfully. Output saved to: {output_file}")

    except Exception as e:
        print(f"Error occurred during merging: {e}", file=sys.stderr)

if __name__ == "__main__":
    if len(sys.argv) < 4:
        print("Usage: python merge_excel.py <sheet_name> <file1> <file2> ... <output_file>", file=sys.stderr)
    else:
        sheet_name, *input_files, output_file = sys.argv[1:]
        merge_specific_sheets(input_files, sheet_name, output_file)
