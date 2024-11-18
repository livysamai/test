import pandas as pd
import os
import sys

def copy_column_a_to_b(input_file, output_file):
    try:
        # 读取Excel文件，没有标题行
        df = pd.read_excel(input_file, header=None)
        print("Excel File Loaded Successfully")
        print(df.head())  # 查看加载的数据

        # 动态生成列名
        df.columns = ['A', 'B'] + [f"Col{i}" for i in range(2, df.shape[1])]
        print("Updated Columns:", df.columns)

        # 确保A列存在
        if 'A' not in df.columns:
            print("Error: Column 'A' not found in the file.")
            return

        # 将A列的内容复制到B列
        df['B'] = df['A']
        print("Column 'A' copied to 'B'.")
        print(df)  # 打印修改后的 DataFrame

        # 检查输出目录是否存在
        output_dir = os.path.dirname(output_file)
        if not os.path.exists(output_dir):
            print(f"Error: Output directory does not exist: {output_dir}")
            return

        # 保存到新的Excel文件
        df.to_excel(output_file, index=False)
        print(f"File saved successfully to {output_file}")

    except Exception as e:
        print(f"Error occurred: {e}")

if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Usage: python test.py <input_file> <output_file>")
    else:
        input_file = sys.argv[1]
        output_file = sys.argv[2]
        copy_column_a_to_b(input_file, output_file)
