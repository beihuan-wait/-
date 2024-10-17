import pandas as pd

# 配置部分
file_path = r'D:\Python Project\chonfuxiangbidui\1.xlsx'  # 替换为实际的文件路径
sheet_name = 'Sheet1'  # 替换为实际的工作表名称

# 需要检查的列名
check_column = '平台ID'  # 替换为 A 列的列名
compare_column = '小红书号'  # 替换为 B 列的列名

output_file_path = r'D:\Python Project\chonfuxiangbidui\结果.xlsx'  # 替换为你希望保存的文件路径

def read_data(file_path, sheet_name):
    """读取 Excel 数据，并返回 DataFrame"""
    try:
        return pd.read_excel(file_path, sheet_name=sheet_name)
    except Exception as e:
        print(f"读取文件时发生错误: {e}")
        raise

def preprocess_data(df, columns):
    """将指定列的数据转换为文本格式，并去除两端空白字符"""
    for col in columns:
        df[col] = df[col].astype(str).str.strip()
    return df

def validate_columns(df, columns):
    """确保 DataFrame 包含指定的列名"""
    missing_cols = [col for col in columns if col not in df.columns]
    if missing_cols:
        raise ValueError(f"列名 {', '.join(missing_cols)} 不存在于数据中，请检查列名。")

def create_compare_set(df, column):
    """创建用于比较的集合，区分大小写"""
    return set(df[column])

def is_duplicate(value, compare_set):
    """检查值是否在比较集合中"""
    return '重复' if value in compare_set else '不重复'

def main():
    # 读取 Excel 文件
    df = read_data(file_path, sheet_name)

    # 确保列名在 DataFrame 中存在
    validate_columns(df, [check_column, compare_column])

    # 将所有数据转换为文本格式，确保区分大小写和字符类型
    df = preprocess_data(df, [check_column, compare_column])

    # 打印转换后的数据以检查数据格式
    print(f"{check_column} 列的数据预览：")
    print(df[check_column].head(20))  # 增加输出的行数以检查数据
    print(f"{compare_column} 列的数据预览：")
    print(df[compare_column].head(20))  # 增加输出的行数以检查数据

    # 创建一个集合来存储 B 列中的所有值（区分大小写）
    compare_set = create_compare_set(df, compare_column)
    print("B 列的唯一值集合预览：")
    print(compare_set)

    # 检查 A 列中的每个值是否存在于 B 列的集合中，区分大小写
    df['标记'] = df[check_column].apply(lambda x: is_duplicate(x, compare_set))

    # 打印出标记为“重复”的数据
    print("标记为“重复”的数据：")
    print(df[df['标记'] == '重复'])

    # 将结果保存到新的 Excel 文件
    try:
        df.to_excel(output_file_path, index=False)
        print(f"处理结果已保存到 {output_file_path}")
    except Exception as e:
        print(f"保存文件时发生错误: {e}")

if __name__ == "__main__":
    main()
