import pandas as pd
import os

# 定义根目录
root_dir = '/Users/caiwenshuo/Desktop/大数据公司实习/关于报送地方征信平台1-6月融资数据明细的通知-20240711163400'

# 定义要提取的列
required_columns = [
    '地方平台编码', '获贷企业名称', '统一社会信用代码', '获贷时间', 
    '是否为首贷', '放贷机构', '放贷机构统一社会信用代码'
]

# 定义列类型
dtype = {
    '统一社会信用代码': str,
    '放贷机构统一社会信用代码': str
}

# 创建一个空的DataFrame来存放所有数据
combined_df = pd.DataFrame()

# 定义处理日期的函数
def parse_date(date_str):
    date_formats = ['%Y/%m/%d', '%Y%m%d', '%Y年%m月%d日', '%Y-%m-%d']
    for fmt in date_formats:
        try:
            return pd.to_datetime(date_str, format=fmt).strftime('%Y%m%d')
        except (ValueError, TypeError):
            continue
    return ''  # 返回空字符串以表示无效日期

# 定义规范化统一社会信用代码的函数
def normalize_code(code):
    if isinstance(code, (int, float)):
        # 如果是数字，转换为字符串，确保为纯文本形式
        return '{:.0f}'.format(code)
    return str(code)

# 遍历根目录中的所有子文件夹和Excel文件
for subdir, _, files in os.walk(root_dir):
    for file in files:
        if file.endswith('.xls') or file.endswith('.xlsx'):
            file_path = os.path.join(subdir, file)
            try:
                # 尝试读取Excel文件，指定列类型
                if file.endswith('.xls'):
                    df = pd.read_excel(file_path, dtype=dtype, engine='xlrd')
                else:
                    df = pd.read_excel(file_path, dtype=dtype, engine='openpyxl')
                
                # 提取所需的列
                df = df[required_columns]
                
                # 删除“地方平台编码”列的所有内容
                df['地方平台编码'] = ''
                
                # 填充“是否为首贷”中的空白值为“否”
                df['是否为首贷'].fillna('否', inplace=True)
                
                # 处理获贷时间列
                df['获贷时间'] = df['获贷时间'].apply(parse_date)

                # 规范化统一社会信用代码列
                df['统一社会信用代码'] = df['统一社会信用代码'].apply(normalize_code)
                df['放贷机构统一社会信用代码'] = df['放贷机构统一社会信用代码'].apply(normalize_code)
                
                # 添加来源文件列
                df['银行'] = file

                # 拆分多行获贷时间
                df = df.assign(获贷时间=df['获贷时间'].str.split('\n')).explode('获贷时间')

                # 打印调试信息
                print(f"Successfully read {file_path}")
                print(df.head())  # 打印前5行数据以检查内容
                
                # 合并数据
                combined_df = pd.concat([combined_df, df], ignore_index=True)
            except Exception as e:
                print(f"Error reading {file_path}: {e}")

# 数据格式统一化
if not combined_df.empty:
    # 例如：将所有列名转换为小写，去掉多余空格
    combined_df.columns = combined_df.columns.str.strip().str.lower()

    # 删除超过3个单元格为空的行
    combined_df['null_count'] = combined_df[required_columns].isnull().sum(axis=1)
    combined_df = combined_df[combined_df['null_count'] <= 3].drop(columns=['null_count'])

    # 查找统一社会信用代码为非18位的条目（去除空格和不可见字符后）
    invalid_codes_df = combined_df[combined_df['统一社会信用代码'].apply(lambda x: len(x.replace(' ', '').replace('\u200b', '').replace('\u200c', '').replace('\u200d', '').replace('\ufeff', '')) != 18)]
    invalid_codes_df.to_excel('invalid_codes.xlsx', index=False)

    # 打印调试信息
    print("Invalid unified social credit codes and their sources:")
    print(invalid_codes_df[['统一社会信用代码', '银行']])

    # 保存合并后的数据到一个新的Excel文件
    output_file = 'combined_data.xlsx'
    combined_df.to_excel(output_file, index=False)
    print(f"所有Excel文件已合并，结果保存为 {output_file}")
else:
    print("No data found in the Excel files.")
    



























