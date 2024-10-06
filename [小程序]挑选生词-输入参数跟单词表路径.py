import pandas as pd
import os
from tqdm import tqdm
import argparse

# 输入与熟练度的映射
proficiency_map = {
    '0': '0%',
    '1': '20%',
    '2': '40%',
    '3': '60%',
    '4': '80%',
    '5': '100%'
}

# 设置命令行参数解析
parser = argparse.ArgumentParser(description='Process an Excel file for word proficiency.')
parser.add_argument('input_file', type=str, help='Path to the input Excel file')
args = parser.parse_args()

input_file = args.input_file

try:
    df = pd.read_excel(input_file, sheet_name=None)
    df = df[list(df.keys())[0]]  # 假设你想处理第一个表格
except FileNotFoundError:
    print("File not found.")
    exit()

# 确保需要的列存在
if '熟练度' not in df.columns:
    df['熟练度'] = None

# 打乱数据框的顺序
df = df.sample(frac=1).reset_index(drop=True)

# 获取控制台尺寸
console_size = os.get_terminal_size()
console_width = console_size.columns
console_height = console_size.lines

# 初始化进度条
progress_bar = tqdm(total=len(df), desc="Progress", ncols=100)

index = 0
while index < len(df):
    try:
        row = df.iloc[index]

        if pd.isna(row['熟练度']):  # 仅处理未打分的行
            os.system('cls')  # 清屏
            word = row['单词']

            # 计算垂直居中的行数
            vertical_padding = (console_height // 2) - 1

            # 打印空行以实现垂直居中
            print("\n" * vertical_padding)
            # 居中显示单词
            print(word.center(console_width))

            # 打印空行以将输入框移到左下角
            print("\n" * (console_height - vertical_padding - 4))

            # 输入行
            score = input("Enter the proficiency score (0-5, 'back', 'pre', 'next', or 'exit' to quit): ")

            if score.lower() == 'exit':
                break
            elif score.lower() in ['back', 'pre']:
                index = max(0, index - 1)
                continue
            elif score.lower() == 'next':
                index = min(len(df) - 1, index + 1)
                continue

            # 将输入转换为熟练度
            proficiency = proficiency_map.get(score, None)
            if proficiency is not None:
                df.at[index, '熟练度'] = proficiency
                # 显示单词的释义
                meaning = row['释义']
                print(f"Meaning: {meaning}")
                input("Press Enter to continue...")

            # 更新进度条
            progress_bar.update(1)

        # 移动到下一个单词
        index = min(len(df) - 1, index + 1)

    except Exception as e:
        print(f"An error occurred: {e}")
        break

# 关闭进度条
progress_bar.close()

# 保存进度到原文件的新表格
with pd.ExcelWriter(input_file, mode='a', engine='openpyxl') as writer:
    df.to_excel(writer, sheet_name='Scored Words', index=False)

print("Progress saved to the original file in 'Scored Words' sheet.")
