import pandas as pd
from docxtpl import DocxTemplate
from pathlib import Path
import datetime
import time

# ================= ⚙️ 用户配置区域 (修改这里) =================

# 1. Excel 数据文件路径
EXCEL_PATH = '/Users/mac/Library/CloudStorage/OneDrive-个人/1.项目/攀枝花米易撒莲丙谷光伏发电项目（35kV 集电线路）/6.过程资料/7.相关数据/（数据）表D.0.4 灌注桩基础检查记录表.xlsx'

# 2. Word 模板文件路径
TEMPLATE_PATH = '/Users/mac/Desktop/work/表D.0.4  灌注桩基础检查记录表(线基1)py.docx'

# 3. 指定 Excel 中哪一列的内容作为生成文件的文件名
FILENAME_COLUMN = '设计桩号'

# 4. 结果输出文件夹名称
OUTPUT_DIR = '灌注桩基础检查记录表'

# 5. 日期格式设置
DATE_FORMAT_STR = '%Y年%m月%d日'

# 6. 需要“强制去小数”的列名列表
INT_COLUMNS = [
    '根设AB', '根设BC', '根设CD', '根设DA', '根设AC', '根设BD', '间距',
    # 如果还有其他需要取整的列，继续加在这里
]

# 7. 【新功能】指定工作表名称 (Sheet Name)
# 如果你的数据在第一个表，可以填 None (不带引号)
# 如果数据在特定表，请填入名称，例如 'Sheet1' 或 '数据录入'
SHEET_NAME = '检验批数据'


# =============================================================

def clean_filename(filename):
    """清理文件名中包含的非法字符"""
    invalid_chars = '<>:"/\\|?*\n\r\t'
    filename = str(filename)
    for char in invalid_chars:
        filename = filename.replace(char, '_')
    return filename.strip()


def process_data(key, value):
    """
    数据清洗核心逻辑 (包含 V2.1 的强力数字修复)
    """
    if pd.isna(value):
        return ""

    # 处理日期
    if isinstance(value, (datetime.datetime, pd.Timestamp)):
        return value.strftime(DATE_FORMAT_STR)

    # 【重要】尝试将字符串数字转换为浮点数 (修复 Excel 文本格式数字问题)
    if isinstance(value, str):
        try:
            value = float(value)
        except ValueError:
            pass

    # 处理数值 (强力取整)
    if isinstance(value, (int, float)):
        if key in INT_COLUMNS:
            return int(value)  # 强制去掉小数
        else:
            if value == int(value):
                return int(value)
            return round(value, 2)

    return value


def main():
    start_time = time.time()

    # 1. 路径处理
    base_path = Path(EXCEL_PATH).parent
    excel_file = Path(EXCEL_PATH)
    template_file = Path(TEMPLATE_PATH)
    output_path = base_path / OUTPUT_DIR

    print("=" * 50)
    print(f"🚀 自动化填充工具 V3.0 (多Sheet版)")
    print("=" * 50)

    # 检查文件
    if not excel_file.exists():
        print(f"❌ 错误：找不到 Excel 文件\n路径: {excel_file}")
        return
    if not template_file.exists():
        print(f"❌ 错误：找不到模板文件\n路径: {template_file}")
        return

    output_path.mkdir(parents=True, exist_ok=True)

    # 2. 读取 Excel 信息
    print("⏳ 正在分析 Excel 文件结构...")
    try:
        # 先加载 Excel 文件对象，查看有哪些 Sheet
        xls = pd.ExcelFile(excel_file)
        sheet_names = xls.sheet_names
        print(f"📄 发现工作表: {sheet_names}")

        # 确定要读取哪个 Sheet
        target_sheet = SHEET_NAME

        # 如果用户填了 None，默认读第一个
        if target_sheet is None:
            target_sheet = sheet_names[0]
            print(f"👉 未指定工作表，默认读取第一个: [{target_sheet}]")

        # 检查指定的 Sheet 是否存在
        if target_sheet not in sheet_names:
            print(f"❌ 错误：找不到名为 '{target_sheet}' 的工作表！")
            print(f"   当前 Excel 中只有: {sheet_names}")
            print(f"   请修改代码第 29 行的 SHEET_NAME 配置。")
            return

        # 读取指定 Sheet 的数据
        print(f"📖 正在读取工作表: [{target_sheet}] ...")
        df = pd.read_excel(excel_file, sheet_name=target_sheet)

    except Exception as e:
        print(f"❌ 读取 Excel 失败: {e}")
        return

    # 检查文件名列
    if FILENAME_COLUMN not in df.columns:
        print(f"❌ 错误：在表 [{target_sheet}] 中找不到列名: [{FILENAME_COLUMN}]")
        print(f"   当前表包含列名: {list(df.columns)}")
        return

    # 3. 批量生成
    total = len(df)
    print(f"✅ 读取成功，共 {total} 条数据，开始生成...\n")

    success_count = 0

    for index, row in df.iterrows():
        try:
            context = {k: process_data(k, v) for k, v in row.items()}

            doc = DocxTemplate(template_file)
            doc.render(context)

            fname = clean_filename(context.get(FILENAME_COLUMN, f'Result_{index}'))
            save_path = output_path / f"{fname}.docx"

            doc.save(save_path)
            success_count += 1
            print(f"  [{(index + 1):03d}/{total}] 🟢 {fname}.docx")

        except Exception as e:
            print(f"  [{(index + 1):03d}/{total}] 🔴 失败: {e}")

    duration = time.time() - start_time
    print("\n" + "=" * 50)
    print(f"🎉 处理完成！耗时: {duration:.2f} 秒")
    print(f"📂 文件已保存在: {output_path}")


if __name__ == '__main__':
    main()
