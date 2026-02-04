###塔基钢筋数据合并器
import pandas as pd
import numpy as np
from collections import defaultdict

# 1. 读取数据（已填入你的文件路径）
file_path = "/Users/mac/Desktop/work/工作簿1.xlsx"
df = pd.read_excel(file_path)

# 填充合并单元格的塔号
df["塔号"] = df["塔号"].ffill()

# 2. 数据清洗
core_columns = ["塔号", "塔腿", "规格", "长度(mm)", "数量"]
df = df[core_columns].dropna(subset=core_columns)
df["塔腿"] = df["塔腿"].astype(str).str.strip().str.upper()
df["长度(mm)"] = df["长度(mm)"].astype(float).astype(int)
df["数量"] = df["数量"].astype(float).astype(int)

# 3. 核心处理逻辑（动态前缀+合并数量）
result_data = []
original_tower_order = df["塔号"].drop_duplicates().tolist()

for tower_num, group in df.groupby("塔号"):
    row = {"塔号": tower_num, "塔腿A": "", "塔腿B": "", "塔腿C": "", "塔腿D": ""}
    leg_data = {}  # 存储各腿原始数据：{腿号: (规格, 长度, 数量)}

    # 第一步：填充塔腿列（带A：/B：标识）
    for _, item in group.iterrows():
        leg = item["塔腿"]
        spec, length, count = item["规格"], item["长度(mm)"], item["数量"]
        row[f"塔腿{leg}"] = f"{leg}：{spec}*{length}*{count}"
        leg_data[leg] = (spec, length, count)

    # 第二步：按【规格+长度】分组，统计所有腿的数量和对应腿号
    merge_dict = defaultdict(lambda: {"total_count": 0, "legs": []})
    for leg, (spec, length, count) in leg_data.items():
        key = (spec, length)
        merge_dict[key]["total_count"] += count
        merge_dict[key]["legs"].append(leg)

    # 第三步：动态生成合并列文本
    merge_parts = []
    for (spec, length), info in merge_dict.items():
        # 动态生成前缀（比如A、B、AB、C、D等）
        leg_prefix = "".join(sorted(info["legs"]))  # 排序保证AB而非BA
        # 拼接单条合并项
        merge_item = f"{leg_prefix}:{spec}*{length}*{info['total_count']}"
        merge_parts.append(merge_item)

    # 最终合并列（多个项用、分隔）
    row["合并"] = "、".join(merge_parts)
    result_data.append(row)

# 4. 按原始顺序排序
result_df = pd.DataFrame(result_data)
result_df["塔号排序键"] = result_df["塔号"].map(lambda x: original_tower_order.index(x))
result_df = result_df.sort_values("塔号排序键").drop("塔号排序键", axis=1).reset_index(drop=True)

# 5. 输出结果
output_path = "/Users/mac/Desktop/整理后_钢筋数据_动态前缀版.xlsx"
result_df.to_excel(output_path, index=False)

print(f"✅ 数据整理完成！文件保存至：{output_path}")
print("\n===== 动态合并逻辑示例 =====")
print("场景1：只有A腿 → A:C22*6900*28")
print("场景2：A+B腿规格长度相同 → AB:C22*6900*56")
print("场景3：A+B腿规格长度不同 → A:C22*6900*28、B:C22*7400*28")
print("场景4：B+C腿规格长度相同 → BC:C22*9400*56")
print("场景5：A+C+D腿规格长度相同 → ACD:C22*8400*84")
