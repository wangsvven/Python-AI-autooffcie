###  Word 表格坐标探测脚本 
# 从 python-docx 库中导入 Document 类，用于操作 .docx 格式的 Word 文档
from docx import Document

# 初始化 Document 实例，加载目标 Word 文档（用于探测表格坐标）
# 提示：替换为你需要探测坐标的 Word 文档路径（相对路径/绝对路径均可）
doc = Document("你的表格文档.docx")

# 获取文档中的第 1 个表格（索引 0 对应文档中第一个表格，如需其他表格可修改索引）
target_table = doc.tables[0]

# 遍历表格的所有行和单元格，为每个单元格填入对应的「行索引,列索引」坐标
# r_idx：行索引（从 0 开始），row：当前遍历到的行对象
for r_idx, row in enumerate(target_table.rows):
    # c_idx：列索引（从 0 开始），cell：当前遍历到的单元格对象
    for c_idx, cell in enumerate(row.cells):
        # 给当前单元格写入「行索引,列索引」格式的坐标值
        # 如需保留原有内容，可修改为：cell.text = f"{cell.text} ({r_idx},{c_idx})"
        cell.text = f"{r_idx},{c_idx}"

# 保存生成的坐标探测文档，避免覆盖原文档
# 提示：可自定义输出文档名，方便识别
doc.save("表格坐标探测结果.docx")

# 打印运行完成提示，引导用户查看结果
print("探测文档已生成，请打开 '表格坐标探测结果.docx' 查看数据起始位置的数字。")
