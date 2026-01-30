### Word 表格结构查询脚本
# 从 python-docx 库中导入 Document 类，用于操作 .docx 格式的 Word 文档
from docx import Document

# 初始化 Document 实例，加载目标 Word 文档
# 提示：此处替换为你的 Word 文档路径（相对路径或绝对路径均可）
doc = Document("你的目标文档.docx")

# 遍历文档中所有的表格，通过 enumerate 同时获取表格索引和表格对象
# 表格索引从 0 开始，对应文档中第 1 个表格、第 2 个表格……
for table_index, table in enumerate(doc.tables):
    # 获取并打印当前表格的核心信息：索引、行数、列数
    table_rows = len(table.rows)  # 获取表格总行数
    table_cols = len(table.columns)  # 获取表格总列数
    print(f"表格索引 {table_index}, 行数: {table_rows}, 列数: {table_cols}")
