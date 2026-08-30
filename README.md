# latex2doc_formula
latex公式批量转docx专业格式
# 基本用法：生成 文件名_converted.docx
python word_latex_converter.py 论文.docx

# 指定输出文件名
python word_latex_converter.py 论文.docx -o 终稿.docx

# 批量转换多个文件到指定目录
python word_latex_converter.py *.docx --outdir ./converted

# 直接修改原文件（自动创建 .bak 备份）
python word_latex_converter.py 论文.docx --in-place
