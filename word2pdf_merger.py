import os
from docx2pdf import convert
from PyPDF2 import PdfMerger
import tempfile
import datetime  # 新增时间模块

# 配置路径
word_folder = r'D:\work'
output_pdf = os.path.join(word_folder, '合并结果.pdf')

# 记录开始时间
start_time = datetime.datetime.now()
print(f"🕒 程序启动时间：{start_time.strftime('%Y-%m-%d %H:%M:%S')}")

# 创建临时文件夹
temp_dir = tempfile.TemporaryDirectory()

# 第一步：转换所有Word为PDF
for filename in os.listdir(word_folder):
    if filename.lower().endswith(('.doc', '.docx')):
        word_path = os.path.join(word_folder, filename)
        try:
            # 转换到临时目录
            convert(word_path, os.path.join(temp_dir.name, f"{filename}.pdf"))
            print(f"✅ 已转换：{filename}")
        except Exception as e:
            print(f"❌ 转换失败：{filename} - {str(e)}")

# 第二步：合并所有PDF
pdf_merger = PdfMerger()
for pdf_file in os.listdir(temp_dir.name):
    if pdf_file.endswith('.pdf'):
        pdf_path = os.path.join(temp_dir.name, pdf_file)
        try:
            pdf_merger.append(pdf_path)
            print(f"📑 已合并：{pdf_file}")
        except Exception as e:
            print(f"⚠️ 合并失败：{pdf_file} - {str(e)}")

# 保存最终结果
pdf_merger.write(output_pdf)
pdf_merger.close()
temp_dir.cleanup()

# 计算运行时间
end_time = datetime.datetime.now()
time_cost = end_time - start_time

# 将耗时转换为易读格式
hours, remainder = divmod(time_cost.seconds, 3600)
minutes, seconds = divmod(remainder, 60)
time_str = f"{hours}小时{minutes}分{seconds}秒" if hours > 0 else f"{minutes}分{seconds}秒"

print(f"\n🎉 合并完成！")
print(f"📂 输出路径：{output_pdf}")
print(f"⏰ 开始时间：{start_time.strftime('%H:%M:%S')}")
print(f"🕒 完成时间：{end_time.strftime('%H:%M:%S')}")
print(f"⏱ 总耗时：{time_str}")
