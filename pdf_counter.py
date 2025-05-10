import os
from PyPDF2 import PdfReader

# 指定文件夹路径
folder_path = r'C:\Users\hello\Desktop\'

total_pages = 0

# 遍历文件夹中的所有文件
for filename in os.listdir(folder_path):
    if filename.lower().endswith('.pdf'):
        file_path = os.path.join(folder_path, filename)
        try:
            with open(file_path, 'rb') as f:
                # 创建PDF阅读器对象
                reader = PdfReader(f)
                # 获取页数并累加
                num_pages = len(reader.pages)
                total_pages += num_pages
                print(f"✅ 已处理：{filename} ({num_pages}页)")
        except Exception as e:
            print(f"❌ 错误：无法读取文件 {filename} - {str(e)}")

# 输出最终结果
print(f"\n📊 总页数统计完成！共 {total_pages} 页")
