import pdfplumber
import pandas as pd
import re
import pytesseract
import numpy as np
from PIL import Image, ImageEnhance
import io
import time
import sys

# 配置参数
PDF_PATH = r"D:\work\test.pdf"
OUTPUT_EXCEL = r"D:\work\提取结果.xlsx"
PATTERN = r'报告编号[:：]\s*([A-Z0-9\-]{10,20})'
SAMPLE_PAGES = 3
RESOLUTION = 400

# OCR配置
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

def enhance_image(img):
    """图像增强处理"""
    img = img.convert('L')
    enhancer = ImageEnhance.Contrast(img)
    return enhancer.enhance(2.5)

def detect_keyword_region(pdf):
    """动态检测关键词区域（修复图像处理）"""
    coordinates = []
    
    for page in pdf.pages[:SAMPLE_PAGES]:
        try:
            # 定义搜索区域
            search_area = (
                page.width * 0.5,
                page.height * 0.0,
                page.width * 1.0,
                page.height * 0.3
            )
            
            # 正确获取图像字节流
            page_image = page.crop(search_area).to_image(resolution=RESOLUTION)
            img_bytes = io.BytesIO()
            page_image.save(img_bytes, format='PNG')  # 转换为PNG格式
            
            img = Image.open(img_bytes)
            img = enhance_image(img)
            
            # OCR识别坐标
            data = pytesseract.image_to_data(img, output_type=pytesseract.Output.DICT, lang='chi_sim')
            
            # 查找关键词位置
            for i, text in enumerate(data['text']):
                if '报告编号' in text:
                    x = data['left'][i] + search_area[0]
                    y = data['top'][i] + search_area[1]
                    coordinates.append((x/page.width, y/page.height))
        except Exception as e:
            print(f"区域检测异常：{str(e)}")
            continue
    
    # 计算平均坐标
    if coordinates:
        avg_x = np.mean([c[0] for c in coordinates])
        avg_y = np.mean([c[1] for c in coordinates])
        return (
            max(0.0, avg_x - 0.15),
            max(0.0, avg_y - 0.05),
            min(1.0, avg_x + 0.25),
            min(1.0, avg_y + 0.08)
        )
    
    # 默认区域
    return (0.7, 0.05, 0.95, 0.15)

def extract_report_number(page, region):
    """提取报告编号（修复语法错误）"""
    try:
        # 计算实际坐标
        x0 = page.width * region[0]
        y0 = page.height * region[1]
        x1 = page.width * region[2]
        y1 = page.height * region[3]
        
        # 生成图像字节流
        page_image = page.crop((x0, y0, x1, y1)).to_image(resolution=RESOLUTION)
        img_bytes = io.BytesIO()
        page_image.save(img_bytes, format='PNG')
        
        # 图像增强处理
        img = enhance_image(Image.open(img_bytes))
        
        # OCR识别
        text = pytesseract.image_to_string(
            img,
            lang='chi_sim+eng',
            config='--psm 6 --oem 3 -c tessedit_char_whitelist=0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ-:：'
        ).replace(' ', '').strip()
        
        # 文本清洗（正确缩进在try块内）
        clean_text = re.sub(r'[^\w\-:：]', '', text)
        
        # 正则匹配
        if match := re.search(PATTERN, clean_text):
            return match.group(1), text
        return "未匹配到格式", text
    
    except Exception as e:
        return f"提取失败：{str(e)}", ""

def main():
    # 依赖检查
    missing = []
    try:
        import openpyxl, pdfplumber, pytesseract, PIL
    except ImportError as e:
        missing.append(e.name)
    
    if missing:
        print("缺少依赖库：", ", ".join(missing))
        print("请执行：pip install openpyxl pdfplumber pytesseract pillow")
        sys.exit(1)
    
    # 处理流程
    results = []
    start_time = time.time()
    
    try:
        with pdfplumber.open(PDF_PATH) as pdf:
            # 动态检测区域
            region = detect_keyword_region(pdf)
            print(f"🔍 检测到识别区域：{tuple(round(x,2) for x in region)}")
            
            # 批量处理页面
            total = len(pdf.pages)
            for idx, page in enumerate(pdf.pages, 1):
                record = {
                    "文件名": PDF_PATH.split("\\")[-1],
                    "页码": idx,
                    "报告编号": "",
                    "原始文本": "",
                    "状态": ""
                }
                
                num, text = extract_report_number(page, region)
                record.update({
                    "报告编号": num if isinstance(num, str) else "",
                    "原始文本": text,
                    "状态": "成功" if isinstance(num, str) and "失败" not in num else num
                })
                results.append(record)
                print(f"\r处理进度：{idx}/{total}页", end='')
            
            # 生成结果
            df = pd.DataFrame(results)
            df.to_excel(OUTPUT_EXCEL, index=False)
            
            # 统计信息
            success = df[df['状态'] == '成功'].shape[0]
            print(f"\n✅ 完成！成功：{success}/{total} ({success/total:.1%})")
            print(f"⏱ 耗时：{time.time()-start_time:.1f}秒")
            print(f"📂 保存路径：{OUTPUT_EXCEL}")
    
    except Exception as e:
        print(f"\n❌ 严重错误：{str(e)}")
        sys.exit(1)

if __name__ == "__main__":
    main()
