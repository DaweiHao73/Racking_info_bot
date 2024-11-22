import pandas as pd
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN

try:
    # 讀取Excel文件
    file_path = r"C:\Users\dwhao\OneDrive - The Dairy Farm Company Ltd\HCS VM Stuff\IMS\各店貨架借用表.xlsx"
    df = pd.read_excel(file_path, sheet_name="借出 & 借用表")
    
    # 創建PowerPoint
    ppt = Presentation()
    
    # 使用空白版面
    slide_layout = ppt.slide_layouts[6]
    slide = ppt.slides.add_slide(slide_layout)
    
    # 設置標題
    title_box = slide.shapes.add_textbox(Inches(1), Inches(0.5), Inches(8), Inches(1))
    title_box.text_frame.text = "貨架借用表"
    title_box.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    title_box.text_frame.paragraphs[0].font.size = Pt(24)
    title_box.text_frame.paragraphs[0].font.bold = True
    
    # 獲取前10行數據（你可以調整這個數字）
    rows = len(df.head(10)) + 1  # +1 是為了表頭
    cols = len(df.columns)
    
    # 創建表格
    table = slide.shapes.add_table(rows, cols, Inches(1), Inches(1.5), Inches(8), Inches(5)).table
    
    # 添加表頭
    for col_idx, column in enumerate(df.columns):
        cell = table.cell(0, col_idx)
        cell.text = column
        paragraph = cell.text_frame.paragraphs[0]
        paragraph.font.bold = True
        paragraph.font.size = Pt(11)
        paragraph.alignment = PP_ALIGN.CENTER
    
    # 添加數據
    for row_idx, row in enumerate(df.head(10).itertuples()):
        for col_idx, value in enumerate(row[1:]):  # [1:] 跳過索引
            cell = table.cell(row_idx + 1, col_idx)
            cell.text = str(value)
            paragraph = cell.text_frame.paragraphs[0]
            paragraph.font.size = Pt(10)
            paragraph.alignment = PP_ALIGN.CENTER
    
    # 保存PowerPoint
    ppt.save('貨架借用表.pptx')
    print("PowerPoint 創建成功！")

except Exception as e:
    print(f"發生錯誤: {str(e)}")