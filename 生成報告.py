import os
import pandas as pd
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ROW_HEIGHT_RULE, WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docxtpl import DocxTemplate, InlineImage
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox

# ============================================================
# 1. 表格樣式設定 (粗體控制、對齊、邊框)
# ============================================================

def set_cell_border(cell):
    """設置儲存格黑色細邊框"""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = OxmlElement('w:tcBorders')
    for edge in ['top', 'left', 'bottom', 'right']:
        edge_el = OxmlElement(f'w:{edge}')
        edge_el.set(qn('w:val'), 'single')
        edge_el.set(qn('w:sz'), '4')
        edge_el.set(qn('w:space'), '0')
        edge_el.set(qn('w:color'), '000000')
        tcBorders.append(edge_el)
    tcPr.append(tcBorders)

def create_table_structure(doc, table_type='測量照片'):
    """創建表格結構"""
    table = doc.add_table(rows=7, cols=4)
    table.style = 'Table Grid'
    
    # 格式化儲存格函式
    def format_cell(cell, text="", font_size=None, bold=False, align=WD_ALIGN_PARAGRAPH.CENTER):
        # 1. 設定文字
        cell.text = text
        p = cell.paragraphs[0]
        
        # 2. 設定水平對齊
        p.alignment = align
        
        # 3. 垂直置中
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        
        # 4. 設定邊框
        set_cell_border(cell)
        
        # 5. 設定字體與樣式
        for run in p.runs:
            run.font.name = '標楷體'
            run._element.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')
            
            # 指定大小
            if font_size:
                run.font.size = font_size
            
            # 指定粗體 (True/False)
            run.font.bold = bold

    # 第1行：標題區
    # ▼▼▼ 修改處：標題不加粗 (bold=False)，維持 18 號字 ▼▼▼
    row1 = table.rows[0].cells
    row1[0].merge(row1[3])
    format_cell(row1[0], '欣中天然氣(股)公司 測量作業項目照片', font_size=Pt(18), bold=False)

    # 第2-3行：資訊區 (全部加粗)
    table.rows[1].cells[0].text = '工程案號'
    table.rows[1].cells[1].text = '{{ project_number }}'
    table.rows[1].cells[2].text = '申請書編號'
    table.rows[1].cells[3].text = '{{ application_number }}'
    
    table.rows[2].cells[0].text = '施工地址'
    table.rows[2].cells[1].text = '{{ construction_address }}'
    table.rows[2].cells[2].text = '承攬商'
    table.rows[2].cells[3].text = '庭安科技'

    for r in range(1, 3):
        for c in range(4):
            # 判斷對齊方式 (地址靠左，其他置中)
            target_align = WD_ALIGN_PARAGRAPH.CENTER
            if r == 2 and c == 1:
                target_align = WD_ALIGN_PARAGRAPH.LEFT
            
            # ▼▼▼ 修改處：內容全部加粗 (bold=True) ▼▼▼
            format_cell(table.rows[r].cells[c], table.rows[r].cells[c].text, bold=True, align=target_align)

    # 第4行：類型標題 (加粗)
    # ▼▼▼ 修改處：bold=True ▼▼▼
    format_cell(table.rows[3].cells[0].merge(table.rows[3].cells[3]), table_type, bold=True)

    # 第5-7行：照片區 (加粗，雖然圖片沒粗體，但若有替代文字會加粗)
    if table_type == '測量照片':
        for i, row_idx in enumerate([4, 5, 6]):
            row = table.rows[row_idx]
            row.height_rule = WD_ROW_HEIGHT_RULE.EXACTLY
            row.height = Inches(2.4)
            
            row.cells[0].merge(row.cells[1])
            row.cells[2].merge(row.cells[3])
            
            # ▼▼▼ 修改處：bold=True ▼▼▼
            format_cell(row.cells[0], f'{{{{ photo_{i*2+1} }}}}', bold=True)
            format_cell(row.cells[2], f'{{{{ photo_{i*2+2} }}}}', bold=True)
    else:
        # 點位圖與系統截圖
        for row_idx, tag in [(4, '{{ location_map }}'), (6, '{{ system_screenshot }}')]:
            row = table.rows[row_idx]
            row.height_rule = WD_ROW_HEIGHT_RULE.EXACTLY
            row.height = Inches(3.5)
            row.cells[0].merge(row.cells[3])
            # ▼▼▼ 修改處：bold=True ▼▼▼
            format_cell(row.cells[0], tag, bold=True)
        
        # ▼▼▼ 修改處：說明文字加粗 (bold=True) ▼▼▼
        format_cell(table.rows[5].cells[0].merge(table.rows[5].cells[3]), '道挖系統上傳完成截圖', bold=True)

    return table

# ============================================================
# 2. 智慧搜尋檔案功能
# ============================================================

def find_excel_file(selected_folder):
    """依序搜尋：選定資料夾 -> 上一層 -> 程式目錄"""
    search_paths = [Path(selected_folder), Path(selected_folder).parent, Path('.')]
    for path in search_paths:
        if not path.exists(): continue
        files = list(path.glob('*.xlsx')) + list(path.glob('*.csv'))
        valid_files = [f for f in files if not f.name.startswith('~$')]
        if valid_files: return valid_files[0]
    return None

# ============================================================
# 3. 主程式邏輯
# ============================================================

def process_selected_folder():
    print(">>> 程式啟動...")
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    
    try:
        # STEP 1: 選擇資料夾
        print(">>> 請選擇包含照片或 Excel 的資料夾 (例如: 搶修)...")
        folder_path = filedialog.askdirectory(title="請選擇資料夾", parent=root)
        
        if not folder_path:
            print("❌ 使用者取消選擇")
            return
        
        print(f">>> 已選擇資料夾: {folder_path}")
        project_dir = Path(folder_path)
        
        # STEP 2: 自動搜尋 Excel
        data_file = find_excel_file(folder_path)
        
        if data_file:
            print(f">>> ✅ 成功找到資料檔: {data_file}")
        else:
            messagebox.showinfo("提示", "找不到 Excel/CSV，請手動選擇。")
            file_path = filedialog.askopenfilename(filetypes=[("Excel/CSV", "*.xlsx *.csv")])
            if not file_path: return
            data_file = Path(file_path)

        # STEP 3: 讀取 Excel 並決定案號
        if data_file.suffix == '.csv':
            df = pd.read_csv(data_file)
        else:
            df = pd.read_excel(data_file)
            
        df['工程案號'] = df['工程案號'].astype(str).str.strip()
        
        folder_name_as_id = project_dir.name
        info = df[df['工程案號'] == str(folder_name_as_id)]
        
        final_project_id = folder_name_as_id
        
        if info.empty:
            if len(df) == 1:
                print(f">>> 資料夾名稱 '{folder_name_as_id}' 不在 Excel 中，使用 Excel 內唯一案號。")
                info = df.iloc[[0]]
                final_project_id = info.iloc[0]['工程案號']
                print(f">>> ✅ 確定案號為: {final_project_id}")
            else:
                print(f"⚠️  警告: 資料夾名稱 '{folder_name_as_id}' 找不到，且 Excel 有多筆資料，無法自動判斷。")
                messagebox.showwarning("提醒", f"Excel 中找不到案號 '{folder_name_as_id}'。\n將使用資料夾名稱生成空白報告。")
                context = {'project_number': folder_name_as_id, 'application_number': '', 'construction_address': ''}
        
        if not info.empty:
            data = info.iloc[0]
            context = {
                'project_number': str(data['工程案號']),
                'application_number': str(data['申請書編號']),
                'construction_address': str(data['施工地址'])
            }

        # STEP 4: 決定照片讀取路徑
        photo_root = project_dir
        possible_subfolder = project_dir / str(final_project_id)
        if possible_subfolder.exists() and possible_subfolder.is_dir():
            print(f">>> 發現案號子資料夾，切換路徑至: {possible_subfolder.name}")
            photo_root = possible_subfolder
        
        print(f">>> 最終照片讀取路徑: {photo_root}")

        # STEP 5: 生成範本與填充
        # 使用 v8 檔名，確保更新粗體設定
        template_name = 'report_template.docx' 
        if not os.path.exists(template_name):
            doc = Document()
            create_table_structure(doc, '測量照片')
            doc.add_page_break()
            create_table_structure(doc, '點位圖')
            doc.save(template_name)

        tpl = DocxTemplate(template_name)
        
        # 讀取測量照
        photo_dir = photo_root / '測量照'
        imgs = sorted(list(photo_dir.glob('*.jpg')) + list(photo_dir.glob('*.png'))) if photo_dir.exists() else []
        
        if not imgs:
            print(f"⚠️  警告: 在 {photo_dir} 找不到任何照片！")

        for i in range(1, 7):
            if (i-1) < len(imgs):
                print(f"    - 填入照片 {i}: {imgs[i-1].name}")
                context[f'photo_{i}'] = InlineImage(tpl, str(imgs[i-1]), width=Inches(3.0))
            else:
                context[f'photo_{i}'] = ""

        # 讀取其他照片
        def get_single_img(sub, width):
            d = photo_root / sub
            f = list(d.glob('*.*')) if d.exists() else []
            return InlineImage(tpl, str(f[0]), width=Inches(width)) if f else ""

        context['location_map'] = get_single_img('點位圖', 6.0)
        context['system_screenshot'] = get_single_img('道管截圖', 6.0)

        # STEP 6: 存檔
        tpl.render(context)
        output_filename = f"{final_project_id}_報告書.docx"
        output_path = project_dir / output_filename
        
        tpl.save(output_path)
        print(f"✅ 成功！報告已儲存: {output_path}")
        messagebox.showinfo("成功", f"報告書已生成！\n位置: {output_path}")

    except Exception as e:
        print(f"❌ 嚴重錯誤: {e}")
        messagebox.showerror("錯誤", f"發生錯誤: {str(e)}")
    finally:
        root.destroy()

if __name__ == '__main__':
    process_selected_folder()