import os
import sys
import re
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import fitz
import pytesseract
from PIL import Image, ImageEnhance
import io
import openpyxl
from difflib import SequenceMatcher

# ===== Tesseract路径设置（支持打包） =====
def get_tesseract_path():
    if getattr(sys, 'frozen', False):
        exe_dir = os.path.dirname(sys.executable)
        tess_path = os.path.join(exe_dir, 'tesseract.exe')
        if os.path.exists(tess_path):
            return tess_path
        if hasattr(sys, '_MEIPASS'):
            tess_path = os.path.join(sys._MEIPASS, 'tesseract.exe')
            if os.path.exists(tess_path):
                return tess_path
    possible_paths = [
        r'C:\Program Files\Tesseract-OCR\tesseract.exe',
        r'C:\Program Files (x86)\Tesseract-OCR\tesseract.exe',
    ]
    for path in possible_paths:
        if os.path.exists(path):
            return path
    return 'tesseract'

tesseract_path = get_tesseract_path()
if os.path.exists(tesseract_path):
    pytesseract.pytesseract.tesseract_cmd = tesseract_path

class FBAExtractorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("FBA仓库代码提取器 - 日本跨境物流")
        self.root.geometry("850x700")
        self.root.resizable(False, False)
        
        self.excel_path = tk.StringVar()
        self.pdf_folder = tk.StringVar()
        self.output_path = tk.StringVar()
        self.progress_var = tk.DoubleVar()

        self.create_widgets()
        self.check_tesseract()

    def check_tesseract(self):
        try:
            version = pytesseract.get_tesseract_version()
            self.status_label.config(text=f"✓ Tesseract OCR 已就绪", foreground="green")
        except Exception as e:
            self.status_label.config(text="✗ Tesseract OCR 未找到", foreground="red")

    def create_widgets(self):
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)

        title_label = ttk.Label(main_frame, text="📦 FBA仓库代码批量提取工具", font=("Arial", 18, "bold"))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))

        ttk.Label(main_frame, text="1. Excel文件 (含FBA箱号):").grid(row=1, column=0, sticky=tk.W, pady=5)
        ttk.Entry(main_frame, textvariable=self.excel_path, width=50).grid(row=1, column=1, padx=5, pady=5)
        ttk.Button(main_frame, text="📁 浏览", command=self.browse_excel).grid(row=1, column=2, padx=5, pady=5)

        ttk.Label(main_frame, text="2. PDF图片文件夹:").grid(row=2, column=0, sticky=tk.W, pady=5)
        ttk.Entry(main_frame, textvariable=self.pdf_folder, width=50).grid(row=2, column=1, padx=5, pady=5)
        ttk.Button(main_frame, text="📁 浏览", command=self.browse_pdf_folder).grid(row=2, column=2, padx=5, pady=5)

        ttk.Label(main_frame, text="3. 输出Excel文件:").grid(row=3, column=0, sticky=tk.W, pady=5)
        ttk.Entry(main_frame, textvariable=self.output_path, width=50).grid(row=3, column=1, padx=5, pady=5)
        ttk.Button(main_frame, text="💾 保存为", command=self.browse_output).grid(row=3, column=2, padx=5, pady=5)

        ttk.Separator(main_frame, orient='horizontal').grid(row=4, column=0, columnspan=3, sticky=tk.W+tk.E, pady=15)

        self.progress = ttk.Progressbar(main_frame, orient=tk.HORIZONTAL, length=500, mode='determinate', variable=self.progress_var)
        self.progress.grid(row=5, column=0, columnspan=3, pady=10)

        self.status_label = ttk.Label(main_frame, text="就绪", foreground="gray")
        self.status_label.grid(row=6, column=0, columnspan=3, pady=5)

        log_frame = ttk.LabelFrame(main_frame, text="处理日志", padding="5")
        log_frame.grid(row=7, column=0, columnspan=3, pady=10, sticky=tk.W+tk.E)
        
        self.log_text_widget = tk.Text(log_frame, height=10, width=80, font=("Consolas", 9))
        self.log_text_widget.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        scrollbar = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_text_widget.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text_widget.config(yscrollcommand=scrollbar.set)
        self.log_text_widget.insert('1.0', "准备就绪，请选择文件开始处理...\n")

        self.run_button = ttk.Button(main_frame, text="🚀 开始提取", command=self.run_extraction, width=20)
        self.run_button.grid(row=8, column=0, columnspan=3, pady=20)

    def browse_excel(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if path:
            self.excel_path.set(path)
            self.log_message(f"已选择Excel: {os.path.basename(path)}")

    def browse_pdf_folder(self):
        path = filedialog.askdirectory()
        if path:
            self.pdf_folder.set(path)
            pdf_count = len([f for f in os.listdir(path) if f.lower().endswith('.pdf')])
            self.log_message(f"已选择文件夹，含 {pdf_count} 个PDF")

    def browse_output(self):
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if path:
            self.output_path.set(path)
            self.log_message(f"输出: {os.path.basename(path)}")

    def log_message(self, message):
        timestamp = pd.Timestamp.now().strftime("%H:%M:%S")
        self.log_text_widget.insert(tk.END, f"[{timestamp}] {message}\n")
        self.log_text_widget.see(tk.END)
        self.root.update()

    def enhance_image(self, image):
        if image.mode != 'L':
            image = image.convert('L')
        enhancer = ImageEnhance.Contrast(image)
        image = enhancer.enhance(2.5)
        enhancer = ImageEnhance.Sharpness(image)
        image = enhancer.enhance(2.0)
        image = image.point(lambda x: 0 if x < 150 else 255, '1')
        return image

    def extract_text_from_pdf_image(self, pdf_path):
        try:
            doc = fitz.open(pdf_path)
            all_text = ""
            for page_num in range(min(len(doc), 3)):
                page = doc[page_num]
                zoom = 4.0
                mat = fitz.Matrix(zoom, zoom)
                pix = page.get_pixmap(matrix=mat)
                img_data = pix.tobytes("png")
                img = Image.open(io.BytesIO(img_data))
                enhanced_img = self.enhance_image(img)
                text = pytesseract.image_to_string(enhanced_img, lang='eng+chi_sim', config='--psm 6 --oem 3')
                all_text += text + "\n"
            doc.close()
            return all_text
        except Exception as e:
            self.log_message(f"❌ PDF处理失败: {str(e)}")
            return None

    def extract_all_fba_numbers_from_text(self, text):
        if not text:
            return []
        text = text.replace('\n', ' ').replace('\r', ' ')
        patterns = [r'FBA[A-Z0-9]{12,20}', r'FBA[A-Z0-9]+']
        all_fba_numbers = []
        for pattern in patterns:
            matches = re.findall(pattern, text, re.IGNORECASE)
            for match in matches:
                fba = match.strip().upper()
                if fba.startswith('FBA') and len(fba) >= 12:
                    all_fba_numbers.append(fba)
        return list(set(all_fba_numbers))

    def correct_ocr_errors(self, text):
        if not text:
            return text
        corrections = {
            'VFU4': 'XHD4', 'VFU': 'XHD', 'VHD4': 'XHD4', 'XFU4': 'XHD4',
            'XHU4': 'XHD4', 'VHD': 'XHD', 'XHDA': 'XHD4', 'XH04': 'XHD4',
            'XH D4': 'XHD4', 'TVD4': 'TYO2', 'TYD2': 'TYO2', 'TY02': 'TYO2',
            'XJVV1': 'XJW1', 'XJVV': 'XJW1', 'XJW': 'XJW1', 'XJWI': 'XJW1',
        }
        for wrong, correct in corrections.items():
            text = text.replace(wrong, correct)
        return text

    def extract_warehouse_code_from_text(self, text):
        if not text:
            return None
        text_corrected = self.correct_ocr_errors(text)
        text_clean = text_corrected.replace('\n', ' ').replace('\r', ' ')
        
        sta_patterns = [
            r'FBA\s*STA[^\-]*-\s*([A-Z0-9]{2,8})',
            r'FBA\s*STA.*?-\s*([A-Z0-9]{2,8})',
            r'STA[^\-]*-\s*([A-Z0-9]{2,8})',
        ]
        for pattern in sta_patterns:
            match = re.search(pattern, text_clean, re.IGNORECASE)
            if match:
                code = match.group(1).strip().upper()
                if 2 <= len(code) <= 8:
                    return code
        
        dest_patterns = [
            r'目的地\s*[:：]\s*([A-Z0-9]{2,8})',
            r'FBA仓库\s*[:：]\s*([A-Z0-9]{2,8})',
            r'Destination\s*[:：]\s*([A-Z0-9]{2,8})',
        ]
        for pattern in dest_patterns:
            match = re.search(pattern, text_clean, re.IGNORECASE)
            if match:
                code = match.group(1).strip().upper()
                if 2 <= len(code) <= 8:
                    return code
        
        jp_codes = ['XJW1', 'XHD4', 'TYO2', 'TY02', 'NRT1', 'NRT2', 'HND1', 'HND2', 'KIX1', 'KIX2']
        text_upper = text_clean.upper()
        for code in jp_codes:
            if code in text_upper:
                return code
        
        generic = r'\b([A-Z]{2,4}[0-9]{1,2})\b'
        matches = re.findall(generic, text_upper)
        for match in matches:
            if len(match) >= 3 and not match.isdigit() and match not in ['FBA']:
                return match
        return None

    def find_best_match(self, excel_fba, pdf_fba_list):
        if not pdf_fba_list:
            return None, 0
        excel_fba = excel_fba.strip().upper()
        best_match = None
        best_score = 0
        for pdf_fba in pdf_fba_list:
            pdf_fba = pdf_fba.strip().upper()
            if excel_fba == pdf_fba:
                return pdf_fba, 1.0
            if excel_fba in pdf_fba:
                score = len(excel_fba) / len(pdf_fba)
                if score > best_score:
                    best_score = score
                    best_match = pdf_fba
            if pdf_fba in excel_fba:
                score = len(pdf_fba) / len(excel_fba)
                if score > best_score:
                    best_score = score
                    best_match = pdf_fba
            similarity = SequenceMatcher(None, excel_fba, pdf_fba).ratio()
            if similarity > best_score and similarity >= 0.6:
                best_score = similarity
                best_match = pdf_fba
        return best_match, best_score

    def run_extraction(self):
        excel_file = self.excel_path.get().strip()
        pdf_folder = self.pdf_folder.get().strip()
        output_file = self.output_path.get().strip()

        if not excel_file or not pdf_folder or not output_file:
            messagebox.showerror("错误", "请完整填写所有路径")
            return

        self.log_text_widget.delete('1.0', tk.END)
        self.log_message("=" * 50)
        self.log_message("🚀 开始处理...")
        
        try:
            df = pd.read_excel(excel_file, engine='openpyxl')
            self.log_message(f"✓ 读取Excel成功，共 {len(df)} 行")
        except Exception as e:
            self.log_message(f"❌ 读取失败: {str(e)}")
            return

        if 'FBA号' not in df.columns:
            self.log_message("❌ Excel缺少 'FBA号' 列")
            return

        self.run_button.config(state=tk.DISABLED)
        self.progress_var.set(0)

        pdf_files = [os.path.join(pdf_folder, f) for f in os.listdir(pdf_folder) if f.lower().endswith('.pdf')]
        self.log_message(f"✓ 找到 {len(pdf_files)} 个PDF")

        pdf_content_map = {}
        for idx, pdf_path in enumerate(pdf_files):
            self.log_message(f"读取PDF ({idx+1}/{len(pdf_files)}): {os.path.basename(pdf_path)}")
            text = self.extract_text_from_pdf_image(pdf_path)
            if text:
                fba_list = self.extract_all_fba_numbers_from_text(text)
                if fba_list:
                    warehouse = self.extract_warehouse_code_from_text(text)
                    for fba in fba_list:
                        pdf_content_map[fba] = (warehouse, pdf_path)
                        self.log_message(f"  ✓ 箱号: {fba}, 仓库: {warehouse}")
            self.progress_var.set((idx + 1) / len(pdf_files) * 30)
            self.root.update()

        self.log_message(f"✓ 提取到 {len(pdf_content_map)} 个箱号")
        self.log_message("开始匹配...")

        warehouse_codes = []
        success_count = 0
        pdf_fba_list = list(pdf_content_map.keys())

        for idx, row in df.iterrows():
            excel_fba = str(row['FBA号']).strip()
            best_match, score = self.find_best_match(excel_fba, pdf_fba_list)
            
            if best_match and score >= 0.6:
                warehouse, _ = pdf_content_map[best_match]
                if warehouse:
                    warehouse_codes.append(warehouse)
                    success_count += 1
                    self.log_message(f"✅ 匹配成功: {excel_fba} → {warehouse}")
                else:
                    warehouse_codes.append("未提取到仓库代码")
            else:
                warehouse_codes.append("未匹配到PDF")
            
            self.progress_var.set(30 + (idx + 1) / len(df) * 70)
            self.root.update()

        df['仓库代码'] = warehouse_codes
        df.to_excel(output_file, index=False, engine='openpyxl')
        
        self.log_message("=" * 50)
        self.log_message(f"✅ 完成！成功匹配 {success_count}/{len(df)} 条")
        messagebox.showinfo("完成", f"成功匹配 {success_count}/{len(df)} 条！")
        self.run_button.config(state=tk.NORMAL)

if __name__ == "__main__":
    root = tk.Tk()
    app = FBAExtractorApp(root)
    root.mainloop()