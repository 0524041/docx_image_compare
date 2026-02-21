import os
import sys
import zipfile
import xml.etree.ElementTree as ET
from PIL import Image
import imagehash
import io
import datetime

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QPushButton, QSlider, QProgressBar, QTextEdit,
    QFileDialog, QMessageBox
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal

# --- 核心邏輯 ---
NS = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'pic': 'http://schemas.openxmlformats.org/drawingml/2006/picture',
    'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
}

def extract_images_from_docx(docx_path):
    images_info = []
    try:
        with zipfile.ZipFile(docx_path, 'r') as docx_zip:
            rels_path = 'word/_rels/document.xml.rels'
            if rels_path not in docx_zip.namelist():
                return images_info
            
            rels_xml = docx_zip.read(rels_path)
            rels_tree = ET.fromstring(rels_xml)
            
            rel_map = {}
            for rel in rels_tree.findall('.//{http://schemas.openxmlformats.org/package/2006/relationships}Relationship'):
                rel_id = rel.get('Id')
                target = rel.get('Target')
                if target.startswith('media/'):
                    rel_map[rel_id] = target

            doc_path = 'word/document.xml'
            if doc_path not in docx_zip.namelist():
                return images_info
                
            doc_xml = docx_zip.read(doc_path)
            doc_tree = ET.fromstring(doc_xml)
            
            current_chapter = "開頭/未命名章節"
            recent_text_buffer = []

            # 嘗試計算頁數：Word 在分頁時通常會插入 <w:lastRenderedPageBreak> 或 <w:br w:type="page"/>
            current_page = 1

            body = doc_tree.find('w:body', NS)
            if body is None:
                return images_info

            # 遞迴或線性尋找段落與分頁符號
            # 這裡我們用簡單的迭代 w:p 和其他可能有分頁符號的元素
            for elem in body.iter():
                # 計算頁碼
                if elem.tag == f"{{{NS['w']}}}lastRenderedPageBreak":
                    current_page += 1
                elif elem.tag == f"{{{NS['w']}}}br":
                    br_type = elem.get(f"{{{NS['w']}}}type")
                    if br_type == "page":
                        current_page += 1

                # 處理段落
                if elem.tag == f"{{{NS['w']}}}p":
                    texts = [t.text for t in elem.findall('.//w:t', NS) if t.text]
                    para_text = "".join(texts).strip()
                    
                    if para_text:
                        pPr = elem.find('w:pPr', NS)
                        if pPr is not None:
                            pStyle = pPr.find('w:pStyle', NS)
                            if pStyle is not None:
                                style_val = pStyle.get(f"{{{NS['w']}}}val")
                                if style_val and style_val.startswith('Heading'):
                                    current_chapter = para_text
                                    recent_text_buffer = []
                        
                        recent_text_buffer.append(para_text)
                        if len(recent_text_buffer) > 2:
                            recent_text_buffer.pop(0)

                # 處理圖片
                if elem.tag == f"{{{NS['w']}}}drawing":
                    blips = elem.findall('.//a:blip', NS)
                    for blip in blips:
                        embed_id = blip.get(f"{{{NS['r']}}}embed")
                        if embed_id and embed_id in rel_map:
                            target_media = 'word/' + rel_map[embed_id]
                            if target_media in docx_zip.namelist():
                                img_bytes = docx_zip.read(target_media)
                                 
                                context = current_chapter
                                if current_chapter == "開頭/未命名章節" and recent_text_buffer:
                                    context = f"上下文: {' '.join(recent_text_buffer)}"
                                    
                                images_info.append({
                                    'filename': os.path.basename(docx_path),
                                    'image_name': target_media.split('/')[-1],
                                    'context': context[:50] + "..." if len(context) > 50 else context,
                                    'page': current_page,
                                    'bytes': img_bytes
                                })
                                
    except Exception as e:
        print(f"處理檔案時發生錯誤 {docx_path}: {e}")
        
    return images_info

# --- 背景任務執行緒 ---
class WorkerThread(QThread):
    log_signal = pyqtSignal(str)
    progress_signal = pyqtSignal(int, int)
    finished_signal = pyqtSignal()

    def __init__(self, folder_path, threshold):
        super().__init__()
        self.folder_path = folder_path
        self.threshold = threshold

    def run(self):
        try:
            self.log_signal.emit("啟動比對任務...")
            docx_files = [os.path.join(self.folder_path, f) for f in os.listdir(self.folder_path) if f.lower().endswith('.docx') and not f.startswith('~')]
            
            if not docx_files:
                self.log_signal.emit(f"錯誤：在 '{self.folder_path}' 中找不到任何 docx 檔案。")
                self.finished_signal.emit()
                return

            self.log_signal.emit(f"找到 {len(docx_files)} 個 docx 檔案，開始解析並提取圖片...")

            all_images = []
            
            total_files = len(docx_files)
            for i, df in enumerate(docx_files):
                self.log_signal.emit(f"  處理讀取: {os.path.basename(df)}")
                extracted = extract_images_from_docx(df)
                for img_info in extracted:
                    try:
                        img = Image.open(io.BytesIO(img_info['bytes']))
                        img_hash = imagehash.phash(img)
                        img_info['hash'] = img_hash
                        all_images.append(img_info)
                    except Exception as e:
                        self.log_signal.emit(f"    無法解析圖片 {img_info['image_name']}: {e}")
                
                self.progress_signal.emit(i + 1, total_files)

            self.log_signal.emit(f"\n共提取並計算了 {len(all_images)} 張圖片。開始進行相似度比對 (目前的容忍閥值為: {self.threshold})...")

            groups = []
            for img in all_images:
                found_group = False
                for group in groups:
                    if img['hash'] - group[0]['hash'] <= self.threshold:
                        group.append(img)
                        found_group = True
                        break
                
                if not found_group:
                    groups.append([img])

            dup_count = 0
            duplicate_groups = []
            
            self.log_signal.emit("\n" + "="*60)
            self.log_signal.emit(" 📊 圖片重複檢查報告")
            self.log_signal.emit("="*60)
            
            for i, group in enumerate(groups, 1):
                if len(group) > 1:
                    dup_count += 1
                    duplicate_groups.append(group)
                    
                    self.log_signal.emit(f"\n[發現重複群組 #{dup_count}] 共 {len(group)} 張相似度極高的圖片:")
                    for img in group:
                        self.log_signal.emit(f"  📂 檔案來源: {img['filename']}")
                        self.log_signal.emit(f"  📄 所在頁數: 第 {img['page']} 頁")
                        self.log_signal.emit(f"  📍 所在節錄: {img['context']}")
                        self.log_signal.emit(f"  🖼 圖片名稱: {img['image_name']}")
                        self.log_signal.emit(f"  🔑 Hash: {img['hash']}")
                    self.log_signal.emit("-" * 60)

            self.log_signal.emit("\n" + "="*60)
            if dup_count == 0:
                self.log_signal.emit("🎉 太棒了！所有的檔案中沒有發現任何重複且相似的圖片。")
            else:
                self.log_signal.emit(f"⚠️  檢查完畢，總共發現 {dup_count} 組重複/相似的圖片。")
            self.log_signal.emit("="*60 + "\n")
            
            self.generate_html_report(total_files, len(all_images), duplicate_groups)

        except Exception as e:
            self.log_signal.emit(f"\n執行中發生錯誤: {e}")
        finally:
            self.finished_signal.emit()

    def generate_html_report(self, file_count, image_count, dup_groups):
        report_dir = os.path.join(self.folder_path, "report")
        if not os.path.exists(report_dir):
            os.makedirs(report_dir)
            
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        report_path = os.path.join(report_dir, f"Duplicate_Image_Report_{timestamp}.html")
        
        html_content = f"""<!DOCTYPE html>
<html lang="zh-TW">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Docx 圖片重複檢測報告</title>
    <style>
        body {{ font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; line-height: 1.6; color: #333; max-width: 1000px; margin: 0 auto; padding: 20px; }}
        h1 {{ color: #2c3e50; border-bottom: 2px solid #3498db; padding-bottom: 10px; }}
        h2 {{ color: #2980b9; margin-top: 30px; }}
        .summary {{ background: #f8f9fa; padding: 15px; border-radius: 8px; margin-bottom: 30px; border-left: 4px solid #3498db; }}
        .group {{ background: #fff; border: 1px solid #ddd; border-radius: 8px; margin-bottom: 20px; padding: 15px; box-shadow: 0 2px 4px rgba(0,0,0,0.05); }}
        .group-title {{ font-size: 1.2em; font-weight: bold; color: #e74c3c; margin-top: 0; margin-bottom: 15px; }}
        ul {{ list-style-type: none; padding: 0; margin: 0; }}
        li {{ margin-bottom: 15px; padding-bottom: 15px; border-bottom: 1px dashed #eee; }}
        li:last-child {{ margin-bottom: 0; border-bottom: none; padding-bottom: 0; }}
        .detail-label {{ font-weight: bold; color: #555; display: inline-block; width: 150px; }}
        .success-msg {{ font-size: 1.2em; color: #27ae60; font-weight: bold; text-align: center; padding: 20px; background: #e8f8f5; border-radius: 8px; }}
        code {{ background: #f4f4f4; padding: 2px 5px; border-radius: 4px; font-family: monospace; color: #d63031; }}
    </style>
</head>
<body>
    <h1>Docx 圖片重複檢測報告</h1>
    <div class="summary">
        <p><span class="detail-label">產生時間:</span> {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</p>
        <p><span class="detail-label">掃描資料夾:</span> <code>{self.folder_path}</code></p>
        <p><span class="detail-label">相似度閥值:</span> {self.threshold}</p>
        <p><span class="detail-label">掃描文件數量:</span> {file_count}</p>
        <p><span class="detail-label">提取圖片數量:</span> {image_count}</p>
        <p><span class="detail-label">發現重複群組:</span> {len(dup_groups)}</p>
    </div>
"""
        with open(report_path, "w", encoding="utf-8") as f:
            f.write(html_content)
            
            if not dup_groups:
                f.write('    <div class="success-msg">🎉 太棒了！所有的檔案中沒有發現任何重複且相似的圖片。</div>\n')
            else:
                f.write('    <h2>⚠️ 重複圖片詳細資料</h2>\n')
                for i, group in enumerate(dup_groups, 1):
                    f.write(f'    <div class="group">\n')
                    f.write(f'        <div class="group-title">發現重複群組 #{i} (共 {len(group)} 張高度相似圖片)</div>\n')
                    f.write('        <ul>\n')
                    for img in group:
                        f.write(f'            <li>\n')
                        f.write(f'                <div><span class="detail-label">檔案來源:</span> <code>{img["filename"]}</code></div>\n')
                        f.write(f'                <div><span class="detail-label">所在頁數:</span> 第 {img["page"]} 頁</div>\n')
                        f.write(f'                <div><span class="detail-label">所在節錄:</span> {img["context"]}</div>\n')
                        f.write(f'                <div><span class="detail-label">內部資源名稱:</span> <code>{img["image_name"]}</code></div>\n')
                        f.write(f'                <div><span class="detail-label">特徵雜湊碼:</span> <code>{img["hash"]}</code></div>\n')
                        f.write(f'            </li>\n')
                    f.write('        </ul>\n')
                    f.write('    </div>\n')
            f.write('</body>\n</html>\n')
                    
        self.log_signal.emit(f"\n[系統提示] 詳細 HTML 報告已儲存至: \n{report_path}")


# --- GUI 應用程式 ---
class DuplicateFinderApp(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Docx 圖片重複檢測工具")
        self.resize(750, 600)

        # 中心 Widget
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # 主垂直佈局
        main_layout = QVBoxLayout(central_widget)

        # 1. 頂部選擇資料夾區域
        folder_layout = QHBoxLayout()
        lbl_folder = QLabel("目標資料夾:")
        self.entry_folder_path = QLineEdit()
        self.entry_folder_path.setPlaceholderText("請選擇含有 docx 檔案的資料夾...")
        btn_browse = QPushButton("瀏覽...")
        btn_browse.clicked.connect(self.browse_folder)
        
        folder_layout.addWidget(lbl_folder)
        folder_layout.addWidget(self.entry_folder_path)
        folder_layout.addWidget(btn_browse)
        main_layout.addLayout(folder_layout)

        # 2. 設定區域
        settings_layout = QHBoxLayout()
        lbl_threshold = QLabel("相似度閥值 (0~20):")
        
        self.slider_threshold = QSlider(Qt.Orientation.Horizontal)
        self.slider_threshold.setMinimum(0)
        self.slider_threshold.setMaximum(20)
        self.slider_threshold.setValue(3)
        self.slider_threshold.setTickPosition(QSlider.TickPosition.TicksBelow)
        self.slider_threshold.setTickInterval(1)
        self.slider_threshold.valueChanged.connect(self.update_threshold_label)
        
        self.lbl_threshold_val = QLabel("3")
        self.lbl_threshold_val.setMinimumWidth(30)
        
        self.btn_run = QPushButton("開始比對")
        self.btn_run.setStyleSheet("background-color: #2E8B57; color: white; font-weight: bold; padding: 5px;")
        self.btn_run.clicked.connect(self.start_processing)
        
        settings_layout.addWidget(lbl_threshold)
        settings_layout.addWidget(self.slider_threshold)
        settings_layout.addWidget(self.lbl_threshold_val)
        settings_layout.addStretch()
        settings_layout.addWidget(self.btn_run)
        main_layout.addLayout(settings_layout)

        # 3. 進度條
        self.progressbar = QProgressBar()
        self.progressbar.setValue(0)
        main_layout.addWidget(self.progressbar)

        # 4. 資訊輸出區
        self.textbox_log = QTextEdit()
        self.textbox_log.setReadOnly(True)
        self.textbox_log.setStyleSheet("font-family: 'Courier New'; font-size: 13px;")
        main_layout.addWidget(self.textbox_log)

        # Thread reference
        self.worker = None

    def update_threshold_label(self, value):
        self.lbl_threshold_val.setText(str(value))

    def browse_folder(self):
        folder_selected = QFileDialog.getExistingDirectory(self, "選擇目標資料夾")
        if folder_selected:
            self.entry_folder_path.setText(folder_selected)

    def log(self, text):
        self.textbox_log.append(text)
        # Scroll to bottom
        scrollbar = self.textbox_log.verticalScrollBar()
        scrollbar.setValue(scrollbar.maximum())

    def update_progress(self, current, total):
        pct = int((current / total) * 100)
        self.progressbar.setValue(pct)

    def task_finished(self):
        self.btn_run.setEnabled(True)
        self.progressbar.setValue(100)

    def start_processing(self):
        folder_path = self.entry_folder_path.text().strip()
        if not folder_path or not os.path.isdir(folder_path):
            QMessageBox.critical(self, "錯誤", "請選擇有效的資料夾")
            return
            
        threshold = self.slider_threshold.value()
        
        self.btn_run.setEnabled(False)
        self.textbox_log.clear()
        self.progressbar.setValue(0)
        
        # 啟動背景處理
        self.worker = WorkerThread(folder_path, threshold)
        self.worker.log_signal.connect(self.log)
        self.worker.progress_signal.connect(self.update_progress)
        self.worker.finished_signal.connect(self.task_finished)
        self.worker.start()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle("Fusion") # 給一個看起來乾淨現代的樣式
    window = DuplicateFinderApp()
    window.show()
    sys.exit(app.exec())
