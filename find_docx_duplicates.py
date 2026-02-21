import os
import sys
import zipfile
import xml.etree.ElementTree as ET
from PIL import Image
import imagehash
import io
import argparse

# Docx XML 檔案中常用的命名空間
NS = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'pic': 'http://schemas.openxmlformats.org/drawingml/2006/picture',
    'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
}

def extract_images_from_docx(docx_path):
    """
    解析 Docx 壓縮檔，提取裡面的圖片以及其所在的章節或上下文。
    """
    images_info = []
    try:
        with zipfile.ZipFile(docx_path, 'r') as docx_zip:
            # 1. 讀取關聯檔 (_rels) 來取得關聯 ID 與實體圖檔路徑的映射關係
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

            # 2. 讀取主文件內容，依序解析段落與圖片
            doc_path = 'word/document.xml'
            if doc_path not in docx_zip.namelist():
                return images_info
                
            doc_xml = docx_zip.read(doc_path)
            doc_tree = ET.fromstring(doc_xml)
            
            current_chapter = "開頭/未命名章節"
            recent_text_buffer = []

            # 找到文件的 body
            body = doc_tree.find('w:body', NS)
            if body is None:
                return images_info

            # 依序走訪所有元素
            for elem in body:
                if elem.tag == f"{{{NS['w']}}}p": # 是一個段落
                    # 提取這段的文字
                    texts = [t.text for t in elem.findall('.//w:t', NS) if t.text]
                    para_text = "".join(texts).strip()
                    
                    if para_text:
                        # 檢查這段文字的樣式是不是標題 (Heading)
                        pPr = elem.find('w:pPr', NS)
                        if pPr is not None:
                            pStyle = pPr.find('w:pStyle', NS)
                            if pStyle is not None:
                                style_val = pStyle.get(f"{{{NS['w']}}}val")
                                if style_val and style_val.startswith('Heading'):
                                    current_chapter = para_text
                                    recent_text_buffer = [] # 遇到新標題就清空上下文
                        
                        recent_text_buffer.append(para_text)
                        # 只保留最近兩段有文字的段落作為上下文參考
                        if len(recent_text_buffer) > 2:
                            recent_text_buffer.pop(0)

                    # 檢查這段落裡有沒有夾帶圖片繪圖 (drawing)
                    for drawing in elem.findall('.//w:drawing', NS):
                        blips = drawing.findall('.//a:blip', NS)
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
                                        'bytes': img_bytes
                                    })
                                    
    except Exception as e:
        print(f"處理檔案時發生錯誤 {docx_path}: {e}")
        
    return images_info

def main():
    parser = argparse.ArgumentParser(description="比對目標資料夾中所有 docx 檔案內的圖片使否重複。")
    parser.add_argument("folder", help="包含 docx 檔案的資料夾絕對或相對路徑")
    parser.add_argument("--threshold", type=int, default=5, help="圖片相似度寬容閥值 (預設 5，越小越嚴格，0 代表完全一模一樣)")
    args = parser.parse_args()

    folder_path = args.folder
    threshold = args.threshold

    if not os.path.isdir(folder_path):
        print(f"錯誤：找不到指定的資料夾 '{folder_path}'")
        sys.exit(1)

    docx_files = [os.path.join(folder_path, f) for f in os.listdir(folder_path) if f.lower().endswith('.docx') and not f.startswith('~')]
    
    if not docx_files:
        print(f"在 '{folder_path}' 中找不到任何 docx 檔案。")
        sys.exit(0)

    print(f"找到 {len(docx_files)} 個 docx 檔案，開始解析並提取圖片...\n")

    all_images = []
    
    for df in docx_files:
        print(f"  處理讀取: {os.path.basename(df)}")
        extracted = extract_images_from_docx(df)
        for img_info in extracted:
            try:
                # 讀取圖片 Bytes，並透過 Pillow 將其轉成圖片物件
                img = Image.open(io.BytesIO(img_info['bytes']))
                
                # 計算 Perceptual Hash (感知雜湊)
                # Phash 對於圖片稍微壓縮、調整大小等微小變動具有很強的抵抗力
                img_hash = imagehash.phash(img)
                img_info['hash'] = img_hash
                all_images.append(img_info)
            except Exception as e:
                print(f"    無法解析圖片 {img_info['image_name']}: {e}")

    print(f"\n共提取並計算了 {len(all_images)} 張圖片的 Hash。開始進行相似度比對 (目前的容忍閥值為: {threshold})...")

    # 利用分群演算法將相似的圖片分類
    groups = []
    
    for img in all_images:
        found_group = False
        for group in groups:
            # 與群組內的第一張代表圖片進行比較
            # ImageHash 可以直接透過減號計算兩個 Hash 之間的涵明距離 (Hamming distance)
            if img['hash'] - group[0]['hash'] <= threshold:
                group.append(img)
                found_group = True
                break
        
        # 若與所有現有群組都不相似，就自己建立一個新群組
        if not found_group:
            groups.append([img])

    # 輸出簡易報告到終端機
    print("\n" + "="*60)
    print(" 📊 圖片重複檢查報告")
    print("="*60)
    
    dup_count = 0
    for i, group in enumerate(groups, 1):
        if len(group) > 1:
            dup_count += 1
            print(f"\n[發現重複群組 #{dup_count}] 共 {len(group)} 張相似度極高的圖片:")
            for img in group:
                print(f"  📂 檔案來源: {img['filename']}")
                print(f"  📍 所在章節/位置段落: {img['context']}")
                print(f"  🖼 內部資源名稱: {img['image_name']}")
                print(f"  🔑 Hash: {img['hash']}")
            print("-" * 60)

    print("\n" + "="*60)
    if dup_count == 0:
        print("🎉 太棒了！所有的檔案中沒有發現任何重複且相似的圖片。")
    else:
        print(f"⚠️  檢查完畢，總共發現 {dup_count} 組重複/相似的圖片。")
    print("="*60 + "\n")

if __name__ == "__main__":
    main()
