# Docx Duplicate Image Finder

這是一個用 Python 撰寫的工具，專門用於掃描指定資料夾內的所有 Microsoft Word (`.docx`) 檔案，並找出其中「重複」或「高度相似」的圖片。

本工具提供**純字元介面 (CLI)** 與**圖形化介面 (GUI)** 兩種操作方式，並且支援產出 Markdown 格式的完整檢查報告。

## 🌟 功能亮點

- **快速掃描**：直接把 `.docx` 當作 ZIP 解析，不需要依賴或開啟 Microsoft Word 或 LibreOffice。
- **精準相似度比對**：採用 **Perceptual Hash (感知雜湊, phash)** 核心演算法。即使圖片被稍微調整過大小、壓縮過，只要視覺上雷同，程式都能準確判定為同一張圖片。
- **詳細溯源資訊**：不僅抓出圖片，還能告知您圖片存在於哪個檔案的「第幾頁」，以及上下文標題或內容為何。
- **跨平台 GUI**：以 PyQt6 打造現代化的圖形介面，輕鬆選擇資料夾並調整相似度容忍閥值。
- **自動產出報告**：檢測完畢後，自動於目標資料夾下建立 `report` 目錄，產出 Markdown 詳細分析報告供存查。

---

## 🚀 安裝與環境準備

本工具使用 `uv` 來作為套件與虛擬環境管理工具，確保環境乾淨且一致。

### 1. 安裝 `uv`
若您尚未安裝 `uv`，請根據您的作業系統參考 [uv 官方安裝指南](https://github.com/astral-sh/uv)。

### 2. 下載專案
```bash
git clone git@github.com:0524041/docx_image_compare.git
cd docx_image_compare
```

### 3. 初始化依賴
在第一次執行前，`uv` 會自動根據 `pyproject.toml` 或 `uv.lock` 安裝所需套件（例如 `Pillow`, `ImageHash`, `PyQt6` 等）。

---

## 🎨 使用方式

### GUI 圖形化介面 (推薦)

您可以直接執行包裹好的 Shell 腳本，它會自動呼叫 `uv` 並開啟圖形介面：

```bash
./run_gui.sh
```

**操作步驟**：
1. 點擊「瀏覽...」選擇含有您想比對的 `.docx` 的資料夾。
2. 調整「相似度閥值」（預設為 3）。
   - **數值越小**：越嚴格（0 代表必須完全一模一樣）。
   - **數值越大**：能容忍更多的壓縮或微調變形，但誤判機率會些微增加。
3. 點擊「開始比對」。下方的進度條與日誌區會即時顯示掃描狀態。
4. 完成後，您可以在該資料夾底下的 `report/` 目錄中找到生成的檢測報告 (`.md` 檔)。

### CLI 指令操作

如果您偏好使用終端機指令（例如自動化作業）：

```bash
# 預設閾值為 5
uv run find_docx_duplicates.py /您的/目標/資料夾路徑

# 自訂相似度閾值 (例如設定為 3)
uv run find_docx_duplicates.py /您的/目標/資料夾路徑 --threshold 3
```

終端機將會列出完整的檢查結果報告。

---

## 🛠 技術規格

- **語言**: Python 3
- **套件管理**: `uv`
- **圖形套件**: `PyQt6`
- **核心依賴**: `Pillow`, `ImageHash`
- **解析方式**: `xml.etree.ElementTree`, `zipfile`

---

## 授權與版權
MIT License.
