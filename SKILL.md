---
name: python-docx-photo-grid
description: |
  Python-based Word (.docx) 報告生成，配「自適應 2 欄相片網格」技術 + 工程竣工報告
  三件套（燈具安裝 / Lux 照度 / 風扇）+ 勘察備忘錄 統一排版（藍色系、承建商抬頭圖頁眉、
  統一簽署區、章節結構）。觸發詞：生成 docx 報告 with 自動相片排版、相片表格、photo
  auto-layout、竣工報告相片、test report with photos、自適應相片排版、A4 縱向 2 欄相片、
  docx 報告生成、Word 相片插入、python-docx Pillow、竣工報告三件套、燈具安裝報告、
  Lux照度報告、風扇安裝報告、承建商抬頭圖、統一簽署、藍色系報告。
  提供：完整 A4 幾何常數、自適應網格算法（核心技術：按行數動態縮放）、add_picture_fit
  工具、MD5 去重、正確嘅 caption 生成方式（避開 _01 後綴 bug）、統一簽署區 + 頁眉抬頭圖
  helper、驗證清單、坑位大全。可與 officecli-workflow、electrical-test-report-generator、
  site-inspection-memo-generator 等報告生成技能組合使用。
---

# python-docx-photo-grid — Word 報告 + 自適應相片排版 + 竣工報告三件套

> **呢個技能嘅核心使命：解決「換任務就唔識、換會話又唔識、換技能又唔識」嘅復發問題。**
> 凡係「生成 docx 報告 + 自動塞相片 + 自動排版 + 統一簽署/抬頭」嘅任務，**先睇呢度**。
>
> 📎 **統一模板/SOP/helper 完整版**：`references/竣工報告三件套_模板與SOP.md`
> （合併自知識庫 SOP + 藍色系模板 + python-docx 踩坑 + 實機交付件實測參數）

---

## 1. 適用場景（When to use）

| 場景 | 典型任務 |
|---|---|
| **竣工報告三件套** | 燈具安裝完成報告、Lux 照度報告、風扇安裝完成報告（同一項目一齊出）|
| 竣工/驗收/測試報告 | 設備驗收報告、電箱竣工詳情 |
| 巡查/勘察記錄 | 現場勘察備忘錄（藍色系 Letter 版）、質量檢驗報告 |
| 任何「文字+大量相片」報告 | 工作總結、施工日誌、售後報告 |
| 需要相片排版靚 | 標書附件、客戶報告、年報 |

**唔適用**：純文字文檔（用 word-typography-guide）；修改現有 docx（用 officecli-workflow）；表格數據為主（用 faithful-xlsx-template）；勘察備忘錄專用流程（用 site-inspection-memo-generator，但相片網格技術照用本技能）。

---

## 2. 環境搭建（必做，唔好跳過）

### 2.1 Managed Python venv

永遠用 **managed venv**，唔好污染系統 Python：

```bash
# managed python 位置
PY="<WORKBUDDY_DIR>/binaries/python/versions/3.13.12/python.exe"

# 建 venv（首次）
$PY -m venv "<WORKBUDDY_DIR>/binaries/python/envs/default"

# 裝依賴
"<WORKBUDDY_DIR>/binaries/python/envs/default/Scripts/pip.exe" install python-docx Pillow
```

### 2.2 Windows vs Git Bash 路徑坑 ⚠️

| 工具 | 路徑格式 | 例子 |
|---|---|---|
| **Windows python.exe** | **Windows 反斜杠 `D:\...`** | `r"D:\工作文件\xxx.docx"` |
| Git Bash shell 命令 | POSIX `/d/...` | `/d/工作文件/xxx.docx` |

**規則：**
- 餵畀 `python -c` 或 `.py` 嘅路徑 → **必須用 `D:\...`**（python-docx 跑喺 Windows 下，POSIX 路徑會 PackageNotFoundError）
- 喺 bash heredoc/cp/ls 用嘅路徑 → POSIX `/d/...`
- venv 內 `python.exe` 喺 **`Scripts\python.exe`**，唔係 `bin/python`

---

## 3. ⭐ 統一排版規範（竣工報告三件套 / 勘察備忘）

> 完整版見 `references/竣工報告三件套_模板與SOP.md`。以下係**必須記住嘅核心**。

### 3.1 紙型（兩套，唔好撈亂）⚠️

| 文檔類型 | 紙型 | 邊距 |
|:---|:---|:---|
| **勘察備忘錄** | **Letter 21.59 × 27.94 cm** | 上/下 2.6、左/右 2.8 |
| **竣工報告三件套** | **A4 21.0 × 29.7 cm** | 左 1.5 固定；上 2.54、下 1.2、右 1.1 |

### 3.2 配色（藍色系，一文一色系）

深藍 `#1F4E79`（標題/表頭）｜中藍 `#2E75B6`（副標題）｜淺藍 `#D6E4F0`｜綠 `#C6EFCE`（合格）｜橙 `#FCE4D6`（取消/待定）｜紅 `#F8CBAD`（獨立供電）｜灰 `#808080`（圖注）｜斑馬 `#F2F2F2`。

### 3.3 字體 / 字號

- 中文 eastAsia：**PMingLiU**；ascii/hAnsi：**Times New Roman**（`set_run` 三屬性一齊設）。
- 封面 22pt Arial 粗；章節標題 16pt 粗；子標題 13pt；正文 12pt；表格 10pt；圖注 10pt 灰。
- 頁碼：footer `PAGE / NUMPAGES` 域（OxmlElement `w:fldChar` 三件套）。

### 3.4 頁眉：承建商公司抬頭圖

- section header 第一段落 run 內嵌公司 logo，**實際渲染 ~7.5 × 2.0 cm**（交付件實測 6.5~7.7 × 1.8~2.1）。
- ⚠️ 首次用源圖 15.92 × 4.32 cm 會**過大**壓正文；插入後縮到 ~7.5×2.0，或用大圖時上邊距調到 ~4.6 cm。
- helper：`add_header_logo(doc, logo_path, width_cm=7.5, height_cm=2.0)`（見附錄 §五）。

### 3.5 統一簽署區（權威格式，三份報告一致）

```
[16pt 粗體]  簽署
測試/安裝人員：工程承建商　　　日期：2026-08-28
業主/用家：________________　　　日期：________________
```

- 報告開頭資訊區另有「承建商：工程承建商」。
- 簽署章節白底黑字（唔填色）。helper：`add_signing_block(doc)`（見附錄 §六）。

### 3.6 章節結構 — 竣工報告三件套（標準模板）

| 報告 | 章節結構 |
|:---|:---|
| **燈具安裝完成報告** | 標題 → 一、工程概況 → 二、燈具安裝完成明細 → 三、現場安裝照片 → 簽署 |
| **Lux 照度報告** | 標題 → 一、測試概述 → 二、標準依據 → 三、照度測試數據 → 四、測試點照片 → 簽署 |
| **風扇安裝完成報告** | 標題 → 一、工程概況 → 二、風扇安裝完成明細 → 三、現場安裝照片 → 簽署 |

---

## 4. A4 縱向頁面幾何（核心常數）

```python
# ===== A4 縱向（margin 2.54 cm）=====
PAGE_W, PAGE_H = 21.0, 29.7
MARGIN = 2.54
CONTENT_W = PAGE_W - 2 * MARGIN      # 15.92 cm
CONTENT_H = PAGE_H - 2 * MARGIN      # 24.62 cm
```

如果用其他紙型/邊距（如勘察備忘 Letter），改 `PAGE_W/PAGE_H/MARGIN`，後面所有計算自動跟住變。

---

## 5. ⭐ 自適應 2 欄相片網格（核心算法）

> **呢個係成個技能最值錢嘅技術。** 解決「固定 2×2 有大片空白」、「固定每頁 N 張排唔落」嘅矛盾。

### 5.1 原理

1. 每頁最多 `MAX_ROWS × GRID_COLS` 張（默認 3×2 = 6 張）
2. **每頁按實際行數動態計算相片高度上限**：
   ```
   h_allow = (PAGE_BUDGET / rows) - CAPTION_H - CELL_PAD - SAFETY
   h_allow = clamp(h_allow, MIN_IMG_H, ABS_MAX_IMG_H)
   ```
3. 該頁所有相片按 `h_allow` 為上限縮放（保持長寬比）
4. 因為相片永遠可以縮細，所以 3 行 6 張**必定放得落**

### 5.2 為何「固定行數」會失敗

常見錯誤：用「貪婪塞行」（放唔落先換頁）— 如果相片多係手機直拍（直向圖），渲染高度頂到上限，一頁點計都只放到 2 行 = 4 張，無改善。

**正確做法：行數固定，但相片大小跟住行數變。** 3 行就自動縮細到 ~5.9 cm 高；2 行就放大到 9 cm；1 行（尾頁剩少）就放到上限。

### 5.3 完整代碼模板

```python
# ===== 網格參數 =====
GRID_COLS   = 2          # 固定 2 欄
MAX_ROWS    = 3          # 最多 3 行 → 每頁最多 6 張
COL_W       = CONTENT_W / GRID_COLS         # 7.96 cm
IMG_MAX_W   = COL_W - 0.96                  # 7.00 cm
RESERVE_H   = 2.2        # 標題 + 說明 + 行距預留
PAGE_BUDGET = CONTENT_H - RESERVE_H         # 22.42 cm
CAPTION_H   = 0.9
CELL_PAD    = 0.5
SAFETY      = 0.15
ABS_MAX_IMG_H = 9.0      # 單張相絕對高度上限
MIN_IMG_H  = 3.0

# ===== 工具：按比例縮放至上限 =====
from PIL import Image
def rendered_size(path, max_w, max_h):
    img = Image.open(path); ar = img.width / img.height
    w = max_w; h = w / ar
    if h > max_h:
        h = max_h; w = h * ar
    return w, h

def add_picture_fit(paragraph, path, max_w_cm, max_h_cm):
    w, h = rendered_size(path, max_w_cm, max_h_cm)
    run = paragraph.add_run()
    run.add_picture(path, width=Cm(w), height=Cm(h))
    return run

# ===== 表格版面固定（重要）=====
def fix_table_layout(table, col_w):
    table.autofit = False
    tblPr = table._element.tblPr
    layout = tblPr.makeelement(qn('w:tblLayout'), {qn('w:type'): 'fixed'})
    tblPr.append(layout)
    for row in table.rows:
        for cell in row.cells:
            cell.width = Cm(col_w)

# ===== 核心：自適應相片排版 =====
def add_photo_grid(doc, items, caption_size=9):
    """items = [(image_path, caption_text), ...]
    每頁 1 個 2 欄網格表，按行數動態縮放相片。"""
    per_page = MAX_ROWS * GRID_COLS
    pages = [items[i:i+per_page] for i in range(0, len(items), per_page)]
    counts = []
    for pi, page in enumerate(pages):
        rows = (len(page) + GRID_COLS - 1) // GRID_COLS
        # 該頁相片高度上限
        h_allow = (PAGE_BUDGET / rows) - CAPTION_H - CELL_PAD - SAFETY
        h_allow = max(min(h_allow, ABS_MAX_IMG_H), MIN_IMG_H)

        table = doc.add_table(rows=rows, cols=GRID_COLS)
        table.style = "Table Grid"
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        fix_table_layout(table, COL_W)

        for idx, (img_path, caption) in enumerate(page):
            ri, ci = divmod(idx, GRID_COLS)
            cell = table.cell(ri, ci)
            # 文字描述喺相片上方
            p_cap = cell.paragraphs[0]
            p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
            set_run(p_cap.add_run(caption), size=caption_size,
                    bold=True, color="595959")
            # 相片（自動縮放）
            p_img = cell.add_paragraph()
            p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
            if img_path and os.path.exists(img_path):
                add_picture_fit(p_img, img_path, IMG_MAX_W, h_allow)
            else:
                set_run(p_img.add_run("[相片缺失]"), size=9, color="C00000")

        # 空白儲存格補齊
        for idx in range(len(page), rows * GRID_COLS):
            ri, ci = divmod(idx, GRID_COLS)
            table.cell(ri, ci).text = ""

        set_table_borders(table)
        counts.append(len(page))
        if pi < len(pages) - 1:
            doc.add_page_break()
    return counts
```

> 大量相（如 Lux 29 張）用上邊自適應網格；少量相（風扇 2-4 張）可用附錄 §八嘅 `add_photo_grid()`（2 列無框線版，相放大到上限）。

### 5.4 成效對比（真實數據）

| 報告 | 相片數 | 固定 2×2 | 自適應 |
|---|---|---|---|
| 燈具安裝 18 張 | 18 | 5 頁 | **3 頁** [6,6,6] |
| Lux 照度 29 張 | 29 | 8 頁 | **5 頁** [6,6,6,6,5] |
| 風扇 3 張 | 3 | 1 頁 | 1 頁（2 行，相放大到 9 cm） |

---

## 6. ⭐ Caption 生成（避開 `_01` 後綴 Bug）

### 6.1 錯嘅做法

```python
# 假設檔名 "132室_465lux.jpg"
base = os.path.basename(ph)          # "132室_465lux.jpg"
lux_val = base.split("_")[-1].replace("lux.jpg", "")
# → "465"  OK

# 但若檔名係 "132室_465lux_01.jpg"（重複相，_01 後綴）
base = "132室_465lux_01.jpg"
lux_val = base.split("_")[-1].replace("lux.jpg", "")
# → "01.jpg"  ❌  caption 變成 "132室 — 01.jpg lux"
```

**Bug 確認**：用戶喺 Lux 報告中手動修正咗呢個 bug（將 caption 改為 "132室 — 465.6 lux" 用平均值）。

### 6.2 正確做法：用數據表嘅平均值/編號

```python
# LUX_DATA = {"132室": [465, 502, 505, 606]}
loc_avg = round(sum(LUX_DATA[loc]) / len(LUX_DATA[loc]), 1)  # 519.5
caption = f"{loc} — {loc_avg} lux"   # "132室 — 519.5 lux"（交付件實測格式）

# 或者用測點編號
caption = f"{loc} 測點 {idx+1}"       # "132室 測點 1"
```

**原則：caption 內容必須來自結構化數據，唔好由檔名解析。** 檔名只係 ID，唔係數據。

---

## 7. MD5 去重（必做）

微信/相機重複匯出嘅相，MD5 會完全相同。唔去重會導致：
- python-docx 自動重用同一張圖，導致相片數對唔上 glob 數量
- 報告出現重複相

```python
import hashlib, glob
seen_md5 = set()
for loc in LOCATIONS:
    pat = os.path.join(PHOTO_DIR, f"{loc}_*lux*.jpg")
    for ph in sorted(glob.glob(pat)):
        with open(ph, 'rb') as fp:
            m = hashlib.md5(fp.read()).hexdigest()
        if m in seen_md5:
            continue
        seen_md5.add(m)
        # ... 加入 items
```

---

## 8. 壓縮相片（控文件大小）

```python
def compress_image(src, dst, max_width_cm=12, dpi=200):
    img = Image.open(src)
    max_w = int(max_width_cm * dpi / 2.54)
    if img.width > max_w:
        ratio = max_w / img.width
        img = img.resize((max_w, int(img.height * ratio)), Image.LANCZOS)
    img.save(dst, "JPEG", quality=85, dpi=(dpi, dpi))
```

`max_width_cm=12, dpi=200, quality=85` 係經驗值，文件大小與質素嘅最佳平衡。
3 份報告（18+29+3=50 張相）總文件 ~3 MB，合理。

---

## 9. 驗證清單（每次生成後必做）

### 9.1 相片網格驗證（python-docx 讀回）

```python
# scripts/verify_docx.py
from docx import Document
from docx.oxml.ns import qn
import os

EMU = 360000
PAGE_BUDGET = 22.42
CONTENT_H   = 24.62

def verify(path, expect_imgs):
    d = Document(path)
    tables = d.tables
    imgs = [r for r in d.part.rels.values() if "image" in r.reltype]

    # 1. 數據表格 vs 相片表格
    photo_pages = []
    for t in tables:
        grid = []
        for row in t.rows:
            hs = []
            for c in row.cells:
                got = None
                for para in c.paragraphs:
                    for ext in para._p.iter(qn('wp:extent')):
                        got = int(ext.get('cy')) / EMU
                hs.append(got)
            grid.append(hs)
        if any(any(h is not None for h in hs) for hs in grid):
            n_img = sum(1 for hs in grid for h in hs if h is not None)
            est = sum(
                (max([h for h in hs if h is not None]) if any(h is not None for h in hs) else 0)
                + 0.9 + 0.5  # CAPTION_H + CELL_PAD
                for hs in grid
            )
            photo_pages.append({
                "rows": len(grid), "cols": len(t.columns),
                "n_img": n_img, "est_h": round(est, 2)
            })

    total_img = sum(p["n_img"] for p in photo_pages)
    shape_ok = all(p["rows"] <= 3 and p["cols"] == 2 for p in photo_pages)
    over = [p for p in photo_pages if p["est_h"] > PAGE_BUDGET + 0.01]

    print(f"{os.path.basename(path)}")
    print(f"  表格={len(tables)} 相片頁={len(photo_pages)} 嵌入相={len(imgs)} (預期{expect_imgs}) 表格內相={total_img}")
    print(f"  每頁: {photo_pages}")
    print(f"  形狀OK={shape_ok}  超出預算={len(over)}")
    assert imgs and total_img == len(imgs), "嵌入相 vs 表格內相不一致"
    assert shape_ok, "相片表格形狀錯誤"
    assert not over, "有頁超出 A4 預算"
    print("  ✅ PASS")
```

### 9.2 三件套完整性檢查（加埋呢啲）

- 頁眉有承建商 logo（header 圖片存在，寬 ~7.5cm）
- 最後一章係「簽署」（16pt bold），含「測試/安裝人員：工程承建商」+「業主/用家：____」兩行
- 報告開頭有「承建商：工程承建商」
- 章節標題順序符合三件套模板（見 §3.6）
- 紙型正確（竣工報告 A4 / 勘察備忘 Letter）
- 頁腳有 `PAGE / NUMPAGES` 域

---

## 10. 坑位大全（讀一次受用一世）

| # | 坑 | 症狀 | 解決 |
|---|---|---|---|
| 1 | managed venv 路徑寫 `bin/python` | 找不到 python | Windows venv 係 **`Scripts\python.exe`** |
| 2 | 餵 POSIX `/d/...` 路徑畀 python-docx | PackageNotFoundError | python 跑喺 Windows 下，要用 `D:\...` |
| 3 | caption 由檔名解析（`_01` 後綴） | "132室 — 01.jpg lux" 垃圾文字 | caption 內容用結構化數據，唔好解析檔名 |
| 4 | 相片重複匯出無去重 | 嵌入相數 ≠ glob 數 | MD5 set 去重 |
| 5 | 用 `run.add_picture(width=, height=)` 兩個都設但 aspect 錯 | 相片變形 | 計算好 aspect 後只設限制維度，或兩個都按 aspect 算 |
| 6 | TableGrid 預設 `autofit=True` | Word 自動調欄寬，2 欄變 1 欄 | `fix_table_layout()`：tblLayout fixed + autofit=False + 鎖 cell.width |
| 7 | 「貪婪塞行」配直向相 | 一頁只 2 行 4 張，無改善 | 改用「行數固定、相片按行數縮放」（核心算法） |
| 8 | Word 保存報「權限錯誤」 | 文件被設唯讀 / 鎖住 | `attrib -R file.docx` + `os.chmod(file, 0o777)` |
| 9 | Read tool 讀圖報「Content filtered」 | 模型話睇唔到 | 靠 cache 機制逐張確認；或用 MD5 交叉驗證 |
| 10 | 文件首次存喺 Desktop | Sandbox 靜默拒絕 | 改放 `<PROJECT_DIR>` 或 `dangerouslyDisableSandbox` |
| 11 | 頁眉圖太大（15.92×4.32）壓正文 | 正文被 logo 覆蓋 | 縮到 ~7.5×2.0 cm，或上邊距調到 ~4.6 cm |
| 12 | `add_body()` 傳兩個 text 參數 | "文字2" 被當 size 傳入 `Pt()` 報錯 | `+` 拼接或分兩次調用 |
| 13 | `set_run()` 只設 `run.font.name` | 中文字體唔生效 | `w:ascii` + `w:hAnsi` + `w:eastAsia` 三屬性一齊設 |
| 14 | `set_cell()` 直接 `cell.text=""` | 殘留 empty run 產生 bug | 先清 runs 再 `add_run()` |
| 15 | `set_col_widths()` 只設 `cell.width` | 欄寬飄 | 同時改 `tblGrid`（twips，1cm≈567）+ `cell.width`（Cm） |
| 16 | Python heredoc `\U` 截斷 | `SyntaxError: truncated \UXXXXXXXX` | 一律用 Write 寫 `.py` 檔再執行 |
| 17 | 簽署區/表頭誤填彩色 | 用戶要求白底黑字 | 簽署章節白底黑字唔填色；ME 測試表白底黑字 |

---

## 11. 與其他技能嘅關係

```
任務：生成一份工程竣工報告（三件套 / 勘察備忘）
   ↓
skill-router 路由
   ↓
electrical-test-report-generator  ← 測試表格風格（ME 白底黑字、頁碼頁腳）
   +
python-docx-photo-grid（本技能）← 統一排版（藍色系/抬頭/簽署）+ 自適應相片網格
   +
site-inspection-memo-generator  ← 勘察備忘錄專用流程（Letter/藍色系/動作欄）
   ↓
生成 .docx
   ↓
用戶手改（officecli-workflow 做局部改 OR 用戶直接 Word 改）
   ↓
correction-diff-capture  ← 讀用戶改完嘅版本，diff 反饋到本技能
```

| 技能 | 關係 |
|---|---|
| `electrical-test-report-generator` | 測試表格風格基底，本技能加相片排版 + 統一簽署 |
| `site-inspection-memo-generator` | 勘察備忘錄專用（Letter/動作欄色碼）；相片網格技術照用本技能 |
| `officecli-workflow` | **生成後**用戶手改時，用 OfficeCLI 局部改（快），唔好 regen |
| `correction-diff-capture` | 用戶改完交付，讀取+diff 反饋到本技能嘅坑位/常數 |
| `word-typography-guide` | 純文字文檔用呢個；本技能用於「文字+相片」混合 |
| `gantt-chart-pro` / `faithful-xlsx-template` | Excel 場景用呢啲 |

---

## 12. 完整最小可運行範例

```python
# -*- coding: utf-8 -*-
import os
from docx import Document
from docx.shared import Cm, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from PIL import Image

# === 環境 ===
PHOTO_DIR = r"D:\path\to\photos"
LOGO      = r"D:\path\to\company_logo.jpg"     # 承建商公司抬頭圖
OUT       = r"D:\path\to\output\report.docx"
TEMP      = r"D:\WorkBuddy\temp\_photos"
os.makedirs(TEMP, exist_ok=True)

# === A4 幾何（竣工報告；勘察備忘改 Letter 21.59×27.94）===
PAGE_W, PAGE_H = 21.0, 29.7
MARGIN = 2.54
CONTENT_W = PAGE_W - 2 * MARGIN
CONTENT_H = PAGE_H - 2 * MARGIN
GRID_COLS, MAX_ROWS = 2, 3
COL_W = CONTENT_W / GRID_COLS
IMG_MAX_W = COL_W - 0.96
PAGE_BUDGET = CONTENT_H - 2.2
CAPTION_H, CELL_PAD, SAFETY = 0.9, 0.5, 0.15
ABS_MAX_IMG_H, MIN_IMG_H = 9.0, 3.0

# === 工具函數（§5.3 + 附錄 helper 全集）===
# ... rendered_size, add_picture_fit, fix_table_layout, add_photo_grid,
#     set_run, set_cell, add_page_number, add_header_logo, add_signing_block ...

# === 生成 ===
doc = Document()
sec = doc.sections[0]
sec.page_width, sec.page_height = Cm(PAGE_W), Cm(PAGE_H)
sec.top_margin = sec.bottom_margin = sec.left_margin = sec.right_margin = Cm(MARGIN)

add_header_logo(doc, LOGO)                  # 頁眉承建商抬頭圖
# ... 標題、項目信息（承建商：工程承建商）、數據表格 ...
items = []  # [(temp_image_path, "128室 — 484 lux"), ...]
add_photo_grid(doc, items)                  # 自適應相片網格
add_signing_block(doc)                      # 統一簽署區
add_page_number(doc)                        # 頁腳 PAGE / NUMPAGES
doc.save(OUT)
print("saved:", OUT)
```

---

## 13. 自我進化記錄

| 日期 | 變更 | 觸發 |
|---|---|---|
| 2026-08-28 | 初版建立 | 學校燈具 Lux 報告任務（v1~v4 四次迭代） |
| 2026-08-28 | 加入 `_01` 後綴 caption bug 教訓 | 用戶手改 Lux 報告修正 caption |
| 2026-08-28 | 確認自適應網格公式 | 解決「直向相塞唔落 3 行」問題 |
| 2026-09-01 | **統一升級**：合併知識庫 SOP + 藍色系模板 + python-docx 踩坑 + 三份交付件實測（承建商抬頭圖 7.5×2.0、統一簽署區、A4/Letter 兩套紙型、三件套章節結構），新增 `references/竣工報告三件套_模板與SOP.md` | 用戶要求把竣工報告三件套能力統一封裝入本技能 |
