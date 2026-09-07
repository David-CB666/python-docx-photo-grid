# 竣工報告三件套 模板與 SOP（python-docx-photo-grid 附錄）

> 本附錄合併三份權威來源，供 `python-docx-photo-grid` 技能主文件引用：
> ① 知識庫 `流程/工程勘察備忘與竣工報告生成_SOP.md`
> ② 知識庫 `模板/DOCX_TEMPLATE_藍色系勘察備忘與竣工報告.md`
> ③ 知識庫 `经验/python-docx工程文檔踩坑.md`
> ④ **項目交付件實測** → 統一簽署/頁眉/紙型/章節參數以此為準。

---

## 一、9 步工作流程（每份報告都走）

1. **讀路徑** — `ls` 睇用戶文件夾有咩（照片 / PDF / 文檔 / 手稿）。
2. **試讀圖** — 用 Read 工具讀 jpg，確認讀圖能力是否可用（極不穩定，見踩坑）。
3. **掌握規範** — 讀本附錄 + 前序交付件（上一份報告就係最好嘅模板）。
4. **寫生成腳本** — 到 `<TEMP_DIR>/build_xxx.py`，復用本附錄 helper。
5. **執行** — managed venv：
   ```
   "<WORKBUDDY_DIR>/binaries/python/envs/default/Scripts/python.exe" "<TEMP_DIR>/build_xxx.py"
   ```
6. **驗證** — 讀回 docx 檢查 tables/images/頁面設置/章節標題/簽署區（見主文件 §8）。
7. **交付** — `present_files` 呈現 docx。
8. **寫記憶** — daily log + 長期入 MEMORY.md。
9. **用戶手改後** — 只讀唔改（python-docx 讀取掌握），等明確指示先改。

---

## 二、紙型（兩套，唔好撈亂）⚠️

| 文檔類型 | 紙型 | 邊距 | 來源 |
|:---|:---|:---|:---|
| **勘察備忘錄** | **Letter 21.59 × 27.94 cm** | 上/下 2.6、左/右 2.8 | 用戶手定稿 |
| **竣工報告三件套** | **A4 21.0 × 29.7 cm** | 左 1.5（固定）；上 2.54、下 1.2、右 1.1 | 三份交付件實測 |

> 竣工報告邊距實測（三份略有差異，係用戶 Word 手改所致）：**統一建議：左 1.5 / 右 1.1 / 上 2.54 / 下 1.2。**

---

## 三、配色（藍色系，一文一色系）

| 用途 | 色碼 | 說明 |
|:---|:---|:---|
| 深藍 DEEP_BLUE | `#1F4E79` | 封面標題、章節標題、表頭底色（白字）|
| 中藍 MED_BLUE | `#2E75B6` | 副標題、次級強調 |
| 淺藍 LIGHT_BLUE | `#D6E4F0` | 動作欄底色（遷移）|
| 綠 GREEN | `#C6EFCE` | 驗收「合格」標記 |
| 橙 ORANGE | `#FCE4D6` | 取消/待定/分離 動作欄 |
| 紅 RED | `#F8CBAD` | 獨立供電 動作欄 |
| 灰 GRAY | `#808080` | 圖注/caption 文字 |
| 斑馬紋 | `#F2F2F2` | 表格交替行 |

---

## 四、字體 / 字號

- 中文 eastAsia：**PMingLiU**；ascii/hAnsi：**Times New Roman**（`set_run` 三屬性一齊設）。
- 封面標題 22pt Arial 粗體；章節標題 16pt 粗體；子標題 13pt；正文 12pt；表格 10pt；圖注 10pt 灰。
- 頁碼：footer 用 `PAGE / NUMPAGES` 域（OxmlElement `w:fldChar` 三件套），唔好用靜態數字。

---

## 五、頁眉：承建商公司抬頭圖（權威參數）

- **位置**：section header 第一個段落，段落 run 內嵌公司 logo 圖片。
- **實際渲染尺寸**（三份交付件實測）：**寬 6.5 ~ 7.7 cm × 高 1.8 ~ 2.1 cm**。
  - **統一建議：7.5 × 2.0 cm**。
- ⚠️ 首次插入用源圖 15.92 × 4.32 cm（= A4 CONTENT_W）會**過大**，正文會被壓；**插入後必須縮到 ~7.5×2.0**，或同步把上邊距調大到 ~4.6 cm。最終交付版係細圖 + 上邊距 2.25~2.54。
- HeaderDist：1.27 cm（或 0）均可，圖片置左。

```python
def add_header_logo(doc, logo_path, width_cm=7.5, height_cm=2.0):
    """頁眉插入承建商抬頭圖（居中或置左）。"""
    sec = doc.sections[0]
    hdr = sec.header
    p = hdr.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run()
    run.add_picture(logo_path, width=Cm(width_cm), height=Cm(height_cm))
    # 頁眉圖較高時，需加大上邊距避免正文被壓：
    # sec.top_margin = Cm(4.6)  # 只有用大圖時先需要
```

---

## 六、統一簽署區（權威格式，三份報告一致）

最後一章固定係「簽署」，格式如下（**Lux 版係權威，已套用到燈具/風扇**）：

```
[章節標題 16pt 粗體]  簽署

測試/安裝人員：工程承建商　　　日期：2026-08-28
業主/用家：________________　　　日期：________________
```

- 上方另有「承建商：工程承建商」（報告開頭資訊區）。
- 簽署章節用 **白底黑字**（唔填色），正文 12pt PMingLiU。

```python
def add_signing_block(doc):
    """統一簽署區（三件套通用）。"""
    doc.add_page_break()
    h = doc.add_paragraph()
    set_run(h.add_run("簽署"), size=16, bold=True, color=DEEP_BLUE)
    h.paragraph_format.space_before = Pt(18)

    p1 = doc.add_paragraph()
    set_run(p1.add_run("測試/安裝人員：工程承建商　　　日期：________________"),
            size=12)
    p1.paragraph_format.space_before = Pt(24)

    p2 = doc.add_paragraph()
    set_run(p2.add_run("業主/用家：________________　　　日期：________________"), size=12)
    p2.paragraph_format.space_before = Pt(18)
```

---

## 七、章節結構 — 竣工報告三件套（標準模板）

| 報告 | 章節結構 |
|:---|:---|
| **燈具安裝完成報告** | 標題 → 一、工程概況 → 二、燈具安裝完成明細 → 三、現場安裝照片 → 簽署 |
| **Lux 照度報告** | 標題 → 一、測試概述 → 二、標準依據 → 三、照度測試數據 → 四、測試點照片 → 簽署 |
| **風扇安裝完成報告** | 標題 → 一、工程概況 → 二、風扇安裝完成明細 → 三、現場安裝照片 → 簽署 |

相片 caption 規範（Lux 版）：`"128室 — 484 lux"` = `位置 — 數值 lux`，數值來自**結構化數據表平均值**（唔好由檔名解析，見主文件 §5）。

勘察備忘錄章節（藍色系）：一、備忘目的 → 二、電箱總覽 → 三、分路明細 → 四、勘察結果摘要 → 五、後續動作 → **六、前期備忘**（⚠️ 唔係「校方須知」）→ 七、勘察手稿記錄。

---

## 八、python-docx 完整 helper（三件套 + 勘察備忘通用）

```python
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from PIL import Image
import io, os

DEEP_BLUE = RGBColor(0x1F, 0x4E, 0x79)
MED_BLUE  = RGBColor(0x2E, 0x75, 0xB6)
GRAY      = RGBColor(0x80, 0x80, 0x80)

def set_run(run, size=12, bold=False, color=None, font_cn='PMingLiU', font_en='Times New Roman'):
    run.font.name = font_en
    run.font.size = Pt(size)
    run.font.bold = bold
    if color is not None:
        run.font.color.rgb = color
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts'); rPr.append(rFonts)
    rFonts.set(qn('w:ascii'), font_en)
    rFonts.set(qn('w:hAnsi'), font_en)
    rFonts.set(qn('w:eastAsia'), font_cn)

def set_cell_shading(cell, color_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd'); shd.set(qn('w:fill'), color_hex); tcPr.append(shd)

def set_cell(cell, text, size=10, align=WD_ALIGN_PARAGRAPH.LEFT, bold=False):
    cell.text = ''
    p = cell.paragraphs[0]; p.alignment = align
    run = p.add_run(text); set_run(run, size=size, bold=bold)
    for r in list(p.runs)[1:]:
        r._element.getparent().remove(r._element)

def set_table_borders(table):
    tbl = table._tbl; tblPr = tbl.tblPr
    borders = OxmlElement('w:tblBorders')
    for edge in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        el = OxmlElement(f'w:{edge}')
        el.set(qn('w:val'), 'single'); el.set(qn('w:sz'), '4')
        el.set(qn('w:space'), '0'); el.set(qn('w:color'), '000000')
        borders.append(el)
    tblPr.append(borders)

def add_page_number(doc):
    section = doc.sections[0]; footer = section.footer
    p = footer.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    def _fld(p, instr):
        r = p.add_run()
        f1 = OxmlElement('w:fldChar'); f1.set(qn('w:fldCharType'), 'begin')
        it = OxmlElement('w:instrText'); it.set(qn('xml:space'), 'preserve'); it.text = instr
        f2 = OxmlElement('w:fldChar'); f2.set(qn('w:fldCharType'), 'end')
        r._element.append(f1); r._element.append(it); r._element.append(f2)
    _fld(p, ' PAGE '); p.add_run(' / '); _fld(p, ' NUMPAGES ')

def add_photo(doc, path, width_cm=12, caption=None):
    img = Image.open(path).convert('RGB'); dpi = 200
    tw = int(width_cm / 2.54 * dpi); th = int(img.height * (tw / img.width))
    mh = int(16 / 2.54 * dpi)
    if th > mh: th = mh; tw = int(img.width * (mh / img.height))
    img = img.resize((tw, th), Image.LANCZOS)
    tmp = io.BytesIO(); img.save(tmp, format='JPEG', quality=85); tmp.seek(0)
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run().add_picture(tmp, width=Cm(tw / dpi * 2.54))
    if caption:
        cp = doc.add_paragraph(); cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_run(cp.add_run(caption), size=10, color=GRAY)

# 相片網格（2 列無框線，一頁 2-4 張；大量相用主文件 §4 自適應網格）
def add_photo_grid(doc, items, cols=2, width_cm=8.0, max_h_cm=9.0):
    for start in range(0, len(items), cols):
        group = items[start:start + cols]
        table = doc.add_table(rows=1, cols=cols)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        tblPr = table._tbl.tblPr
        borders = OxmlElement('w:tblBorders')
        for edge in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
            el = OxmlElement(f'w:{edge}')
            el.set(qn('w:val'), 'none'); el.set(qn('w:sz'), '0')
            borders.append(el)
        tblPr.append(borders)
        for ci in range(cols):
            cell = table.cell(0, ci); cell.width = Cm(width_cm + 0.3)
            if ci < len(group):
                path, caption = group[ci]
                p = cell.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                if os.path.exists(path):
                    img = Image.open(path).convert('RGB'); dpi = 200
                    tw = int(width_cm / 2.54 * dpi); th = int(img.height * (tw / img.width))
                    mh = int(max_h_cm / 2.54 * dpi)
                    if th > mh: th = mh; tw = int(img.width * (mh / img.height))
                    img = img.resize((tw, th), Image.LANCZOS)
                    tmp = io.BytesIO(); img.save(tmp, format='JPEG', quality=85); tmp.seek(0)
                    p.add_run().add_picture(tmp, width=Cm(tw / dpi * 2.54))
                cp = cell.add_paragraph(); cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
                set_run(cp.add_run(caption), size=9, color=GRAY)
```

> 表格欄寬規則：`set_col_widths()` 必須**同時設 `tblGrid`（twips，1cm≈567）+ `cell.width`（Cm）**，否則欄寬會飄。
