#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
香山國小社團申請文件生成器
用法：python3 generate.py <設定檔.json>
      python3 generate.py config_114下.json
"""

import json
import sys
import os
from datetime import date, timedelta
from pathlib import Path
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from copy import deepcopy

# ── 輔助函數 ────────────────────────────────────────────────────────────────

def roc_date(d):
    """西元日期 → 民國年 (e.g. 114/03/18)"""
    return f"{d.year - 1911}/{d.month:02d}/{d.day:02d}"

def short_date(d):
    """短格式 m/dd (e.g. 3/18)"""
    return f"{d.month}/{d.day:02d}"

def full_date_cjk(d):
    """中文完整日期 (e.g. 3月18日（三）)"""
    wk = ["一","二","三","四","五","六","日"][d.weekday()]
    return f"{d.month}月{d.day:02d}日（{wk}）"

def get_class_dates(first: date, count: int):
    return [first + timedelta(weeks=i) for i in range(count)]

def set_cell_bg(cell, hex_color):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    tcPr.append(shd)

def set_para_spacing(para, before=0, after=0):
    pPr = para._p.get_or_add_pPr()
    spacing = OxmlElement("w:spacing")
    spacing.set(qn("w:before"), str(before))
    spacing.set(qn("w:after"), str(after))
    pPr.append(spacing)

def add_bold_run(para, text, size=None, color=None):
    run = para.add_run(text)
    run.bold = True
    if size:
        run.font.size = Pt(size)
    if color:
        run.font.color.rgb = RGBColor(*bytes.fromhex(color))
    return run

def set_col_width(table, col_idx, width_cm):
    for row in table.rows:
        row.cells[col_idx].width = Cm(width_cm)

# ── 文件 1：上課時間表 ──────────────────────────────────────────────────────

def make_doc1(cfg, output_path):
    y = cfg["year"]
    sem = cfg["semester"]
    teacher_contact = cfg["school_contact"]
    schedule = cfg["schedule"]
    class_days = cfg["class_days"]          # list of {weekday, start_date, end_date, count, time}

    doc = Document()

    # 設定頁面邊距
    for section in doc.sections:
        section.top_margin    = Cm(2)
        section.bottom_margin = Cm(2)
        section.left_margin   = Cm(2.5)
        section.right_margin  = Cm(2.5)

    # 標題
    h1 = doc.add_paragraph()
    h1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = h1.add_run(f"{y}學年度{sem}學期課後社團上課時間")
    r.bold = True; r.font.size = Pt(16)

    h2 = doc.add_paragraph()
    h2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r2 = h2.add_run("課後社團準備日程表")
    r2.bold = True; r2.font.size = Pt(13)

    note = doc.add_paragraph()
    note.alignment = WD_ALIGN_PARAGRAPH.CENTER
    note.add_run("若有任何問題，請撥打電話或用line私訊我，感謝!\n")
    note.add_run(f"香山國小 學務處活動組 {teacher_contact}")

    doc.add_paragraph()  # 空行

    # 上課時間表格
    tbl1 = doc.add_table(rows=1+len(class_days), cols=3)
    tbl1.style = "Table Grid"
    hdr = tbl1.rows[0].cells
    for idx, h in enumerate(["", "日期", "時間"]):
        hdr[idx].text = h
        for p in hdr[idx].paragraphs:
            p.runs[0].bold = True if h else False
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_cell_bg(hdr[idx], "D9E1F2")

    for i, cd in enumerate(class_days):
        row = tbl1.rows[i+1].cells
        row[0].text = f"每週{cd['weekday']}的社團，共{cd['count']}次"
        row[1].text = f"{cd['start_date']}-{cd['end_date']}"
        row[2].text = cd["time"]
        for cell in row:
            cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_paragraph()  # 空行

    # 行程表格
    tbl2 = doc.add_table(rows=1+len(schedule), cols=2)
    tbl2.style = "Table Grid"
    hdr2 = tbl2.rows[0].cells
    hdr2[0].text = "日期"; hdr2[1].text = "工作內容"
    for cell in hdr2:
        for p in cell.paragraphs:
            p.runs[0].bold = True
        set_cell_bg(cell, "D9E1F2")

    for i, item in enumerate(schedule):
        row = tbl2.rows[i+1].cells
        row[0].text = item["date"]
        row[1].text = item["content"]

    doc.save(output_path)
    print(f"  ✓ {Path(output_path).name}")

# ── 文件 2：校外老師入校申請表 ───────────────────────────────────────────────

def make_doc2(cfg, output_path):
    y = cfg["year"]
    sem = cfg["semester"]
    t = cfg["teacher"]
    apply_date = cfg["apply_date"]   # e.g. "114/12/17"
    club = cfg["club_name"]
    course_time = cfg["course_time"] # e.g. "每週三12:50~14:30"
    course_weeks = cfg["course_weeks"] # e.g. "約12或13週課程"
    target = cfg["target_grades"]
    location = cfg["location"]

    doc = Document()
    for section in doc.sections:
        section.top_margin    = Cm(2)
        section.bottom_margin = Cm(2)
        section.left_margin   = Cm(2)
        section.right_margin  = Cm(2)

    # 附件標題
    h = doc.add_paragraph()
    h.alignment = WD_ALIGN_PARAGRAPH.CENTER
    h.add_run("（附件1）     新竹市香山國小校外人士協助教學或活動申請表").bold = True

    doc.add_paragraph()

    # 基本資訊表格
    info_data = [
        ["申請處室\n/班級", "學務處", "申請人", t["name"]],
        ["申請日期", f"中華民國   {apply_date.replace('/', '   年    ').replace('/', '   月    ')}    日", "", ""],
    ]
    tbl = doc.add_table(rows=2, cols=4)
    tbl.style = "Table Grid"
    for r_idx, row_data in enumerate(info_data):
        for c_idx, val in enumerate(row_data):
            cell = tbl.rows[r_idx].cells[c_idx]
            cell.text = val
            if c_idx in [0, 2]:
                cell.paragraphs[0].runs[0].bold = True
                set_cell_bg(cell, "F2F2F2")
    # 合併申請日期跨欄
    tbl.rows[1].cells[1].merge(tbl.rows[1].cells[3])

    doc.add_paragraph()

    # 協助教學人士資訊 (大表格)
    rows_data = [
        ("協助教學或\n活動人士",
         f"姓名：{t['name']}\n連絡電話：{t['phone']}\n個人學經歷：{t['company']} {t['title']}\n出生年月日：{t['dob']}\n身分證字號：{t['id']}",
         f"服務單位：{t['company']}\n\n\n是否施打過疫苗：█是      □否"),
        ("協助教學或\n活動人士資格",
         "█無「犯性侵害犯罪防治法第二條第一項所定之罪，經有罪判決確定」\n"
         "█無「受兒童及少年性剝削防制條例規定處罰，或受性騷擾防治法第二十條或第二十五條規定處罰」\n"
         "█無「經各級社政主管機關依兒童及少年福利與權益保障法第九十七條規定處罰」\n"
         "█無「曾體罰或霸凌學生，造成其身心嚴重侵害」\n"
         "█無「有性別平等教育法第二十七之一條第一項第一、二款及同條第三項之情形者」\n"
         "(曾犯任何1項，學校不得進用或運用)", ""),
        ("協助教學\n活動時間",
         f"{course_time}\n{course_weeks}", ""),
        ("教學對象與地點",
         f"國小{target}學生\n{location}", ""),
        ("課程大綱", cfg.get("course_summary", ""), ""),
        ("教材形式",
         "教學計畫書    □教學簡報    □印刷品   □影音光碟   □其他於課程或活動中使用之教學資料", ""),
        ("教材內容簡介", cfg.get("teaching_material", ""), ""),
        ("申請結果\n(由學校填寫)",
         "□通過。\n□修正後再審(請於＿＿年＿＿月＿＿日前提出修正資料)。\n□修正後通過。\n□不通過。", ""),
    ]

    tbl2 = doc.add_table(rows=len(rows_data), cols=3)
    tbl2.style = "Table Grid"
    # 設欄寬
    for i, (label, content, extra) in enumerate(rows_data):
        cells = tbl2.rows[i].cells
        cells[0].text = label
        cells[0].paragraphs[0].runs[0].bold = True
        set_cell_bg(cells[0], "F2F2F2")
        cells[1].text = content
        if extra:
            cells[2].text = extra
        else:
            cells[1].merge(cells[2])

    doc.add_paragraph()

    # 備註
    note_p = doc.add_paragraph()
    note_p.add_run("【備註】\n"
                   "    1. 校外人士協助教學或活動違反相關法規或本要點規定者，本校應終止契約關係或運用關係，並依相關法令處理。\n"
                   "    2. 本表由承辦處室或老師協助填寫。               "
                   f"申請人：＿＿＿＿＿＿＿(簽章)")

    doc.add_paragraph()
    sign_p = doc.add_paragraph()
    sign_p.add_run("   校長:                人事主任:              教務主任:              教學組:   ")

    doc.save(output_path)
    print(f"  ✓ {Path(output_path).name}")

# ── 文件 3：社團企劃書 ───────────────────────────────────────────────────────

def make_doc3(cfg, output_path):
    y = cfg["year"]
    sem = cfg["semester"]
    t = cfg["teacher"]
    club = cfg["club_name"]
    target = cfg["target_grades"]
    fee = cfg["fee"]
    course_time_display = cfg["course_time_display"]  # e.g. "每週三12:50~14:30"
    class_dates = get_class_dates(
        date.fromisoformat(cfg["first_class_date"]),
        cfg["total_classes"]
    )
    curriculum = cfg["curriculum"]
    location = cfg["location"]

    doc = Document()
    for section in doc.sections:
        section.top_margin    = Cm(1.8)
        section.bottom_margin = Cm(1.8)
        section.left_margin   = Cm(2.2)
        section.right_margin  = Cm(2.2)

    # ─ 師資簡介 ─
    h1 = doc.add_paragraph()
    h1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    h1.add_run(f"新竹市香山國小 {y}學年度 {sem}學期 學生課後社團師資簡介").bold = True

    tbl = doc.add_table(rows=7, cols=5)
    tbl.style = "Table Grid"

    def fill_row(row_idx, label, *values):
        cells = tbl.rows[row_idx].cells
        cells[0].text = label
        cells[0].paragraphs[0].runs[0].bold = True
        set_cell_bg(cells[0], "F2F2F2")
        if len(values) == 1:
            cells[1].merge(cells[4])
            cells[1].text = values[0]
        elif len(values) == 3:
            cells[1].text = values[0]
            cells[2].merge(cells[3])
            cells[2].text = values[1]
            cells[4].text = values[2]
        elif len(values) == 2:
            cells[1].merge(cells[2])
            cells[1].text = values[0]
            cells[3].merge(cells[4])
            cells[3].text = values[1]

    tbl.rows[0].cells[0].text = "請填寫實際授課教師資料"
    tbl.rows[0].cells[0].merge(tbl.rows[0].cells[0])  # placeholder
    # Re-build with correct structure
    tbl2 = doc.add_table(rows=7, cols=5)
    tbl2.style = "Table Grid"

    def set_cell(cell, text, bold=False, bg=None):
        cell.text = text
        if bold and cell.paragraphs[0].runs:
            cell.paragraphs[0].runs[0].bold = True
        if bg:
            set_cell_bg(cell, bg)

    rows_info = [
        ("請填寫實際授課教師資料", "教師姓名", t["name"], "性  別", "男"),
        ("請填寫實際授課教師資料", "出生年月日", t["dob"], t["dob"], t["dob"]),
        ("請填寫實際授課教師資料", "聯絡住址\n（含鄰里）",
         t["address"], t["address"], t["address"]),
        ("請填寫實際授課教師資料", "聯絡電話", t["phone"], t["phone"], t["phone"]),
        ("請填寫實際授課教師資料", "電子信箱", t["email"], t["email"], t["email"]),
        ("請填寫實際授課教師資料", "學經歷",
         t["experience"], t["experience"], t["experience"]),
        ("請填寫實際授課教師資料", "  身分證正反證照照片",
         "\n\n\n\n\n\n", "\n\n\n\n\n\n", "\n\n\n\n\n\n"),
    ]

    # Remove the dummy table
    tbl._element.getparent().remove(tbl._element)

    tbl_info = doc.add_table(rows=7, cols=5)
    tbl_info.style = "Table Grid"
    for r_idx, (col0, col1, col2, col3, col4) in enumerate(rows_info):
        cells = tbl_info.rows[r_idx].cells
        if r_idx == 0:
            # First column spans all rows (vertical)
            cells[0].text = col0
            cells[0].paragraphs[0].runs[0].bold = True
            set_cell_bg(cells[0], "F2F2F2")
        else:
            cells[0].text = ""
        cells[1].text = col1
        cells[1].paragraphs[0].runs[0].bold = True if col1 not in ["", " "] else False
        set_cell_bg(cells[1], "F2F2F2")
        if col2 == col3 == col4:
            cells[2].merge(cells[4])
            cells[2].text = col2
        else:
            cells[2].text = col2
            cells[3].text = col3
            cells[4].text = col4

    # Merge first column
    for r_idx in range(1, 7):
        tbl_info.rows[0].cells[0].merge(tbl_info.rows[r_idx].cells[0])

    doc.add_paragraph()

    # ─ 招生簡章 ─
    h2 = doc.add_paragraph()
    h2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    h2.add_run(f"新竹市香山國民小學  {y} 學年度{sem}學期").bold = True

    h3 = doc.add_paragraph()
    h3.alignment = WD_ALIGN_PARAGRAPH.CENTER
    h3.add_run(f"『{club}』招生簡章").bold = True

    items = [
        f"一、 招生對象：_{target}_ 不到 6 人不開課",
        "二、 報名方式：       (請貼上QR code) \n"
        f"報名費用__{fee}__    (已包含場租及耗材費用)\n"
        "請填上包含貴社繳交場租成本之要向學生收的學費+材料費 註明學費金額及材料費金額(繳交社團老師學校不經手)",
        "三、 上課日期：約開學後第三週開始為期12或13週(第一週報名、第二週第二次報名、第三週申請場地通知學生)\n"
        f"上課時間：{course_time_display}",
        "四、 上課地點： 學校總務處事務組安排\n（希望安排在中庭飛行）",
        "五、 上課內容如下：約12或13週課程簡介進度(依星期幾不同)",
        "六、 ★注意事項:",
        "七、 報名QR code (課程簡介.班級.姓名.家長電話.下課回家方式)\n  下課時請家長親自準時接回！",
    ]
    for item in items:
        p = doc.add_paragraph(item)
        p.paragraph_format.left_indent = Cm(0.5)

    doc.add_paragraph()

    # 課程特色
    features_title = doc.add_paragraph()
    features_title.add_run("課程特色介紹").bold = True

    features = [
        ("一、結合運動與科技",
         "以「足球競技」為核心，融合無人機操作，讓孩子在遊戲與對抗中自然學習科技概念，三軸空間感，提升學習動機。"),
        ("二、零基礎也能上手",
         "不需程式或無人機經驗，從操作、安全觀念開始，循序漸進，讓每位孩子都能成功參與。"),
        ("三、強調安全與規範",
         "使用足球防護框無人機，搭配完整飛行規則與場地規劃，培養正確使用科技的態度。"),
        ("四、培養專注與判斷力",
         "透過即時操控與比賽情境，訓練專注力、空間感與快速判斷能力。"),
        ("五、團隊合作與溝通",
         "設計雙人或小組任務，學習分工、討論與合作，強化人際互動能力。"),
        ("六、接軌競賽與未來發展",
         "課程內容銜接足球無人機競賽規則，為未來校際或全國競賽打下基礎。"),
        ("七、寓教於樂、成就感高",
         "每堂課都有明確任務與成果，讓孩子在完成挑戰中建立自信與成就感。"),
    ]
    for title, desc in features:
        p = doc.add_paragraph()
        p.add_run(f"{title}\n").bold = True
        p.add_run(desc)
        p.paragraph_format.left_indent = Cm(0.5)

    doc.add_paragraph()
    doc.add_paragraph()

    # ─ 課程進度表 ─
    prog_title = doc.add_paragraph()
    prog_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    prog_title.add_run(
        f"新竹市香山國小 {y}學年度 {sem}學期 {club} 社課程進度表/教學日誌"
    ).bold = True

    # 教學目標和課程內容簡介表格
    meta_tbl = doc.add_table(rows=2, cols=4)
    meta_tbl.style = "Table Grid"

    goal_cells = meta_tbl.rows[0].cells
    goal_cells[0].text = "教學目標"
    goal_cells[0].paragraphs[0].runs[0].bold = True
    set_cell_bg(goal_cells[0], "D9E1F2")
    goal_cells[0].merge(goal_cells[1])
    goal_cells[2].text = ("學生在每週課後社團中輕鬆接觸前瞻科技，逐步建立學習興趣與科技概念，不需具備程式基礎，"
                          "也能以安全、有系統的方式學習無人機操作。透過循序漸進的飛行訓練與團隊合作活動，"
                          "培養判斷力、觀察力與問題解決能力，並以足球無人機為主題進行規則與戰術入門，"
                          "未來更有機會代表學校參加足球無人機競賽，為校爭光。")
    goal_cells[2].merge(goal_cells[3])

    intro_cells = meta_tbl.rows[1].cells
    intro_cells[0].text = "課程內容簡介"
    intro_cells[0].paragraphs[0].runs[0].bold = True
    set_cell_bg(intro_cells[0], "D9E1F2")
    intro_cells[0].merge(intro_cells[1])
    intro_cells[2].text = ("AI足球無人機社團，以最火紅的電子競技〈足球無人機〉為出發點，"
                           "引導孩子在自主探索中理解無人機為何需要 AI，"
                           "以及哪些情境能由無人機協助提升生活的安全與效率。")
    intro_cells[2].merge(intro_cells[3])

    # 課程進度表格
    prog_tbl = doc.add_table(rows=1+len(class_dates), cols=4)
    prog_tbl.style = "Table Grid"

    hdr_cells = prog_tbl.rows[0].cells
    for ci, h in enumerate(["次數", "日期", "授課內容", "教師簽名"]):
        hdr_cells[ci].text = h
        hdr_cells[ci].paragraphs[0].runs[0].bold = True
        set_cell_bg(hdr_cells[ci], "D9E1F2")

    for i, d in enumerate(class_dates):
        row = prog_tbl.rows[i+1].cells
        row[0].text = str(i+1)
        row[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        row[1].text = full_date_cjk(d)
        row[2].text = curriculum[i] if i < len(curriculum) else ""
        row[3].text = ""

    doc.save(output_path)
    print(f"  ✓ {Path(output_path).name}")

# ── 文件 4：入校須知 ─────────────────────────────────────────────────────────

def make_doc4(cfg, output_path):
    t = cfg["teacher"]
    doc = Document()
    for section in doc.sections:
        section.top_margin    = Cm(2)
        section.bottom_margin = Cm(2)
        section.left_margin   = Cm(2.5)
        section.right_margin  = Cm(2.5)

    h = doc.add_paragraph()
    h.alignment = WD_ALIGN_PARAGRAPH.CENTER
    h.add_run("香山國小校外人士協助教學或活動入校須知範本").bold = True

    doc.add_paragraph(
        "歡迎您至本校協助教學或活動，基於維護學生權益，請您詳閱本須知內容，"
        "並於下方簽名確認，感謝您的配合！"
    )
    doc.add_paragraph()

    # 資格表格
    qual_title = doc.add_paragraph()
    qual_title.add_run("一、資格").bold = True

    qual_items = [
        "曾犯性侵害犯罪防治法第二條第一項所定之罪，經有罪判決確定",
        "曾受兒童及少年性剝削防制條例規定處罰，或受性騷擾防治法第二十條或第二十五條規定處罰",
        "曾經各級社政主管機關依兒童及少年福利與權益保障法第九十七條規定處罰",
        "曾體罰或霸凌學生，造成其身心嚴重侵害",
        "有性別平等教育法第二十七之一條第一項第一、二款及同條第三項之情形者",
    ]

    qual_tbl = doc.add_table(rows=1+len(qual_items), cols=4)
    qual_tbl.style = "Table Grid"
    hdr = qual_tbl.rows[0].cells
    for ci, h_text in enumerate(["一、資格", "自我檢核", "", "備註"]):
        hdr[ci].text = h_text
        if h_text:
            hdr[ci].paragraphs[0].runs[0].bold = True
        set_cell_bg(hdr[ci], "F2F2F2")

    for i, item in enumerate(qual_items):
        row = qual_tbl.rows[i+1].cells
        row[0].text = item
        row[1].text = "□是"
        row[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        row[2].text = "■否"
        row[2].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        if i == 0:
            row[3].text = "任何1項勾選「是」，學校不得進用或運用"

    doc.add_paragraph()

    # 義務及重要事項
    duty_title = doc.add_paragraph()
    duty_title.add_run("二、義務及重要事項").bold = True

    duty_items = [
        "需遵守十二年國民基本教育課程綱要總綱與各領綱規定",
        "需遵守相關法規(如教育基本法、性別平等教育法)及國際人權公約(如消除對婦女一切形式歧視公約、兒童權利公約、身心障礙者權利公約)之規定",
        "不得為特定政治團體或宗教信仰從事宣傳或活動",
        "不得有商業或為其他利益衝突之行為",
        "需遵守學校訂定之規章，並尊重學生之權利",
        "需參與地方教育主管機關或學校所提供之教育訓練",
        "原授課教師為學校課程之主要授課者，校外人士係為協助教學之角色",
        "本校由○○處○○組負責校外人士協助教學或活動及家長諮詢或申訴之相關事項",
        "校外人士協助教學或活動違反相關法規或本要點規定者，本校應終止契約關係或運用關係，並依相關法令處理",
    ]

    duty_tbl = doc.add_table(rows=1+len(duty_items), cols=3)
    duty_tbl.style = "Table Grid"
    duty_hdr = duty_tbl.rows[0].cells
    for ci, h_text in enumerate(["二、義務及重要事項", "檢視確認", ""]):
        duty_hdr[ci].text = h_text
        if h_text:
            duty_hdr[ci].paragraphs[0].runs[0].bold = True
        set_cell_bg(duty_hdr[ci], "F2F2F2")

    understand_items = {6, 7, 8}
    for i, item in enumerate(duty_items):
        row = duty_tbl.rows[i+1].cells
        row[0].text = item
        row[1].text = "■瞭解" if i in understand_items else "■可以"
        row[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        if i == 0:
            row[2].text = "任何1項未勾選，學校不予進用或運用"
        else:
            row[2].text = ""
        if i < len(duty_items)-1:
            row[2].merge(duty_tbl.rows[i+2].cells[2]) if i == 0 else None

    doc.add_paragraph()
    sign_p = doc.add_paragraph()
    sign_p.add_run(f"簽名：___________________________")

    doc.save(output_path)
    print(f"  ✓ {Path(output_path).name}")

# ── 主程式 ───────────────────────────────────────────────────────────────────

def main():
    if len(sys.argv) < 2:
        print("用法：python3 generate.py <設定檔.json>")
        sys.exit(1)

    config_path = sys.argv[1]
    with open(config_path, "r", encoding="utf-8") as f:
        cfg = json.load(f)

    y = cfg["year"]
    sem = cfg["semester"]
    folder_name = f"{y}{sem}學期申請文件"
    base_dir = Path(__file__).parent
    out_dir = base_dir / folder_name
    out_dir.mkdir(exist_ok=True)

    print(f"\n🏫 正在產生 {y}學年度{sem}學期 申請文件...")

    make_doc1(cfg, str(out_dir / f"1.{y}學年度{sem}學期課後社團上課時間.docx"))
    make_doc2(cfg, str(out_dir / f"2.{y}學年度{sem}學期校外老師入校協助教學申請表.docx"))
    make_doc3(cfg, str(out_dir / f"3.{y}學年度{sem}學期香山國小社團企劃書.docx"))
    make_doc4(cfg, str(out_dir / f"4.校外人士入校須知.docx"))

    print(f"\n✅ 完成！文件已儲存至：{folder_name}/")

if __name__ == "__main__":
    main()
