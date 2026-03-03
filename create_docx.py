from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import copy

doc = Document()

# ページ余白設定
section = doc.sections[0]
section.top_margin = Cm(2.5)
section.bottom_margin = Cm(2.5)
section.left_margin = Cm(3)
section.right_margin = Cm(2.5)

def set_font(run, size=11, bold=False, color=None):
    run.font.name = '游明朝'
    run.font.size = Pt(size)
    run.font.bold = bold
    r = run._r
    rPr = r.get_or_add_rPr()
    rFonts = OxmlElement('w:rFonts')
    rFonts.set(qn('w:eastAsia'), '游明朝')
    rPr.insert(0, rFonts)
    if color:
        run.font.color.rgb = RGBColor(*color)

def add_paragraph(text, size=11, bold=False, align=WD_ALIGN_PARAGRAPH.LEFT, space_before=0, space_after=6):
    p = doc.add_paragraph()
    p.alignment = align
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after = Pt(space_after)
    run = p.add_run(text)
    set_font(run, size=size, bold=bold)
    return p

# ========== タイトル ==========
add_paragraph('稟　議　書', size=18, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, space_before=6, space_after=12)

# ========== 基本情報テーブル ==========
info_table = doc.add_table(rows=4, cols=2)
info_table.style = 'Table Grid'
info_data = [
    ('起案日', '令和７年３月３日'),
    ('起案者', '　　　　　（教諭）'),
    ('所　属', '○○農業高等学校'),
    ('提出先', '校　長　殿'),
]
for i, (label, value) in enumerate(info_data):
    row = info_table.rows[i]
    row.cells[0].width = Cm(3)
    row.cells[1].width = Cm(11)
    for cell, text in [(row.cells[0], label), (row.cells[1], value)]:
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p = cell.paragraphs[0]
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(2)
        run = p.add_run(text)
        set_font(run, size=10.5, bold=(cell == row.cells[0]))

doc.add_paragraph('')

# ========== 件名 ==========
add_paragraph('件　名', size=11, bold=True, space_before=4, space_after=2)
add_paragraph('生成AIを活用した業務効率化研修会の実施について', size=12, bold=True,
              align=WD_ALIGN_PARAGRAPH.CENTER, space_before=2, space_after=10)

# 区切り線
p = doc.add_paragraph('─' * 40)
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(8)

# ========== 1. 目的 ==========
add_paragraph('１．目　的', size=11, bold=True, space_before=4, space_after=4)
add_paragraph(
    '　近年、生成AI技術の急速な普及に伴い、教育現場においてもその活用が求められています。'
    '本研修会では、教職員が生成AIを使った業務効率化の手法を習得し、日常業務の時間短縮を図ることを目的とします。',
    size=10.5, space_after=8
)

# ========== 2. 概要テーブル ==========
add_paragraph('２．概　要', size=11, bold=True, space_before=4, space_after=4)
summary_data = [
    ('研修名', '業務効率化・1時間早く帰れる生成AI術'),
    ('日　時', '令和７年３月９日（月）　16:00〜17:00（60分）'),
    ('対　象', '教職員（任意参加）'),
    ('場　所', '　　　　　（会場は後日連絡）'),
    ('講　師', '　　　　　教諭（自主研修形式）'),
    ('参加費', '無料'),
]
tbl = doc.add_table(rows=len(summary_data), cols=2)
tbl.style = 'Table Grid'
for i, (label, value) in enumerate(summary_data):
    row = tbl.rows[i]
    for cell, text, bold in [(row.cells[0], label, True), (row.cells[1], value, False)]:
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p = cell.paragraphs[0]
        p.paragraph_format.space_before = Pt(3)
        p.paragraph_format.space_after = Pt(3)
        run = p.add_run(text)
        set_font(run, size=10.5, bold=bold)

doc.add_paragraph('')

# ========== 3. 研修内容 ==========
add_paragraph('３．研修内容', size=11, bold=True, space_before=4, space_after=4)
add_paragraph('　以下の４テーマについて、実際に操作しながら学びます。', size=10.5, space_after=4)
items = [
    ('①　ルーブリック評価表の作成', '生成AIを使い、観点別評価規準を効率的に作成する方法'),
    ('②　稟議書・公文書の作成', '定型文書を素早く、高品質に仕上げるプロンプト活用術'),
    ('③　授業案の作成', '学習指導要領に沿った授業展開案を短時間で作成する方法'),
    ('④　ワークシートの作成', '生徒の実態に応じたワークシートをAIで効率的に作成する方法'),
]
for title, desc in items:
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(1)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.left_indent = Cm(0.5)
    r1 = p.add_run(f'{title}　')
    set_font(r1, size=10.5, bold=True)
    r2 = p.add_run(desc)
    set_font(r2, size=10.5)

doc.add_paragraph('')

# ========== 4. 期待効果 ==========
add_paragraph('４．期待される効果', size=11, bold=True, space_before=4, space_after=4)
effects = [
    '教材・文書作成にかかる時間の大幅な短縮（目標：1時間以上／日）',
    'AIツールへの理解促進と学校全体の業務効率化',
    '教職員のワークライフバランスの改善',
]
for e in effects:
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(1)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.left_indent = Cm(0.5)
    run = p.add_run(f'・　{e}')
    set_font(run, size=10.5)

doc.add_paragraph('')

# ========== 5. 費用 ==========
add_paragraph('５．費　用', size=11, bold=True, space_before=4, space_after=4)
add_paragraph('　なし（既存のPCおよび無料AIツールを使用）', size=10.5, space_after=8)

# ========== 6. その他 ==========
add_paragraph('６．その他', size=11, bold=True, space_before=4, space_after=4)
notes = [
    '参加は任意とし、事前申し込みは不要です。',
    '使用するAIツール：ChatGPT（無料版）等、アカウント不要でも参加可能',
]
for n in notes:
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(1)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.left_indent = Cm(0.5)
    run = p.add_run(f'・　{n}')
    set_font(run, size=10.5)

doc.add_paragraph('')

# ========== 承認欄 ==========
add_paragraph('承　認　欄', size=11, bold=True, space_before=8, space_after=4)
approval_table = doc.add_table(rows=3, cols=4)
approval_table.style = 'Table Grid'
headers = ['校　長', '教　頭', '教務主任', '起案者']
for i, h in enumerate(headers):
    cell = approval_table.rows[0].cells[i]
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(3)
    run = p.add_run(h)
    set_font(run, size=10.5, bold=True)
# 空白行を２行
for row in approval_table.rows[1:]:
    for cell in row.cells:
        p = cell.paragraphs[0]
        p.paragraph_format.space_before = Pt(12)
        p.paragraph_format.space_after = Pt(12)

# 起案日（末尾）
doc.add_paragraph('')
add_paragraph('令和７年３月３日　起案', size=10, align=WD_ALIGN_PARAGRAPH.RIGHT, space_before=4)

# 保存
output_path = '/home/user/normal-task/稟議書_生成AI研修.docx'
doc.save(output_path)
print(f'保存完了: {output_path}')
