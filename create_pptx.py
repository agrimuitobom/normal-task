from pptx import Presentation
from pptx.util import Inches, Pt, Cm
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Emu

# ========== カラーパレット ==========
COLOR_BG       = RGBColor(0x1A, 0x1A, 0x2E)   # 紺（背景）
COLOR_ACCENT   = RGBColor(0x16, 0x21, 0x3E)   # やや明るい紺
COLOR_KEY      = RGBColor(0xE9, 0x4C, 0x4C)   # 赤（アクセント）
COLOR_WHITE    = RGBColor(0xFF, 0xFF, 0xFF)
COLOR_LIGHT    = RGBColor(0xE0, 0xE0, 0xE8)   # 薄グレー
COLOR_YELLOW   = RGBColor(0xFF, 0xD7, 0x00)   # 黄

FONT_MAIN = '游ゴシック'
W = Inches(13.33)   # ワイド16:9
H = Inches(7.5)

prs = Presentation()
prs.slide_width  = W
prs.slide_height = H

BLANK = prs.slide_layouts[6]  # blank

# ---------- ユーティリティ ----------

def add_rect(slide, x, y, w, h, fill_color=None, alpha=None):
    shape = slide.shapes.add_shape(1, x, y, w, h)   # MSO_SHAPE_TYPE.RECTANGLE = 1
    shape.line.fill.background()
    shape.line.color.rgb = RGBColor(0, 0, 0)
    shape.line.width = 0
    if fill_color:
        shape.fill.solid()
        shape.fill.fore_color.rgb = fill_color
    else:
        shape.fill.background()
    return shape

def add_textbox(slide, x, y, w, h, text, font_size=18, bold=False,
                color=COLOR_WHITE, align=PP_ALIGN.LEFT, font=FONT_MAIN,
                word_wrap=True):
    txBox = slide.shapes.add_textbox(x, y, w, h)
    tf = txBox.text_frame
    tf.word_wrap = word_wrap
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.name = font
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.color.rgb = color
    return txBox

def slide_bg(slide, color=COLOR_BG):
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = color


# ======================================================
# スライド 1 ── タイトル
# ======================================================
slide1 = prs.slides.add_slide(BLANK)
slide_bg(slide1)

# 上部アクセントバー
add_rect(slide1, 0, 0, W, Inches(0.12), fill_color=COLOR_KEY)

# 中央帯
add_rect(slide1, 0, Inches(2.6), W, Inches(2.5), fill_color=COLOR_ACCENT)

# タイトル
add_textbox(slide1, Inches(1), Inches(2.75),
            Inches(11.33), Inches(1.2),
            '業務効率化・1時間早く帰れる\n生成AI術',
            font_size=38, bold=True, color=COLOR_WHITE, align=PP_ALIGN.CENTER)

# サブタイトル
add_textbox(slide1, Inches(1), Inches(4.05),
            Inches(11.33), Inches(0.6),
            '教職員向け研修会　　令和7年3月9日（月）16:00〜17:00',
            font_size=20, color=COLOR_LIGHT, align=PP_ALIGN.CENTER)

# 下部ライン
add_rect(slide1, Inches(3), Inches(4.8), Inches(7.33), Inches(0.04), fill_color=COLOR_KEY)

# 主催者
add_textbox(slide1, Inches(1), Inches(5.0),
            Inches(11.33), Inches(0.5),
            '○○農業高等学校　　主催',
            font_size=16, color=COLOR_LIGHT, align=PP_ALIGN.CENTER)

# 下アクセントバー
add_rect(slide1, 0, H - Inches(0.12), W, Inches(0.12), fill_color=COLOR_KEY)


# ======================================================
# スライド 2 ── アジェンダ
# ======================================================
slide2 = prs.slides.add_slide(BLANK)
slide_bg(slide2)
add_rect(slide2, 0, 0, W, Inches(0.12), fill_color=COLOR_KEY)
add_rect(slide2, 0, H - Inches(0.12), W, Inches(0.12), fill_color=COLOR_KEY)

# ヘッダー帯
add_rect(slide2, 0, Inches(0.12), W, Inches(1.1), fill_color=COLOR_ACCENT)
add_textbox(slide2, Inches(0.5), Inches(0.2), Inches(12), Inches(0.9),
            '本日のアジェンダ', font_size=28, bold=True, color=COLOR_WHITE)

agenda_items = [
    ('16:00', '① イントロ・研修の目的'),
    ('16:10', '② ルーブリック評価表の作成'),
    ('16:22', '③ 稟議書・公文書の作成'),
    ('16:34', '④ 授業案の作成'),
    ('16:46', '⑤ ワークシートの作成'),
    ('16:55', '⑥ Q&A・まとめ'),
]

for i, (time, item) in enumerate(agenda_items):
    y = Inches(1.4) + i * Inches(0.87)
    # 時刻バッジ
    add_rect(slide2, Inches(0.5), y, Inches(1.1), Inches(0.6), fill_color=COLOR_KEY)
    add_textbox(slide2, Inches(0.5), y + Pt(4), Inches(1.1), Inches(0.55),
                time, font_size=15, bold=True, color=COLOR_WHITE, align=PP_ALIGN.CENTER)
    # 項目
    add_textbox(slide2, Inches(1.8), y, Inches(10.5), Inches(0.65),
                item, font_size=20, color=COLOR_WHITE)
    # 区切り線
    if i < len(agenda_items) - 1:
        add_rect(slide2, Inches(0.5), y + Inches(0.68),
                 Inches(11.8), Inches(0.02), fill_color=RGBColor(0x44, 0x44, 0x66))


# ======================================================
# スライド 3 ── 研修の目的
# ======================================================
slide3 = prs.slides.add_slide(BLANK)
slide_bg(slide3)
add_rect(slide3, 0, 0, W, Inches(0.12), fill_color=COLOR_KEY)
add_rect(slide3, 0, H - Inches(0.12), W, Inches(0.12), fill_color=COLOR_KEY)
add_rect(slide3, 0, Inches(0.12), W, Inches(1.1), fill_color=COLOR_ACCENT)
add_textbox(slide3, Inches(0.5), Inches(0.2), Inches(12), Inches(0.9),
            'なぜ今、生成AI？', font_size=28, bold=True, color=COLOR_WHITE)

problems = [
    ('😓', '問題', '教材・文書作成に毎日1〜2時間以上かかっている'),
    ('🤖', '解決策', '生成AIを使えば同じ品質の成果物を10〜20分で作れる'),
    ('🎯', '目標', '本研修後、今日から業務時間を1時間以上短縮する'),
]

for i, (icon, label, desc) in enumerate(problems):
    y = Inches(1.5) + i * Inches(1.6)
    add_rect(slide3, Inches(0.4), y, Inches(0.7), Inches(1.1),
             fill_color=COLOR_KEY)
    add_textbox(slide3, Inches(0.4), y, Inches(0.7), Inches(1.1),
                icon, font_size=26, align=PP_ALIGN.CENTER, color=COLOR_WHITE)
    add_rect(slide3, Inches(1.2), y, Inches(1.2), Inches(1.1),
             fill_color=COLOR_ACCENT)
    add_textbox(slide3, Inches(1.2), y, Inches(1.2), Inches(1.1),
                label, font_size=18, bold=True, align=PP_ALIGN.CENTER, color=COLOR_YELLOW)
    add_textbox(slide3, Inches(2.6), y + Inches(0.22), Inches(10), Inches(0.7),
                desc, font_size=20, color=COLOR_WHITE)


# ======================================================
# スライド 4 ── ルーブリック評価表
# ======================================================
def content_slide(prs, title, subtitle, steps, note=''):
    slide = prs.slides.add_slide(BLANK)
    slide_bg(slide)
    add_rect(slide, 0, 0, W, Inches(0.12), fill_color=COLOR_KEY)
    add_rect(slide, 0, H - Inches(0.12), W, Inches(0.12), fill_color=COLOR_KEY)
    add_rect(slide, 0, Inches(0.12), W, Inches(1.3), fill_color=COLOR_ACCENT)
    add_textbox(slide, Inches(0.5), Inches(0.18), Inches(12), Inches(0.65),
                title, font_size=26, bold=True, color=COLOR_WHITE)
    add_textbox(slide, Inches(0.5), Inches(0.8), Inches(12), Inches(0.5),
                subtitle, font_size=16, color=COLOR_LIGHT)

    for i, (num, step_title, detail) in enumerate(steps):
        y = Inches(1.6) + i * Inches(1.15)
        add_rect(slide, Inches(0.4), y, Inches(0.55), Inches(0.55),
                 fill_color=COLOR_KEY)
        add_textbox(slide, Inches(0.4), y, Inches(0.55), Inches(0.55),
                    num, font_size=18, bold=True, align=PP_ALIGN.CENTER, color=COLOR_WHITE)
        add_textbox(slide, Inches(1.1), y, Inches(4), Inches(0.55),
                    step_title, font_size=18, bold=True, color=COLOR_YELLOW)
        add_textbox(slide, Inches(1.1), y + Inches(0.52), Inches(11.5), Inches(0.55),
                    detail, font_size=16, color=COLOR_LIGHT)

    if note:
        add_rect(slide, Inches(0.4), Inches(6.5), Inches(12.4), Inches(0.7),
                 fill_color=RGBColor(0x2A, 0x2A, 0x4A))
        add_textbox(slide, Inches(0.6), Inches(6.55), Inches(12), Inches(0.6),
                    f'💡 {note}', font_size=15, color=COLOR_YELLOW)
    return slide

content_slide(prs,
    '① ルーブリック評価表の作成',
    '観点別評価規準をAIで一瞬に作る',
    [
        ('1', 'プロンプトを入力する',
         '「農業科○○単元のルーブリック評価表を4観点で作って」と入力'),
        ('2', 'AIが表を生成',
         '知識・技能・思考・主体性の4列×5段階の評価表が数秒で完成'),
        ('3', '修正・微調整',
         '「○○の観点をより具体的に」と追加指示して仕上げる'),
        ('4', 'コピー&ペースト',
         'Wordに貼り付けて完成。所要時間：約10分→従来の1/5'),
    ],
    note='ポイント：学習指導要領の文言をそのまま使うと精度が上がる'
)


# ======================================================
# スライド 5 ── 稟議書・公文書
# ======================================================
content_slide(prs,
    '② 稟議書・公文書の作成',
    '定型文書を素早く・高品質に仕上げるプロンプト術',
    [
        ('1', '目的・件名を伝える',
         '「生成AI研修会の実施について、校長宛の稟議書を作って」と入力'),
        ('2', '詳細条件を追加',
         '日時・対象者・費用（無料）・期待効果などを箇条書きで渡す'),
        ('3', '文体の調整',
         '「公用文形式で、漢字を多めに」など体裁を指示して再生成'),
        ('4', '確認・保存',
         '内容を確認してWordに貼り付け。所要時間：約15分→従来の1/4'),
    ],
    note='ポイント：「〇〇形式で」「〇〇向けに」など制約を加えると品質UP'
)


# ======================================================
# スライド 6 ── 授業案
# ======================================================
content_slide(prs,
    '③ 授業案の作成',
    '学習指導要領に沿った授業展開案を短時間で',
    [
        ('1', '単元・学年・時数を伝える',
         '「農業科・○○単元・高校2年・50分×3時間の授業案を作って」'),
        ('2', '展開案が自動生成',
         '導入→展開→まとめの3段構成で指導内容・発問・評価が出力される'),
        ('3', '実態に合わせて修正',
         '「○○農業高校の生徒向けに農場実習の例を入れて」と追加指示'),
        ('4', '完成・印刷',
         '学習指導案の形式に整えてWordへ貼り付け。約20分→従来の1/3'),
    ],
    note='ポイント：指導要録の観点（知・技・思・主）を明示すると精度が上がる'
)


# ======================================================
# スライド 7 ── ワークシート
# ======================================================
content_slide(prs,
    '④ ワークシートの作成',
    '生徒の実態に合わせたワークシートをAIで',
    [
        ('1', '単元・目標・形式を指定',
         '「○○単元、理解確認用、穴埋め＋記述の混合、A4・1枚」'),
        ('2', 'シートが生成される',
         '設問・解答欄・キーワード枠が自動配置されたワークシートが完成'),
        ('3', '難易度・量を調整',
         '「難しすぎるので設問を3問減らして、ヒントを追加して」'),
        ('4', 'コピー→印刷',
         'Word/Googleドキュメントに貼り付けて印刷。約30分→約10分'),
    ],
    note='ポイント：「○○が苦手な生徒向けに」など生徒像を伝えるとベスト'
)


# ======================================================
# スライド 8 ── まとめ・持ち帰り
# ======================================================
slide8 = prs.slides.add_slide(BLANK)
slide_bg(slide8)
add_rect(slide8, 0, 0, W, Inches(0.12), fill_color=COLOR_KEY)
add_rect(slide8, 0, H - Inches(0.12), W, Inches(0.12), fill_color=COLOR_KEY)
add_rect(slide8, 0, Inches(0.12), W, Inches(1.1), fill_color=COLOR_ACCENT)
add_textbox(slide8, Inches(0.5), Inches(0.2), Inches(12), Inches(0.9),
            '今日から使える！まとめ', font_size=28, bold=True, color=COLOR_WHITE)

takeaways = [
    ('ルーブリック', '10分で4観点・5段階の評価表'),
    ('稟議書・公文書', '15分でそのまま使える公用文'),
    ('授業案', '20分で50分授業の展開案'),
    ('ワークシート', '10分で穴埋め＋記述の混合シート'),
]

for i, (label, value) in enumerate(takeaways):
    col = i % 2
    row = i // 2
    x = Inches(0.5) + col * Inches(6.4)
    y = Inches(1.5) + row * Inches(2.1)
    add_rect(slide8, x, y, Inches(6.0), Inches(1.8), fill_color=COLOR_ACCENT)
    add_rect(slide8, x, y, Inches(6.0), Inches(0.55), fill_color=COLOR_KEY)
    add_textbox(slide8, x, y, Inches(6.0), Inches(0.55),
                label, font_size=18, bold=True, align=PP_ALIGN.CENTER, color=COLOR_WHITE)
    add_textbox(slide8, x, y + Inches(0.6), Inches(6.0), Inches(1.0),
                value, font_size=20, bold=True, align=PP_ALIGN.CENTER, color=COLOR_YELLOW)

# 合計節約時間
add_rect(slide8, Inches(0.4), Inches(5.8), Inches(12.5), Inches(0.85),
         fill_color=COLOR_KEY)
add_textbox(slide8, Inches(0.4), Inches(5.82), Inches(12.5), Inches(0.75),
            '4つの業務で合計 約1時間以上の節約　→　毎日使えば年間200時間以上の削減！',
            font_size=20, bold=True, align=PP_ALIGN.CENTER, color=COLOR_WHITE)


# ======================================================
# スライド 9 ── おわり
# ======================================================
slide9 = prs.slides.add_slide(BLANK)
slide_bg(slide9)
add_rect(slide9, 0, 0, W, Inches(0.12), fill_color=COLOR_KEY)
add_rect(slide9, 0, H - Inches(0.12), W, Inches(0.12), fill_color=COLOR_KEY)
add_rect(slide9, 0, Inches(2.8), W, Inches(2.0), fill_color=COLOR_ACCENT)

add_textbox(slide9, Inches(1), Inches(2.85), Inches(11.33), Inches(1.0),
            'ご参加ありがとうございました！',
            font_size=34, bold=True, align=PP_ALIGN.CENTER, color=COLOR_WHITE)
add_textbox(slide9, Inches(1), Inches(3.8), Inches(11.33), Inches(0.6),
            'まずは今日、1つだけ試してみましょう 🚀',
            font_size=20, align=PP_ALIGN.CENTER, color=COLOR_LIGHT)

add_rect(slide9, Inches(3.5), Inches(4.6), Inches(6.33), Inches(0.04), fill_color=COLOR_KEY)

add_textbox(slide9, Inches(1), Inches(4.8), Inches(11.33), Inches(0.5),
            '質問・相談はいつでも声をかけてください',
            font_size=18, align=PP_ALIGN.CENTER, color=COLOR_LIGHT)
add_textbox(slide9, Inches(1), Inches(5.3), Inches(11.33), Inches(0.5),
            '○○農業高等学校　　令和7年3月9日',
            font_size=15, align=PP_ALIGN.CENTER, color=RGBColor(0x88, 0x88, 0xAA))

# ======================================================
# 保存
# ======================================================
output_path = '/home/user/normal-task/生成AI研修_スライド.pptx'
prs.save(output_path)
print(f'保存完了: {output_path}')
print(f'スライド数: {len(prs.slides)} 枚')
