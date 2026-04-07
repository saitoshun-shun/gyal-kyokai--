"""
NotebookLM用PDFを生成するスクリプト
日本語フォントはIPAexGothicを使用（なければNoto Sans JP相当を探す）
"""
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    HRFlowable, KeepTogether
)
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import os, glob

# --- 日本語フォントを探す ---
def find_font():
    candidates = [
        "/usr/share/fonts/opentype/ipafont-gothic/ipagp.ttf",
        "/usr/share/fonts/opentype/ipafont-gothic/ipag.ttf",
        "/usr/share/fonts/truetype/fonts-japanese-gothic.ttf",
        "/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc",
        "/usr/share/fonts/noto-cjk/NotoSansCJK-Regular.ttc",
    ]
    for c in candidates:
        if os.path.exists(c):
            return c
    # globで探す
    for pattern in ["/usr/share/fonts/**/*.ttf", "/usr/share/fonts/**/*.otf", "/usr/share/fonts/**/*.ttc"]:
        found = glob.glob(pattern, recursive=True)
        jp = [f for f in found if any(k in f.lower() for k in ["ipa","noto","gothic","cjk","japanese"])]
        if jp:
            return jp[0]
    return None

font_path = find_font()
if font_path:
    print(f"Using font: {font_path}")
    pdfmetrics.registerFont(TTFont("JP", font_path))
    FONT = "JP"
else:
    print("Warning: No Japanese font found, using Helvetica (text may not render)")
    FONT = "Helvetica"

# --- カラー ---
PINK   = colors.HexColor("#FF69B4")
LPINK  = colors.HexColor("#FFD6E8")
ACCENT = colors.HexColor("#FF1493")
DARK   = colors.HexColor("#2D2D2D")
GRAY   = colors.HexColor("#888888")
WHITE  = colors.white

W, H = A4

# --- スタイル ---
def s(name, **kw):
    kw.setdefault("fontName", FONT)
    return ParagraphStyle(name, **kw)

ST = {
    "h1":   s("h1",   fontSize=20, textColor=WHITE,  backColor=PINK,
               spaceAfter=4, spaceBefore=12, leading=28,
               leftIndent=6, rightIndent=6, borderPad=6),
    "h2":   s("h2",   fontSize=14, textColor=ACCENT, spaceAfter=4,
               spaceBefore=10, leading=20, fontName=FONT),
    "body": s("body", fontSize=10, textColor=DARK,   spaceAfter=3,
               spaceBefore=2, leading=16),
    "bullet":s("bullet",fontSize=10,textColor=DARK,  spaceAfter=2,
               spaceBefore=1, leading=15, leftIndent=14, bulletIndent=4),
    "note": s("note", fontSize=9,  textColor=GRAY,   spaceAfter=2,
               leading=14),
    "title":s("title",fontSize=28, textColor=PINK,   spaceAfter=6,
               spaceBefore=20, leading=36, alignment=1),
    "sub":  s("sub",  fontSize=16, textColor=DARK,   spaceAfter=4,
               leading=22, alignment=1),
    "label":s("label",fontSize=9,  textColor=WHITE,  backColor=PINK,
               leading=13),
}

def h1(text):
    return [Paragraph(text, ST["h1"]), Spacer(1, 2*mm)]

def h2(text):
    return [Paragraph(text, ST["h2"]),
            HRFlowable(width="100%", thickness=1, color=PINK, spaceAfter=2)]

def body(text):
    return Paragraph(text, ST["body"])

def bullet(text):
    return Paragraph(f"▸　{text}", ST["bullet"])

def note(text):
    return Paragraph(text, ST["note"])

def sp(n=4):
    return Spacer(1, n*mm)

# --- テーブルスタイル共通 ---
def tbl_style(header_rows=1):
    return TableStyle([
        ("BACKGROUND",  (0,0), (-1, header_rows-1), PINK),
        ("TEXTCOLOR",   (0,0), (-1, header_rows-1), WHITE),
        ("FONTNAME",    (0,0), (-1,-1), FONT),
        ("FONTSIZE",    (0,0), (-1, header_rows-1), 9),
        ("FONTSIZE",    (0, header_rows), (-1,-1), 9),
        ("ROWBACKGROUNDS",(0, header_rows),(-1,-1),[LPINK, WHITE]),
        ("GRID",        (0,0), (-1,-1), 0.5, PINK),
        ("VALIGN",      (0,0), (-1,-1), "MIDDLE"),
        ("TOPPADDING",  (0,0), (-1,-1), 4),
        ("BOTTOMPADDING",(0,0),(-1,-1), 4),
        ("LEFTPADDING", (0,0), (-1,-1), 6),
    ])

# ============================================================ コンテンツ構築
story = []

# ── 表紙ページ相当 ──────────────────────────────────────────
story.append(sp(20))
story.append(Paragraph("ギャル × インバウンド", ST["title"]))
story.append(Paragraph("コンテンツビジネス事業計画書", ST["sub"]))
story.append(sp(4))
story.append(Paragraph("ギャル協会　×　サイバーバズ　×　ENTIAL", ST["sub"]))
story.append(sp(2))
story.append(Paragraph("2026年4月　／　NotebookLM用資料", ST["note"]))
story.append(HRFlowable(width="100%", thickness=2, color=PINK, spaceAfter=6))
story.append(sp(8))

# ── 1. ビジョン・事業概要 ──────────────────────────────────
story += h1("1. ビジョン・事業概要")
story.append(body("「ギャル文化」を日本最強のインバウンドコンテンツIPとして世界に届ける"))
story.append(sp(2))
story += h2("事業概要")
story.append(body("訪日外国人が求める「本物の日本サブカルチャー体験」を、ギャル協会／うさたにパイセンのIPをコアに、以下4軸で収益化するビジネス。"))
for t in ["ギャル協会 / うさたにパイセンのIPをコアに",
          "TikTokショートドラマメディア（サイバーバズ名義）で拡散",
          "リアル体験コンテンツ（渋谷ツアー・ギャル旅）で差別化",
          "ENTIALが企業・自治体へタイアップ営業"]:
    story.append(bullet(t))
story.append(sp(4))

# ── 2. 市場背景 ───────────────────────────────────────────
story += h1("2. 市場背景　〜なぜ今ギャル×インバウンドか〜")
mkt = [
    ["#", "背景"],
    ["01", "訪日外国人数が過去最高を更新。インバウンド消費が拡大継続中"],
    ["02", "海外での「ギャル文化」認知がSNSを中心に急上昇（Y2Kブーム・日本サブカル人気）"],
    ["03", "観光庁・自治体がインバウンド向けコンテンツに積極的に予算を投下中"],
    ["04", "ギャル × インバウンドを商品化している競合は現時点でゼロ"],
]
t = Table(mkt, colWidths=[15*mm, 145*mm])
t.setStyle(tbl_style())
story.append(t)
story.append(sp(4))

# ── 3. 3社の役割 ──────────────────────────────────────────
story += h1("3. 3社の役割")
roles = [
    ["プレイヤー", "役割", "主なアウトプット"],
    ["ギャル協会 / うさたにパイセン", "IPオーナー・クリエイター",
     "世界観・企画・出演・ギャル文化の発信源"],
    ["サイバーバズ", "メディアオーナー・実行",
     "TikTokアカウント運営・広告販売・拡散・制作費負担"],
    ["ENTIAL", "営業部隊・プロジェクト管理",
     "企業・自治体への営業、営業代理店の開拓・管理、進行管理"],
]
t = Table(roles, colWidths=[45*mm, 50*mm, 65*mm])
t.setStyle(tbl_style())
story.append(t)
story.append(sp(4))

# ── 4. 収益モデル ─────────────────────────────────────────
story += h1("4. 収益の流れ・お金の設計")
story += h2("お金の流れ")
for line in [
    "企業・自治体（広告主）がタイアップ費・協賛費・ライセンス料を支払う",
    "ENTIALが受注・請求窓口として受け取り、サイバーバズへ入金",
    "サイバーバズがギャル協会へ出演費＋レベニューシェアを支払う",
    "ENTIALはサイバーバズから営業手数料（固定費＋売上連動インセンティブ）を受け取る",
]:
    story.append(bullet(line))
story.append(sp(2))
money = [
    ["項目", "内容"],
    ["コンテンツ制作費・TikTok運営費", "サイバーバズが全額負担"],
    ["ギャル協会の財務リスク", "ゼロ（先行投資なし）"],
    ["レベニューシェア比率", "サイバーバズ：ギャル協会 ＝ 6:4〜7:3（要交渉）"],
    ["ENTIALの報酬", "固定費 ＋ 売上連動インセンティブ"],
]
t = Table(money, colWidths=[70*mm, 90*mm])
t.setStyle(tbl_style())
story.append(t)
story.append(sp(4))

# ── 5. 広告商品パッケージ ────────────────────────────────
story += h1("5. 広告商品パッケージ（インバウンド特化）")

pkgs = [
    ("Package 1", "ギャル旅 × 地方観光タイアップ",
     "自治体・観光協会",
     "300〜700万/地域",
     "ギャル協会メンバーが地方を訪問しショートドラマを制作・TikTok配信。英語字幕付きで海外SNSに拡散。観光庁・文化庁補助金の活用で自治体負担を圧縮できる可能性あり。"),
    ("Package 2", "渋谷ギャル文化体験ツアー × メディア化",
     "ホテル・旅行代理店・OTA",
     "50〜150万",
     "うさたにパイセン監修の体験ツアーを商品化し、ツアー密着をコンテンツ化。ホテルのアクティビティメニューやOTAのユニーク体験枠に掲載。"),
    ("Package 3", "インバウンド向けショートドラマ広告",
     "免税店・コスメ・決済サービス等",
     "150〜400万",
     "外国人観光客が登場するギャル×インバウンドドラマに企業商品・サービスを自然な形で組み込む。"),
    ("Package 4", "「ギャル協会公認」コンテンツライセンス",
     "海外メディア・観光アプリ",
     "月額20〜80万",
     "ギャル協会公認コンテンツの使用権・監修権を海外向けメディアやアプリに販売。"),
]

for pkg, name, target, price, detail in pkgs:
    tbl = Table(
        [["ターゲット", "参考単価"], [target, price]],
        colWidths=[90*mm, 70*mm]
    )
    tbl.setStyle(tbl_style())
    story.append(KeepTogether([
        *h2(f"{pkg}｜{name}"),
        tbl,
        sp(1),
        body(detail),
        sp(3),
    ]))

story.append(sp(4))

# ── 6. 営業戦略 ───────────────────────────────────────────
story += h1("6. 営業戦略（ENTIALの動き方）")
story += h2("ターゲット優先順位")
tgt = [
    ["優先度", "ターゲット", "理由"],
    ["★★★", "ホテル・旅行代理店", "意思決定が速く予算あり。インバウンド需要でピーク"],
    ["★★★", "免税店・百貨店", "インバウンド消費に直結。高単価・継続案件になりやすい"],
    ["★★",  "地方自治体・観光協会", "予算大だが意思決定遅い。補助金活用でハードル下げられる"],
    ["★",   "OTA・体験予約サービス", "拡散力大。TikTokが育ってから本格攻略"],
]
t = Table(tgt, colWidths=[18*mm, 52*mm, 90*mm])
t.setStyle(tbl_style())
story.append(t)
story.append(sp(3))
story += h2("ENTIALの動き方（ステップ）")
for step in [
    "①　営業代理店を1〜2社開拓（インバウンド・観光系に強い代理店）",
    "②　代理店経由でホテル・旅行代理店・百貨店へ提案",
    "③　自治体案件は補助金スキームとセットで直接提案",
]:
    story.append(bullet(step))
story.append(sp(4))

# ── 7. ロードマップ ───────────────────────────────────────
story += h1("7. ロードマップ")
rm = [
    ["フェーズ", "期間", "主なアクション"],
    ["Phase 1\n仕込み期", "0〜3ヶ月",
     "・3社間の契約・役割・収益分配を合意\n・TikTokアカウント開設・初期コンテンツ10本制作\n・ENTIALが営業代理店1〜2社を開拓\n・ホテル・旅行代理店への初回提案開始"],
    ["Phase 2\n初収益期", "3〜6ヶ月",
     "・タイアップ案件 初回受注\n・TikTokフォロワー1万人達成\n・自治体案件 1件提案・交渉開始"],
    ["Phase 3\nスケール期", "6〜12ヶ月",
     "・月次タイアップ案件を安定受注\n・TikTok英語字幕で海外フォロワー獲得\n・ギャルIPの海外ライセンス販売開始\n・神7組成・複数案件を並走できる体制構築"],
]
t = Table(rm, colWidths=[25*mm, 22*mm, 113*mm])
t.setStyle(tbl_style())
story.append(t)
story.append(sp(4))

# ── 8. サイバーバズ打ち合わせ確認事項 ────────────────────
story += h1("8. サイバーバズ打ち合わせ　確認事項")
items = [
    ("①", "TikTokアカウントの名義・権利関係", "バズ名義で進める方向で合意確認"),
    ("②", "コンテンツ制作費の上限設定", "Phase 1の先行投資額を決める"),
    ("③", "レベニューシェア比率", "ギャル協会との分配。6:4をたたき台に交渉"),
    ("④", "ENTIALへの営業手数料スキーム", "固定費＋売上連動インセンティブの設計"),
    ("⑤", "KPI設定", "TikTokフォロワー数・月次受注額・1年の売上目標"),
]
conf = [["番号", "確認項目", "内容"]] + [[n,t,d] for n,t,d in items]
t = Table(conf, colWidths=[12*mm, 65*mm, 83*mm])
t.setStyle(tbl_style())
story.append(t)
story.append(sp(6))
story.append(HRFlowable(width="100%", thickness=1, color=PINK))
story.append(note("本資料はNotebookLM用のソース資料です。ギャル協会×サイバーバズ×ENTIAL 打ち合わせ用たたき台（2026年4月）"))

# ============================================================ 出力
out = "/home/user/gyal-kyokai--/事業計画書_ギャル×インバウンド_NotebookLM用.pdf"
doc = SimpleDocTemplate(
    out, pagesize=A4,
    rightMargin=18*mm, leftMargin=18*mm,
    topMargin=15*mm, bottomMargin=15*mm,
    title="ギャル×インバウンド事業計画書",
    author="ギャル協会×サイバーバズ×ENTIAL",
)
doc.build(story)
print(f"Saved: {out}")
