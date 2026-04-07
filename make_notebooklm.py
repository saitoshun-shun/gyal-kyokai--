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
    ["#", "背景", "データ・根拠"],
    ["01", "インバウンド消費が過去最高を更新",
     "2024年の訪日外国人消費額は8.1兆円（前年比+53.4%）、訪日者数3,687万人。いずれも過去最高（観光庁、2025年1月）"],
    ["02", "海外でのギャル文化認知が急上昇",
     "#gyaru #gyarufashionがTikTok・Instagramで数百万再生。米国・欧州・東南アジアでコミュニティ形成。Hyper Japan（英）・Anime Expo（米）でも展示"],
    ["03", "Y2Kブームが追い風",
     "#y2kのTikTok投稿数460万件超。ギャルはY2Kファッションの象徴として2025〜2026年にかけて再注目"],
    ["04", "サブカル観光の成功実績あり",
     "アニメ聖地巡礼で訪日外国人の8.1%（推計299万人）が参加。メイドカフェはTripAdvisor千代田区1位にランクイン"],
    ["05", "ギャル × インバウンド商品化の競合はゼロ",
     "メイドカフェ・忍者体験・アニメ観光と異なり、ギャル文化をインバウンド向けに体系化した事業者は現時点で存在しない"],
]
t = Table(mkt, colWidths=[12*mm, 55*mm, 93*mm])
t.setStyle(tbl_style())
story.append(t)
story.append(sp(4))

# ── 2.5 ターゲット顧客ペルソナ ─────────────────────────────
story += h1("3. ターゲット顧客ペルソナ")
story.append(body("以下の国・地域からの訪日外国人が主要ターゲット。"))
story.append(sp(2))
persona = [
    ["国・地域", "特徴", "ギャルへの興味接点"],
    ["米国・カナダ", "消費単価トップ圏（英国38万円、米国高水準）。Y2K・日本サブカル人気が高い",
     "LA Gyaru Circle等の現地コミュニティあり。TikTok経由でギャルを知った層が多い"],
    ["東南アジア\n（タイ・マレーシア等）", "訪日者数上位、リピーターが多い。日本ポップカルチャーへの親和性が高い",
     "日本のファッション・コスメへの憧れが強く、ギャルメイクの再現動画が人気"],
    ["欧州\n（英・仏・独）", "1人当たり消費額が高い。体験型・文化体験を好む傾向",
     "Hyper Japan等のイベントでギャルファン層が形成済み"],
    ["中国・台湾・香港", "消費額シェアトップ（中国1.7兆円）。Z世代の日本カルチャー熱が高い",
     "ギャルメイク・ファッションのSNS発信への関心が高い"],
]
t = Table(persona, colWidths=[28*mm, 62*mm, 70*mm])
t.setStyle(tbl_style())
story.append(t)
story.append(sp(4))

# ── 3. 3社の役割 ──────────────────────────────────────────
story += h1("4. 3社の役割")
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

# ── 競合相場データ ────────────────────────────────────────
story += h1("5. 競合・相場データ（価格設定の根拠）")
story.append(body("広告商品パッケージの単価設定は以下の市場相場をベースにしています。"))
story.append(sp(2))
comp = [
    ["参考カテゴリ", "相場・実績", "出典・根拠"],
    ["TikTokショートドラマ制作費", "数万ドル〜（標準）。プレミアム品質で$400,000〜$600,000",
     "Business of Apps / Variety 2025年レポート"],
    ["ショートドラマ市場規模", "2024年グローバル市場$14億→2030年$95億予測（CAGR 28.4%）",
     "Market Report Analytics 2025"],
    ["インフルエンサータイアップ", "ブランドがキャラクターの衣装提供・商品をプロット内に組込み直販連動",
     "TikTok Minis広告事例（2025）"],
    ["アニメ聖地巡礼の経済効果", "アニメ放映後10年で約31億円の経済波及効果（鷲宮神社事例）",
     "日本政策投資銀行レポート"],
    ["体験型インバウンド観光", "忍者・サムライ体験が外国人人気観光スポット2位。メイドカフェTripAdvisor千代田区1位",
     "訪日ラボ・JapanTicket 2026年調査"],
]
t = Table(comp, colWidths=[40*mm, 68*mm, 52*mm])
t.setStyle(tbl_style())
story.append(t)
story.append(sp(4))

# ── 5. 収益モデル ─────────────────────────────────────────
story += h1("6. 収益の流れ・お金の設計")
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
story += h1("7. 広告商品パッケージ（インバウンド特化）")

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
story += h1("8. 営業戦略（ENTIALの動き方）")
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
story += h1("9. ロードマップ")
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
story += h1("10. サイバーバズ打ち合わせ　確認事項")
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
