"""
生成 30 份學生考卷 PDF——用於 Gemini CLI demo「掃資料夾摘要成 Excel」。

題目：海洋生態 × 馬祖藍眼淚 10 題單選，每題 10 分。
產出：學生考卷/座號-姓名.pdf × 30 份
"""
from __future__ import annotations
import os
import random
import subprocess
import tempfile
from pathlib import Path

# ---------- 配置 ----------
BASE = Path(__file__).resolve().parent.parent
OUT_DIR = BASE / "學生考卷"
CHROMIUM = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"
SCHOOL = "國立馬祖高級中學"
CLASS_NAME = "高一甲班"
SUBJECT = "海洋生態 × 馬祖藍眼淚 隨堂測驗"
EXAM_DATE = "115 年 5 月 14 日"

# ---------- 題目與答案 ----------
QUESTIONS = [
    {
        "q": "馬祖藍眼淚的主要發光生物是？",
        "options": ["螢火蟲", "夜光蟲", "水母", "珊瑚"],
        "ans": 1,  # B
    },
    {
        "q": "夜光蟲在生物分類上屬於？",
        "options": ["脊椎動物", "節肢動物", "陸生植物", "單細胞甲藻"],
        "ans": 3,  # D
    },
    {
        "q": "馬祖藍眼淚最佳觀賞月份是？",
        "options": ["1–2 月", "4–6 月", "7–9 月", "10–12 月"],
        "ans": 1,  # B
    },
    {
        "q": "夜光蟲體內負責發光的化學系統是？",
        "options": ["螢光素—螢光素酶", "葉綠素 a", "黑色素", "維生素 D"],
        "ans": 0,  # A
    },
    {
        "q": "下列何者最有利於藍眼淚被觀賞到？",
        "options": ["漲潮時的海水擾動", "滿月強光照射", "海岸強烈光害", "冬季低水溫"],
        "ans": 0,  # A
    },
    {
        "q": "蔣國平團隊 2021 年揭密藍眼淚機制的論文發表於？",
        "options": ["Nature", "Science", "Frontiers in Marine Science", "Cell"],
        "ans": 2,  # C
    },
    {
        "q": "2021 年研究指出，夜光蟲爆發與消失的關鍵是？",
        "options": ["餌料不足觸發有性生殖", "月相週期", "颱風擾動", "海水酸鹼值劇變"],
        "ans": 0,  # A
    },
    {
        "q": "為什麼馬祖藍眼淚難以準確預測出現時間？",
        "options": ["受多重環境因素交互影響", "完全隨機事件", "僅在週末出現", "政府刻意控管"],
        "ans": 0,  # A
    },
    {
        "q": "「暗空公園」概念的核心目標是？",
        "options": ["增設路燈", "減少光害以利觀星與藍眼淚", "限制夜間活動", "種植螢光植物"],
        "ans": 1,  # B
    },
    {
        "q": "下列哪個觀點代表「自然現象論」立場？",
        "options": [
            "藍眼淚是海洋污染的指標",
            "藍眼淚是自然食物鏈循環的產物",
            "藍眼淚對人類有害應禁止觀賞",
            "藍眼淚是外來種入侵",
        ],
        "ans": 1,  # B
    },
]
LETTER = "ABCD"

# ---------- 30 位虛構學生 ----------
SURNAMES = list("王李張陳林黃吳劉楊蔡周葉鄭蘇呂許賴何宋鄧邱徐丁潘簡藍曹尤魏馬")
GIVEN = [
    "佳穎", "宥廷", "詠晴", "庭瑄", "承恩", "羽彤", "宜蓁", "柏翰", "子涵", "宇恩",
    "雅雯", "冠霖", "采蓁", "睿哲", "予晴", "之翔", "亭儀", "凱翔", "茉莉", "詠新",
    "辰穎", "怡蓁", "皓宇", "芷晴", "建宏", "若菲", "亦凡", "曉彤", "崇翰", "品妍",
]
SKIP_SEATS = {23}  # 已退選／轉學的座號（保留缺號，不重編）

STUDENTS = [
    {"seat": i + 1, "name": SURNAMES[i] + GIVEN[i]}
    for i in range(30)
    if (i + 1) not in SKIP_SEATS
]

# ---------- HTML 模板 ----------
HTML_TPL = """<!DOCTYPE html>
<html lang="zh-Hant">
<head>
<meta charset="utf-8">
<title>{school}-{subject}-{name}</title>
<style>
  @page {{ size: A4; margin: 2cm; }}
  body {{
    font-family: "PingFang TC", "Heiti TC", "STHeiti", "Microsoft JhengHei", sans-serif;
    font-size: 12pt; color: #222; line-height: 1.5;
  }}
  .header {{ border-bottom: 2px solid #333; padding-bottom: 10px; margin-bottom: 18px; }}
  .header h1 {{ font-size: 16pt; margin: 0 0 6px; text-align: center; }}
  .meta {{
    display: grid; grid-template-columns: repeat(3, 1fr); gap: 4px 12px;
    font-size: 11pt; margin-top: 8px;
  }}
  .meta .score {{
    grid-column: 3; text-align: right; font-weight: bold; font-size: 14pt; color: #c00;
  }}
  .question {{
    margin: 10px 0; padding: 8px 10px; border-left: 3px solid #ddd;
    page-break-inside: avoid;
  }}
  .q-num {{ font-weight: bold; }}
  .options {{ margin: 4px 0 0 18px; }}
  .option {{ margin: 2px 0; }}
  .picked {{ color: #06c; font-weight: bold; }}
  .correct-mark {{ color: #0a0; font-weight: bold; }}
  .wrong-mark   {{ color: #c00; font-weight: bold; }}
  .answer-row {{
    margin-top: 4px; padding-left: 18px; font-size: 10.5pt; color: #555;
  }}
  .footer-grade {{
    margin-top: 16px; padding-top: 10px; border-top: 1px dashed #888;
    font-size: 11pt; display: grid; grid-template-columns: 1fr 1fr; gap: 4px;
  }}
</style>
</head>
<body>
  <div class="header">
    <h1>{school} {class_name}</h1>
    <h1 style="font-size:14pt; font-weight:normal;">{subject}</h1>
    <div class="meta">
      <div>座號：{seat:02d}</div>
      <div>姓名：{name}</div>
      <div class="score">總分：{score}</div>
      <div>日期：{date}</div>
      <div>題數：10 題，每題 10 分</div>
      <div>答對：{correct}／10</div>
    </div>
  </div>
  {questions_html}
  <div class="footer-grade">
    <div>批改老師：顏永進</div>
    <div style="text-align:right;">家長簽名：________________</div>
  </div>
</body>
</html>
"""


def render_question(idx: int, q: dict, picked: int) -> str:
    """渲染單題 HTML，含學生答案與對錯標記。"""
    opts_html = []
    for i, opt in enumerate(q["options"]):
        cls = []
        suffix = ""
        if i == picked:
            cls.append("picked")
            if picked == q["ans"]:
                suffix = " <span class='correct-mark'>✓</span>"
            else:
                suffix = " <span class='wrong-mark'>✗</span>"
        cls_attr = f" class='option {' '.join(cls)}'" if cls else " class='option'"
        opts_html.append(f"<div{cls_attr}>({LETTER[i]}) {opt}{suffix}</div>")

    correct_note = (
        ""
        if picked == q["ans"]
        else f"<div class='answer-row'>正確答案：（{LETTER[q['ans']]}）</div>"
    )
    return f"""
  <div class="question">
    <div class="q-num">{idx + 1}. {q['q']}</div>
    <div class="options">{''.join(opts_html)}</div>
    {correct_note}
  </div>"""


def simulate_answers(seed: int) -> tuple[list[int], int, int]:
    """模擬一位學生的作答——隨機但偏向常態分佈，有人高分有人低分。"""
    rng = random.Random(seed)
    # 目標答對率：在 [0.3, 1.0] 取 beta-like 分佈
    target_correct = int(round(rng.triangular(3, 10, 7)))
    target_correct = max(2, min(10, target_correct))

    picks = []
    correct_indices = list(range(10))
    wrong_indices = list(range(10))
    rng.shuffle(correct_indices)
    rng.shuffle(wrong_indices)
    will_correct = set(correct_indices[:target_correct])

    for i, q in enumerate(QUESTIONS):
        if i in will_correct:
            picks.append(q["ans"])
        else:
            wrong = [j for j in range(4) if j != q["ans"]]
            picks.append(rng.choice(wrong))

    correct_count = sum(1 for i, p in enumerate(picks) if p == QUESTIONS[i]["ans"])
    score = correct_count * 10
    return picks, correct_count, score


def html_to_pdf(html: str, pdf_path: Path) -> None:
    with tempfile.NamedTemporaryFile(
        mode="w", suffix=".html", delete=False, encoding="utf-8"
    ) as f:
        f.write(html)
        html_path = f.name
    try:
        subprocess.run(
            [
                CHROMIUM,
                "--headless",
                "--disable-gpu",
                "--no-sandbox",
                "--no-pdf-header-footer",
                f"--print-to-pdf={pdf_path}",
                f"file://{html_path}",
            ],
            check=True,
            capture_output=True,
        )
    finally:
        os.unlink(html_path)


def main() -> None:
    OUT_DIR.mkdir(parents=True, exist_ok=True)

    for s in STUDENTS:
        picks, correct, score = simulate_answers(seed=s["seat"] * 17 + 31)
        questions_html = "".join(
            render_question(i, q, picks[i]) for i, q in enumerate(QUESTIONS)
        )
        html = HTML_TPL.format(
            school=SCHOOL,
            class_name=CLASS_NAME,
            subject=SUBJECT,
            date=EXAM_DATE,
            seat=s["seat"],
            name=s["name"],
            score=score,
            correct=correct,
            questions_html=questions_html,
        )
        pdf_name = f"{s['seat']:02d}-{s['name']}.pdf"
        html_to_pdf(html, OUT_DIR / pdf_name)
        print(f"  ✓ {pdf_name}  總分 {score:3d}  答對 {correct}/10")

    print(f"\n共產出 {len(STUDENTS)} 份 PDF：{OUT_DIR}")


if __name__ == "__main__":
    main()
