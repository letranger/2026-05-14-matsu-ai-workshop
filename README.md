# 馬祖 × AI：從藍眼淚寫一篇科學小論文

> 教育部中小學推廣教育計畫 AI 教學與實作推廣研習教材
> 2026-05-14（四）17:30–20:30｜國立馬祖高級中學

## 場次資訊

| 項目  | 內容 |
|-------|------|
| 主題  | 馬祖 × AI——從藍眼淚寫一篇科學小論文 |
| 時間  | 2026-05-14（四）17:30–20:30（共 3 小時） |
| 地點  | 國立馬祖高級中學（連江縣南竿鄉 374 號） |
| 對象  | 連江縣高中職學校教師（各科） |
| 主辦  | 國立臺南大學理工學院 AI 教育推廣暨研究發展中心 |
| 授課  | 顏永進（臺南一中） |

## 課程主軸

研習圍繞兩條主線：**會用** 與 **會看懂**。

### 一、AI 責任轉移四階段模型（L1–L4）

| 階段 | 名稱    | 一秒辨識點                  | 今晚位置                  |
|------|---------|-----------------------------|---------------------------|
| L1   | 對話 AI | 你打字、AI 回字              | 開場 demo 三邊對照        |
| L2   | 協作 AI | 你還在點滑鼠，但 app 裡有 AI | **實作一全程**            |
| L3   | 代理 AI | AI 在動你的鍵盤／滑鼠／檔案  | 實作二 Section 2          |
| L4   | 常駐 AI | 你不在電腦前，AI 還在背景跑  | 實作二 Section 1 & 4      |

核心判準：**誰在動鍵盤**。

### 二、實作一：藍眼淚科學小論文（五步流程）

| 步驟  | 工具                | 在做什麼                            | 產出                |
|-------|---------------------|-------------------------------------|---------------------|
| 1 讀  | NotebookLM          | 馬祖在地素材 → 研究筆記             | `research_notes.docx` |
| 2 畫  | Napkin × 2          | 因果關係圖＋時間軸                  | 兩張 PNG            |
| 3 寫  | Claude              | 用筆記＋圖 → 小論文初稿             | `paper_draft.docx`  |
| 4 變  | Claude              | 同一份內容換骨架（教案／心得／報告）| `report_variant.docx` |
| 5 報  | Claude vs Gamma     | 同一份做投影片，兩工具對比          | `slides.html` + `.pptx` |

### 三、實作二：模型、代理人與你

- **Section 1**：Skill（Lv.4）——把今晚流程包成 SOP
- **Section 2**：Gemini CLI（Lv.3）——把 AI 從瀏覽器拉出來
- **Section 3**：風險與隱私——AI 代理失控與資料保護
- **Section 4**：L4 demo——Claude Code、OpenClaw

### 四、原理段：為什麼 AI 能生成文字、影音

- 關於模型你應該知道的事（商業 vs 開源）
- AI 到底怎麼辦到的？——函數視角
- 為什麼會生出幻覺
- RAG 如何解決幻覺問題

## 資料夾結構

```
.
├── README.md                                # 本檔
├── CLAUDE.md                                # Claude Code 工作說明
├── index.org                                # 研習主控頁（Org-mode 原始檔）
├── index.html                               # 主控頁（HTML 匯出）
├── 素材包.md                                # NotebookLM 素材包指引
├── 小論文範本.docx                          # Step 3 範本（全國高中小論文比賽格式）
├── 小論文格式說明.pdf                       # 比賽官方規範
├── 教案範本-橋頭糖廠版.docx                 # Step 4 教案格式範本
├── 換骨架-prompts.md                        # Step 4 四個變體 prompt
├── RAG-flow-diagram.md                      # 原理段 RAG 流程圖
├── 馬祖小論文.skill/                        # Claude Skill 包
├── examples/                                # 範例輸出
│   ├── 研究筆記-範例.md                     # NotebookLM 預期產出
│   └── 小論文-範例.md                       # Claude 預期產出
├── scripts/                                 # 演示用腳本
│   ├── generate_exam_pdfs.py                # 生成 30 份學生考卷 PDF
│   └── generate_ground_truth_xlsx.py        # 生成正解 Excel
├── 學生考卷/                                # 29 份 demo 學生 PDF
├── 學生考卷-正解.xlsx                       # Gemini CLI demo 對照組
├── images/, img/                            # 教學圖（函數、機率示意）
└── 藍眼淚成因因果關係圖.png, 時間軸.png     # Napkin 預生成範例
```

## 學員會前準備

1. **註冊免費帳號**：
   - [NotebookLM](https://notebooklm.google.com)（Google 帳號即可）
   - [Napkin](https://www.napkin.ai)
   - [Claude](https://claude.ai)（建議用 Google 登入）
   - [Gamma](https://gamma.app)

2. **下載這份 repo**（Code → Download ZIP 或 `git clone`）

3. **筆電帶來**：能上網、能裝 Chrome／Edge

## 工具與成本

| 工具       | 免費版能做的                          | 進階建議               |
|------------|---------------------------------------|------------------------|
| NotebookLM | 50 個 source、5 個 notebook、Audio 摘要 | 對 80% 老師夠用        |
| Napkin     | 無限張視覺圖（有浮水印）              | 對 80% 老師夠用        |
| Claude     | 每 5 小時數則 Sonnet＋無限 Haiku      | 重度用戶推薦 Pro($20) |
| Gemini CLI | 1500 req/day                          | 對 80% 老師夠用        |
| Gamma      | 400 點數一輩子                        | 教師輕度使用夠         |

## 衍生作品授權

本 repo 教材以 [CC BY-NC-SA 4.0](https://creativecommons.org/licenses/by-nc-sa/4.0/deed.zh-Hant) 授權：
- 教師研習用途 **可自由改寫使用**
- 引用請註明作者（顏永進）與本 repo 連結
- 不可作商業用途
- 衍生作品需採相同授權

## 相關場次

本研習延續自 2026 春季三場連續教師研習，同一套 NotebookLM→Napkin→Claude 流程不同素材：

- 2026-04-30 雲林：和 AI 作朋友（高中生版，主題：虎尾糖廠）
- 2026-05-07 高雄市輔導團：AI 與閱讀力（教師版，主題：橋頭糖廠）
- 2026-05-14 馬祖：從藍眼淚寫小論文（教師版，本場）

## 聯絡

- 主辦單位：aik12.edu@gmail.com / 06-2606123 #7027
- 教師個人：letranger@gm.tnfsh.tn.edu.tw

---

📚 配套教材：[《和 AI 做朋友——相知篇：揭秘生成式 AI》](https://market.cloud.edu.tw/resources/web/1841948)（教育部出版／教學資源市集）
