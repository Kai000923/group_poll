# -*- coding: utf-8 -*-
"""
================================================================================
 實驗室 Group Meeting 時段調查（仿 Doodle）— Streamlit 單頁應用程式
================================================================================
功能總覽：
  1. 仿 Doodle 投票介面（姓名 + 身份下拉 + 時段複選）
  2. 防呆覆蓋機制（同一姓名重複投票會覆蓋舊紀錄）
  3. 身份加權演算法（田老師為硬性條件；碩博 2 分、專題生 1 分）
  4. 排除時段黑名單（EXCLUDE_SLOTS）
  5. Seaborn 熱力圖即時看板 + 各時段投票名單
  6. 一鍵將最佳結果推播至 Slack（Incoming Webhook）

資料持久化：寫入 Google Sheets（透過 gspread + Service Account），
            避免 Streamlit Community Cloud 休眠導致記憶體資料遺失。
================================================================================
"""

import html
import json
import os
from datetime import datetime

import gspread
import matplotlib
import matplotlib.font_manager as fm
import matplotlib.patheffects as path_effects
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import requests
import seaborn as sns
import streamlit as st
from google.oauth2.service_account import Credentials
from matplotlib.colors import LinearSegmentedColormap
from matplotlib.patches import FancyBboxPatch

# ------------------------------------------------------------------------------
# 中文字型設定（解決熱力圖中文變成方框 / 豆腐字的問題）
# 策略：1) 嘗試把系統內常見的中文字型檔註冊給 matplotlib
#       2) 從已安裝字型中挑一個支援 CJK 的設為預設
# 部署到 Streamlit Cloud 時，請搭配本專案的 packages.txt（內含 fonts-noto-cjk），
# 雲端就會安裝 Noto Sans CJK，下方即可自動偵測到。
# ------------------------------------------------------------------------------
@st.cache_resource(show_spinner=False)
def setup_chinese_font():
    """回傳一個指向中文字型「檔案」的 FontProperties 物件（找不到回傳 None）。

    為什麼用 FontProperties(fname=...) 而不是只設字型名稱？
      因為雲端有時會有「同名但壞掉的字型快取」，只指定名稱仍可能畫成方框。
      直接綁定字型檔路徑，並把它套用到每一個文字元素上，就能 100% 避免豆腐字。

    偵測順序：
      1. 本專案內 fonts/ 資料夾的字型檔（最可靠，跟著 repo 走，部署到哪都有中文）
      2. 系統內常見的 CJK 字型（Linux / macOS / Windows）
    """
    here = os.path.dirname(os.path.abspath(__file__))
    candidate_files = [
        "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",  # Linux / Streamlit Cloud
        "/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc",
        "/System/Library/Fonts/PingFang.ttc",                       # macOS
        "C:/Windows/Fonts/msjh.ttc",                                # Windows 微軟正黑體
    ]
    # 把 bundled 字型放最前面（優先使用）
    bundled_dir = os.path.join(here, "fonts")
    if os.path.isdir(bundled_dir):
        for fn in sorted(os.listdir(bundled_dir)):
            if fn.lower().endswith((".ttc", ".ttf", ".otf")):
                candidate_files.insert(0, os.path.join(bundled_dir, fn))

    for path in candidate_files:
        if os.path.exists(path):
            try:
                fm.fontManager.addfont(path)
                fp = fm.FontProperties(fname=path)
                # 同步設定 rcParams 當作備援
                matplotlib.rcParams["font.sans-serif"] = (
                    [fp.get_name()] + matplotlib.rcParams["font.sans-serif"]
                )
                matplotlib.rcParams["axes.unicode_minus"] = False
                return fp
            except Exception:
                continue

    matplotlib.rcParams["axes.unicode_minus"] = False
    return None


# 模組載入時即執行一次，之後畫圖都用這個字型物件
_FONT_PROP = setup_chinese_font()

# ==============================================================================
# 【區塊 0】全域設定 — 你平常只需要改這一段
# ==============================================================================

# --- 排除時段黑名單（Blacklist）---------------------------------------------
# 寫進這個清單的時段「不會出現在前台投票選項」中。
# 例如某堂共同必修課的時段，直接貼進來即可。
# 字串必須與下方 ALL_SLOTS 內的時段標籤「完全一致」。
EXCLUDE_SLOTS = [
    "週四 13:00-15:00",   # 範例：共同必修課，排除掉
]

# --- 所有候選時段 -------------------------------------------------------------
# 每個時段用 dict 描述：
#   date  -> 熱力圖 X 軸（日期 / 星期）
#   time  -> 熱力圖 Y 軸（時段）
#   label -> 顯示與儲存用的唯一名稱（= date + " " + time）
def _slot(date, time):
    return {"date": date, "time": time, "label": f"{date} {time}"}

ALL_SLOTS = [
    _slot("下週一", "10:00-12:00"),
    _slot("下週一", "14:00-16:00"),
    _slot("下週二", "10:00-12:00"),
    _slot("下週三", "10:00-12:00"),
    _slot("下週三", "14:00-16:00"),
    _slot("下週四", "13:00-15:00"),   # 注意：會被 EXCLUDE_SLOTS 過濾掉（範例）
    _slot("下週五", "10:00-12:00"),
    _slot("下週五", "15:00-17:00"),
]

# --- 身份選項與權重 -----------------------------------------------------------
ROLE_TEACHER = "田老師"
ROLE_GRAD = "碩/博班"
ROLE_PROJECT = "專題生"
ROLE_OPTIONS = [ROLE_TEACHER, ROLE_GRAD, ROLE_PROJECT]

# 軟性權重：碩博 2 分、專題生 1 分。田老師為硬性條件，不計入加權分（僅判斷出席）。
ROLE_WEIGHT = {
    ROLE_TEACHER: 0,
    ROLE_GRAD: 2,
    ROLE_PROJECT: 1,
}

# Google Sheets 設定（試算表名稱 / 工作表名稱）
GSHEET_NAME = "GroupMeetingPoll"   # 你建立的 Google 試算表檔名
WORKSHEET_NAME = "votes"           # 工作表（分頁）名稱

# 試算表欄位：name(姓名) | role(身份) | slots(勾選時段，以 ; 分隔) | timestamp
HEADER = ["name", "role", "slots", "timestamp"]


# ==============================================================================
# 【區塊 1】Google Sheets 連線與讀寫
# ==============================================================================

# Google API 權限範圍（讀寫試算表 + 存取雲端硬碟檔案）
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]


@st.cache_resource(show_spinner=False)
def get_worksheet():
    """
    建立並快取 gspread 連線，回傳目標工作表物件。
    使用 @st.cache_resource，避免每次 rerun 都重新驗證（連線只建立一次）。
    金鑰從 st.secrets["gcp_service_account"] 讀取（TOML 格式，詳見教學文件）。
    """
    # 從 Streamlit Secrets 取得 Service Account 憑證（為 dict）
    service_account_info = dict(st.secrets["gcp_service_account"])
    creds = Credentials.from_service_account_info(service_account_info, scopes=SCOPES)
    client = gspread.authorize(creds)

    # 開啟試算表；若找不到工作表分頁就自動建立並寫入表頭
    spreadsheet = client.open(GSHEET_NAME)
    try:
        worksheet = spreadsheet.worksheet(WORKSHEET_NAME)
    except gspread.WorksheetNotFound:
        worksheet = spreadsheet.add_worksheet(title=WORKSHEET_NAME, rows=1000, cols=10)
        worksheet.append_row(HEADER)

    # 確保第一列是表頭（空白工作表時補上）
    if worksheet.row_values(1) != HEADER:
        if not worksheet.get_all_values():
            worksheet.append_row(HEADER)
    return worksheet


def load_votes_df():
    """
    從 Google Sheets 讀取所有投票紀錄，回傳 pandas DataFrame。
    若無資料則回傳僅含表頭欄位的空 DataFrame。
    """
    worksheet = get_worksheet()
    records = worksheet.get_all_records()  # list[dict]，自動以第一列為欄名
    if not records:
        return pd.DataFrame(columns=HEADER)
    df = pd.DataFrame(records)
    # 確保欄位齊全
    for col in HEADER:
        if col not in df.columns:
            df[col] = ""
    return df[HEADER]


def submit_vote(name, role, selected_slots):
    """
    送出投票，內建「防呆與覆蓋機制」：
      - 若該姓名已存在 → 更新（覆蓋）那一列
      - 若不存在 → 新增一列
    selected_slots：list[str]，會以 ";" 串接後存入單一儲存格。
    """
    worksheet = get_worksheet()
    name = name.strip()
    slots_str = ";".join(selected_slots)
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    new_row = [name, role, slots_str, timestamp]

    # 取得所有資料（含表頭），用來尋找是否已有同名紀錄
    all_values = worksheet.get_all_values()
    # all_values[0] 為表頭，資料從第 2 列（index 1）開始
    target_row_index = None
    for i, row in enumerate(all_values[1:], start=2):  # start=2 對應 Google Sheets 列號
        if row and row[0].strip() == name:
            target_row_index = i
            break

    if target_row_index:
        # 覆蓋：更新整列（A:D）
        worksheet.update(f"A{target_row_index}:D{target_row_index}", [new_row])
        return "updated"
    else:
        # 新增
        worksheet.append_row(new_row)
        return "created"


# ==============================================================================
# 【區塊 2】加權演算法 — 核心運算
# ==============================================================================

def get_active_slots():
    """回傳「未被黑名單排除」的時段清單（list[dict]）。"""
    return [s for s in ALL_SLOTS if s["label"] not in EXCLUDE_SLOTS]


def compute_scores(df):
    """
    依投票紀錄計算每個時段的加權分數與投票名單。

    規則：
      - 硬性條件：田老師沒勾選該時段 → 分數判定為 -1（老師不行）
      - 軟性條件：田老師有勾選時 → 分數 = 2 * 碩博票數 + 1 * 專題生票數

    回傳：dict，key 為時段 label，value 為：
      {
        "score":  int,                 # 加權分數（-1 表示老師不行）
        "teacher_ok": bool,            # 田老師是否出席
        "voters": list[(name, role)],  # 投此時段的名單（不含分數判定）
      }
    """
    active_labels = [s["label"] for s in get_active_slots()]
    # 初始化
    result = {
        label: {"score": 0, "teacher_ok": False, "voters": []}
        for label in active_labels
    }

    # 逐筆統計每個人勾選的時段
    for _, row in df.iterrows():
        name = str(row.get("name", "")).strip()
        role = str(row.get("role", "")).strip()
        slots_str = str(row.get("slots", ""))
        if not name or not slots_str:
            continue
        voted = [s for s in slots_str.split(";") if s]

        for label in voted:
            if label not in result:
                continue  # 跳過已被黑名單排除或不存在的時段
            result[label]["voters"].append((name, role))
            if role == ROLE_TEACHER:
                result[label]["teacher_ok"] = True
            # 累加軟性權重（田老師權重為 0，不影響）
            result[label]["score"] += ROLE_WEIGHT.get(role, 0)

    # 套用硬性條件：田老師沒出席 → 分數設為 -1
    for label, info in result.items():
        if not info["teacher_ok"]:
            info["score"] = -1

    return result


def get_top3(scores):
    """
    依加權分數由高到低排序，回傳前三名（僅取分數 > 0 且田老師可出席的時段）。
    回傳 list[(label, info)]。
    """
    valid = [
        (label, info)
        for label, info in scores.items()
        if info["teacher_ok"] and info["score"] > 0
    ]
    valid.sort(key=lambda x: x[1]["score"], reverse=True)
    return valid[:3]


# ==============================================================================
# 【區塊 3】Seaborn 熱力圖
# ==============================================================================

def draw_heatmap(scores, top_label=None):
    """
    繪製美化版時段熱力圖（圓角卡片風格）：
      X 軸 = 日期、Y 軸 = 時段、顏色越深綠 = 加權分數越高（越推薦）。
      - 田老師不行的時段：灰底 + 「老師不行」
      - 目前最佳時段：金色外框 + 「★ 最佳」
      - 沒有該時段組合：留白
    top_label：(date, time) tuple，會標示為最佳。
    回傳 matplotlib figure。
    """
    sns.set_style("white")
    active = get_active_slots()

    # 依 ALL_SLOTS 原始順序排列日期；時段由早到晚
    dates = []
    for s in ALL_SLOTS:
        if s in active and s["date"] not in dates:
            dates.append(s["date"])
    times = sorted({s["time"] for s in active})

    # 建立 (時段) x (日期) 的分數矩陣與票數矩陣
    score = pd.DataFrame(index=times, columns=dates, dtype=float)
    nvotes = {}
    for s in active:
        info = scores.get(s["label"], {})
        score.loc[s["time"], s["date"]] = info.get("score", np.nan)
        nvotes[(s["time"], s["date"])] = len(info.get("voters", []))

    n_rows, n_cols = len(times), len(dates)
    fig, ax = plt.subplots(figsize=(1.7 * n_cols + 1.4, 1.28 * n_rows + 1.6), dpi=170)
    fig.patch.set_facecolor("#ffffff")
    ax.set_facecolor("#ffffff")

    # 深海軍藍漸層：分數越高越深（= 越推薦）
    cmap = LinearSegmentedColormap.from_list(
        "meet", ["#dbe6f3", "#a9c0dd", "#5b85b8", "#2a4d7a", "#14253d"]
    )
    valid_vals = score.where(score > 0).values
    vmax = np.nanmax(valid_vals) if np.isfinite(np.nanmax(valid_vals)) else 1
    vmin = 0
    pad, r = 0.07, 0.16
    soft_shadow = [path_effects.withSimplePatchShadow(offset=(1.4, -1.6),
                                                      shadow_rgbFace="#0e2238", alpha=0.22)]

    for i, t in enumerate(times):
        for j, d in enumerate(dates):
            v = score.loc[t, d]
            if pd.isna(v):
                continue  # 沒有此時段，留白
            if v <= -1:  # 田老師不行 —— 刻意低調，讓有效時段更突出
                cell = FancyBboxPatch(
                    (j - 0.5 + pad, i - 0.5 + pad), 1 - 2 * pad, 1 - 2 * pad,
                    boxstyle=f"round,pad=0,rounding_size={r}",
                    linewidth=0, facecolor="#f1f4f8",
                )
                ax.add_patch(cell)
                ax.text(j, i, "老師不行", ha="center", va="center", color="#b6bfcc",
                        fontsize=10.5, fontweight="bold", fontproperties=_FONT_PROP)
            else:
                frac = (v - vmin) / (vmax - vmin) if vmax > vmin else 1.0
                face = cmap(0.18 + 0.82 * frac)
                cell = FancyBboxPatch(
                    (j - 0.5 + pad, i - 0.5 + pad), 1 - 2 * pad, 1 - 2 * pad,
                    boxstyle=f"round,pad=0,rounding_size={r}",
                    linewidth=0, facecolor=face,
                )
                cell.set_path_effects(soft_shadow)
                ax.add_patch(cell)
                tcolor = "white" if frac > 0.4 else "#14253d"
                subcolor = (1, 1, 1, 0.78) if frac > 0.4 else "#5b85b8"
                ax.text(j, i - 0.07, f"{int(v)}", ha="center", va="center", color=tcolor,
                        fontsize=25, fontweight="bold", fontproperties=_FONT_PROP)
                n = nvotes.get((t, d), 0)
                ax.text(j, i + 0.28, f"{n} 票", ha="center", va="center", color=subcolor,
                        fontsize=9.5, fontproperties=_FONT_PROP)

    # 標出目前最佳時段（金框 + ★ 最佳）
    if top_label:
        td, tt = top_label
        if td in dates and tt in times:
            jx, iy = dates.index(td), times.index(tt)
            ax.add_patch(FancyBboxPatch(
                (jx - 0.5 + pad, iy - 0.5 + pad), 1 - 2 * pad, 1 - 2 * pad,
                boxstyle=f"round,pad=0,rounding_size={r}",
                linewidth=3.4, edgecolor="#e0a800", facecolor="none", zorder=5,
            ))
            ax.text(jx, iy - 0.36, "★ 最佳", ha="center", va="center", fontsize=11,
                    fontweight="bold", color="#e0a800", zorder=6, fontproperties=_FONT_PROP)

    # 日期表頭（軍藍粗體）與時段標籤
    ax.set_xlim(-0.5, n_cols - 0.5)
    ax.set_ylim(n_rows - 0.4, -0.6)
    ax.set_xticks(range(n_cols))
    ax.set_xticklabels(dates, fontsize=13.5, color="#1e3a5f",
                       fontweight="bold", fontproperties=_FONT_PROP)
    ax.set_yticks(range(n_rows))
    ax.set_yticklabels(times, fontsize=11.5, color="#5b6b7f", fontproperties=_FONT_PROP)
    ax.xaxis.set_ticks_position("top")
    ax.xaxis.set_label_position("top")
    for spine in ax.spines.values():
        spine.set_visible(False)
    ax.tick_params(length=0)
    ax.set_xlabel("")
    ax.set_title("Group Meeting 時段投票熱力圖", fontsize=17,
                 color="#1e3a5f", pad=30, fontweight="bold", fontproperties=_FONT_PROP)
    fig.text(0.5, 0.012,
             "顏色越深 = 加權分數越高（金框 ★ 為目前最佳）　｜　淺灰 = 田老師無法出席",
             ha="center", fontsize=9.5, color="#9aa7b5", fontproperties=_FONT_PROP)
    fig.tight_layout(rect=[0, 0.035, 1, 1])
    return fig


# ==============================================================================
# 【區塊 4】Slack 推播
# ==============================================================================

def push_to_slack(best_label, best_score):
    """
    透過 Slack Incoming Webhook 發送排版整齊的訊息。
    Webhook URL 從 st.secrets["slack"]["webhook_url"] 讀取。
    """
    webhook_url = st.secrets.get("slack", {}).get("webhook_url", "")
    if not webhook_url:
        return False, "尚未設定 Slack Webhook URL（請於 Secrets 設定 slack.webhook_url）"

    text = (
        f"根據網頁調查結果，下週 Group Meeting 推薦時間為："
        f"*{best_label}*，加權得分：*{best_score} 分*！ <!channel>"
    )
    payload = {
        "text": text,
        "blocks": [
            {
                "type": "section",
                "text": {
                    "type": "mrkdwn",
                    "text": (
                        ":mega: *Group Meeting 時段調查結果出爐！*\n\n"
                        f"> 推薦時間：*{best_label}*\n"
                        f"> 加權得分：*{best_score} 分*\n\n"
                        "請大家準時出席 :calendar: <!channel>"
                    ),
                },
            }
        ],
    }
    try:
        resp = requests.post(webhook_url, data=json.dumps(payload),
                             headers={"Content-Type": "application/json"}, timeout=10)
        if resp.status_code == 200:
            return True, "已成功推播至 Slack！"
        return False, f"Slack 回傳錯誤：{resp.status_code} {resp.text}"
    except Exception as e:
        return False, f"推播失敗：{e}"


# ==============================================================================
# 【區塊 5】UI 美化：CSS 樣式與 HTML 元件
# ==============================================================================

# 各身份對應的標籤顏色（chip）
ROLE_CHIP_CLASS = {
    ROLE_TEACHER: "chip-teacher",
    ROLE_GRAD: "chip-grad",
    ROLE_PROJECT: "chip-proj",
}

CUSTOM_CSS = """
<style>
/* ---- 全域 ---- */
.block-container { max-width: 1080px; padding-top: 1.4rem; padding-bottom: 4rem; }
html, body, [class*="css"] { font-family: -apple-system, "PingFang TC", "Microsoft JhengHei", "Noto Sans TC", sans-serif; }

/* ---- Hero 標題 ---- */
.hero {
  background: linear-gradient(135deg, #2a4d7a 0%, #14253d 100%);
  border-radius: 22px; padding: 30px 34px; color: #fff;
  box-shadow: 0 12px 30px rgba(20,37,61,.28); margin-bottom: 26px;
}
.hero h1 { margin: 0; font-size: 30px; font-weight: 800; letter-spacing: .5px; }
.hero p { margin: 10px 0 0; font-size: 15px; opacity: .92; }
.hero .pills { margin-top: 16px; display: flex; flex-wrap: wrap; gap: 8px; }
.hero .pill {
  background: rgba(255,255,255,.18); border: 1px solid rgba(255,255,255,.28);
  padding: 5px 13px; border-radius: 999px; font-size: 12.5px; font-weight: 600;
}

/* ---- 區段標題 ---- */
.section-head { display: flex; align-items: center; gap: 12px; margin: 8px 0 18px; }
.section-num {
  width: 34px; height: 34px; border-radius: 11px; flex: none;
  background: linear-gradient(135deg, #1e3a5f, #14253d); color: #fff;
  display: flex; align-items: center; justify-content: center; font-weight: 800; font-size: 17px;
  box-shadow: 0 4px 10px rgba(20,37,61,.25);
}
.section-title { font-size: 21px; font-weight: 800; color: #1f2933; }

/* ---- Top 3 排名卡 ---- */
.top3-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 16px; }
.rank-card {
  border-radius: 18px; padding: 20px 22px; background: #fff;
  border: 1px solid #eceff1; box-shadow: 0 6px 18px rgba(15,40,30,.06);
  position: relative; overflow: hidden;
}
.rank-card::before { content:""; position:absolute; left:0; top:0; bottom:0; width:6px; }
.rank-1::before { background: linear-gradient(#f6c945,#e0a800); }
.rank-2::before { background: linear-gradient(#cfd8dc,#9aa7ad); }
.rank-3::before { background: linear-gradient(#e7b58a,#cf915c); }
.rank-1 { box-shadow: 0 10px 26px rgba(224,168,0,.18); }
.rank-top { display: flex; align-items: center; gap: 8px; margin-bottom: 6px; }
.rank-medal { font-size: 22px; }
.rank-tag { font-size: 12px; font-weight: 700; color: #90a4ae; letter-spacing: .5px; }
.rank-slot { font-size: 17px; font-weight: 800; color: #1f2933; margin: 2px 0 8px; }
.rank-score { font-size: 34px; font-weight: 900; color: #1e3a5f; line-height: 1; }
.rank-score span { font-size: 15px; font-weight: 700; color: #78909c; margin-left: 3px; }
.rank-sub { margin-top: 9px; font-size: 12.5px; color: #90a4ae; }

/* ---- 投票名單卡 ---- */
.slot-list { display: grid; grid-template-columns: repeat(auto-fill, minmax(330px, 1fr)); gap: 14px; }
.slot-card {
  border: 1px solid #eceff1; border-radius: 16px; padding: 15px 17px; background: #fff;
  box-shadow: 0 4px 14px rgba(15,40,30,.05);
}
.slot-card.is-best { border-color: #f4b400; box-shadow: 0 6px 18px rgba(244,180,0,.18); }
.slot-card.is-no { background: #fafbfc; }
.slot-head { display: flex; align-items: center; justify-content: space-between; gap: 10px; margin-bottom: 11px; }
.slot-name { font-size: 15.5px; font-weight: 800; color: #1f2933; }
.badge { font-size: 12px; font-weight: 800; padding: 4px 11px; border-radius: 999px; white-space: nowrap; }
.badge-ok { background: #e7eef7; color: #14253d; }
.badge-no { background: #eceff1; color: #90a4ae; }
.chips { display: flex; flex-wrap: wrap; gap: 6px; }
.chip { font-size: 12px; font-weight: 600; padding: 4px 10px; border-radius: 8px; }
.chip-teacher { background: #fff5d6; color: #a87900; border: 1px solid #ffe7a3; }
.chip-grad { background: #e7eef7; color: #1e3a5f; border: 1px solid #c2d3e8; }
.chip-proj { background: #e0f2f1; color: #00695c; border: 1px solid #b2dfdb; }
.chip-empty { color: #b0bec5; font-size: 12.5px; }

/* ---- 按鈕 ---- */
.stButton > button, .stFormSubmitButton > button {
  border-radius: 12px; font-weight: 700; border: none; padding: 10px 18px;
}
.stFormSubmitButton > button {
  background: linear-gradient(135deg, #1e3a5f, #14253d); color: #fff;
}
.stFormSubmitButton > button:hover { filter: brightness(1.06); color: #fff; }

/* ---- 表單容器 ---- */
[data-testid="stForm"] {
  border: 1px solid #eceff1; border-radius: 18px; padding: 22px 24px;
  box-shadow: 0 6px 18px rgba(15,40,30,.05); background: #fff;
}

/* ---- 最佳結果橫幅 ---- */
.best-banner {
  background: linear-gradient(135deg, #fff8e6, #fff2cc); border: 1px solid #ffe39a;
  border-radius: 16px; padding: 16px 20px; margin-bottom: 14px;
  display: flex; align-items: center; gap: 14px;
}
.best-banner .bb-icon { font-size: 26px; }
.best-banner .bb-text { font-size: 15px; color: #6b5200; }
.best-banner .bb-text b { color: #1f2933; font-size: 16.5px; }
</style>
"""


def section_header(num, title):
    """輸出統一風格的區段標題。"""
    st.markdown(
        f'<div class="section-head"><div class="section-num">{num}</div>'
        f'<div class="section-title">{html.escape(title)}</div></div>',
        unsafe_allow_html=True,
    )


def render_top3(top3):
    """以排名卡呈現 Top 3。"""
    medals = ["🥇", "🥈", "🥉"]
    tags = ["第 1 推薦", "第 2 推薦", "第 3 推薦"]
    cards = []
    for i, (label, info) in enumerate(top3):
        n_grad = sum(1 for _, r in info["voters"] if r == ROLE_GRAD)
        n_proj = sum(1 for _, r in info["voters"] if r == ROLE_PROJECT)
        cards.append(
            f'<div class="rank-card rank-{i+1}">'
            f'<div class="rank-top"><span class="rank-medal">{medals[i]}</span>'
            f'<span class="rank-tag">{tags[i]}</span></div>'
            f'<div class="rank-slot">{html.escape(label)}</div>'
            f'<div class="rank-score">{info["score"]}<span>分</span></div>'
            f'<div class="rank-sub">碩博 {n_grad} 票 · 專題 {n_proj} 票 · 田老師可出席 ✓</div>'
            f'</div>'
        )
    st.markdown(f'<div class="top3-grid">{"".join(cards)}</div>', unsafe_allow_html=True)


def render_voter_list(active_slots, scores, best_label=None):
    """以卡片呈現各時段投票名單。"""
    cards = []
    for slot in active_slots:
        label = slot["label"]
        info = scores.get(label, {"score": 0, "teacher_ok": False, "voters": []})
        voters = info["voters"]
        is_best = (label == best_label)
        if info["teacher_ok"]:
            badge = f'<span class="badge badge-ok">加權 {info["score"]} 分</span>'
            card_cls = "slot-card is-best" if is_best else "slot-card"
        else:
            badge = '<span class="badge badge-no">老師不行</span>'
            card_cls = "slot-card is-no"
        if voters:
            chips = "".join(
                f'<span class="chip {ROLE_CHIP_CLASS.get(r, "chip-proj")}">{html.escape(n)}</span>'
                for n, r in voters
            )
        else:
            chips = '<span class="chip-empty">尚無人投票</span>'
        star = "★ " if is_best else ""
        cards.append(
            f'<div class="{card_cls}"><div class="slot-head">'
            f'<span class="slot-name">{star}{html.escape(label)}</span>{badge}</div>'
            f'<div class="chips">{chips}</div></div>'
        )
    st.markdown(f'<div class="slot-list">{"".join(cards)}</div>', unsafe_allow_html=True)


# ==============================================================================
# 【區塊 6】Streamlit 前台主介面
# ==============================================================================

def main():
    st.set_page_config(page_title="實驗室 Group Meeting 時段調查", page_icon="🗳️", layout="wide")
    st.markdown(CUSTOM_CSS, unsafe_allow_html=True)

    # Hero 標題
    st.markdown(
        '<div class="hero"><h1>🗳️ 實驗室 Group Meeting 時段調查</h1>'
        '<p>仿 Doodle 投票，依身份自動加權，田老師出席為必要條件，結果可一鍵推播 Slack。</p></div>',
        unsafe_allow_html=True,
    )

    active_slots = get_active_slots()
    active_labels = [s["label"] for s in active_slots]

    # ----------------------------------------------------------------------
    # （一）投票區
    # ----------------------------------------------------------------------
    section_header(1, "填寫你的可出席時段")
    with st.form("vote_form", clear_on_submit=False):
        col1, col2 = st.columns([2, 1])
        with col1:
            name = st.text_input("你的姓名", placeholder="請輸入姓名（重複送出會覆蓋舊投票）")
        with col2:
            role = st.selectbox("你的身份", ROLE_OPTIONS, index=0)

        st.markdown("**勾選你可以出席的時段（可複選）：**")
        # 用 checkbox 列表呈現待選時段
        selected = []
        # 以兩欄排列，畫面較整齊
        cols = st.columns(2)
        for i, slot in enumerate(active_slots):
            with cols[i % 2]:
                if st.checkbox(slot["label"], key=f"slot_{slot['label']}"):
                    selected.append(slot["label"])

        submitted = st.form_submit_button("✅ 送出 / 更新我的投票", use_container_width=True)

    if submitted:
        if not name.strip():
            st.error("請先輸入姓名再送出。")
        elif not selected:
            st.warning("你尚未勾選任何時段，仍可送出（代表目前都不行），確定的話請再次送出。")
            action = submit_vote(name, role, selected)
            st.success("已記錄你目前沒有可出席的時段。")
        else:
            action = submit_vote(name, role, selected)
            if action == "updated":
                st.success(f"已覆蓋更新 {name} 的投票（共 {len(selected)} 個時段）。")
            else:
                st.success(f"已記錄 {name} 的投票（共 {len(selected)} 個時段）。")

    st.divider()

    # ----------------------------------------------------------------------
    # 讀取目前所有投票並計算分數
    # ----------------------------------------------------------------------
    df = load_votes_df()
    scores = compute_scores(df)
    top3 = get_top3(scores)

    # ----------------------------------------------------------------------
    # （二）系統推薦 Top 3
    # ----------------------------------------------------------------------
    section_header(2, "🏆 系統當前推薦最佳時段 Top 3")
    if not top3:
        st.info("目前還沒有「田老師有勾選且有加權分數」的時段，等大家投票後就會出現推薦。")
    else:
        render_top3(top3)

    st.divider()

    # ----------------------------------------------------------------------
    # （三）熱力圖看板
    # ----------------------------------------------------------------------
    section_header(3, "📊 當前投票結果看板（熱力圖）")
    best_label = top3[0][0] if top3 else None
    if df.empty:
        st.info("目前還沒有任何投票紀錄。")
    else:
        # 把 Top1 的 label 轉成 (date, time) 以便在熱力圖標示「★ 最佳」
        top_label = None
        if best_label:
            best_slot = next((s for s in ALL_SLOTS if s["label"] == best_label), None)
            if best_slot:
                top_label = (best_slot["date"], best_slot["time"])
        c1, c2, c3 = st.columns([1, 8, 1])
        with c2:
            fig = draw_heatmap(scores, top_label=top_label)
            st.pyplot(fig)

    # 各時段投票名單（卡片化，身份以顏色標籤區分）
    st.markdown("##### 各時段已投票名單")
    render_voter_list(active_slots, scores, best_label=best_label)

    st.divider()

    # ----------------------------------------------------------------------
    # （四）Slack 推播
    # ----------------------------------------------------------------------
    section_header(4, "📢 公布結果")
    if top3:
        best_label, best_info = top3[0]
        st.markdown(
            f'<div class="best-banner"><span class="bb-icon">🏆</span>'
            f'<span class="bb-text">目前最佳時段：<b>{html.escape(best_label)}</b>'
            f'（加權 {best_info["score"]} 分）</span></div>',
            unsafe_allow_html=True,
        )
        if st.button("📢 將目前最佳結果推播至 Slack", type="primary"):
            ok, msg = push_to_slack(best_label, best_info["score"])
            if ok:
                st.success(msg)
            else:
                st.error(msg)
    else:
        st.info("尚無可推播的最佳時段。")

    # 黑名單提示（方便管理者確認）
    if EXCLUDE_SLOTS:
        st.caption(f"🚫 已排除時段：{'、'.join(EXCLUDE_SLOTS)}")


if __name__ == "__main__":
    main()
