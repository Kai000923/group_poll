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

import json
import os
from datetime import datetime

import gspread
import matplotlib
import matplotlib.font_manager as fm
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

    # 建立 (時段) x (日期) 的分數矩陣
    score = pd.DataFrame(index=times, columns=dates, dtype=float)
    for s in active:
        score.loc[s["time"], s["date"]] = scores.get(s["label"], {}).get("score", np.nan)

    n_rows, n_cols = len(times), len(dates)
    fig, ax = plt.subplots(figsize=(1.55 * n_cols + 1.2, 1.15 * n_rows + 1.2), dpi=160)
    fig.patch.set_facecolor("#ffffff")
    ax.set_facecolor("#ffffff")

    # 綠色漸層：分數越高越深綠（= 越推薦）
    cmap = LinearSegmentedColormap.from_list(
        "meet", ["#e8f6ef", "#9fdcc0", "#4cb18a", "#1f8f63", "#0c6b46"]
    )
    valid_vals = score.where(score > 0).values
    vmax = np.nanmax(valid_vals) if np.isfinite(np.nanmax(valid_vals)) else 1
    vmin = 0
    pad = 0.06

    for i, t in enumerate(times):
        for j, d in enumerate(dates):
            v = score.loc[t, d]
            if pd.isna(v):
                continue  # 沒有此時段，留白
            if v <= -1:  # 田老師不行
                face, edge = "#eceff1", "#cfd8dc"
                txt, tcolor, fsize = "老師不行", "#90a4ae", 12
            else:
                frac = (v - vmin) / (vmax - vmin) if vmax > vmin else 1.0
                face, edge = cmap(0.15 + 0.85 * frac), "white"
                txt, fsize = f"{int(v)}", 20
                tcolor = "white" if frac > 0.45 else "#0c6b46"
            ax.add_patch(FancyBboxPatch(
                (j - 0.5 + pad, i - 0.5 + pad), 1 - 2 * pad, 1 - 2 * pad,
                boxstyle="round,pad=0,rounding_size=0.12",
                linewidth=2, edgecolor=edge, facecolor=face,
            ))
            ax.text(j, i, txt, ha="center", va="center", color=tcolor,
                    fontsize=fsize, fontweight="bold", fontproperties=_FONT_PROP)

    # 標出目前最佳時段（金框 + ★ 最佳）
    if top_label:
        td, tt = top_label
        if td in dates and tt in times:
            jx, iy = dates.index(td), times.index(tt)
            ax.add_patch(FancyBboxPatch(
                (jx - 0.5 + pad, iy - 0.5 + pad), 1 - 2 * pad, 1 - 2 * pad,
                boxstyle="round,pad=0,rounding_size=0.12",
                linewidth=3.2, edgecolor="#f4b400", facecolor="none", zorder=5,
            ))
            ax.text(jx, iy - 0.33, "★ 最佳", ha="center", va="center", fontsize=11,
                    fontweight="bold", color="#f4b400", zorder=6, fontproperties=_FONT_PROP)

    ax.set_xlim(-0.5, n_cols - 0.5)
    ax.set_ylim(n_rows - 0.5, -0.5)
    ax.set_xticks(range(n_cols))
    ax.set_xticklabels(dates, fontsize=13, color="#37474f", fontproperties=_FONT_PROP)
    ax.set_yticks(range(n_rows))
    ax.set_yticklabels(times, fontsize=12, color="#37474f", fontproperties=_FONT_PROP)
    ax.xaxis.set_ticks_position("top")
    ax.xaxis.set_label_position("top")
    for spine in ax.spines.values():
        spine.set_visible(False)
    ax.tick_params(length=0)
    ax.set_title("Group Meeting 時段投票熱力圖", fontsize=16,
                 color="#1f8f63", pad=26, fontproperties=_FONT_PROP)
    fig.text(0.5, 0.015,
             "顏色越深 = 加權分數越高（金框 ★ 為目前最佳）　｜　灰底 = 田老師無法出席",
             ha="center", fontsize=9.5, color="#78909c", fontproperties=_FONT_PROP)
    fig.tight_layout(rect=[0, 0.04, 1, 1])
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
# 【區塊 5】Streamlit 前台介面
# ==============================================================================

def main():
    st.set_page_config(page_title="實驗室 Group Meeting 時段調查", page_icon="🗳️", layout="wide")
    st.title("🗳️ 實驗室 Group Meeting 時段調查")
    st.caption("仿 Doodle 投票 · 身份加權 · 田老師為必要條件 · 結果可一鍵推播 Slack")

    active_slots = get_active_slots()
    active_labels = [s["label"] for s in active_slots]

    # ----------------------------------------------------------------------
    # （一）投票區
    # ----------------------------------------------------------------------
    st.header("① 填寫你的可出席時段")
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
    st.header("② 🏆 系統當前推薦最佳時段 Top 3")
    if not top3:
        st.info("目前還沒有「田老師有勾選且有加權分數」的時段，等大家投票後就會出現推薦。")
    else:
        medals = ["🥇", "🥈", "🥉"]
        cols = st.columns(len(top3))
        for i, (label, info) in enumerate(top3):
            with cols[i]:
                st.metric(label=f"{medals[i]} {label}", value=f"{info['score']} 分",
                          help=f"投票人數：{len(info['voters'])}")

    st.divider()

    # ----------------------------------------------------------------------
    # （三）熱力圖看板
    # ----------------------------------------------------------------------
    st.header("③ 📊 當前投票結果看板（熱力圖）")
    if df.empty:
        st.info("目前還沒有任何投票紀錄。")
    else:
        # 把 Top1 的 label 轉成 (date, time) 以便在熱力圖標示「★ 最佳」
        top_label = None
        if top3:
            best_slot = next((s for s in ALL_SLOTS if s["label"] == top3[0][0]), None)
            if best_slot:
                top_label = (best_slot["date"], best_slot["time"])
        fig = draw_heatmap(scores, top_label=top_label)
        st.pyplot(fig)

    # 各時段投票名單
    st.subheader("各時段已投票名單")
    for slot in active_slots:
        label = slot["label"]
        info = scores.get(label, {"score": 0, "teacher_ok": False, "voters": []})
        voters = info["voters"]
        if info["teacher_ok"]:
            score_tag = f"加權 {info['score']} 分"
        else:
            score_tag = "❌ 老師不行（-1）"
        names = "、".join([f"{n}（{r}）" for n, r in voters]) if voters else "（尚無人投票）"
        st.markdown(f"**{label}** — {score_tag}　｜　{names}")

    st.divider()

    # ----------------------------------------------------------------------
    # （四）Slack 推播
    # ----------------------------------------------------------------------
    st.header("④ 📢 公布結果")
    if top3:
        best_label, best_info = top3[0]
        st.write(f"目前最佳時段：**{best_label}**（加權 {best_info['score']} 分）")
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
