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
from datetime import datetime

import gspread
import matplotlib
import matplotlib.pyplot as plt
import pandas as pd
import requests
import seaborn as sns
import streamlit as st
from google.oauth2.service_account import Credentials

# ------------------------------------------------------------------------------
# matplotlib 後端與中文字型設定
# Streamlit Cloud 預設沒有中文字型，若 X/Y 軸標籤含中文會變成方框（豆腐字）。
# 這裡先嘗試常見中文字型；雲端可在 requirements 安裝字型或改用英文時段名稱。
# ------------------------------------------------------------------------------
matplotlib.rcParams["font.sans-serif"] = [
    "Microsoft JhengHei",  # Windows 微軟正黑體
    "PingFang TC",          # macOS
    "Noto Sans CJK TC",     # Linux（雲端建議安裝）
    "Heiti TC",
    "Arial Unicode MS",
    "DejaVu Sans",
]
matplotlib.rcParams["axes.unicode_minus"] = False  # 修正負號顯示

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

def draw_heatmap(scores):
    """
    以 Seaborn 繪製時段熱力圖：
      X 軸 = 日期、Y 軸 = 時段、顏色深淺 = 加權分數。
    回傳 matplotlib figure。
    """
    active = get_active_slots()
    dates = sorted({s["date"] for s in active}, key=lambda d: [t["date"] for t in ALL_SLOTS].index(d))
    times = sorted({s["time"] for s in active})

    # 建立 (時段 time) x (日期 date) 的分數矩陣
    matrix = pd.DataFrame(index=times, columns=dates, dtype=float)
    for s in active:
        score = scores.get(s["label"], {}).get("score", 0)
        matrix.loc[s["time"], s["date"]] = score

    fig, ax = plt.subplots(figsize=(max(6, len(dates) * 1.4), max(3, len(times) * 0.9)))
    sns.heatmap(
        matrix,
        annot=True,            # 顯示分數數字
        fmt=".0f",
        cmap="YlOrRd",         # 顏色越深 = 分數越高
        linewidths=0.5,
        linecolor="white",
        cbar_kws={"label": "加權分數"},
        mask=matrix.isnull(),  # 沒有該組合的格子留白
        ax=ax,
    )
    ax.set_xlabel("日期")
    ax.set_ylabel("時段")
    ax.set_title("Group Meeting 時段加權熱力圖（-1 = 老師不行）")
    plt.xticks(rotation=0)
    plt.yticks(rotation=0)
    fig.tight_layout()
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
        fig = draw_heatmap(scores)
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