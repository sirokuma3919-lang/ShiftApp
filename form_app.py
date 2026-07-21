import streamlit as st
import pandas as pd
import os
import datetime
import calendar
import unicodedata
from filelock import FileLock
import requests
import io  # 空中でExcelを作るための部品

# ==========================================
# ⚙️ 1. 設定値（定数）の定義
# ==========================================
today = datetime.date.today()
TARGET_YEAR = today.year + 1 if today.month == 12 else today.year
TARGET_MONTH = 1 if today.month == 12 else today.month + 1

# ※ローカル環境用のファイル名（クラウド上ではバックアップ用として動作します）
CSV_REQUESTS = f"【プログラム用】{TARGET_MONTH}月シフト提出状況.csv"
EXCEL_REQUESTS = f"【店長確認用】{TARGET_MONTH}月シフト提出状況.xlsx"
LOCK_FILE = f"{TARGET_MONTH}月シフト提出状況.lock"

DEPARTMENTS = ["選択してください", "家電", "季節AV", "情報", "通信"]
ADMIN_PASSWORD = "password"

# 🌟【超重要】取得した「ウェブアプリのURL」をここに貼り付けてください！
GAS_URL = "https://script.google.com/macros/s/AKfycbx20gcPFY7CKjGRjNMNHI9zNgwzmC_i8u1Wsw1r2BrpTtYUmB06ejgWFGKtLbJaTlPkGw/exec"

st.set_page_config(page_title="シフト希望提出フォーム", layout="wide")

# ==========================================
# 2. 関数の定義
# ==========================================
def init_session_state():
    """セッションステートの初期化"""
    if "confirm_mode" not in st.session_state:
        st.session_state.confirm_mode = False
    if "is_submitted" not in st.session_state:
        st.session_state.is_submitted = False
    if "excel_warning" not in st.session_state:
        st.session_state.excel_warning = False

def get_month_days():
    """来月の日数と日付ラベルのリストを取得する"""
    _, num_days = calendar.monthrange(TARGET_YEAR, TARGET_MONTH)
    weekdays_ja = ["月", "火", "水", "木", "金", "土", "日"]
    
    day_labels = []
    for d in range(1, num_days + 1):
        dt = datetime.date(TARGET_YEAR, TARGET_MONTH, d)
        day_labels.append(f"{d}日（{weekdays_ja[dt.weekday()]}）")
        
    return num_days, day_labels

def save_shift_data(emp_code, name, department, target_days, shift_requests):
    """シフトデータを保存する（Googleスプレッドシートへ自動送信）"""
    safe_emp_code = unicodedata.normalize('NFKC', emp_code).strip()
    safe_name = name.strip()
    
    # 送信用のデータを作成
    data = {
        "対象月": f"{TARGET_MONTH}",
        "提出日時": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "従業員コード": safe_emp_code,
        "名前": safe_name, 
        "部門": department,
        "希望出勤時間": target_days
    }
    data.update(shift_requests)
    
    # 🌟 1. Googleスプレッドシート（GAS）へデータを送信
    try:
        response = requests.post(GAS_URL, json=data, timeout=10)
        if response.status_code != 200 or response.json().get("status") != "success":
            return "gas_error"
    except Exception:
        return "gas_error"
    
    # 🌟 2. 念のためサーバーのローカル環境にもバックアップを残す
    df_submit = pd.DataFrame([data])
    lock = FileLock(LOCK_FILE)
    with lock:
        if os.path.exists(CSV_REQUESTS):
            try:
                df_existing = pd.read_csv(CSV_REQUESTS, encoding="utf-8-sig")
                df_final = pd.concat([df_existing, df_submit], ignore_index=True)
            except Exception:
                df_final = df_submit
        else:
            df_final = df_submit
            
        try:
            df_final.to_csv(CSV_REQUESTS, index=False, encoding="utf-8-sig")
            df_final.to_excel(EXCEL_REQUESTS, index=False)
        except Exception:
            pass # クラウド環境ではエラーを無視して進める
            
    return "success"

def generate_styled_calendar(day_labels, shift_requests):
    """色付きのカレンダー表を作成する"""
    calendar.setfirstweekday(calendar.SUNDAY)
    cal_matrix = calendar.monthcalendar(TARGET_YEAR, TARGET_MONTH)
    cal_data = []
    
    for week in cal_matrix:
        week_data = []
        for d in week:
            if d == 0:
                week_data.append("")
            else:
                label = day_labels[d-1]
                req = shift_requests.get(label, "")
                display_text = "出" if req == "" else req
                week_data.append(f"{d}日: {display_text}")
        cal_data.append(week_data)
        
    df = pd.DataFrame(cal_data, columns=["日", "月", "火", "水", "木", "金", "土"])
    
    def style_cells(data):
        styles = pd.DataFrame("", index=data.index, columns=data.columns)
        for row in data.index:
            for col in data.columns:
                val = str(data.loc[row, col])
                css = []
                if col == "土": css.append("background-color: #E6F2FF")
                elif col == "日": css.append("background-color: #FFE6E6")
                if "休" in val: css.append("color: #FF0000; font-weight: bold")
                styles.loc[row, col] = "; ".join(css)
        return styles

    return df.style.apply(style_cells, axis=None)

def show_admin_panel():
    """右上の店長用確認パネル（原因調査モード付き）"""
    with st.popover("店長専用メニュー", use_container_width=True):
        admin_pass = st.text_input("店長用パスワードを入力", type="password")
        if admin_pass == ADMIN_PASSWORD:
            st.write("---")
            st.markdown("#### 📥 シフトデータのダウンロード")
            
            if st.button("最新のExcelを作成する", use_container_width=True):
                with st.spinner("クラウドからデータを取得中..."):
                    try:
                        response = requests.get(f"{GAS_URL}?type=shift&month={TARGET_MONTH}", timeout=15)
                        if response.status_code == 200:
                            raw_data = response.json()
                            if len(raw_data) > 1:
                                df_dl = pd.DataFrame(raw_data[1:], columns=raw_data[0])
                                output = io.BytesIO()
                                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                                    df_dl.to_excel(writer, index=False, sheet_name=f'{TARGET_MONTH}月シフト提出')
                                excel_data = output.getvalue()
                                st.success("✅ 準備完了！下のボタンから保存してください。")
                                st.download_button(
                                    label="📊 Excelファイル（.xlsx）を保存",
                                    data=excel_data,
                                    file_name=EXCEL_REQUESTS,
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                    use_container_width=True,
                                    type="primary"
                                )
                            else:
                                st.info("まだ誰もシフトを提出していません。")
                        else:
                            st.error("データの取得に失敗しました。")
                    except Exception as e:
                        st.error(f"通信エラーが発生しました: {e}")

            st.write("---")
            st.markdown("#### 👤 提出状況一覧（未提出チェック）")
            
            with st.spinner("名簿と提出状況を照合中..."):
                try:
                    res_member = requests.get(f"{GAS_URL}?type=member", timeout=15)
                    res_shift = requests.get(f"{GAS_URL}?type=shift&month={TARGET_MONTH}", timeout=15)
                    
                    if res_member.status_code == 200 and res_shift.status_code == 200:
                        raw_member = res_member.json()
                        raw_shift = res_shift.json()
                        
                        # 🌟【調査エリア】システムが何を読み取っているか画面に表示します！
                        st.write(f"🔍 【調査1】名簿のデータ行数: {len(raw_member)}行（見出し含む）")
                        
                        if len(raw_member) > 1:
                            df_members = pd.DataFrame(raw_member[1:], columns=raw_member[0])
                            df_members.columns = df_members.columns.str.strip()
                            
                            st.write(f"🔍 【調査2】読み取った列名: {df_members.columns.tolist()}")
                            
                            if set(["従業員コード", "名前", "部門"]).issubset(df_members.columns):
                                df_status = df_members[["従業員コード", "名前", "部門"]].copy()
                                df_status["従業員コード"] = df_status["従業員コード"].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
                                
                                submitted_codes = []
                                if len(raw_shift) > 1:
                                    df_submitted = pd.DataFrame(raw_shift[1:], columns=raw_shift[0])
                                    df_submitted.columns = df_submitted.columns.str.strip()
                                    if "従業員コード" in df_submitted.columns:
                                        submitted_codes = df_submitted["従業員コード"].astype(str).str.replace(r'\.0$', '', regex=True).str.strip().tolist()
                                
                                df_status["提出状況"] = df_status["従業員コード"].apply(
                                    lambda x: "提出済" if x in submitted_codes else "未提出"
                                )
                                df_status = df_status.sort_values("提出状況", ascending=False)
                                
                                def highlight_unsubmitted(row):
                                    return ['background-color: #FFE6E6' if row['提出状況'] == '未提出' else ''] * len(row)
                                
                                st.dataframe(df_status.style.apply(highlight_unsubmitted, axis=1), hide_index=True, use_container_width=True)
                            else:
                                st.error("❌ エラー：必須の3つの列（従業員コード, 名前, 部門）のどれかが見つかりません。調査2の文字を確認してください。")
                        else:
                            st.warning("⚠️ スプレッドシートの「名簿」タブが空っぽ（見出しのみ、または0行）です。")
                    else:
                        st.error(f"❌ 通信エラー: 名簿({res_member.status_code}), シフト({res_shift.status_code})")
                except Exception as e:
                    st.error(f"❌ 読み込みエラー: {e}")

# ==========================================
# 3. メインの画面描画
# ==========================================
init_session_state()
input_disabled = st.session_state.confirm_mode or st.session_state.is_submitted
num_days, day_labels = get_month_days()

# --- タイトル ＆ 管理者パネル ---
col_title, col_admin = st.columns([4, 1])
with col_title:
    st.markdown(f"### {TARGET_YEAR}年{TARGET_MONTH}月分 シフト希望提出フォーム")
    st.write("時間の希望がある日だけ選択してください。希望がない日は「希望なし」のままでOKです。\n\nシステムの不具合がありましたら、不具合の画面の写真とともに、箭内にご連絡ください。")
with col_admin:
    show_admin_panel()

st.divider()

# --- 入力エリア ---
col_left, col_right = st.columns([1, 2])

with col_left:
    st.subheader("基本情報")
    name = st.text_input("1. お名前", disabled=input_disabled)
    emp_code = st.text_input("2. 従業員コード（数字）", disabled=input_disabled)
    department = st.selectbox("3. 部門を選択", DEPARTMENTS, disabled=input_disabled)
    target_days = st.number_input("4. 希望出勤時間（希望がない場合は0を入力してください）", min_value=0, max_value=120, value=0, step=1, disabled=input_disabled)

with col_right:
    st.subheader("日ごとの希望")
    tab_titles = ["1日〜7日", "8日〜14日", "15日〜21日", "22日〜28日"]
    day_groups = [day_labels[0:7], day_labels[7:14], day_labels[14:21], day_labels[21:28]]
    if num_days > 28:
        tab_titles.append(f"29日〜{num_days}日")
        day_groups.append(day_labels[28:num_days])
    
    tabs = st.tabs(tab_titles)
    shift_requests = {}
    
    for i, tab in enumerate(tabs):
        with tab:
            for label in day_groups[i]:
                choice = st.radio(f"**{label}**", ["希望なし", "休", "早", "遅", "時間指定"], horizontal=True, disabled=input_disabled)
                if choice == "時間指定":
                    specific_time = st.text_input(f"↳ 【{label}】希望時間を入力（例：11-19, 早-15, 14-L）", key=f"time_{label}", disabled=input_disabled)
                    shift_requests[label] = specific_time if specific_time else "時間指定(未入力)"
                elif choice == "希望なし":
                    shift_requests[label] = "出" 
                else:
                    shift_requests[label] = choice
                st.write("---")

st.divider()

# --- プレビュー ＆ ボタンエリア ---
if st.session_state.is_submitted:
    st.success(f"{name}さん、シフトの提出が完了しました。")
    st.info("※修正が必要な場合は店長または箭内へ連絡してください。")
    st.markdown("#### 提出されたシフト内容")
    st.table(generate_styled_calendar(day_labels, shift_requests))

elif st.session_state.confirm_mode:
    st.warning("以下の内容で確定してよろしいですか？")
    st.markdown("#### シフト希望プレビュー")
    st.table(generate_styled_calendar(day_labels, shift_requests))
    
    col_btn1, col_btn2 = st.columns(2)
    with col_btn1:
        if st.button("戻って修正する", use_container_width=True):
            st.session_state.confirm_mode = False
            st.rerun()
    with col_btn2:
        if st.button("この内容で確定・提出する", type="primary", use_container_width=True):
            
            result = save_shift_data(emp_code, name, department, target_days, shift_requests)
            
            if result == "gas_error":
                st.error("【通信エラー】Googleスプレッドシートへの送信に失敗しました。時間をおいてもう一度お試しいただくか、管理者へ連絡してください。")
            else:
                st.session_state.is_submitted = True
                st.session_state.confirm_mode = False
                st.rerun()

else:
    st.markdown("#### ライブプレビュー")
    st.table(generate_styled_calendar(day_labels, shift_requests))
    st.write("") 
    if st.button("確認画面へ進む", type="primary", use_container_width=True):
        if not name: st.error("お名前が入力されていません。")
        if department == "選択してください": st.error("部門が選択されていません。")
        if not emp_code: st.error("従業員コードが入力されていません。")
        
        if name and department != '選択してください' and emp_code:
            st.session_state.confirm_mode = True
            st.rerun()
