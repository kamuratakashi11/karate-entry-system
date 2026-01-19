import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Alignment
import json
import datetime
import io
import os
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import time
import random

# 安全なインポート（環境によるエラー回避）
try:
    from openpyxl.cell import MergedCell
except ImportError:
    try:
        from openpyxl.cell.cell import MergedCell
    except ImportError:
        class MergedCell: pass

# ---------------------------------------------------------
# 1. 定数・初期設定
# ---------------------------------------------------------
KEY_FILE = 'secrets.json'
SHEET_NAME = 'tournament_db'
# ※ ADMIN_PASSWORD 定数は廃止し、configから読み込む仕様に変更

# 大会設定
DEFAULT_TOURNAMENTS = {
    "kantou": {
        "name": "関東高等学校空手道大会 埼玉県予選",
        "template": "template_kantou.xlsx",
        "type": "standard", 
        "grades": [1, 2, 3],
        "active": True
    },
    "interhigh": {
        "name": "インターハイ 埼玉県予選",
        "template": "template_interhigh.xlsx",
        "type": "standard",
        "grades": [1, 2, 3],
        "active": False
    },
    "shinjin": {
        "name": "新人大会",
        "template": "template_shinjin.xlsx",
        "type": "shinjin",
        "grades": [1, 2],
        "weights_m": "-55,-61,-68,-76,+76", 
        "weights_w": "-48,-53,-59,-66,+66", 
        "active": False
    },
    "senbatsu": {
        "name": "全国選抜 埼玉県予選",
        "template": "template_senbatsu.xlsx",
        "type": "division", 
        "grades": [1, 2],
        "active": False
    }
}

# 人数制限設定
DEFAULT_LIMITS = {
    "team_kata": {"min": 3, "max": 3},
    "team_kumite_5": {"min": 3, "max": 5},
    "team_kumite_3": {"min": 2, "max": 3},
    "ind_kata_reg": {"max": 4},
    "ind_kata_sub": {"max": 2},
    "ind_kumi_reg": {"max": 4},
    "ind_kumi_sub": {"max": 2}
}

# Excel座標設定
COORD_DEF = {
    "year": "E3", "tournament_name": "I3", "date": "M7",
    "school_name": "C8", "principal": "C9", "head_advisor": "O9",
    "advisors": [
        {"name": "B42", "d1": "C42", "d2": "F42"},
        {"name": "B43", "d1": "C43", "d2": "F43"},
        {"name": "K42", "d1": "Q42", "d2": "U42"},
        {"name": "K43", "d1": "Q43", "d2": "U43"}
    ],
    "start_row": 16, "cap": 22, "offset": 46,
    "cols": {
        "name": 2, "grade": 3, "dob": 4, "jkf_no": 19,
        "m_team_kata": 11, "m_team_kumite": 12, "m_kata": 13, "m_kumite": 14,
        "w_team_kata": 15, "w_team_kumite": 16, "w_kata": 17, "w_kumite": 18
    }
}

# ---------------------------------------------------------
# 2. Google Sheets 接続 & リトライ
# ---------------------------------------------------------
@st.cache_resource
def get_gsheet_client():
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    if os.path.exists(KEY_FILE):
        creds = ServiceAccountCredentials.from_json_keyfile_name(KEY_FILE, scope)
    else:
        try:
            key_dict = json.loads(st.secrets["gcp_key"])
            creds = ServiceAccountCredentials.from_json_keyfile_dict(key_dict, scope)
        except Exception as e:
            st.error(f"認証設定エラー: {e}"); st.stop()
    return gspread.authorize(creds)

def retry_api(func):
    def wrapper(*args, **kwargs):
        for i in range(3):
            try: return func(*args, **kwargs)
            except Exception as e:
                if i == 2: raise e
                time.sleep(1 + random.random())
    return wrapper

@retry_api
def get_worksheet_safe(tab_name):
    client = get_gsheet_client()
    try: sh = client.open(SHEET_NAME)
    except gspread.exceptions.SpreadsheetNotFound:
        st.error(f"スプレッドシート '{SHEET_NAME}' が見つかりません。"); st.stop()
    try: ws = sh.worksheet(tab_name)
    except: 
        try: ws = sh.add_worksheet(title=tab_name, rows=100, cols=20)
        except: ws = sh.worksheet(tab_name)
    return ws

# ---------------------------------------------------------
# 3. データ操作
# ---------------------------------------------------------
def load_json(tab_name, default):
    try:
        ws = get_worksheet_safe(tab_name)
        val = ws.acell('A1').value
        return json.loads(val) if val else default
    except: return default

def save_json(tab_name, data):
    ws = get_worksheet_safe(tab_name)
    ws.update_acell('A1', json.dumps(data, ensure_ascii=False))

def load_members_master():
    if "master_cache" in st.session_state: return st.session_state["master_cache"]
    cols = ["school", "name", "sex", "grade", "dob", "jkf_no", "active"]
    try:
        recs = get_worksheet_safe("members").get_all_records()
        df = pd.DataFrame(recs) if recs else pd.DataFrame(columns=cols)
    except:
        return pd.DataFrame(columns=cols)
    df['grade'] = pd.to_numeric(df['grade'], errors='coerce').fillna(0).astype(int)
    df['jkf_no'] = df['jkf_no'].astype(str)
    st.session_state["master_cache"] = df
    return df

def save_members_master(df):
    ws = get_worksheet_safe("members"); ws.clear()
    df = df.fillna("")
    df['jkf_no'] = df['jkf_no'].astype(str)
    ws.update([df.columns.tolist()] + df.astype(str).values.tolist())
    st.session_state["master_cache"] = df

def load_entries(tournament_id):
    if f"entry_cache_{tournament_id}" in st.session_state:
        return st.session_state[f"entry_cache_{tournament_id}"]
    try:
        ws = get_worksheet_safe(f"entry_{tournament_id}")
        val = ws.acell('A1').value
        data = json.loads(val) if val else {}
    except: data = {}
    st.session_state[f"entry_cache_{tournament_id}"] = data
    return data

def save_entries(tournament_id, data):
    ws = get_worksheet_safe(f"entry_{tournament_id}")
    ws.update_acell('A1', json.dumps(data, ensure_ascii=False))
    st.session_state[f"entry_cache_{tournament_id}"] = data

def load_auth(): return load_json("auth", {})
def save_auth(d): save_json("auth", d)
def load_schools(): return load_json("schools", {})
def save_schools(d): save_json("schools", d)

def load_conf():
    # v1.20.0: admin_password を設定に追加 (初期値: "1234")
    default_conf = {
        "year": "6", 
        "tournaments": DEFAULT_TOURNAMENTS, 
        "limits": DEFAULT_LIMITS,
        "admin_password": "1234" 
    }
    data = load_json("config", default_conf)
    if "limits" not in data: data["limits"] = DEFAULT_LIMITS
    if "tournaments" not in data: data["tournaments"] = DEFAULT_TOURNAMENTS
    if "year" not in data: data["year"] = "6"
    if "admin_password" not in data: data["admin_password"] = "1234"
    return data

def save_conf(d): save_json("config", d)

# ---------------------------------------------------------
# 4. ロジック (バックアップ・復元・更新・ソート)
# ---------------------------------------------------------
def create_backup():
    df = load_members_master()
    ws_bk_mem = get_worksheet_safe("members_backup")
    ws_bk_mem.clear()
    df = df.fillna("")
    df['jkf_no'] = df['jkf_no'].astype(str)
    ws_bk_mem.update([df.columns.tolist()] + df.astype(str).values.tolist())
    
    conf = load_conf()
    ws_bk_conf = get_worksheet_safe("config_backup")
    ws_bk_conf.update_acell('A1', json.dumps(conf, ensure_ascii=False))

def restore_from_backup():
    try:
        ws_bk_mem = get_worksheet_safe("members_backup")
        recs = ws_bk_mem.get_all_records()
        df = pd.DataFrame(recs) if recs else pd.DataFrame()
        if not df.empty:
            df['grade'] = pd.to_numeric(df['grade'], errors='coerce').fillna(0).astype(int)
            save_members_master(df)
    except: return "名簿の復元に失敗しました"

    try:
        ws_bk_conf = get_worksheet_safe("config_backup")
        val = ws_bk_conf.acell('A1').value
        if val:
            conf = json.loads(val)
            save_conf(conf)
    except: return "設定の復元に失敗しました"
    
    return "✅ バックアップから復元しました"

def perform_year_rollover():
    create_backup()
    if "master_cache" in st.session_state: del st.session_state["master_cache"]
    df = load_members_master()
    if not df.empty:
        df['grade'] = df['grade'] + 1
        df = df[df['grade'] <= 3]
        save_members_master(df)
    conf = load_conf()
    for tid in conf["tournaments"].keys():
        save_entries(tid, {})
    try:
        conf["year"] = str(int(conf["year"]) + 1)
        save_conf(conf)
    except: pass
    return "✅ 新年度更新完了（直前の状態をバックアップしました）"

def get_merged_data(school_name, tournament_id):
    master = load_members_master()
    if master.empty: return pd.DataFrame()
    my_members = master[master['school'] == school_name].copy()
    
    if f"entry_cache_{tournament_id}" in st.session_state:
        entries = st.session_state[f"entry_cache_{tournament_id}"]
    else:
        entries = load_entries(tournament_id)

    def get_ent(row, key):
        uid = f"{row['school']}_{row['name']}"
        return entries.get(uid, {}).get(key, None)
    
    cols_to_add = ["team_kata_chk", "team_kata_role", "team_kumi_chk", "team_kumi_role",
                   "kata_chk", "kata_val", "kata_rank", "kumi_chk", "kumi_val", "kumi_rank"]
    for c in cols_to_add:
        my_members[f"last_{c}"] = my_members.apply(lambda r: get_ent(r, c), axis=1)
    return my_members

def validate_counts(members_df, entries_data, limits, t_type, school_meta):
    errs = []
    for sex in ["男子", "女子"]:
        sex_df = members_df[members_df['sex'] == sex]
        cnt_tk = 0; cnt_tku = 0
        cnt_ind_k_reg = 0; cnt_ind_k_sub = 0
        cnt_ind_ku_reg = 0; cnt_ind_ku_sub = 0
        
        for _, r in sex_df.iterrows():
            uid = f"{r['school']}_{r['name']}"
            ent = entries_data.get(uid, {})
            
            if ent.get("team_kata_chk") and ent.get("team_kata_role") == "正選手": cnt_tk += 1
            if ent.get("team_kumi_chk") and ent.get("team_kumi_role") == "正選手": cnt_tku += 1
            
            if ent.get("kata_chk"):
                if ent.get("kata_val") == "補欠": cnt_ind_k_sub += 1
                else: cnt_ind_k_reg += 1
            if ent.get("kumi_chk"):
                val = ent.get("kumi_val", "")
                if val == "補欠": cnt_ind_ku_sub += 1
                elif val and val != "出場しない": cnt_ind_ku_reg += 1

        if cnt_tk > 0:
            mn, mx = limits["team_kata"]["min"], limits["team_kata"]["max"]
            if not (mn <= cnt_tk <= mx):
                errs.append(f"❌ {sex}団体形: 正選手は {mn}～{mx}名で登録してください。(現在{cnt_tk}名)")

        if cnt_tku > 0:
            mode = "5"
            if t_type == "shinjin":
                mode_key = "m_kumite_mode" if sex == "男子" else "w_kumite_mode"
                mode = school_meta.get(mode_key, "none")
            
            if mode == "5":
                mn, mx = limits["team_kumite_5"]["min"], limits["team_kumite_5"]["max"]
                if not (mn <= cnt_tku <= mx):
                    errs.append(f"❌ {sex}団体組手(5人制): 正選手は {mn}～{mx}名で登録してください。(現在{cnt_tku}名)")
            elif mode == "3":
                mn, mx = limits["team_kumite_3"]["min"], limits["team_kumite_3"]["max"]
                if not (mn <= cnt_tku <= mx):
                    errs.append(f"❌ {sex}団体組手(3人制): 正選手は {mn}～{mx}名で登録してください。(現在{cnt_tku}名)")
        
        if cnt_ind_k_reg > limits["ind_kata_reg"]["max"]: errs.append(f"❌ {sex}個人形(正): 上限 {limits['ind_kata_reg']['max']}名を超えています。")
        if cnt_ind_k_sub > limits["ind_kata_sub"]["max"]: errs.append(f"❌ {sex}個人形(補): 上限 {limits['ind_kata_sub']['max']}名を超えています。")
        if cnt_ind_ku_reg > limits["ind_kumi_reg"]["max"]: errs.append(f"❌ {sex}個人組手(正): 上限 {limits['ind_kumi_reg']['max']}名を超えています。")
        if cnt_ind_ku_sub > limits["ind_kumi_sub"]["max"]: errs.append(f"❌ {sex}個人組手(補): 上限 {limits['ind_kumi_sub']['max']}名を超えています。")

    return errs

# ---------------------------------------------------------
# 5. Excel生成
# ---------------------------------------------------------
def safe_write(ws, target, value, align_center=False):
    if value is None: return
    if isinstance(target, str): cell = ws[target]
    else: cell = ws.cell(row=target[0], column=target[1])
    
    if isinstance(cell, MergedCell):
        for r in ws.merged_cells.ranges:
            if cell.coordinate in r:
                cell = ws.cell(row=r.min_row, column=r.min_col); break
    
    val_str = str(value)
    if val_str.endswith("年") and val_str[:-1].isdigit(): val_str = val_str.replace("年", "")
    cell.value = val_str
    if align_center: cell.alignment = Alignment(horizontal='center', vertical='center')

def generate_excel(school_name, school_data, members_df, t_id, t_conf):
    coords = COORD_DEF
    template_file = t_conf.get("template", "template.xlsx")
    try: wb = openpyxl.load_workbook(template_file); ws = wb.active
    except: return None, f"{template_file} が見つかりません。"
    
    conf = load_conf()
    safe_write(ws, coords["year"], conf.get("year", ""))
    safe_write(ws, coords["tournament_name"], t_conf.get("name", ""))
    safe_write(ws, coords["date"], f"令和{datetime.date.today().year-2018}年{datetime.date.today().month}月{datetime.date.today().day}日")
    safe_write(ws, coords["school_name"], school_name)
    safe_write(ws, coords["principal"], school_data.get("principal", ""))
    
    advs = school_data.get("advisors", [])
    head = advs[0]["name"] if advs else ""
    safe_write(ws, coords["head_advisor"], head)
    
    for i, a in enumerate(advs[:4]):
        c = coords["advisors"][i]
        safe_write(ws, c["name"], a["name"])
        safe_write(ws, c["d1"], "○" if a.get("d1") else "×", True)
        safe_write(ws, c["d2"], "○" if a.get("d2") else "×", True)
    
    cols = coords["cols"]
    # Excel出力用ソート (男子->女子、学年降順)
    members_df['sex_rank'] = members_df['sex'].map({'男子': 0, '女子': 1})
    members_df['grade_rank'] = members_df['grade'].map({3: 0, 2: 1, 1: 2})
    entries = members_df[
        (members_df['last_team_kata_chk']==True) | (members_df['last_team_kumi_chk']==True) |
        (members_df['last_kata_chk']==True) | (members_df['last_kumi_chk']==True)
    ].sort_values(by=['sex_rank', 'grade_rank', 'name'])

    for i, (_, row) in enumerate(entries.iterrows()):
        r = coords["start_row"] + (i // coords["cap"] * coords["offset"]) + (i % coords["cap"])
        safe_write(ws, (r, cols["name"]), row["name"])
        safe_write(ws, (r, cols["grade"]), row["grade"])
        safe_write(ws, (r, cols["dob"]), row["dob"])
        safe_write(ws, (r, cols["jkf_no"]), row["jkf_no"])
        
        sex = row["sex"]
        tk_col = cols["m_team_kata"] if sex=="男子" else cols["w_team_kata"]
        tku_col = cols["m_team_kumite"] if sex=="男子" else cols["w_team_kumite"]
        if row.get("last_team_kata_chk"):
            safe_write(ws, (r, tk_col), "補" if row.get("last_team_kata_role")=="補欠" else "○", True)
        if row.get("last_team_kumi_chk"):
            safe_write(ws, (r, tku_col), "補" if row.get("last_team_kumi_role")=="補欠" else "○", True)
            
        k_col = cols["m_kata"] if sex=="男子" else cols["w_kata"]
        ku_col = cols["m_kumite"] if sex=="男子" else cols["w_kumite"]
        
        if row.get("last_kata_chk"):
            val = row.get("last_kata_val")
            rank = row.get("last_kata_rank", "")
            if val == "補欠": txt = "補"
            elif t_conf["type"] == "standard": txt = f"○{rank}" if val=="一般" else f"シ{rank}"
            else: txt = "○"
            safe_write(ws, (r, k_col), txt, True)

        if row.get("last_kumi_chk"):
            val = row.get("last_kumi_val")
            rank = row.get("last_kumi_rank", "")
            if val == "補欠": txt = "補"
            elif t_conf["type"] == "standard": txt = f"○{rank}" if val=="一般" else f"シ{rank}"
            elif t_conf["type"] == "weight": txt = str(val)
            elif t_conf["type"] == "division": txt = str(val)
            else: txt = "○"
            safe_write(ws, (r, ku_col), txt, True)
    
    fname = f"申込書_{school_name}.xlsx"
    wb.save(fname)
    return fname, "作成成功"

# ---------------------------------------------------------
# 6. UI: 学校ページ
# ---------------------------------------------------------
def school_page(s_name):
    col_h1, col_h2 = st.columns([3, 1])
    with col_h1: st.markdown(f"### 🏫 {s_name} 様")
    with col_h2:
        if st.button("🚪 ログアウト", type="secondary", use_container_width=True):
            st.query_params.clear()
            st.session_state.clear()
            st.rerun()
    st.divider()

    conf = load_conf()
    active_tid = next((k for k, v in conf["tournaments"].items() if v["active"]), None)
    t_conf = conf["tournaments"].get(active_tid, {}) if active_tid else {}
    
    if not active_tid: st.error("現在受付中の大会はありません。"); return
    
    disp_year = conf.get("year", "〇")
    st.markdown(f"## 🥋 **令和{disp_year}年度 {t_conf['name']}** <small>エントリー画面</small>", unsafe_allow_html=True)
    
    if st.button("🔄 データを最新にする (うまく表示されない場合)"):
        if "master_cache" in st.session_state: del st.session_state["master_cache"]
        if f"entry_cache_{active_tid}" in st.session_state: del st.session_state[f"entry_cache_{active_tid}"]
        st.success("最新データを読み込みました")
        time.sleep(1); st.rerun()

    if "schools_data" not in st.session_state: st.session_state.schools_data = load_schools()
    s_data = st.session_state.schools_data.get(s_name, {"principal":"", "advisors":[]})
    
    if "current_view" not in st.session_state: st.session_state["current_view"] = "① 大会エントリー"

    # UI改善: 並び順変更
    menu = ["① 大会エントリー", "② 部員名簿", "③ 顧問登録"]
    try: idx = menu.index(st.session_state["current_view"])
    except: idx = 0
    selected_view = st.radio("メニュー選択", menu, index=idx, horizontal=True, label_visibility="collapsed")
    st.session_state["current_view"] = selected_view
    st.markdown("---")

    # --- ① 大会エントリー ---
    if selected_view == "① 大会エントリー":
        target_grades = [int(g) for g in t_conf['grades']]
        st.markdown(f"**出場対象学年:** {target_grades} 年生")
        merged = get_merged_data(s_name, active_tid)
        
        # ソートロジック強化: 男子(3->2->1) -> 女子(3->2->1)
        merged['sex_rank'] = merged['sex'].map({'男子': 0, '女子': 1})
        merged['grade_rank'] = merged['grade'].map({3: 0, 2: 1, 1: 2})
        
        valid_members = merged[merged['grade'].isin(target_grades)].sort_values(by=['sex_rank', 'grade_rank', 'name']).copy()
        
        if valid_members.empty: st.warning("部員名簿が空です。"); return
        
        entries_update = load_entries(active_tid)
        
        meta_key = f"_meta_{s_name}"
        school_meta = entries_update.get(meta_key, {"m_kumite_mode": "none", "w_kumite_mode": "none"})
        m_mode = "5"; w_mode = "5"
        if t_conf["type"] == "shinjin":
            with st.expander("団体組手の設定 (新人戦)", expanded=True):
                c_m, c_w = st.columns(2)
                curr_m = school_meta.get("m_kumite_mode", "none")
                idx_m = ["none", "5", "3"].index(curr_m) if curr_m in ["none", "5", "3"] else 0
                new_m = c_m.radio("男子 団体組手", ["出場しない", "5人制", "3人制"], index=idx_m, horizontal=True)
                m_mode = "none" if new_m == "出場しない" else ("5" if new_m == "5人制" else "3")
                curr_w = school_meta.get("w_kumite_mode", "none")
                idx_w = ["none", "5", "3"].index(curr_w) if curr_w in ["none", "5", "3"] else 0
                new_w = c_w.radio("女子 団体組手", ["出場しない", "5人制", "3人制"], index=idx_w, horizontal=True)
                w_mode = "none" if new_w == "出場しない" else ("5" if new_w == "5人制" else "3")
                school_meta["m_kumite_mode"] = m_mode; school_meta["w_kumite_mode"] = w_mode
                entries_update[meta_key] = school_meta

        with st.form("entry_form_unified"):
            cols = st.columns([1.5, 2.3, 2.3, 2.7, 3.2])
            cols[0].markdown("**氏名**")
            cols[1].markdown("**団体形** (なし/正/補)")
            kumi_label = f"**団体組手({m_mode if m_mode==w_mode else '選択'})**"
            cols[2].markdown(f"{kumi_label} (なし/正/補)")
            cols[3].markdown("**個人形** (なし/正/補)[順位]")
            cols[4].markdown("**個人組手** (階級)[順位]")

            form_buffer = {}

            for i, r in valid_members.iterrows():
                uid = f"{r['school']}_{r['name']}"
                name_style = 'background-color:#e8f5e9; color:#1b5e20; padding:2px 6px; border-radius:4px; font-weight:bold;' if r['sex'] == "男子" else 'background-color:#ffebee; color:#b71c1c; padding:2px 6px; border-radius:4px; font-weight:bold;'
                
                def_tk = "なし"
                if r.get("last_team_kata_chk"): def_tk = "正" if r.get("last_team_kata_role") == "正選手" else "補"
                
                def_tku = "なし"
                if r.get("last_team_kumi_chk"): def_tku = "正" if r.get("last_team_kumi_role") == "正選手" else "補"
                
                def_k = "なし"
                if r.get("last_kata_chk"):
                    val = r.get("last_kata_val")
                    def_k = "補" if val == "補欠" else "正"

                c = st.columns([1.5, 2.3, 2.3, 2.7, 3.2])
                c[0].markdown(f'<span style="{name_style}">{r["grade"]}年 {r["name"]}</span>', unsafe_allow_html=True)
                
                opts_tk = ["なし", "正", "補"]
                idx_tk = opts_tk.index(def_tk) if def_tk in opts_tk else 0
                val_tk = c[1].radio(f"tk_{uid}", opts_tk, index=idx_tk, horizontal=True, key=f"rd_tk_{uid}", label_visibility="collapsed")
                
                mode = m_mode if r['sex']=="男子" else w_mode
                if mode != "none":
                    opts_tku = ["なし", "正", "補"]
                    idx_tku = opts_tku.index(def_tku) if def_tku in opts_tku else 0
                    val_tku = c[2].radio(f"tku_{uid}", opts_tku, index=idx_tku, horizontal=True, key=f"rd_tku_{uid}", label_visibility="collapsed")
                else:
                    val_tku = "なし"; c[2].caption("-")

                if t_conf["type"] != "division":
                    opts_k = ["なし", "正", "補"]
                    idx_k = opts_k.index(def_k) if def_k in opts_k else 0
                    ck1, ck2 = c[3].columns([1.5, 1])
                    val_k = ck1.radio(f"k_{uid}", opts_k, index=idx_k, horizontal=True, key=f"rd_k_{uid}", label_visibility="collapsed")
                    rank_k = ck2.text_input("順位", r.get("last_kata_rank",""), key=f"rk_k_{uid}", label_visibility="collapsed", placeholder="順位")
                else:
                    val_k = "なし"; rank_k = ""; c[3].caption("-")
                
                c4a, c4b = c[4].columns([1.8, 1])
                w_key = "weights_m" if r['sex'] == "男子" else "weights_w"
                w_str = t_conf.get(w_key, "")
                w_list = ["出場しない"] + [f"{w.strip()}kg級" for w in w_str.split(",")] + ["補欠"]
                if t_conf["type"] == "standard":
                    w_list = ["出場しない", "一般", "シード", "補欠"]
                
                def_val = r.get("last_kumi_val", "出場しない")
                if def_val not in w_list:
                     if def_val in ["一般","シード","補欠"] and t_conf["type"] == "weight": def_val = "出場しない" 
                     elif "kg" in def_val and t_conf["type"] == "standard": def_val = "出場しない"
                     elif t_conf["type"] == "weight" and def_val!="出場しない": def_val = f"{def_val}kg級"

                try: idx = w_list.index(def_val)
                except: idx = 0
                ku_val = c4a.selectbox("階級", w_list, index=idx, key=f"sel_ku_{uid}", label_visibility="collapsed")
                rank_ku = c4b.text_input("順位", r.get("last_kumi_rank",""), key=f"rk_ku_{uid}", label_visibility="collapsed", placeholder="順位")

                form_buffer[uid] = {
                    "val_tk": val_tk, "val_tku": val_tku, 
                    "val_k": val_k, "rank_k": rank_k,
                    "ku_val": ku_val, "rank_ku": rank_ku,
                }

            if st.form_submit_button("✅ エントリーを保存 (全員分)"):
                has_error = False
                processed = {}
                temp_processed = {}
                for uid, raw in form_buffer.items():
                    tk_chk = (raw["val_tk"] != "なし")
                    tk_role = "正選手" if raw["val_tk"] == "正" else ("補欠" if raw["val_tk"] == "補" else "")
                    tku_chk = (raw["val_tku"] != "なし")
                    tku_role = "正選手" if raw["val_tku"] == "正" else ("補欠" if raw["val_tku"] == "補" else "")
                    k_chk = (raw["val_k"] != "なし")
                    k_role = "一般" if raw["val_k"] == "正" else ("補欠" if raw["val_k"] == "補" else "")
                    k_rank = raw["rank_k"]
                    ku_chk = (raw["ku_val"] != "出場しない")
                    ku_role = raw["ku_val"]
                    ku_rank = raw["rank_ku"]

                    name = uid.split('_')[1]
                    if k_chk and k_role == "一般":
                        if not k_rank: st.error(f"❌ {name}さん: 個人形の順位が入力されていません。"); has_error = True
                    elif not k_chk or k_role == "補欠": k_rank = ""

                    if ku_chk:
                        is_reg = (t_conf["type"] == "weight" and ku_role != "補欠") or \
                                 (t_conf["type"] == "standard" and ku_role in ["一般", "シード"])
                        if is_reg and not ku_rank: st.error(f"❌ {name}さん: 個人組手の順位が入力されていません。"); has_error = True
                    
                    if not ku_chk or ku_role == "補欠": ku_rank = ""

                    temp_processed[uid] = {
                        "team_kata_chk": tk_chk, "team_kata_role": tk_role,
                        "team_kumi_chk": tku_chk, "team_kumi_role": tku_role,
                        "kata_chk": k_chk, "kata_val": k_role, "kata_rank": k_rank,
                        "kumi_chk": ku_chk, "kumi_val": ku_role, "kumi_rank": ku_rank
                    }
                
                current_entries = load_entries(active_tid)
                current_entries.update(temp_processed)
                errs = validate_counts(valid_members, current_entries, conf["limits"], t_conf["type"], {"m_kumite_mode":m_mode, "w_kumite_mode":w_mode})
                if errs:
                    has_error = True
                    for e in errs: st.error(e)
                    st.error("⚠️ 保存できませんでした。人数超過などを修正してください。")

                if not has_error:
                    save_entries(active_tid, current_entries)
                    st.success("✅ データを保存しました！")
                    time.sleep(1); st.rerun()

        st.markdown("---")
        if st.button("📥 Excel申込書を作成する", type="primary"):
             latest_entries = load_entries(active_tid)
             final_merged = get_merged_data(s_name, active_tid)
             fp, msg = generate_excel(s_name, s_data, final_merged, active_tid, t_conf)
             if fp:
                 with open(fp, "rb") as f:
                     st.download_button("📥 ダウンロード開始", f, fp, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
             else: st.error(msg)

    # --- ② 部員名簿 ---
    elif selected_view == "② 部員名簿":
        st.info("💡 ここは「全大会共通」の名簿です。")
        with st.expander("👤 新しい部員を追加する", expanded=False):
            with st.form("add_member"):
                c = st.columns(3)
                nn = c[0].text_input("氏名")
                ns = c[1].selectbox("性別", ["男子", "女子"])
                ng = c[2].selectbox("学年", [1, 2, 3])
                c2 = st.columns(2)
                nd = c2[0].text_input("生年月日")
                nj = c2[1].text_input("JKF番号")
                if st.form_submit_button("追加"):
                    if nn:
                        if "master_cache" in st.session_state: del st.session_state["master_cache"]
                        master = load_members_master()
                        new_row = pd.DataFrame([{"school":s_name, "name":nn, "sex":ns, "grade":ng, "dob":nd, "jkf_no":nj, "active":True}])
                        save_members_master(pd.concat([master, new_row], ignore_index=True))
                        st.success(f"{nn} さんを追加しました"); st.rerun()
        master = load_members_master()
        my_m = master[master['school']==s_name].reset_index()
        if my_m.empty: st.warning("部員が登録されていません。")
        else:
            st.markdown("##### 登録済み部員リスト")
            for i, r in my_m.iterrows():
                c = st.columns([2, 1, 1, 1])
                c[0].write(r['name'])
                c[1].write(r['sex'])
                c[2].write(f"{r['grade']}年")
                if c[3].button("削除", key=f"mdel_{r['index']}"):
                    if "master_cache" in st.session_state: del st.session_state["master_cache"]
                    save_members_master(master.drop(r['index'])); st.rerun()

    # --- ③ 顧問登録 ---
    elif selected_view == "③ 顧問登録":
        c_p = st.columns([1, 2])
        np = c_p[0].text_input("校長名", s_data.get("principal", ""))
        st.markdown("#### 顧問リスト")
        advs = s_data.get("advisors", [])
        for i, a in enumerate(advs):
            with st.container():
                c = st.columns([0.8, 2, 1.5, 0.5, 0.5, 0.7])
                if i == 0: c[0].info("筆頭顧問")
                else: c[0].caption("顧問")
                a["name"] = c[1].text_input("氏名", a["name"], key=f"n{i}", label_visibility="collapsed", placeholder="氏名")
                a["role"] = c[2].selectbox("役割", ["審判","競技記録","係員"], index=["審判","競技記録","係員"].index(a.get("role","審判")), key=f"r{i}", label_visibility="collapsed")
                a["d1"] = c[3].checkbox("1日目", a.get("d1"), key=f"d1{i}")
                a["d2"] = c[4].checkbox("2日目", a.get("d2"), key=f"d2{i}")
                if c[5].button("削除", key=f"del_{i}"):
                    advs.pop(i)
                    s_data["advisors"] = advs
                    save_schools(st.session_state.schools_data); st.rerun()
        if st.button("＋ 顧問を追加"):
            advs.append({"name":"", "role":"審判", "d1":True, "d2":True})
            s_data["advisors"] = advs
            save_schools(st.session_state.schools_data); st.rerun()
        if st.button("顧問情報を保存", type="primary"):
            s_data["principal"] = np; s_data["advisors"] = advs
            st.session_state.schools_data[s_name] = s_data
            save_schools(st.session_state.schools_data); st.success("保存しました")

# ---------------------------------------------------------
# 7. UI: 管理者ページ
# ---------------------------------------------------------
def admin_page():
    st.title("🔧 管理者画面")
    
    conf = load_conf()
    current_admin_pw = conf.get("admin_password", "1234")
    
    input_pw = st.text_input("Admin Password", type="password")
    if input_pw != current_admin_pw:
        return # パスワード不一致なら何も表示しない

    auth = load_auth()
    t1, t2, t3, t4 = st.tabs(["🏆 大会設定", "📥 データ出力", "🏫 アカウント", "📅 年次処理"])
    
    with t1:
        st.subheader("基本設定")
        new_year = st.text_input("現在の年度", conf.get("year", "6"))
        st.subheader("大会切り替え")
        t_opts = list(conf["tournaments"].keys())
        active_now = next((k for k, v in conf["tournaments"].items() if v["active"]), None)
        new_active = st.radio("受付中の大会", t_opts, index=t_opts.index(active_now) if active_now else 0, format_func=lambda x: conf["tournaments"][x]["name"])
        if st.button("設定を保存 & 大会切替"):
            conf["year"] = new_year
            if new_active != active_now:
                for k in conf["tournaments"]: conf["tournaments"][k]["active"] = (k == new_active)
            save_conf(conf); st.success("保存しました"); st.rerun()
        st.divider()
        with st.expander("参加人数制限の設定", expanded=True):
            lm = conf["limits"]
            c1, c2 = st.columns(2)
            lm["team_kata"]["min"] = c1.number_input("団体形 下限", 0, 10, lm["team_kata"]["min"])
            lm["team_kata"]["max"] = c2.number_input("団体形 上限", 0, 10, lm["team_kata"]["max"])
            c1, c2 = st.columns(2)
            lm["team_kumite_5"]["min"] = c1.number_input("団体組手(5人) 下限", 0, 10, lm["team_kumite_5"]["min"])
            lm["team_kumite_5"]["max"] = c2.number_input("団体組手(5人) 上限", 0, 10, lm["team_kumite_5"]["max"])
            c1, c2 = st.columns(2)
            lm["team_kumite_3"]["min"] = c1.number_input("団体組手(3人) 下限", 0, 10, lm["team_kumite_3"]["min"])
            lm["team_kumite_3"]["max"] = c2.number_input("団体組手(3人) 上限", 0, 10, lm["team_kumite_3"]["max"])
            st.caption("個人戦 (上限のみ)")
            c1, c2 = st.columns(2)
            lm["ind_kata_reg"]["max"] = c1.number_input("個人形(正) 上限", 0, 10, lm["ind_kata_reg"]["max"])
            lm["ind_kata_sub"]["max"] = c2.number_input("個人形(補) 上限", 0, 10, lm["ind_kata_sub"]["max"])
            c1, c2 = st.columns(2)
            lm["ind_kumi_reg"]["max"] = c1.number_input("個人組手(正) 上限", 0, 10, lm["ind_kumi_reg"]["max"])
            lm["ind_kumi_sub"]["max"] = c2.number_input("個人組手(補) 上限", 0, 10, lm["ind_kumi_sub"]["max"])
            if st.button("人数制限を保存"):
                conf["limits"] = lm; save_conf(conf); st.success("保存しました")
        st.caption("新人戦 階級設定 (男女別)")
        t_data = conf["tournaments"]["shinjin"]
        with st.form("edit_t"):
            wm_in = st.text_area("男子階級リスト", t_data.get("weights_m", ""))
            ww_in = st.text_area("女子階級リスト", t_data.get("weights_w", ""))
            if st.form_submit_button("階級を保存"):
                conf["tournaments"]["shinjin"]["weights_m"] = wm_in
                conf["tournaments"]["shinjin"]["weights_w"] = ww_in
                save_conf(conf); st.success("保存しました")

    with t2:
        st.subheader("全データダウンロード")
        tid = next((k for k, v in conf["tournaments"].items() if v["active"]), "kantou")
        master = load_members_master(); entries = load_entries(tid)
        full_data = []
        for _, m in master.iterrows():
            uid = f"{m['school']}_{m['name']}"
            ent = entries.get(uid, {})
            if ent and (ent.get("kata_chk") or ent.get("kumi_chk") or ent.get("team_kata_chk") or ent.get("team_kumi_chk")):
                row = m.to_dict(); row.update(ent)
                row["school_no"] = auth.get(m['school'], {}).get("school_no", 999)
                full_data.append(row)
        df_out = pd.DataFrame(full_data)
        if not df_out.empty:
            df_out = df_out.sort_values(by=["school_no", "grade"])
            csv = df_out.to_csv(index=False).encode('utf-8_sig')
            st.download_button("エントリー一覧 (CSV)", csv, "entries.csv")
        else: st.warning("エントリーデータがありません")

    with t3:
        st.subheader("学校番号 & パスワード管理")
        
        # 1. 管理者PW変更
        with st.expander("🔑 管理者パスワード変更"):
            new_admin_pw = st.text_input("新しい管理者パスワード", type="password")
            if st.button("管理者パスワードを変更"):
                if len(new_admin_pw) < 6: st.error("6文字以上にしてください")
                else:
                    conf["admin_password"] = new_admin_pw
                    save_conf(conf); st.success("変更しました。次回から新しいパスワードを使用してください。")

        # 2. 学校アカウント管理
        st.markdown("---")
        s_list = [{"学校名":k, "No":v.get("school_no",999), "Password": v.get("password","")} for k,v in auth.items()]
        edf = st.data_editor(pd.DataFrame(s_list), key="sed", num_rows="fixed")
        if st.button("学校情報を保存 (PW変更含む)"):
            has_pw_error = False
            for i, r in edf.iterrows():
                if r["学校名"] in auth:
                    if len(str(r["Password"])) < 6:
                        st.error(f"❌ {r['学校名']} のパスワードが短すぎます (6文字以上)"); has_pw_error = True
                    else:
                        auth[r["学校名"]]["school_no"] = int(r["No"])
                        auth[r["学校名"]]["password"] = str(r["Password"])
            
            if not has_pw_error:
                save_auth(auth); st.success("保存しました")
            
    with t4:
        st.subheader("🌸 年度更新処理")
        st.warning("【注意】実行すると学年+1、3年削除、全エントリーリセットされます。")
        col_act1, col_act2 = st.columns(2)
        if col_act1.button("新年度を開始する (実行確認)"):
            res = perform_year_rollover(); st.success(res)
        
        st.markdown("---")
        st.subheader("⏪ 復元 (Undo)")
        st.info("間違えて年度更新してしまった場合、ここから元に戻せます。")
        if st.button("バックアップから復元する"):
            res = restore_from_backup(); st.warning(res)

# ---------------------------------------------------------
# 8. Main
# ---------------------------------------------------------
def main():
    st.set_page_config(page_title="大会エントリー", layout="wide")
    qp = st.query_params
    if "school" in qp: st.session_state["logged_in_school"] = qp["school"]
    if "logged_in_school" in st.session_state:
        st.query_params["school"] = st.session_state["logged_in_school"]
        school_page(st.session_state["logged_in_school"]); return

    st.title("🔐 エントリーシステム"); auth = load_auth()
    t1, t2, t3 = st.tabs(["ログイン", "新規登録", "管理者"])
    with t1:
        s = st.selectbox("学校名", list(auth.keys()))
        pw = st.text_input("パスワード", type="password")
        if st.button("ログイン"):
            if s in auth and auth[s]["password"] == pw:
                st.session_state["logged_in_school"] = s; st.rerun()
            else: st.error("パスワードが違います")
    with t2:
        n = st.text_input("学校名 (新規)"); p = st.text_input("校長名"); new_pw = st.text_input("パスワード (設定)", type="password")
        if st.button("登録"):
            if n and new_pw:
                if len(new_pw) < 6: st.error("パスワードは6文字以上にしてください")
                else:
                    auth[n] = {"password": new_pw, "principal": p, "school_no": 999}
                    save_auth(auth); st.success("登録しました"); st.rerun()
    with t3: admin_page()

if __name__ == "__main__": main()