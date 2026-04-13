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
import base64
import requests

# 安全なインポート
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
V2_PREFIX = "v2_" 
# ★ステップ1で取得したURLをここに貼り付けてください
GAS_WEBAPP_URL = "https://script.google.com/macros/s/AKfycbwTegYveIaIKagvcsJBcLlxbjVx7siHoeUmh_3YrRSu9uOpvl6Uo8X3NifGinnzuxSA/exec"

MEMBERS_COLS = ["school_id", "name", "sex", "grade", "dob", "jkf_no", "display_order", "active"]

def to_half_width(text):
    if not text: return ""
    return str(text).translate(str.maketrans('０１２３４５６７８９', '0123456789')).strip()

def to_safe_int(val):
    try:
        s = to_half_width(str(val))
        return int(s)
    except: return 999

def generate_school_id():
    return f"sch_{datetime.datetime.now().strftime('%Y%m%d%H%M%S')}"

# 大会設定
DEFAULT_TOURNAMENTS = {
    "kantou": {
        "name": "関東高等学校空手道大会 埼玉県予選",
        "template": "template_kantou.xlsx",
        "type": "standard", "grades": [1, 2, 3], "active": True
    },
    "interhigh": {
        "name": "インターハイ 埼玉県予選",
        "template": "template_interhigh.xlsx",
        "type": "standard", "grades": [1, 2, 3], "active": False
    },
    "shinjin": {
        "name": "新人大会",
        "template": "template_shinjin.xlsx",
        "type": "shinjin", "grades": [1, 2],
        "weights_m": "-55,-61,-68,-76,+76", "weights_w": "-48,-53,-59,-66,+66", "active": False
    },
    "senbatsu": {
        "name": "全国選抜 埼玉県予選",
        "template": "template_senbatsu.xlsx",
        "type": "division", "grades": [1, 2], "active": False
    }
}

DEFAULT_LIMITS = {
    "team_kata": {"min": 3, "max": 3},
    "team_kumite_5": {"min": 3, "max": 5},
    "team_kumite_3": {"min": 2, "max": 3},
    "ind_kata_reg": {"max": 4}, "ind_kata_sub": {"max": 2},
    "ind_kumi_reg": {"max": 4}, "ind_kumi_sub": {"max": 2}
}

COORD_DEF = {
    "year": "E3", "tournament_name": "I3", "date": "M7",
    "school_name": "C8", "principal": "C9", "head_advisor": "O9",
    "advisors": [
        {"name": "B42", "d1": "C42", "d2": "F42"}, {"name": "B43", "d1": "C43", "d2": "F43"},
        {"name": "K42", "d1": "Q42", "d2": "U42"}, {"name": "K43", "d1": "Q43", "d2": "U43"}
    ],
    "start_row": 16, "cap": 22, "offset": 46,
    "cols": {
        "name": 2, "grade": 3, "dob": 4, "jkf_no": 19,
        "m_team_kata": 11, "m_team_kumite": 12, "m_kata": 13, "m_kumite": 14,
        "w_team_kata": 15, "w_team_kumite": 16, "w_kata": 17, "w_kumite": 18
    }
}

# ---------------------------------------------------------
# 2. Google Sheets 接続
# ---------------------------------------------------------
@st.cache_resource
def get_gsheet_client():
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    if os.path.exists(KEY_FILE):
        creds = ServiceAccountCredentials.from_json_keyfile_name(KEY_FILE, scope)
    else:
        try:
            vals = st.secrets["gcp_key"]
            if isinstance(vals, str): key_dict = json.loads(vals)
            else: key_dict = vals
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
    target_tab = f"{V2_PREFIX}{tab_name}"
    try:
        ws = get_worksheet_safe(target_tab)
        recs = ws.get_all_values()
        if not recs: return default
        if len(recs) == 1 and len(recs[0]) >= 1:
            val = str(recs[0][0])
            if val.startswith("{") or val.startswith("["):
                parsed = json.loads(val)
                return parsed if parsed is not None else default
        result = {}
        for row in recs:
            if len(row) >= 2:
                key = row[0]; val_str = row[1]
                try: result[key] = json.loads(val_str)
                except: result[key] = val_str
        return result if result else default
    except: return default

def save_json(tab_name, data):
    target_tab = f"{V2_PREFIX}{tab_name}"
    ws = get_worksheet_safe(target_tab)
    if not isinstance(data, dict):
        ws.clear(); ws.update_acell('A1', json.dumps(data, ensure_ascii=False)); return
    rows = [[str(k), json.dumps(v, ensure_ascii=False)] for k, v in data.items()]
    ws.clear()
    if rows: ws.update(rows)

def load_members_master(force_reload=False):
    if not force_reload and "v2_master_cache" in st.session_state:
        return st.session_state["v2_master_cache"]
    try:
        ws = get_worksheet_safe(f"{V2_PREFIX}members")
        recs = ws.get_all_records()
        if not recs: df = pd.DataFrame(columns=MEMBERS_COLS)
        else:
             df = pd.DataFrame(recs)
             for c in MEMBERS_COLS:
                 if c not in df.columns: df[c] = ""
    except: return pd.DataFrame(columns=MEMBERS_COLS)
    df['grade'] = pd.to_numeric(df['grade'], errors='coerce').fillna(0).astype(int)
    df['jkf_no'] = df['jkf_no'].astype(str).replace('nan', '')
    df['dob'] = df['dob'].astype(str).replace('nan', '')
    df['display_order'] = df['display_order'].astype(str).replace('nan', '') 
    df = df[MEMBERS_COLS]
    st.session_state["v2_master_cache"] = df
    return df

def save_members_master(df):
    ws = get_worksheet_safe(f"{V2_PREFIX}members"); ws.clear()
    df = df.fillna("")
    df['jkf_no'] = df['jkf_no'].astype(str); df['dob'] = df['dob'].astype(str); df['display_order'] = df['display_order'].astype(str) 
    for c in MEMBERS_COLS:
        if c not in df.columns: df[c] = ""
    df_to_save = df[MEMBERS_COLS]
    ws.update([df_to_save.columns.tolist()] + df_to_save.astype(str).values.tolist())
    st.session_state["v2_master_cache"] = df_to_save

# ★追加：GAS経由でアップロードする関数
def upload_file_to_gas(uploaded_file, school_name):
    if not GAS_WEBAPP_URL or GAS_WEBAPP_URL == "ここに貼り付け":
        return False, "GASのURLが設定されていません。"
    try:
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        ext = os.path.splitext(uploaded_file.name)[1]
        file_name = f"【{school_name}】_申込書_{timestamp}{ext}"
        
        # ファイルをBase64形式に変換
        file_content = uploaded_file.getvalue()
        base64_content = base64.b64encode(file_content).decode('utf-8')
        
        payload = {
            "fileName": file_name,
            "mimeType": uploaded_file.type,
            "base64": base64_content
        }
        
        response = requests.post(GAS_WEBAPP_URL, json=payload)
        res_data = response.json()
        
        if res_data.get("status") == "success":
            return True, res_data.get("id")
        else:
            return False, res_data.get("message", "不明なエラー")
    except Exception as e:
        return False, str(e)

# ---------------------------------------------------------
# 4. ロジック (バックアップ・年次処理)
# ---------------------------------------------------------
def create_backup():
    df = load_members_master(force_reload=False)
    ws_bk_mem = get_worksheet_safe(f"{V2_PREFIX}members_backup"); ws_bk_mem.clear()
    df_bk = df.fillna("")[MEMBERS_COLS]
    ws_bk_mem.update([df_bk.columns.tolist()] + df_bk.astype(str).values.tolist())
    conf = load_conf()
    ws_bk_conf = get_worksheet_safe(f"{V2_PREFIX}config_backup")
    ws_bk_conf.update_acell('A1', json.dumps(conf, ensure_ascii=False))

def perform_year_rollover():
    create_backup()
    if "v2_master_cache" in st.session_state: del st.session_state["v2_master_cache"]
    df = load_members_master(force_reload=True)
    if df.empty: return "データがありません"
    df['grade'] = df['grade'] + 1
    graduates = df[df['grade'] > 3].copy()
    current = df[df['grade'] <= 3].copy()
    if not graduates.empty:
        auth = load_auth(); ws_grad = get_worksheet_safe(f"{V2_PREFIX}graduates")
        graduates["archived_school_name"] = graduates["school_id"].apply(lambda sid: auth.get(sid, {}).get("base_name", "不明"))
        graduates["archived_date"] = datetime.date.today().strftime("%Y-%m-%d")
        if not ws_grad.get_all_values(): ws_grad.append_row(graduates.columns.tolist())
        ws_grad.append_rows(graduates.astype(str).values.tolist())
    save_members_master(current)
    conf = load_conf()
    for tid in conf["tournaments"].keys(): save_entries(tid, {})
    try: conf["year"] = str(int(conf["year"]) + 1); save_conf(conf)
    except: pass
    return f"✅ 新年度更新完了。{len(graduates)}名の卒業生データをアーカイブしました。"

def get_merged_data(school_id, tournament_id):
    master = load_members_master(force_reload=False)
    if master.empty: return pd.DataFrame()
    my_members = master[master['school_id'] == school_id].copy()
    entries = load_entries(tournament_id, force_reload=False)
    cols_to_add = ["team_kata_chk", "team_kata_role", "team_kumi_chk", "team_kumi_role",
                   "kata_chk", "kata_val", "kata_rank", "kumi_chk", "kumi_val", "kumi_rank"]
    for c in cols_to_add:
        my_members[f"last_{c}"] = my_members.apply(lambda r: entries.get(f"{r['school_id']}_{r['name']}", {}).get(c, None), axis=1)
    return my_members

def validate_counts(members_df, entries_data, limits, t_type, school_meta, school_id):
    errs = []
    for sex in ["男子", "女子"]:
        sex_df = members_df[members_df['sex'] == sex]
        cnt_tk = 0; cnt_tku = 0; cnt_ind_k_reg = 0; cnt_ind_k_sub = 0; cnt_ind_ku_reg = 0; cnt_ind_ku_sub = 0
        for _, r in sex_df.iterrows():
            uid = f"{school_id}_{r['name']}"; ent = entries_data.get(uid, {})
            if ent.get("team_kata_chk") and ent.get("team_kata_role") == "正": cnt_tk += 1
            if ent.get("team_kumi_chk") and ent.get("team_kumi_role") == "正": cnt_tku += 1
            if ent.get("kata_chk"):
                if ent.get("kata_val") == "補": cnt_ind_k_sub += 1
                elif ent.get("kata_val") == "正": cnt_ind_k_reg += 1 
            if ent.get("kumi_chk"):
                v = ent.get("kumi_val")
                if v == "補": cnt_ind_ku_sub += 1
                elif v == "正": cnt_ind_ku_reg += 1
                elif t_type != "standard" and v and v not in ["出場しない", "なし", "シード", "補"]: cnt_ind_ku_reg += 1
        if cnt_tk > 0:
            mn, mx = limits["team_kata"]["min"], limits["team_kata"]["max"]
            if not (mn <= cnt_tk <= mx): errs.append(f"❌ {sex}団体形: 正選手は {mn}～{mx}名で登録してください。(現在{cnt_tk}名)")
        if cnt_tku > 0:
            mode = school_meta.get("m_kumite_mode" if sex == "男子" else "w_kumite_mode", "none") if t_type == "shinjin" else "5"
            l_key = "team_kumite_3" if mode == "3" else "team_kumite_5"; mn, mx = limits[l_key]["min"], limits[l_key]["max"]
            if not (mn <= cnt_tku <= mx): errs.append(f"❌ {sex}団体組手({mode}人制): 正選手は {mn}～{mx}名で登録してください。(現在{cnt_tku}名)")
        if cnt_ind_k_reg > limits["ind_kata_reg"]["max"]: errs.append(f"❌ {sex}個人形(正): 上限超え")
        if cnt_ind_ku_reg > limits["ind_kumi_reg"]["max"]: errs.append(f"❌ {sex}個人組手(正): 上限超え")
    return errs

# ---------------------------------------------------------
# 5. Excel生成
# ---------------------------------------------------------
def safe_write(ws, target, value, align_center=False):
    if value is None: value = ""
    cell = ws[target] if isinstance(target, str) else ws.cell(row=target[0], column=target[1])
    if isinstance(cell, MergedCell):
        for r in ws.merged_cells.ranges:
            if cell.coordinate in r: cell = ws.cell(row=r.min_row, column=r.min_col); break
    val_str = str(value)
    if val_str.endswith("年") and val_str[:-1].isdigit(): val_str = val_str.replace("年", "")
    cell.value = val_str
    if align_center: cell.alignment = Alignment(horizontal='center', vertical='center')

def generate_excel(school_id, school_data, members_df, t_id, t_conf):
    coords = COORD_DEF; template_file = t_conf.get("template", "template.xlsx")
    try: wb = openpyxl.load_workbook(template_file); ws = wb.active
    except: return None, f"{template_file} が見つかりません。"
    conf = load_conf(); safe_write(ws, coords["year"], conf.get("year", "")); safe_write(ws, coords["tournament_name"], t_conf.get("name", ""))
    safe_write(ws, coords["date"], f"令和{datetime.date.today().year-2018}年{datetime.date.today().month}月{datetime.date.today().day}日")
    bn = school_data.get("base_name", ""); safe_write(ws, coords["school_name"], bn); safe_write(ws, coords["principal"], school_data.get("principal", ""))
    advs = school_data.get("advisors", []); safe_write(ws, coords["head_advisor"], advs[0]["name"] if advs else "")
    for i, a in enumerate(advs[:4]):
        c = coords["advisors"][i]; safe_write(ws, c["name"], a["name"]); safe_write(ws, c["d1"], "○" if a.get("d1") else "×", True); safe_write(ws, c["d2"], "○" if a.get("d2") else "×", True)
    cols = coords["cols"]; members_df['sex_rank'] = members_df['sex'].map({'男子': 0, '女子': 1}); members_df['grade_rank'] = members_df['grade'].map({3: 0, 2: 1, 1: 2})
    def get_sort_key(row):
        try: return float(row['display_order']) if pd.notna(row['display_order']) and str(row['display_order']).strip() else 999999.0
        except: return 999999.0
    members_df['custom_order'] = members_df.apply(get_sort_key, axis=1)
    target_grades = [int(g) for g in t_conf.get('grades', [1, 2, 3])]
    entries = members_df[members_df['grade'].isin(target_grades)].sort_values(by=['custom_order', 'sex_rank', 'grade_rank', 'name'])
    for i, (_, row) in enumerate(entries.iterrows()):
        r = coords["start_row"] + (i // coords["cap"] * coords["offset"]) + (i % coords["cap"])
        safe_write(ws, (r, cols["name"]), row["name"]); safe_write(ws, (r, cols["grade"]), row["grade"]); safe_write(ws, (r, cols["dob"]), row["dob"]); safe_write(ws, (r, cols["jkf_no"]), row["jkf_no"])
        sex = row["sex"]; tk_c = cols["m_team_kata"] if sex=="男子" else cols["w_team_kata"]; tku_c = cols["m_team_kumite"] if sex=="男子" else cols["w_team_kumite"]
        if row.get("last_team_kata_chk"): safe_write(ws, (r, tk_c), "補" if row.get("last_team_kata_role")=="補" else "○", True)
        if row.get("last_team_kumi_chk"): safe_write(ws, (r, tku_c), "補" if row.get("last_team_kumi_role")=="補" else "○", True)
        k_c = cols["m_kata"] if sex=="男子" else cols["w_kata"]; ku_c = cols["m_kumite"] if sex=="男子" else cols["w_kumite"]
        if row.get("last_kata_chk"):
            v, rk = row.get("last_kata_val"), row.get("last_kata_rank", "")
            txt = "補" if v=="補" else (f"シ{rk}" if v=="シード" else f"○{rk}")
            safe_write(ws, (r, k_c), txt, True)
        if row.get("last_kumi_chk"):
            v, rk = row.get("last_kumi_val"), row.get("last_kumi_rank", "")
            if v=="補": txt = "補"
            elif t_conf["type"]=="standard": txt = f"シ{rk}" if v=="シード" else f"○{rk}"
            else: txt = str(v)
            safe_write(ws, (r, ku_c), txt, True)
    fname = f"申込書_{bn}.xlsx"; wb.save(fname); return fname, "成功"

# ---------------------------------------------------------
# 7. UI
# ---------------------------------------------------------
def school_page(s_id):
    st.markdown("""<style>div[data-testid="stRadio"] > div { flex-direction: row; }</style>""", unsafe_allow_html=True)
    auth = load_auth(); s_data = auth.get(s_id, {}); base_name = s_data.get("base_name", "")
    col_h1, col_h2 = st.columns([3, 1])
    with col_h1: st.markdown(f"### {base_name}高等学校")
    with col_h2:
        if st.button("🚪 ログアウト", type="secondary", use_container_width=True): st.query_params.clear(); st.session_state.clear(); st.rerun()
    st.divider()
    conf = load_conf(); active_tid = next((k for k, v in conf["tournaments"].items() if v["active"]), None)
    if not active_tid: st.error("現在受付中の大会はありません。"); return
    t_conf = conf["tournaments"][active_tid]; st.markdown(f"## 🥋 **令和{conf.get('year','〇')}年度 {t_conf['name']}**", unsafe_allow_html=True)
    
    if st.button("🔄 データを最新にする"):
        if "v2_master_cache" in st.session_state: del st.session_state["v2_master_cache"]
        if f"v2_entry_cache_{active_tid}" in st.session_state: del st.session_state[f"v2_entry_cache_{active_tid}"]
        st.rerun()

    menu = ["① 顧問登録", "② 部員名簿登録", "③ 大会エントリー"]
    if "school_menu_idx" not in st.session_state: st.session_state["school_menu_idx"] = 0
    selected_view = st.radio("メニュー選択", menu, index=st.session_state["school_menu_idx"], key="school_menu_radio", horizontal=True, label_visibility="collapsed")
    st.session_state["school_menu_idx"] = menu.index(selected_view)
    st.markdown("---")

    if selected_view == "① 顧問登録":
        st.warning("⚠️ **重要:** 編集内容は自動保存されません。変更後は必ず下の **『💾 顧問情報を保存』** ボタンを押してください。")
        np = st.text_input("校長名", s_data.get("principal", ""))
        adv_df = pd.DataFrame(s_data.get("advisors", []))
        for c in ["name", "role", "d1", "d2"]:
            if c not in adv_df.columns: adv_df[c] = ""
        edited_adv_df = st.data_editor(adv_df[["name", "role", "d1", "d2"]], column_config={"name": "氏名", "role": st.column_config.SelectboxColumn("役割", options=["審判", "競技記録", "係員"], required=True), "d1": "1日目", "d2": "2日目"}, num_rows="dynamic", use_container_width=True, hide_index=True)
        if st.button("💾 顧問情報を保存", type="primary"):
            if edited_adv_df["name"].isnull().any() or (edited_adv_df["name"] == "").any(): st.error("❌ 氏名が未入力です"); return
            with st.spinner("保存中..."):
                load_auth_cached.clear(); latest_auth = load_auth(); cur_s = latest_auth.get(s_id, s_data)
                cur_s["principal"], cur_s["advisors"] = np, edited_adv_df.to_dict(orient="records")
                latest_auth[s_id] = cur_s; save_auth(latest_auth); st.success("✅ 保存完了"); time.sleep(1); st.rerun()

    elif selected_view == "② 部員名簿登録":
        st.warning("⚠️ **重要:** 編集内容は自動保存されません。変更後は必ず下の **『💾 名簿を保存して更新』** ボタンを押してください。")
        master = load_members_master(force_reload=False); my_m = master[master['school_id']==s_id].copy()
        disp_df = my_m[["display_order", "name", "sex", "grade", "dob", "jkf_no"]].copy()
        edited_mem_df = st.data_editor(disp_df, column_config={"display_order": st.column_config.NumberColumn("No.", step=1), "name": "氏名", "sex": st.column_config.SelectboxColumn("性別", options=["男子", "女子"], required=True), "grade": st.column_config.SelectboxColumn("学年", options=[1, 2, 3], required=True), "dob": "生年月日", "jkf_no": "JKF番号"}, num_rows="dynamic", use_container_width=True, hide_index=True)
        if st.button("💾 名簿を保存して更新", type="primary"):
            with st.spinner("保存中..."):
                create_backup(); edited_mem_df["school_id"], edited_mem_df["active"] = s_id, True
                edited_mem_df['display_order'] = edited_mem_df['display_order'].apply(lambda x: str(int(x)) if pd.notnull(x) and str(x).strip() != "" else "")
                for c in MEMBERS_COLS:
                    if c not in edited_mem_df.columns: edited_mem_df[c] = ""
                latest_master = load_members_master(force_reload=True); new_master = pd.concat([latest_master[latest_master['school_id'] != s_id], edited_mem_df[MEMBERS_COLS]], ignore_index=True)
                save_members_master(new_master); st.success("✅ 更新完了"); time.sleep(1); st.rerun()

    elif selected_view == "③ 大会エントリー":
        merged = get_merged_data(s_id, active_tid)
        if merged.empty: st.warning("名簿を登録してください。"); return
        target_grades = [int(g) for g in t_conf['grades']]
        merged['sex_rank'] = merged['sex'].map({'男子': 0, '女子': 1}); merged['grade_rank'] = merged['grade'].map({3: 0, 2: 1, 1: 2})
        valid_members = merged[merged['grade'].isin(target_grades)].copy()
        def get_sort_key_ent(row):
            try: return float(row['display_order']) if pd.notna(row['display_order']) and str(row['display_order']).strip() else 999999.0
            except: return 999999.0
        valid_members['custom_order'] = valid_members.apply(get_sort_key_ent, axis=1)
        valid_members = valid_members.sort_values(by=['custom_order', 'sex_rank', 'grade_rank', 'name'])
        
        entries_update = load_entries(active_tid, force_reload=False); school_meta = entries_update.get(f"_meta_{s_id}", {"m_kumite_mode": "none", "w_kumite_mode": "none"})
        m_mode, w_mode = "5", "5"
        if t_conf["type"] == "shinjin":
            with st.expander("団体組手の設定 (新人戦)", expanded=True):
                c_m, c_w = st.columns(2)
                cur_m = school_meta.get("m_kumite_mode", "none"); idx_m = ["none", "5", "3"].index(cur_m) if cur_m in ["none", "5", "3"] else 0
                new_m = c_m.radio("男子 団体組手", ["出場しない", "5人制", "3人制"], index=idx_m, horizontal=True)
                m_mode = "none" if new_m == "出場しない" else ("5" if new_m == "5人制" else "3")
                cur_w = school_meta.get("w_kumite_mode", "none"); idx_w = ["none", "5", "3"].index(cur_w) if cur_w in ["none", "5", "3"] else 0
                new_w = c_w.radio("女子 団体組手", ["出場しない", "5人制", "3人制"], index=idx_w, horizontal=True)
                w_mode = "none" if new_w == "出場しない" else ("5" if new_w == "5人制" else "3")
                if new_m != cur_m or new_w != cur_w:
                    school_meta["m_kumite_mode"] = m_mode; school_meta["w_kumite_mode"] = w_mode; entries_update[f"_meta_{s_id}"] = school_meta; save_entries(active_tid, entries_update)

        with st.form("entry_form_unified"):
            cols = st.columns([1.7, 1.4, 1.4, 0.1, 3.1, 3.1]); cols[0].markdown("**氏名**"); cols[1].markdown("**団体形**"); cols[2].markdown("**団体組手**"); cols[4].markdown("**個人形**"); cols[5].markdown("**個人組手**")
            form_buffer = {}
            for i, r in valid_members.iterrows():
                uid = f"{s_id}_{r['name']}"; ns = 'background-color:#e8f5e9; padding:2px; font-weight:bold;' if r['sex']=="男子" else 'background-color:#ffebee; padding:2px; font-weight:bold;'
                c = st.columns([1.7, 1.4, 1.4, 0.1, 3.1, 3.1])
                c[0].markdown(f'<span style="{ns}">{r["grade"]}年 {r["name"]}</span>', unsafe_allow_html=True)
                val_tk = c[1].radio(f"tk_{uid}", ["なし", "正", "補"], index=["なし", "正", "補"].index(r.get("last_team_kata_role", "なし") if r.get("last_team_kata_role") in ["なし", "正", "補"] else "なし"), horizontal=True, label_visibility="collapsed")
                val_tku = c[2].radio(f"tku_{uid}", ["なし", "正", "補"], index=["なし", "正", "補"].index(r.get("last_team_kumi_role", "なし") if r.get("last_team_kumi_role") in ["なし", "正", "補"] else "なし"), horizontal=True, label_visibility="collapsed") if (m_mode if r['sex']=="男子" else w_mode) != "none" else "なし"
                if (m_mode if r['sex']=="男子" else w_mode) == "none": c[2].caption("-")
                opts_k = ["なし", "正", "補", "シード"] if t_conf["type"]=="standard" else ["なし", "正", "補"]
                ck1, ck2 = c[4].columns([1.5, 1])
                val_k = ck1.radio(f"k_{uid}", opts_k, index=opts_k.index(r.get("last_kata_val","なし") if r.get("last_kata_val") in opts_k else "なし"), horizontal=True, label_visibility="collapsed")
                rk_k = ck2.text_input("順位", r.get("last_kata_rank",""), key=f"rk_k_{uid}", label_visibility="collapsed", placeholder="順位")
                c5a, c5b = c[5].columns([1.8, 1])
                w_list = ["出場しない"] + [f"{w.strip()}kg級" for w in t_conf.get("weights_m" if r['sex']=="男子" else "weights_w", "").split(",")] + ["補欠"]
                if t_conf["type"]=="standard":
                    ku_v = c5a.radio(f"ku_{uid}", ["なし", "正", "補", "シード"], index=["なし", "正", "補", "シード"].index(r.get("last_kumi_val","なし") if r.get("last_kumi_val") in ["なし", "正", "補", "シード"] else "なし"), horizontal=True, label_visibility="collapsed")
                else:
                    cur_ku = r.get("last_kumi_val", "出場しない"); idx_ku = w_list.index(cur_ku) if cur_ku in w_list else 0
                    ku_v = c5a.selectbox("階級", w_list, index=idx_ku, key=f"sel_ku_{uid}", label_visibility="collapsed")
                rk_ku = c5b.text_input("順位", r.get("last_kumi_rank",""), key=f"rk_ku_{uid}", label_visibility="collapsed", placeholder="順位")
                form_buffer[uid] = {"val_tk": val_tk, "val_tku": val_tku, "val_k": val_k, "rank_k": rk_k, "ku_val": ku_v, "rank_ku": rk_ku, "name": r["name"], "sex": r["sex"]}
            if st.form_submit_button("✅ エントリーを保存 (全員分)"):
                with st.spinner("保存中..."):
                    cur_e = load_entries(active_tid, force_reload=True)
                    for uid, raw in form_buffer.items():
                        cur_e[uid] = {"team_kata_chk": raw["val_tk"]!="なし", "team_kata_role": raw["val_tk"] if raw["val_tk"]!="なし" else "", "team_kumi_chk": raw["val_tku"]!="なし", "team_kumi_role": raw["val_tku"] if raw["val_tku"]!="なし" else "", "kata_chk": raw["val_k"]!="なし", "kata_val": raw["val_k"] if raw["val_k"]!="なし" else "", "kata_rank": raw["rank_k"], "kumi_chk": raw["ku_val"] not in ["なし", "出場しない"], "kumi_val": raw["ku_val"] if raw["ku_val"] not in ["なし", "出場しない"] else "", "kumi_rank": raw["rank_ku"]}
                    save_entries(active_tid, cur_e); st.success("✅ 保存完了"); time.sleep(1); st.rerun()

        st.markdown("---")
        st.markdown("#### 📥 申込書の出力と提出")
        st.info("出力したExcelファイルに公印を押し、PDFや画像にしてから右側の枠へ提出してください。")
        c1, c2 = st.columns(2)
        with c1:
            st.markdown("##### 1. 申込書の作成")
            if st.button("📄 Excel申込書を作成する", type="secondary", use_container_width=True):
                 final_m = get_merged_data(s_id, active_tid); fp, msg = generate_excel(s_id, s_data, final_m, active_tid, t_conf)
                 if fp:
                     with open(fp, "rb") as f: st.download_button("📥 ダウンロード", f, fp, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
        with c2:
            st.markdown("##### 2. 申込書のアップロード")
            u_file = st.file_uploader("ファイルを選択 (PDF, JPG, PNG 等)", type=['pdf', 'jpg', 'jpeg', 'png'], label_visibility="collapsed")
            if u_file:
                if st.button("✅ 申込書を提出する", type="primary", use_container_width=True):
                    with st.spinner("送信中..."):
                        ok, res = upload_file_to_gas(u_file, base_name)
                        if ok: st.success("🎉 提出完了！管理者が確認します。")
                        else: st.error(f"❌ 失敗: {res}")

def admin_page():
    st.title("🔧 管理者画面")
    conf = load_conf()
    if not st.session_state.get("admin_ok", False):
        pw = st.text_input("Admin Password", type="password")
        if st.button("ログイン"):
            if pw == conf.get("admin_password", "1234"): st.session_state["admin_ok"] = True; st.rerun()
            else: st.error("不一致")
        return
    auth = load_auth(); admin_menu = ["🏆 大会設定", "📥 データ出力", "🏫 アカウント", "📅 年次処理"]
    admin_tab = st.radio("メニュー", admin_menu, index=st.session_state.get("admin_menu_idx", 0), horizontal=True)
    st.session_state["admin_menu_idx"] = admin_menu.index(admin_tab)
    st.divider()
    if admin_tab == "📥 データ出力":
        # (管理者用データ出力ロジック - 略)
        st.write("各校のエントリー状況の集計が可能です。")
        pass

def main():
    st.set_page_config(page_title="Entry System", layout="wide")
    st.title("🥋 高体連空手エントリーシステム")
    nav = st.radio("Nav", ["🏠 学校ログイン", "🆕 新規登録", "🔧 管理者"], horizontal=True, label_visibility="collapsed")
    auth = load_auth()
    if nav == "🏠 学校ログイン":
        if "logged_in_school" in st.session_state: school_page(st.session_state["logged_in_school"])
        else:
            with st.form("login_form"):
                sorted_auth = sorted(auth.items(), key=lambda x: to_safe_int(x[1].get('school_no', 999)))
                name_map = {f"{v.get('base_name')}高等学校": k for k, v in sorted_auth}
                ph = "（こちらから選択。ない場合は新規登録をしてください）"
                s_name = st.selectbox("学校名", [ph] + list(name_map.keys()))
                pw = st.text_input("パスワード", type="password")
                if st.form_submit_button("ログイン"):
                    if s_name != ph and name_map.get(s_name) and auth[name_map[s_name]]["password"] == pw:
                        st.session_state["logged_in_school"] = name_map[s_name]; st.rerun()
                    else: st.error("入力不備")
    elif nav == "🆕 新規登録":
        with st.form("reg"):
            bn = st.text_input("学校名 (「高等学校」不要)"); p = st.text_input("校長名"); pw = st.text_input("PW", type="password")
            if st.form_submit_button("登録"):
                if bn and pw:
                    nid = generate_school_id(); auth[nid] = {"base_name": bn, "password": pw, "principal": p, "school_no": 999, "advisors": []}
                    save_auth(auth); st.success("完了"); st.rerun()
    elif nav == "🔧 管理者": admin_page()

if __name__ == "__main__": main()
