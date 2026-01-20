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

# 安全なインポート
try:
    from openpyxl.cell import MergedCell
except ImportError:
    try:
        from openpyxl.cell.cell import MergedCell
    except ImportError:
        class MergedCell: pass

# ---------------------------------------------------------
# 1. 定数・初期設定・ヘルパー
# ---------------------------------------------------------
KEY_FILE = 'secrets.json'
SHEET_NAME = 'tournament_db' 
V2_PREFIX = "v2_" 

MEMBERS_COLS = ["school_id", "name", "sex", "grade", "dob", "jkf_no", "active"]

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
        val = ws.acell('A1').value
        if not val: return default
        parsed = json.loads(val)
        return parsed if parsed is not None else default
    except: return default

def save_json(tab_name, data):
    target_tab = f"{V2_PREFIX}{tab_name}"
    ws = get_worksheet_safe(target_tab)
    ws.update_acell('A1', json.dumps(data, ensure_ascii=False))

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
    df = df[MEMBERS_COLS]
    st.session_state["v2_master_cache"] = df
    return df

def save_members_master(df):
    ws = get_worksheet_safe(f"{V2_PREFIX}members"); ws.clear()
    df = df.fillna("")
    df['jkf_no'] = df['jkf_no'].astype(str)
    df['dob'] = df['dob'].astype(str)
    for c in MEMBERS_COLS:
        if c not in df.columns: df[c] = ""
    df_to_save = df[MEMBERS_COLS]
    ws.update([df_to_save.columns.tolist()] + df_to_save.astype(str).values.tolist())
    st.session_state["v2_master_cache"] = df_to_save

def archive_graduates(grad_df, auth_data):
    if grad_df.empty: return
    ws_grad = get_worksheet_safe(f"{V2_PREFIX}graduates")
    def get_school_name(sid):
        d = auth_data.get(sid, {})
        return d.get("base_name", "不明")
    grad_df = grad_df.copy()
    grad_df["archived_school_name"] = grad_df["school_id"].apply(get_school_name)
    grad_df["archived_date"] = datetime.date.today().strftime("%Y-%m-%d")
    if not ws_grad.get_all_values(): ws_grad.append_row(grad_df.columns.tolist())
    ws_grad.append_rows(grad_df.astype(str).values.tolist())

def clear_graduates_archive():
    ws_grad = get_worksheet_safe(f"{V2_PREFIX}graduates"); ws_grad.clear()

def get_graduates_df():
    try:
        ws = get_worksheet_safe(f"{V2_PREFIX}graduates")
        recs = ws.get_all_records()
        return pd.DataFrame(recs) if recs else pd.DataFrame()
    except: return pd.DataFrame()

def load_entries(tournament_id, force_reload=False):
    key = f"v2_entry_cache_{tournament_id}"
    if not force_reload and key in st.session_state: return st.session_state[key]
    data = load_json(f"entry_{tournament_id}", {})
    st.session_state[key] = data
    return data

def save_entries(tournament_id, data):
    save_json(f"entry_{tournament_id}", data)
    st.session_state[f"v2_entry_cache_{tournament_id}"] = data

@st.cache_data
def load_auth_cached(): return load_json("auth", {})
def load_auth(): return load_auth_cached()
def save_auth(d):
    save_json("auth", d)
    load_auth_cached.clear()

def load_schools(): return load_json("schools", {})

def load_conf():
    default_conf = {"year": "6", "tournaments": DEFAULT_TOURNAMENTS, "limits": DEFAULT_LIMITS, "admin_password": "1234"}
    data = load_json("config", default_conf)
    if "limits" not in data: data["limits"] = DEFAULT_LIMITS
    if "tournaments" not in data: data["tournaments"] = DEFAULT_TOURNAMENTS
    return data

def save_conf(d): save_json("config", d)

# ---------------------------------------------------------
# 4. ロジック
# ---------------------------------------------------------
def create_backup():
    df = load_members_master(force_reload=False)
    ws_bk_mem = get_worksheet_safe(f"{V2_PREFIX}members_backup"); ws_bk_mem.clear()
    df = df.fillna("")
    df_bk = df[MEMBERS_COLS]
    ws_bk_mem.update([df_bk.columns.tolist()] + df_bk.astype(str).values.tolist())
    conf = load_conf()
    ws_bk_conf = get_worksheet_safe(f"{V2_PREFIX}config_backup")
    ws_bk_conf.update_acell('A1', json.dumps(conf, ensure_ascii=False))

def restore_from_backup():
    try:
        ws_bk_mem = get_worksheet_safe(f"{V2_PREFIX}members_backup")
        recs = ws_bk_mem.get_all_records()
        df = pd.DataFrame(recs) if recs else pd.DataFrame(columns=MEMBERS_COLS)
        if not df.empty:
            df['grade'] = pd.to_numeric(df['grade'], errors='coerce').fillna(0).astype(int)
            save_members_master(df)
    except: return "名簿の復元に失敗しました"
    try:
        ws_bk_conf = get_worksheet_safe(f"{V2_PREFIX}config_backup")
        val = ws_bk_conf.acell('A1').value
        if val: save_conf(json.loads(val))
    except: return "設定の復元に失敗しました"
    return "✅ バックアップから復元しました"

def perform_year_rollover():
    create_backup()
    if "v2_master_cache" in st.session_state: del st.session_state["v2_master_cache"]
    df = load_members_master(force_reload=True)
    if df.empty: return "データがありません"
    
    df['grade'] = df['grade'] + 1
    graduates = df[df['grade'] > 3].copy()
    current = df[df['grade'] <= 3].copy()
    if not graduates.empty:
        auth = load_auth()
        archive_graduates(graduates, auth)
    save_members_master(current)
    conf = load_conf()
    for tid in conf["tournaments"].keys(): save_entries(tid, {})
    try:
        conf["year"] = str(int(conf["year"]) + 1)
        save_conf(conf)
    except: pass
    return f"✅ 新年度更新完了。{len(graduates)}名の卒業生データをアーカイブしました。"

def get_merged_data(school_id, tournament_id):
    master = load_members_master(force_reload=False)
    if master.empty: return pd.DataFrame()
    my_members = master[master['school_id'] == school_id].copy()
    entries = load_entries(tournament_id, force_reload=False)
    def get_ent(row, key):
        uid = f"{row['school_id']}_{row['name']}"
        val = entries.get(uid, {}).get(key, None)
        return val
    cols_to_add = ["team_kata_chk", "team_kata_role", "team_kumi_chk", "team_kumi_role",
                   "kata_chk", "kata_val", "kata_rank", "kumi_chk", "kumi_val", "kumi_rank"]
    for c in cols_to_add:
        my_members[f"last_{c}"] = my_members.apply(lambda r: get_ent(r, c), axis=1)
    return my_members

def validate_counts(members_df, entries_data, limits, t_type, school_meta, school_id):
    errs = []
    for sex in ["男子", "女子"]:
        sex_df = members_df[members_df['sex'] == sex]
        cnt_tk = 0; cnt_tku = 0
        cnt_ind_k_reg = 0; cnt_ind_k_sub = 0
        cnt_ind_ku_reg = 0; cnt_ind_ku_sub = 0
        for _, r in sex_df.iterrows():
            uid = f"{school_id}_{r['name']}"
            ent = entries_data.get(uid, {})
            if ent.get("team_kata_chk") and ent.get("team_kata_role") == "正": cnt_tk += 1
            if ent.get("team_kumi_chk") and ent.get("team_kumi_role") == "正": cnt_tku += 1
            if ent.get("kata_chk"):
                k_val = ent.get("kata_val")
                if k_val == "補": cnt_ind_k_sub += 1
                elif k_val == "正": cnt_ind_k_reg += 1 
            if ent.get("kumi_chk"):
                val = ent.get("kumi_val")
                if val == "補": cnt_ind_ku_sub += 1
                elif val == "正": cnt_ind_ku_reg += 1
                elif t_type != "standard" and val and val != "出場しない" and val != "なし" and val != "シード" and val != "補": cnt_ind_ku_reg += 1

        if cnt_tk > 0:
            mn, mx = limits["team_kata"]["min"], limits["team_kata"]["max"]
            if not (mn <= cnt_tk <= mx): errs.append(f"❌ {sex}団体形: 正選手は {mn}～{mx}名で登録してください。(現在{cnt_tk}名)")
        if cnt_tku > 0:
            mode = "5"
            if t_type == "shinjin":
                mode_key = "m_kumite_mode" if sex == "男子" else "w_kumite_mode"
                mode = school_meta.get(mode_key, "none")
            if mode == "5":
                mn, mx = limits["team_kumite_5"]["min"], limits["team_kumite_5"]["max"]
                if not (mn <= cnt_tku <= mx): errs.append(f"❌ {sex}団体組手(5人制): 正選手は {mn}～{mx}名で登録してください。(現在{cnt_tku}名)")
            elif mode == "3":
                mn, mx = limits["team_kumite_3"]["min"], limits["team_kumite_3"]["max"]
                if not (mn <= cnt_tku <= mx): errs.append(f"❌ {sex}団体組手(3人制): 正選手は {mn}～{mx}名で登録してください。(現在{cnt_tku}名)")
        
        if cnt_ind_k_reg > limits["ind_kata_reg"]["max"]: errs.append(f"❌ {sex}個人形(正): 上限 {limits['ind_kata_reg']['max']}名を超えています。(シード除く)")
        if cnt_ind_k_sub > limits["ind_kata_sub"]["max"]: errs.append(f"❌ {sex}個人形(補): 上限 {limits['ind_kata_sub']['max']}名を超えています。")
        if cnt_ind_ku_reg > limits["ind_kumi_reg"]["max"]: errs.append(f"❌ {sex}個人組手(正): 上限 {limits['ind_kumi_reg']['max']}名を超えています。(シード除く)")
        if cnt_ind_ku_sub > limits["ind_kumi_sub"]["max"]: errs.append(f"❌ {sex}個人組手(補): 上限 {limits['ind_kumi_sub']['max']}名を超えています。")
    return errs

# ---------------------------------------------------------
# 5. Excel生成
# ---------------------------------------------------------
def safe_write(ws, target, value, align_center=False):
    if value is None: value = ""
    if isinstance(target, str): cell = ws[target]
    else: cell = ws.cell(row=target[0], column=target[1])
    if isinstance(cell, MergedCell):
        for r in ws.merged_cells.ranges:
            if cell.coordinate in r: cell = ws.cell(row=r.min_row, column=r.min_col); break
    val_str = str(value)
    if val_str.endswith("年") and val_str[:-1].isdigit(): val_str = val_str.replace("年", "")
    cell.value = val_str
    if align_center: cell.alignment = Alignment(horizontal='center', vertical='center')

def generate_excel(school_id, school_data, members_df, t_id, t_conf):
    coords = COORD_DEF
    template_file = t_conf.get("template", "template.xlsx")
    try: wb = openpyxl.load_workbook(template_file); ws = wb.active
    except: return None, f"{template_file} が見つかりません。"
    conf = load_conf()
    safe_write(ws, coords["year"], conf.get("year", ""))
    safe_write(ws, coords["tournament_name"], t_conf.get("name", ""))
    safe_write(ws, coords["date"], f"令和{datetime.date.today().year-2018}年{datetime.date.today().month}月{datetime.date.today().day}日")
    base_name = school_data.get("base_name", "")
    safe_write(ws, coords["school_name"], base_name)
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
            role = row.get("last_team_kata_role")
            safe_write(ws, (r, tk_col), "補" if role=="補" else "○", True)
        if row.get("last_team_kumi_chk"):
            role = row.get("last_team_kumi_role")
            safe_write(ws, (r, tku_col), "補" if role=="補" else "○", True)
        k_col = cols["m_kata"] if sex=="男子" else cols["w_kata"]
        ku_col = cols["m_kumite"] if sex=="男子" else cols["w_kumite"]
        if row.get("last_kata_chk"):
            val = row.get("last_kata_val")
            rank = row.get("last_kata_rank", "")
            if val == "補": txt = "補"
            elif t_conf["type"] == "standard": txt = f"シ{rank}" if val == "シード" else f"○{rank}"
            else: txt = "○"
            safe_write(ws, (r, k_col), txt, True)
        if row.get("last_kumi_chk"):
            val = row.get("last_kumi_val")
            rank = row.get("last_kumi_rank", "")
            if val == "補": txt = "補"
            elif t_conf["type"] == "standard": txt = f"シ{rank}" if val == "シード" else f"○{rank}"
            elif t_conf["type"] == "weight": txt = str(val)
            elif t_conf["type"] == "division": txt = str(val)
            else: txt = "○"
            safe_write(ws, (r, ku_col), txt, True)
    fname = f"申込書_{base_name}.xlsx"
    wb.save(fname)
    return fname, "作成成功"

def generate_tournament_excel(all_data, t_type, auth_data):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        sheets_data = {}
        for row in all_data:
            name = row['name']
            sid = row['school_id']
            s_data = auth_data.get(sid, {})
            school_short = s_data.get("short_name")
            if not school_short: school_short = s_data.get("base_name", "")
            sex = row['sex']
            if row.get('kata_chk'):
                k_val = row.get('kata_val')
                k_rank = row.get('kata_rank', '')
                if k_val and k_val != '補' and k_val != 'なし' and k_val != '出場しない':
                    sheet_name = f"{sex}個人形"
                    rank_cell = k_rank if k_val == '正' else ''
                    seed_cell = k_rank if k_val == 'シード' else ''
                    record = {"個人形_順位": rank_cell, "名前": name, "学校名": school_short, "シード順位": seed_cell}
                    if sheet_name not in sheets_data: sheets_data[sheet_name] = []
                    sheets_data[sheet_name].append(record)
            if row.get('kumi_chk'):
                ku_val = row.get('kumi_val')
                ku_rank = row.get('kumi_rank', '')
                if ku_val and ku_val != '補' and ku_val != 'なし' and ku_val != '出場しない':
                    if t_type == 'standard':
                        sheet_name = f"{sex}個人組手"
                        is_seed = (ku_val == 'シード'); is_reg = (ku_val == '正')
                    else:
                        sheet_name = f"{sex}個人組手_{ku_val}"
                        is_seed = False; is_reg = True 
                    rank_cell = ku_rank if is_reg else ''
                    seed_cell = ku_rank if is_seed else ''
                    record = {"個人組手_順位": rank_cell, "名前": name, "学校名": school_short, "シード順位": seed_cell}
                    if sheet_name not in sheets_data: sheets_data[sheet_name] = []
                    sheets_data[sheet_name].append(record)
        sorted_sheet_names = sorted(sheets_data.keys())
        for s_name in sorted_sheet_names:
            recs = sheets_data[s_name]
            header_rank = "個人組手_順位" if "組手" in s_name else "個人形_順位"
            df_out = pd.DataFrame(recs, columns=[header_rank, "名前", "学校名", "シード順位"])
            df_out.to_excel(writer, sheet_name=s_name, index=False)
    return output.getvalue()

def generate_summary_excel(master_df, entries, auth_data, t_type):
    summary_rows = []
    sorted_schools = sorted(auth_data.items(), key=lambda x: to_safe_int(x[1].get('school_no')))
    for s_id, s_data in sorted_schools:
        s_no = s_data.get('school_no', '')
        s_name = s_data.get("short_name")
        if not s_name: s_name = s_data.get("base_name", "")
        s_members = master_df[master_df['school_id'] == s_id]
        m_tk_flag = ""; m_tku_flag = ""; w_tk_flag = ""; w_tku_flag = ""
        m_k_cnt = 0; m_ku_cnt = 0; w_k_cnt = 0; w_ku_cnt = 0
        reg_player_names = set()
        for _, r in s_members.iterrows():
            uid = f"{s_id}_{r['name']}"
            ent = entries.get(uid, {})
            sex = r['sex']
            if sex == "男子":
                if ent.get("team_kata_chk"): m_tk_flag = "○"
                if ent.get("team_kumi_chk"): m_tku_flag = "○"
            else:
                if ent.get("team_kata_chk"): w_tk_flag = "○"
                if ent.get("team_kumi_chk"): w_tku_flag = "○"
            if ent.get("kata_chk"):
                val = ent.get("kata_val")
                if val and val != "補" and val != "なし" and val != "出場しない":
                    if sex == "男子": m_k_cnt += 1
                    else: w_k_cnt += 1
            if ent.get("kumi_chk"):
                val = ent.get("kumi_val")
                if val and val != "補" and val != "なし" and val != "出場しない":
                    if sex == "男子": m_ku_cnt += 1
                    else: w_ku_cnt += 1
            is_reg = False
            if ent.get("team_kata_chk") and ent.get("team_kata_role") == "正": is_reg = True
            if ent.get("team_kumi_chk") and ent.get("team_kumi_role") == "正": is_reg = True
            kv = ent.get("kata_val")
            if ent.get("kata_chk") and kv and kv != "補" and kv != "なし" and kv != "出場しない": is_reg = True
            kuv = ent.get("kumi_val")
            if ent.get("kumi_chk") and kuv and kuv != "補" and kuv != "なし" and kuv != "出場しない": is_reg = True
            if is_reg: reg_player_names.add(r['name'])
        summary_rows.append({
            "学校No": s_no, "学校名": s_name, 
            "男団体形": m_tk_flag, "男団体組手": m_tku_flag,
            "男個人形": m_k_cnt if m_k_cnt > 0 else "", "男個人組手": m_ku_cnt if m_ku_cnt > 0 else "",
            "女団体形": w_tk_flag, "女団体組手": w_tku_flag,
            "女個人形": w_k_cnt if w_k_cnt > 0 else "", "女個人組手": w_ku_cnt if w_ku_cnt > 0 else "",
            "正選手合計": len(reg_player_names)
        })
    df_out = pd.DataFrame(summary_rows)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_out.to_excel(writer, sheet_name="参加校一覧", index=False)
    return output.getvalue()

def generate_advisor_excel(schools_data, auth_data):
    rows = []
    sorted_schools = sorted(auth_data.items(), key=lambda x: to_safe_int(x[1].get('school_no')))
    cnt_judge = 0; cnt_staff = 0
    for s_id, s_auth in sorted_schools:
        s_no = s_auth.get('school_no', '')
        s_name = s_auth.get("short_name")
        if not s_name: s_name = s_auth.get("base_name", "")
        advs = s_auth.get("advisors", [])
        for a in advs:
            name = a.get("name", "")
            if not name: continue
            role = a.get("role", "審判")
            d1 = "○" if a.get("d1") else "×"
            d2 = "○" if a.get("d2") else "×"
            if role == "審判": cnt_judge += 1
            if role == "係員": cnt_staff += 1
            rows.append({"No": s_no, "学校名": s_name, "顧問氏名": name, "役割": role, "1日目": d1, "2日目": d2})
    df_list = pd.DataFrame(rows)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_list.to_excel(writer, sheet_name="顧問一覧", index=False, startcol=0)
        df_summary = pd.DataFrame([{"項目": "審判 合計", "人数": cnt_judge}, {"項目": "係員 合計", "人数": cnt_staff}])
        df_summary.to_excel(writer, sheet_name="顧問一覧", index=False, startcol=7)
    return output.getvalue()

# ---------------------------------------------------------
# 7. UI (v2) - Lazy Loading & 負荷軽減徹底 & UX改善
# ---------------------------------------------------------
def school_page(s_id):
    st.markdown("""<style>div[data-testid="stRadio"] > div { flex-direction: row; }</style>""", unsafe_allow_html=True)
    auth = load_auth()
    s_data = auth.get(s_id, {})
    base_name = s_data.get("base_name", "")
    
    col_h1, col_h2 = st.columns([3, 1])
    with col_h1: st.markdown(f"### {base_name}高等学校")
    with col_h2:
        if st.button("🚪 ログアウト", type="secondary", use_container_width=True):
            st.query_params.clear(); st.session_state.clear(); st.rerun()
    st.divider()

    conf = load_conf()
    active_tid = next((k for k, v in conf["tournaments"].items() if v["active"]), None)
    t_conf = conf["tournaments"].get(active_tid, {}) if active_tid else {}
    if not active_tid: st.error("現在受付中の大会はありません。"); return
    
    disp_year = conf.get("year", "〇")
    st.markdown(f"## 🥋 **令和{disp_year}年度 {t_conf['name']}** <small>エントリー画面</small>", unsafe_allow_html=True)
    
    if st.button("🔄 データを最新にする"):
        if "v2_master_cache" in st.session_state: del st.session_state["v2_master_cache"]
        if f"v2_entry_cache_{active_tid}" in st.session_state: del st.session_state[f"v2_entry_cache_{active_tid}"]
        st.success("最新データを読み込みました"); time.sleep(0.5); st.rerun()

    menu = ["① 顧問登録", "② 部員名簿登録", "③ 大会エントリー"]
    selected_view = st.radio("メニュー選択", menu, key="school_menu_key", horizontal=True, label_visibility="collapsed")
    st.markdown("---")

    if selected_view == "① 顧問登録":
        st.warning("⚠️ **重要:** 編集内容は自動保存されません。変更後は必ず下の **『💾 顧問情報を保存』** ボタンを押してください。")
        st.caption("※下記の表を直接編集し、最後に「保存」ボタンを押してください。")
        
        with st.form("advisor_form"):
            np = st.text_input("校長名", s_data.get("principal", ""))
            st.form_submit_button("校長名を反映(仮保存)")

        current_advs = s_data.get("advisors", [])
        adv_df = pd.DataFrame(current_advs)
        for c in ["name", "role", "d1", "d2"]:
            if c not in adv_df.columns: adv_df[c] = ""
        
        col_conf_adv = {
            "name": st.column_config.TextColumn("氏名"),
            "role": st.column_config.SelectboxColumn("役割", options=["審判", "競技記録", "係員"], required=True),
            "d1": st.column_config.CheckboxColumn("1日目"),
            "d2": st.column_config.CheckboxColumn("2日目"),
        }
        
        # 修正: ヘッダー名削除
        edited_adv_df = st.data_editor(adv_df[["name", "role", "d1", "d2"]], 
                                       column_config=col_conf_adv, 
                                       num_rows="dynamic", use_container_width=True, key="adv_editor")
        
        st.caption("💡 **削除するには:** 表の左端（行番号）をクリックして行を選び、キーボードの **Delete** キーを押してください。その後、保存ボタンで確定します。")
        
        if st.button("💾 顧問情報を保存", type="primary"):
            if edited_adv_df["name"].isnull().any() or (edited_adv_df["name"] == "").any():
                st.error("❌ 氏名が未入力の行があります。"); return
            with st.spinner("💾 データを保存しています..."):
                new_advs = edited_adv_df.to_dict(orient="records")
                new_advs = [x for x in new_advs if x["name"]]
                s_data["principal"] = np
                s_data["advisors"] = new_advs
                auth[s_id] = s_data
                save_auth(auth); time.sleep(1)
            st.success("✅ 保存しました！"); time.sleep(2); st.rerun()

    elif selected_view == "② 部員名簿登録":
        st.warning("⚠️ **重要:** 編集内容は自動保存されません。変更後は必ず下の **『💾 名簿を保存して更新』** ボタンを押してください。")
        st.info("💡 ここは「全大会共通」の名簿です。")
        
        master = load_members_master(force_reload=False)
        my_m = master[master['school_id']==s_id].copy()
        other_m = master[master['school_id']!=s_id].copy()
        
        disp_df = my_m[["name", "sex", "grade", "dob", "jkf_no"]].copy()
        col_config_mem = {
            "name": st.column_config.TextColumn("氏名"), 
            "sex": st.column_config.SelectboxColumn("性別", options=["男子", "女子"], required=True),
            "grade": st.column_config.SelectboxColumn("学年", options=[1, 2, 3], required=True),
            "dob": st.column_config.TextColumn("生年月日(任意)"),
            "jkf_no": st.column_config.TextColumn("JKF番号(任意)")
        }
        
        # 修正: ヘッダー名削除
        edited_mem_df = st.data_editor(disp_df, column_config=col_config_mem, num_rows="dynamic", use_container_width=True, key="mem_editor")
        
        st.caption("💡 **削除するには:** 表の左端（行番号）をクリックして行を選び、キーボードの **Delete** キーを押してください。その後、保存ボタンで確定します。")

        if st.button("💾 名簿を保存して更新", type="primary"):
            if edited_mem_df["name"].isnull().any() or (edited_mem_df["name"] == "").any():
                st.error("❌ 氏名が未入力の行があります。"); return
            if edited_mem_df["sex"].isnull().any() or (edited_mem_df["sex"] == "").any():
                st.error("❌ 性別が未選択の行があります。"); return
            if edited_mem_df["grade"].isnull().any() or (edited_mem_df["grade"] == "").any():
                st.error("❌ 学年が未選択の行があります。"); return
            
            with st.spinner("💾 データを保存しています..."):
                create_backup() 
                edited_mem_df["school_id"] = s_id
                edited_mem_df["active"] = True
                for c in MEMBERS_COLS:
                    if c not in edited_mem_df.columns: edited_mem_df[c] = ""
                new_master = pd.concat([other_m, edited_mem_df[MEMBERS_COLS]], ignore_index=True)
                save_members_master(new_master)
                time.sleep(1)
            st.success("✅ 名簿を更新しました（自動バックアップ完了）"); time.sleep(2); st.rerun()
        
        st.divider()
        st.markdown("##### 📋 登録済み部員リスト（確認用）")
        master_check = load_members_master(force_reload=False)
        my_check = master_check[master_check['school_id']==s_id]
        
        rename_map = {'grade': '学年', 'name': '氏名', 'jkf_no': 'JKF番号'}
        
        c_male, c_female = st.columns(2)
        with c_male:
            st.markdown("###### 🚹 男子部員")
            m_df = my_check[my_check['sex'] == '男子'].sort_values(by=['grade', 'name'], ascending=[False, True])
            if not m_df.empty: st.dataframe(m_df[['grade','name','jkf_no']].rename(columns=rename_map), hide_index=True, use_container_width=True)
            else: st.caption("登録なし")
        with c_female:
            st.markdown("###### 🚺 女子部員")
            w_df = my_check[my_check['sex'] == '女子'].sort_values(by=['grade', 'name'], ascending=[False, True])
            if not w_df.empty: st.dataframe(w_df[['grade','name','jkf_no']].rename(columns=rename_map), hide_index=True, use_container_width=True)
            else: st.caption("登録なし")

    elif selected_view == "③ 大会エントリー":
        target_grades = [int(g) for g in t_conf['grades']]
        st.markdown(f"**出場対象学年:** {target_grades} 年生")
        
        st.markdown("""
            <div style="background-color: #ffebee; border: 1px solid #ef9a9a; padding: 10px; border-radius: 5px; color: #c62828;">
                <h4 style="margin:0;">⚠️ 順位入力について（重要）</h4>
                <p style="font-size: 1.1em; font-weight: bold; margin: 5px 0;">
                    トーナメント作成の優先順位に使用します。シード権を持っているものはシード順位を入力し、正選手の場合はトーナメント作成における優先順位をつけてください。この順位をもとにトーナメントの位置を決めます。補欠の場合は入力はいりません。<br>
                    （例：1, 2, 3...）
                </p>
            </div>
            <br>
        """, unsafe_allow_html=True)

        merged = get_merged_data(s_id, active_tid)
        if merged.empty: st.warning("エントリー可能な部員がいません。名簿を登録してください。"); return
        
        merged['sex_rank'] = merged['sex'].map({'男子': 0, '女子': 1})
        merged['grade_rank'] = merged['grade'].map({3: 0, 2: 1, 1: 2})
        valid_members = merged[merged['grade'].isin(target_grades)].sort_values(by=['sex_rank', 'grade_rank', 'name']).copy()
        
        entries_update = load_entries(active_tid, force_reload=False)
        meta_key = f"_meta_{s_id}"
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
                if new_m != curr_m or new_w != curr_w:
                    school_meta["m_kumite_mode"] = m_mode; school_meta["w_kumite_mode"] = w_mode
                    entries_update[meta_key] = school_meta
                    save_entries(active_tid, entries_update)

        with st.form("entry_form_unified"):
            cols = st.columns([1.7, 1.4, 1.4, 0.1, 3.1, 3.1])
            cols[0].markdown("**氏名**")
            cols[1].markdown("**団体形**")
            cols[2].markdown("**団体組手**")
            cols[4].markdown("**個人形**")
            cols[5].markdown("**個人組手**")
            form_buffer = {}
            for i, r in valid_members.iterrows():
                uid = f"{s_id}_{r['name']}"
                name_style = 'background-color:#e8f5e9; color:#1b5e20; padding:2px 6px; border-radius:4px; font-weight:bold;' if r['sex'] == "男子" else 'background-color:#ffebee; color:#b71c1c; padding:2px 6px; border-radius:4px; font-weight:bold;'
                c = st.columns([1.7, 1.4, 1.4, 0.1, 3.1, 3.1])
                c[0].markdown(f'<span style="{name_style}">{r["grade"]}年 {r["name"]}</span>', unsafe_allow_html=True)
                def_tk = r.get("last_team_kata_role", "なし"); opts_tk = ["なし", "正", "補"]
                if def_tk not in opts_tk: def_tk = "なし"
                val_tk = c[1].radio(f"tk_{uid}", opts_tk, index=opts_tk.index(def_tk), horizontal=True, label_visibility="collapsed")
                mode = m_mode if r['sex']=="男子" else w_mode
                if mode != "none":
                    def_tku = r.get("last_team_kumi_role", "なし"); opts_tku = ["なし", "正", "補"]
                    if def_tku not in opts_tku: def_tku = "なし"
                    val_tku = c[2].radio(f"tku_{uid}", opts_tku, index=opts_tku.index(def_tku), horizontal=True, label_visibility="collapsed")
                else: val_tku = "なし"; c[2].caption("-")
                def_k = r.get("last_kata_val", "なし")
                if t_conf["type"] == "standard": opts_k = ["なし", "シード", "正", "補"]
                else: opts_k = ["なし", "正", "補"]
                if def_k not in opts_k: def_k = "なし"
                ck1, ck2 = c[4].columns([1.5, 1])
                val_k = ck1.radio(f"k_{uid}", opts_k, index=opts_k.index(def_k), horizontal=True, label_visibility="collapsed")
                rank_k = ck2.text_input("順位", r.get("last_kata_rank",""), key=f"rk_k_{uid}", label_visibility="collapsed", placeholder="順位")
                c5a, c5b = c[5].columns([1.8, 1])
                w_key = "weights_m" if r['sex'] == "男子" else "weights_w"
                w_list = ["出場しない"] + [f"{w.strip()}kg級" for w in t_conf.get(w_key, "").split(",")] + ["補欠"]
                raw_kumi = r.get("last_kumi_val")
                def_val = str(raw_kumi) if raw_kumi else "出場しない"
                if t_conf["type"] == "standard":
                    opts_ku = ["なし", "シード", "正", "補"]
                    if def_val == "出場しない" or def_val not in opts_ku: def_val = "なし"
                    ku_val = c5a.radio(f"ku_{uid}", opts_ku, index=opts_ku.index(def_val), horizontal=True, label_visibility="collapsed")
                else:
                    if t_conf["type"] == "weight" and def_val not in w_list: def_val = f"{def_val}kg級"
                    try: idx = w_list.index(def_val)
                    except: idx = 0
                    ku_val = c5a.selectbox("階級", w_list, index=idx, key=f"sel_ku_{uid}", label_visibility="collapsed")
                rank_ku = c5b.text_input("順位", r.get("last_kumi_rank",""), key=f"rk_ku_{uid}", label_visibility="collapsed", placeholder="順位")
                form_buffer[uid] = {"val_tk": val_tk, "val_tku": val_tku, "val_k": val_k, "rank_k": rank_k, "ku_val": ku_val, "rank_ku": rank_ku, "name": r["name"], "sex": r["sex"]}

            if st.form_submit_button("✅ エントリーを保存 (全員分)"):
                has_error = False; temp_processed = {}
                duplicate_checker = {} 

                for uid, raw in form_buffer.items():
                    k_chk = (raw["val_k"] != "なし"); k_val = raw["val_k"] if k_chk else ""
                    if k_chk:
                        if (k_val == "正" or k_val == "シード") and not raw["rank_k"]:
                            st.error(f"❌ {uid} 個人形: {k_val}選手の実績順位が入力されていません。"); has_error = True
                        
                        if (k_val == "正" or k_val == "シード") and raw["rank_k"]:
                            check_key = f"{raw['sex']}_kata_{k_val}"
                            clean_rank = to_half_width(raw["rank_k"])
                            if check_key not in duplicate_checker: duplicate_checker[check_key] = {}
                            if clean_rank not in duplicate_checker[check_key]: duplicate_checker[check_key][clean_rank] = []
                            duplicate_checker[check_key][clean_rank].append(raw["name"])

                    ku_chk = (raw["ku_val"] not in ["なし", "出場しない"]); ku_val = raw["ku_val"] if ku_chk else ""
                    if ku_chk:
                        is_reg = (t_conf["type"] == "weight" and ku_val != "補欠") or (t_conf["type"] == "standard" and ku_val == "正")
                        is_seed = (t_conf["type"] == "standard" and ku_val == "シード")
                        if (is_reg or is_seed) and not raw["rank_ku"]:
                            st.error(f"❌ {uid} 個人組手: 実績順位が入力されていません。"); has_error = True
                        
                        if t_conf["type"] == "standard" and (is_reg or is_seed) and raw["rank_ku"]:
                            role_key = "シード" if is_seed else "正"
                            check_key = f"{raw['sex']}_kumite_{role_key}"
                            clean_rank = to_half_width(raw["rank_ku"])
                            if check_key not in duplicate_checker: duplicate_checker[check_key] = {}
                            if clean_rank not in duplicate_checker[check_key]: duplicate_checker[check_key][clean_rank] = []
                            duplicate_checker[check_key][clean_rank].append(raw["name"])

                    temp_processed[uid] = {
                        "team_kata_chk": (raw["val_tk"]!="なし"), "team_kata_role": raw["val_tk"] if raw["val_tk"]!="なし" else "",
                        "team_kumi_chk": (raw["val_tku"]!="なし"), "team_kumi_role": raw["val_tku"] if raw["val_tku"]!="なし" else "",
                        "kata_chk": k_chk, "kata_val": k_val, "kata_rank": to_half_width(raw["rank_k"]),
                        "kumi_chk": ku_chk, "kumi_val": ku_val, "kumi_rank": to_half_width(raw["rank_ku"])
                    }
                
                for key, ranks in duplicate_checker.items():
                    for rank_val, names in ranks.items():
                        if len(names) > 1:
                            parts = key.split("_")
                            sex_lbl = parts[0]; type_lbl = "形" if parts[1]=="kata" else "組手"; role_lbl = parts[2]
                            name_list_str = ", ".join(names)
                            st.error(f"❌ {sex_lbl} 個人{type_lbl} ({role_lbl}選手) で順位『{rank_val}』が重複しています: {name_list_str}")
                            has_error = True

                if not has_error:
                    with st.spinner("💾 エントリーを保存しています..."):
                        current_entries = load_entries(active_tid, force_reload=True)
                        current_entries.update(temp_processed)
                        errs = validate_counts(valid_members, current_entries, conf["limits"], t_conf["type"], {"m_kumite_mode":m_mode, "w_kumite_mode":w_mode}, s_id)
                        if errs:
                            for e in errs: st.error(e)
                        else:
                            save_entries(active_tid, current_entries)
                            time.sleep(1)
                    st.success("✅ 保存しました！"); time.sleep(2); st.rerun()

        st.markdown("---")
        if st.button("📥 Excel申込書を作成する", type="primary"):
             final_merged = get_merged_data(s_id, active_tid)
             fp, msg = generate_excel(s_id, s_data, final_merged, active_tid, t_conf)
             if fp:
                 with open(fp, "rb") as f: st.download_button("📥 ダウンロード", f, fp, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

def admin_page():
    st.title("🔧 管理者画面 (v2)")
    conf = load_conf()
    pw_input = st.text_input("Admin Password", type="password")
    st.caption("※パスワードを忘れた場合は、secrets.jsonを確認するか、システム開発者に連絡してください")
    if not st.button("管理者ログイン"): 
        if "admin_ok" not in st.session_state: st.warning("パスワードを入力してログインボタンを押してください"); return
    
    if pw_input == conf.get("admin_password", "1234"): st.session_state["admin_ok"] = True
    else: st.error("パスワードが違います"); return

    auth = load_auth()
    admin_menu = ["🏆 大会設定", "📥 データ出力", "🏫 アカウント(編集)", "📅 年次処理"]
    admin_tab = st.radio("メニュー", admin_menu, key="admin_menu_key", horizontal=True)
    st.divider()

    if admin_tab == "🏆 大会設定":
        st.subheader("基本設定")
        with st.form("conf_basic"):
            new_year = st.text_input("現在の年度", conf.get("year", "6"))
            t_opts = list(conf["tournaments"].keys())
            active_now = next((k for k, v in conf["tournaments"].items() if v["active"]), None)
            new_active = st.radio("受付中の大会", t_opts, index=t_opts.index(active_now) if active_now else 0, format_func=lambda x: conf["tournaments"][x]["name"])
            if st.form_submit_button("設定を保存 & 大会切替"):
                conf["year"] = new_year
                if new_active != active_now:
                    for k in conf["tournaments"]: conf["tournaments"][k]["active"] = (k == new_active)
                save_conf(conf); st.success("保存しました"); time.sleep(0.5); st.rerun()
        st.divider()
        with st.expander("参加人数制限の設定", expanded=True):
            with st.form("conf_limits"):
                lm = conf["limits"]
                c1, c2 = st.columns(2)
                lm["team_kata"]["min"] = c1.number_input("団体形 下限", 0, 10, lm["team_kata"]["min"])
                lm["team_kata"]["max"] = c2.number_input("団体形 上限", 0, 10, lm["team_kata"]["max"])
                c1, c2 = st.columns(2)
                lm["team_kumite_5"]["min"] = c1.number_input("団体組手(5人) 下限", 0, 10, lm["team_kumite_5"]["min"])
                lm["team_kumite_5"]["max"] = c2.number_input("団体組手(5人) 上限", 0, 10, lm["team_kumite_5"]["max"])
                st.caption("個人戦 (上限のみ)")
                c1, c2 = st.columns(2)
                lm["ind_kata_reg"]["max"] = c1.number_input("個人形(正) 上限", 0, 10, lm["ind_kata_reg"]["max"])
                lm["ind_kata_sub"]["max"] = c2.number_input("個人形(補) 上限", 0, 10, lm["ind_kata_sub"]["max"])
                if st.form_submit_button("人数制限を保存"):
                    conf["limits"] = lm; save_conf(conf); st.success("保存しました")
        with st.expander("⚙️ 詳細設定（ファイル名・階級など ※通常は変更不要）"):
            t_data = conf["tournaments"]["shinjin"]
            with st.form("edit_t_advanced"):
                wm_in = st.text_area("新人戦 男子階級リスト", t_data.get("weights_m", ""))
                ww_in = st.text_area("新人戦 女子階級リスト", t_data.get("weights_w", ""))
                if st.form_submit_button("詳細設定を保存"):
                    conf["tournaments"]["shinjin"]["weights_m"] = wm_in
                    conf["tournaments"]["shinjin"]["weights_w"] = ww_in
                    save_conf(conf); st.success("保存しました")

    elif admin_tab == "📥 データ出力":
        st.subheader("トーナメント・集計データの出力")
        tid = next((k for k, v in conf["tournaments"].items() if v["active"]), "kantou")
        st.info("ℹ️ サーバー負荷軽減のため、ファイルはボタンを押した時のみ作成されます。まず下の「集計を開始」ボタンを押してください。")
        if st.button("🔄 最新データで集計を開始"):
            with st.spinner("集計中..."):
                master = load_members_master(force_reload=True)
                entries = load_entries(tid, force_reload=True)
                full_data = []
                for _, m in master.iterrows():
                    uid = f"{m['school_id']}_{m['name']}"
                    ent = entries.get(uid, {})
                    if ent and (ent.get("kata_chk") or ent.get("kumi_chk")):
                        row = m.to_dict(); row.update(ent)
                        full_data.append(row)
                st.session_state["admin_xlsx_tour"] = generate_tournament_excel(full_data, conf["tournaments"][tid]["type"], auth)
                st.session_state["admin_xlsx_summ"] = generate_summary_excel(master, entries, auth, conf["tournaments"][tid]["type"])
                st.session_state["admin_xlsx_adv"] = generate_advisor_excel(load_schools(), auth)
                st.session_state["admin_data_ts"] = datetime.datetime.now().strftime("%Y/%m/%d %H:%M:%S")
        if "admin_data_ts" in st.session_state:
            st.success(f"✅ 集計完了 ({st.session_state['admin_data_ts']})")
            c1, c2, c3 = st.columns(3)
            c1.download_button("📥 トーナメント用データ", st.session_state["admin_xlsx_tour"], "v2_tournament.xlsx")
            c2.download_button("📊 参加校一覧集計", st.session_state["admin_xlsx_summ"], "v2_summary.xlsx")
            c3.download_button("👔 顧問リスト", st.session_state["admin_xlsx_adv"], "v2_advisors.xlsx")

    elif admin_tab == "🏫 アカウント(編集)":
        st.subheader("アカウント管理 (v2仕様)")
        st.caption("※「基本名」は申込書用、「略称」はトーナメント表用です。")
        recs = []
        for sid, d in auth.items():
            recs.append({
                "ID": sid, "基本名(申込書)": d.get("base_name",""), "略称(集計用)": d.get("short_name", d.get("base_name","")),
                "No": d.get("school_no", 999), "Password": d.get("password",""), "校長名": d.get("principal","")
            })
        edited = st.data_editor(pd.DataFrame(recs), disabled=["ID"], key="v2_auth_edit")
        if st.button("変更を保存"):
            has_error = False
            for _, row in edited.iterrows():
                sid = row["ID"]
                if len(str(row["Password"])) < 6: st.error(f"❌ {row['基本名(申込書)']} のパスワードが短すぎます"); has_error = True
                if sid in auth:
                    auth[sid]["base_name"] = row["基本名(申込書)"]
                    auth[sid]["short_name"] = row["略称(集計用)"]
                    auth[sid]["school_no"] = to_safe_int(row["No"])
                    auth[sid]["password"] = row["Password"]
                    auth[sid]["principal"] = row["校長名"]
            if not has_error: save_auth(auth); st.success("保存しました")
    
    elif admin_tab == "📅 年次処理":
        st.subheader("🌸 年度更新処理 (v2)")
        st.info("実行すると、学年を+1し、卒業生(新4年生)を「卒業生アーカイブ」へ移動します。")
        col_act1, col_act2 = st.columns(2)
        if col_act1.button("新年度を開始する (実行確認)"): res = perform_year_rollover(); st.success(res)
        st.markdown("---")
        st.subheader("🎓 卒業生データ管理")
        col_dl, col_del = st.columns(2)
        with col_dl:
            grad_df = get_graduates_df()
            if not grad_df.empty:
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer: grad_df.to_excel(writer, index=False)
                st.download_button("📥 卒業生名簿をダウンロード", output.getvalue(), "v2_graduates_archive.xlsx")
            else: st.caption("卒業生データはありません")
        with col_del:
            if not grad_df.empty:
                st.warning("⚠️ この操作は取り消せません")
                if st.button("🗑️ 卒業生データを全て削除する"):
                    clear_graduates_archive(); st.success("削除しました"); time.sleep(0.5); st.rerun()
        st.markdown("---")
        st.subheader("⏪ 復元 (Undo)")
        if st.button("バックアップから復元する"): res = restore_from_backup(); st.warning(res)

def main():
    st.set_page_config(page_title="Entry System v2", layout="wide")
    st.title("🥋 エントリーシステム v2 (Sandbox)")
    if "logged_in_school" in st.session_state:
        school_page(st.session_state["logged_in_school"]); return

    auth = load_auth()
    t1, t2, t3 = st.tabs(["ログイン", "新規登録(v2)", "管理者"])
    
    with t1:
        st.info("💡 初めての方は「新規登録(v2)」タブから登録を行ってください。")
        with st.form("login_form"):
            name_map = {f"{v.get('base_name')}高等学校": k for k, v in auth.items()}
            s_name = st.selectbox("学校名", list(name_map.keys()))
            pw = st.text_input("パスワード", type="password")
            if st.form_submit_button("ログイン"):
                if s_name:
                    sid = name_map[s_name]
                    if auth[sid]["password"] == pw:
                        st.session_state["logged_in_school"] = sid; st.rerun()
                    else: st.error("パスワードが違います")

    with t2:
        st.markdown("###### 新規登録")
        with st.form("register_form"):
            c1, c2 = st.columns([3, 1])
            base_name = c1.text_input("学校名 (「高等学校」は入力不要)")
            c2.markdown("<br>**高等学校**", unsafe_allow_html=True)
            p = st.text_input("校長名")
            new_pw = st.text_input("パスワード設定", type="password")
            st.caption("※パスワードは6文字以上で登録してください。")
            
            if st.form_submit_button("登録"):
                if base_name and new_pw:
                    if len(new_pw) < 6: st.error("パスワードは6文字以上にしてください")
                    else:
                        exists = any(v.get('base_name') == base_name for v in auth.values())
                        if exists: st.error("その学校名は既に登録されています")
                        else:
                            new_id = generate_school_id()
                            auth[new_id] = {
                                "base_name": base_name, "short_name": base_name, 
                                "password": new_pw, "principal": p, "school_no": 999, "advisors": []
                            }
                            save_auth(auth); st.success(f"登録完了! ID: {new_id}"); time.sleep(1); st.rerun()
                else: st.error("入力を確認してください")

    with t3: admin_page()

if __name__ == "__main__": main()