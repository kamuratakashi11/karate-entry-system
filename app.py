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
# ADMIN_PASSWORD は load_conf() で管理

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
            vals = st.secrets["gcp_key"]
            if isinstance(vals, str):
                key_dict = json.loads(vals)
            else:
                key_dict = vals
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
        if not val: return default
        parsed = json.loads(val)
        return parsed if parsed is not None else default
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
        if data is None: data = {}
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
# 4. ロジック
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
        val = entries.get(uid, {}).get(key, None)
        return val
    
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
            
            if ent.get("team_kata_chk") and ent.get("team_kata_role") == "正": cnt_tk += 1
            if ent.get("team_kumi_chk") and ent.get("team_kumi_role") == "正": cnt_tku += 1
            
            if ent.get("kata_chk"):
                k_val = ent.get("kata_val")
                if k_val == "補": cnt_ind_k_sub += 1
                elif k_val == "正": cnt_ind_k_reg += 1 
                
            if ent.get("kumi_chk"):
                val = ent.get("kumi_val", "")
                if val == "補": cnt_ind_ku_sub += 1
                elif val == "正": cnt_ind_ku_reg += 1
                elif t_type != "standard" and val and val != "出場しない" and val != "なし" and val != "シード" and val != "補":
                    cnt_ind_ku_reg += 1

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
        
        if cnt_ind_k_reg > limits["ind_kata_reg"]["max"]: errs.append(f"❌ {sex}個人形(正): 上限 {limits['ind_kata_reg']['max']}名を超えています。(シード除く)")
        if cnt_ind_k_sub > limits["ind_kata_sub"]["max"]: errs.append(f"❌ {sex}個人形(補): 上限 {limits['ind_kata_sub']['max']}名を超えています。")
        if cnt_ind_ku_reg > limits["ind_kumi_reg"]["max"]: errs.append(f"❌ {sex}個人組手(正): 上限 {limits['ind_kumi_reg']['max']}名を超えています。(シード除く)")
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
            elif t_conf["type"] == "standard": 
                if val == "シード": txt = f"シ{rank}"
                else: txt = f"○{rank}"
            else: txt = "○"
            safe_write(ws, (r, k_col), txt, True)

        if row.get("last_kumi_chk"):
            val = row.get("last_kumi_val")
            rank = row.get("last_kumi_rank", "")
            if val == "補": txt = "補"
            elif t_conf["type"] == "standard": 
                if val == "シード": txt = f"シ{rank}"
                else: txt = f"○{rank}"
            elif t_conf["type"] == "weight": txt = str(val)
            elif t_conf["type"] == "division": txt = str(val)
            else: txt = "○"
            safe_write(ws, (r, ku_col), txt, True)
    
    fname = f"申込書_{school_name}.xlsx"
    wb.save(fname)
    return fname, "作成成功"

# ---------------------------------------------------------
# 6. トーナメントデータ・集計表出力
# ---------------------------------------------------------
def generate_tournament_excel(all_data, t_type):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        sheets_data = {}

        for row in all_data:
            name = row['name']
            school = row['school']
            sex = row['sex']
            
            # --- 個人形 ---
            if row.get('kata_chk'):
                k_val = row.get('kata_val')
                k_rank = row.get('kata_rank', '')
                
                if k_val and k_val != '補' and k_val != 'なし' and k_val != '出場しない':
                    sheet_name = f"{sex}個人形"
                    rank_cell = k_rank if k_val == '正' else ''
                    seed_cell = k_rank if k_val == 'シード' else ''
                    
                    record = {
                        "個人形_順位": rank_cell,
                        "名前": name,
                        "学校名": school,
                        "シード順位": seed_cell
                    }
                    if sheet_name not in sheets_data: sheets_data[sheet_name] = []
                    sheets_data[sheet_name].append(record)

            # --- 個人組手 ---
            if row.get('kumi_chk'):
                ku_val = row.get('kumi_val')
                ku_rank = row.get('kumi_rank', '')
                
                if ku_val and ku_val != '補' and ku_val != 'なし' and ku_val != '出場しない':
                    if t_type == 'standard':
                        sheet_name = f"{sex}個人組手"
                        is_seed = (ku_val == 'シード')
                        is_reg = (ku_val == '正')
                    else:
                        sheet_name = f"{sex}個人組手_{ku_val}"
                        is_seed = False
                        is_reg = True 

                    rank_cell = ku_rank if is_reg else ''
                    seed_cell = ku_rank if is_seed else ''
                    
                    record = {
                        "個人組手_順位": rank_cell,
                        "名前": name,
                        "学校名": school,
                        "シード順位": seed_cell
                    }
                    if sheet_name not in sheets_data: sheets_data[sheet_name] = []
                    sheets_data[sheet_name].append(record)

        sorted_sheet_names = sorted(sheets_data.keys())
        for s_name in sorted_sheet_names:
            recs = sheets_data[s_name]
            header_rank = "個人組手_順位" if "組手" in s_name else "個人形_順位"
            df_out = pd.DataFrame(recs, columns=[header_rank, "名前", "学校名", "シード順位"])
            df_out.to_excel(writer, sheet_name=s_name, index=False)
            
    return output.getvalue()

def to_safe_int(val):
    try:
        s = to_half_width(str(val))
        return int(s)
    except:
        return 999

def generate_summary_excel(master_df, entries, auth_data, t_type):
    summary_rows = []
    
    # 学校番号順にソート(安全版)
    sorted_schools = sorted(auth_data.items(), key=lambda x: to_safe_int(x[1].get('school_no')))
    
    for s_name, s_auth in sorted_schools:
        s_no = s_auth.get('school_no', '')
        s_members = master_df[master_df['school'] == s_name]
        
        m_tk_flag = ""; m_tku_flag = ""
        w_tk_flag = ""; w_tku_flag = ""
        m_k_cnt = 0; m_ku_cnt = 0
        w_k_cnt = 0; w_ku_cnt = 0
        
        reg_player_names = set()
        
        for _, r in s_members.iterrows():
            uid = f"{s_name}_{r['name']}"
            ent = entries.get(uid, {})
            sex = r['sex']
            
            if sex == "男子":
                if ent.get("team_kata_chk"): m_tk_flag = "○"
                if ent.get("team_kumi_chk"): m_tku_flag = "○"
            else:
                if ent.get("team_kata_chk"): w_tk_flag = "○"
                if ent.get("team_kumi_chk"): w_tku_flag = "○"
            
            # 個人カウント (正 or シード or 階級) -> 補欠以外
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
            
            # 正選手合計計算
            is_reg = False
            if ent.get("team_kata_chk") and ent.get("team_kata_role") == "正": is_reg = True
            if ent.get("team_kumi_chk") and ent.get("team_kumi_role") == "正": is_reg = True
            kv = ent.get("kata_val")
            if ent.get("kata_chk") and kv and kv != "補" and kv != "なし" and kv != "出場しない": is_reg = True
            kuv = ent.get("kumi_val")
            if ent.get("kumi_chk") and kuv and kuv != "補" and kuv != "なし" and kuv != "出場しない": is_reg = True
            
            if is_reg:
                reg_player_names.add(r['name'])

        summary_rows.append({
            "学校No": s_no,
            "学校名": s_name,
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
    # 学校番号順にソート(安全版)
    sorted_schools = sorted(auth_data.items(), key=lambda x: to_safe_int(x[1].get('school_no')))
    
    cnt_judge = 0
    cnt_staff = 0
    
    for s_name, s_auth in sorted_schools:
        s_no = s_auth.get('school_no', '')
        s_info = schools_data.get(s_name, {})
        advs = s_info.get("advisors", [])
        
        for a in advs:
            name = a.get("name", "")
            if not name: continue
            role = a.get("role", "審判")
            d1 = "○" if a.get("d1") else "×"
            d2 = "○" if a.get("d2") else "×"
            
            if role == "審判": cnt_judge += 1
            if role == "係員": cnt_staff += 1
            
            rows.append({
                "No": s_no,
                "学校名": s_name,
                "顧問氏名": name,
                "役割": role,
                "1日目": d1,
                "2日目": d2
            })
            
    df_list = pd.DataFrame(rows)
    
    # 1シート化：リストの右側に集計を表示
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_list.to_excel(writer, sheet_name="顧問一覧", index=False, startcol=0)
        
        # 集計表を作成して同じシートのH列あたりに書き込む
        df_summary = pd.DataFrame([
            {"項目": "審判 合計", "人数": cnt_judge},
            {"項目": "係員 合計", "人数": cnt_staff}
        ])
        df_summary.to_excel(writer, sheet_name="顧問一覧", index=False, startcol=7) # H列=7
        
    return output.getvalue()

# ---------------------------------------------------------
# 7. UI
# ---------------------------------------------------------
def to_half_width(text):
    if not text: return ""
    return text.translate(str.maketrans('０１２３４５６７８９', '0123456789')).strip()

def school_page(s_name):
    st.markdown("""
    <style>
    div[data-testid="stRadio"] > div {
        flex-direction: row; 
    }
    </style>
    """, unsafe_allow_html=True)

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

    menu = ["① 大会エントリー", "② 部員名簿", "③ 顧問登録"]
    try: idx = menu.index(st.session_state["current_view"])
    except: idx = 0
    selected_view = st.radio("メニュー選択", menu, index=idx, horizontal=True, label_visibility="collapsed")
    st.session_state["current_view"] = selected_view
    st.markdown("---")

    if selected_view == "① 大会エントリー":
        target_grades = [int(g) for g in t_conf['grades']]
        st.markdown(f"**出場対象学年:** {target_grades} 年生  \n<small>※順位のところは、シード順位を入れてください。シードでない場合は出場選手の優先順位を入れてください。優先順位をもとにトーナメントは組まれます。</small>", unsafe_allow_html=True)
        
        merged = get_merged_data(s_name, active_tid)
        
        merged['sex_rank'] = merged['sex'].map({'男子': 0, '女子': 1})
        merged['grade_rank'] = merged['grade'].map({3: 0, 2: 1, 1: 2})
        
        valid_members = merged[merged['grade'].isin(target_grades)].sort_values(by=['sex_rank', 'grade_rank', 'name']).copy()
        
        if valid_members.empty: st.warning("部員名簿が空です。名簿タブから部員を登録してください。"); return
        
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
            cols = st.columns([2.0, 2.0, 2.0, 0.2, 2.2, 3.2])
            cols[0].markdown("**氏名**")
            cols[1].markdown("**団体形**")
            cols[2].markdown("**団体組手**")
            cols[4].markdown("**個人形**")
            cols[5].markdown("**個人組手**")

            form_buffer = {}

            for i, r in valid_members.iterrows():
                uid = f"{r['school']}_{r['name']}"
                name_style = 'background-color:#e8f5e9; color:#1b5e20; padding:2px 6px; border-radius:4px; font-weight:bold;' if r['sex'] == "男子" else 'background-color:#ffebee; color:#b71c1c; padding:2px 6px; border-radius:4px; font-weight:bold;'
                
                def_tk = r.get("last_team_kata_role", "なし")
                if not def_tk or def_tk not in ["正", "補"]: def_tk = "なし"
                
                def_tku = r.get("last_team_kumi_role", "なし")
                if not def_tku or def_tku not in ["正", "補"]: def_tku = "なし"
                
                def_k = r.get("last_kata_val", "なし")
                if not def_k or def_k not in ["正", "補", "シード"]: def_k = "なし"

                c = st.columns([2.0, 2.0, 2.0, 0.2, 2.2, 3.2])
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
                    if t_conf["type"] == "standard":
                        opts_k = ["なし", "シード", "正", "補"]
                    else:
                        opts_k = ["なし", "正", "補"]
                    idx_k = opts_k.index(def_k) if def_k in opts_k else 0
                    ck1, ck2 = c[4].columns([1.5, 1])
                    val_k = ck1.radio(f"k_{uid}", opts_k, index=idx_k, horizontal=True, key=f"rd_k_{uid}", label_visibility="collapsed")
                    rank_k = ck2.text_input("順位", r.get("last_kata_rank",""), key=f"rk_k_{uid}", label_visibility="collapsed", placeholder="順位")
                else:
                    val_k = "なし"; rank_k = ""; c[4].caption("-")
                
                c5a, c5b = c[5].columns([1.8, 1])
                w_key = "weights_m" if r['sex'] == "男子" else "weights_w"
                w_str = t_conf.get(w_key, "")
                w_list = ["出場しない"] + [f"{w.strip()}kg級" for w in w_str.split(",")] + ["補欠"]
                
                raw_kumi = r.get("last_kumi_val")
                if raw_kumi is None or pd.isna(raw_kumi):
                    def_val = "出場しない"
                else:
                    def_val = str(raw_kumi)
                
                if t_conf["type"] == "standard":
                    opts_ku = ["なし", "シード", "正", "補"]
                    if def_val == "出場しない": def_val = "なし"
                    if def_val not in opts_ku: def_val = "なし"
                    idx = opts_ku.index(def_val)
                    ku_val = c5a.radio(f"ku_{uid}", opts_ku, index=idx, horizontal=True, key=f"rd_ku_{uid}", label_visibility="collapsed")
                else:
                    if "kg" in def_val and t_conf["type"] == "standard": def_val = "出場しない"
                    elif t_conf["type"] == "weight" and def_val not in w_list and def_val != "補欠" and def_val != "出場しない": 
                        def_val = f"{def_val}kg級"
                    try: idx = w_list.index(def_val)
                    except: idx = 0
                    ku_val = c5a.selectbox("階級", w_list, index=idx, key=f"sel_ku_{uid}", label_visibility="collapsed")
                
                rank_ku = c5b.text_input("順位", r.get("last_kumi_rank",""), key=f"rk_ku_{uid}", label_visibility="collapsed", placeholder="順位")

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
                    tk_role = raw["val_tk"] if tk_chk else ""
                    
                    tku_chk = (raw["val_tku"] != "なし")
                    tku_role = raw["val_tku"] if tku_chk else ""
                    
                    k_chk = (raw["val_k"] != "なし")
                    k_role = raw["val_k"] if k_chk else ""
                    k_rank = to_half_width(raw["rank_k"])
                    
                    if t_conf["type"] == "standard":
                        ku_chk = (raw["ku_val"] != "なし")
                        ku_role = raw["ku_val"] if ku_chk else ""
                    else:
                        ku_chk = (raw["ku_val"] != "出場しない")
                        ku_role = raw["ku_val"] if ku_chk else ""
                    
                    ku_rank = to_half_width(raw["rank_ku"])

                    name = uid.split('_')[1]
                    if k_chk and k_role == "正":
                        if not k_rank: st.error(f"❌ {name}さん: 個人形の順位が入力されていません。"); has_error = True
                    elif not k_chk or k_role == "補": k_rank = ""

                    if ku_chk:
                        is_reg = (t_conf["type"] == "weight" and ku_role != "補欠") or \
                                 (t_conf["type"] == "standard" and ku_role == "正")
                        if is_reg and not ku_rank: st.error(f"❌ {name}さん: 個人組手の順位が入力されていません。"); has_error = True
                    
                    if not ku_chk or ku_role == "補": ku_rank = ""

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

    elif selected_view == "② 部員名簿":
        st.info("💡 ここは「全大会共通」の名簿です。")
        
        # 1. 新規追加フォーム
        with st.expander("👤 新しい部員を追加する", expanded=False):
            with st.form("add_member"):
                c = st.columns(3)
                nn = c[0].text_input("氏名")
                # 性別選択: 初期値を空欄に
                ns = c[1].selectbox("性別", ["", "男子", "女子"])
                ng = c[2].selectbox("学年", [1, 2, 3])
                c2 = st.columns(2)
                nd = c2[0].text_input("生年月日")
                nj = c2[1].text_input("JKF番号")
                
                if st.form_submit_button("追加"):
                    if not nn:
                        st.error("❌ 氏名を入力してください")
                    elif not ns:
                        st.error("❌ 性別を選択してください")
                    else:
                        if "master_cache" in st.session_state: del st.session_state["master_cache"]
                        master = load_members_master()
                        new_row = pd.DataFrame([{"school":s_name, "name":nn, "sex":ns, "grade":ng, "dob":nd, "jkf_no":nj, "active":True}])
                        save_members_master(pd.concat([master, new_row], ignore_index=True))
                        st.success(f"{nn} さんを追加しました"); st.rerun()

        st.divider()
        st.markdown("##### 📝 名簿編集 (修正・削除)")
        st.caption("※データを直接書き換えて「保存」ボタンを押してください。行を選んでDeleteキーで削除できます。")
        
        master = load_members_master()
        # この学校のデータだけ抽出
        my_m = master[master['school']==s_name].copy()
        
        # エディタで編集
        edited_df = st.data_editor(my_m[['name','sex','grade','dob','jkf_no','active']], num_rows="dynamic", use_container_width=True)
        
        if st.button("💾 修正を保存する", type="primary"):
            # 保存処理: 他校のデータ + 編集後の自校データ
            other_m = master[master['school']!=s_name]
            # edited_dfにschool列を付与して結合
            edited_df['school'] = s_name
            new_master = pd.concat([other_m, edited_df], ignore_index=True)
            
            save_members_master(new_master)
            st.success("✅ 名簿を更新しました"); time.sleep(1); st.rerun()

        st.divider()
        st.markdown("##### 📋 登録部員リスト (確認用)")
        
        # 編集後のデータをリロードして表示に使用
        view_m = master[master['school']==s_name]
        
        # 左右分割表示 (左:男子 / 右:女子)
        c_male, c_female = st.columns(2)
        
        with c_male:
            st.markdown("###### 🚹 男子部員")
            # 男子または性別不明なデータ (安全策)
            m_df = view_m[view_m['sex'] != '女子'].sort_values(by=['grade', 'name'], ascending=[False, True])
            if not m_df.empty:
                st.dataframe(m_df[['grade','name','jkf_no']], hide_index=True, use_container_width=True)
            else:
                st.caption("登録なし")
                
        with c_female:
            st.markdown("###### 🚺 女子部員")
            w_df = view_m[view_m['sex'] == '女子'].sort_values(by=['grade', 'name'], ascending=[False, True])
            if not w_df.empty:
                st.dataframe(w_df[['grade','name','jkf_no']], hide_index=True, use_container_width=True)
            else:
                st.caption("登録なし")

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

def admin_page():
    st.title("🔧 管理者画面")
    conf = load_conf()
    current_admin_pw = conf.get("admin_password", "1234")
    input_pw = st.text_input("Admin Password", type="password")
    if input_pw != current_admin_pw:
        return 

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
        st.subheader("トーナメントデータ出力")
        tid = next((k for k, v in conf["tournaments"].items() if v["active"]), "kantou")
        master = load_members_master(); entries = load_entries(tid)
        full_data = []
        for _, m in master.iterrows():
            uid = f"{m['school']}_{m['name']}"
            ent = entries.get(uid, {})
            if ent and (ent.get("kata_chk") or ent.get("kumi_chk")):
                row = m.to_dict(); row.update(ent)
                row["school_no"] = auth.get(m['school'], {}).get("school_no", 999)
                full_data.append(row)
        
        t_type = conf["tournaments"][tid]["type"]
        if st.button("📥 トーナメント用Excelをダウンロード"):
            if not full_data:
                st.warning("エントリーデータがありません")
            else:
                xlsx_data = generate_tournament_excel(full_data, t_type)
                st.download_button("Excelダウンロード開始", xlsx_data, "tournament_entries.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        st.divider()
        st.subheader("集計・運営資料出力")
        
        col_r1, col_r2 = st.columns(2)
        with col_r1:
            if st.button("📊 参加校一覧 (集計表)"):
                if "schools_data" not in st.session_state: st.session_state.schools_data = load_schools()
                xlsx = generate_summary_excel(master, entries, auth, t_type)
                st.download_button("集計表ダウンロード", xlsx, "summary_participation.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        with col_r2:
            if st.button("👔 顧問出欠リスト"):
                if "schools_data" not in st.session_state: st.session_state.schools_data = load_schools()
                xlsx = generate_advisor_excel(st.session_state.schools_data, auth)
                st.download_button("顧問リストダウンロード", xlsx, "summary_advisors.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    with t3:
        st.subheader("学校番号 & パスワード管理")
        with st.expander("🔑 管理者パスワード変更"):
            new_admin_pw = st.text_input("新しい管理者パスワード", type="password")
            if st.button("管理者パスワードを変更"):
                if len(new_admin_pw) < 6: st.error("6文字以上にしてください")
                else:
                    conf["admin_password"] = new_admin_pw
                    save_conf(conf); st.success("変更しました。次回から新しいパスワードを使用してください。")

        st.markdown("---")
        st.markdown("#### アカウント一覧・編集")
        st.caption("※学校名自体を書き換えると、システム上は「古い学校を削除して新しい学校を追加」した扱いになります。")
        
        s_list = []
        for k, v in auth.items():
            s_list.append({
                "学校名": k, 
                "No": v.get("school_no", 999), 
                "Password": v.get("password", ""),
                "校長名": v.get("principal", "") 
            })
            
        edf = st.data_editor(pd.DataFrame(s_list), key="sed", num_rows="fixed")
        
        if st.button("全データを保存 (修正を反映)"):
            new_auth = {}
            has_error = False
            for i, r in edf.iterrows():
                s_name = str(r["学校名"]).strip()
                if not s_name: continue
                if len(str(r["Password"])) < 6:
                    st.error(f"❌ {s_name} のパスワードが短すぎます (6文字以上)"); has_error = True
                new_auth[s_name] = {
                    "school_no": int(r["No"]),
                    "password": str(r["Password"]),
                    "principal": str(r["校長名"])
                }
            if not has_error:
                save_auth(new_auth)
                st.success("✅ 保存しました！学校名の変更も反映されました。")
                time.sleep(1); st.rerun()

        st.divider()
        with st.expander("🗑️ アカウント削除"):
            del_target = st.selectbox("削除する学校を選択", [""] + list(auth.keys()))
            if del_target:
                st.warning(f"⚠️ 本当に「{del_target}」を削除しますか？")
                if st.button(f"はい、{del_target} を削除します", type="primary"):
                    if del_target in auth:
                        del auth[del_target]
                        save_auth(auth)
                        st.success(f"{del_target} を削除しました。")
                        time.sleep(1); st.rerun()
            
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

def main():
    st.set_page_config(page_title="大会エントリー", layout="wide")
    qp = st.query_params
    if "school" in qp: st.session_state["logged_in_school"] = qp["school"]
    if "logged_in_school" in st.session_state:
        st.query_params["school"] = st.session_state["logged_in_school"]
        school_page(st.session_state["logged_in_school"]); return

    st.title("🥋埼玉県高体連空手道エントリーシステム"); auth = load_auth()
    t1, t2, t3 = st.tabs(["ログイン", "新規登録", "管理者"])
    with t1:
        s = st.selectbox("学校名", list(auth.keys()))
        pw = st.text_input("パスワード", type="password")
        if st.button("ログイン"):
            if s in auth and auth[s]["password"] == pw:
                st.session_state["logged_in_school"] = s; st.rerun()
            else: st.error("パスワードが違います")
        st.caption("※パスワードを忘れた場合は競技部長へ連絡をしてください。")
    with t2:
        n = st.text_input("学校名 (新規)"); p = st.text_input("校長名"); new_pw = st.text_input("パスワード (設定)", type="password")
        st.caption("※パスワードは6文字以上で登録してください。")
        if st.button("登録"):
            if n and new_pw:
                if len(new_pw) < 6: st.error("パスワードは6文字以上にしてください")
                else:
                    auth[n] = {"password": new_pw, "principal": p, "school_no": 999}
                    save_auth(auth); st.success("登録しました"); st.rerun()
    with t3: admin_page()

if __name__ == "__main__": main()