import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Alignment, PatternFill
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
# 1. 定数・初期設定・ヘルパー
# ---------------------------------------------------------
KEY_FILE = 'secrets.json'
SHEET_NAME = 'tournament_db' 
V2_PREFIX = "v2_" 

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

DEFAULT_TOURNAMENTS = {
    "kantou": {"name": "関東高等学校空手道大会 埼玉県予選", "template": "template.xlsx", "coords": "coords_standard.json", "type": "standard", "grades": [1, 2, 3], "active": True},
    "interhigh": {"name": "学校総合体育大会兼全国高等学校総合体育大会空手道競技県予選会", "template": "template.xlsx", "coords": "coords_standard.json", "type": "standard", "grades": [1, 2, 3], "active": False},
    "shinjin": {"name": "新人大会", "template": "template_shinjin.xlsx", "coords": "coords_shinjin.json", "type": "shinjin", "grades": [1, 2], "weights_m": "-55,-61,-68,-76,+76", "weights_w": "-48,-53,-59,-66,+66", "active": False},
    "senbatsu": {"name": "全国選抜 埼玉県予選", "template": "template_senbatsu.xlsx", "coords": "coords_senbatsu.json", "type": "weight", "grades": [1, 2], "weights_m": "選抜の部,一年生の部,高入生の部", "weights_w": "選抜の部,一年生の部,高入生の部", "active": False}
}

DEFAULT_LIMITS = {
    "team_kata": {"min": 3, "max": 3, "sub_max": 1},
    "team_kumite_5": {"min": 3, "max": 5, "sub_max": 2},
    "team_kumite_3": {"min": 2, "max": 3, "sub_max": 1},
    "ind_kata_reg": {"max": 4}, "ind_kata_sub": {"max": 2},
    "ind_kumi_reg": {"max": 4}, "ind_kumi_sub": {"max": 2}
}

# ---------------------------------------------------------
# 2. Google Sheets 接続
# ---------------------------------------------------------
@st.cache_resource
def get_gsheet_client():
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    if os.path.exists(KEY_FILE): creds = ServiceAccountCredentials.from_json_keyfile_name(KEY_FILE, scope)
    else:
        try:
            vals = st.secrets["gcp_key"]
            key_dict = json.loads(vals) if isinstance(vals, str) else vals
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
    except: st.error(f"スプレッドシート '{SHEET_NAME}' が見つかりません。"); st.stop()
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
        ws = get_worksheet_safe(target_tab); recs = ws.get_all_values()
        if not recs: return default
        if len(recs) == 1 and len(recs[0]) >= 1:
            val = str(recs[0][0])
            if val.startswith("{") or val.startswith("["): return json.loads(val) if json.loads(val) is not None else default
        result = {}
        for row in recs:
            if len(row) >= 2:
                key = row[0]; val_str = row[1]
                try: result[key] = json.loads(val_str)
                except: result[key] = val_str
        return result if result else default
    except: return default

def save_json(tab_name, data):
    target_tab = f"{V2_PREFIX}{tab_name}"; ws = get_worksheet_safe(target_tab)
    if not isinstance(data, dict):
        ws.clear(); ws.update_acell('A1', json.dumps(data, ensure_ascii=False)); return
    rows = [[str(k), json.dumps(v, ensure_ascii=False)] for k, v in data.items()]
    ws.clear()
    if rows: ws.update(rows)

def load_members_master(force_reload=False):
    if not force_reload and "v2_master_cache" in st.session_state: return st.session_state["v2_master_cache"]
    try:
        ws = get_worksheet_safe(f"{V2_PREFIX}members"); recs = ws.get_all_records()
        if not recs: df = pd.DataFrame(columns=MEMBERS_COLS)
        else:
             df = pd.DataFrame(recs)
             for c in MEMBERS_COLS:
                 if c not in df.columns: df[c] = ""
    except: return pd.DataFrame(columns=MEMBERS_COLS)
    df['grade'] = pd.to_numeric(df['grade'], errors='coerce').fillna(0).astype(int); df['jkf_no'] = df['jkf_no'].astype(str).replace('nan', ''); df['dob'] = df['dob'].astype(str).replace('nan', ''); df['display_order'] = df['display_order'].astype(str).replace('nan', '') 
    df = df[MEMBERS_COLS]; st.session_state["v2_master_cache"] = df; return df

def save_members_master(df):
    ws = get_worksheet_safe(f"{V2_PREFIX}members"); ws.clear()
    df = df.fillna(""); df['jkf_no'] = df['jkf_no'].astype(str); df['dob'] = df['dob'].astype(str); df['display_order'] = df['display_order'].astype(str) 
    for c in MEMBERS_COLS:
        if c not in df.columns: df[c] = ""
    df_to_save = df[MEMBERS_COLS]
    ws.update([df_to_save.columns.tolist()] + df_to_save.astype(str).values.tolist())
    st.session_state["v2_master_cache"] = df_to_save

def archive_graduates(grad_df, auth_data):
    if grad_df.empty: return
    ws_grad = get_worksheet_safe(f"{V2_PREFIX}graduates")
    grad_df = grad_df.copy(); grad_df["archived_school_name"] = grad_df["school_id"].apply(lambda sid: auth_data.get(sid, {}).get("base_name", "不明")); grad_df["archived_date"] = datetime.date.today().strftime("%Y-%m-%d")
    if not ws_grad.get_all_values(): ws_grad.append_row(grad_df.columns.tolist())
    ws_grad.append_rows(grad_df.astype(str).values.tolist())

def clear_graduates_archive(): get_worksheet_safe(f"{V2_PREFIX}graduates").clear()

def get_graduates_df():
    try:
        recs = get_worksheet_safe(f"{V2_PREFIX}graduates").get_all_records()
        return pd.DataFrame(recs) if recs else pd.DataFrame()
    except: return pd.DataFrame()

def load_entries(tournament_id, force_reload=False):
    key = f"v2_entry_cache_{tournament_id}"
    if not force_reload and key in st.session_state: return st.session_state[key]
    data = load_json(f"entry_{tournament_id}", {}); st.session_state[key] = data; return data

def save_entries(tournament_id, data):
    save_json(f"entry_{tournament_id}", data); st.session_state[f"v2_entry_cache_{tournament_id}"] = data

@st.cache_data
def load_auth_cached(): return load_json("auth", {})

def load_auth(): return load_auth_cached()

def save_auth(d): save_json("auth", d); load_auth_cached.clear()

def load_schools(): return load_json("schools", {})

def load_conf():
    default_conf = {"year": "6", "tournaments": DEFAULT_TOURNAMENTS, "limits": DEFAULT_LIMITS, "admin_password": "1234"}
    data = load_json("config", default_conf)
    if "limits" not in data: data["limits"] = DEFAULT_LIMITS
    if "tournaments" not in data: 
        data["tournaments"] = DEFAULT_TOURNAMENTS
    else:
        for t_key, default_t in DEFAULT_TOURNAMENTS.items():
            if t_key not in data["tournaments"]:
                data["tournaments"][t_key] = default_t
            else:
                for sub_k, sub_v in default_t.items():
                    if sub_k not in data["tournaments"][t_key]:
                        data["tournaments"][t_key][sub_k] = sub_v
    for k in ["team_kata", "team_kumite_5", "team_kumite_3"]:
        if k in data["limits"] and "sub_max" not in data["limits"][k]:
            data["limits"][k]["sub_max"] = DEFAULT_LIMITS[k]["sub_max"]
    return data

def save_conf(d): save_json("config", d)

def upload_file_to_gas(uploaded_file, school_name):
    if not GAS_WEBAPP_URL or GAS_WEBAPP_URL == "ここに貼り付け": return False, "GASのURLが設定されていません。"
    try:
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        ext = os.path.splitext(uploaded_file.name)[1]; file_name = f"【{school_name}】_申込書_{timestamp}{ext}"
        base64_content = base64.b64encode(uploaded_file.getvalue()).decode('utf-8')
        payload = {"fileName": file_name, "mimeType": uploaded_file.type, "base64": base64_content}
        res_data = requests.post(GAS_WEBAPP_URL, json=payload).json()
        if res_data.get("status") == "success": return True, res_data.get("id")
        else: return False, res_data.get("message", "不明なエラー")
    except Exception as e: return False, str(e)

# ---------------------------------------------------------
# 4. ロジック 
# ---------------------------------------------------------
def create_backup():
    df = load_members_master(force_reload=False)
    ws_bk_mem = get_worksheet_safe(f"{V2_PREFIX}members_backup"); ws_bk_mem.clear()
    df_bk = df.fillna("")[MEMBERS_COLS]; ws_bk_mem.update([df_bk.columns.tolist()] + df_bk.astype(str).values.tolist())
    conf = load_conf(); ws_bk_conf = get_worksheet_safe(f"{V2_PREFIX}config_backup"); ws_bk_conf.update_acell('A1', json.dumps(conf, ensure_ascii=False))

def restore_from_backup():
    try:
        ws_bk_mem = get_worksheet_safe(f"{V2_PREFIX}members_backup"); recs = ws_bk_mem.get_all_records()
        df = pd.DataFrame(recs) if recs else pd.DataFrame(columns=MEMBERS_COLS)
        if not df.empty: df['grade'] = pd.to_numeric(df['grade'], errors='coerce').fillna(0).astype(int); save_members_master(df)
    except: return "名簿の復元に失敗しました"
    try:
        ws_bk_conf = get_worksheet_safe(f"{V2_PREFIX}config_backup"); val = ws_bk_conf.acell('A1').value
        if val: save_conf(json.loads(val))
    except: return "設定の復元に失敗しました"
    return "✅ バックアップから復元しました"

def perform_year_rollover():
    create_backup()
    if "v2_master_cache" in st.session_state: del st.session_state["v2_master_cache"]
    df = load_members_master(force_reload=True)
    if df.empty: return "データがありません"
    df['grade'] = df['grade'] + 1
    graduates = df[df['grade'] > 3].copy(); current = df[df['grade'] <= 3].copy()
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
    cols_to_add = ["team_kata_chk", "team_kata_role", "team_kumi_chk", "team_kumi_role", "kata_chk", "kata_val", "kata_rank", "kumi_chk", "kumi_val", "kumi_rank"]
    for c in cols_to_add: my_members[f"last_{c}"] = my_members.apply(lambda r: entries.get(f"{r['school_id']}_{r['name']}", {}).get(c, None), axis=1)
    return my_members

def validate_counts(members_df, entries_data, limits, t_type, school_meta, school_id):
    errs = []
    for sex in ["男子", "女子"]:
        sex_df = members_df[members_df['sex'] == sex]
        cnt_tk = 0; cnt_tk_sub = 0; cnt_tku = 0; cnt_tku_sub = 0; cnt_ind_k_reg = 0; cnt_ind_k_sub = 0; cnt_ind_ku_reg = 0; cnt_ind_ku_sub = 0
        for _, r in sex_df.iterrows():
            uid = f"{school_id}_{r['name']}"; ent = entries_data.get(uid, {})
            if ent.get("team_kata_chk"):
                if ent.get("team_kata_role") == "正": cnt_tk += 1
                elif ent.get("team_kata_role") == "補": cnt_tk_sub += 1
            if ent.get("team_kumi_chk"):
                if ent.get("team_kumi_role") == "正": cnt_tku += 1
                elif ent.get("team_kumi_role") == "補": cnt_tku_sub += 1
            if ent.get("kata_chk"):
                if ent.get("kata_val") == "補": cnt_ind_k_sub += 1
                elif ent.get("kata_val") == "正": cnt_ind_k_reg += 1 
            if ent.get("kumi_chk"):
                v = ent.get("kumi_val")
                if v == "補": cnt_ind_ku_sub += 1
                elif v == "正": cnt_ind_ku_reg += 1
                elif t_type != "standard" and v and v not in ["出場しない", "なし", "シード", "補"]: cnt_ind_ku_reg += 1
        if cnt_tk > 0 or cnt_tk_sub > 0:
            mn, mx = limits["team_kata"]["min"], limits["team_kata"]["max"]
            if not (mn <= cnt_tk <= mx): errs.append(f"❌ {sex}団体形: 正選手は {mn}～{mx}名で登録してください。(現在{cnt_tk}名)")
            s_mx = limits["team_kata"].get("sub_max", 1)
            if cnt_tk_sub > s_mx: errs.append(f"❌ {sex}団体形: 補欠は上限 {s_mx}名までです。(現在{cnt_tk_sub}名)")
        if cnt_tku > 0 or cnt_tku_sub > 0:
            mode = school_meta.get("m_kumite_mode" if sex == "男子" else "w_kumite_mode", "none") if t_type == "shinjin" else "5"
            l_key = "team_kumite_3" if mode == "3" else "team_kumite_5"; mn, mx = limits[l_key]["min"], limits[l_key]["max"]
            if not (mn <= cnt_tku <= mx): errs.append(f"❌ {sex}団体組手({mode}人制): 正選手は {mn}～{mx}名で登録してください。(現在{cnt_tku}名)")
            s_mx = limits[l_key].get("sub_max", 2)
            if cnt_tku_sub > s_mx: errs.append(f"❌ {sex}団体組手({mode}人制): 補欠は上限 {s_mx}名までです。(現在{cnt_tku_sub}名)")
        if cnt_ind_k_reg > limits["ind_kata_reg"]["max"]: errs.append(f"❌ {sex}個人形(正): 上限 {limits['ind_kata_reg']['max']}名を超えています。(シード除く)")
        if cnt_ind_k_sub > limits["ind_kata_sub"]["max"]: errs.append(f"❌ {sex}個人形(補): 上限 {limits['ind_kata_sub']['max']}名を超えています。")
        if cnt_ind_ku_reg > limits["ind_kumi_reg"]["max"]: errs.append(f"❌ {sex}個人組手(正): 上限 {limits['ind_kumi_reg']['max']}名を超えています。(シード除く)")
        if cnt_ind_ku_sub > limits["ind_kumi_sub"]["max"]: errs.append(f"❌ {sex}個人組手(補): 上限 {limits['ind_kumi_sub']['max']}名を超えています。")
    return errs

# ---------------------------------------------------------
# 5. Excel生成
# ---------------------------------------------------------
def safe_write(ws, target, value, align_center=False):
    if not target: return
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
    template_file = t_conf.get("template", "template.xlsx")
    coords_file = t_conf.get("coords", "coords_standard.json")
    try:
        with open(coords_file, "r", encoding="utf-8") as _f:
            coords = json.load(_f)
    except: return None, f"{coords_file} が見つかりません。"
    try: wb = openpyxl.load_workbook(template_file); ws = wb.active
    except: return None, f"{template_file} が見つかりません。"
    conf = load_conf(); entries = load_entries(t_id); school_meta = entries.get(f"_meta_{school_id}", {}); m_mode = school_meta.get("m_kumite_mode", "5"); w_mode = school_meta.get("w_kumite_mode", "5"); safe_write(ws, coords["year"], conf.get("year", "")); safe_write(ws, coords["tournament_name"], t_conf.get("name", ""))
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
        sex = row["sex"]
        
        tk_c = cols.get(f"m_team_kata" if sex=="男子" else f"w_team_kata")
        if tk_c and row.get("last_team_kata_chk"): 
            safe_write(ws, (r, tk_c), "補" if row.get("last_team_kata_role")=="補" else "○", True)
            
        if row.get("last_team_kumi_chk"):
            if t_conf["type"] in ["shinjin", "weight"]:
                mode = m_mode if sex=="男子" else w_mode
                tku_c = cols.get(f"m_team_kumite_{mode}" if sex=="男子" else f"w_team_kumite_{mode}")
            else:
                tku_c = cols.get(f"m_team_kumite" if sex=="男子" else f"w_team_kumite")
            if tku_c: safe_write(ws, (r, tku_c), "補" if row.get("last_team_kumi_role")=="補" else "○", True)

        k_c = cols.get(f"m_kata" if sex=="男子" else f"w_kata")
        if k_c and row.get("last_kata_chk"):
            v, rk = row.get("last_kata_val"), row.get("last_kata_rank", "")
            txt = "補" if v=="補" else (f"シ{rk}" if v=="シード" else f"○{rk}")
            safe_write(ws, (r, k_c), txt, True)

        if row.get("last_kumi_chk"):
            v, rk = row.get("last_kumi_val"), row.get("last_kumi_rank", "")
            if t_conf["type"] in ["shinjin", "weight"]:
                ku_c = cols.get(f"m_kumite_{v}" if sex=="男子" else f"w_kumite_{v}")
            else:
                ku_c = cols.get(f"m_kumite" if sex=="男子" else f"w_kumite")
            
            if ku_c:
                sub_v = row.get("last_kumi_sub_val", v) # 階級制なら sub_val、標準なら val
                if sub_v == "補": txt = "補"
                else: txt = f"シ{rk}" if sub_v=="シード" else f"○{rk}"
                safe_write(ws, (r, ku_c), txt, True)
            else:
                weight = row.get("last_kumi_val")
                sub_v = row.get("last_kumi_sub_val", "正")
                rk = row.get("last_kumi_rank", "")
                txt = "補" if sub_v=="補" else (f"シ{rk}" if sub_v=="シード" else f"○{rk}")
                key = f"m_kumite_{weight}" if sex=="男子" else f"w_kumite_{weight}"
                ku_c = cols.get(key)
                if ku_c:
                    safe_write(ws, (r, ku_c), txt, True)

    import io
    output = io.BytesIO()
    fname = f"申込書_{bn}.xlsx"
    wb.save(output)
    output.seek(0)
    return output, fname

def generate_tournament_excel(all_data, t_type, auth_data):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        sheets_data = {}
        for row in all_data:
            name, sid, sex = row['name'], row['school_id'], row['sex']
            s_data = auth_data.get(sid, {}); school_short = s_data.get("short_name", s_data.get("base_name", ""))
            if row.get('kata_chk') and row.get('kata_val') not in ['補', 'なし', '出場しない']:
                sn = f"{sex}個人形"; rec = {"個人形_順位": row.get('kata_rank','') if row.get('kata_val')=='正' else '', "名前": name, "学校名": school_short, "シード順位": row.get('kata_rank','') if row.get('kata_val')=='シード' else ''}
                if sn not in sheets_data: sheets_data[sn] = []
                sheets_data[sn].append(rec)
            if row.get('kumi_chk') and row.get('kumi_val') not in ['補', 'なし', '出場しない']:
                sn = f"{sex}個人組手" if t_type=='standard' else f"{sex}個人組手_{row.get('kumi_val')}"
                rec = {"個人組手_順位": row.get('kumi_rank','') if row.get('kumi_val')=='正' or t_type!='standard' else '', "名前": name, "学校名": school_short, "シード順位": row.get('kumi_rank','') if row.get('kumi_val')=='シード' else ''}
                if sn not in sheets_data: sheets_data[sn] = []
                sheets_data[sn].append(rec)
        for sn in sorted(sheets_data.keys()):
            pd.DataFrame(sheets_data[sn]).to_excel(writer, sheet_name=sn, index=False)
    return output.getvalue()

def generate_summary_excel(master_df, entries, auth_data, t_type):
    rows = []
    flags = {}
    for s_id, s_data in sorted(auth_data.items(), key=lambda x: to_safe_int(x[1].get('school_no'))):
        s_name = s_data.get("short_name", s_data.get("base_name", ""))
        s_members = master_df[master_df['school_id'] == s_id]
        meta = entries.get(f"_meta_{s_id}", {})
        m_tk, m_tku, w_tk, w_tku, m_k, m_ku, w_k, w_ku = "", "", "", "", 0, 0, 0, 0
        regs = set()
        for _, r in s_members.iterrows():
            ent = entries.get(f"{s_id}_{r['name']}", {})
            sex = r['sex']
            if sex == "男子":
                if ent.get("team_kata_chk"): m_tk = "○"
                if ent.get("team_kumi_chk"): m_tku = "○"
                if ent.get("kata_chk") and ent.get("kata_val") not in ["補","なし","出場しない"]: m_k += 1
                if ent.get("kumi_chk") and ent.get("kumi_val") not in ["補","なし","出場しない"]: m_ku += 1
            else:
                if ent.get("team_kata_chk"): w_tk = "○"
                if ent.get("team_kumi_chk"): w_tku = "○"
                if ent.get("kata_chk") and ent.get("kata_val") not in ["補","なし","出場しない"]: w_k += 1
                if ent.get("kumi_chk") and ent.get("kumi_val") not in ["補","なし","出場しない"]: w_ku += 1
            if (ent.get("team_kata_chk") and ent.get("team_kata_role")=="正") or \
               (ent.get("team_kumi_chk") and ent.get("team_kumi_role")=="正") or \
               (ent.get("kata_chk") and ent.get("kata_val") not in ["補","なし","出場しない"]) or \
               (ent.get("kumi_chk") and ent.get("kumi_val") not in ["補","なし","出場しない"]):
                regs.add(r['name'])
        
        row_dict = {
            "学校No": s_data.get('school_no',''), "学校名": s_name,
            "男団体形": m_tk, "男団体組手": m_tku, "男個人形": m_k if m_k>0 else "", "男個人組手": m_ku if m_ku>0 else "",
            "女団体形": w_tk, "女団体組手": w_tku, "女個人形": w_k if w_k>0 else "", "女個人組手": w_ku if w_ku>0 else "",
            "正選手合計": len(regs)
        }
        # 不参加フラグの適用
        if not meta.get("part_m_tk", True): row_dict["男団体形"] = "GRAY"
        if not meta.get("part_m_tku", True): row_dict["男団体組手"] = "GRAY"
        if not meta.get("part_m_k", True): row_dict["男個人形"] = "GRAY"
        if not meta.get("part_m_ku", True): row_dict["男個人組手"] = "GRAY"
        if not meta.get("part_w_tk", True): row_dict["女団体形"] = "GRAY"
        if not meta.get("part_w_tku", True): row_dict["女団体組手"] = "GRAY"
        if not meta.get("part_w_k", True): row_dict["女個人形"] = "GRAY"
        if not meta.get("part_w_ku", True): row_dict["女個人組手"] = "GRAY"
        rows.append(row_dict)

    import pandas as pd
    import io
    df = pd.DataFrame(rows)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name="参加校一覧", index=False)
        ws = writer.sheets["参加校一覧"]
        gray_fill = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid")
        for row in ws.iter_rows(min_row=2, max_col=10):
            for cell in row:
                if cell.value == "GRAY":
                    cell.value = ""
                    cell.fill = gray_fill
    return output.getvalue()

def generate_advisor_excel(schools_data, auth_data):
    rows = []
    for s_id, s_auth in sorted(auth_data.items(), key=lambda x: to_safe_int(x[1].get('school_no'))):
        s_name = s_auth.get("short_name", s_auth.get("base_name", ""))
        for a in s_auth.get("advisors", []):
            if a.get("name"): rows.append({"No": s_auth.get('school_no',''), "学校名": s_name, "顧問氏名": a["name"], "役割": a.get("role","審判"), "1日目": "○" if a.get("d1") else "×", "2日目": "○" if a.get("d2") else "×"})
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer: pd.DataFrame(rows).to_excel(writer, sheet_name="顧問一覧", index=False)
    return output.getvalue()

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
    t_conf = conf["tournaments"][active_tid]; st.markdown(f"## 🥋 **令和{conf.get('year','〇')}年度 {t_conf['name']}** <small>エントリー画面</small>", unsafe_allow_html=True)
    
    if st.button("🔄 データを最新にする"):
        if "v2_master_cache" in st.session_state: del st.session_state["v2_master_cache"]
        if f"v2_entry_cache_{active_tid}" in st.session_state: del st.session_state[f"v2_entry_cache_{active_tid}"]
        st.success("最新データを読み込みました"); time.sleep(0.5); st.rerun()

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
        edited_adv_df = st.data_editor(
    adv_df[["name", "role", "d1", "d2"]], 
    column_config={
        "name": "氏名", 
        "role": st.column_config.SelectboxColumn("役割", options=["審判", "競技記録", "係員"], required=True), 
        "d1": st.column_config.CheckboxColumn("1日目", default=False),  # ここを明示的にチェックボックス化
        "d2": st.column_config.CheckboxColumn("2日目", default=False)   # ここを明示的にチェックボックス化
    }, 
    num_rows="dynamic", 
    use_container_width=True, 
    hide_index=True
)
        if st.button("💾 顧問情報を保存", type="primary"):
            if edited_adv_df["name"].isnull().any() or (edited_adv_df["name"] == "").any(): st.error("❌ 氏名が未入力です"); return
            with st.spinner("保存中..."):
                load_auth_cached.clear(); latest_auth = load_auth(); cur_s = latest_auth.get(s_id, s_data)
                cur_s["principal"], cur_s["advisors"] = np, edited_adv_df.to_dict(orient="records")
                latest_auth[s_id] = cur_s; save_auth(latest_auth); st.success("✅ 保存完了"); time.sleep(1); st.rerun()

    elif selected_view == "② 部員名簿登録":
        st.warning("⚠️ **重要:** 編集内容は自動保存されません。変更後は必ず下の **『💾 名簿を保存して更新』** ボタンを押してください。")
        st.caption("💡 **表示順について:** 「No.」列に数字を入力すると、その順番に表示されます。指定がない場合は学年・名前順になります。")
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
        st.divider()
        st.markdown("##### 📋 登録済み部員リスト（確認用）")
        master_check = load_members_master(force_reload=False)
        my_check = master_check[master_check['school_id']==s_id]
        def sort_key(row):
            try: return float(row['display_order']) if pd.notna(row['display_order']) and str(row['display_order']).strip() else 999999.0
            except: return 999999.0
        my_check['sort_k'] = my_check.apply(sort_key, axis=1)
        c_male, c_female = st.columns(2)
        with c_male:
            st.markdown("###### 🚹 男子部員")
            m_df = my_check[my_check['sex'] == '男子'].sort_values(by=['sort_k', 'grade', 'name'], ascending=[True, False, True])
            if not m_df.empty: st.dataframe(m_df[['display_order', 'grade','name','jkf_no']].rename(columns={'display_order':'No.','grade':'学年','name':'氏名','jkf_no':'JKF番号'}), hide_index=True, use_container_width=True)
            else: st.caption("登録なし")
        with c_female:
            st.markdown("###### 🚺 女子部員")
            w_df = my_check[my_check['sex'] == '女子'].sort_values(by=['sort_k', 'grade', 'name'], ascending=[True, False, True])
            if not w_df.empty: st.dataframe(w_df[['display_order', 'grade','name','jkf_no']].rename(columns={'display_order':'No.','grade':'学年','name':'氏名','jkf_no':'JKF番号'}), hide_index=True, use_container_width=True)
            else: st.caption("登録なし")

    elif selected_view == "③ 大会エントリー":
        target_grades = [int(g) for g in t_conf['grades']]
        st.markdown(f"**出場対象学年:** {target_grades} 年生")

        merged = get_merged_data(s_id, active_tid)
        if merged.empty: st.warning("名簿を登録してください。"); return
        valid_members = merged[merged['grade'].isin(target_grades)].copy()
        
        # 男子、女子の名前リストを作成
        m_names = valid_members[valid_members['sex'] == '男子']['name'].tolist()
        w_names = valid_members[valid_members['sex'] == '女子']['name'].tolist()
        
        entries_update = load_entries(active_tid, force_reload=False)
        school_meta = entries_update.get(f"_meta_{s_id}", {"m_kumite_mode": "none", "w_kumite_mode": "none"})
        
        m_mode, w_mode = "5", "5"
        st.info("💡 個人戦の「順位」には、シードの選手はシード順位、そうでない場合は優先順位を必ず入れてください。優先順位を見てそれぞれの学校の１と４、２と３がうまく当たるようにトーナメント表の位置決めをします。")
        st.markdown("#### ⚙️ 参加種目・階級の事前設定")
        st.info("💡 まずはじめに、出場する種目・階級をすべて選択し、下のボタンで確定してください。選んだ種目だけが下のエントリー表に表示されます。")
        with st.form("pre_setting_form_all"):
            c1, c2 = st.columns(2)
            sel_res = {}
            with c1:
                st.markdown("##### 🚹 男子")
                sel_res["part_m_tk"] = st.checkbox("団体形", value=school_meta.get("part_m_tk", False), key="pre_m_tk")
                sel_res["part_m_k"] = st.checkbox("個人形", value=school_meta.get("part_m_k", False), key="pre_m_k")
                
                if t_conf["type"] == "shinjin" or t_conf["type"] == "weight":
                    sel_res["m_kumi_5"] = st.checkbox("団体組手 (5人制)", value=(school_meta.get("m_kumite_mode")=="5"), key="pre_m_kumi_5")
                    sel_res["m_kumi_3"] = st.checkbox("団体組手 (3人制)", value=(school_meta.get("m_kumite_mode")=="3"), key="pre_m_kumi_3")
                    
                    st.write("個人組手 階級")
                    wm_list = [f"{w.strip()}kg級" for w in t_conf.get("weights_m", "-55,-61,-68,-76,+76").split(",")]
                    cols_m = st.columns(len(wm_list))
                    for i, w in enumerate(wm_list):
                        sel_res[f"part_m_ku_{w}"] = cols_m[i].checkbox(w.replace("kg級", ""), value=school_meta.get(f"part_m_ku_{w}", False), key=f"pre_m_ku_{w}")
                else:
                    sel_res["part_m_tku"] = st.checkbox("団体組手", value=school_meta.get("part_m_tku", False), key="pre_m_tku")
                    sel_res["part_m_ku"] = st.checkbox("個人組手", value=school_meta.get("part_m_ku", False), key="pre_m_ku")

            with c2:
                st.markdown("##### 🚺 女子")
                sel_res["part_w_tk"] = st.checkbox("団体形", value=school_meta.get("part_w_tk", False), key="pre_w_tk")
                sel_res["part_w_k"] = st.checkbox("個人形", value=school_meta.get("part_w_k", False), key="pre_w_k")
                
                if t_conf["type"] == "shinjin" or t_conf["type"] == "weight":
                    sel_res["w_kumi_5"] = st.checkbox("団体組手 (5人制)", value=(school_meta.get("w_kumite_mode")=="5"), key="pre_w_kumi_5")
                    sel_res["w_kumi_3"] = st.checkbox("団体組手 (3人制)", value=(school_meta.get("w_kumite_mode")=="3"), key="pre_w_kumi_3")
                    
                    st.write("個人組手 階級")
                    ww_list = [f"{w.strip()}kg級" for w in t_conf.get("weights_w", "-48,-53,-59,-66,+66").split(",")]
                    cols_w = st.columns(len(ww_list))
                    for i, w in enumerate(ww_list):
                        sel_res[f"part_w_ku_{w}"] = cols_w[i].checkbox(w.replace("kg級", ""), value=school_meta.get(f"part_w_ku_{w}", False), key=f"pre_w_ku_{w}")
                else:
                    sel_res["part_w_tku"] = st.checkbox("団体組手", value=school_meta.get("part_w_tku", False), key="pre_w_tku")
                    sel_res["part_w_ku"] = st.checkbox("個人組手", value=school_meta.get("part_w_ku", False), key="pre_w_ku")

            st.markdown("<br>", unsafe_allow_html=True)
            if st.form_submit_button("✅ 事前設定を確定する", type="primary"):
                has_err = False
                if t_conf["type"] == "shinjin" or t_conf["type"] == "weight":
                    if sel_res["m_kumi_5"] and sel_res["m_kumi_3"]:
                        st.error("❌ 男子団体組手の5人制と3人制はどちらか一方しか出場できません。")
                        has_err = True
                    if sel_res["w_kumi_5"] and sel_res["w_kumi_3"]:
                        st.error("❌ 女子団体組手の5人制と3人制はどちらか一方しか出場できません。")
                        has_err = True
                
                if not has_err:
                    school_meta["pre_setting_done"] = True
                    school_meta["part_m_tk"] = sel_res["part_m_tk"]
                    school_meta["part_m_k"] = sel_res["part_m_k"]
                    school_meta["part_w_tk"] = sel_res["part_w_tk"]
                    school_meta["part_w_k"] = sel_res["part_w_k"]
                    
                    if t_conf["type"] == "shinjin" or t_conf["type"] == "weight":
                        school_meta["m_kumite_mode"] = "5" if sel_res["m_kumi_5"] else ("3" if sel_res["m_kumi_3"] else "none")
                        school_meta["w_kumite_mode"] = "5" if sel_res["w_kumi_5"] else ("3" if sel_res["w_kumi_3"] else "none")
                        school_meta["part_m_tku"] = school_meta["m_kumite_mode"] != "none"
                        school_meta["part_w_tku"] = school_meta["w_kumite_mode"] != "none"
                        
                        m_ku_any, w_ku_any = False, False
                        for w in wm_list:
                            school_meta[f"part_m_ku_{w}"] = sel_res[f"part_m_ku_{w}"]
                            if sel_res[f"part_m_ku_{w}"]: m_ku_any = True
                        for w in ww_list:
                            school_meta[f"part_w_ku_{w}"] = sel_res[f"part_w_ku_{w}"]
                            if sel_res[f"part_w_ku_{w}"]: w_ku_any = True
                        school_meta["part_m_ku"] = m_ku_any
                        school_meta["part_w_ku"] = w_ku_any
                    else:
                        school_meta["part_m_tku"] = sel_res["part_m_tku"]
                        school_meta["part_m_ku"] = sel_res["part_m_ku"]
                        school_meta["part_w_tku"] = sel_res["part_w_tku"]
                        school_meta["part_w_ku"] = sel_res["part_w_ku"]
                    
                    entries_update[f"_meta_{s_id}"] = school_meta
                    save_entries(active_tid, entries_update)
                    st.rerun()

        m_mode = school_meta.get("m_kumite_mode", "none")
        w_mode = school_meta.get("w_kumite_mode", "none")
        p_m_tk = school_meta.get("part_m_tk", False)
        p_m_tku = school_meta.get("part_m_tku", False)
        p_m_k = school_meta.get("part_m_k", False)
        p_m_ku = school_meta.get("part_m_ku", False)
        p_w_tk = school_meta.get("part_w_tk", False)
        p_w_tku = school_meta.get("part_w_tku", False)
        p_w_k = school_meta.get("part_w_k", False)
        p_w_ku = school_meta.get("part_w_ku", False)

        st.markdown("### エントリー入力")
        with st.form("entry_form_unified"):
            tabs = st.tabs(["🥋 団体形", "🥊 団体組手", "🥋 個人形", "🥊 個人組手"])
            limits = conf["limits"]
            results = {}

            # メタデータの取得
            p_m_tk = school_meta.get("part_m_tk", True)
            p_m_tku = school_meta.get("part_m_tku", True)
            p_m_k = school_meta.get("part_m_k", True)
            p_m_ku = school_meta.get("part_m_ku", True)
            p_w_tk = school_meta.get("part_w_tk", True)
            p_w_tku = school_meta.get("part_w_tku", True)
            p_w_k = school_meta.get("part_w_k", True)
            p_w_ku = school_meta.get("part_w_ku", True)

            def build_team_df(sex_names, role_key, chk_key, max_reg, max_sub):
                reg_l, sub_l = [], []
                for n in sex_names:
                    ent = entries_update.get(f"{s_id}_{n}", {})
                    if ent.get(chk_key):
                        r = ent.get(role_key)
                        if r == "正": reg_l.append(n)
                        elif r == "補": sub_l.append(n)
                reg_l = (reg_l + [""] * max_reg)[:max_reg]
                sub_l = (sub_l + [""] * max_sub)[:max_sub]
                return pd.DataFrame({"役割": ["正"] * max_reg + ["補"] * max_sub, "選手名": reg_l + sub_l})

            def build_ind_df(sex_names, chk_key, val_key, rank_key, def_reg, def_sub):
                rows = []
                for n in sex_names:
                    ent = entries_update.get(f"{s_id}_{n}", {})
                    if ent.get(chk_key):
                        rows.append({"区分": ent.get(val_key, ""), "順位": ent.get(rank_key, ""), "選手名": n})
                cur_reg = sum(1 for r in rows if r["区分"] == "正")
                cur_sub = sum(1 for r in rows if r["区分"] == "補")
                for _ in range(max(0, def_reg - cur_reg)): rows.append({"区分": "正", "順位": "", "選手名": ""})
                for _ in range(max(0, def_sub - cur_sub)): rows.append({"区分": "補", "順位": "", "選手名": ""})
                if sum(1 for r in rows if r["区分"] == "シード") == 0: rows.append({"区分": "シード", "順位": "", "選手名": ""})
                return pd.DataFrame(rows)

            def build_weight_df(sex_names, chk_key, val_key, rank_key, sub_key, w_name):
                rows = []
                for n in sex_names:
                    ent = entries_update.get(f"{s_id}_{n}", {})
                    if ent.get(chk_key) and ent.get(val_key) == w_name:
                        rows.append({"区分": ent.get(sub_key, "正"), "順位": ent.get(rank_key, ""), "選手名": n})
                for _ in range(max(0, 5 - len(rows))):
                    rows.append({"区分": "正", "順位": "", "選手名": ""})
                return pd.DataFrame(rows)

            # ガイドテキスト用定数
            guide_txt = "💡 事前設定で参加が選択されていないため、エントリーできません。"

            # 1. 団体形
            with tabs[0]:
                c1, c2 = st.columns(2)
                with c1:
                    st.markdown("#### 🚹 男子 団体形")
                    if p_m_tk:
                        df_m_tk = build_team_df(m_names, "team_kata_role", "team_kata_chk", limits["team_kata"]["max"], limits["team_kata"]["sub_max"])
                        results["m_tk"] = st.data_editor(df_m_tk, column_config={"役割": st.column_config.Column(disabled=True), "選手名": st.column_config.SelectboxColumn(options=[""] + m_names)}, hide_index=True, key="ed_m_tk", use_container_width=True)
                    else: st.info(guide_txt)
                with c2:
                    st.markdown("#### 🚺 女子 団体形")
                    if p_w_tk:
                        df_w_tk = build_team_df(w_names, "team_kata_role", "team_kata_chk", limits["team_kata"]["max"], limits["team_kata"]["sub_max"])
                        results["w_tk"] = st.data_editor(df_w_tk, column_config={"役割": st.column_config.Column(disabled=True), "選手名": st.column_config.SelectboxColumn(options=[""] + w_names)}, hide_index=True, key="ed_w_tk", use_container_width=True)
                    else: st.info(guide_txt)

            # 2. 団体組手
            with tabs[1]:
                c1, c2 = st.columns(2)
                with c1:
                    st.markdown("#### 🚹 男子 団体組手")
                    if p_m_tku:
                        l_k = "team_kumite_3" if m_mode == "3" else "team_kumite_5"
                        df_m_tku = build_team_df(m_names, "team_kumi_role", "team_kumi_chk", limits[l_k]["max"], limits[l_k]["sub_max"])
                        results["m_tku"] = st.data_editor(df_m_tku, column_config={"役割": st.column_config.Column(disabled=True), "選手名": st.column_config.SelectboxColumn(options=[""] + m_names)}, hide_index=True, key="ed_m_tku", use_container_width=True)
                    else: st.info(guide_txt)
                with c2:
                    st.markdown("#### 🚺 女子 団体組手")
                    if p_w_tku:
                        l_k = "team_kumite_3" if w_mode == "3" else "team_kumite_5"
                        df_w_tku = build_team_df(w_names, "team_kumi_role", "team_kumi_chk", limits[l_k]["max"], limits[l_k]["sub_max"])
                        results["w_tku"] = st.data_editor(df_w_tku, column_config={"役割": st.column_config.Column(disabled=True), "選手名": st.column_config.SelectboxColumn(options=[""] + w_names)}, hide_index=True, key="ed_w_tku", use_container_width=True)
                    else: st.info(guide_txt)

            # 3. 個人形
            with tabs[2]:
                c1, c2 = st.columns(2)
                opts_k = ["正", "補", "シード", "なし"]
                with c1:
                    st.markdown("#### 🚹 男子 個人形")
                    if p_m_k:
                        df_m_k = build_ind_df(m_names, "kata_chk", "kata_val", "kata_rank", limits["ind_kata_reg"]["max"], limits["ind_kata_sub"]["max"])
                        results["m_k"] = st.data_editor(df_m_k, column_config={"区分": st.column_config.SelectboxColumn(options=opts_k), "選手名": st.column_config.SelectboxColumn(options=[""] + m_names)}, num_rows="dynamic", hide_index=True, key="ed_m_k", use_container_width=True)
                    else: st.info(guide_txt)
                with c2:
                    st.markdown("#### 🚺 女子 個人形")
                    if p_w_k:
                        df_w_k = build_ind_df(w_names, "kata_chk", "kata_val", "kata_rank", limits["ind_kata_reg"]["max"], limits["ind_kata_sub"]["max"])
                        results["w_k"] = st.data_editor(df_w_k, column_config={"区分": st.column_config.SelectboxColumn(options=opts_k), "選手名": st.column_config.SelectboxColumn(options=[""] + w_names)}, num_rows="dynamic", hide_index=True, key="ed_w_k", use_container_width=True)
                    else: st.info(guide_txt)

            # 4. 個人組手
            with tabs[3]:
                c1, c2 = st.columns(2)
                with c1:
                    st.markdown("#### 🚹 男子 個人組手")
                    if t_conf["type"] == "standard":
                        if p_m_ku:
                            df_m_ku = build_ind_df(m_names, "kumi_chk", "kumi_val", "kumi_rank", limits["ind_kumi_reg"]["max"], limits["ind_kumi_sub"]["max"])
                            results["m_ku_std"] = st.data_editor(df_m_ku, column_config={"区分": st.column_config.SelectboxColumn(options=["正", "補", "シード", "なし"]), "選手名": st.column_config.SelectboxColumn(options=[""] + m_names)}, num_rows="dynamic", hide_index=True, key="ed_m_ku", use_container_width=True)
                        else: st.info(guide_txt)
                    else:
                        if not p_m_ku: st.info("💡 事前設定で参加階級が選択されていないため、エントリーできません。")
                        else:
                            st.caption("💡 上の事前設定で選択した階級のみ表示されています。")
                            weights_m = t_conf.get("weights_m", "-55,-61,-68,-76,+76").split(",")
                            for w in weights_m:
                                w_name = f"{w.strip()}kg級"
                                if school_meta.get(f"part_m_ku_{w_name}", False):
                                    with st.expander(f"▼ {w_name}", expanded=True):
                                        st.caption("人数を追加する場合は表下部の「＋」を押してください")
                                        df_w = build_weight_df(m_names, "kumi_chk", "kumi_val", "kumi_rank", "kumi_sub_val", w_name)
                                        results[f"m_ku_{w_name}"] = st.data_editor(df_w, column_config={"区分": st.column_config.SelectboxColumn(options=["正", "補", "シード", "なし"]), "選手名": st.column_config.SelectboxColumn(options=[""] + m_names)}, num_rows="dynamic", hide_index=True, key=f"ed_m_ku_{w_name}", use_container_width=True)
                with c2:
                    st.markdown("#### 🚺 女子 個人組手")
                    if t_conf["type"] == "standard":
                        if p_w_ku:
                            df_w_ku = build_ind_df(w_names, "kumi_chk", "kumi_val", "kumi_rank", limits["ind_kumi_reg"]["max"], limits["ind_kumi_sub"]["max"])
                            results["w_ku_std"] = st.data_editor(df_w_ku, column_config={"区分": st.column_config.SelectboxColumn(options=["正", "補", "シード", "なし"]), "選手名": st.column_config.SelectboxColumn(options=[""] + w_names)}, num_rows="dynamic", hide_index=True, key="ed_w_ku", use_container_width=True)
                        else: st.info(guide_txt)
                    else:
                        if not p_w_ku: st.info("💡 事前設定で参加階級が選択されていないため、エントリーできません。")
                        else:
                            st.caption("💡 上の事前設定で選択した階級のみ表示されています。")
                            weights_w = t_conf.get("weights_w", "-48,-53,-59,-66,+66").split(",")
                            for w in weights_w:
                                w_name = f"{w.strip()}kg級"
                                if school_meta.get(f"part_w_ku_{w_name}", False):
                                    with st.expander(f"▼ {w_name}", expanded=True):
                                        st.caption("人数を追加する場合は表下部の「＋」を押してください")
                                        df_w = build_weight_df(w_names, "kumi_chk", "kumi_val", "kumi_rank", "kumi_sub_val", w_name)
                                        results[f"w_ku_{w_name}"] = st.data_editor(df_w, column_config={"区分": st.column_config.SelectboxColumn(options=["正", "補", "シード", "なし"]), "選手名": st.column_config.SelectboxColumn(options=[""] + w_names)}, num_rows="dynamic", hide_index=True, key=f"ed_w_ku_{w_name}", use_container_width=True)

            st.markdown("<br>", unsafe_allow_html=True)
            if st.form_submit_button("✅ エントリーを保存 (全員分)", type="primary", use_container_width=True):
                has_error = False
                temp_processed = {}
                duplicate_checker = {}

                # 初期化（全メンバーのチェックを一度外す）
                for _, r in valid_members.iterrows():
                    uid = f"{s_id}_{r['name']}"
                    temp_processed[uid] = {
                        "team_kata_chk": False, "team_kata_role": "",
                        "team_kumi_chk": False, "team_kumi_role": "",
                        "kata_chk": False, "kata_val": "", "kata_rank": "",
                        "kumi_chk": False, "kumi_val": "", "kumi_rank": "", "kumi_sub_val": ""
                    }

                def apply_team(df, chk_k, role_k, sex_str, t_name):
                    nonlocal has_error
                    if df is None: return
                    used = set()
                    for _, row in df.iterrows():
                        n = row.get("選手名", "")
                        if n:
                            if n in used:
                                st.error(f"❌ {sex_str} {t_name}: 「{n}」が重複して選択されています。")
                                has_error = True
                            used.add(n)
                            uid = f"{s_id}_{n}"
                            if uid in temp_processed:
                                temp_processed[uid][chk_k] = True
                                temp_processed[uid][role_k] = row.get("役割", "")

                def apply_ind(df, chk_k, val_k, rank_k, sub_k, sex_str, t_name, is_weight, w_name=None):
                    nonlocal has_error
                    if df is None: return
                    used = set()
                    for _, row in df.iterrows():
                        n = row.get("選手名", "")
                        v = row.get("区分", "")
                        rk = str(row.get("順位", "")).strip()
                        if n and v not in ["なし", "出場しない"]:
                            if n in used:
                                st.error(f"❌ {sex_str} {t_name}{'('+w_name+')' if w_name else ''}: 「{n}」が重複して選択されています。")
                                has_error = True
                            used.add(n)
                            uid = f"{s_id}_{n}"
                            if uid in temp_processed:
                                need_rank = False
                                if v in ["正", "シード"]: need_rank = True
                                
                                if need_rank and not rk:
                                    st.error(f"❌ {n} {t_name}: 順位が入力されていません。")
                                    has_error = True
                                if not need_rank and rk:
                                    st.error(f"❌ {n} {t_name}: 「{v}」ですが順位が入力されています。順位を削除してください。")
                                    has_error = True
                                    
                                if need_rank and rk:
                                    key = f"{sex_str}_{t_name}_{v}" + (f"_{w_name}" if w_name else "")
                                    if key not in duplicate_checker: duplicate_checker[key] = {}
                                    if rk not in duplicate_checker[key]: duplicate_checker[key][rk] = []
                                    duplicate_checker[key][rk].append(n)

                                temp_processed[uid][chk_k] = True
                                if is_weight and w_name:
                                    temp_processed[uid][val_k] = w_name
                                    temp_processed[uid][sub_k] = v
                                else:
                                    temp_processed[uid][val_k] = v
                                temp_processed[uid][rank_k] = to_half_width(rk)

                # メタデータ(参加フラグ)の保存
                temp_processed[f"_meta_{s_id}"] = school_meta

                # データの適用
                apply_team(results.get("m_tk"), "team_kata_chk", "team_kata_role", "男子", "団体形")
                apply_team(results.get("m_tku"), "team_kumi_chk", "team_kumi_role", "男子", "団体組手")
                apply_ind(results.get("m_k"), "kata_chk", "kata_val", "kata_rank", "", "男子", "個人形", False)
                if t_conf["type"] == "standard":
                    apply_ind(results.get("m_ku_std"), "kumi_chk", "kumi_val", "kumi_rank", "", "男子", "個人組手", False)
                    apply_ind(results.get("w_ku_std"), "kumi_chk", "kumi_val", "kumi_rank", "", "女子", "個人組手", False)
                else:
                    weights_m = t_conf.get("weights_m", "-55,-61,-68,-76,+76").split(",")
                    for w in weights_m:
                        w_name = f"{w.strip()}kg級"
                        apply_ind(results.get(f"m_ku_{w_name}"), "kumi_chk", "kumi_val", "kumi_rank", "kumi_sub_val", "男子", "個人組手", True, w_name)
                    weights_w = t_conf.get("weights_w", "-48,-53,-59,-66,+66").split(",")
                    for w in weights_w:
                        w_name = f"{w.strip()}kg級"
                        apply_ind(results.get(f"w_ku_{w_name}"), "kumi_chk", "kumi_val", "kumi_rank", "kumi_sub_val", "女子", "個人組手", True, w_name)

                apply_ind(results.get("w_k"), "kata_chk", "kata_val", "kata_rank", "", "女子", "個人形", False)
                apply_team(results.get("w_tk"), "team_kata_chk", "team_kata_role", "女子", "団体形")
                apply_team(results.get("w_tku"), "team_kumi_chk", "team_kumi_role", "女子", "団体組手")

                for key, ranks in duplicate_checker.items():
                    for rank_val, names in ranks.items():
                        if len(names) > 1:
                            st.error(f"❌ {key} で順位『{rank_val}』が重複しています: {', '.join(names)}")
                            has_error = True

                if not has_error:
                    with st.spinner("💾 エントリーを保存しています..."):
                        cur_entries = load_entries(active_tid, force_reload=True)
                        cur_entries.update(temp_processed)
                        errs = validate_counts(valid_members, cur_entries, conf["limits"], t_conf["type"], {"m_kumite_mode":m_mode, "w_kumite_mode":w_mode}, s_id)
                        if errs:
                            for e in errs: st.error(e)
                        else:
                            save_entries(active_tid, cur_entries)
                            st.success("✅ 保存しました！")
                            time.sleep(1)
                            st.rerun()


        st.markdown("---")
        st.markdown("#### 📥 申込書の出力と提出")
        st.info("出力したExcelファイルに公印を押し、PDFや画像にしてから右側の枠へ提出してください。")
        c1, c2 = st.columns(2)
        with c1:
            st.markdown("##### 1. 申込書の作成")
            final_m = get_merged_data(s_id, active_tid)
            file_data, fname = generate_excel(s_id, s_data, final_m, active_tid, t_conf)
            if file_data:
                st.download_button("📄 Excel申込書をダウンロード", file_data, fname, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="secondary", use_container_width=True)
            else:
                st.error(f"作成失敗: {fname}")
        with c2:
            st.markdown("##### 2. 申込書のアップロード")
            u_file = st.file_uploader("ファイルを選択 (PDF, JPG, PNG 等)", type=['pdf', 'jpg', 'jpeg', 'png'], label_visibility="collapsed")
            if u_file:
                if st.button("✅ 申込書を提出する", type="primary", use_container_width=True):
                    with st.spinner("安全に送信中..."):
                        ok, res = upload_file_to_gas(u_file, base_name)
                        if ok: st.success("🎉 提出完了しました！管理者が確認いたします。")
                        else: st.error(f"❌ 失敗: {res}")

def admin_page():
    st.title("🔧 管理者画面")
    conf = load_conf()
    if not st.session_state.get("admin_ok", False):
        pw = st.text_input("Admin Password", type="password")
        if st.button("ログイン"):
            if pw == conf.get("admin_password", "1234"): st.session_state["admin_ok"] = True; st.rerun()
            else: st.error("パスワードが違います")
        return
    auth = load_auth(); admin_menu = ["🏆 大会設定", "📥 データ出力", "🏫 アカウント", "📅 年次処理"]
    admin_tab = st.radio("メニュー", admin_menu, index=st.session_state.get("admin_menu_idx", 0), horizontal=True)
    st.session_state["admin_menu_idx"] = admin_menu.index(admin_tab)
    st.divider()

    if admin_tab == "📥 データ出力":
        st.subheader("大会データの出力")
        tid = next((k for k, v in conf["tournaments"].items() if v["active"]), "kantou")
        if st.button("🔄 最新データで集計を開始"):
            with st.spinner("集計中..."):
                master = load_members_master(force_reload=True); entries = load_entries(tid, force_reload=True); full_data = []
                for _, m in master.iterrows():
                    ent = entries.get(f"{m['school_id']}_{m['name']}", {})
                    if ent.get("kata_chk") or ent.get("kumi_chk"):
                        row = m.to_dict(); row.update(ent); full_data.append(row)
                st.session_state["xlsx_tour"] = generate_tournament_excel(full_data, conf["tournaments"][tid]["type"], auth)
                st.session_state["xlsx_summ"] = generate_summary_excel(master, entries, auth, conf["tournaments"][tid]["type"])
                st.session_state["xlsx_adv"] = generate_advisor_excel(load_schools(), auth)
                st.session_state["xlsx_ts"] = datetime.datetime.now().strftime("%H:%M:%S")
        if "xlsx_ts" in st.session_state:
            st.success(f"✅ 集計完了 ({st.session_state['xlsx_ts']})")
            c1, c2, c3 = st.columns(3)
            c1.download_button("📥 トーナメントデータ", st.session_state["xlsx_tour"], "tournament.xlsx")
            c2.download_button("📊 参加校一覧集計", st.session_state["xlsx_summ"], "summary.xlsx")
            c3.download_button("👔 顧問リスト", st.session_state["xlsx_adv"], "advisors.xlsx")
            
        st.divider()
        st.subheader("📂 提出された申込書の確認")
        st.write("各校からアップロードされたファイルは、設定したGoogleドライブのフォルダに保存されています。")
        st.info("直接Googleドライブを開いて、ファイルを一括ダウンロードして管理してください。https://drive.google.com/drive/u/0/folders/1n5CM_Jh9g3MiYfRU1yugdhSFWJZyhDdT")
        
    elif admin_tab == "🏆 大会設定":
        st.subheader("基本設定")
        with st.form("conf_basic"):
            new_year = st.text_input("現在の年度", conf.get("year", "6"))
            t_opts = list(conf["tournaments"].keys())
            active_now = next((k for k, v in conf["tournaments"].items() if v["active"]), None)
            new_active = st.radio("受付中の大会", t_opts, index=t_opts.index(active_now) if active_now else 0, format_func=lambda x: conf["tournaments"][x]["name"])
            
            st.markdown("---")
            active_t_name = conf["tournaments"][active_now]["name"] if active_now else ""
            new_name = st.text_input("表示中(受付中)の大会名の変更", active_t_name)
            
            if st.form_submit_button("設定を保存 & 大会切替"):
                conf["year"] = new_year
                if active_now and new_name.strip():
                    conf["tournaments"][active_now]["name"] = new_name.strip()
                if new_active != active_now:
                    for k in conf["tournaments"]: conf["tournaments"][k]["active"] = (k == new_active)
                save_conf(conf); st.success("保存しました"); time.sleep(0.5); st.rerun()
        st.divider()
        with st.expander("参加人数制限の設定", expanded=True):
            with st.form("conf_limits"):
                lm = conf["limits"]
                st.caption("団体戦")
                c1, c2, c3 = st.columns(3)
                lm["team_kata"]["min"] = c1.number_input("団体形(正) 下限", 0, 10, lm["team_kata"].get("min", 3))
                lm["team_kata"]["max"] = c2.number_input("団体形(正) 上限", 0, 10, lm["team_kata"].get("max", 3))
                lm["team_kata"]["sub_max"] = c3.number_input("団体形(補) 上限", 0, 10, lm["team_kata"].get("sub_max", 1))
                c1, c2, c3 = st.columns(3)
                lm["team_kumite_5"]["min"] = c1.number_input("団体組手5人(正) 下限", 0, 10, lm["team_kumite_5"].get("min", 3))
                lm["team_kumite_5"]["max"] = c2.number_input("団体組手5人(正) 上限", 0, 10, lm["team_kumite_5"].get("max", 5))
                lm["team_kumite_5"]["sub_max"] = c3.number_input("団体組手5人(補) 上限", 0, 10, lm["team_kumite_5"].get("sub_max", 2))
                c1, c2, c3 = st.columns(3)
                lm["team_kumite_3"]["min"] = c1.number_input("団体組手3人(正) 下限", 0, 10, lm["team_kumite_3"].get("min", 2))
                lm["team_kumite_3"]["max"] = c2.number_input("団体組手3人(正) 上限", 0, 10, lm["team_kumite_3"].get("max", 3))
                lm["team_kumite_3"]["sub_max"] = c3.number_input("団体組手3人(補) 上限", 0, 10, lm["team_kumite_3"].get("sub_max", 1))
                st.caption("個人戦 (上限のみ)")
                c1, c2 = st.columns(2)
                lm["ind_kata_reg"]["max"] = c1.number_input("個人形(正) 上限", 0, 10, lm["ind_kata_reg"]["max"])
                lm["ind_kata_sub"]["max"] = c2.number_input("個人形(補) 上限", 0, 10, lm["ind_kata_sub"]["max"])
                c1, c2 = st.columns(2)
                lm["ind_kumi_reg"]["max"] = c1.number_input("個人組手(正) 上限", 0, 100, lm["ind_kumi_reg"]["max"])
                lm["ind_kumi_sub"]["max"] = c2.number_input("個人組手(補) 上限", 0, 10, lm["ind_kumi_sub"]["max"])
                if st.form_submit_button("人数制限を保存"): conf["limits"] = lm; save_conf(conf); st.success("保存しました")
        with st.expander("⚖️ 体重別階級の設定 (新人戦等)", expanded=True):
            with st.form("conf_weights"):
                active_t_conf = conf["tournaments"].get(active_now, {}) if active_now else {}
                cw1, cw2 = st.columns(2)
                wm = cw1.text_input("男子階級 (カンマ区切り)", active_t_conf.get("weights_m", "-55,-61,-68,-76,+76"))
                ww = cw2.text_input("女子階級 (カンマ区切り)", active_t_conf.get("weights_w", "-48,-53,-59,-66,+66"))
                if st.form_submit_button("階級を保存"):
                    if active_now:
                        conf["tournaments"][active_now]["weights_m"] = wm
                        conf["tournaments"][active_now]["weights_w"] = ww
                        save_conf(conf); st.success("保存しました"); time.sleep(1); st.rerun()
        with st.expander("🔐 管理者パスワード変更"):
            with st.form("admin_pw_change"):
                new_pw = st.text_input("新しい管理者パスワード", type="password")
                if st.form_submit_button("パスワードを変更して保存"):
                    if len(new_pw) >= 4: conf["admin_password"] = new_pw; save_conf(conf); st.success("変更しました"); time.sleep(1); st.session_state["admin_ok"] = False; st.rerun()
                    else: st.error("4文字以上にしてください")

    elif admin_tab == "🏫 アカウント":
        st.subheader("アカウント管理")
        recs = []
        for sid, d in auth.items(): recs.append({"ID": sid, "基本名": d.get("base_name",""), "略称": d.get("short_name", d.get("base_name","")), "No": d.get("school_no", 999), "Password": d.get("password",""), "校長名": d.get("principal","")})
        edited = st.data_editor(pd.DataFrame(recs), disabled=["ID"])
        if st.button("変更を保存"):
            for _, row in edited.iterrows():
                sid = row["ID"]
                if sid in auth: auth[sid].update({"base_name": row["基本名"], "short_name": row["略称"], "school_no": to_safe_int(row["No"]), "password": row["Password"], "principal": row["校長名"]})
            save_auth(auth); st.success("保存完了")
        st.divider()
        with st.expander("🗑️ 学校アカウントの削除", expanded=False):
            del_opts = {f"{v['base_name']} ({k})": k for k, v in auth.items()}; t_name = st.selectbox("削除する学校を選択", list(del_opts.keys()))
            if st.button("完全削除する", type="primary") and st.checkbox("理解して削除します"):
                t_sid = del_opts[t_name]; create_backup(); master = load_members_master(force_reload=True); save_members_master(master[master['school_id'] != t_sid])
                if t_sid in auth: del auth[t_sid]; save_auth(auth)
                st.success("削除しました"); time.sleep(1); st.rerun()

    elif admin_tab == "📅 年次処理":
        st.subheader("🌸 年度更新処理")
        if st.button("新年度を開始する"): st.success(perform_year_rollover())
        st.subheader("🎓 卒業生データ")
        grad_df = get_graduates_df()
        if not grad_df.empty:
            out = io.BytesIO(); pd.DataFrame(grad_df).to_excel(out, index=False); st.download_button("ダウンロード", out.getvalue(), "graduates.xlsx")
            if st.button("🗑️ 全て削除"): clear_graduates_archive(); st.success("削除完了")
        else: st.caption("なし")
        st.subheader("⏪ 復元")
        if st.button("バックアップから復元"): st.warning(restore_from_backup())

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
                    elif s_name == ph: st.error("❌ 学校を選択してください")
                    else: st.error("❌ パスワードが違います")
    elif nav == "🆕 新規登録":
        with st.form("reg"):
            bn = st.text_input("学校名 (「高等学校」不要)"); p = st.text_input("校長名"); pw = st.text_input("PW", type="password")
            if st.form_submit_button("登録"):
                if bn and pw:
                    nid = generate_school_id(); auth[nid] = {"base_name": bn, "password": pw, "principal": p, "school_no": 999, "advisors": []}
                    save_auth(auth); st.success("完了"); st.rerun()
    elif nav == "🔧 管理者": admin_page()

if __name__ == "__main__": main()
