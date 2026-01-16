import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.cell.cell import MergedCell
from openpyxl.styles import Alignment
import json
import datetime
import io
import os
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# ---------------------------------------------------------
# 1. 設定 & 定数定義
# ---------------------------------------------------------
TEMPLATE_FILE = 'template.xlsx'
KEY_FILE = 'secrets.json'       
SHEET_NAME = 'tournament_db'    

# Excel座標設定
COL_COORDS = {
    "tournament_name": "I3", "year": "E3", "date": "M7",
    "school_name": "C8", "principal": "C9", "head_advisor": "O9",
    "advisors_list": [
        {"name": "B42", "d1": "C42", "d2": "F42"},
        {"name": "B43", "d1": "C43", "d2": "F43"},
        {"name": "K42", "d1": "Q42", "d2": "U42"},
        {"name": "K43", "d1": "Q43", "d2": "U43"}
    ],
    "name": 2, "grade": 3, "dob": 4, "jkf_no": 19,
    "m_team_kata": 11, "m_team_kumite": 12, "m_kata": 13, "m_kumite": 14,
    "w_team_kata": 15, "w_team_kumite": 16, "w_kata": 17, "w_kumite": 18,
}

ADMIN_PASSWORD = "1234"

# ---------------------------------------------------------
# 2. Google Sheets 接続マネージャー (ハイブリッド対応版)
# ---------------------------------------------------------
@st.cache_resource
def get_gsheet_client():
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    
    # ★修正ポイント: ローカル(ファイル)とクラウド(Secrets)の両対応
    if os.path.exists(KEY_FILE):
        # PCで動かしているとき (secrets.jsonがある)
        creds = ServiceAccountCredentials.from_json_keyfile_name(KEY_FILE, scope)
    else:
        # Streamlit Cloudで動かしているとき (Secrets機能を使う)
        # 設定画面の "gcp_key" という名前の変数から中身を読み込む
        try:
            key_dict = json.loads(st.secrets["gcp_key"])
            creds = ServiceAccountCredentials.from_json_keyfile_dict(key_dict, scope)
        except Exception as e:
            st.error(f"認証エラー: Secretsの設定を確認してください。詳細: {e}")
            st.stop()
            
    client = gspread.authorize(creds)
    return client

def get_worksheet(tab_name):
    client = get_gsheet_client()
    sheet = client.open(SHEET_NAME)
    try:
        ws = sheet.worksheet(tab_name)
    except:
        ws = sheet.add_worksheet(title=tab_name, rows=100, cols=20)
    return ws

# --- A. JSON形式での保存 (Auth, Schools, Settings) ---
def load_json_from_sheet(tab_name, default_data):
    try:
        ws = get_worksheet(tab_name)
        val = ws.acell('A1').value
        if val:
            return json.loads(val)
        return default_data
    except Exception as e:
        return default_data

def save_json_to_sheet(tab_name, data):
    ws = get_worksheet(tab_name)
    json_str = json.dumps(data, ensure_ascii=False)
    ws.update_acell('A1', json_str)

# --- B. DataFrame形式での保存 (Members) ---
def load_members_from_sheet():
    default_cols = [
        "school", "name", "sex", "grade", "dob", "jkf_no", "active",
        "last_kata_chk", "last_kata_type", "last_kata_rank",
        "last_kumi_chk", "last_kumi_type", "last_kumi_rank",
        "last_t_kata_chk", "last_t_kata_role",
        "last_t_kumi_chk", "last_t_kumi_role"
    ]
    try:
        ws = get_worksheet("members")
        records = ws.get_all_records()
        if records:
            df = pd.DataFrame(records)
            for col in df.columns:
                if "chk" in col or "active" in col:
                    df[col] = df[col].apply(lambda x: True if str(x).upper() == "TRUE" else False)
            return df
        else:
            return pd.DataFrame(columns=default_cols)
    except:
        return pd.DataFrame(columns=default_cols)

def save_members_to_sheet(df):
    ws = get_worksheet("members")
    ws.clear()
    df_clean = df.fillna("")
    data = [df_clean.columns.tolist()] + df_clean.astype(str).values.tolist()
    ws.update(range_name='A1', values=data)

# ---------------------------------------------------------
# 3. データアクセスラッパー
# ---------------------------------------------------------
def load_auth(): return load_json_from_sheet("auth", {})
def save_auth(data): save_json_to_sheet("auth", data)

def load_schools(): return load_json_from_sheet("schools", {})
def save_schools(data): save_json_to_sheet("schools", data)

def load_settings():
    default_limits = {
        "ind_kata": {"reg": 4, "sub": 2}, 
        "ind_kumite": {"reg": 4, "sub": 2},
        "team_kata": {"reg": 3, "sub": 1}, 
        "team_kumite": {"reg": 5, "sub": 2}
    }
    default = {
        "year": "", "name": "",
        "limits": default_limits
    }
    data = load_json_from_sheet("settings", default)
    if not isinstance(data, dict): data = default
    if "limits" not in data or not isinstance(data["limits"], dict): data["limits"] = default_limits
    for key, val in default_limits.items():
        if key not in data["limits"] or not isinstance(data["limits"][key], dict):
            data["limits"][key] = val
        else:
            if "reg" not in data["limits"][key]: data["limits"][key]["reg"] = val["reg"]
            if "sub" not in data["limits"][key]: data["limits"][key]["sub"] = val["sub"]
    return data

def save_settings(data): save_json_to_sheet("settings", data)

# ---------------------------------------------------------
# 4. Excel作成 (個別申込書) & 管理者用一括出力
# ---------------------------------------------------------
def safe_write(ws, row, col, value, align_center=False):
    if value is None: return
    try:
        if isinstance(col, str) and not col.isdigit(): cell = ws[col]
        else: cell = ws.cell(row=row, column=col)
        if isinstance(cell, MergedCell):
            for r in ws.merged_cells.ranges:
                if cell.coordinate in r:
                    cell = ws.cell(row=r.min_row, column=r.min_col); break
        if str(value).endswith("年") and str(value)[:-1].isdigit(): value = str(value).replace("年", "")
        cell.value = value
        if align_center: cell.alignment = Alignment(horizontal='center', vertical='center')
    except: pass

def get_today_japanese_date():
    t = datetime.date.today()
    return f"令和{t.year-2018}年{t.month}月{t.day}日"

# --- 個別申込書生成 ---
def generate_excel(entry_list, school_name, school_data, settings):
    try: wb = openpyxl.load_workbook(TEMPLATE_FILE); ws = wb.active
    except: return None, "template.xlsx が見つかりません。"

    safe_write(ws, None, COL_COORDS["year"], settings.get("year", ""))
    safe_write(ws, None, COL_COORDS["tournament_name"], settings.get("name", ""))
    safe_write(ws, None, COL_COORDS["date"], get_today_japanese_date())
    safe_write(ws, None, COL_COORDS["school_name"], school_name)
    safe_write(ws, None, COL_COORDS["principal"], school_data.get("principal", ""))

    advs = school_data.get("advisors", [])
    head_name = advs[0]["name"] if advs else ""
    safe_write(ws, None, COL_COORDS["head_advisor"], head_name)

    for i, a in enumerate(advs[:4]):
        c = COL_COORDS["advisors_list"][i]
        safe_write(ws, None, c["name"], a["name"])
        safe_write(ws, None, c["d1"], "○" if a.get("d1") else "×", True)
        safe_write(ws, None, c["d2"], "○" if a.get("d2") else "×", True)

    START, CAP, OFFSET = 16, 22, 46
    for i, e in enumerate(entry_list):
        r = START + (i // CAP * OFFSET) + (i % CAP)
        safe_write(ws, r, COL_COORDS["name"], e["name"])
        safe_write(ws, r, COL_COORDS["grade"], e["grade"])
        safe_write(ws, r, COL_COORDS["dob"], e["dob"])
        safe_write(ws, r, COL_COORDS["jkf_no"], e["jkf_no"])

        sex = e["sex"]
        tk_c = COL_COORDS["m_team_kata"] if sex=="男子" else COL_COORDS["w_team_kata"]
        tku_c = COL_COORDS["m_team_kumite"] if sex=="男子" else COL_COORDS["w_team_kumite"]
        if e.get("team_kata_chk"): safe_write(ws, r, tk_c, "補" if e.get("team_kata_role")=="補欠" else "○", True)
        if e.get("team_kumi_chk"): safe_write(ws, r, tku_c, "補" if e.get("team_kumi_role")=="補欠" else "○", True)

        ik_c = COL_COORDS["m_kata"] if sex=="男子" else COL_COORDS["w_kata"]
        iku_c = COL_COORDS["m_kumite"] if sex=="男子" else COL_COORDS["w_kumite"]
        safe_write(ws, r, ik_c, format_rank(e.get("kata_type"), e.get("kata_rank")), True)
        safe_write(ws, r, iku_c, format_rank(e.get("kumite_type"), e.get("kumite_rank")), True)

    fname = f"申込書_{school_name}.xlsx"
    wb.save(fname)
    return fname, "作成成功"

# --- 管理者帳票 A: 選手詳細リスト ---
def generate_admin_entry_details(df, auth_data):
    output = io.BytesIO()
    df['school_no'] = df['school'].apply(lambda s: auth_data.get(s, {}).get('school_no', 9999))
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        categories = [
            ("男子個人形", "男子", "last_kata_chk", "last_kata_type", "last_kata_rank"),
            ("女子個人形", "女子", "last_kata_chk", "last_kata_type", "last_kata_rank"),
            ("男子個人組手", "男子", "last_kumi_chk", "last_kumi_type", "last_kumi_rank"),
            ("女子個人組手", "女子", "last_kumi_chk", "last_kumi_type", "last_kumi_rank"),
        ]
        for sheet_name, sex, chk_col, type_col, rank_col in categories:
            sub = df[(df['sex'] == sex) & (df[chk_col] == True)].copy()
            if not sub.empty:
                out_df = sub[['school_no', 'school', 'grade', 'name', type_col, rank_col, 'jkf_no']]
                out_df.columns = ['No', '学校名', '学年', '氏名', '種別', 'シード順位', 'JKF番号']
                out_df = out_df.sort_values(by=['No', '学年'])
                out_df.to_excel(writer, sheet_name=sheet_name, index=False)
            else:
                pd.DataFrame(columns=['No', '学校名', '学年', '氏名', '種別', 'シード順位', 'JKF番号']).to_excel(writer, sheet_name=sheet_name, index=False)

        team_cats = [
            ("男子団体形", "男子", "last_t_kata_chk"),
            ("女子団体形", "女子", "last_t_kata_chk"),
            ("男子団体組手", "男子", "last_t_kumi_chk"),
            ("女子団体組手", "女子", "last_t_kumi_chk"),
        ]
        for sheet_name, sex, chk_col in team_cats:
            sub = df[(df['sex'] == sex) & (df[chk_col] == True)].copy()
            if not sub.empty:
                grouped = sub.groupby(['school', 'school_no'])['name'].apply(list).reset_index()
                grouped['人数'] = grouped['name'].apply(len)
                grouped['メンバー'] = grouped['name'].apply(lambda x: "、".join(x))
                out_df = grouped[['school_no', 'school', '人数', 'メンバー']].rename(columns={'school': '学校名', 'school_no': 'No'})
                out_df = out_df.sort_values(by='No')
                out_df.to_excel(writer, sheet_name=sheet_name, index=False)
            else:
                pd.DataFrame(columns=['No', '学校名', '人数', 'メンバー']).to_excel(writer, sheet_name=sheet_name, index=False)
    return output.getvalue()

# --- 管理者帳票 B: 参加校集計表 ---
def generate_admin_school_summary(df, auth_data):
    output = io.BytesIO()
    schools = []
    for s_name, s_info in auth_data.items():
        schools.append({"name": s_name, "no": s_info.get("school_no", 9999)})
    s_df = pd.DataFrame(schools).sort_values(by="no")
    
    rows = []
    for _, s_row in s_df.iterrows():
        s = s_row["name"]
        my = df[df['school'] == s]
        
        def count_ind(sex, chk): return len(my[(my['sex']==sex) & (my[chk]==True)])
        def has_team(sex, chk): return "○" if not my[(my['sex']==sex) & (my[chk]==True)].empty else ""
        
        m_t_ka = has_team("男子", "last_t_kata_chk")
        m_i_ka = count_ind("男子", "last_kata_chk") or ""
        m_t_ku = has_team("男子", "last_t_kumi_chk")
        m_i_ku = count_ind("男子", "last_kumi_chk") or ""
        w_t_ka = has_team("女子", "last_t_kata_chk")
        w_i_ka = count_ind("女子", "last_kata_chk") or ""
        w_t_ku = has_team("女子", "last_t_kumi_chk")
        w_i_ku = count_ind("女子", "last_kumi_chk") or ""
        total = len(my) 
        rows.append([s_row["no"], s, m_t_ka, m_i_ka, m_t_ku, m_i_ku, w_t_ka, w_i_ka, w_t_ku, w_i_ku, total])
        
    cols = ["No", "学校名", "男団形", "男個形", "男団組", "男個組", "女団形", "女個形", "女団組", "女個組", "合計人数"]
    pd.DataFrame(rows, columns=cols).to_excel(output, index=False)
    return output.getvalue()

# --- 管理者帳票 C: 顧問一覧 ---
def generate_admin_advisor_list(schools_data, auth_data):
    output = io.BytesIO()
    rows = []
    s_list = sorted(auth_data.keys(), key=lambda k: auth_data[k].get("school_no", 9999))
    
    for sch in s_list:
        no = auth_data[sch].get("school_no", 9999)
        advs = schools_data.get(sch, {}).get("advisors", [])
        for i, a in enumerate(advs):
            pos = "筆頭顧問" if i == 0 else "顧問"
            rows.append({
                "No": no, "学校名": sch, "氏名": a["name"], "役職": pos,
                "役割": a.get("role", ""), 
                "1日目": "○" if a.get("d1") else "", 
                "2日目": "○" if a.get("d2") else ""
            })
    pd.DataFrame(rows).to_excel(output, index=False)
    return output.getvalue()

def format_rank(t, r):
    if not t: return None
    if t == "補欠": return "補"
    rs = str(r) if r else ""
    return f"○{rs}" if t == "一般" else f"シ{rs}"

def validate_entries(el, limits):
    errs = []
    cnt = {s: {c: {"reg":0, "sub":0} for c in ["ind_kata","ind_kumite","team_kata","team_kumite"]} for s in ["男子","女子"]}
    for e in el:
        s = e["sex"]
        if e["kata_type"] == "一般": cnt[s]["ind_kata"]["reg"]+=1
        elif e["kata_type"] == "補欠": cnt[s]["ind_kata"]["sub"]+=1
        if e["kumite_type"] == "一般": cnt[s]["ind_kumite"]["reg"]+=1
        elif e["kumite_type"] == "補欠": cnt[s]["ind_kumite"]["sub"]+=1
        if e["team_kata_chk"]: cnt[s]["team_kata"]["sub" if e["team_kata_role"]=="補欠" else "reg"]+=1
        if e["team_kumi_chk"]: cnt[s]["team_kumite"]["sub" if e["team_kumi_role"]=="補欠" else "reg"]+=1
    
    lbl = {"ind_kata":"個人形", "ind_kumite":"個人組手", "team_kata":"団体形", "team_kumite":"団体組手"}
    for s in ["男子","女子"]:
        for c, v in cnt[s].items():
            lr, ls = int(limits[c]["reg"]), int(limits[c]["sub"])
            if v["reg"] > lr: errs.append(f"❌ {s} {lbl[c]} (正選手): {v['reg']}名 (定員{lr})")
            if v["sub"] > ls: errs.append(f"❌ {s} {lbl[c]} (補欠): {v['sub']}名 (定員{ls})")
    return errs

# ---------------------------------------------------------
# 5. UI: Admin & School
# ---------------------------------------------------------
def admin_page():
    st.title("🔧 管理者モード")
    if st.text_input("パスワード", type="password") != ADMIN_PASSWORD: return
    st.success("認証成功")
    
    settings = load_settings()
    auth_data = load_auth()
    schools_data = load_schools()
    
    tab1, tab2, tab3 = st.tabs(["⚙️ 設定", "📊 集計・出力", "🏫 アカウント & No."])
    
    with tab1:
        with st.form("conf"):
            st.subheader("大会基本情報")
            c1, c2 = st.columns(2)
            ny = c1.text_input("年度", settings.get("year",""))
            nn = c2.text_input("大会名", settings.get("name",""))
            st.divider()
            st.subheader("定員設定")
            lm = settings["limits"]
            targets = [("個人形", "ind_kata"), ("個人組手", "ind_kumite"), ("団体形", "team_kata"), ("団体組手", "team_kumite")]
            nl = {}
            for label, key in targets:
                st.markdown(f"**{label}**")
                c_reg, c_sub = st.columns(2)
                try: val_r = int(lm[key].get("reg", 0))
                except: val_r = 0
                try: val_s = int(lm[key].get("sub", 0))
                except: val_s = 0
                r = c_reg.number_input(f"{label} (正選手)", value=val_r, key=f"r_{key}")
                s = c_sub.number_input(f"{label} (補欠)", value=val_s, key=f"s_{key}")
                nl[key] = {"reg": r, "sub": s}
            st.write("")
            if st.form_submit_button("設定を保存"):
                save_settings({"year": ny, "name": nn, "limits": nl})
                st.success("設定を保存しました")

    with tab2:
        st.subheader("帳票ダウンロードステーション")
        st.caption("※ すべて「学校番号順」に出力されます")
        all_members = load_members_from_sheet()
        
        c1, c2 = st.columns(2)
        with c1:
            st.markdown("##### 📄 1. 選手詳細リスト")
            st.caption("トーナメント作成用（種目別シート）")
            if st.button("作成 (Entry Details)"):
                d = generate_admin_entry_details(all_members, auth_data)
                st.download_button("📥 ダウンロード", d, "entry_details.xlsx")
        with c2:
            st.markdown("##### 🏫 2. 参加校集計表")
            st.caption("参加費計算・一覧表用")
            if st.button("作成 (School Summary)"):
                d = generate_admin_school_summary(all_members, auth_data)
                st.download_button("📥 ダウンロード", d, "school_summary.xlsx")
        st.divider()
        c3, c4 = st.columns(2)
        with c3:
            st.markdown("##### 👔 3. 顧問出欠リスト")
            st.caption("お弁当・審判編成用")
            if st.button("作成 (Advisor List)"):
                d = generate_admin_advisor_list(schools_data, auth_data)
                st.download_button("📥 ダウンロード", d, "advisor_list.xlsx")
        with c4:
            st.markdown("##### 🖥️ 4. システム用CSV")
            st.caption("全データ（生データ）")
            if st.button("作成 (Raw CSV)"):
                csv = all_members.to_csv(index=False).encode('utf-8_sig')
                st.download_button("📥 ダウンロード", csv, "raw_data.csv")

    with tab3:
        st.subheader("学校番号の管理")
        st.caption("ここで設定した番号順に帳票が出力されます")
        s_list = []
        for s_name, data in auth_data.items():
            s_list.append({"学校名": s_name, "No": data.get("school_no", 999)})
        df_nums = pd.DataFrame(s_list)
        edited_df = st.data_editor(df_nums, key="editor_nums", num_rows="fixed")
        if st.button("番号を保存して更新"):
            for index, row in edited_df.iterrows():
                nm = row["学校名"]
                if nm in auth_data:
                    auth_data[nm]["school_no"] = int(row["No"])
            save_auth(auth_data)
            st.success("学校番号を更新しました")
        st.divider()
        st.subheader("アカウント管理")
        tgt = st.selectbox("対象学校", list(auth_data.keys()))
        if tgt:
            with st.form("ren"):
                new_n = st.text_input("新学校名")
                if st.form_submit_button("変更"):
                    if new_n and new_n not in auth_data:
                        auth_data[new_n] = auth_data.pop(tgt)
                        save_auth(auth_data)
                        if tgt in schools_data: schools_data[new_n] = schools_data.pop(tgt); save_schools(schools_data)
                        m_df = load_members_from_sheet()
                        if not m_df.empty:
                            m_df.loc[m_df['school'] == tgt, 'school'] = new_n
                            save_members_to_sheet(m_df)
                        st.success("変更完了"); st.rerun()
                    else: st.error("エラー")
            with st.form("del"):
                confirm = st.checkbox("完全に削除する確認")
                if st.form_submit_button("削除実行"):
                    if confirm:
                        del auth_data[tgt]; save_auth(auth_data)
                        if tgt in schools_data: del schools_data[tgt]; save_schools(schools_data)
                        m_df = load_members_from_sheet()
                        save_members_to_sheet(m_df[m_df['school'] != tgt])
                        st.success("削除完了"); st.rerun()
                    else: st.error("確認チェックを入れてください")

def school_page(s_name):
    st.sidebar.markdown(f"**{s_name}**"); st.sidebar.button("ログアウト", on_click=lambda: st.session_state.pop("logged_in_school"))
    settings = load_settings()
    disp_title = f"令和{settings.get('year','〇')}年度 {settings.get('name','未定大会')}"
    st.title(f"🥋 {disp_title}")
    if "schools_data" not in st.session_state: st.session_state.schools_data = load_schools()
    if "members_df" not in st.session_state: st.session_state.members_df = load_members_from_sheet()
    
    s_data = st.session_state.schools_data.get(s_name, {"principal":"", "advisors":[]})
    t1, t2, t3 = st.tabs(["顧問", "部員", "エントリー"])

    with t1:
        np = st.text_input("校長", s_data.get("principal", ""))
        st.markdown("#### 顧問リスト")
        st.caption("※ リストの一番上が自動的に「筆頭顧問」になります")
        advs = s_data.get("advisors", [])
        for i, a in enumerate(advs):
            with st.container():
                c = st.columns([0.5, 2, 1.5, 1, 1, 0.5])
                if i == 0: c[0].markdown("👑")
                else:
                    if c[0].button("↑", key=f"up_{i}"):
                        advs[i], advs[i-1] = advs[i-1], advs[i]
                        s_data["advisors"] = advs
                        save_schools(st.session_state.schools_data); st.rerun()
                a["name"] = c[1].text_input("氏名", a["name"], key=f"n{i}", label_visibility="collapsed", placeholder="氏名")
                a["role"] = c[2].selectbox("役割", ["審判","競技記録","係員"], ["審判","競技記録","係員"].index(a.get("role","審判")), key=f"r{i}", label_visibility="collapsed")
                a["d1"] = c[3].checkbox("1日目", a.get("d1"), key=f"d1{i}")
                a["d2"] = c[4].checkbox("2日目", a.get("d2"), key=f"d2{i}")
                if c[5].button("×", key=f"del_{i}"):
                    advs.pop(i)
                    s_data["advisors"] = advs
                    save_schools(st.session_state.schools_data); st.rerun()
        if len(advs) > 1: st.caption("下へ移動させるには、下の人の「↑」を押してください")
        if st.button("＋ 顧問を追加"):
            advs.append({"name":"", "role":"審判", "d1":True, "d2":True})
            s_data["advisors"] = advs
            save_schools(st.session_state.schools_data); st.rerun()
        if st.button("保存", type="primary"):
            s_data["principal"] = np; s_data["advisors"] = advs
            st.session_state.schools_data[s_name] = s_data
            save_schools(st.session_state.schools_data); st.success("保存完了")

    with t2:
        with st.form("nm"):
            c = st.columns(3); nn = c[0].text_input("名"); ns = c[1].selectbox("性", ["男子","女子"]); ng = c[2].selectbox("学",["1","2","3"])
            c = st.columns(2); nd = c[0].text_input("誕"); nj = c[1].text_input("JKF")
            if st.form_submit_button("追加") and nn:
                st.session_state.members_df = pd.concat([st.session_state.members_df, pd.DataFrame([{"school":s_name, "name":nn, "sex":ns, "grade":ng, "dob":nd, "jkf_no":nj}])], ignore_index=True)
                save_members_to_sheet(st.session_state.members_df); st.success("OK"); st.rerun()
        m_df = st.session_state.members_df
        my = m_df[m_df['school'] == s_name].reset_index()
        for i, r in my.iterrows():
            c = st.columns([2,1,1,2,2,1])
            c[0].write(r['name']); c[1].write(r['sex']); c[2].write(r['grade']); c[5].button("削", key=f"md{r['index']}", on_click=lambda idx=r['index']: (save_members_to_sheet(m_df.drop(idx).reset_index(drop=True)), st.session_state.update({"members_df": load_members_from_sheet()})))

    with t3:
        df = st.session_state.members_df; tdf = df[df['school'] == s_name].copy()
        if tdf.empty: st.info("部員なし"); return
        men, women = tdf[tdf['sex']=="男子"], tdf[tdf['sex']=="女子"]
        ents = []; upds = {}
        def ren(r):
            c = st.columns([2,1.5,1.5,2.5,2.5]); c[0].write(f"{r['grade']} {r['name']}")
            tkc = c[1].checkbox("団体形", r.get("last_t_kata_chk"), key=f"tk{r['name']}")
            tkr = c[1].radio("-", ["正選手","補欠"], 0 if r.get("last_t_kata_role")=="正選手" else 1, key=f"tkr{r['name']}") if tkc else "正選手"
            tkuc = c[2].checkbox("団体組手", r.get("last_t_kumi_chk"), key=f"tku{r['name']}")
            tkur = c[2].radio("-", ["正選手","補欠"], 0 if r.get("last_t_kumi_role")=="正選手" else 1, key=f"tkur{r['name']}") if tkuc else "正選手"
            ikc = c[3].checkbox("個人形", r.get("last_kata_chk"), key=f"ik{r['name']}")
            def_opts = ["一般","シード","補欠"]; val_k = r.get("last_kata_type","一般")
            if val_k not in def_opts: val_k = "一般"
            ikt = "一般"; ikrk = ""
            if ikc:
                sc = c[3].columns([1.5,1])
                ikt = sc[0].radio("-", def_opts, def_opts.index(val_k), horizontal=True, key=f"ikt{r['name']}")
                if ikt!="補欠": ikrk = sc[1].text_input("-", r.get("last_kata_rank",""), key=f"ikr{r['name']}", placeholder="順位(数字)")
            ikuc = c[4].checkbox("個人組手", r.get("last_kumi_chk"), key=f"iku{r['name']}")
            val_ku = r.get("last_kumi_type","一般"); 
            if val_ku not in def_opts: val_ku = "一般"
            ikut = "一般"; ikurk = ""
            if ikuc:
                sc = c[4].columns([1.5,1])
                ikut = sc[0].radio("-", def_opts, def_opts.index(val_ku), horizontal=True, key=f"ikut{r['name']}")
                if ikut!="補欠": ikurk = sc[1].text_input("-", r.get("last_kumi_rank",""), key=f"ikur{r['name']}", placeholder="順位(数字)")
            e = {"name":r['name'], "sex":r['sex'], "grade":r['grade'], "dob":r['dob'], "jkf_no":r['jkf_no'], "team_kata_chk":tkc, "team_kata_role":tkr, "team_kumi_chk":tkuc, "team_kumi_role":tkur, "kata_type":ikt if ikc else None, "kata_rank":ikrk, "kumite_type":ikut if ikuc else None, "kumite_rank":ikurk}
            s = {"last_t_kata_chk":tkc, "last_t_kata_role":tkr, "last_t_kumi_chk":tkuc, "last_t_kumi_role":tkur, "last_kata_chk":ikc, "last_kata_type":ikt, "last_kata_rank":ikrk, "last_kumi_chk":ikuc, "last_kumi_type":ikut, "last_kumi_rank":ikurk}
            return e, s
        for _df, lab in [(men,"男子"),(women,"女子")]:
            if not _df.empty:
                st.subheader(f"{lab}の部")
                st.markdown(":gray[**学年 氏名 | 団体形 | 団体組手 | 個人形 (区分 / 順位) | 個人組手 (区分 / 順位)**]")
                st.markdown("<hr style='margin:0; padding:0;'>", unsafe_allow_html=True)
                for i, r in _df.iterrows(): e, s = ren(r); ents.append(e); upds[r['name']] = s; st.divider()
        if st.button("Excel作成", type="primary"):
            if errs := validate_entries(ents, load_settings()["limits"]): 
                for e in errs: st.error(e)
            else:
                f_df = st.session_state.members_df
                for idx, row in f_df.iterrows():
                    if row['school']==s_name and row['name'] in upds:
                        for k,v in upds[row['name']].items(): f_df.at[idx,k] = v
                save_members_to_sheet(f_df)
                fp, msg = generate_excel(ents, s_name, s_data, load_settings())
                if fp: st.success(msg); st.download_button("DL", open(fp,"rb"), fp)
                else: st.error(msg)

def main():
    st.set_page_config(page_title="大会エントリー", layout="wide")
    if "logged_in_school" in st.session_state: school_page(st.session_state["logged_in_school"]); return
    st.title("🔐 エントリーシステム"); auth = load_auth()
    t1, t2, t3 = st.tabs(["ログイン", "新規", "管理"])
    with t1:
        s = st.selectbox("学校", list(auth.keys()))
        if st.button("ログイン"):
            if s in auth and st.session_state.get("login_pw_val") == auth[s]["password"]:
                 st.session_state["logged_in_school"] = s; st.rerun()
            else: st.error("パスワードが違います")
        st.text_input("パスワード", type="password", key="login_pw_val")
    with t2:
        n = st.text_input("学校名"); p = st.text_input("校長"); pw = st.text_input("Pass", type="password")
        if st.button("登録") and n and pw:
            auth[n]={"password":pw, "principal":p, "school_no": 999}; save_auth(auth)
            sch = load_schools(); sch[n]={"principal":p, "advisors":[]}; save_schools(sch)
            st.session_state["logged_in_school"]=n; st.rerun()
    with t3:
        if st.checkbox("管理者"): admin_page()

if __name__ == "__main__": main()