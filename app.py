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

# ---------------------------------------------------------
# 1. 定数・初期設定
# ---------------------------------------------------------
KEY_FILE = 'secrets.json'
SHEET_NAME = 'tournament_db'
ADMIN_PASSWORD = "1234"

# デフォルトの大会設定
DEFAULT_TOURNAMENTS = {
    "kantou": {
        "name": "関東高等学校空手道大会 埼玉県予選",
        "template": "template_kantou.xlsx",
        "type": "standard",  # standard: 一般/シード/補欠 + 順位
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
        "type": "weight",    # weight: 階級選択
        "grades": [1, 2],
        "weights": "-55,-61,-68,-76,+76", # デフォルト階級
        "active": False
    },
    "senbatsu": {
        "name": "全国選抜 埼玉県予選",
        "template": "template_senbatsu.xlsx",
        "type": "division",  # division: 選抜/1年/高入生 (形なし)
        "grades": [1, 2],
        "active": False
    }
}

# Excel書き込み座標定義 (大会ごとに調整可能)
# ※現状は全大会で同じセル位置を使用する想定です。必要に応じて分岐を追加してください。
COORD_DEF = {
    "year": "E3", "tournament_name": "I3", "date": "M7",
    "school_name": "C8", "principal": "C9", "head_advisor": "O9",
    "advisors": [
        {"name": "B42", "d1": "C42", "d2": "F42"},
        {"name": "B43", "d1": "C43", "d2": "F43"},
        {"name": "K42", "d1": "Q42", "d2": "U42"},
        {"name": "K43", "d1": "Q43", "d2": "U43"}
    ],
    # 選手書き込み開始位置などの定義
    "start_row": 16, "cap": 22, "offset": 46,
    "cols": {
        "name": 2, "grade": 3, "dob": 4, "jkf_no": 19,
        "m_team_kata": 11, "m_team_kumite": 12, "m_kata": 13, "m_kumite": 14,
        "w_team_kata": 15, "w_team_kumite": 16, "w_kata": 17, "w_kumite": 18
    }
}

# ---------------------------------------------------------
# 2. Google Sheets 接続 & データ操作
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
            st.error(f"認証エラー: {e}"); st.stop()
    return gspread.authorize(creds)

def get_worksheet(tab_name):
    client = get_gsheet_client()
    sheet = client.open(SHEET_NAME)
    try: ws = sheet.worksheet(tab_name)
    except: ws = sheet.add_worksheet(title=tab_name, rows=100, cols=20)
    return ws

# --- JSONデータ (設定・アカウント) ---
def load_json(tab_name, default):
    try:
        val = get_worksheet(tab_name).acell('A1').value
        return json.loads(val) if val else default
    except: return default

def save_json(tab_name, data):
    ws = get_worksheet(tab_name)
    ws.update_acell('A1', json.dumps(data, ensure_ascii=False))

# --- 部員マスター (members) ---
def load_members_master():
    cols = ["school", "name", "sex", "grade", "dob", "jkf_no", "active"]
    try:
        recs = get_worksheet("members").get_all_records()
        return pd.DataFrame(recs) if recs else pd.DataFrame(columns=cols)
    except: return pd.DataFrame(columns=cols)

def save_members_master(df):
    ws = get_worksheet("members"); ws.clear()
    df = df.fillna("")
    ws.update([df.columns.tolist()] + df.astype(str).values.tolist())

# --- エントリーデータ (大会別) ---
def load_entries(tournament_id):
    # UUID(学校名_氏名)をキーにした辞書として読み込む
    try:
        ws = get_worksheet(f"entry_{tournament_id}")
        val = ws.acell('A1').value
        return json.loads(val) if val else {}
    except: return {}

def save_entries(tournament_id, data):
    # JSONとして保存 (行数削減のため)
    ws = get_worksheet(f"entry_{tournament_id}")
    ws.update_acell('A1', json.dumps(data, ensure_ascii=False))

# --- ラッパー ---
def load_auth(): return load_json("auth", {})
def save_auth(d): save_json("auth", d)
def load_schools(): return load_json("schools", {})
def save_schools(d): save_json("schools", d)
def load_conf(): return load_json("config", {"year": "6", "tournaments": DEFAULT_TOURNAMENTS})
def save_conf(d): save_json("config", d)

# ---------------------------------------------------------
# 3. ロジック: 年次更新 & データ結合
# ---------------------------------------------------------
def perform_year_rollover():
    # 1. 部員の学年更新 (+1, 3年は削除)
    df = load_members_master()
    if not df.empty:
        df['grade'] = pd.to_numeric(df['grade'], errors='coerce').fillna(0).astype(int)
        df['grade'] = df['grade'] + 1
        df = df[df['grade'] <= 3] # 4年生以上(元3年)を削除
        save_members_master(df)
    
    # 2. 全大会のエントリーシートをクリア
    conf = load_conf()
    for tid in conf["tournaments"].keys():
        save_entries(tid, {})
    
    # 3. 年度の更新
    try:
        conf["year"] = str(int(conf["year"]) + 1)
        save_conf(conf)
    except: pass
    return "新年度更新が完了しました。全学年が+1され、エントリーデータはリセットされました。"

def get_merged_data(school_name, tournament_id):
    # マスターデータとエントリーデータを結合して表示用DFを作る
    master = load_members_master()
    if master.empty: return pd.DataFrame()
    
    my_members = master[master['school'] == school_name].copy()
    entries = load_entries(tournament_id)
    
    # エントリー情報をマージ
    def get_ent(row, key):
        uid = f"{row['school']}_{row['name']}"
        return entries.get(uid, {}).get(key, None)

    cols_to_add = [
        "team_kata_chk", "team_kata_role", "team_kumi_chk", "team_kumi_role",
        "kata_chk", "kata_val", "kata_rank", # valは種別(一般/シード)または区分
        "kumi_chk", "kumi_val", "kumi_rank"  # valは種別(一般/-55kg/選抜)など
    ]
    
    for c in cols_to_add:
        my_members[f"last_{c}"] = my_members.apply(lambda r: get_ent(r, c), axis=1)
        
    return my_members

# ---------------------------------------------------------
# 4. Excel出力
# ---------------------------------------------------------
def generate_excel(school_name, school_data, members_df, t_id, t_conf):
    coords = COORD_DEF
    template_file = t_conf.get("template", "template.xlsx")
    
    try: wb = openpyxl.load_workbook(template_file); ws = wb.active
    except: return None, f"{template_file} が見つかりません。"
    
    conf = load_conf()
    ws[coords["year"]] = conf.get("year", "")
    ws[coords["tournament_name"]] = t_conf.get("name", "")
    ws[coords["date"]] = f"令和{datetime.date.today().year-2018}年{datetime.date.today().month}月{datetime.date.today().day}日"
    ws[coords["school_name"]] = school_name
    ws[coords["principal"]] = school_data.get("principal", "")
    
    # 顧問
    advs = school_data.get("advisors", [])
    head = advs[0]["name"] if advs else ""
    ws[coords["head_advisor"]] = head
    for i, a in enumerate(advs[:4]):
        c = coords["advisors"][i]
        ws[c["name"]] = a["name"]
        ws[c["d1"]] = "○" if a.get("d1") else "×"
        ws[c["d2"]] = "○" if a.get("d2") else "×"
    
    # 選手
    cols = coords["cols"]
    entries = members_df[
        (members_df['last_team_kata_chk']==True) | (members_df['last_team_kumi_chk']==True) |
        (members_df['last_kata_chk']==True) | (members_df['last_kumi_chk']==True)
    ].sort_values(by="grade") # 学年順

    for i, (_, row) in enumerate(entries.iterrows()):
        r = coords["start_row"] + (i // coords["cap"] * coords["offset"]) + (i % coords["cap"])
        
        ws.cell(row=r, column=cols["name"], value=row["name"])
        ws.cell(row=r, column=cols["grade"], value=row["grade"])
        ws.cell(row=r, column=cols["dob"], value=row["dob"])
        ws.cell(row=r, column=cols["jkf_no"], value=row["jkf_no"])
        
        sex = row["sex"]
        # 団体
        tk_col = cols["m_team_kata"] if sex=="男子" else cols["w_team_kata"]
        tku_col = cols["m_team_kumite"] if sex=="男子" else cols["w_team_kumite"]
        if row.get("last_team_kata_chk"):
            ws.cell(row=r, column=tk_col, value="補" if row.get("last_team_kata_role")=="補欠" else "○").alignment = Alignment(horizontal='center')
        if row.get("last_team_kumi_chk"):
            ws.cell(row=r, column=tku_col, value="補" if row.get("last_team_kumi_role")=="補欠" else "○").alignment = Alignment(horizontal='center')
            
        # 個人
        k_col = cols["m_kata"] if sex=="男子" else cols["w_kata"]
        ku_col = cols["m_kumite"] if sex=="男子" else cols["w_kumite"]
        
        # 形の書き込みロジック
        if row.get("last_kata_chk"):
            val = row.get("last_kata_val")
            rank = row.get("last_kata_rank", "")
            if val == "補欠": txt = "補"
            elif t_conf["type"] == "standard": txt = f"○{rank}" if val=="一般" else f"シ{rank}"
            else: txt = "○" # その他の大会
            ws.cell(row=r, column=k_col, value=txt).alignment = Alignment(horizontal='center')

        # 組手の書き込みロジック
        if row.get("last_kumi_chk"):
            val = row.get("last_kumi_val")
            rank = row.get("last_kumi_rank", "")
            if val == "補欠": txt = "補"
            elif t_conf["type"] == "standard": txt = f"○{rank}" if val=="一般" else f"シ{rank}"
            elif t_conf["type"] == "weight": txt = str(val) # 階級をそのまま
            elif t_conf["type"] == "division": txt = str(val) # 区分をそのまま
            else: txt = "○"
            ws.cell(row=r, column=ku_col, value=txt).alignment = Alignment(horizontal='center')

    fname = f"申込書_{school_name}.xlsx"
    wb.save(fname)
    return fname, "作成成功"

# ---------------------------------------------------------
# 5. UI: 学校用ページ
# ---------------------------------------------------------
def school_page(s_name):
    # サイドバー: 情報 & 隠しログアウト
    st.sidebar.title("メニュー")
    st.sidebar.markdown(f"**{s_name}** 様")
    
    conf = load_conf()
    active_tid = next((k for k, v in conf["tournaments"].items() if v["active"]), None)
    t_conf = conf["tournaments"].get(active_tid, {}) if active_tid else {}
    
    if not active_tid:
        st.error("現在受付中の大会はありません。"); return

    st.sidebar.info(f"🏆 {t_conf['name']}")
    
    with st.sidebar.expander("⚙️ その他・ログアウト"):
        if st.button("ログアウト"):
            st.query_params.clear() # URLパラメータ消去
            st.session_state.pop("logged_in_school", None)
            st.rerun()

    st.title(f"🥋 {t_conf['name']} エントリー")

    # データロード
    if "schools_data" not in st.session_state: st.session_state.schools_data = load_schools()
    s_data = st.session_state.schools_data.get(s_name, {"principal":"", "advisors":[]})
    
    tab1, tab2, tab3 = st.tabs(["① 顧問登録", "② 部員名簿", "③ 大会エントリー"])

    # --- ① 顧問 ---
    with tab1:
        np = st.text_input("校長名", s_data.get("principal", ""))
        st.markdown("#### 顧問リスト")
        st.caption("※ 一番上の先生が「筆頭顧問」になります（矢印で入替可能）")
        
        advs = s_data.get("advisors", [])
        for i, a in enumerate(advs):
            with st.container():
                c = st.columns([0.8, 2, 1.5, 0.5, 0.5, 0.7])
                if i == 0: c[0].markdown("**筆頭顧問**")
                else:
                    c[0].markdown(f"<small>顧問({i+1})</small>", unsafe_allow_html=True)
                    if c[0].button("↑", key=f"up_{i}"):
                        advs[i], advs[i-1] = advs[i-1], advs[i]
                        s_data["advisors"] = advs
                        # キャッシュクリア
                        for idx in [i, i-1]:
                            for p in ["n","r","d1","d2"]: st.session_state.pop(f"{p}{idx}", None)
                        save_schools(st.session_state.schools_data); st.rerun()

                a["name"] = c[1].text_input("氏名", a["name"], key=f"n{i}", label_visibility="collapsed", placeholder="氏名")
                a["role"] = c[2].selectbox("役割", ["審判","競技記録","係員"], index=["審判","競技記録","係員"].index(a.get("role","審判")), key=f"r{i}", label_visibility="collapsed")
                a["d1"] = c[3].checkbox("1日", a.get("d1"), key=f"d1{i}")
                a["d2"] = c[4].checkbox("2日", a.get("d2"), key=f"d2{i}")
                if c[5].button("削除", key=f"del_{i}"):
                    advs.pop(i)
                    s_data["advisors"] = advs
                    for k in list(st.session_state.keys()):
                        if k.startswith(("n","r","d1","d2")) and k[2:].isdigit(): del st.session_state[k]
                    save_schools(st.session_state.schools_data); st.rerun()

        if st.button("＋ 顧問を追加"):
            advs.append({"name":"", "role":"審判", "d1":True, "d2":True})
            s_data["advisors"] = advs
            save_schools(st.session_state.schools_data); st.rerun()
        
        if st.button("顧問情報を保存", type="primary"):
            s_data["principal"] = np; s_data["advisors"] = advs
            st.session_state.schools_data[s_name] = s_data
            save_schools(st.session_state.schools_data); st.success("保存しました")

    # --- ② 部員名簿 ---
    with tab2:
        st.caption("※ ここは「全大会共通」の名簿です。大会ごとの出場チェックは「③ エントリー」で行ってください。")
        with st.form("add_member"):
            c = st.columns(3)
            nn = c[0].text_input("氏名")
            ns = c[1].selectbox("性別", ["男子", "女子"])
            ng = c[2].selectbox("学年", [1, 2, 3])
            c2 = st.columns(2)
            nd = c2[0].text_input("生年月日 (例: H20.4.1)")
            nj = c2[1].text_input("JKF会員登録番号")
            if st.form_submit_button("部員を追加"):
                if nn:
                    master = load_members_master()
                    new_row = pd.DataFrame([{"school":s_name, "name":nn, "sex":ns, "grade":ng, "dob":nd, "jkf_no":nj, "active":True}])
                    save_members_master(pd.concat([master, new_row], ignore_index=True))
                    st.success(f"{nn} さんを追加しました"); st.rerun()
        
        # 一覧表示 & 削除
        master = load_members_master()
        my_m = master[master['school']==s_name].reset_index()
        for i, r in my_m.iterrows():
            c = st.columns([2, 1, 1, 2])
            c[0].write(r['name'])
            c[1].write(r['sex'])
            c[2].write(f"{r['grade']}年")
            if c[3].button("削除", key=f"m_del_{r['index']}"):
                save_members_master(master.drop(r['index'])); st.rerun()

    # --- ③ エントリー ---
    with tab3:
        st.markdown(f"**対象学年:** {t_conf['grades']} 年生")
        
        # 対象学年でフィルタ
        merged = get_merged_data(s_name, active_tid)
        if merged.empty: st.info("部員が登録されていません。"); return
        
        # 学年フィルタ
        target_grades = [int(g) for g in t_conf['grades']]
        valid_members = merged[merged['grade'].isin(target_grades)].copy()
        
        if valid_members.empty:
            st.warning("この大会に出場できる学年の部員がいません。")
            return

        men = valid_members[valid_members['sex']=="男子"]
        women = valid_members[valid_members['sex']=="女子"]
        
        entries_update = load_entries(active_tid) # 保存用辞書
        
        def render_entry_row(r):
            # UID
            uid = f"{r['school']}_{r['name']}"
            entry_data = entries_update.get(uid, {})
            
            c = st.columns([2, 1.5, 1.5, 2.5, 2.5])
            c[0].markdown(f"**{r['grade']}年 {r['name']}**")
            
            # --- 団体 ---
            # 新人戦・関東・インハイはあり。選抜は無し？(一旦ありとするが、無しならConfigで制御可能)
            tk = c[1].checkbox("団体形", r.get("last_team_kata_chk"), key=f"tk_{uid}")
            tkr = "正選手"
            if tk: tkr = c[1].radio("役", ["正選手","補欠"], 0 if r.get("last_team_kata_role")=="正選手" else 1, key=f"tkr_{uid}", horizontal=True, label_visibility="collapsed")
            
            tku = c[2].checkbox("団体組手", r.get("last_team_kumi_chk"), key=f"tku_{uid}")
            tkur = "正選手"
            if tku: tkur = c[2].radio("役", ["正選手","補欠"], 0 if r.get("last_team_kumi_role")=="正選手" else 1, key=f"tkur_{uid}", horizontal=True, label_visibility="collapsed")
            
            # --- 個人形 ---
            # 選抜は「形」なし
            k_chk = False; k_val = ""; k_rank = ""
            if t_conf["type"] != "division": # division(選抜)以外は形あり
                k_chk = c[3].checkbox("個人形", r.get("last_kata_chk"), key=f"k_{uid}")
                if k_chk:
                    opts = ["一般","シード","補欠"]
                    def_val = r.get("last_kata_val", "一般")
                    k_val = c[3].selectbox("区分", opts, opts.index(def_val) if def_val in opts else 0, key=f"kv_{uid}", label_visibility="collapsed")
                    if k_val != "補欠":
                        k_rank = c[3].text_input("順位", r.get("last_kata_rank",""), key=f"kr_{uid}", placeholder="数字", label_visibility="collapsed")

            # --- 個人組手 ---
            # 大会タイプによってUI変化
            ku_chk = c[4].checkbox("個人組手", r.get("last_kumi_chk"), key=f"ku_{uid}")
            ku_val = ""; ku_rank = ""
            
            if ku_chk:
                if t_conf["type"] == "standard": # 関東・インハイ
                    opts = ["一般","シード","補欠"]
                    def_val = r.get("last_kumi_val", "一般")
                    ku_val = c[4].selectbox("区分", opts, opts.index(def_val) if def_val in opts else 0, key=f"kuv_{uid}", label_visibility="collapsed")
                    if ku_val != "補欠":
                        ku_rank = c[4].text_input("順位", r.get("last_kumi_rank",""), key=f"kur_{uid}", placeholder="数字", label_visibility="collapsed")
                
                elif t_conf["type"] == "weight": # 新人 (階級)
                    # Configから階級リストを取得
                    w_str = t_conf.get("weights", "-55,-61,-68,-76,+76")
                    w_list = [f"{w.strip()}kg級" for w in w_str.split(",")] + ["補欠"]
                    def_val = r.get("last_kumi_val", w_list[0])
                    # 保存値にkg級がついてない場合のケア
                    if def_val not in w_list and def_val != "補欠": def_val = f"{def_val}kg級"
                    
                    ku_val = c[4].selectbox("階級", w_list, w_list.index(def_val) if def_val in w_list else 0, key=f"kuv_{uid}", label_visibility="collapsed")
                
                elif t_conf["type"] == "division": # 選抜 (区分)
                    d_list = ["選抜の部", "1年生の部", "高入生の部", "補欠"]
                    def_val = r.get("last_kumi_val", "選抜の部")
                    ku_val = c[4].selectbox("出場区分", d_list, d_list.index(def_val) if def_val in d_list else 0, key=f"kuv_{uid}", label_visibility="collapsed")

            # データ更新
            entry_data.update({
                "team_kata_chk": tk, "team_kata_role": tkr,
                "team_kumi_chk": tku, "team_kumi_role": tkur,
                "kata_chk": k_chk, "kata_val": k_val, "kata_rank": k_rank,
                "kumi_chk": ku_chk, "kumi_val": ku_val, "kumi_rank": ku_rank
            })
            entries_update[uid] = entry_data

        st.subheader("男子")
        for i, r in men.iterrows(): render_entry_row(r); st.divider()
        st.subheader("女子")
        for i, r in women.iterrows(): render_entry_row(r); st.divider()
        
        if st.button("エントリー保存 & Excel作成", type="primary"):
            save_entries(active_tid, entries_update)
            fp, msg = generate_excel(s_name, s_data, get_merged_data(s_name, active_tid), active_tid, t_conf)
            if fp:
                st.success("保存完了！Excelをダウンロードできます。")
                st.download_button("Excelダウンロード", open(fp,"rb"), fp)
            else: st.error(msg)

# ---------------------------------------------------------
# 6. UI: 管理者ページ
# ---------------------------------------------------------
def admin_page():
    st.title("🔧 管理者画面")
    if st.text_input("Admin Password", type="password") != ADMIN_PASSWORD: return
    
    conf = load_conf()
    auth = load_auth()
    
    t1, t2, t3, t4 = st.tabs(["🏆 大会設定", "📥 データ出力", "🏫 アカウント", "📅 年次処理"])
    
    with t1:
        st.subheader("大会マスター設定")
        st.caption("現在アクティブにする大会を選択してください")
        
        # 大会選択ラジオボタン
        t_opts = list(conf["tournaments"].keys())
        active_now = next((k for k, v in conf["tournaments"].items() if v["active"]), None)
        new_active = st.radio("受付中の大会", t_opts, index=t_opts.index(active_now) if active_now else 0, format_func=lambda x: conf["tournaments"][x]["name"])
        
        if new_active != active_now:
            if st.button("大会を切り替える"):
                for k in conf["tournaments"]: conf["tournaments"][k]["active"] = (k == new_active)
                save_conf(conf); st.success("切り替えました"); st.rerun()

        st.divider()
        st.subheader("詳細設定 (新人戦の階級など)")
        target_t = st.selectbox("編集する大会", t_opts, format_func=lambda x: conf["tournaments"][x]["name"])
        t_data = conf["tournaments"][target_t]
        
        with st.form("edit_t"):
            st.text_input("大会名", t_data["name"], disabled=True)
            if t_data["type"] == "weight":
                w_in = st.text_area("階級リスト (カンマ区切り, 数字のみでOK)", t_data.get("weights", ""))
                if st.form_submit_button("階級を保存"):
                    conf["tournaments"][target_t]["weights"] = w_in
                    save_conf(conf); st.success("保存しました")
            else:
                st.info("この大会には設定可能な階級リストはありません。")

    with t2:
        st.subheader("帳票ダウンロード")
        # アクティブな大会のデータを使用
        tid = next((k for k, v in conf["tournaments"].items() if v["active"]), "kantou")
        st.caption(f"対象データ: {conf['tournaments'][tid]['name']}")
        
        # 全データ結合
        master = load_members_master()
        entries = load_entries(tid)
        
        # 結合ロジック (簡易版)
        full_data = []
        for _, m in master.iterrows():
            uid = f"{m['school']}_{m['name']}"
            ent = entries.get(uid, {})
            # エントリーがある、かつどれかにチェックが入っている場合のみ
            if ent and (ent.get("kata_chk") or ent.get("kumi_chk") or ent.get("team_kata_chk") or ent.get("team_kumi_chk")):
                row = m.to_dict()
                row.update(ent)
                row["school_no"] = auth.get(m['school'], {}).get("school_no", 999)
                full_data.append(row)
        
        df_out = pd.DataFrame(full_data)
        
        if not df_out.empty:
            df_out = df_out.sort_values(by=["school_no", "grade"])
            csv = df_out.to_csv(index=False).encode('utf-8_sig')
            st.download_button("エントリー一覧 (CSV)", csv, "entries.csv")
            
            # ここにExcel帳票機能 (A,B,C) を実装可能
            # ※長くなるので今回はCSVのみ実装していますが、以前のコードのロジックでExcel化可能です
        else:
            st.warning("エントリーデータがありません")

    with t3:
        st.subheader("学校番号管理")
        s_list = [{"学校名":k, "No":v.get("school_no",999)} for k,v in auth.items()]
        edf = st.data_editor(pd.DataFrame(s_list), key="sed", num_rows="fixed")
        if st.button("番号保存"):
            for i, r in edf.iterrows():
                if r["学校名"] in auth: auth[r["学校名"]]["school_no"] = int(r["No"])
            save_auth(auth); st.success("保存しました")
            
    with t4:
        st.subheader("🌸 年度更新処理")
        st.warning("【注意】これを押すと、全員の学年が+1され、3年生は削除され、全大会のエントリー情報がリセットされます。")
        if st.button("新年度を開始する (実行確認)"):
            res = perform_year_rollover()
            st.success(res)

# ---------------------------------------------------------
# 7. Main
# ---------------------------------------------------------
def main():
    st.set_page_config(page_title="大会エントリー", layout="wide")
    
    # URL自動ログイン
    qp = st.query_params
    if "school" in qp:
        st.session_state["logged_in_school"] = qp["school"]
    
    if "logged_in_school" in st.session_state:
        # URLに状態を反映
        st.query_params["school"] = st.session_state["logged_in_school"]
        school_page(st.session_state["logged_in_school"])
        return

    st.title("🔐 エントリーシステム")
    auth = load_auth()
    
    t1, t2, t3 = st.tabs(["ログイン", "新規登録", "管理者"])
    with t1:
        s = st.selectbox("学校名", list(auth.keys()))
        pw = st.text_input("パスワード", type="password")
        if st.button("ログイン"):
            if s in auth and auth[s]["password"] == pw:
                st.session_state["logged_in_school"] = s
                st.rerun()
            else: st.error("パスワードが違います")
    with t2:
        n = st.text_input("学校名 (新規)")
        p = st.text_input("校長名")
        new_pw = st.text_input("パスワード (設定)", type="password")
        if st.button("登録"):
            if n and new_pw:
                auth[n] = {"password": new_pw, "principal": p, "school_no": 999}
                save_auth(auth); st.success("登録しました"); st.rerun()
    with t3:
        admin_page()

if __name__ == "__main__": main()