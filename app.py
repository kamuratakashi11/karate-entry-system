"""閉鎖後の申込窓口（跡地）。新システム（さくら）への誘導だけを行う。

データへのアクセスが無い＝Google Sheets の認証情報が不要なので、
Streamlit Cloud の Secrets（gcp_key）を消しても動く。

閉鎖するとき（推奨: 8/19 の最終データ取り込みの直後）:
    1. このファイルの中身を app.py に上書きして GitHub へ push
       （Streamlit Cloud が自動で再デプロイする）
    2. Streamlit Cloud の Settings → Secrets から gcp_key を削除してよい
万一の切り戻し:
    閉鎖のコミットを git revert して push（元の申込システムに戻る）

このファイルと app.py は**中身が同じ**（閉鎖後）。直すときは両方に同じものを置く。
"""

import streamlit as st

NEW_URL = "https://saitamahs-karate.sakura.ne.jp/entry/"

st.set_page_config(page_title="大会申込システム（移転しました）",
                   page_icon="🥋", layout="centered")

st.title("大会申込システムは移転しました")
st.markdown(
    f"""
申込の受付は、下の新しいページに変わりました。**このページでは受け付けていません。**

### 👉 [{NEW_URL}]({NEW_URL})
"""
)

try:
    st.link_button("新しい申込システムをひらく", NEW_URL,
                   type="primary", use_container_width=True)
except Exception:
    pass  # 古い streamlit には link_button が無い。上のリンクで足りる

st.markdown(
    """
- **学校名とパスワードはこれまでと同じ**です（そのままログインできます）
- **これまでに入力した部員名簿とエントリーは、そのまま新しいページに移してあります**
- 新人大会の受付は **8月20日から** です
- お気に入り・ブックマークは、新しいページに登録し直してください
- パスワードを忘れた場合や、うまく開けない場合は空手道専門部までご連絡ください
"""
)
