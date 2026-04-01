
# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
from pathlib import Path
import threading, time

DATA_PATH = Path("app_data.xlsx")
SHEET_INSTR = "計測器"
SHEET_TC_USAGE = "熱電対使用履歴"
SHEET_TC_INV = "熱電対在庫"

# === 列定義（係数+単位を2組） ===
NEW_FIELDS = [
    "メーカー","購入日","校正期限",
    "校正係数1","校正係数1_単位",
    "校正係数2","校正係数2_単位",
]
DATE_FIELDS = ["貸出日","返却予定日","購入日","校正期限"]
TEXT_FIELDS_RECOMMENDED = [
    "名称","型式","所属","氏名","使用場所","使用用途","保管場所","備考","メーカー"
]

# === ヘルパ：有効数字で丸め ===
def _round_sig(x, sig=6):
    if x is None or pd.isna(x):
        return x
    try:
        x = float(x)
    except Exception:
        return pd.NA
    if x == 0.0:
        return 0.0
    return float(f"{x:.{sig}g}")

@st.cache_data(ttl=60)
def load_data():
    df_instr = pd.read_excel(DATA_PATH, sheet_name=SHEET_INSTR, engine="openpyxl")

    # Excel由来の不要列（Unnamed:* / 空ヘッダ）を除去
    drop_cols = [c for c in df_instr.columns if str(c).strip()=='' or str(c).lower().startswith('unnamed:')]
    if drop_cols:
        df_instr = df_instr.drop(columns=drop_cols, errors='ignore')

    # 過去互換：旧『構成係数』→『校正係数』
    if '構成係数' in df_instr.columns and '校正係数' not in df_instr.columns:
        df_instr = df_instr.rename(columns={'構成係数':'校正係数'})

    # 新2組列を最低限生やす
    for c in ["校正係数1","校正係数1_単位","校正係数2","校正係数2_単位"]:
        if c not in df_instr.columns:
            df_instr[c] = ""

    # 旧『校正係数』『校正係数_単位』があれば 係数1 側へ移行（空/NaN のときのみ上書き）
    if '校正係数' in df_instr.columns:
        mask_empty = df_instr['校正係数1'].astype(str).str.strip().eq('') | df_instr['校正係数1'].isna()
        df_instr.loc[mask_empty, '校正係数1'] = pd.to_numeric(df_instr.loc[mask_empty, '校正係数'], errors='coerce')
    if '校正係数_単位' in df_instr.columns:
        mask_empty_u = df_instr['校正係数1_単位'].astype(str).str.strip().eq('') | df_instr['校正係数1_単位'].isna()
        df_instr.loc[mask_empty_u, '校正係数1_単位'] = (
            df_instr.loc[mask_empty_u, '校正係数_単位'].astype(str).replace('nan','').fillna('')
        )

    needed_instr = [
        "名称","型式","識別番号","貸出状況","貸出日","返却予定日",
        "所属","氏名","使用場所","使用用途","保管場所","備考"
    ] + NEW_FIELDS
    for c in needed_instr:
        if c not in df_instr.columns:
            df_instr[c] = "" if c not in DATE_FIELDS else pd.NaT

    # 日付・文字列の整形
    for dcol in DATE_FIELDS:
        if dcol in df_instr.columns:
            df_instr[dcol] = pd.to_datetime(df_instr[dcol], errors="coerce")
    if "識別番号" in df_instr.columns:
        df_instr["識別番号"] = df_instr["識別番号"].astype(str).str.strip()
    for _c in TEXT_FIELDS_RECOMMENDED:
        if _c in df_instr.columns:
            df_instr[_c] = df_instr[_c].astype(str).replace("nan", "").fillna("")

    # 貸出状況の表記ゆれ統一
    if '貸出状況' in df_instr.columns:
        _map = {'○':'〇','✕':'×','〇':'〇','×':'×'}
        df_instr['貸出状況'] = df_instr['貸出状況'].astype(str).str.strip().map(lambda v: _map.get(v, v))

    # 係数は数値化、単位は文字列として保証（float推定を防止）
    for coef_col in ["校正係数1","校正係数2"]:
        df_instr[coef_col] = pd.to_numeric(df_instr.get(coef_col, pd.Series(dtype='float')), errors="coerce")
    for unit_col in ["校正係数1_単位","校正係数2_単位"]:
        if unit_col not in df_instr.columns:
            df_instr[unit_col] = ""
        else:
            df_instr[unit_col] = (
                df_instr[unit_col]
                .astype(object)
                .where(pd.notna(df_instr[unit_col]), "")
                .astype(str)
            )

    # --- 熱電対 使用履歴 ---
    try:
        df_tc_usage = pd.read_excel(DATA_PATH, sheet_name=SHEET_TC_USAGE, engine="openpyxl")
    except Exception:
        df_tc_usage = pd.DataFrame(columns=["使用日","所属","氏名","用途","使用数"])
    if "使用日" in df_tc_usage.columns:
        df_tc_usage["使用日"] = pd.to_datetime(df_tc_usage["使用日"], errors="coerce")

    # --- 熱電対 在庫 ---
    try:
        df_tc_inv = pd.read_excel(DATA_PATH, sheet_name=SHEET_TC_INV, engine="openpyxl")
    except Exception:
        df_tc_inv = pd.DataFrame(columns=["種別","在庫","備考"])
    if "在庫" in df_tc_inv.columns:
        df_tc_inv["在庫"] = pd.to_numeric(df_tc_inv["在庫"], errors="coerce").fillna(0).astype(int)

    return df_instr, df_tc_usage, df_tc_inv


def save_data(df_instr, df_tc_usage, df_tc_inv):
    import pandas as _pd, tempfile as _tempfile, os as _os
    from pathlib import Path as _Path
    import time as _time

    target = _Path("app_data.xlsx")
    tmp_dir = target.parent if target.parent.exists() else _Path('.')
    attempts = 5
    last_err = None

    for _ in range(attempts):
        try:
            with _tempfile.NamedTemporaryFile(delete=False, dir=str(tmp_dir), suffix='.xlsx') as tmp:
                tmp_path = _Path(tmp.name)
                with _pd.ExcelWriter(tmp_path, engine="openpyxl", mode="w") as w:
                    df_instr.to_excel(w, sheet_name=SHEET_INSTR, index=False)
                    df_tc_usage.to_excel(w, sheet_name=SHEET_TC_USAGE, index=False)
                    df_tc_inv.to_excel(w, sheet_name=SHEET_TC_INV, index=False)
            if target.exists():
                try:
                    _os.replace(tmp_path, target)
                except PermissionError:
                    _time.sleep(0.4)
                    _os.replace(tmp_path, target)
            else:
                _os.replace(tmp_path, target)
            last_err = None
            break
        except PermissionError as e:
            last_err = e
            _time.sleep(0.8)
        except Exception as e:
            last_err = e
            break
        finally:
            try:
                if 'tmp_path' in locals() and tmp_path.exists():
                    tmp_path.unlink(missing_ok=True)
            except Exception:
                pass
    if last_err is not None:
        raise last_err


def status_icon(s):
    if s in ["〇", "○"]: return "✅"
    if s in ["×", "✕"]: return "❌"
    return "➖"

# アプリ設定：初期ページを『Dashboard(=貸出状況)』
st.set_page_config(page_title="計測器管理アプリ v4.4 r4c_r2", layout="wide")

# === Sidebar ===
with st.sidebar:
    st.title("計測器管理")
    mode = st.radio("モード選択", ["ユーザー", "管理者"], horizontal=True, key="mode")

    if 'is_admin' not in st.session_state:
        st.session_state.is_admin = False

    if mode == "管理者":
        import os
        def get_admin_code():
            env = os.getenv('ADMIN_CODE')
            if env: return env
            try:
                with open('admin_code.txt', 'r', encoding='utf-8') as f:
                    return f.readline().strip()
            except Exception:
                return 'basd4-admin'
        valid_code = get_admin_code()

        if st.session_state.is_admin:
            st.success("管理者としてログイン中")
            if st.button("ログアウト", key="admin_logout"):
                st.session_state.is_admin = False
                try: st.rerun()
                except AttributeError: st.experimental_rerun()
        else:
            admin_code_input = st.text_input("管理者コード", type="password", key="admin_code_input")
            if st.button("ログイン", key="admin_login"):
                if admin_code_input == valid_code:
                    st.session_state.is_admin = True
                    try: st.rerun()
                    except AttributeError: st.experimental_rerun()
                else:
                    st.error("管理者コードが違います。")

    if 'page' not in st.session_state:
        st.session_state.page = "Dashboard"
    page = st.radio("メニュー", ["Dashboard", "計測器一覧", "熱電対 在庫", "熱電対 使用履歴", "管理者"], key="page")

# === Load data ===
df_instr, df_tc_usage, df_tc_inv = load_data()

# === Dashboard（貸出状況） ===
if page == "Dashboard":
    st.header("貸出状況")
    today = pd.Timestamp(datetime.now().date())
    df_out = df_instr[df_instr["貸出状況"].isin(["×","✕"])]
    overdue = df_out[df_out["返却予定日"].notna() & (df_out["返却予定日"] < today)]
    due_today = df_out[df_out["返却予定日"].dt.date == today.date()]
    due_7 = df_out[df_out["返却予定日"].between(today, today + pd.Timedelta(days=7), inclusive="right")]

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("貸出中", len(df_out))
    c2.metric("期限切れ", len(overdue))
    c3.metric("本日期限", len(due_today))
    c4.metric("7日以内期限", len(due_7))

    st.subheader("期限切れ一覧")
    st.dataframe(overdue[["名称","型式","識別番号","氏名","所属","返却予定日","使用用途","使用場所"]], use_container_width=True)

# === 計測器一覧 ===
elif page == "計測器一覧":
    st.header("計測器一覧")
    cols = st.columns([2,2,2,2])
    with cols[0]:
        kw = st.text_input("キーワード（名称/型式/識別番号/メーカー）", key="kw")
    with cols[1]:
        stfilter = st.selectbox("ステータス", ["すべて","貸出可（〇）","貸出中（×）"], key="stfilter")
    with cols[2]:
        only_due = st.checkbox("期限切れのみ", key="only_due")
    with cols[3]:
        only_due7 = st.checkbox("7日以内の返却予定", key="only_due7")

    dfv = df_instr.copy()
    if kw:
        mask = (
            dfv["名称"].astype(str).str.contains(kw, case=False, na=False)
            | dfv["型式"].astype(str).str.contains(kw, case=False, na=False)
            | dfv["識別番号"].astype(str).str.contains(kw, case=False, na=False)
            | dfv["メーカー"].astype(str).str.contains(kw, case=False, na=False)
        )
        dfv = dfv[mask]

    if stfilter != "すべて":
        m = {"貸出可（〇）":["〇","○"], "貸出中（×）":["×","✕"]}
        dfv = dfv[dfv["貸出状況"].isin(m.get(stfilter, []))]

    today = pd.Timestamp(datetime.now().date())
    if only_due:
        dfv = dfv[dfv["返却予定日"].notna() & (dfv["返却予定日"] < today)]
    if only_due7:
        dfv = dfv[dfv["返却予定日"].between(today, today + pd.Timedelta(days=7), inclusive="right")]

    st.dataframe(
        dfv.assign(ステータス=dfv["貸出状況"].map(status_icon))[[
            "ステータス","名称","型式","識別番号","メーカー","購入日","校正期限","氏名","所属",
            "貸出日","返却予定日","使用用途","使用場所","保管場所","備考"
        ]],
        use_container_width=True
    )

    st.divider()
    st.subheader("貸出 / 返却（名称から選択）")
    name_choices = dfv["名称"].dropna().astype(str).unique().tolist()
    selected_name = st.selectbox("名称を選択", [""] + sorted(name_choices), key="sel_name")

    subset = df_instr[df_instr["名称"].astype(str) == selected_name]
    selected_id = None
    if selected_name:
        if len(subset) > 1:
            sub_opts = [f"{r['型式']} / {r['識別番号']}" for _, r in subset.iterrows()]
            sub_sel = st.selectbox("対象（同名が複数ある場合は選択）", sub_opts, key="sel_sub")
            if sub_sel:
                idx = sub_opts.index(sub_sel)
                selected_id = subset.iloc[idx]["識別番号"]
        else:
            if not subset.empty:
                selected_id = subset.iloc[0]["識別番号"]

    if selected_name and selected_id:
        t = df_instr[df_instr["識別番号"].astype(str) == str(selected_id)].iloc[0]
        st.write(f"**{t['名称']} / {t['型式']} / {t['識別番号']}** 現在: {t['貸出状況']}")
        c1, c2 = st.columns(2)
        with c1:
            st.markdown("**貸出**")
            with st.form(f"lend_{selected_id}"):
                所属 = st.text_input("所属", value=str(t.get("所属","")))
                氏名 = st.text_input("氏名", value=str(t.get("氏名","")))
                使用場所 = st.text_input("使用場所", value=str(t.get("使用場所","")))
                使用用途 = st.text_input("使用用途", value=str(t.get("使用用途","")))
                返却予定日 = st.date_input("返却予定日", value=datetime.now().date() + timedelta(days=7))
                submitted = st.form_submit_button("貸出登録")
                if submitted:
                    idx = df_instr.index[df_instr["識別番号"].astype(str) == str(selected_id)][0]
                    df_instr.at[idx, "貸出状況"] = "×"
                    df_instr.at[idx, "所属"] = 所属
                    df_instr.at[idx, "氏名"] = 氏名
                    df_instr.at[idx, "使用場所"] = 使用場所
                    df_instr.at[idx, "使用用途"] = 使用用途
                    df_instr.at[idx, "貸出日"] = pd.Timestamp(datetime.now())
                    df_instr.at[idx, "返却予定日"] = pd.Timestamp(返却予定日)
                    save_data(df_instr, df_tc_usage, df_tc_inv)
                    st.success("貸出を登録しました。")
                    st.cache_data.clear()
        with c2:
            st.markdown("**返却**")
            if st.button("返却処理", type="primary"):
                idx = df_instr.index[df_instr["識別番号"].astype(str) == str(selected_id)][0]
                df_instr.at[idx, "貸出状況"] = "〇"
                for c in ["所属","氏名","使用場所","使用用途"]:
                    df_instr.at[idx, c] = ""
                df_instr.at[idx, "貸出日"] = pd.NaT
                df_instr.at[idx, "返却予定日"] = pd.NaT
                save_data(df_instr, df_tc_usage, df_tc_inv)
                st.success("返却を処理しました。")
                st.cache_data.clear()

# === 熱電対 在庫 ===
elif page == "熱電対 在庫":
    st.header("熱電対 在庫")
    st.subheader("在庫一覧")
    st.caption("※ 管理者のみ編集可能（備考・在庫）。ユーザーは閲覧のみ。")
    edited = st.data_editor(
        df_tc_inv if st.session_state.is_admin else df_tc_inv.copy(),
        num_rows="fixed",
        use_container_width=True,
        column_config={
            "種別": st.column_config.TextColumn(disabled=not st.session_state.is_admin),
            "在庫": st.column_config.NumberColumn(min_value=0, step=1, disabled=not st.session_state.is_admin),
            "備考": st.column_config.TextColumn(disabled=not st.session_state.is_admin)
        },
        hide_index=True,
        key="inv_editor"
    )
    if st.session_state.is_admin and st.button("在庫表を保存"):
        save_data(df_instr, df_tc_usage, edited)
        st.success("在庫表（備考含む）を保存しました。")
        st.cache_data.clear()

    st.divider()
    st.subheader("入出庫フォーム")
    with st.form("io_form", clear_on_submit=True):
        種別 = st.selectbox("種別", edited["種別"].tolist())
        区分 = st.radio("区分", ["入庫", "出庫"], horizontal=True)
        if not st.session_state.is_admin:
            区分 = '出庫'
        数量 = st.number_input("数量", min_value=1, step=1)
        追加メモ = st.text_input("メモ（任意）")
        所属 = st.text_input("（出庫時）所属", value="")
        氏名 = st.text_input("（出庫時）氏名", value="")
        用途 = st.text_input("（出庫時）用途", value="")
        submitted = st.form_submit_button("実行")
        if submitted:
            df_inv = edited.copy()
            idx = df_inv.index[df_inv["種別"] == 種別][0]
            if 区分 == "入庫":
                df_inv.at[idx, "在庫"] = int(df_inv.at[idx, "在庫"]) + int(数量)
                if 追加メモ:
                    note = str(df_inv.at[idx, "備考"]).strip()
                    df_inv.at[idx, "備考"] = (note + "" if note else "") + f"[入庫] {datetime.now():%Y-%m-%d} {追加メモ}"
                st.success(f"{種別} を {数量} 本 入庫しました。")
            else:
                current = int(df_inv.at[idx, "在庫"])
                if 数量 > current:
                    st.error("在庫不足です。数量を見直してください。")
                    st.stop()
                df_inv.at[idx, "在庫"] = current - int(数量)
                row = {"使用日": pd.Timestamp(datetime.now().date()),
                       "所属": 所属, "氏名": 氏名, "用途": 用途 if 用途 else 種別,
                       "使用数": int(数量)}
                df_tc_usage = pd.concat([df_tc_usage, pd.DataFrame([row])], ignore_index=True)
                if 追加メモ:
                    note = str(df_inv.at[idx, "備考"]).strip()
                    df_inv.at[idx, "備考"] = (note + "" if note else "") + f"[出庫] {datetime.now():%Y-%m-%d} {追加メモ}"
                st.success(f"{種別} を {数量} 本 出庫しました。")
            save_data(df_instr, df_tc_usage, df_inv)
            st.cache_data.clear()

# === 熱電対 使用履歴 ===
elif page == "熱電対 使用履歴":
    st.header("熱電対 使用履歴")
    if df_tc_usage.empty:
        st.info("まだ使用履歴がありません。")
    else:
        st.dataframe(df_tc_usage.sort_values("使用日", ascending=False), use_container_width=True)

# === 管理者 ===
elif page == "管理者":
    st.header("管理者専用ページ")
    if not st.session_state.is_admin:
        st.error("管理者コードが必要です。サイドバーで '管理者' を選び、管理者コードを入力してください。")
        st.stop()
    st.success("管理者としてログイン中")

    # バックアップ/ダウンロード
    st.subheader("データバックアップ / ダウンロード")
    from io import BytesIO
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl', mode='w') as w:
        df_instr.to_excel(w, sheet_name=SHEET_INSTR, index=False)
        df_tc_usage.to_excel(w, sheet_name=SHEET_TC_USAGE, index=False)
        df_tc_inv.to_excel(w, sheet_name=SHEET_TC_INV, index=False)
    st.download_button("app_data.xlsx をダウンロード", data=buf.getvalue(), file_name="app_data_backup.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    st.divider()
    st.subheader("計測器マスタ編集（保存 / 削除）")
    names = df_instr["名称"].dropna().astype(str).unique().tolist()
    target_name = st.selectbox("名称を選択", [""] + sorted(names), key="admin_sel_name")

    row = None
    if target_name:
        subset = df_instr[df_instr["名称"].astype(str) == target_name]
        if len(subset) > 1:
            opts = [f"{r['型式']} / {r['識別番号']}" for _, r in subset.iterrows()]
            sel = st.selectbox("対象", opts, key="admin_sel_sub")
            if sel:
                idx_sel = opts.index(sel)
                row = subset.iloc[idx_sel]
        else:
            row = subset.iloc[0]

    if row is not None:
        st.info(f"編集対象: {row['名称']} / {row['型式']} / {row['識別番号']}")
        with st.form("edit_master"):
            名称 = st.text_input("名称", value=str(row.get("名称","")))
            型式 = st.text_input("型式", value=str(row.get("型式","")))
            メーカー = st.text_input("メーカー", value=str(row.get("メーカー","")))
            入力購入日 = st.checkbox("購入日を設定する", value=bool(pd.notna(row.get("購入日"))))
            if 入力購入日:
                購入日 = st.date_input("購入日", value=(row.get("購入日").date() if pd.notna(row.get("購入日")) else datetime.now().date()))
            else:
                購入日 = None
            入力校正 = st.checkbox("校正期限を設定する", value=bool(pd.notna(row.get("校正期限"))))
            if 入力校正:
                校正期限 = st.date_input("校正期限", value=(row.get("校正期限").date() if pd.notna(row.get("校正期限")) else datetime.now().date()))
            else:
                校正期限 = None
            保管場所 = st.text_input("保管場所", value=str(row.get("保管場所","")))
            備考 = st.text_area("備考", value=str(row.get("備考","")))
            c1, c2 = st.columns(2)
            submitted = c1.form_submit_button("保存")
            delete_req = c2.form_submit_button("削除")
            if submitted:
                idx = df_instr.index[df_instr['識別番号'].astype(str) == str(row['識別番号'])][0]
                _status_map = {"○":"〇", "✕":"×", "〇":"〇", "×":"×"}
                df_instr.at[idx, '貸出状況'] = _status_map.get(str(df_instr.at[idx, '貸出状況']).strip(), '〇')
                df_instr.at[idx, '名称'] = 名称
                df_instr.at[idx, '型式'] = 型式
                df_instr.at[idx, 'メーカー'] = メーカー
                df_instr.at[idx, '保管場所'] = 保管場所
                df_instr.at[idx, '備考'] = 備考
                df_instr.at[idx, '購入日'] = pd.Timestamp(購入日) if 購入日 else pd.NaT
                df_instr.at[idx, '校正期限'] = pd.Timestamp(校正期限) if 入力校正 and 校正期限 else pd.NaT
                save_data(df_instr, df_tc_usage, df_tc_inv)
                st.success("計測器マスタを保存しました。")
                st.cache_data.clear()
            if delete_req:
                with st.modal("この計測器を削除します。よろしいですか？"):
                    st.warning(f"削除対象: {row['名称']} / {row['型式']} / {row['識別番号']}")
                    cc1, cc2 = st.columns(2)
                    do_del = cc1.button("削除を確定", type="primary")
                    cancel = cc2.button("やめる")
                    if do_del:
                        df_instr = df_instr[df_instr['識別番号'].astype(str) != str(row['識別番号'])].reset_index(drop=True)
                        save_data(df_instr, df_tc_usage, df_tc_inv)
                        st.success("削除しました。")
                        st.cache_data.clear()
                    elif cancel:
                        st.info("削除をキャンセルしました。")

    # ====== 計測器マスタの一括編集（管理者） ======
    st.divider()
    st.subheader("計測器マスタの一括編集（管理者）")

    if 'master_edit_buf' not in st.session_state:
        st.session_state.master_edit_buf = df_instr.copy()
    if 'master_dirty' not in st.session_state:
        st.session_state.master_dirty = False
    if 'last_autosave_text' not in st.session_state:
        st.session_state.last_autosave_text = '—'
    if 'autosave_lock' not in st.session_state:
        st.session_state.autosave_lock = threading.Lock()

    # オートセーブ設定
    with st.expander("オートセーブ（ベータ）", expanded=False):
        st.checkbox("オートセーブを有効にする", key="autosave_enabled",
                    help="『編集を反映（未保存）』後、一定間隔でExcelへ自動保存します（rerunしないためスクロールは維持）。")
        st.select_slider("保存間隔（秒）", options=[15, 30, 60, 120, 300], value=60, key="autosave_interval")
        st.caption(f"最終自動保存: {st.session_state.get('last_autosave_text', '—')}")

    def start_autosave_thread():
        stop_event = threading.Event()
        st.session_state.autosave_stop = stop_event
        def autosave_loop():
            while not stop_event.is_set():
                time.sleep(st.session_state.get('autosave_interval', 60))
                if not st.session_state.get('autosave_enabled', False):
                    continue
                if not st.session_state.get('master_dirty', False):
                    continue
                with st.session_state.autosave_lock:
                    df_snapshot = st.session_state.master_edit_buf.copy()
                    try:
                        save_data(df_snapshot, df_tc_usage, df_tc_inv)
                        st.session_state.master_dirty = False
                        st.session_state.last_autosave_text = time.strftime("%Y-%m-%d %H:%M:%S")
                    except Exception as e:
                        st.session_state.last_autosave_text = f"エラー: {e}"
        t = threading.Thread(target=autosave_loop, daemon=True)
        t.start()
        st.session_state.autosave_thread = t

    if st.session_state.get('autosave_enabled', False):
        if 'autosave_thread' not in st.session_state:
            start_autosave_thread()
    else:
        if 'autosave_thread' in st.session_state:
            try: st.session_state.autosave_stop.set()
            except Exception: pass
            st.session_state.pop('autosave_thread', None)

    # 編集用コピー
    df_master_edit = st.session_state.master_edit_buf.copy()

    # 表記ゆれ統一
    _map_ui = {'○':'〇','✕':'×','〇':'〇','×':'×'}
    if '貸出状況' in df_master_edit.columns:
        df_master_edit['貸出状況'] = df_master_edit['貸出状況'].astype(str).str.strip().map(lambda v: _map_ui.get(v, '〇'))
    if '識別番号' in df_master_edit.columns:
        df_master_edit['識別番号'] = df_master_edit['識別番号'].astype(str).str.strip()

    # 単位列は最終的に文字列へ矯正（エディタ互換）
    for unit_col in ["校正係数1_単位","校正係数2_単位"]:
        if unit_col in df_master_edit.columns:
            df_master_edit[unit_col] = (
                df_master_edit[unit_col].astype(object).where(pd.notna(df_master_edit[unit_col]), "").astype(str)
            )

    # 表示列ホワイトリスト
    whitelist = [
        '名称','型式','識別番号','貸出状況',
        '貸出日','返却予定日',
        '所属','氏名','使用場所','使用用途',
        '保管場所','備考',
        'メーカー','購入日','校正期限',
        '校正係数1','校正係数1_単位',
        '校正係数2','校正係数2_単位',
    ]
    visible_cols = [c for c in whitelist if c in df_master_edit.columns]
    df_master_for_view = df_master_edit[visible_cols].copy()

    with st.form("master_bulk_form"):
        edited_master = st.data_editor(
            df_master_for_view,
            num_rows="dynamic",
            use_container_width=True,
            hide_index=True,
            key="editor_master",
            column_config={
                "名称": st.column_config.TextColumn(required=True),
                "型式": st.column_config.TextColumn(required=True),
                "識別番号": st.column_config.TextColumn(required=True, help="一意のID。重複不可"),
                "貸出状況": st.column_config.SelectboxColumn(options=["〇","×"], help="〇=貸出可、×=貸出中"),
                "貸出日": st.column_config.DateColumn(format="YYYY-MM-DD"),
                "返却予定日": st.column_config.DateColumn(format="YYYY-MM-DD"),
                "所属": st.column_config.TextColumn(),
                "氏名": st.column_config.TextColumn(),
                "使用場所": st.column_config.TextColumn(),
                "使用用途": st.column_config.TextColumn(),
                "保管場所": st.column_config.TextColumn(),
                "備考": st.column_config.TextColumn(),
                "メーカー": st.column_config.TextColumn(),
                "購入日": st.column_config.DateColumn(format="YYYY-MM-DD"),
                "校正期限": st.column_config.DateColumn(format="YYYY-MM-DD"),
                # 2組の校正係数（有効数字6桁表示 + 6桁小数刻み）
                "校正係数1": st.column_config.NumberColumn(min_value=0.0, step=0.000001, format="%.6g", help="校正係数1（実数、表示は有効数字6桁）"),
                "校正係数1_単位": st.column_config.TextColumn(help="校正係数1の単位（例: 1/K, mV/V, %, ppm）"),
                "校正係数2": st.column_config.NumberColumn(min_value=0.0, step=0.000001, format="%.6g", help="校正係数2（実数、表示は有効数字6桁）"),
                "校正係数2_単位": st.column_config.TextColumn(help="校正係数2の単位（例: 1/K, mV/V, %, ppm）"),
            },
        )

        col1, col2, col3, col4 = st.columns(4)
        apply_clicked = col1.form_submit_button("編集を反映（未保存）")
        save_clicked  = col2.form_submit_button("Excelへ保存", type="primary")
        discard_clicked = col3.form_submit_button("未保存編集を破棄して元に戻す")
        quick_clicked  = col4.form_submit_button("反映して即保存（推奨）")

        def ensure_df(obj):
            if isinstance(obj, pd.DataFrame):
                return obj
            elif isinstance(obj, dict):
                if 'data' in obj and isinstance(obj['data'], list):
                    return pd.DataFrame(obj['data'])
                return pd.DataFrame(obj)
            else:
                return pd.DataFrame(obj)

        if apply_clicked or save_clicked or quick_clicked:
            tmp_df = ensure_df(edited_master).copy()

            # === reset_index 衝突を安全に解決 ===
            idx_name = getattr(tmp_df.index, 'name', None)
            if idx_name == '識別番号' and '識別番号' in tmp_df.columns:
                tmp_df = tmp_df.reset_index(drop=True)
            elif idx_name == '識別番号':
                tmp_df = tmp_df.reset_index(drop=False)
            elif '識別番号' not in tmp_df.columns:
                tmp_df = tmp_df.reset_index(drop=False)
                if 'index' in tmp_df.columns:
                    tmp_df = tmp_df.rename(columns={'index': '識別番号'})

            # 日付列整形
            for c in DATE_FIELDS:
                if c in tmp_df.columns:
                    tmp_df[c] = pd.to_datetime(tmp_df[c], errors='coerce')

            # 校正係数（2組） 数値化
            for coef_col in ['校正係数1','校正係数2']:
                if coef_col in tmp_df.columns:
                    tmp_df[coef_col] = pd.to_numeric(tmp_df[coef_col], errors='coerce')

            # 有効数字6桁へ丸め（保存前正規化）
            for coef_col in ['校正係数1','校正係数2']:
                if coef_col in tmp_df.columns:
                    tmp_df[coef_col] = tmp_df[coef_col].map(lambda v: _round_sig(v, 6))

            # ステータス表記ゆれ統一
            if '貸出状況' in tmp_df.columns:
                status_map = {"○":"〇", "✕":"×", "〇":"〇", "×":"×"}
                tmp_df['貸出状況'] = tmp_df['貸出状況'].astype(str).str.strip().map(lambda v: status_map.get(v, '〇'))

            # === アップサート安全版マージ（新規IDも受け入れ、KeyError回避） ===
            with st.session_state.autosave_lock:
                base = st.session_state.master_edit_buf.copy()
                key = '識別番号'

                # キー正規化
                base[key]   = base[key].astype(str).str.strip()
                tmp_df[key] = tmp_df[key].astype(str).str.strip()

                # バリデーション
                if (tmp_df[key] == '').any():
                    st.error('識別番号に空の行があります。保存できません。')
                    st.stop()
                dupes = tmp_df[key][tmp_df[key].duplicated()].unique().tolist()
                if dupes:
                    st.error(f'識別番号が重複しています: {dupes}')
                    st.stop()

                # index 化と union
                base = base.set_index(key, drop=False)
                tmp_df = tmp_df.set_index(key, drop=False)
                base.index   = base.index.astype(str).str.strip()
                tmp_df.index = tmp_df.index.astype(str).str.strip()
                base = base.reindex(base.index.union(tmp_df.index))

                # 列単位で安全上書き
                for col in tmp_df.columns:
                    base.loc[tmp_df.index, col] = tmp_df[col].values

                base = base.reset_index(drop=True)
                st.session_state.master_edit_buf = base.copy()
                st.session_state.master_dirty = True
            st.success('未保存の編集内容をバッファに反映しました。')

            # 保存
            if quick_clicked or save_clicked or (st.session_state.get('autosave_enabled', False) and apply_clicked):
                try:
                    save_data(st.session_state.master_edit_buf, df_tc_usage, df_tc_inv)
                    st.success('計測器マスタを保存しました。')
                    st.session_state.master_dirty = False
                    st.session_state.last_autosave_text = time.strftime("%Y-%m-%d %H:%M:%S")
                    st.cache_data.clear()
                except Exception as e:
                    st.error(f"保存時にエラー: {e}")

        if discard_clicked:
            with st.session_state.autosave_lock:
                st.session_state.master_edit_buf = df_instr.copy()
                st.session_state.master_dirty = False
            st.info("編集バッファを最新データで再読込しました。")

    # === 貸出状況の一括更新（管理者） ===
    st.divider()
    st.subheader("貸出状況の一括更新（管理者）")
    labels = [f"{r['名称']} / {r['型式']} / {r['識別番号']} 現在:{r['貸出状況']}" for _, r in df_instr.iterrows()]
    values = df_instr['識別番号'].astype(str).tolist()
    pick = st.multiselect("対象を選択", options=values, format_func=lambda v: labels[values.index(str(v))], key="bulk_pick")

    col_bs1, col_bs2 = st.columns([1,1])
    with col_bs1:
        new_status = st.radio("更新後ステータス", ["〇","×"], index=0, horizontal=True, key="bulk_status")
    with col_bs2:
        auto_clear = st.checkbox("返却時に 所属・氏名・使用場所・使用用途・貸出日・返却予定日 を自動クリア", value=True, key="bulk_clear")

    lend_col1, lend_col2, lend_col3 = st.columns(3)
    with lend_col1:
        入力所属 = st.text_input("（×に更新）所属", value="", key="bulk_aff")
    with lend_col2:
        入力氏名 = st.text_input("（×に更新）氏名", value="", key="bulk_name")
    with lend_col3:
        入力返却予定日 = st.date_input("（×に更新）返却予定日", value=datetime.now().date() + timedelta(days=7), key="bulk_due")

    if st.button("一括更新を実行", type="primary", key="bulk_exec"):
        if not pick:
            st.warning("対象が選択されていません。")
        else:
            df_new = df_instr.copy()
            for idv in pick:
                idxs = df_new.index[df_new['識別番号'].astype(str) == str(idv)]
                if not len(idxs):
                    continue
                idx = idxs[0]
                if new_status == '〇':
                    df_new.at[idx, '貸出状況'] = '〇'
                    if auto_clear:
                        for c in ['所属','氏名','使用場所','使用用途']:
                            df_new.at[idx, c] = ''
                        df_new.at[idx, '貸出日'] = pd.NaT
                        df_new.at[idx, '返却予定日'] = pd.NaT
                else:
                    if not 入力所属 or not 入力氏名:
                        st.error('×への更新には「所属」と「氏名」が必要です。')
                        st.stop()
                    df_new.at[idx, '貸出状況'] = '×'
                    df_new.at[idx, '所属'] = 入力所属
                    df_new.at[idx, '氏名'] = 入力氏名
                    df_new.at[idx, '貸出日'] = pd.Timestamp(datetime.now())
                    df_new.at[idx, '返却予定日'] = pd.Timestamp(入力返却予定日)
            save_data(df_new, df_tc_usage, df_tc_inv)
            st.success(f"{len(pick)} 件のレコードを更新しました。")
            st.cache_data.clear()

    # === 熱電対在庫の一括編集 ===
    st.divider()
    st.subheader("熱電対在庫の一括編集")
    edited_inv = st.data_editor(
        df_tc_inv,
        num_rows="dynamic",
        use_container_width=True,
        column_config={
            "種別": st.column_config.TextColumn(),
            "在庫": st.column_config.NumberColumn(min_value=0, step=1),
            "備考": st.column_config.TextColumn()
        },
        hide_index=True,
        key="inv_bulk_editor"
    )
    if st.button("在庫テーブルを保存（管理者）", key="inv_bulk_save"):
        save_data(df_instr, df_tc_usage, edited_inv)
        st.success("在庫テーブルを保存しました。")
        st.cache_data.clear()

    # === 不要列の削除（Excelからも消す） ===
    st.divider()
    st.subheader("不要列の削除（Excelからも消す）")
    st.caption("※ 'Unnamed:*'や空ヘッダ列は自動で候補に含めます。主要な必須列は削除できないよう保護します。")
    essential_cols = set([
        '名称','型式','識別番号','貸出状況','貸出日','返却予定日',
        '所属','氏名','使用場所','使用用途','保管場所','備考',
        'メーカー','購入日','校正期限',
        '校正係数1','校正係数1_単位','校正係数2','校正係数2_単位'
    ])
    current_cols = list(df_instr.columns)
    auto_candidates = [c for c in current_cols if str(c).strip()=='' or str(c).lower().startswith('unnamed:')]
    selectable = [c for c in current_cols if c not in essential_cols]
    cols_to_drop = st.multiselect("削除する列を選択", options=selectable, default=auto_candidates,
                             help="Excelファイルから完全に削除します。※元に戻すにはバックアップから復元してください。")
    if st.button("選択した列を削除して保存", type="primary", key="drop_cols_and_save"):
        if not cols_to_drop:
            st.warning("削除対象が選択されていません。")
        else:
            protected = [c for c in cols_to_drop if c in essential_cols]
            if protected:
                st.error("必須列は削除できません: " + ", ".join(protected))
            else:
                df_new = df_instr.drop(columns=cols_to_drop, errors='ignore')
                save_data(df_new, df_tc_usage, df_tc_inv)
                st.success("列を削除して保存しました: " + ", ".join(cols_to_drop))
                st.cache_data.clear()
