
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

NEW_FIELDS = ["メーカー","購入日","校正期限"]
DATE_FIELDS = ["貸出日","返却予定日","購入日","校正期限"]
TEXT_FIELDS_RECOMMENDED = ["名称","型式","所属","氏名","使用場所","使用用途","保管場所","備考","メーカー"]

@st.cache_data(ttl=60)
def load_data():
    df_instr = pd.read_excel(DATA_PATH, sheet_name=SHEET_INSTR, engine="openpyxl")
    needed_instr = [
        "名称","型式","識別番号","貸出状況","貸出日","返却予定日",
        "所属","氏名","使用場所","使用用途","保管場所","備考"
    ] + NEW_FIELDS
    for c in needed_instr:
        if c not in df_instr.columns:
            df_instr[c] = "" if c not in DATE_FIELDS else pd.NaT
    for dcol in DATE_FIELDS:
        if dcol in df_instr.columns:
            df_instr[dcol] = pd.to_datetime(df_instr[dcol], errors="coerce")
    if "識別番号" in df_instr.columns:
        df_instr["識別番号"] = df_instr["識別番号"].astype(str).str.strip()
    for _c in TEXT_FIELDS_RECOMMENDED:
        if _c in df_instr.columns:
            df_instr[_c] = df_instr[_c].astype(str).replace("nan", "").fillna("")
    if '貸出状況' in df_instr.columns:
        _map = {'○':'〇','✕':'×','〇':'〇','×':'×'}
        df_instr['貸出状況'] = df_instr['貸出状況'].astype(str).str.strip().map(lambda v: _map.get(v, v))
    try:
        df_tc_usage = pd.read_excel(DATA_PATH, sheet_name=SHEET_TC_USAGE, engine="openpyxl")
    except Exception:
        df_tc_usage = pd.DataFrame(columns=["使用日","所属","氏名","用途","使用数"])
    if "使用日" in df_tc_usage.columns:
        df_tc_usage["使用日"] = pd.to_datetime(df_tc_usage["使用日"], errors="coerce")
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

st.set_page_config(page_title="計測器管理アプリ v4.4 r4c_r1 (フォーム編集 + オートセーブ/安定化)", layout="wide")

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
            st.caption("※ 管理者コードは admin_code.txt または環境変数 ADMIN_CODE で変更可能")
            if st.button("ログアウト", key="admin_logout"):
                st.session_state.is_admin = False
                try:
                    st.rerun()
                except AttributeError:
                    st.experimental_rerun()
        else:
            admin_code_input = st.text_input("管理者コード", type="password", key="admin_code_input")
            if st.button("ログイン", key="admin_login"):
                if admin_code_input == valid_code:
                    st.session_state.is_admin = True
                    try:
                        st.rerun()
                    except AttributeError:
                        st.experimental_rerun()
                else:
                    st.error("管理者コードが違います。")
                    st.caption("※ 管理者コードは admin_code.txt または環境変数 ADMIN_CODE で変更可能")
    if 'page' not in st.session_state:
        st.session_state.page = "計測器一覧"
    page = st.radio("メニュー", ["Dashboard", "計測器一覧", "熱電対 在庫", "熱電対 使用履歴", "管理者"], key="page")

# Load data
df_instr, df_tc_usage, df_tc_inv = load_data()

# ---- ページ（Dashboard/一覧/在庫/履歴）は r4c と同等。省略なく実装するのが理想だが、
#   問題の箇所は管理者>一括編集なので、そこを完全版で記述。 ----

if page == "管理者":
    st.header("管理者専用ページ")
    if not st.session_state.is_admin:
        st.error("管理者コードが必要です。サイドバーで '管理者' を選び、管理者コードを入力してください。")
        st.stop()
    st.success("管理者としてログイン中")

    # === オートセーブ状態の初期化 ===
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
        st.checkbox("オートセーブを有効にする", key="autosave_enabled")
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
            try:
                st.session_state.autosave_stop.set()
            except Exception:
                pass
            st.session_state.pop('autosave_thread', None)

    # ====== 一括編集テーブル ======
    df_master_edit = st.session_state.master_edit_buf.copy()
    _map_ui = {'○':'〇','✕':'×','〇':'〇','×':'×'}
    if '貸出状況' in df_master_edit.columns:
        df_master_edit['貸出状況'] = (
            df_master_edit['貸出状況'].astype(str).str.strip().map(lambda v: _map_ui.get(v, '〇'))
        )
    if '識別番号' in df_master_edit.columns:
        df_master_edit['識別番号'] = df_master_edit['識別番号'].astype(str).str.strip()
        df_master_edit = df_master_edit.set_index('識別番号', drop=False)

    with st.form("master_bulk_form"):
        edited_master = st.data_editor(
            df_master_edit,
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
            },
        )
        c1, c2, c3, c4 = st.columns(4)
        apply_clicked   = c1.form_submit_button("編集を反映（未保存）")
        save_clicked    = c2.form_submit_button("Excelへ保存", type="primary")
        discard_clicked = c3.form_submit_button("未保存編集を破棄して元に戻す")
        quick_clicked   = c4.form_submit_button("反映して即保存（推奨）")

    # === ここが修正の肝 ===
    # st.data_editor の戻り値 edited_master は DataFrame のはずだが、環境によっては dict が session_state に入ることがある。
    # その場合でも確実に DataFrame 化してから処理する。
    def ensure_df(obj):
        if isinstance(obj, pd.DataFrame):
            return obj
        elif isinstance(obj, dict):
            # データエディタが dict を返したケース: 可能なら 'data' キーを解釈、それ以外は DataFrame() で受ける
            if 'data' in obj and isinstance(obj['data'], list):
                return pd.DataFrame(obj['data'])
            return pd.DataFrame(obj)
        else:
            return pd.DataFrame(obj)

    if apply_clicked or save_clicked or quick_clicked:
        tmp_df = ensure_df(edited_master).copy()
        # index を列に戻す
        if tmp_df.index.name == '識別番号' or '識別番号' not in tmp_df.columns:
            tmp_df = tmp_df.reset_index(drop=False)
        else:
            tmp_df = tmp_df.reset_index(drop=True)
        # バリデーション
        if tmp_df['識別番号'].isna().any() or (tmp_df['識別番号'].astype(str).str.strip() == '').any():
            st.error('識別番号は空にできません。')
        elif tmp_df['識別番号'].astype(str).duplicated().any():
            st.error('識別番号が重複しています。重複を解消してください。')
        else:
            for c in DATE_FIELDS:
                if c in tmp_df.columns:
                    tmp_df[c] = pd.to_datetime(tmp_df[c], errors='coerce')
            _status_map = {"○":"〇", "✕":"×", "〇":"〇", "×":"×"}
            if '貸出状況' in tmp_df.columns:
                tmp_df['貸出状況'] = tmp_df['貸出状況'].map(lambda v: _status_map.get(str(v).strip(), '〇'))
            if '識別番号' in tmp_df.columns:
                tmp_df['識別番号'] = tmp_df['識別番号'].astype(str).str.strip()
            with st.session_state.autosave_lock:
                st.session_state.master_edit_buf = tmp_df.copy()
            st.session_state.master_dirty = True
            st.success('未保存の編集内容をバッファに反映しました。')
            if quick_clicked or save_clicked or (st.session_state.get('autosave_enabled', False) and apply_clicked):
                try:
                    save_data(tmp_df, df_tc_usage, df_tc_inv)
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

