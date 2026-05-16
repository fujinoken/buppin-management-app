import streamlit as st
import pandas as pd
from datetime import date, datetime
from pathlib import Path
import calendar

st.set_page_config(page_title="物品管理アプリ Ver1.5.1", layout="wide")

DATA = Path("data")
DATA.mkdir(exist_ok=True)

USERS = DATA / "users.xlsx"
ITEMS = DATA / "items.xlsx"
USAGE = DATA / "usage.xlsx"
STOCK = DATA / "stock.xlsx"
ACCOUNTS_FILE = DATA / "accounts.xlsx"
DASH_MEMO = DATA / "dashboard_memo.txt"

USER_COLS = ["利用者ID", "利用者名", "請求先", "備考"]
ITEM_COLS = ["物品ID", "物品名", "単価", "最低在庫", "FEED商品URL", "備考"]
USAGE_COLS = ["記録ID", "日付", "利用者", "物品", "数量", "単価", "金額", "備考", "登録日時"]
STOCK_COLS = ["物品", "現在庫", "更新日時"]
ACCOUNT_COLS = ["ログインID", "パスワード", "権限", "表示名"]

ADMIN_MENUS = [
    "管理ダッシュボード",
    "使用記録 登録",
    "使用記録 検索・更新・削除",
    "現在庫 登録・更新",
    "月間集計",
    "請求書作成",
    "FEED発注候補",
    "利用者マスタ 登録・更新・削除",
    "物品マスタ 登録・更新・削除",
    "ログイン設定",
    "データ確認",
]

STAFF_MENUS = [
    "使用記録 登録",
    "使用記録 検索・更新・削除",
]


def now_id(prefix):
    return f"{prefix}{datetime.now().strftime('%Y%m%d%H%M%S%f')}"


def safe_int(v):
    try:
        if pd.isna(v) or v == "":
            return 0
        return int(float(v))
    except Exception:
        return 0


def load_df(path, cols):
    if path.exists():
        df = pd.read_excel(path)
        for c in cols:
            if c not in df.columns:
                df[c] = ""
        df = df[cols].fillna("")
        df.to_excel(path, index=False)
        return df

    df = pd.DataFrame(columns=cols)
    df.to_excel(path, index=False)
    return df


def save_df(df, path, cols):
    df = df[cols].fillna("")
    df.to_excel(path, index=False)


def init_accounts():
    if not ACCOUNTS_FILE.exists():
        df = pd.DataFrame([
            {"ログインID": "admin", "パスワード": "admin123", "権限": "管理者", "表示名": "管理者"},
            {"ログインID": "staff", "パスワード": "staff123", "権限": "職員", "表示名": "職員"},
        ], columns=ACCOUNT_COLS)
        df.to_excel(ACCOUNTS_FILE, index=False)


def load_accounts():
    init_accounts()
    df = pd.read_excel(ACCOUNTS_FILE)
    for c in ACCOUNT_COLS:
        if c not in df.columns:
            df[c] = ""
    df = df[ACCOUNT_COLS].fillna("")
    df.to_excel(ACCOUNTS_FILE, index=False)
    return df


def save_accounts(df):
    df = df[ACCOUNT_COLS].fillna("")
    df.to_excel(ACCOUNTS_FILE, index=False)


def login_screen():
    st.title("📦 物品管理アプリ Ver1.5.1")
    st.subheader("ログイン")

    with st.form("login_form"):
        login_id = st.text_input("ログインID")
        password = st.text_input("パスワード", type="password")
        ok = st.form_submit_button("ログイン")

    if ok:
        accounts = load_accounts()
        hit = accounts[
            (accounts["ログインID"].astype(str) == str(login_id)) &
            (accounts["パスワード"].astype(str) == str(password))
        ]

        if not hit.empty:
            account = hit.iloc[0]
            st.session_state["logged_in"] = True
            st.session_state["login_id"] = str(account["ログインID"])
            st.session_state["role"] = str(account["権限"])
            st.session_state["user_name"] = str(account["表示名"])
            st.success("ログインしました。")
            st.rerun()
        else:
            st.error("ログインIDまたはパスワードが違います。")

    st.info("初期設定：管理者 admin / admin123　職員 staff / staff123")


def require_login():
    if "logged_in" not in st.session_state:
        st.session_state["logged_in"] = False

    if not st.session_state["logged_in"]:
        login_screen()
        st.stop()


def logout_button():
    with st.sidebar:
        st.markdown("---")
        st.write(f"ログイン：{st.session_state.get('user_name', '')}")
        st.write(f"権限：{st.session_state.get('role', '')}")
        if st.button("ログアウト"):
            st.session_state.clear()
            st.rerun()


def sync_stock(items_df, stock_df):
    rows = []

    for _, item in items_df.iterrows():
        name = str(item["物品名"]).strip()
        if not name:
            continue

        old = stock_df[stock_df["物品"].astype(str) == name]

        if not old.empty:
            rows.append({
                "物品": name,
                "現在庫": safe_int(old.iloc[0]["現在庫"]),
                "更新日時": old.iloc[0].get("更新日時", "")
            })
        else:
            rows.append({
                "物品": name,
                "現在庫": 0,
                "更新日時": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            })

    return pd.DataFrame(rows, columns=STOCK_COLS)


def load_memo():
    if DASH_MEMO.exists():
        return DASH_MEMO.read_text(encoding="utf-8")
    return ""


def save_memo(text):
    DASH_MEMO.write_text(text, encoding="utf-8")


def make_stock_status(stock_df, items_df):
    if stock_df.empty or items_df.empty:
        return pd.DataFrame(columns=[
            "物品", "現在庫", "最低在庫", "不足数", "状態", "単価", "FEED商品URL"
        ])

    merged = stock_df.merge(items_df, left_on="物品", right_on="物品名", how="left")
    merged["現在庫"] = merged["現在庫"].apply(safe_int)
    merged["最低在庫"] = merged["最低在庫"].apply(safe_int)
    merged["単価"] = merged["単価"].apply(safe_int)
    merged["不足数"] = merged["最低在庫"] - merged["現在庫"]

    def status(row):
        current = safe_int(row["現在庫"])
        minimum = safe_int(row["最低在庫"])

        if current <= 0:
            return "🔴 在庫0"
        if current <= minimum:
            return "🔴 発注必要"
        if current <= minimum + max(2, minimum // 2):
            return "🟡 注意"
        return "🟢 OK"

    merged["状態"] = merged.apply(status, axis=1)

    return merged[[
        "物品", "現在庫", "最低在庫", "不足数", "状態", "単価", "FEED商品URL"
    ]]


def month_schedule():
    today = date.today()
    last_day = calendar.monthrange(today.year, today.month)[1]

    return pd.DataFrame([
        {"時期": "月初 1〜5日", "やること": "棚卸・現在庫確認", "目的": "前月のズレを直す", "状態": "今月の土台作り"},
        {"時期": "10日前後", "やること": "中間在庫チェック", "目的": "急な減りに気づく", "状態": "ショート予防"},
        {"時期": "20日前後", "やること": "FEED発注候補確認", "目的": "月末前に不足を防ぐ", "状態": "発注判断"},
        {"時期": f"月末 {last_day}日前後", "やること": "利用者別請求確認", "目的": "月末請求を作成", "状態": "請求処理"},
    ])


require_login()

users = load_df(USERS, USER_COLS)
items = load_df(ITEMS, ITEM_COLS)
usage = load_df(USAGE, USAGE_COLS)
stock = load_df(STOCK, STOCK_COLS)

if not users.empty:
    mask = users["利用者ID"].astype(str).str.strip() == ""
    users.loc[mask, "利用者ID"] = [now_id("U") for _ in range(mask.sum())]

if not items.empty:
    mask = items["物品ID"].astype(str).str.strip() == ""
    items.loc[mask, "物品ID"] = [now_id("I") for _ in range(mask.sum())]

if not usage.empty:
    mask = usage["記録ID"].astype(str).str.strip() == ""
    usage.loc[mask, "記録ID"] = [now_id("R") for _ in range(mask.sum())]

save_df(users, USERS, USER_COLS)
save_df(items, ITEMS, ITEM_COLS)
save_df(usage, USAGE, USAGE_COLS)

stock = sync_stock(items, stock)
save_df(stock, STOCK, STOCK_COLS)

st.title("📦 物品管理アプリ Ver1.5.1")
st.caption("ログイン設定対応／職員権限制限／在庫ショート予防／月末請求／FEED発注補助")

role = st.session_state.get("role", "職員")
available_menus = ADMIN_MENUS if role == "管理者" else STAFF_MENUS

menu = st.sidebar.radio("メニュー", available_menus)
logout_button()


if menu == "管理ダッシュボード":
    st.subheader("📋 物品担当 管理ダッシュボード")

    stock_status = make_stock_status(stock, items)

    total_items = len(items)
    stock_zero = len(stock_status[stock_status["状態"] == "🔴 在庫0"]) if not stock_status.empty else 0
    order_needed = len(stock_status[stock_status["状態"] == "🔴 発注必要"]) if not stock_status.empty else 0

    work_usage = usage.copy()
    if not work_usage.empty:
        work_usage["日付"] = pd.to_datetime(work_usage["日付"], errors="coerce")
        this_month = date.today().strftime("%Y-%m")
        month_usage = work_usage[work_usage["日付"].dt.strftime("%Y-%m") == this_month].copy()
        month_usage["金額"] = month_usage["金額"].apply(safe_int)
        month_total = int(month_usage["金額"].sum())
    else:
        month_usage = pd.DataFrame(columns=USAGE_COLS)
        month_total = 0

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("登録物品数", f"{total_items}件")
    col2.metric("在庫0", f"{stock_zero}件")
    col3.metric("発注必要", f"{order_needed}件")
    col4.metric("今月物品費", f"{month_total:,}円")

    st.markdown("---")
    left, right = st.columns([1.3, 1])

    with left:
        st.markdown("### 🔴 発注・注意が必要な物品")
        if stock_status.empty:
            st.info("物品マスタと在庫を登録してください。")
        else:
            alert = stock_status[stock_status["状態"].isin(["🔴 在庫0", "🔴 発注必要", "🟡 注意"])].copy()
            if alert.empty:
                st.success("現在、在庫ショートの注意物品はありません。")
            else:
                st.dataframe(alert[["状態", "物品", "現在庫", "最低在庫", "不足数", "FEED商品URL"]], use_container_width=True)

                st.markdown("#### FEED商品ページ")
                for _, r in alert.iterrows():
                    url = str(r.get("FEED商品URL", "")).strip()
                    if url.startswith("http"):
                        st.link_button(f"{r['物品']} を開く", url)

    with right:
        st.markdown("### 🗓 月間注文スケジュール")
        st.dataframe(month_schedule(), use_container_width=True, hide_index=True)

    st.markdown("---")
    c1, c2 = st.columns(2)

    with c1:
        st.markdown("### 📊 今月の物品別使用量")
        if month_usage.empty:
            st.info("今月の使用記録はまだありません。")
        else:
            month_usage["数量"] = month_usage["数量"].apply(safe_int)
            item_summary = month_usage.groupby("物品")[["数量", "金額"]].sum().reset_index()
            st.dataframe(item_summary.sort_values("数量", ascending=False), use_container_width=True)

    with c2:
        st.markdown("### 👤 今月の利用者別物品費")
        if month_usage.empty:
            st.info("今月の使用記録はまだありません。")
        else:
            user_summary = month_usage.groupby("利用者")[["金額"]].sum().reset_index()
            st.dataframe(user_summary.sort_values("金額", ascending=False), use_container_width=True)

    st.markdown("---")
    st.markdown("### 📝 物品担当メモ・申し送り")
    memo = load_memo()
    new_memo = st.text_area("次回発注予定日、注意物品、申し送りなどを記録できます。", value=memo, height=180)

    if st.button("担当メモを保存"):
        save_memo(new_memo)
        st.success("担当メモを保存しました。")


elif menu == "使用記録 登録":
    st.subheader("使用記録 登録")

    if users.empty:
        st.warning("先に利用者マスタを登録してください。")
    elif items.empty:
        st.warning("先に物品マスタを登録してください。")
    else:
        with st.form("usage_create"):
            d = st.date_input("日付", date.today())
            user = st.selectbox("利用者", users["利用者名"].dropna().astype(str))
            item = st.selectbox("物品", items["物品名"].dropna().astype(str))
            qty = st.number_input("数量", min_value=1, max_value=9999, value=1)
            note = st.text_input("備考")

            item_row = items[items["物品名"].astype(str) == str(item)].iloc[0]
            price = safe_int(item_row["単価"])
            amount = qty * price

            st.write(f"単価：{price:,}円")
            st.write(f"金額：{amount:,}円")

            ok = st.form_submit_button("登録する")

        if ok:
            new = pd.DataFrame([{
                "記録ID": now_id("R"),
                "日付": d,
                "利用者": user,
                "物品": item,
                "数量": qty,
                "単価": price,
                "金額": amount,
                "備考": note,
                "登録日時": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            }])

            usage = pd.concat([usage, new], ignore_index=True)
            save_df(usage, USAGE, USAGE_COLS)

            if item in stock["物品"].astype(str).values:
                idx = stock[stock["物品"].astype(str) == str(item)].index[0]
                stock.loc[idx, "現在庫"] = safe_int(stock.loc[idx, "現在庫"]) - qty
                stock.loc[idx, "更新日時"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                save_df(stock, STOCK, STOCK_COLS)

            st.success("使用記録を登録しました。在庫も減算しました。")
            st.rerun()


elif menu == "使用記録 検索・更新・削除":
    st.subheader("使用記録 検索・更新・削除")

    if usage.empty:
        st.info("使用記録がありません。")
    else:
        work = usage.copy()
        work["日付_dt"] = pd.to_datetime(work["日付"], errors="coerce")

        col1, col2, col3 = st.columns(3)
        keyword = col1.text_input("検索語（利用者・物品・備考）")
        year = col2.number_input("年", 2024, 2035, date.today().year)
        month = col3.number_input("月", 1, 12, date.today().month)

        filtered = work[work["日付_dt"].dt.strftime("%Y-%m") == f"{year}-{month:02d}"].copy()

        if keyword:
            filtered = filtered[
                filtered["利用者"].astype(str).str.contains(keyword, na=False) |
                filtered["物品"].astype(str).str.contains(keyword, na=False) |
                filtered["備考"].astype(str).str.contains(keyword, na=False)
            ]

        st.dataframe(filtered[USAGE_COLS], use_container_width=True)

        if not filtered.empty:
            selected_id = st.selectbox("更新・削除する記録ID", filtered["記録ID"].astype(str))
            row = usage[usage["記録ID"].astype(str) == selected_id].iloc[0]

            with st.form("usage_edit"):
                new_date = st.date_input("日付", pd.to_datetime(row["日付"]).date())

                user_list = list(users["利用者名"].dropna().astype(str))
                item_list = list(items["物品名"].dropna().astype(str))

                new_user = st.selectbox(
                    "利用者",
                    user_list,
                    index=user_list.index(str(row["利用者"])) if str(row["利用者"]) in user_list else 0
                )

                new_item = st.selectbox(
                    "物品",
                    item_list,
                    index=item_list.index(str(row["物品"])) if str(row["物品"]) in item_list else 0
                )

                new_qty = st.number_input("数量", min_value=1, max_value=9999, value=max(1, safe_int(row["数量"])))
                new_note = st.text_input("備考", str(row.get("備考", "")))

                c1, c2 = st.columns(2)
                update = c1.form_submit_button("更新する")
                delete = c2.form_submit_button("削除する")

            if update:
                price = safe_int(items[items["物品名"].astype(str) == str(new_item)].iloc[0]["単価"])
                amount = new_qty * price

                idx = usage[usage["記録ID"].astype(str) == selected_id].index[0]
                usage.loc[idx, ["日付", "利用者", "物品", "数量", "単価", "金額", "備考"]] = [
                    new_date, new_user, new_item, new_qty, price, amount, new_note
                ]

                save_df(usage, USAGE, USAGE_COLS)
                st.success("使用記録を更新しました。※在庫は必要に応じて現在庫画面で調整してください。")
                st.rerun()

            if delete:
                usage = usage[usage["記録ID"].astype(str) != selected_id]
                save_df(usage, USAGE, USAGE_COLS)
                st.success("使用記録を削除しました。※在庫は必要に応じて現在庫画面で調整してください。")
                st.rerun()


elif menu == "現在庫 登録・更新":
    st.subheader("現在庫 登録・更新")

    stock = sync_stock(items, stock)
    save_df(stock, STOCK, STOCK_COLS)

    stock_status = make_stock_status(stock, items)

    if not stock_status.empty:
        st.markdown("### 在庫状況")
        st.dataframe(stock_status[["状態", "物品", "現在庫", "最低在庫", "不足数"]], use_container_width=True)
    else:
        st.info("物品マスタを登録してください。")

    if not stock.empty:
        st.markdown("### 入庫・棚卸修正")
        with st.form("stock_update"):
            item = st.selectbox("物品", stock["物品"].dropna().astype(str))
            mode = st.radio("処理", ["入庫として加算", "実在庫数に修正"])
            qty = st.number_input("数量", min_value=0, max_value=99999, value=1)
            ok = st.form_submit_button("在庫を更新する")

        if ok:
            idx = stock[stock["物品"].astype(str) == str(item)].index[0]

            if mode == "入庫として加算":
                stock.loc[idx, "現在庫"] = safe_int(stock.loc[idx, "現在庫"]) + qty
            else:
                stock.loc[idx, "現在庫"] = qty

            stock.loc[idx, "更新日時"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            save_df(stock, STOCK, STOCK_COLS)

            st.success("在庫を更新しました。")
            st.rerun()


elif menu == "月間集計":
    st.subheader("月間集計")

    if usage.empty:
        st.info("使用記録がありません。")
    else:
        work = usage.copy()
        work["日付"] = pd.to_datetime(work["日付"], errors="coerce")

        ym = st.selectbox("対象月", sorted(work["日付"].dt.strftime("%Y-%m").dropna().unique(), reverse=True))

        target = work[work["日付"].dt.strftime("%Y-%m") == ym].copy()
        target["数量"] = target["数量"].apply(safe_int)
        target["金額"] = target["金額"].apply(safe_int)

        st.markdown("### 利用者別合計")
        st.dataframe(target.groupby("利用者")[["金額"]].sum().reset_index(), use_container_width=True)

        st.markdown("### 物品別合計")
        st.dataframe(target.groupby("物品")[["数量", "金額"]].sum().reset_index(), use_container_width=True)

        st.markdown("### 利用者別・物品別")
        st.dataframe(target.groupby(["利用者", "物品"])[["数量", "金額"]].sum().reset_index(), use_container_width=True)


elif menu == "請求書作成":
    st.subheader("請求書作成")

    if usage.empty:
        st.info("使用記録がありません。")
    else:
        work = usage.copy()
        work["日付"] = pd.to_datetime(work["日付"], errors="coerce")

        ym = st.selectbox("請求月", sorted(work["日付"].dt.strftime("%Y-%m").dropna().unique(), reverse=True))
        user = st.selectbox("利用者", sorted(work["利用者"].dropna().astype(str).unique()))

        target = work[
            (work["日付"].dt.strftime("%Y-%m") == ym) &
            (work["利用者"].astype(str) == str(user))
        ].copy()

        target["数量"] = target["数量"].apply(safe_int)
        target["金額"] = target["金額"].apply(safe_int)

        bill = target.groupby("物品")[["数量", "金額"]].sum().reset_index()
        total = int(bill["金額"].sum()) if not bill.empty else 0

        st.dataframe(bill, use_container_width=True)
        st.markdown(f"## 合計：{total:,}円")

        text = f"請求書\n\n対象月：{ym}\n利用者：{user}\n\n"

        for _, r in bill.iterrows():
            text += f"{r['物品']}　数量：{int(r['数量'])}　金額：{int(r['金額']):,}円\n"

        text += f"\n合計：{total:,}円"

        st.text_area("請求書本文", text, height=300)
        st.download_button("請求書をダウンロード", text, file_name=f"invoice_{user}_{ym}.txt", mime="text/plain")


elif menu == "FEED発注候補":
    st.subheader("FEED発注候補")

    stock_status = make_stock_status(stock, items)

    if stock_status.empty:
        st.info("物品マスタと在庫データを登録してください。")
    else:
        order = stock_status[stock_status["状態"].isin(["🔴 在庫0", "🔴 発注必要", "🟡 注意"])].copy()

        if order.empty:
            st.success("現在、発注候補はありません。")
        else:
            st.warning("発注確認が必要な物品があります。")
            show = order[["状態", "物品", "現在庫", "最低在庫", "不足数", "単価", "FEED商品URL"]]
            st.dataframe(show, use_container_width=True)

            st.markdown("### FEED商品ページ")
            for _, r in order.iterrows():
                url = str(r.get("FEED商品URL", "")).strip()
                if url.startswith("http"):
                    st.link_button(f"{r['物品']} の商品ページを開く", url)

            csv = show.to_csv(index=False).encode("utf-8-sig")
            st.download_button("発注候補CSVをダウンロード", csv, file_name="feed_order_candidates.csv", mime="text/csv")


elif menu == "利用者マスタ 登録・更新・削除":
    st.subheader("利用者マスタ 登録・更新・削除")

    st.markdown("### 登録")
    with st.form("user_create"):
        name = st.text_input("利用者名")
        billing = st.text_input("請求先")
        note = st.text_input("備考")
        ok = st.form_submit_button("登録する")

    if ok:
        if not name.strip():
            st.error("利用者名を入力してください。")
        else:
            new = pd.DataFrame([{
                "利用者ID": now_id("U"),
                "利用者名": name.strip(),
                "請求先": billing.strip(),
                "備考": note.strip()
            }])
            users = pd.concat([users, new], ignore_index=True)
            save_df(users, USERS, USER_COLS)
            st.success("利用者を登録しました。")
            st.rerun()

    st.markdown("### 検索・更新・削除")
    keyword = st.text_input("検索語（利用者名・請求先）", key="user_search")

    filtered = users.copy()
    if keyword:
        filtered = filtered[
            filtered["利用者名"].astype(str).str.contains(keyword, na=False) |
            filtered["請求先"].astype(str).str.contains(keyword, na=False)
        ]

    st.dataframe(filtered, use_container_width=True)

    if not filtered.empty:
        selected = st.selectbox("更新・削除する利用者ID", filtered["利用者ID"].astype(str))
        row = users[users["利用者ID"].astype(str) == selected].iloc[0]

        with st.form("user_edit"):
            new_name = st.text_input("利用者名", str(row["利用者名"]))
            new_billing = st.text_input("請求先", str(row["請求先"]))
            new_note = st.text_input("備考", str(row["備考"]))

            c1, c2 = st.columns(2)
            update = c1.form_submit_button("更新する")
            delete = c2.form_submit_button("削除する")

        if update:
            old_name = str(row["利用者名"])
            idx = users[users["利用者ID"].astype(str) == selected].index[0]
            users.loc[idx, ["利用者名", "請求先", "備考"]] = [new_name, new_billing, new_note]

            if old_name != new_name:
                usage.loc[usage["利用者"].astype(str) == old_name, "利用者"] = new_name
                save_df(usage, USAGE, USAGE_COLS)

            save_df(users, USERS, USER_COLS)
            st.success("利用者を更新しました。")
            st.rerun()

        if delete:
            if str(row["利用者名"]) in usage["利用者"].astype(str).values:
                st.error("使用記録に使われている利用者は削除できません。")
            else:
                users = users[users["利用者ID"].astype(str) != selected]
                save_df(users, USERS, USER_COLS)
                st.success("利用者を削除しました。")
                st.rerun()


elif menu == "物品マスタ 登録・更新・削除":
    st.subheader("物品マスタ 登録・更新・削除")

    st.markdown("### 登録")
    with st.form("item_create"):
        name = st.text_input("物品名")
        price = st.number_input("単価", min_value=0, max_value=999999, value=0)
        min_stock = st.number_input("最低在庫", min_value=0, max_value=99999, value=0)
        url = st.text_input("FEED商品URL")
        note = st.text_input("備考")
        ok = st.form_submit_button("登録する")

    if ok:
        if not name.strip():
            st.error("物品名を入力してください。")
        else:
            new = pd.DataFrame([{
                "物品ID": now_id("I"),
                "物品名": name.strip(),
                "単価": price,
                "最低在庫": min_stock,
                "FEED商品URL": url.strip(),
                "備考": note.strip()
            }])
            items = pd.concat([items, new], ignore_index=True)
            save_df(items, ITEMS, ITEM_COLS)

            stock = sync_stock(items, stock)
            save_df(stock, STOCK, STOCK_COLS)

            st.success("物品を登録しました。")
            st.rerun()

    st.markdown("### 検索・更新・削除")
    keyword = st.text_input("検索語（物品名・URL・備考）", key="item_search")

    filtered = items.copy()
    if keyword:
        filtered = filtered[
            filtered["物品名"].astype(str).str.contains(keyword, na=False) |
            filtered["FEED商品URL"].astype(str).str.contains(keyword, na=False) |
            filtered["備考"].astype(str).str.contains(keyword, na=False)
        ]

    st.dataframe(filtered, use_container_width=True)

    if not filtered.empty:
        selected = st.selectbox("更新・削除する物品ID", filtered["物品ID"].astype(str))
        row = items[items["物品ID"].astype(str) == selected].iloc[0]

        with st.form("item_edit"):
            new_name = st.text_input("物品名", str(row["物品名"]))
            new_price = st.number_input("単価", min_value=0, max_value=999999, value=safe_int(row["単価"]))
            new_min = st.number_input("最低在庫", min_value=0, max_value=99999, value=safe_int(row["最低在庫"]))
            new_url = st.text_input("FEED商品URL", str(row["FEED商品URL"]))
            new_note = st.text_input("備考", str(row["備考"]))

            c1, c2 = st.columns(2)
            update = c1.form_submit_button("更新する")
            delete = c2.form_submit_button("削除する")

        if update:
            old_name = str(row["物品名"])
            idx = items[items["物品ID"].astype(str) == selected].index[0]
            items.loc[idx, ["物品名", "単価", "最低在庫", "FEED商品URL", "備考"]] = [
                new_name, new_price, new_min, new_url, new_note
            ]

            if old_name != new_name:
                usage.loc[usage["物品"].astype(str) == old_name, "物品"] = new_name
                stock.loc[stock["物品"].astype(str) == old_name, "物品"] = new_name
                save_df(usage, USAGE, USAGE_COLS)

            save_df(items, ITEMS, ITEM_COLS)
            stock = sync_stock(items, stock)
            save_df(stock, STOCK, STOCK_COLS)

            st.success("物品を更新しました。")
            st.rerun()

        if delete:
            if str(row["物品名"]) in usage["物品"].astype(str).values:
                st.error("使用記録に使われている物品は削除できません。")
            else:
                items = items[items["物品ID"].astype(str) != selected]
                save_df(items, ITEMS, ITEM_COLS)
                stock = sync_stock(items, stock)
                save_df(stock, STOCK, STOCK_COLS)

                st.success("物品を削除しました。")
                st.rerun()


elif menu == "ログイン設定":
    st.subheader("ログイン設定")
    st.caption("管理者のみ利用できます。職員用ID・パスワードの変更や追加ができます。")

    if st.session_state.get("role") != "管理者":
        st.error("このメニューは管理者のみ使用できます。")
        st.stop()

    accounts = load_accounts()

    st.markdown("### 現在のログイン一覧")
    safe_view = accounts.copy()
    safe_view["パスワード"] = "********"
    st.dataframe(safe_view, use_container_width=True)

    st.markdown("---")
    st.markdown("### 新規アカウント追加")

    with st.form("account_create"):
        new_id = st.text_input("ログインID")
        new_pw = st.text_input("パスワード", type="password")
        new_role = st.selectbox("権限", ["職員", "管理者"])
        new_name = st.text_input("表示名")
        create_ok = st.form_submit_button("追加する")

    if create_ok:
        if not new_id.strip() or not new_pw.strip():
            st.error("ログインIDとパスワードを入力してください。")
        elif new_id.strip() in accounts["ログインID"].astype(str).values:
            st.error("同じログインIDがすでに存在します。")
        else:
            new_row = pd.DataFrame([{
                "ログインID": new_id.strip(),
                "パスワード": new_pw.strip(),
                "権限": new_role,
                "表示名": new_name.strip() if new_name.strip() else new_id.strip(),
            }])
            accounts = pd.concat([accounts, new_row], ignore_index=True)
            save_accounts(accounts)
            st.success("アカウントを追加しました。")
            st.rerun()

    st.markdown("---")
    st.markdown("### パスワード・権限の変更 / 削除")

    if accounts.empty:
        st.info("アカウントがありません。")
    else:
        selected_id = st.selectbox("変更するログインID", accounts["ログインID"].astype(str))
        row = accounts[accounts["ログインID"].astype(str) == selected_id].iloc[0]

        with st.form("account_edit"):
            st.text_input("ログインID", str(row["ログインID"]), disabled=True)
            edit_pw = st.text_input("新しいパスワード", value=str(row["パスワード"]), type="password")
            edit_role = st.selectbox(
                "権限",
                ["職員", "管理者"],
                index=0 if str(row["権限"]) == "職員" else 1
            )
            edit_name = st.text_input("表示名", str(row["表示名"]))

            c1, c2 = st.columns(2)
            update_ok = c1.form_submit_button("更新する")
            delete_ok = c2.form_submit_button("削除する")

        if update_ok:
            idx = accounts[accounts["ログインID"].astype(str) == selected_id].index[0]
            accounts.loc[idx, ["パスワード", "権限", "表示名"]] = [
                edit_pw.strip(),
                edit_role,
                edit_name.strip()
            ]
            save_accounts(accounts)
            st.success("ログイン情報を更新しました。")
            st.rerun()

        if delete_ok:
            admin_count = len(accounts[accounts["権限"].astype(str) == "管理者"])
            target_role = str(row["権限"])

            if str(row["ログインID"]) == st.session_state.get("login_id"):
                st.error("現在ログイン中の自分自身は削除できません。")
            elif target_role == "管理者" and admin_count <= 1:
                st.error("管理者が0人になるため削除できません。")
            else:
                accounts = accounts[accounts["ログインID"].astype(str) != selected_id]
                save_accounts(accounts)
                st.success("アカウントを削除しました。")
                st.rerun()

    st.warning("注意：このVer1.5.1ではパスワードはExcelに保存されます。施設内の簡易運用向けです。")


elif menu == "データ確認":
    st.subheader("データ確認")

    tab1, tab2, tab3, tab4, tab5 = st.tabs(["利用者", "物品", "使用記録", "在庫", "ログイン"])

    with tab1:
        st.dataframe(users, use_container_width=True)
        st.download_button("利用者CSV", users.to_csv(index=False).encode("utf-8-sig"), "users.csv", "text/csv")

    with tab2:
        st.dataframe(items, use_container_width=True)
        st.download_button("物品CSV", items.to_csv(index=False).encode("utf-8-sig"), "items.csv", "text/csv")

    with tab3:
        st.dataframe(usage, use_container_width=True)
        st.download_button("使用記録CSV", usage.to_csv(index=False).encode("utf-8-sig"), "usage.csv", "text/csv")

    with tab4:
        st.dataframe(stock, use_container_width=True)
        st.download_button("在庫CSV", stock.to_csv(index=False).encode("utf-8-sig"), "stock.csv", "text/csv")

    with tab5:
        accounts = load_accounts()
        masked = accounts.copy()
        masked["パスワード"] = "********"
        st.dataframe(masked, use_container_width=True)
