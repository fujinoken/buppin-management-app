import streamlit as st
import pandas as pd
from datetime import date, datetime
from pathlib import Path

st.set_page_config(page_title="物品管理アプリ Ver1.2", layout="wide")

DATA = Path("data")
DATA.mkdir(exist_ok=True)

USERS = DATA / "users.xlsx"
ITEMS = DATA / "items.xlsx"
USAGE = DATA / "usage.xlsx"
STOCK = DATA / "stock.xlsx"

USER_COLS = ["利用者ID", "利用者名", "請求先", "備考"]
ITEM_COLS = ["物品ID", "物品名", "単価", "最低在庫", "FEED商品URL", "備考"]
USAGE_COLS = ["記録ID", "日付", "利用者", "物品", "数量", "単価", "金額", "備考", "登録日時"]
STOCK_COLS = ["物品", "現在庫", "更新日時"]

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

def sync_stock(items, stock):
    current_names = set(stock["物品"].astype(str)) if not stock.empty else set()
    rows = []
    for _, r in items.iterrows():
        name = str(r["物品名"]).strip()
        if not name:
            continue
        if name in current_names:
            old = stock[stock["物品"].astype(str) == name].iloc[0]
            rows.append({
                "物品": name,
                "現在庫": safe_int(old["現在庫"]),
                "更新日時": old.get("更新日時", "")
            })
        else:
            rows.append({
                "物品": name,
                "現在庫": 0,
                "更新日時": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            })
    return pd.DataFrame(rows, columns=STOCK_COLS)

users = load_df(USERS, USER_COLS)
items = load_df(ITEMS, ITEM_COLS)
usage = load_df(USAGE, USAGE_COLS)
stock = load_df(STOCK, STOCK_COLS)

# 旧Verからの移行補助
if "利用者ID" in users.columns:
    mask = users["利用者ID"].astype(str).str.strip() == ""
    users.loc[mask, "利用者ID"] = [now_id("U") for _ in range(mask.sum())]
if "物品ID" in items.columns:
    mask = items["物品ID"].astype(str).str.strip() == ""
    items.loc[mask, "物品ID"] = [now_id("I") for _ in range(mask.sum())]
if "記録ID" in usage.columns:
    mask = usage["記録ID"].astype(str).str.strip() == ""
    usage.loc[mask, "記録ID"] = [now_id("R") for _ in range(mask.sum())]

save_df(users, USERS, USER_COLS)
save_df(items, ITEMS, ITEM_COLS)
save_df(usage, USAGE, USAGE_COLS)

st.title("📦 物品管理アプリ Ver1.2")
st.caption("登録・検索・更新・削除／月末請求／FEED発注補助")

menu = st.sidebar.radio(
    "メニュー",
    [
        "ホーム",
        "使用記録 登録",
        "使用記録 検索・更新・削除",
        "現在庫 登録・更新",
        "月間集計",
        "請求書作成",
        "FEED発注候補",
        "利用者マスタ 登録・更新・削除",
        "物品マスタ 登録・更新・削除",
        "データ確認",
    ],
)

if menu == "ホーム":
    st.subheader("ホーム")
    col1, col2, col3 = st.columns(3)

    with col1:
        st.metric("利用者数", len(users))
    with col2:
        st.metric("物品数", len(items))
    with col3:
        st.metric("使用記録数", len(usage))

    st.markdown("### 発注注意")
    if stock.empty or items.empty:
        st.info("物品マスタと在庫を登録してください。")
    else:
        merged = stock.merge(items, left_on="物品", right_on="物品名", how="left")
        merged["現在庫"] = merged["現在庫"].apply(safe_int)
        merged["最低在庫"] = merged["最低在庫"].apply(safe_int)
        alert = merged[merged["現在庫"] <= merged["最低在庫"]]
        if alert.empty:
            st.success("最低在庫以下の物品はありません。")
        else:
            st.warning("発注確認が必要な物品があります。")
            st.dataframe(alert[["物品", "現在庫", "最低在庫", "FEED商品URL"]], use_container_width=True)

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
        with col1:
            keyword = st.text_input("検索語（利用者・物品・備考）")
        with col2:
            year = st.number_input("年", 2024, 2035, date.today().year)
        with col3:
            month = st.number_input("月", 1, 12, date.today().month)

        filtered = work[work["日付_dt"].dt.strftime("%Y-%m") == f"{year}-{month:02d}"].copy()
        if keyword:
            k = keyword
            filtered = filtered[
                filtered["利用者"].astype(str).str.contains(k, na=False) |
                filtered["物品"].astype(str).str.contains(k, na=False) |
                filtered["備考"].astype(str).str.contains(k, na=False)
            ]

        st.dataframe(filtered[USAGE_COLS], use_container_width=True)

        if not filtered.empty:
            selected_id = st.selectbox("更新・削除する記録ID", filtered["記録ID"].astype(str))
            row = usage[usage["記録ID"].astype(str) == selected_id].iloc[0]

            with st.form("usage_edit"):
                new_date = st.date_input("日付", pd.to_datetime(row["日付"]).date())
                new_user = st.selectbox(
                    "利用者",
                    users["利用者名"].dropna().astype(str),
                    index=list(users["利用者名"].dropna().astype(str)).index(str(row["利用者"])) if str(row["利用者"]) in list(users["利用者名"].dropna().astype(str)) else 0
                )
                new_item = st.selectbox(
                    "物品",
                    items["物品名"].dropna().astype(str),
                    index=list(items["物品名"].dropna().astype(str)).index(str(row["物品"])) if str(row["物品"]) in list(items["物品名"].dropna().astype(str)) else 0
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

    st.dataframe(stock, use_container_width=True)

    if not stock.empty:
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

        target = work[(work["日付"].dt.strftime("%Y-%m") == ym) & (work["利用者"].astype(str) == str(user))].copy()
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

    if stock.empty or items.empty:
        st.info("物品マスタと在庫データを登録してください。")
    else:
        merged = stock.merge(items, left_on="物品", right_on="物品名", how="left")
        merged["現在庫"] = merged["現在庫"].apply(safe_int)
        merged["最低在庫"] = merged["最低在庫"].apply(safe_int)
        merged["不足数"] = merged["最低在庫"] - merged["現在庫"]
        order = merged[merged["現在庫"] <= merged["最低在庫"]].copy()

        if order.empty:
            st.success("現在、発注候補はありません。")
        else:
            st.warning("発注確認が必要な物品があります。")
            show = order[["物品", "現在庫", "最低在庫", "不足数", "単価", "FEED商品URL"]]
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
            idx = users[users["利用者ID"].astype(str) == selected].index[0]
            users.loc[idx, ["利用者名", "請求先", "備考"]] = [new_name, new_billing, new_note]
            save_df(users, USERS, USER_COLS)
            st.success("利用者を更新しました。")
            st.rerun()

        if delete:
            if str(row["利用者名"]) in usage["利用者"].astype(str).values:
                st.error("使用記録に使われている利用者は削除できません。先に使用記録を確認してください。")
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
                st.error("使用記録に使われている物品は削除できません。先に使用記録を確認してください。")
            else:
                items = items[items["物品ID"].astype(str) != selected]
                save_df(items, ITEMS, ITEM_COLS)
                stock = sync_stock(items, stock)
                save_df(stock, STOCK, STOCK_COLS)
                st.success("物品を削除しました。")
                st.rerun()

elif menu == "データ確認":
    st.subheader("データ確認")

    tab1, tab2, tab3, tab4 = st.tabs(["利用者", "物品", "使用記録", "在庫"])

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
