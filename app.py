import streamlit as st
import pandas as pd
from datetime import date
from pathlib import Path

st.set_page_config(page_title="物品管理アプリ Ver1.1", layout="wide")

DATA = Path("data")
DATA.mkdir(exist_ok=True)

USERS = DATA / "users.xlsx"
ITEMS = DATA / "items.xlsx"
USAGE = DATA / "usage.xlsx"
STOCK = DATA / "stock.xlsx"

USER_COLS = ["利用者名", "請求先"]
ITEM_COLS = ["物品名", "単価", "最低在庫", "FEED商品URL"]
USAGE_COLS = ["日付", "利用者", "物品", "数量", "単価", "金額"]
STOCK_COLS = ["物品", "現在庫"]

def load_df(path, cols):
    if path.exists():
        df = pd.read_excel(path)
        for c in cols:
            if c not in df.columns:
                df[c] = ""
        df = df[cols]
        df.to_excel(path, index=False)
        return df
    df = pd.DataFrame(columns=cols)
    df.to_excel(path, index=False)
    return df

def safe_int(v):
    try:
        if pd.isna(v) or v == "":
            return 0
        return int(float(v))
    except Exception:
        return 0

users = load_df(USERS, USER_COLS)
items = load_df(ITEMS, ITEM_COLS)
usage = load_df(USAGE, USAGE_COLS)
stock = load_df(STOCK, STOCK_COLS)

st.title("📦 物品管理アプリ Ver1.1")
st.caption("使用記録・在庫管理・月末請求・FEED発注補助")

menu = st.sidebar.radio(
    "メニュー",
    [
        "使用記録入力",
        "現在庫",
        "月間集計",
        "請求書作成",
        "FEED発注候補",
        "マスタ管理",
    ],
)

if menu == "使用記録入力":

    st.subheader("使用記録入力")

    if users.empty:
        st.warning("先に利用者マスタを登録してください。")
    elif items.empty:
        st.warning("先に物品マスタを登録してください。")
    else:
        with st.form("usage_form"):
            d = st.date_input("日付", date.today())
            user = st.selectbox("利用者", users["利用者名"].dropna())
            item = st.selectbox("物品", items["物品名"].dropna())
            qty = st.number_input("数量", min_value=1, max_value=999, value=1)

            row = items[items["物品名"] == item].iloc[0]
            price = safe_int(row["単価"])
            amount = qty * price

            st.write(f"単価：{price:,}円")
            st.write(f"金額：{amount:,}円")

            ok = st.form_submit_button("登録")

        if ok:
            new = pd.DataFrame([{
                "日付": d,
                "利用者": user,
                "物品": item,
                "数量": qty,
                "単価": price,
                "金額": amount,
            }])

            usage = pd.concat([usage, new], ignore_index=True)
            usage.to_excel(USAGE, index=False)

            if item in stock["物品"].values:
                idx = stock[stock["物品"] == item].index[0]
                stock.loc[idx, "現在庫"] = safe_int(stock.loc[idx, "現在庫"]) - qty

            stock.to_excel(STOCK, index=False)

            st.success("登録しました。在庫も減算しました。")
            st.rerun()

elif menu == "現在庫":

    st.subheader("現在庫")

    st.dataframe(stock, use_container_width=True)

    if not stock.empty:
        with st.form("stock_form"):
            item = st.selectbox("物品", stock["物品"].dropna())
            qty = st.number_input("入庫数", min_value=1, max_value=9999, value=1)
            ok = st.form_submit_button("入庫する")

        if ok:
            idx = stock[stock["物品"] == item].index[0]
            stock.loc[idx, "現在庫"] = safe_int(stock.loc[idx, "現在庫"]) + qty
            stock.to_excel(STOCK, index=False)
            st.success("在庫を更新しました。")
            st.rerun()

elif menu == "月間集計":

    st.subheader("月間集計")

    if usage.empty:
        st.info("使用記録がありません。")
    else:
        usage["日付"] = pd.to_datetime(usage["日付"], errors="coerce")

        ym = st.selectbox(
            "対象月",
            sorted(usage["日付"].dt.strftime("%Y-%m").dropna().unique(), reverse=True)
        )

        target = usage[usage["日付"].dt.strftime("%Y-%m") == ym]

        st.markdown("### 利用者別合計")
        by_user = target.groupby("利用者")[["金額"]].sum().reset_index()
        st.dataframe(by_user, use_container_width=True)

        st.markdown("### 物品別合計")
        by_item = target.groupby("物品")[["数量", "金額"]].sum().reset_index()
        st.dataframe(by_item, use_container_width=True)

        st.markdown("### 利用者別・物品別")
        detail = target.groupby(["利用者", "物品"])[["数量", "金額"]].sum().reset_index()
        st.dataframe(detail, use_container_width=True)

elif menu == "請求書作成":

    st.subheader("請求書作成")

    if usage.empty:
        st.info("使用記録がありません。")
    else:
        usage["日付"] = pd.to_datetime(usage["日付"], errors="coerce")

        ym = st.selectbox(
            "請求月",
            sorted(usage["日付"].dt.strftime("%Y-%m").dropna().unique(), reverse=True)
        )

        user = st.selectbox("利用者", sorted(usage["利用者"].dropna().unique()))

        target = usage[
            (usage["日付"].dt.strftime("%Y-%m") == ym) &
            (usage["利用者"] == user)
        ]

        bill = target.groupby("物品")[["数量", "金額"]].sum().reset_index()
        total = int(bill["金額"].sum()) if not bill.empty else 0

        st.dataframe(bill, use_container_width=True)
        st.markdown(f"## 合計：{total:,}円")

        text = f"請求書\n\n対象月：{ym}\n利用者：{user}\n\n"

        for _, r in bill.iterrows():
            text += f"{r['物品']}　数量：{int(r['数量'])}　金額：{int(r['金額']):,}円\n"

        text += f"\n合計：{total:,}円"

        st.text_area("請求書本文", text, height=300)

        st.download_button(
            "請求書をダウンロード",
            text,
            file_name=f"invoice_{user}_{ym}.txt",
            mime="text/plain",
        )

elif menu == "FEED発注候補":

    st.subheader("FEED発注候補")
    st.caption("現在庫が最低在庫以下の物品を表示します。FEED商品URLを登録しておくと、ここから商品ページを開けます。")

    if stock.empty or items.empty:
        st.info("物品マスタと在庫データを登録してください。")
    else:
        merged = stock.merge(
            items,
            left_on="物品",
            right_on="物品名",
            how="left"
        )

        merged["現在庫"] = merged["現在庫"].apply(safe_int)
        merged["最低在庫"] = merged["最低在庫"].apply(safe_int)
        merged["不足数"] = merged["最低在庫"] - merged["現在庫"]

        order = merged[merged["現在庫"] <= merged["最低在庫"]].copy()

        if order.empty:
            st.success("現在、発注候補はありません。")
        else:
            st.warning("発注確認が必要な物品があります。")

            show_cols = ["物品", "現在庫", "最低在庫", "不足数", "単価", "FEED商品URL"]
            st.dataframe(
                order[show_cols],
                use_container_width=True,
                column_config={
                    "FEED商品URL": st.column_config.LinkColumn(
                        "FEED商品URL",
                        display_text="商品ページを開く"
                    )
                }
            )

            csv = order[show_cols].to_csv(index=False).encode("utf-8-sig")

            st.download_button(
                "発注候補CSVをダウンロード",
                csv,
                file_name="feed_order_candidates.csv",
                mime="text/csv",
            )

elif menu == "マスタ管理":

    st.subheader("マスタ管理")

    tab1, tab2 = st.tabs(["利用者マスタ", "物品マスタ"])

    with tab1:
        st.markdown("### 利用者マスタ")
        edit_users = st.data_editor(
            users,
            num_rows="dynamic",
            use_container_width=True
        )

        if st.button("利用者マスタを保存"):
            edit_users.to_excel(USERS, index=False)
            st.success("利用者マスタを保存しました。")
            st.rerun()

    with tab2:
        st.markdown("### 物品マスタ")
        st.caption("FEED商品URL欄に商品ページのURLを貼ると、発注候補画面から開けます。")

        edit_items = st.data_editor(
            items,
            num_rows="dynamic",
            use_container_width=True,
            column_config={
                "FEED商品URL": st.column_config.LinkColumn(
                    "FEED商品URL",
                    display_text="開く"
                )
            }
        )

        if st.button("物品マスタを保存"):
            edit_items.to_excel(ITEMS, index=False)

            new_stock = pd.DataFrame({
                "物品": edit_items["物品名"],
                "現在庫": 0
            })

            if not stock.empty:
                for i, row in new_stock.iterrows():
                    name = row["物品"]
                    if name in stock["物品"].values:
                        new_stock.loc[i, "現在庫"] = safe_int(
                            stock[stock["物品"] == name]["現在庫"].iloc[0]
                        )

            new_stock.to_excel(STOCK, index=False)

            st.success("物品マスタを保存しました。在庫マスタも更新しました。")
            st.rerun()
