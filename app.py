import streamlit as st
import pandas as pd
from datetime import date
from pathlib import Path

st.set_page_config(page_title="物品管理アプリ Ver1.0", layout="wide")

DATA = Path("data")
DATA.mkdir(exist_ok=True)

USERS = DATA / "users.xlsx"
ITEMS = DATA / "items.xlsx"
USAGE = DATA / "usage.xlsx"
STOCK = DATA / "stock.xlsx"

def load_df(path, cols):
    if path.exists():
        return pd.read_excel(path)
    df = pd.DataFrame(columns=cols)
    df.to_excel(path, index=False)
    return df

users = load_df(USERS, ["利用者名","請求先"])
items = load_df(ITEMS, ["物品名","単価","最低在庫"])
usage = load_df(USAGE, ["日付","利用者","物品","数量","単価","金額"])
stock = load_df(STOCK, ["物品","現在庫"])

st.title("📦 物品管理アプリ Ver1.0")

menu = st.sidebar.radio("メニュー",[
    "使用記録入力",
    "現在庫",
    "月間集計",
    "請求書作成",
    "マスタ管理"
])

if menu == "使用記録入力":

    st.subheader("使用記録入力")

    if users.empty:
        st.warning("先に利用者マスタを登録してください")
    elif items.empty:
        st.warning("先に物品マスタを登録してください")
    else:
        with st.form("usage"):

            d = st.date_input("日付",date.today())
            user = st.selectbox("利用者",users["利用者名"])
            item = st.selectbox("物品",items["物品名"])
            qty = st.number_input("数量",1,999,1)

            row = items[items["物品名"]==item].iloc[0]
            price = int(row["単価"])
            amount = qty * price

            st.write(f"単価: {price}円")
            st.write(f"金額: {amount}円")

            ok = st.form_submit_button("登録")

        if ok:

            new = pd.DataFrame([{
                "日付":d,
                "利用者":user,
                "物品":item,
                "数量":qty,
                "単価":price,
                "金額":amount
            }])

            usage = pd.concat([usage,new],ignore_index=True)
            usage.to_excel(USAGE,index=False)

            if item in stock["物品"].values:
                idx = stock[stock["物品"]==item].index[0]
                stock.loc[idx,"現在庫"] -= qty

            stock.to_excel(STOCK,index=False)

            st.success("登録しました")
            st.rerun()

elif menu == "現在庫":

    st.subheader("現在庫")

    st.dataframe(stock,use_container_width=True)

    if not stock.empty:

        with st.form("stock_add"):

            item = st.selectbox("物品",stock["物品"])
            qty = st.number_input("追加数",1,999,1)

            ok = st.form_submit_button("入庫")

        if ok:
            idx = stock[stock["物品"]==item].index[0]
            stock.loc[idx,"現在庫"] += qty
            stock.to_excel(STOCK,index=False)
            st.success("在庫更新")
            st.rerun()

elif menu == "月間集計":

    st.subheader("月間集計")

    if usage.empty:
        st.info("データなし")
    else:

        usage["日付"] = pd.to_datetime(usage["日付"])

        ym = st.selectbox(
            "対象月",
            sorted(usage["日付"].dt.strftime("%Y-%m").unique(), reverse=True)
        )

        target = usage[usage["日付"].dt.strftime("%Y-%m")==ym]

        st.markdown("### 利用者別")
        u = target.groupby("利用者")[["金額"]].sum().reset_index()
        st.dataframe(u,use_container_width=True)

        st.markdown("### 物品別")
        i = target.groupby("物品")[["数量","金額"]].sum().reset_index()
        st.dataframe(i,use_container_width=True)

elif menu == "請求書作成":

    st.subheader("請求書作成")

    if usage.empty:
        st.info("データなし")
    else:

        usage["日付"] = pd.to_datetime(usage["日付"])

        ym = st.selectbox(
            "請求月",
            sorted(usage["日付"].dt.strftime("%Y-%m").unique(), reverse=True)
        )

        user = st.selectbox(
            "利用者",
            sorted(usage["利用者"].unique())
        )

        target = usage[
            (usage["日付"].dt.strftime("%Y-%m")==ym) &
            (usage["利用者"]==user)
        ]

        bill = target.groupby("物品")[["数量","金額"]].sum().reset_index()

        total = int(bill["金額"].sum())

        st.dataframe(bill,use_container_width=True)

        st.markdown(f"## 合計: {total:,}円")

        text = f"請求書\n\n対象月:{ym}\n利用者:{user}\n\n"

        for _,r in bill.iterrows():
            text += f"{r['物品']} 数量:{r['数量']} 金額:{int(r['金額']):,}円\n"

        text += f"\n合計:{total:,}円"

        st.text_area("請求書",text,height=300)

        st.download_button(
            "請求書ダウンロード",
            text,
            file_name=f"invoice_{user}_{ym}.txt"
        )

elif menu == "マスタ管理":

    tab1,tab2 = st.tabs(["利用者","物品"])

    with tab1:

        st.subheader("利用者マスタ")

        edit = st.data_editor(users,num_rows="dynamic")

        if st.button("利用者保存"):
            edit.to_excel(USERS,index=False)
            st.success("保存しました")

    with tab2:

        st.subheader("物品マスタ")

        edit = st.data_editor(items,num_rows="dynamic")

        if st.button("物品保存"):

            edit.to_excel(ITEMS,index=False)

            new_stock = pd.DataFrame({
                "物品":edit["物品名"],
                "現在庫":0
            })

            if not stock.empty:
                for i,row in new_stock.iterrows():
                    name = row["物品"]
                    if name in stock["物品"].values:
                        new_stock.loc[i,"現在庫"] = int(
                            stock[stock["物品"]==name]["現在庫"].iloc[0]
                        )

            new_stock.to_excel(STOCK,index=False)

            st.success("保存しました")
