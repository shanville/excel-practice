"""ExcelファイルをアップロードしてPandasで統計分析するWebアプリ"""
import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(page_title="Excel統計分析アプリ", page_icon="📊", layout="wide")

st.title("📊 Excel統計分析アプリ")
st.markdown("Excelファイルをアップロードすると、自動で統計情報とグラフを表示します")

# ファイルアップロード
uploaded_file = st.file_uploader(
    "Excelファイルをアップロードしてください (.xlsx)",
    type=['xlsx']
)

if uploaded_file is not None:
    try:
        # Excelファイルを読み込み
        df = pd.read_excel(uploaded_file)

        st.success(f"✓ ファイルを読み込みました: {uploaded_file.name}")

        # 2カラムレイアウト
        col1, col2 = st.columns([2, 1])

        with col1:
            st.subheader("📋 データプレビュー")
            st.dataframe(df, use_container_width=True)

        with col2:
            st.subheader("📈 基本情報")
            st.metric("行数", f"{len(df):,}")
            st.metric("列数", f"{len(df.columns):,}")

        st.divider()

        # 数値列の統計情報
        numeric_cols = df.select_dtypes(include=['number']).columns.tolist()

        if numeric_cols:
            st.subheader("📊 数値列の統計情報")

            # タブで切り替え
            tab1, tab2, tab3 = st.tabs(["📈 統計サマリー", "📉 グラフ", "🔢 詳細統計"])

            with tab1:
                # 統計サマリー
                stats_df = df[numeric_cols].describe().T
                stats_df['合計'] = df[numeric_cols].sum()
                st.dataframe(stats_df, use_container_width=True)

            with tab2:
                # グラフ表示
                selected_col = st.selectbox("グラフ化する列を選択", numeric_cols)

                col_a, col_b = st.columns(2)

                with col_a:
                    # ヒストグラム
                    fig_hist = px.histogram(
                        df,
                        x=selected_col,
                        title=f"{selected_col} の分布",
                        labels={selected_col: selected_col, 'count': '件数'}
                    )
                    st.plotly_chart(fig_hist, use_container_width=True)

                with col_b:
                    # ボックスプロット
                    fig_box = px.box(
                        df,
                        y=selected_col,
                        title=f"{selected_col} のボックスプロット"
                    )
                    st.plotly_chart(fig_box, use_container_width=True)

            with tab3:
                # 各列の詳細統計
                for col in numeric_cols:
                    with st.expander(f"📊 {col} の詳細"):
                        metric_col1, metric_col2, metric_col3, metric_col4 = st.columns(4)

                        with metric_col1:
                            st.metric("平均", f"{df[col].mean():.2f}")
                        with metric_col2:
                            st.metric("中央値", f"{df[col].median():.2f}")
                        with metric_col3:
                            st.metric("最小値", f"{df[col].min():.2f}")
                        with metric_col4:
                            st.metric("最大値", f"{df[col].max():.2f}")

                        # ミニグラフ
                        mini_fig = px.line(
                            y=df[col].values,
                            title=f"{col} の推移",
                            labels={'x': 'インデックス', 'y': col}
                        )
                        st.plotly_chart(mini_fig, use_container_width=True)

        else:
            st.warning("数値列が見つかりませんでした")

        # 文字列列の情報
        text_cols = df.select_dtypes(include=['object']).columns.tolist()

        if text_cols:
            st.divider()
            st.subheader("📝 テキスト列の情報")

            for col in text_cols:
                with st.expander(f"📄 {col}"):
                    unique_count = df[col].nunique()
                    st.write(f"ユニーク数: {unique_count}")

                    # 上位5件の出現頻度
                    value_counts = df[col].value_counts().head(5)

                    col_chart, col_table = st.columns([2, 1])

                    with col_chart:
                        fig_bar = px.bar(
                            x=value_counts.index,
                            y=value_counts.values,
                            title=f"{col} の上位5件",
                            labels={'x': col, 'y': '件数'}
                        )
                        st.plotly_chart(fig_bar, use_container_width=True)

                    with col_table:
                        st.dataframe(
                            pd.DataFrame({
                                '値': value_counts.index,
                                '件数': value_counts.values
                            }),
                            use_container_width=True
                        )

        # ダウンロードボタン
        st.divider()
        st.subheader("💾 統計情報のダウンロード")

        # 統計情報をCSVに変換
        if numeric_cols:
            stats_csv = df[numeric_cols].describe().to_csv(encoding='utf-8-sig')
            st.download_button(
                label="📥 統計情報をCSVでダウンロード",
                data=stats_csv,
                file_name="statistics.csv",
                mime="text/csv"
            )

    except Exception as e:
        st.error(f"エラーが発生しました: {str(e)}")
        st.info("Excelファイルの形式を確認してください")

else:
    # サンプル用の説明
    st.info("👆 上のボタンからExcelファイルをアップロードしてください")

    with st.expander("💡 使い方"):
        st.markdown("""
        1. **Excelファイルをアップロード** - `.xlsx`形式のファイルに対応
        2. **自動で分析** - データのプレビュー、統計情報、グラフが表示されます
        3. **インタラクティブに探索** - グラフ化する列を選択したり、詳細を確認できます
        4. **ダウンロード** - 統計情報をCSVでダウンロード可能

        **試してみる**: 既にある `sample_data.xlsx` をアップロードしてみてください!
        """)
