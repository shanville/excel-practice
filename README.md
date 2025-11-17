# Excel操作の基本練習

## プロジェクト概要
PythonでExcelファイルを扱う基本操作を学ぶためのプロジェクトです。

## セットアップ

```bash
# プロジェクトの初期化
uv init excel-practice
cd excel-practice

# 必要なパッケージのインストール
uv add pandas openpyxl
```

## ファイル構成

- `create_sample.py` - サンプルExcelファイルを作成
- `basic_operations.py` - Excelの基本操作を実践
- `sample_data.xlsx` - サンプルデータ
- `updated_data.xlsx` - 操作後のデータ

## 基本操作

### 1. Excelファイルの読み込み
```python
import pandas as pd
df = pd.read_excel('sample_data.xlsx')
```

### 2. データの確認
```python
# データフレーム全体を表示
print(df)

# 基本情報
print(f"行数: {len(df)}")
print(f"列名: {list(df.columns)}")
```

### 3. データの抽出・フィルタリング
```python
# 特定の列を取得
商品名 = df['商品名']

# 条件でフィルタリング
高額商品 = df[df['価格'] >= 150]
```

### 4. データの追加
```python
new_item = pd.DataFrame({
    '商品名': ['メロン'],
    '価格': [500],
    '在庫数': [10],
    'カテゴリ': ['果物']
})
df_updated = pd.concat([df, new_item], ignore_index=True)
```

### 5. Excelファイルへの書き込み
```python
df_updated.to_excel('updated_data.xlsx', index=False)
```

### 6. 統計情報の取得
```python
平均価格 = df['価格'].mean()
最高価格 = df['価格'].max()
最低価格 = df['価格'].min()
合計 = df['在庫数'].sum()
```

## VS CodeでExcelを扱うベストプラクティス

### 1. 拡張機能の活用
- **Excel Viewer** - VS Code内でExcelファイルをプレビュー
- **Rainbow CSV** - CSV形式での確認時に便利

### 2. Excelファイルの確認方法
VS CodeでExcelファイルを開くと:
- バイナリビューになるが、Excel Viewer拡張機能で表形式で見られる
- 編集は避け、Pythonコードで操作する

### 3. ワークフロー
1. Pythonスクリプトでデータを操作
2. `uv run python スクリプト名.py` で実行
3. 生成されたExcelファイルをExcelアプリケーションで確認
4. 必要に応じてスクリプトを修正して再実行

### 4. エンコーディング対策
Windows環境では文字化け対策が必要:
```python
import sys
if sys.platform == 'win32':
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
```

### 5. gitでの管理
`.gitignore`に以下を追加推奨:
```
*.xlsx
*.xls
.venv/
__pycache__/
```

データファイルはgit管理せず、生成スクリプトのみ管理する。

## 実行方法

```bash
# サンプルExcelファイルを作成
uv run python create_sample.py

# 基本操作を実行
uv run python basic_operations.py
```

## 学んだこと

### pandasの基本
- `pd.read_excel()` でExcel読み込み
- `df.to_excel()` でExcel書き込み
- データフレームの操作(フィルタリング、追加、統計)

### エンコーディング
- Windows環境での日本語表示対策
- UTF-8エンコーディングの設定

### uv の使い方
- `uv init` でプロジェクト作成
- `uv add` でパッケージ追加
- `uv run` でスクリプト実行

## 次のステップ

- [ ] 複数シートの操作
- [ ] CSVとExcelの相互変換
- [ ] グラフの作成(plotly)
- [ ] 実際の業務データを使った演習
