# cobol-logic-hierarchy

COBOL ソースの `IF / EVALUATE / PERFORM / SEARCH` の入れ子構造を Excel 上で可視化する、単機能・軽量な VBA ツールです。
**外部依存ゼロ**（Excel + VBA のみ、PowerShell / .NET 不要）。

## 何ができる

| シート | 内容 |
|---|---|
| **コントロール** | 解析開始ボタン |
| **COBOLソース** | 全行 + 物理行番号、ロジック階層からのハイパーリンク先 |
| **ロジック階層** | `├── / └── / │` の罫線ツリー。`IF` 下に `[THEN]` / `[ELSE]`、`EVALUATE` / `SEARCH` 下に `WHEN` / `[AT END]` を分岐ノード化。`END-xx` / `ELSE` / `WHEN` の対応ミスを `★ERROR` で赤表示 |

特徴:

- 行末ピリオドの有無を吸収（`END-IF.` も `END-IF` も統一処理）
- 多行 `IF` 条件は IF ノードのテキストに自動連結
- 文字コード自動判定（UTF-8 / UTF-16 BOM + UTF-8 妥当性検証 + Shift_JIS フォールバック）
- COBOL 固定形式（1〜6 桁=行番号、7 桁=標識、8〜72 桁=コード）前提
- `PROCEDURE DIVISION` 以降だけ解析、`DATA DIVISION` 等は除外

## インストール

### 方法 A: インポート（おすすめ）

1. 任意の `.xlsm`（マクロ有効ブック）を新規作成
2. `Alt + F11` で VBE を開く
3. メニュー `ファイル → ファイルのインポート` で `vba/LogicHierarchy.bas` を選択
4. `デバッグ → VBAProject のコンパイル` でエラーが無いことを確認
5. `SetupControlSheet` にカーソルを置き `F5` を一度実行 → 「コントロール」シートとボタンが生成される

### 方法 B: 手書きコピー（環境が物理隔離されている場合）

`vba/LogicHierarchy.bas` の中身を VBE の標準モジュールに手で入力してください。
1 行目の `Attribute VB_Name = "LogicHierarchy"` は VBE が自動管理するので入力不要、`Option Explicit` から書き始めます。

## 使い方

1. 「コントロール」シートのボタン **「COBOL ソースを選択して解析」** をクリック
2. ファイル選択ダイアログで `.cbl` / `.cob` / `.txt` / `.src` / `.cpy` を選択
3. **COBOLソース** と **ロジック階層** シートが自動生成
4. **ロジック階層** の「開始行」列の数字をクリック → **COBOLソース** の該当行へジャンプ

サンプルは `samples/input/ICASE2.cbl`。

## 制限事項

- COBOL 固定形式が前提。最左 1〜7 桁がコードではなく行番号/標識である想定です。コードが 1 桁目から始まる自由形式は未対応。
- 動詞ホワイトリストに無い文（特殊な拡張文）は `ロジック階層` には表示されません（`COBOLソース` には全行表示）。
- `PERFORM 段落名` は out-of-line として葉ノード扱い、入れ子は展開しません。
- `GO TO`/`PERFORM THRU` の制御フロー追跡、`COPY` 句展開、テストケース生成、分岐カバレッジは対象外。

## ライセンス

MIT License. 詳細は `LICENSE` を参照。

## 動作確認

- Windows 11 + Excel 2021 (Microsoft 365) で動作確認
- 必要な参照: Office Object Library（既定で有効）、Microsoft ActiveX Data Objects（COM 経由で `ADODB.Stream` を使用、参照設定追加不要）
