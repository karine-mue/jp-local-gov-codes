# jp-local-gov-codes

総務省「全国地方公共団体コード」の配布ファイル（Excel）を、そのまま機械的に使える形に変換する前処理パイプライン（Python 3.10+、要 openpyxl）。

配布ファイルは2シートに分かれ、カナは半角、履歴（改正一覧表）は「〃」や「同左」が混ざる人間向けの表です。本ツールはそれを **CSV + JSON + JSON Schema** に整形し、変換時に全コードの検査数字を検証します。廃止・合併で消えたコードの履歴も含むので、過去のデータと現在のコード表を突き合わせる時系列用途にも使えます。

## 使い方

```sh
pip install openpyxl
python3 build_codes.py            # 総務省サイトから現行ファイルと改正一覧表を取得して変換
python3 build_codes.py --current source.xlsx --history history_source.xlsx --out-dir out
```

出力ファイル:

| ファイル | 内容 |
|---|---|
| `local_gov_codes.csv` / `.json` | 現行コード表（1,965件：都道府県・市区町村・政令市の区） |
| `history_events.csv` / `.json` | 改正イベント（2005-04-01 以降、改正一覧表の1行＝1レコード） |
| `code_lifespans.csv` / `.json` | コードごとの生存期間（valid_from / valid_to） |
| `schema.json` | `local_gov_codes` の JSON Schema |

### 現行コード表の項目

`code`（6桁）/ `code5`（検査数字なし5桁）/ `prefecture_code` / `prefecture_name` / `prefecture_kana`（全角）/ `prefecture_kana_halfwidth`（元の半角）/ `municipality_code` / `municipality_name` / `municipality_kana`（全角）/ `municipality_kana_halfwidth`（元の半角）/ `type`

`type` は `prefecture` / `designated_city`（政令指定都市）/ `ward`（政令市の区）/ `special_ward`（特別区）/ `city` / `town` / `village` のいずれかです。

### 履歴の項目

- `history_events`: `group_id`（同じ改正を束ねる連番）/ `date`（施行年月日）/ `prefecture_name` / `change_type`（`created` 新設・`retired` 欠番および「削除」・`continued` 存続）/ `code_before` / `name_before` / `code_after` / `name_after` / `reason`（改正事由等）
- `code_lifespans`: `code` / `name` / `prefecture_name` / `valid_from`（null は「対象期間の前から存在」）/ `valid_to`（null は「現在も有効」）

## テストと検証結果

```sh
python3 test_build_codes.py   # 単位テスト（検査数字・日付解釈・カナ変換・種別判定）
python3 test_real_data.py     # 実ファイルでの統合テスト（要 source.xlsx / history_source.xlsx）
```

検証の方法と結果は次のとおりです（2026-11-15 実施、2026-11-16 に再レビューと修正を反映、令和6年1月1日更新版に対して）。

- 検査数字の式（重み 6,5,4,3,2、合計が11未満の場合は 11−合計、それ以外は (11 − 余り) の下1桁）は、現行ファイル全1,985行で照合し**不一致ゼロ**でした。変換時にも全件検証し、1件でも合わなければ変換を中止します。
- 種別件数は実測どおり（都道府県47、政令指定都市20、区171、特別区23、市772、町743、村189、合計1,965）。
- 改正一覧表は417ブロック（改正のかたまり）に分割でき、全ブロックで施行年月日がちょうど1種類に定まることを確認しました。期間は 2005-04-01 〜 2024-01-01 です。
- 整合性チェック：欠番になったコードは現行コード表に存在せず、現行でないコードは全て生存期間に valid_to を持ちます。生存期間は2,943件（現行1,965＋廃止978）です。2026-11-16 の再レビューで、改定区分「削除」の行と合併の写像行で消滅した6コードが生存期間から漏れていた問題を修正しました。

## 注記

- 履歴の対象期間は改正一覧表の範囲（2005-04-01 以降）です。それ以前から存在するコードの `valid_from` は null で、「期間前から存在」を意味します。
- 改正一覧表のブロック分割は「都道府県名が現れる行から新しい改正が始まる」という原典の書式規則に基づきます。書式が変わってブロック内に日付が定まらない場合は、変換を中止して報告します（黙って誤った履歴は作りません）。
- カナは全角を本カラムとし、元の半角も別カラムに保持しています（どちらが正ではなく、用途で選ぶためのものです）。
- 一部事務組合等のコードは対象外です（総務省の別ファイル。本ツールは都道府県・市区町村本体のみ）。
- 元データの出典は総務省「全国地方公共団体コード」（https://www.soumu.go.jp/denshijiti/code.html ）です。

## ライセンス

MIT
