# ExGM_EB会社ｺｰﾄﾞ変換TB 仕様書

> **ファイル種別**: .xlsm（マクロ付き）
> **用途**: ガンマ（照射管理システム）の顧客コードとEB（外部システム）の顧客コードの対応関係（変換テーブル）を登録・管理するツール。
> **VBA プロジェクト**: モジュール 10 本（.bas 7 / .cls 2 / .frm 0）
> **外部連携**: DSN=ricdb（Oracle）、DB接続先はODBCデータソース「ricdb」で解決（UID=ric）
> **解析日**: 2026-06-18（excel-to-md スキルによる自動解析）

---

## 凡例（本仕様書の表記ルール）

本仕様書では、保守時の判別を容易にするため、以下の表記ルールを使用します。

| 種別 | 表記 | 例 |
|---|---|---|
| モジュール（.bas / .cls） | **太字** | **登録.bas** |
| プロシージャ / イベント | `コード体()` | `登録更新()` |
| シート名 | 「」 | 「登録」 |
| セル参照 | `コード体` | `$B$6` |
| 名前付き範囲 | `コード体` | `AllData` |
| DB テーブル / カラム | `コード体` | `kcdcnvmst` / `kaisyacd` |
| ユーザー操作 | （操作名） | （登録 Click） |
| 主要マーク | ✓ | ✓ = 保守時に最初に確認すべき項目 |

### データフロー 場所マーク（9章）

9章のデータフロー（テーブル・ツリー図）では、処理が行われる場所を以下のアイコンで区別します。

| アイコン | 種別 | 意味 |
|---|---|---|
| 📊 | シート操作 | ワークシート上のセル書込み・読取り・表示変更 |
| 🖥️ | 画面操作 | ユーザーフォーム（.frm）の表示・入力・操作 |
| 🗄️ | DB操作 | DB への SELECT / INSERT / UPDATE / DELETE |
| 📄 | VBA内部処理 | 変数計算・条件分岐など、画面・シートに直接関与しない処理 |

### ✓（主要マーク）の判定基準

✓ は **保守時に最初に確認すべき項目** を示します。
判定基準は対象の種類ごとに以下のとおりです。

| 章 | 対象 | ✓ の判定基準 |
|---|---|---|
| 1.1 | シート | ユーザーが直接操作する、または VBA が動的に表示/非表示を切り替える |
| 1.3 / 6.0 | VBA モジュール | ① ユーザー操作の起点 ② DB I/O を含む ③ 他モジュールから呼び出される ④ コード行数上位 25%　のいずれか |
| 2 | セル / 名前付き範囲 | VBA から `Range()` で代入または参照され、業務ロジックに直結する |
| 3 | 名前付き範囲 | VBA から `Range()` で代入または参照され、業務ロジックに直結する |
| 5 | ボタン / コントロール | DB 更新・画面遷移・計算実行など副作用のある操作を起動する |
| 6.0（全プロシージャ） | プロシージャ | ① ユーザー操作の起点（Click イベント等） ② DB I/O を実行 ③ 他モジュールから呼び出される Public　のいずれか |
| 8.2 | DB テーブル | INSERT / UPDATE / DELETE の対象（参照のみのテーブルは ✓ なし） |

---

## 目次

1. [ファイル構成](#1-ファイル構成)
2. [シート詳細](#2-シート詳細)
3. [名前付き範囲一覧](#3-名前付き範囲一覧)
4. [数式一覧](#4-数式一覧)
5. [ボタン・マクロ対応](#5-ボタンマクロ対応)
6. [VBA モジュール仕様](#6-vba-モジュール仕様)
7. [ユーザーフォーム仕様](#7-ユーザーフォーム仕様)
8. [DB 接続・外部連携](#8-db-接続外部連携)
9. [データフロー](#9-データフロー)
10. [セキュリティ注意事項](#10-セキュリティ注意事項)

---

## 1. ファイル構成

### 1.1 シート一覧

| ✓ | シート名 | 最大行 | 最大列 | 保存時 Visible | VBA による動的切替 |
|---|---|---|---|---|---|
| ✓ | 「登録」 | 193 | 6 (F) | visible | — |

### 1.2 ユーザーフォーム一覧

ユーザーフォームは存在しない。

### 1.3 VBA モジュール一覧

| ✓ | モジュール | 種別 | プロシージャ数 | 主な役割 |
|---|---|---|---|---|
| ✓ | **ThisWorkbook** | .cls | 2 | ブック開閉時の初期化・終了処理 |
| ✓ | **Sheet1** | .cls | 1 | セル変更イベントでガンマ/EB顧客コード入力をハンドリング |
| ✓ | **終了処理** | .bas | 1 | ブックを閉じる処理 |
| ✓ | **SQL_Execution** | .bas | 6 | ADODB経由のDB接続・SQL実行の共通基盤 |
| ✓ | **登録** | .bas | 1 | 顧客コード変換テーブルへの登録・更新処理 |
| ✓ | **リンク解消** | .bas | 1 | 顧客コード変換テーブルからのペア削除処理 |
| ✓ | **クリア** | .bas | 2 | 画面クリアおよびイベント復旧ユーティリティ |
| ✓ | **該当表示** | .bas | 2 | ガンマ顧客コードからの会社名表示・既存紐付チェック |
| ✓ | **共通変数** | .bas | 0 | モジュール間共有のPublic変数宣言 |
| ✓ | **全データ表示** | .bas | 1 | 登録済みデータ一覧をDBから取得しシートに表示 |

> **✓判定理由**: **SQL_Execution**・**登録**・**該当表示** は行数上位25%かつDB I/O含む。**ThisWorkbook**・**Sheet1** はAutoExecイベント起点。**終了処理**・**リンク解消**・**クリア** はユーザー操作起点またはDB I/O含む。**全データ表示** は複数モジュールから呼出かつDB I/O含む。**共通変数** は全モジュールが参照するPublic変数を宣言。

---

## 2. シート詳細

### 2.0 シート可視性一覧

| シート名 | 可視性 | VBAによる動的切替 |
|----------|--------|-------------------|
| 「登録」 | 表示 | なし |

### 2.1 「登録」シート

**レイアウト構造**:

| 行 | 列 | セル | 内容 | 備考 |
|----|----|------|------|------|
| 1 | A | `A1` | "Ex4" | システム識別子 |
| 2 | B | `B2` | "顧客ｺｰﾄﾞ変換テーブル登録" | 画面タイトル |
| 5 | B〜F | `B5:F5` | ヘッダー行（入力エリア） | "ガンマ顧客ｺｰﾄﾞ" / "EB顧客ｺｰﾄﾞ" / "会社名" / "住所１" / "住所2" |
| 6 | B〜F | `B6:F6` | ✓ 入力・表示エリア | VBAが動的に読み書きする行（名前付き範囲 `Touroku`） |
| 7 | B | `B7` | "登録済一覧" | 一覧表示ヘッダーラベル |
| 9 | B〜F | `B9:F9` | ヘッダー行（一覧エリア） | "ガンマ顧客ｺｰﾄﾞ" / "EB顧客ｺｰﾄﾞ" / "会社名" / "住所１" / "住所2" |
| 10〜193 | B〜F | `B10:F193` | ✓ 一覧データ表示エリア | DBから取得した登録済みデータ（名前付き範囲 `AllData`）、AutoFilter適用 |

**ボタン配置**（VML Drawing）:

| ボタン名 | マクロ | 配置 |
|----------|--------|------|
| 登録 | `登録更新()` | シート上 |
| リンク解消 | `PairDelete()` | シート上 |
| 終了 | `Bookを閉じる()` | シート上 |

**シート保護**: `Workbook_Open` で `Protect UserInterfaceOnly:=True` を設定（VBAからの操作は許可、UI操作は制限）。描画オブジェクト保護・フィルタリング許可・ソート許可。

---

## 3. 名前付き範囲一覧

| ✓ | 名前 | 参照先 | 業務的意味 |
|---|------|--------|-----------|
| ✓ | `AllData` | `登録!$B$10:$F$193` | DB取得した登録済ペア一覧の表示領域。`画面クリア()` で全消去、`AllHyouji()` で再描画される |
| ✓ | `EBTNo` | `登録!$C$6` | EB顧客コード入力セル。3桁ゼロ埋めで使用される |
| ✓ | `GMTNo` | `登録!$B$6` | ガンマ顧客コード入力セル。4桁ゼロ埋めで使用される |
| ✓ | `Jyuusyo1` | `登録!$E$6` | 得意先マスタ (`tokumst`) から取得した住所1の表示先 |
| ✓ | `Jyuusyo2` | `登録!$F$6` | 得意先マスタ (`tokumst`) から取得した住所2の表示先 |
| ✓ | `KaiName` | `登録!$D$6` | 得意先マスタ (`tokumst`) から取得した会社名の表示先 |
| ✓ | `Touroku` | `登録!$B$6:$F$6` | 入力行全体（B6:F6）。`画面クリア()` や `PairDelete()` で一括クリアに使用 |

---

## 4. 数式一覧

数式は存在しない。全てのセル値はVBAによる動的な書き込みで管理されている。

---

## 5. ボタン・マクロ対応

### 5.1 シート上ボタン

| ✓ | ボタンテキスト | 呼出マクロ | 動作概要 |
|---|---------------|-----------|----------|
| ✓ | 登録 | `登録更新()` | ガンマ⇔EB顧客コードのペアをDBに登録または更新する |
| ✓ | リンク解消 | `PairDelete()` | 入力中のガンマ⇔EBペアのリンクをDBから削除する |
|   | 終了 | `Bookを閉じる()` | ブックを上書き保存せずに閉じる |

### 5.2 フォーム上ボタン

ユーザーフォームは存在しない。

---

## 6. VBA モジュール仕様

### 6.0 全プロシージャ一覧

| ✓ | モジュール | プロシージャ | 種別 | 概要 |
|---|---|---|---|---|
| ✓ | **ThisWorkbook** | `Workbook_Open()` | Event | ブック開 → シート保護設定→`画面クリア()`→`AllHyouji()`→AutoFilter設定→B6選択 |
|   | **ThisWorkbook** | `Workbook_BeforeClose()` | Event | ブック閉 → 警告非表示で保存済みフラグを立てる |
| ✓ | **Sheet1** | `Worksheet_Change()` | Event | セル変更 → B6変更→ガンマコード4桁埋め＋`会社名表示()`、C6変更→EB3桁埋め＋`既存紐付チェック()` |
| ✓ | **終了処理** | `Bookを閉じる()` | Sub | ボタン「終了」→ ブックを上書き保存せず閉じる。最後のブックならアプリ終了 |
| ✓ | **SQL_Execution** | `Open_oraconDB()` | Sub | 各DB操作Sub → DSN=ricdb でADODB接続を開く（R/W） |
| ✓ | **SQL_Execution** | `SQL_Exe()` | Sub | 各DB操作Sub → SQL文を `oraconn.Execute` で実行しRecordsetを返す（R/W） |
| ✓ | **SQL_Execution** | `SQL_INSERT_UPDATE()` | Sub | `登録更新()` → テーブル・キー・データ配列を受け取り、既存チェック後INSERT/UPDATE実行（W） |
| ✓ | **SQL_Execution** | `SQL_Delete()` | Sub | `登録更新()`, `PairDelete()` → テーブル名とWHERE句を受け取りDELETE実行（W） |
| ✓ | **SQL_Execution** | `Disp_Sheet()` | Sub | `AllHyouji()` → SQL結果をRecordsetで取得し、指定シート・行・列に `CopyFromRecordset` で展開（R） |
| ✓ | **SQL_Execution** | `Set_Array()` | Sub | `登録更新()`, `PairDelete()`, `会社名表示()`, `既存紐付チェック()` → SQL結果を2次元配列に格納して返す（R） |
| ✓ | **登録** | `登録更新()` | Sub | ボタン「登録」→ ガンマ⇔EBペアの重複チェック→INSERT/UPDATE/既存リンク削除→一覧再表示（R/W） |
| ✓ | **リンク解消** | `PairDelete()` | Sub | ボタン「リンク解消」→ ガンマ⇔EBペアの存在確認→確認ダイアログ→DELETE→一覧再表示（R/W） |
| ✓ | **クリア** | `画面クリア()` | Sub | `Workbook_Open()` → `Touroku`（B6:F6）と `AllData`（B10:F193）を空欄にする |
|   | **クリア** | `EventOn()` | Sub | 手動復旧用 → `Application.EnableEvents = True` でイベントを再有効化 |
| ✓ | **該当表示** | `会社名表示()` | Sub | `Worksheet_Change()`（B6変更時）→ ガンマ顧客コードで `tokumst` + `kcdcnvmst` を外部結合し、会社名・住所・既存EBコードを表示（R） |
| ✓ | **該当表示** | `既存紐付チェック()` | Sub | `Worksheet_Change()`（C6変更時）→ EBコードで `kcdcnvmst` + `tokumst` を外部結合し、既存リンクがあれば警告表示（R） |
|   | **共通変数** | （宣言のみ） | Public | 全モジュール参照 → `mpGMKaicd`（ガンマ会社コード）、`mpEBKaicd`（EB会社コード）、`mpKaiName`（会社名） |
| ✓ | **全データ表示** | `AllHyouji()` | Sub | `Workbook_Open()`, `登録更新()`, `PairDelete()` → `tokumst` JOIN `kcdcnvmst` の全件を取得し「登録」シートB10以降に展開（R） |

### 6.1 **ThisWorkbook**（ThisWorkbook.cls）

#### `Workbook_Open()`
1. `ActiveSheet.Unprotect` でシート保護を一旦解除
2. `ActiveSheet.Protect UserInterfaceOnly:=True` でシート保護を再設定（VBA操作は許可）
3. `画面クリア()` で入力欄・一覧エリアをクリア
4. `AllHyouji()` でDBから全登録済みペアを取得・表示
5. `Range("B9").AutoFilter` でAutoFilterを設定
6. `Range("B6").Select` でガンマ顧客コード入力セルにカーソル移動

#### `Workbook_BeforeClose()`
- `Application.DisplayAlerts = False` + `ActiveWorkbook.Saved = True` で保存確認ダイアログをスキップ

### 6.2 **Sheet1**（Sheet1.cls）

#### `Worksheet_Change()`
- **B6（ガンマ顧客コード）変更時**:
  - 空欄なら `Touroku`（B6:F6）全体をクリア
  - 値ありなら4桁ゼロ埋め後、`会社名表示()` を呼出
- **C6（EB顧客コード）変更時**（値が空でない場合）:
  - 3桁ゼロ埋め
  - ガンマコードが未入力ならエラーメッセージ表示＋EBコードクリア
  - ガンマコード入力済みなら `既存紐付チェック()` を呼出
- イベント処理中は `Application.EnableEvents = False` で再帰呼出を防止

### 6.3 **終了処理**（終了処理.bas）

#### `Bookを閉じる()`
- `Application.DisplayAlerts = False` で警告非表示
- 開いているブックが1つだけなら `Application.Quit`（Excel自体を終了）
- 複数ブックが開いていれば `ActiveWorkbook.Close`（当該ブックのみ閉じる）

### 6.4 **SQL_Execution**（SQL_Execution.bas）

Public変数:
- `mpErrDes As String` — 直前のSQL実行エラーメッセージ

Private変数:
- `oraconn As New ADODB.Connection` — DB接続オブジェクト
- `rs As ADODB.Recordset` — レコードセット

#### `Open_oraconDB()`
- 接続文字列: `DSN=ricdb;UID=ric;PWD=t6101`
- `oraconn.CursorLocation = adUseClient`（クライアントサイドカーソル）

#### `SQL_Exe(mySQL As String)`
- `oraconn.Execute(mySQL)` で実行し、結果をモジュールレベル変数 `rs` に格納
- エラー時は `mpErrDes` にエラーメッセージを記録

#### `SQL_INSERT_UPDATE(myTBL, myKey, myD(), myN)`
- `Open_oraconDB` → `BeginTrans`
- キー条件で `SELECT COUNT(*)` → 0件なら `INSERT`、1件以上なら `UPDATE`
- `CommitTrans` でコミット

#### `SQL_Delete(myTBL, myWhere)`
- `Open_oraconDB` → `BeginTrans`
- `DELETE myTBL myWhere` を実行
- `CommitTrans` でコミット

#### `Disp_Sheet(mySQL, mySH, myRow, myRecordCount, myColumn, myFieldCount, myF)`
- `Open_oraconDB` → SQL実行
- `myF = 1` の場合、フィールド名をヘッダー行に出力
- `CopyFromRecordset` でレコードセットをシートに一括転記
- `myRecordCount`, `myFieldCount` をByRefで返す

#### `Set_Array(mySQL, myData(), myRecordCount, myFldCount)`
- `Open_oraconDB` → SQL実行
- 結果を `myData(レコード番号, フィールド番号)` の2次元配列に格納
- `myRecordCount`, `myFldCount` をByRefで返す

### 6.5 **登録**（登録.bas）

#### `登録更新()`
1. ガンマ顧客コード（`GMTNo`）未入力チェック → 4桁ゼロ埋め
2. EB顧客コード（`EBTNo`）未入力チェック → 3桁ゼロ埋め
3. **重複チェック①**: ガンマ＋EBの組み合わせが `kcdcnvmst` に既存 → "既にリンク済" で終了
4. **重複チェック②**: ガンマコードのみで `kcdcnvmst` を検索 → 既存EB紐付件数を取得
5. **重複チェック③**: EBコードのみで `kcdcnvmst` を検索 → 既存ガンマ紐付件数を取得
6. **分岐処理**:
   - 両方リンク済 → 登録不可メッセージ
   - ガンマのみリンク済 → 確認ダイアログ → 既存リンク削除（`SQL_Delete`）
   - EBのみリンク済 → 確認ダイアログ → 既存リンク削除（`SQL_Delete`）
   - 両方未リンク → 新規登録確認ダイアログ
7. `SQL_INSERT_UPDATE` で `kcdcnvmst` に登録/更新
8. `AllHyouji()` で一覧を再表示

### 6.6 **リンク解消**（リンク解消.bas）

#### `PairDelete()`
1. ガンマ・EB顧客コードの入力チェック（4桁/3桁ゼロ埋め、"0000"/"000" は未入力扱い）
2. `kcdcnvmst` から該当ペアの存在確認（`Set_Array`）
3. 該当なし → "該当するﾃﾞｰﾀはありません" で終了
4. 該当あり → 確認ダイアログ（会社名付き）
5. `SQL_Delete` で `kcdcnvmst` から削除
6. エラー時はエラーメッセージ表示＋ `End`（プログラム強制終了）
7. `Touroku`（B6:F6）をクリア
8. `AllHyouji()` で一覧を再表示

### 6.7 **クリア**（クリア.bas）

#### `画面クリア()`
- `Range("Touroku") = ""` で入力行（B6:F6）をクリア
- `Range("AllData") = ""` で一覧表示エリア（B10:F193）をクリア
- ショートカットキー `Ctrl+Shift+E` が割り当てられている（`VB_Invoke_Func = "e\n14"`）

#### `EventOn()`
- `Application.EnableEvents = True` でイベントを再有効化
- デバッグ・障害復旧用のユーティリティプロシージャ

### 6.8 **該当表示**（該当表示.bas）

#### `会社名表示()`
1. `GMTNo` からガンマ顧客コードを取得（空なら終了）
2. `KaiName`, `Jyuusyo1`, `Jyuusyo2`, `EBTNo` をクリア
3. ガンマコードを4桁ゼロ埋め
4. `tokumst` と `kcdcnvmst` を外部結合（Oracle構文 `(+)`）でSELECT
5. 0件 → "照射管理ｼｽﾃﾑに登録されていません" メッセージ＋ガンマコードクリア
6. 2件以上 → "管理者に連絡してください" メッセージ
7. 1件 → 会社名・住所1・住所2・EBコードをシートに表示

#### `既存紐付チェック()`
1. `EBTNo` からEBコードを取得（3桁ゼロ埋め）
2. `tokumst` と `kcdcnvmst` を外部結合でSELECT
3. 0件 → 正常終了（紐付なし）
4. 2件以上 → "管理者に連絡してください" メッセージ
5. 1件かつ現在のガンマコードと同じ → 正常終了（同一ペア）
6. 1件かつ別のガンマコード → "リンクを解消してから登録してください" メッセージ＋EBコードクリア

### 6.9 **共通変数**（共通変数.bas）

| 変数名 | 型 | 用途 |
|--------|-----|------|
| `mpGMKaicd` | String | ガンマの会社コード（`登録更新()`, `会社名表示()`, `PairDelete()` で使用） |
| `mpEBKaicd` | String | EBの会社コード（`登録更新()`, `既存紐付チェック()`, `PairDelete()` で使用） |
| `mpKaiName` | String | 会社名（`会社名表示()` で格納、`PairDelete()` の確認ダイアログで表示） |

### 6.10 **全データ表示**（全データ表示.bas）

#### `AllHyouji()`
1. `tokumst` と `kcdcnvmst` を内部結合し、`kaisyacd` 昇順で全件取得
2. `AllData`（B10:F193）をクリア
3. `Disp_Sheet()` で結果をB10から展開（ヘッダー行出力なし: `myF = 0`）

---

## 7. ユーザーフォーム仕様

ユーザーフォーム（.frm）は存在しない。全ての操作はシート上のボタンとセル入力で行われる。

---

## 8. DB接続・外部連携

### 8.1 ODBC接続情報

| 項目 | 値 |
|------|-----|
| 接続方式 | ADODB（ActiveX Data Objects） |
| 接続文字列 | `DSN=ricdb;UID=ric;PWD=t6101` |
| DSN名 | `ricdb` |
| ユーザーID | `ric` |
| パスワード | `t6101` |
| カーソル位置 | `adUseClient`（クライアントサイド） |
| DB種別 | Oracle（SQL構文に `(+)` 外部結合、`DELETE テーブル名 WHERE ...` 形式を使用） |
| トランザクション管理 | `BeginTrans` / `CommitTrans`（INSERT/UPDATE/DELETE時） |

### 8.2 テーブル一覧

| ✓ | テーブル名 | 用途 | 操作種別 |
|---|-----------|------|----------|
| ✓ | `kcdcnvmst` | 会社コード変換マスタ（ガンマ⇔EBの顧客コード対応表） | SELECT / INSERT / UPDATE / DELETE |
|   | `tokumst` | 得意先マスタ（ガンマ側の顧客情報） | SELECT |

**`kcdcnvmst`（会社コード変換マスタ）推定カラム**:

| カラム名 | 型（推定） | 用途 |
|----------|-----------|------|
| `kaisyacd` | VARCHAR | ガンマ顧客コード（4桁） |
| `ebkaisyacd` | VARCHAR | EB顧客コード（3桁） |

**`tokumst`（得意先マスタ）推定カラム**:

| カラム名 | 型（推定） | 用途 |
|----------|-----------|------|
| `kaisyacd` | VARCHAR | ガンマ顧客コード（4桁） |
| `coname` | VARCHAR | 会社名 |
| `jyus1` | VARCHAR | 住所1 |
| `jyus2` | VARCHAR | 住所2 |

### 8.3 SQL一覧

| # | 発行元 | SQL概要 | 対象テーブル | 操作 |
|---|--------|---------|-------------|------|
| 1 | `登録更新()` | ガンマ＋EBペアの存在確認 | `kcdcnvmst` | SELECT |
| 2 | `登録更新()` | ガンマコードの既存リンク確認 | `kcdcnvmst` | SELECT |
| 3 | `登録更新()` | EBコードの既存リンク確認 | `kcdcnvmst` | SELECT |
| 4 | `SQL_INSERT_UPDATE()` | 既存レコード件数チェック | （引数で指定） | SELECT COUNT |
| 5 | `SQL_INSERT_UPDATE()` | 新規登録 | `kcdcnvmst` | INSERT |
| 6 | `SQL_INSERT_UPDATE()` | 既存更新 | `kcdcnvmst` | UPDATE |
| 7 | `SQL_Delete()` | ペア削除 | `kcdcnvmst` | DELETE |
| 8 | `会社名表示()` | ガンマコードで会社情報取得（外部結合） | `tokumst` + `kcdcnvmst` | SELECT |
| 9 | `既存紐付チェック()` | EBコードで既存リンク確認（外部結合） | `tokumst` + `kcdcnvmst` | SELECT |
| 10 | `AllHyouji()` | 全登録済みペア一覧取得（内部結合） | `tokumst` + `kcdcnvmst` | SELECT |

**代表的なSQL文**:

```sql
-- #8: 会社名表示（ガンマコードで得意先＋変換マスタを外部結合）
SELECT g.coname, g.jyus1, g.jyus2, e.ebkaisyacd
FROM tokumst g, kcdcnvmst e
WHERE g.kaisyacd = '{mpGMKaicd}' AND g.kaisyacd = e.kaisyacd(+)

-- #10: 全データ表示（登録済みペア一覧）
SELECT k.kaisyacd, k.ebkaisyacd, t.coname, t.jyus1, t.jyus2
FROM tokumst t, kcdcnvmst k
WHERE t.kaisyacd = k.kaisyacd
ORDER BY k.kaisyacd
```

### 8.4 外部ファイル参照

外部ファイル参照・外部リンクは存在しない。

---

## 9. データフロー

### 9.1 データフローテーブル

| # | トリガー | 起点 | 処理 | 終点 | 場所マーク |
|---|---------|------|------|------|-----------|
| 1 | ブック開 | 🖥️ Excel起動 | 📄 `Workbook_Open()` → `画面クリア()` | 📊 `Touroku`, `AllData` クリア | 📄📊 |
| 2 | ブック開 | 📄 `Workbook_Open()` | 📄 `AllHyouji()` → 🗄️ SELECT tokumst+kcdcnvmst | 📊 `AllData`（B10:F193）に展開 | 📄🗄️📊 |
| 3 | B6入力 | 🖥️ ガンマ顧客コード入力 | 📄 `Worksheet_Change()` → `会社名表示()` → 🗄️ SELECT tokumst+kcdcnvmst | 📊 `KaiName`, `Jyuusyo1`, `Jyuusyo2`, `EBTNo` に表示 | 🖥️📄🗄️📊 |
| 4 | C6入力 | 🖥️ EB顧客コード入力 | 📄 `Worksheet_Change()` → `既存紐付チェック()` → 🗄️ SELECT tokumst+kcdcnvmst | 🖥️ 既存リンク警告 or 正常通過 | 🖥️📄🗄️ |
| 5 | ボタン「登録」 | 🖥️ ボタンクリック | 📄 `登録更新()` → 🗄️ SELECT/INSERT/UPDATE/DELETE kcdcnvmst → `AllHyouji()` | 📊 `AllData` 再表示 | 🖥️📄🗄️📊 |
| 6 | ボタン「リンク解消」 | 🖥️ ボタンクリック | 📄 `PairDelete()` → 🗄️ SELECT/DELETE kcdcnvmst → `AllHyouji()` | 📊 `Touroku` クリア + `AllData` 再表示 | 🖥️📄🗄️📊 |
| 7 | ボタン「終了」 | 🖥️ ボタンクリック | 📄 `Bookを閉じる()` | 🖥️ ブック閉 | 🖥️📄 |

### 9.2 データフローツリー図

```
🖥️ ユーザー操作
├── 📊 ブックを開く
│   └── 📄 Workbook_Open()
│       ├── 📄 画面クリア()
│       │   └── 📊 Touroku / AllData クリア
│       ├── 📄 AllHyouji()
│       │   ├── 🗄️ SELECT tokumst JOIN kcdcnvmst（全件）
│       │   └── 📊 AllData（B10:F193）に CopyFromRecordset
│       └── 📊 B9 AutoFilter 設定 → B6 選択
│
├── 📊 B6（ガンマ顧客コード）入力
│   └── 📄 Worksheet_Change()
│       ├── [空欄] → 📊 Touroku クリア
│       └── [値あり] → 📄 会社名表示()
│           ├── 🗄️ SELECT tokumst LEFT JOIN kcdcnvmst（ガンマコード検索）
│           ├── [0件] → 🖥️ "照射管理ｼｽﾃﾑに未登録" メッセージ
│           ├── [2件+] → 🖥️ "管理者に連絡" メッセージ
│           └── [1件] → 📊 KaiName, Jyuusyo1, Jyuusyo2, EBTNo に表示
│
├── 📊 C6（EB顧客コード）入力
│   └── 📄 Worksheet_Change()
│       ├── [ガンマ未入力] → 🖥️ "最初にｶﾞﾝﾏ顧客ｺｰﾄﾞを指定" メッセージ
│       └── [ガンマ入力済] → 📄 既存紐付チェック()
│           ├── 🗄️ SELECT tokumst LEFT JOIN kcdcnvmst（EBコード検索）
│           ├── [0件] → 正常通過
│           ├── [同一ペア] → 正常通過
│           └── [別ペアあり] → 🖥️ "リンクを解消してから登録" メッセージ
│
├── 🖥️ ボタン「登録」クリック
│   └── 📄 登録更新()
│       ├── 🗄️ SELECT kcdcnvmst（ペア存在確認 ×3パターン）
│       ├── [既存リンク済] → 🖥️ 確認ダイアログ → 🗄️ DELETE kcdcnvmst
│       ├── 🗄️ INSERT / UPDATE kcdcnvmst（SQL_INSERT_UPDATE）
│       └── 📄 AllHyouji() → 🗄️📊 一覧再表示
│
├── 🖥️ ボタン「リンク解消」クリック
│   └── 📄 PairDelete()
│       ├── 🗄️ SELECT kcdcnvmst（存在確認）
│       ├── 🖥️ 確認ダイアログ（会社名付き）
│       ├── 🗄️ DELETE kcdcnvmst
│       ├── 📊 Touroku クリア
│       └── 📄 AllHyouji() → 🗄️📊 一覧再表示
│
└── 🖥️ ボタン「終了」クリック
    └── 📄 Bookを閉じる()
        └── 🖥️ ブック閉 / Excel終了
```

---

## 10. セキュリティ注意事項

### olevba 警告一覧

| 種別 | キーワード | 説明 |
|------|-----------|------|
| AutoExec | `Workbook_Open` | ブック開時に自動実行される |
| AutoExec | `Workbook_BeforeClose` | ブック閉時に自動実行される |
| AutoExec | `Worksheet_Change` | セル変更時にActiveXオブジェクト経由でイベント発火 |
| Suspicious | `Open` | ファイルを開く可能性のあるコード（実際はDB接続の `oraconn.Open`） |
| Suspicious | `Call` | Excel 4 マクロ（XLM/XLF）によるDLL呼出の可能性（実際はVBA内のSubプロシージャ呼出） |
| Suspicious | `Chr` | 文字列難読化の可能性（実際は `Chr(13)` による改行コード生成） |
| Suspicious | Hex Strings | 16進エンコード文字列の検出（フォームコントロールのバイナリデータ由来） |
| Suspicious | Base64 Strings | Base64エンコード文字列の検出（フォームコントロールのバイナリデータ由来） |

### DB接続情報のハードコーディング

- `SQL_Execution.bas` 内に接続文字列 `DSN=ricdb;UID=ric;PWD=t6101` が平文で記述されている
- パスワード `t6101` がソースコードに直接埋め込まれており、ファイルを入手した第三者に漏洩するリスクがある

### エラーハンドリングの懸念

- `On Error Resume Next` が多用されており、DB操作のエラーが無視される箇所がある
- `SQL_INSERT_UPDATE()` 内で `CommitTrans` 後に `rs.Close` が呼ばれるが、`Set rs = Nothing` が先に実行されるため、正常系でもオブジェクト参照エラーが発生する（`On Error Resume Next` により無視）
- `PairDelete()` 内でエラー発生時に `End` ステートメントで強制終了するため、DB接続のクリーンアップが行われない可能性がある
