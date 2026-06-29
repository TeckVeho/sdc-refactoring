# ExRic3詰替作業 仕様書

| 項目 | 内容 |
|------|------|
| 対象ファイル | ExRic3詰替作業.xlsm |
| 解析日 | 2026-06-29 |
| マクロ | あり (VBA) |
| 業務概要 | 照射済み製品の積替え（詰替え）作業管理。DBから積替え対象製品を抽出・一覧表示し、積替え要/不要フラグの登録更新を行う |

---

## 1. ファイル構成

### 1.1 シート一覧

| # | シート名 | codeName | 用途 | ✓ |
|---|----------|----------|------|---|
| 1 | 「積替品」 | Sheet5 | 積替え品一覧の表示・印刷シート | ✓ |
| 2 | 「積替TB」 | Sheet4 | 積替え製品マスタテーブル（会社・製品・積替えフラグ管理） | ✓ |
| 3 | 「WorkTB」 | Sheet6 | DB抽出データの一時格納領域 | |

### 1.2 ユーザーフォーム一覧

なし

### 1.3 VBAモジュール一覧

| # | モジュール名 | 種別 | 行数 | 用途 | ✓ |
|---|-------------|------|------|------|---|
| 1 | **ThisWorkbook** | cls | 15 | ブック開閉時イベント（初期化・画面クリア・積替品抽出呼び出し） | ✓ |
| 2 | **SQL_Execution** | bas | 170 | DB接続・SQL実行・シート転記・配列格納の汎用ルーチン | ✓ |
| 3 | **Ex画面クリア** | bas | 12 | 名前付き範囲のクリア・印刷範囲初期化・シート遷移 | ✓ |
| 4 | **Ex詰替品更新** | bas | 50 | 積替TB上の変更差分をDBテーブル`ExSeihinZ`へINSERT/UPDATE | ✓ |
| 5 | **Ex積替え品抽出** | bas | 132 | DBから積替え対象品を抽出しWorkTB経由で積替品シートに加工表示 | ✓ |
| 6 | **Ex積替品表示** | bas | 30 | 積替TBシートに全製品マスタ＋積替えフラグを表示 | ✓ |
| 7 | **ExFunction** | bas | 15 | 日付フォーマット変換ユーティリティ | |
| 8 | **印刷範囲** | bas | 12 | 積替品シートの印刷範囲設定 | |
| 9 | **終了処理** | bas | 10 | ブックを閉じる（上書きなし） | |
| 10 | **Sheet4** | cls | 1 | 「積替TB」シートモジュール（コードなし） | |
| 11 | **Sheet5** | cls | 1 | 「積替品」シートモジュール（コードなし） | |
| 12 | **Sheet6** | cls | 1 | 「WorkTB」シートモジュール（コードなし） | |

---

## 2. シート詳細

### 2.0 シート可視性一覧

| シート名 | 状態 |
|----------|------|
| 「積替品」 | 表示 |
| 「積替TB」 | 表示 |
| 「WorkTB」 | 表示 |

### 2.1 「積替品」シート

| 用途 | 積替え品の照射状況一覧表示・印刷用シート |
|------|----------------------------------------|
| 範囲 | 235行 × 15列 |
| codeName | Sheet5 |

#### ヘッダ領域（行1〜4）

| セル | 内容 | 備考 | ✓ |
|------|------|------|---|
| `B1` | "New" | ※推論: バージョンラベル | |
| `B3` | "積替え品一覧" | シートタイトル | |
| `F3` | 32 | 印刷範囲行数（名前付き範囲`Innsatu`） | ✓ |
| `G3:H3` | "まで印刷　Max" | 結合セル | |
| `I3` | `=MAX(B6:B235)` | 現在のNo最大値 | |
| `L3` | `=NOW()` | 現在日時 | |
| `M3` | "Ex4" | ※推論: システム識別子 | |

#### データ領域ヘッダ（行5）

| 列 | ヘッダ | 業務的意味 | ✓ |
|----|--------|-----------|---|
| B | No | 連番 | |
| D | 最終積替え日 | 出荷日-1日で算出 | ✓ |
| E | 照射状況 | i/0/C→未、1→中、2→済 に変換 | ✓ |
| F | 線量計番号 | 下4桁表示 | |
| G | 受付番号 | 下4桁表示 | |
| H | 会社名 | 「株式会社」除去済み | |
| I | 納期 | mm/dd形式 | |
| J | 出荷日 | DB値をDate型変換 | ✓ |
| K | ﾊﾟｽ数 | 照射パス数 | |
| L | 備考 | DB bikou1 | |
| M | ﾁｪｯｸ | 作業確認用 | |

#### データ領域

| 範囲 | 内容 | ✓ |
|------|------|---|
| `B6:M235` | 積替え品一覧データ（名前付き範囲`TumikaeTB`） | ✓ |

### 2.2 「積替TB」シート

| 用途 | 積替え対象製品マスタの表示・編集シート |
|------|--------------------------------------|
| 範囲 | 10000行 × 12列 |
| codeName | Sheet4 |

#### ヘッダ領域

| セル | 内容 | 備考 | ✓ |
|------|------|------|---|
| `E1` | "積み替え製品テーブル" | シートタイトル | |
| `D3` | `=COUNTA(SeihinnSuu)` | 登録製品数 | ✓ |
| `D4` | "登録製品毎に積替え品登録してください" | 操作案内 | |

#### データ領域ヘッダ（行5）

| 列 | ヘッダ | 業務的意味 | ✓ |
|----|--------|-----------|---|
| D | 会社ｺｰﾄﾞ | 4桁ゼロ埋め | ✓ |
| E | 製品ｺｰﾄﾞ | 3桁ゼロ埋め | ✓ |
| F | 会社ｺｰﾄﾞ&製品コード | 結合キー | |
| G | 会社名 | tokumst.coname | |
| H | 製品名 | sehmst.seiname | |
| I | 詰替え要不要 | ユーザー編集対象（"1"=要） | ✓ |
| J | ExSeihin Kaisyacd | DB現在値（比較用） | ✓ |

#### データ領域

| 範囲 | 内容 | ✓ |
|------|------|---|
| `D6:J10000` | 製品マスタデータ（名前付き範囲`TumikaeHinn`） | ✓ |

#### 数式

| セル | 数式 | 業務的意味 |
|------|------|-----------|
| `I3` | `=IF(ISERROR(VLOOKUP(#REF!,TumiFlg,2,FALSE)),"",VLOOKUP(#REF!,TumiFlg,2,FALSE))` | ※参照エラー状態（`TumiFlg`は未定義） |
| `J3` | `=IF(ISERROR(VLOOKUP(#REF!,TumiFlg,2,FALSE)),"",VLOOKUP(#REF!,TumiFlg,2,FALSE))` | 同上 |

### 2.3 「WorkTB」シート

| 用途 | DB抽出結果の一時格納ワークシート |
|------|-------------------------------|
| 範囲 | 133行 × 9列 |
| codeName | Sheet6 |

#### セル

| セル | 内容 | 備考 |
|------|------|------|
| `I1` | 1 | ※推論: フラグまたはカウンタ |

VBAからは名前付き範囲`Work`(`A1:H201`)および`Wtb`(`A1:H250`)で参照される。

---

## 3. 名前付き範囲一覧

| # | 名前 | 参照先 | 業務的意味 | ✓ |
|---|------|--------|-----------|---|
| 1 | `DebugFlg` | 積替品!$A$1 | デバッグフラグ | |
| 2 | `Innsatu` | 積替品!$F$3 | 印刷範囲行数（デフォルト32） | ✓ |
| 3 | `SeihinnSuu` | 積替TB!$D$6:$D$1005 | 登録製品数カウント用範囲 | |
| 4 | `SeiKennsuu` | 積替TB!$D$3 | 登録製品件数（COUNTA結果） | ✓ |
| 5 | `TumeTB` | 積替TB!$G$6:$I$1005 | 会社名・製品名・積替えフラグの表示範囲 | |
| 6 | `TumikaeHinn` | 積替TB!$D$6:$J$10000 | 積替え製品マスタ全データ範囲 | ✓ |
| 7 | `TumikaeTB` | 積替品!$B$6:$M$235 | 積替品一覧表示データ範囲 | ✓ |
| 8 | `Work` | WorkTB!$A$1:$H$201 | ワーク領域（DB抽出一時格納） | ✓ |
| 9 | `Wtb` | WorkTB!$A$1:$H$250 | ワーク領域拡張（クリア用） | ✓ |

---

## 4. 数式一覧

| # | シート | セル | 数式 | 業務的意味 |
|---|--------|------|------|-----------|
| 1 | 「積替品」 | `I3` | `=MAX(B6:B235)` | 一覧の最大No表示 |
| 2 | 「積替品」 | `L3` | `=NOW()` | 現在日時の表示 |
| 3 | 「積替TB」 | `D3` | `=COUNTA(SeihinnSuu)` | 登録製品件数のカウント |
| 4 | 「積替TB」 | `I3` | `=IF(ISERROR(VLOOKUP(#REF!,TumiFlg,2,FALSE)),"",VLOOKUP(#REF!,TumiFlg,2,FALSE))` | 参照エラー状態 |
| 5 | 「積替TB」 | `J3` | `=IF(ISERROR(VLOOKUP(#REF!,TumiFlg,2,FALSE)),"",VLOOKUP(#REF!,TumiFlg,2,FALSE))` | 参照エラー状態 |

---

## 5. ボタン・マクロ対応

### 5.1 シート上のボタン

#### 「積替品」シート（vmlDrawing1.vml）

| # | ボタンテキスト | 呼び出しマクロ | 機能 | ✓ |
|---|---------------|---------------|------|---|
| 1 | 現状の積替品の照射状況 | `TumikaeHinn()` | DBから積替え対象品を抽出し一覧表示 | ✓ |
| 2 | 積替する製品の登録 | `積替製品TB表示()` | 積替TBシートに製品マスタを表示し遷移 | ✓ |
| 3 | 画面を閉じる | `Bookを閉じる()` | ブックを保存せず閉じる | |

#### 「積替TB」シート（vmlDrawing2.vml）

| # | ボタンテキスト | 呼び出しマクロ | 機能 | ✓ |
|---|---------------|---------------|------|---|
| 1 | 戻る | `Modori()` | 積替品シートに戻る | |
| 2 | 登録内容更新 | `詰替品データ更新()` | 変更行をDBへINSERT/UPDATE | ✓ |
| 3 | データ表示製品検索 | `積替製品TB表示()` | 製品マスタを再取得表示 | ✓ |

### 5.2 フォーム上のボタン

なし

### 5.3 CommandBar

コメントアウト済み（`Application.CommandBars("Worksheet Menu Bar").Enabled = True`）。現在は無効。

---

## 6. VBAモジュール仕様

### 6.0 全プロシージャ一覧

| # | モジュール | プロシージャ | 種別 | 場所 | ✓ |
|---|-----------|-------------|------|------|---|
| 1 | **ThisWorkbook** | `Workbook_BeforeClose()` | Event | 📄VBA内部処理 | |
| 2 | **ThisWorkbook** | `Workbook_Open()` | Event | 📄VBA内部処理 | ✓ |
| 3 | **SQL_Execution** | `Open_oraconDB()` | Sub | 🗄️DB操作 | ✓ |
| 4 | **SQL_Execution** | `SQL_Exe()` | Sub | 🗄️DB操作 | ✓ |
| 5 | **SQL_Execution** | `SQL_INSERT_UPDATE()` | Sub | 🗄️DB操作 | ✓ |
| 6 | **SQL_Execution** | `SQL_Delete()` | Sub | 🗄️DB操作 | ✓ |
| 7 | **SQL_Execution** | `Disp_Sheet()` | Sub | 🗄️DB操作 📊シート操作 | ✓ |
| 8 | **SQL_Execution** | `Set_Array()` | Sub | 🗄️DB操作 | ✓ |
| 9 | **Ex画面クリア** | `画面クリア3詰替()` | Sub | 📊シート操作 | ✓ |
| 10 | **Ex画面クリア** | `modori()` | Sub | 🖥️画面操作 | |
| 11 | **Ex詰替品更新** | `詰替品データ更新()` | Sub | 🗄️DB操作 | ✓ |
| 12 | **Ex積替え品抽出** | `TumikaeHinn()` | Sub | 🗄️DB操作 📊シート操作 | ✓ |
| 13 | **Ex積替え品抽出** | `DataKakou()` | Sub | 📄VBA内部処理 📊シート操作 | ✓ |
| 14 | **Ex積替品表示** | `積替製品TB表示()` | Sub | 🗄️DB操作 📊シート操作 | ✓ |
| 15 | **ExFunction** | `ExchengeDATE()` | Function | 📄VBA内部処理 | |
| 16 | **印刷範囲** | `InsatuHanni()` | Sub | 📊シート操作 | |
| 17 | **終了処理** | `Bookを閉じる()` | Sub | 🖥️画面操作 | |

### 6.1 **ThisWorkbook**（ThisWorkbook.cls）

#### `Workbook_BeforeClose(Cancel As Boolean)`
- DisplayAlertsをFalseにし保存済みフラグを設定して無警告で閉じる

#### `Workbook_Open()`
- ブック保護を解除
- ウィンドウ最大化
- 「積替品」シートを選択
- `画面クリア3詰替()` を呼び出し（全範囲クリア）
- `TumikaeHinn()` を呼び出し（DB抽出→一覧表示）

### 6.2 **SQL_Execution**（SQL_Execution.bas）

モジュール変数:
- `mpErrDes As String` — エラー記述（Public）
- `oraconn As New ADODB.Connection` — DB接続オブジェクト
- `rs As ADODB.Recordset` — レコードセット

#### `Open_oraconDB()`
- ODBC DSN `ricdb` でDB接続を確立
- CursorLocationを`adUseClient`に設定

#### `SQL_Exe(mySQL As String)`
- 引数のSQL文をExecuteで実行し結果を`rs`に格納
- エラー発生時は`mpErrDes`にエラー内容を記録

#### `SQL_INSERT_UPDATE(myTBL, myKey, myD(), myN)`
- テーブル`myTBL`に対し、`myKey`条件でレコード存在チェック
- 存在しない場合: INSERT文を動的生成し実行
- 存在する場合: UPDATE文を動的生成し実行
- トランザクション制御あり（BeginTrans→CommitTrans）

#### `SQL_Delete(myTBL, myWhere)`
- テーブル`myTBL`から`myWhere`条件でDELETE実行
- トランザクション制御あり

#### `Disp_Sheet(mySQL, mySH, myRow, myRecordCount, myColumn, myFieldCount, myF)`
- SQL実行結果をシート`mySH`の`myRow`行`myColumn`列から転記
- `myF=1`の場合はフィールド名ヘッダも出力
- `CopyFromRecordset`でデータを一括貼り付け
- 出力パラメータ: `myRecordCount`（レコード数）, `myFieldCount`（フィールド数）

#### `Set_Array(mySQL, myData(), myRecordCount, myFldCount)`
- SQL実行結果を二次元配列`myData(i,j)`に格納
- i=レコード番号, j=フィールド番号

### 6.3 **Ex画面クリア**（Ex画面クリア.bas）

#### `画面クリア3詰替()`
- `TumikaeHinn`, `TumikaeTB`, `Work` をクリア
- `Innsatu` を32にリセット
- `InsatuHanni()` を呼び出し

#### `modori()`
- 「積替品」シートを選択（画面遷移）

### 6.4 **Ex詰替品更新**（Ex詰替品更新.bas）

#### `詰替品データ更新()`
- 「積替TB」シートの`TumikaeHinn`範囲をループ
- I列（ユーザー編集値）とJ列（DB現在値）を比較
- 差分がある行のみ`ExSeihinZ`テーブルへINSERT/UPDATE
- キー: `kaisyacd`（4桁ゼロ埋め）＋ `sehncd`（3桁ゼロ埋め）
- 更新カラム: `tumikae`（積替えフラグ）
- 更新後にJ列をI列の値で上書き（差分解消）

### 6.5 **Ex積替え品抽出**（Ex積替え品抽出.bas）

#### `TumikaeHinn()`
- `Wtb`範囲をクリア
- SQL: `syoukj3`, `zaiko`, `ExKeikakuX`, `ExSeihinZ` を結合
- 条件: `ExSeihinZ.tumikae='1'`（積替え対象のみ）
- `Disp_Sheet()`でWorkTBに転記
- `DataKakou()`を呼び出しデータ加工

#### `DataKakou(myDataN)`
- WorkTBからデータを配列に読み込み
- 加工処理:
  - 出荷日→最終積替え日（出荷日-1日）
  - 照射状況コード変換（i/0/C→未、1→中、2→済）
  - 線量計番号→下4桁
  - 受付番号→下4桁
  - 会社名→「株式会社」除去
  - 納期→mm/dd形式
- 加工後データを「積替品」シート`B6:M`に書き込み
- ソート: 照射状況（昇順）→ 最終積替え日（昇順）→ 線量計番号（昇順）

### 6.6 **Ex積替品表示**（Ex積替品表示.bas）

#### `積替製品TB表示()`
- 「積替TB」シートを選択し`TumikaeHinn`をクリア
- SQL: `sehmst`, `tokumst`, `ExSeihinZ` を結合
- 条件: `sehmst.syouso='3'` AND `sehmst.ric<>'**********'`
- 取得項目: 会社コード、製品コード、結合キー、会社名、製品名、積替えフラグ×2
- `Disp_Sheet()`で積替TBシートD6から転記

### 6.7 **ExFunction**（ExFunction.bas）

#### `ExchengeDATE(ByVal myDate, ByVal myType)`
- `myType="mm/dd"`: 数値を"MM/DD"形式に変換
- `myType="yyyy/mm/dd"`: 8桁数値を"YYYY/MM/DD"形式に変換
- 数値0またはブランクの場合はTrimのみ

### 6.8 **印刷範囲**（印刷範囲.bas）

#### `InsatuHanni()`
- 名前付き範囲`Innsatu`の値に基づき印刷範囲を`$B$6:$M$n`に設定
- 32以外の場合はメッセージボックスで通知

### 6.9 **終了処理**（終了処理.bas）

#### `Bookを閉じる()`
- DisplayAlertsをFalseに設定
- 開いているブックが1つなら`Application.Quit`
- 複数なら`ActiveWorkbook.Close`

### 6.10 **Sheet4**（Sheet4.cls）

- 「積替TB」のシートモジュール。コードなし（`Option Explicit`のみ）

### 6.11 **Sheet5**（Sheet5.cls）

- 「積替品」のシートモジュール。コードなし（`Option Explicit`のみ）

### 6.12 **Sheet6**（Sheet6.cls）

- 「WorkTB」のシートモジュール。コードなし（`Option Explicit`のみ）

---

## 7. ユーザーフォーム仕様

なし（本ブックにユーザーフォームは含まれない）

---

## 8. DB接続・外部連携

### 8.1 ODBC接続

| 項目 | 値 |
|------|-----|
| DSN | `ricdb` |
| UID | `ric` |
| PWD | `t6101` |
| ライブラリ | ADODB (Microsoft ActiveX Data Objects) |
| CursorLocation | adUseClient |

### 8.2 テーブル一覧

| # | テーブル名 | 用途 | INSERT | UPDATE | DELETE | SELECT | ✓ |
|---|-----------|------|--------|--------|--------|--------|---|
| 1 | `ExSeihinZ` | 製品別積替えフラグ管理 | ○ | ○ | — | ○ | ✓ |
| 2 | `syoukj3` | 照射工程管理（照射状況・線量計・受付番号・会社名） | — | — | — | ○ | |
| 3 | `zaiko` | 在庫管理（受付番号・会社コード・製品コード・納期・パス数） | — | — | — | ○ | |
| 4 | `ExKeikakuX` | 計画管理（出荷日・備考） | — | — | — | ○ | |
| 5 | `tokumst` | 得意先マスタ（会社名） | — | — | — | ○ | |
| 6 | `sehmst` | 製品マスタ（製品名・照射所・会社コード） | — | — | — | ○ | |

### 8.3 SQL一覧

#### SQL-1: 積替え品照射状況抽出（`TumikaeHinn()`）

```sql
SELECT s.syoush_f, s.sykno, s.uno1, s.kainame1,
       z.nouki, k.syukkabi, z.pass, k.bikou1
FROM syoukj3 s, zaiko z, ExKeikakuX k, ExSeihinz e
WHERE s.uno1 = z.uno
  AND s.uno1 = k.uno(+)
  AND z.kaisyacd = e.kaisyacd(+)
  AND z.sehncd = e.sehncd(+)
  AND e.tumikae = '1'
ORDER BY s.syoush_f DESC, s.sykno
```

| 結合 | 条件 | 種別 |
|------|------|------|
| syoukj3 ↔ zaiko | s.uno1 = z.uno | 内部結合 |
| syoukj3 ↔ ExKeikakuX | s.uno1 = k.uno(+) | 外部結合（Oracle構文） |
| zaiko ↔ ExSeihinZ | z.kaisyacd = e.kaisyacd(+) AND z.sehncd = e.sehncd(+) | 外部結合（Oracle構文） |

#### SQL-2: 積替製品マスタ取得（`積替製品TB表示()`）

```sql
SELECT s.kaisyacd, s.sehncd, s.kaisyacd || s.sehncd,
       TRIM(t.coname), TRIM(s.seiname), e.tumikae, e.tumikae
FROM tokumst t, sehmst s, ExSeihinZ e
WHERE s.kaisyacd = t.kaisyacd(+)
  AND s.syouso = '3'
  AND s.kaisyacd = e.kaisyacd(+)
  AND s.sehncd = e.sehncd(+)
  AND s.ric <> '**********'
ORDER BY s.kaisyacd, s.sehncd
```

| 結合 | 条件 | 種別 |
|------|------|------|
| sehmst ↔ tokumst | s.kaisyacd = t.kaisyacd(+) | 外部結合 |
| sehmst ↔ ExSeihinZ | s.kaisyacd = e.kaisyacd(+) AND s.sehncd = e.sehncd(+) | 外部結合 |

#### SQL-3: INSERT/UPDATE（`詰替品データ更新()`→`SQL_INSERT_UPDATE()`）

```sql
-- 存在チェック
SELECT COUNT(*) FROM ExSeihinZ WHERE kaisyacd='XXXX' AND sehncd='XXX'

-- INSERT（レコードなしの場合）
INSERT INTO ExSeihinZ (kaisyacd, sehncd, tumikae)
VALUES ('XXXX', 'XXX', '1')

-- UPDATE（レコードありの場合）
UPDATE ExSeihinZ SET kaisyacd='XXXX', sehncd='XXX', tumikae='1'
WHERE kaisyacd='XXXX' AND sehncd='XXX'
```

### 8.4 外部ファイル

なし（外部リンク・外部ファイル参照は検出されず）

---

## 9. データフロー

### 9.1 データフローテーブル

| # | 起点 | → | 終点 | トリガー | 内容 |
|---|------|---|------|---------|------|
| 1 | 🗄️ DB (syoukj3, zaiko, ExKeikakuX, ExSeihinZ) | → | 📊「WorkTB」 | ブック起動 / ボタン「現状の積替品の照射状況」 | 積替え対象品の照射状況データ |
| 2 | 📊「WorkTB」 | → | 📄 VBA配列 (myData) | `DataKakou()` | ワーク読み込み＋加工処理 |
| 3 | 📄 VBA配列 | → | 📊「積替品」B6:M235 | `DataKakou()` | 加工済み積替え品一覧の表示 |
| 4 | 🗄️ DB (sehmst, tokumst, ExSeihinZ) | → | 📊「積替TB」D6:J | ボタン「積替する製品の登録」/ ボタン「データ表示製品検索」 | 全製品マスタ＋積替えフラグ |
| 5 | 📊「積替TB」I列（ユーザー編集） | → | 🗄️ DB ExSeihinZ | ボタン「登録内容更新」 | 差分行の積替えフラグ更新 |

### 9.2 データフローツリー図

```
📂 ExRic3詰替作業
├── 🖥️ ブック起動 (Workbook_Open)
│   ├── 📄 画面クリア3詰替() ─── 📊「積替品」「積替TB」「WorkTB」全クリア
│   └── 📄 TumikaeHinn()
│       ├── 🗄️ SELECT syoukj3+zaiko+ExKeikakuX+ExSeihinZ
│       ├── 📊「WorkTB」← DB結果転記 (Disp_Sheet)
│       └── 📄 DataKakou()
│           ├── 📊「WorkTB」→ 配列読み込み
│           ├── 📄 データ加工（日付変換・コード変換・文字列整形）
│           └── 📊「積替品」B6:M235 ← 加工済みデータ書き込み＋ソート
│
├── 🖥️ ボタン「積替する製品の登録」
│   └── 📄 積替製品TB表示()
│       ├── 🗄️ SELECT sehmst+tokumst+ExSeihinZ
│       └── 📊「積替TB」D6:J ← 製品マスタ転記 (Disp_Sheet)
│
├── 🖥️ ボタン「登録内容更新」
│   └── 📄 詰替品データ更新()
│       ├── 📊「積替TB」I列/J列 比較ループ
│       ├── 🗄️ INSERT/UPDATE ExSeihinZ (差分行のみ)
│       └── 📊「積替TB」J列 ← I列で上書き（同期）
│
├── 🖥️ ボタン「現状の積替品の照射状況」
│   └── 📄 TumikaeHinn() ─── （上記と同一フロー）
│
├── 🖥️ ボタン「戻る」
│   └── 📄 modori() ─── 📊「積替品」シート選択
│
└── 🖥️ ボタン「画面を閉じる」
    └── 📄 Bookを閉じる() ─── 🖥️ ブック終了（保存なし）
```

---

## 10. セキュリティ注意事項

| # | 区分 | 内容 | リスク |
|---|------|------|--------|
| 1 | 認証情報ハードコード | DSN=`ricdb`, UID=`ric`, PWD=`t6101` が **SQL_Execution** モジュールに平文記載 | 高 |
| 2 | SQLインジェクション | `SQL_INSERT_UPDATE()`はセル値を文字列連結でSQL構築。シングルクォートのエスケープなし | 中 |
| 3 | エラーハンドリング | `On Error Resume Next` の多用によりDB操作失敗が無視される可能性 | 中 |
| 4 | トランザクション | `SQL_INSERT_UPDATE()`内で1件ずつCommitするため、ループ途中のエラーで不整合が発生する可能性 | 中 |
| 5 | ブック保護 | `Workbook_Open`で`Protect Structure:=False`を実行しブック保護を解除 | 低 |
| 6 | 保存なし終了 | `Workbook_BeforeClose`で`Saved=True`を強制設定し変更破棄を無警告化 | 低 |

---

*以上*
