# Ex線量検索 仕様書

> **ファイル種別**: .xlsm（マクロ付き）
> **用途**: 照射実績データ（在庫・履歴）をDBから条件指定で検索し、結果を新規ブックに出力する
> **VBA プロジェクトサイズ**: 8モジュール（ThisWorkbook, Sheet1, SQL_Execution, DBRead, 実行, 終了処理, 共通変数, Function_Sub）
> **外部連携ファイル**: なし（DB直接接続）

---

## 目次

1. [ファイル構成](#1-ファイル構成)
2. [シート詳細](#2-シート詳細)
3. [名前付き範囲一覧](#3-名前付き範囲一覧)
4. [数式一覧](#4-数式一覧)
5. [ボタン・マクロ対応](#5-ボタンマクロ対応)
6. [VBA モジュール仕様](#6-vba-モジュール仕様)
7. [DB 接続・外部連携](#7-db-接続外部連携)
8. [データフロー](#8-データフロー)
9. [セキュリティ注意事項](#9-セキュリティ注意事項)

---

## 1. ファイル構成

```
Ex線量検索.xlsm
├── シート
│   └── 条件（検索条件入力・実行）
└── VBA モジュール
    ├── ThisWorkbook.cls    （Workbook_Open / Workbook_BeforeClose）
    ├── Sheet1.cls          （空）
    ├── 共通変数.bas        （Public 変数定義）
    ├── Function_Sub.bas    （BookName 関数）
    ├── SQL_Execution.bas   （ADO/ODBC DB接続・SQL実行基盤）
    ├── DBRead.bas          （条件設定・受付番号範囲・製品コード指定・条件表示）
    ├── 実行.bas            （検索実行処理）
    └── 終了処理.bas        （ブック終了処理）
```

---

## 2. シート詳細

### 2.1 条件

**目的**: 照射実績データの検索条件（抽出データ種別・表示列・検索条件・ソート）を入力し、「実行」ボタンでDB検索を起動する。

#### レイアウト構造

| セル範囲 | 名前付き範囲 | 内容 |
|---|---|---|
| I1 | — | システム識別子（`Ex4`） |
| B2:D2 | — | タイトル「照射実績ﾃﾞｰﾀ抽出」（セル結合） |
| B4:D4 | `K_Rireki` | ラベル「①抽出ﾃﾞｰﾀ」（セル結合） |
| E4 | `Rireki` | 抽出データ種別入力（`在庫` または `履歴`） |
| B6 | — | 列ヘッダー「項目」 |
| D6 | — | 列ヘッダー「②表示」 |
| E6 | — | 列ヘッダー「④検索条件」 |
| F6 | — | 列ヘッダー「ソート」 |
| L6 | — | 参照テーブル列ヘッダー「ﾃｰﾌﾞﾙ先」 |
| M6 | — | 参照テーブル列「在庫」 |
| N6 | — | 参照テーブル列「履歴」 |
| B7 | `K_Uno` | 項目ラベル「③受付番号」 |
| C7 | `s.UNO` | DBフィールド名（`s.UNO`） |
| D7 | — | 表示設定「必須」 |
| E7 | `Uno` | 受付番号入力（単一または範囲: `XXXXXXXXXX-XXXXXXXXXX`） |
| F7 | — | ソート「昇順」 |
| B8 | `K_Kaisyacd` | 項目ラベル「会社ｺｰﾄﾞ」 |
| C8 | `s.KAISYACD` | DBフィールド名 |
| D8 | `H_KAISYACD` | 表示設定（`する`/`しない`） |
| E8 | `KAISYACD` | 会社コード入力（空=全会社） |
| L8 | `SYOUSO` | 装置コード算出（HLOOKUPで変換） |
| B9 | `K_SEHNCD` | 項目ラベル「製品ｺｰﾄﾞ」 |
| C9 | `z.SEHNCD` | DBフィールド名 |
| D9 | `H_SEHNCD` | 表示設定 |
| E9 | `SEHNCD` | 製品コード入力（複数は「,」区切り） |
| B10 | `K_SYOUSO` | 項目ラベル「装置」 |
| C10 | `s.SYOUSO` | DBフィールド名 |
| D10 | `H_SYOUSO` | 表示設定 |
| E10 | `J_SYOUSO` | 装置選択入力（`Ric2`/`Ric3`/`Ric2,3`） |
| F10:F24 | — | 使い方説明テキスト（セル結合） |
| B11 | `K_TUIKAFLG` | 項目ラベル「追加照射」 |
| C11 | `s.TUIKAFLG` | DBフィールド名 |
| D11 | `H_TUIKAFLG` | 表示設定 |
| E11 | `TUIKAFLG` | 追加照射フラグ（`含む`/`含まない`/`のみ`） |
| B12 | — | 項目ラベル「線量計番号」 |
| C12 | `s.SENKNO` | DBフィールド名 |
| D12 | `H_SENKNO` | 表示設定 |
| B13 | — | 項目ラベル「厚さ」 |
| C13 | `s.ATUSA` | DBフィールド名 |
| D13 | `H_ATUSA` | 表示設定 |
| B14 | `K_KEISASK` | 項目ラベル「計算式」 |
| C14 | `s.KEISASK` | DBフィールド名 |
| D14 | `H_KEISASK` | 表示設定 |
| E14 | `KEISASK` | 計算コード入力（空=全コード） |
| B15 | `K_PASS` | 項目ラベル「指定ﾊﾟｽ」 |
| C15 | `s.PASS` | DBフィールド名 |
| D15 | `H_PASS` | 表示設定 |
| E15 | `PASS` | 指定パス入力 |
| B16 | — | 項目ラベル「実ﾊﾟｽ」 |
| C16 | `s.JITUNO` | DBフィールド名 |
| D16 | `H_JITUNO` | 表示設定 |
| B17 | `K_SITEISN` | 項目ラベル「指定線量」 |
| C17 | `s.SITEISN` | DBフィールド名 |
| D17 | `H_SITEISN` | 表示設定 |
| E17 | `SITEISN` | 指定線量入力 |
| B18 | — | 項目ラベル「下限線量」 |
| C18 | `s.KAGENSN` | DBフィールド名 |
| D18 | `H_KAGENSN` | 表示設定 |
| E18:E22 | — | 下限〜温度 検索条件入力欄（セル結合なし） |
| B19 | — | 項目ラベル「上限線量」 |
| C19 | `s.JYOUGSN` | DBフィールド名 |
| D19 | `H_JYOUGSN` | 表示設定 |
| B20 | — | 項目ラベル「ABS」 |
| C20 | `s.SOKUTTI` | DBフィールド名 |
| D20 | `H_SOKUTTI` | 表示設定 |
| B21 | — | 項目ラベル「DOSE」 |
| C21 | `s.SOKUTSN` | DBフィールド名 |
| D21 | `H_SOKUTSN` | 表示設定 |
| B22 | — | 項目ラベル「温度」 |
| C22 | `s.ONDOK` | DBフィールド名 |
| D22 | `H_ONDOK` | 表示設定 |
| B23 | — | 項目ラベル「装置投入日時」 |
| C23 | `s_SYOUSYABI` | DBフィールド名 |
| D23 | `H_SYOUSYABI` | 表示設定 |
| E23 | — | 注記「ddhhmm：日時分」 |
| B24 | — | ラベル「⑥抽出件数上限」 |
| D24 | `DataN` | 抽出件数上限値（デフォルト8000） |
| L7:O8 | — | 装置コード変換テーブル（HLOOKUP参照用） |
| L10:O11 | — | ソート方向変換テーブル（ASC/DESC） |
| M13 | `ThisB_Name` | 自ファイル名（BookName関数で設定） |
| L11 | `SortH` | ソート方向（HLOOKUPで算出） |

---

## 3. 名前付き範囲一覧

| 名前 | 参照先 | 説明 |
|---|---|---|
| `DataN` | 条件!$D$24 | 抽出件数上限値 |
| `H_ATUSA` | 条件!$D$13 | 厚さ列の表示フラグ |
| `H_JITUNO` | 条件!$D$16 | 実パス列の表示フラグ |
| `H_JYOUGSN` | 条件!$D$19 | 上限線量列の表示フラグ |
| `H_KAGENSN` | 条件!$D$18 | 下限線量列の表示フラグ |
| `H_KAISYACD` | 条件!$D$8 | 会社コード列の表示フラグ |
| `H_KEISASK` | 条件!$D$14 | 計算式列の表示フラグ |
| `H_ONDOK` | 条件!$D$22 | 温度列の表示フラグ |
| `H_PASS` | 条件!$D$15 | 指定パス列の表示フラグ |
| `H_SEHNCD` | 条件!$D$9 | 製品コード列の表示フラグ |
| `H_SENKNO` | 条件!$D$12 | 線量計番号列の表示フラグ |
| `H_SITEISN` | 条件!$D$17 | 指定線量列の表示フラグ |
| `H_SOKUTSN` | 条件!$D$21 | DOSE列の表示フラグ |
| `H_SOKUTTI` | 条件!$D$20 | ABS列の表示フラグ |
| `H_SYOUSO` | 条件!$D$10 | 装置列の表示フラグ |
| `H_SYOUSYABI` | 条件!$D$23 | 装置投入日時列の表示フラグ |
| `H_TUIKAFLG` | 条件!$D$11 | 追加フラグ列の表示フラグ |
| `J_SYOUSO` | 条件!$E$10 | 装置選択（表示用ラベル） |
| `K_Kaisyacd` | 条件!$B$8 | 会社コードの項目ラベルセル |
| `K_KEISASK` | 条件!$B$14 | 計算式の項目ラベルセル |
| `K_PASS` | 条件!$B$15 | 指定パスの項目ラベルセル |
| `K_Rireki` | 条件!$B$4 | 抽出データの項目ラベルセル |
| `K_SEHNCD` | 条件!$B$9 | 製品コードの項目ラベルセル |
| `K_SITEISN` | 条件!$B$17 | 指定線量の項目ラベルセル |
| `K_SYOUSO` | 条件!$B$10 | 装置の項目ラベルセル |
| `K_TUIKAFLG` | 条件!$B$11 | 追加照射の項目ラベルセル |
| `K_Uno` | 条件!$B$7 | 受付番号の項目ラベルセル |
| `KAISYACD` | 条件!$E$8 | 会社コード検索条件入力セル |
| `KEISASK` | 条件!$E$14 | 計算コード検索条件入力セル |
| `PASS` | 条件!$E$15 | 指定パス検索条件入力セル |
| `Rireki` | 条件!$E$4 | 抽出データ種別（在庫/履歴） |
| `SEHNCD` | 条件!$E$9 | 製品コード検索条件入力セル |
| `SITEISN` | 条件!$E$17 | 指定線量検索条件入力セル |
| `SortH` | 条件!$L$11 | ソート方向（ASC/DESC） |
| `SYOUSO` | 条件!$L$8 | 装置コード（変換後の数値） |
| `ThisB_Name` | 条件!$M$13 | 自ファイル名（`Ex線量検索.xlsm`） |
| `TUIKAFLG` | 条件!$E$11 | 追加照射フラグ検索条件 |
| `Uno` | 条件!$E$7 | 受付番号入力セル |
| `s.ATUSA` ～ `z.SEHNCD` | 条件!$C$xx | 各DBフィールド名参照セル（VBA動的SQL構築用） |
| `ShName` | 条件!#REF! | 参照エラー（不使用） |

---

## 4. 数式一覧

### 条件シート

| セル | 数式 | 説明 |
|---|---|---|
| L8 | `=IF(ISERROR(HLOOKUP(E10,M7:O8,2,FALSE)),"",HLOOKUP(E10,M7:O8,2,FALSE))` | 装置選択（E10）をDB用装置コード（数値）に変換。変換テーブルはM7:O8 |
| L11 | `=IF(F7="","ASC",HLOOKUP(F7,M10:O11,2,FALSE))` | ソート選択（F7）をSQL ORDER BY 用文字列（ASC/DESC）に変換 |

---

## 5. ボタン・マクロ対応

| シート | ボタンラベル | 割り当てマクロ | 動作概要 |
|---|---|---|---|
| 条件 | 実行 | `検索実行` | 入力条件でDBを検索し、結果を新規ブックに出力する |
| 条件 | 終了 | `終了処理Sub` | ブックを保存せずに終了する |

---

## 6. VBA モジュール仕様

### 6.1 ThisWorkbook.cls

#### `Workbook_Open()`

**処理概要**: ブック起動時に条件シートを表示・保護設定し、ファイル名を `ThisB_Name` セルに書き込む。

```vba
Private Sub Workbook_Open()
    Worksheets("条件").Select
    ActiveSheet.Unprotect
    ActiveSheet.Protect UserInterfaceOnly:=True, DrawingObjects:=False
    Range("ThisB_Name") = BookName
End Sub
```

#### `Workbook_BeforeClose(Cancel As Boolean)`

**処理概要**: 保存ダイアログを抑制してブックを閉じる。

---

### 6.2 Function_Sub.bas

#### `BookName() As String`

**処理概要**: `Application.Volatile` で揮発性関数として定義。ブック名（`Ex線量検索.xlsm`）を返す。

```vba
Function BookName() As String
    Application.Volatile
    BookName = Application.ThisWorkbook.Name
End Function
```

---

### 6.3 SQL_Execution.bas

**処理概要**: ADO/ODBC DB接続・SQL実行の共通基盤（`ExGM_EB会社ｺｰﾄﾞ変換TB.xlsm` と同構造）。

`Open_oraconDB` では接続文字列をPublic変数 `mpDSN` から取得する（在庫/履歴で切り替え）。

| プロシージャ | 役割 |
|---|---|
| `Open_oraconDB` | `mpDSN` で指定したDSNに接続 |
| `SQL_Exe` | SQL文を実行しrsに格納（エラー時はStop） |
| `SQL_INSERT_UPDATE` | INSERT/UPDATE分岐処理 |
| `SQL_Delete` | DELETE実行 |
| `Disp_Sheet` | SQL結果をシートに貼り付け |
| `Set_Array` | SQL結果を2次元配列に格納 |

---

### 6.4 DBRead.bas

#### `条件設定(myKouM, 条件, myErr)`

**処理概要**: シートの入力値からSQL WHERE句と SELECT列リストを動的に組み立てる。

**処理フロー**:
1. 受付番号（`Uno`）が空の場合エラー
2. `受付番号範囲` で開始・終了番号を取得
3. WHERE句: `s.UNO >= '<uno1>' AND s.UNO <= '<uno2>'`
4. 会社コードが入力あり: `AND s.KAISYACD='<cd>'` を追加（4桁ゼロ埋め）
5. 製品コードが入力あり: `製品コード指定` で OR条件を生成して追加
6. 装置が2または3: `AND s.SYOUSO='<no>'` を追加
7. 追加フラグ: `含まない`→`'0'`、`のみ`→`'1'` で条件追加
8. 計算コード・指定パス・指定線量: 入力があれば条件追加
9. 表示列（myKouM）: 各 `H_xxx` フラグが「する」の列をカンマ区切りで連結

#### `受付番号範囲(myUno, Uno1, Uno2, myErr)`

**処理概要**: 受付番号の単一指定（10桁）・範囲指定（`-` または `,` 区切りの21桁）を解析して開始・終了番号を返す。

#### `製品コード指定(myOR, myErr)`

**処理概要**: カンマ区切りの複数製品コードを `z.SEHNCD='xxx' OR z.SEHNCD='yyy'` の形式に変換する（3桁ゼロ埋め）。

#### `条件表示(mySH)`

**処理概要**: 抽出シートのA1から、使用した検索条件（抽出日時・抽出種別・受付番号・各条件）を一覧表示する。

---

### 6.5 実行.bas

#### `検索実行()`

**処理概要**: 検索条件を組み立てDBにクエリを発行し、結果件数を確認後、新規ブックに抽出データを出力する。

**処理フロー**:
1. `Rireki` の値で接続先DSNとFROMテーブルを設定:
   - 在庫: `DSN=ricdb;UID=ric;PWD=t6101` / `zaiko z, syouj2 s`
   - 履歴: `DSN=ricdb;UID=rich;PWD=t6101` / `zaikor z, syoujr2 s`
2. `条件設定` でWHERE句と SELECT列リストを組み立て
3. `SELECT count(*)` で件数確認
4. 0件: 「データなし」で終了
5. 上限（`DataN`）超過: エラーで中止
6. 問題なし: 新規ブック作成確認ダイアログ
7. 本番SELECT実行 → 一時シートに `Disp_Sheet` で書き込み
8. `条件表示` で抽出条件をシート頭部に記載
9. 新規ブックを作成し、一時シートをコピー
10. 元ブックの一時シートを削除・条件シートに戻る

```sql
-- 件数確認
SELECT count(*) FROM zaiko z, syouj2 s
WHERE s.UNO>='<uno1>' AND s.UNO<='<uno2>' AND z.uno=s.uno [AND ...]

-- データ抽出
SELECT s.UNO 受付番号, [選択列...] FROM zaiko z, syouj2 s
WHERE ...
ORDER BY s.senkno, s.uno ASC
```

---

### 6.6 終了処理.bas

#### `終了処理Sub()`

**処理概要**: ブック数が1の場合はExcelを終了、複数の場合は当ブックのみを閉じる。

---

### 6.7 共通変数.bas

```vba
Public mpBook2 As String  '新規作成ブック名
Public mpBook1 As String  '当ブック名
```

---

## 7. DB 接続・外部連携

### ODBC 接続設定

| DSN 名 | UID | PWD | 用途 |
|---|---|---|---|
| `ricdb` | `ric` | `t6101` | 在庫データ取得（zaiko, syouj2） |
| `ricdb` | `rich` | `t6101` | 履歴データ取得（zaikor, syoujr2） |

### 参照テーブル一覧

| テーブル名 | 主な用途 | 主要カラム |
|---|---|---|
| `zaiko` | 在庫データ（ヘッダー部） | `uno`（受付番号）, `sehncd`（製品CD） |
| `syouj2` | 在庫照射詳細データ | `uno`, `kaisyacd`, `syouso`, `senkno`, `atusa`, `keisask`, `pass`, `jituno`, `siteisn`, `kagensn`, `jyougsn`, `sokutti`, `sokutsn`, `ondok`, `tonyutime`, `tuikaflg` |
| `zaikor` | 履歴データ（ヘッダー部） | `uno` |
| `syoujr2` | 履歴照射詳細データ | 同上 |

---

## 8. データフロー

```
【起動フロー】
  Workbook_Open
       ↓
  条件シートを選択 → シート保護設定
       ↓
  ThisB_Name セルに BookName() の結果を書き込み

【検索実行フロー】
  ユーザーが条件シートに入力
  （抽出先・受付番号・会社CD・製品CD・装置等）
       ↓
  「実行」ボタン押下 → 検索実行()
       ↓
  Rireki値でDSN/FROMを切り替え
       ↓
  条件設定(): WHERE句 + SELECT列リスト構築
       ↓
  SELECT count(*) でデータ件数確認
       ↓ （0件 → エラー終了 / 上限超 → 中止）
  ユーザー確認ダイアログ（新規BOOK作成）
       ↓
  DB(zaiko/zaikor + syouj2/syoujr2): SELECT実行
       ↓
  一時シートにデータ貼り付け + 条件表示
       ↓
  新規ブック作成 → 一時シートをコピー
       ↓
  元ブックから一時シート削除 → 条件シートに戻る
       ↓
  完了メッセージ表示
```

---

## 9. セキュリティ注意事項

| 種別 | キーワード | 内容 |
|---|---|---|
| AutoExec | `Workbook_Open` | ブックオープン時に自動実行。シート保護設定とファイル名取得を行う |
| AutoExec | `Workbook_BeforeClose` | ブック終了時に自動実行。保存ダイアログを非表示にする |
| Suspicious | `Open` | ADO接続の `oraconn.Open` が該当。ファイル操作ではなくDB接続 |
| Suspicious | `Call` | 各サブルーチン呼び出しで使用 |
| Suspicious | `Windows` | `Windows(mpBook1).Activate` でウィンドウ操作。ブック間データコピーのため |
| Suspicious | `Chr` | `Chr(13)` による改行文字生成（メッセージ整形） |
| Suspicious | Hex Strings / Base64 Strings | VBAバイナリ内のエンコード。実際のエンコード処理なし |
| 注意 | DB認証情報 | 接続文字列にUID/PWDをハードコード。在庫用: `ric/t6101`、履歴用: `rich/t6101` |
