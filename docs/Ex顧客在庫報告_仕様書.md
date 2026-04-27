# Ex顧客在庫報告 仕様書

> **ファイル種別**: .xlsm（マクロ付き）
> **用途**: 顧客（日機装・オリエンタル酵母・標準）向けの預かり在庫状況を照射管理DBから抽出し、在庫報告書を生成・記録するツール
> **VBA プロジェクトサイズ**: 約160KB
> **外部連携**: DSN=ricdb（Oracle DB）、記録先 `\\RNTSVR-FS\Sv_cup\営業部照射課\女子事務用\顧客在庫報告\報告データ`

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
Ex顧客在庫報告.xlsm
├── シート
│   ├── 標準          （標準顧客向け在庫報告書・入力・集計表示）
│   ├── 日機装        （日機装向け在庫報告書）
│   ├── オリエンタル   （オリエンタル酵母向け在庫報告書）
│   ├── ITパック      （オリエンタルITパック専用シート）
│   └── 抽出          （DB抽出データ格納・作業用シート）
│
├── VBA モジュール
│   ├── ThisWorkbook.cls  （Workbook_Open / Workbook_BeforeClose）
│   ├── Sheet1〜6.cls     （空）
│   ├── 在庫抽出設定.bas  （DB接続・共通データ抽出サブルーチン）
│   ├── 日機装.bas        （日機装データ抽出・表示）
│   ├── オリエンタル.bas   （オリエンタル酵母データ抽出・IT選別）
│   ├── 標準書.bas        （標準顧客向けデータ抽出・報告書生成）
│   ├── 画面初期化.bas    （全シートの入力値・表示クリア）
│   └── 終了処理.bas      （ブック終了処理）
│
└── 外部ファイル連携
    └── 記録先フォルダ: \\RNTSVR-FS\Sv_cup\営業部照射課\...（报告書保存）
```

---

## 2. シート詳細

### 2.1 標準

**目的**: 標準顧客向けの在庫報告書作成シート。会社コード・在庫基準日を入力してDBから在庫情報を抽出し、報告書フォーマットに表示する。

#### レイアウト構造

| セル範囲 | 名前付き範囲 | 内容 |
|---|---|---|
| D1 | — | ラベル「会社ｺｰﾄﾞ」 |
| E1 | `KaisyaNo` | 会社コード入力欄 |
| G1 | — | ラベル「当日出荷分」 |
| H1 | `HSyuF` | 出荷フラグ（含まない/含む） |
| O1 | — | 「含む」ラジオ選択肢 |
| P1 | `HDay` | 在庫基準日 |
| Q1 | `MaxH` | 最大行数管理値 |
| G2 | — | ラベル「当日入荷分」 |
| H2 | `Tnyu` | 入荷フラグ（含む/含まない） |
| D3 | — | ラベル「記録先ﾌｫﾙﾀﾞｰ名」 |
| E3 | `FNameH` | 記録先フォルダパス |
| E12 | `FilNameH` | 記録ファイル名 |
| L13 | `Hakkoubi` | 発行日 |
| C16 | `KaisyaName` | 顧客社名（マクロが設定） |
| C18 | `NenGappi` | 年月日表示 |
| C19 | `Fukumu` | 出荷/入荷含む・含まない説明文 |
| E14:K14 | — | タイトル「在　庫　報　告　書」（結合） |
| C20:L20 | — | テーブルヘッダー行（No/受付番号/入荷日/製品ｺｰﾄﾞ/製品名/指定線量/入荷数/納品済数/預り在庫数/ロットNo） |
| C21:L200 | `ZaikoHyou` | 在庫一覧表示領域（VBAが書き込み） |

（行番号 201〜220 は空白・フッター領域）

---

### 2.2 日機装

**目的**: 日機装専用の在庫報告書シート。入荷日・受付番号・品名・ケース数・本数・ロット番号を表示する。

#### レイアウト構造

| セル範囲 | 名前付き範囲 | 内容 |
|---|---|---|
| E1 | `KaisyaCD` | 日機装会社コード |
| F1:G1 | — | 結合セル（ファイル名エリア） |
| E2 | `SyukkaF` | 出荷フラグ |
| G2:H2 | `FilNameN` | ファイル名（結合） |
| H1 | `FNameN` | 保存先フォルダ |
| D5 | `NDay` | 在庫基準日 |
| R3 | `Hizuke` | 日付 |
| Q1:Q2 | — | 結合セル |
| B8:H199 | `Nikki` | データ表示領域（VBAが書き込み） |

（行番号 7 はヘッダー行: No/入荷日/受付番号/品名/ケース数/本数/ロットNo）

---

### 2.3 オリエンタル

**目的**: オリエンタル酵母向けの在庫報告書シート。IT選別機能付き。

#### レイアウト構造

| セル範囲 | 名前付き範囲 | 内容 |
|---|---|---|
| B1:C1 | — | 結合セル |
| D1 | `KaisyaCD747` | 会社コード（0747） |
| D2 | `KaisyaCDO` | 会社コード（0226） |
| G1:N1 | `FNameO` | ファイル名 |
| G2:N2 | `FilNameO` | 保存先フォルダ |
| T1 | `ITPack` | ITパック開始行 |
| T2 | `ITend` | ITパック終了行 |
| W1 | `SyuuDay` | 収受日 |
| U1:U2 | — | 結合セル |
| C9:E9 | `HizukeO` | 在庫日付（結合） |
| C11:J11 | `Kpumoku` | 列ヘッダー行 |
| B12:K204 | `Orient` | データ表示領域 |
| E2 | `SyukkaFO` | 出荷フラグ |

---

### 2.4 ITパック

**目的**: オリエンタル酵母のITパック製品を抽出・表示する専用シート。

#### レイアウト構造

| セル範囲 | 名前付き範囲 | 内容 |
|---|---|---|
| B1:C1 | — | 結合セル |
| G1:N1 | — | 結合セル（タイトル） |
| G2:N2 | — | 結合セル |
| C9:E9 | `ITDate` | 在庫日付（結合） |
| C12:J204 | `OrientIT` | ITパックデータ表示領域 |

---

### 2.5 抽出

**目的**: DB から取得したデータを一時格納する作業用シート（非表示推奨）。最大6,687行×39列。

#### DB読み込み列

| セル範囲 | 名前付き範囲 | 内容 |
|---|---|---|
| B6:C1002 | `SeihinnTB` | 製品マスタ（製品コード・製品名） |
| E6:N205 | `Zaiko` | 在庫テーブルデータ |
| T6:AC205 | `ZaikoR` | 在庫履歴テーブルデータ |
| AI6:AJ1005 | `SyukkaZ` | 在庫用出荷数集計 |
| AL6:AM100000 | `SyukkaR` | 在庫履歴用出荷数集計 |
| I2 | `KaiName` | 顧客名 |
| I3 | `ZaiDatasuu` | 在庫データ件数 |
| X3 | `RirwkiDataSuu` | 履歴データ件数 |
| AC2 | `UnoMin` | 受付番号最小値 |
| AC3 | `UnoMax` | 受付番号最大値 |
| AI1:AJ3 | — | 結合セル（SyukkaZ ヘッダー） |
| AL1:AM3 | — | 結合セル（SyukkaR ヘッダー） |

（行番号 4〜5 は空。行番号 206〜6687 は空行）

---

## 3. 名前付き範囲一覧

| 名前 | 参照先 | 説明 |
|---|---|---|
| `FilNameH` | 標準!$E$12 | 標準報告書のファイル名 |
| `FilNameN` | 日機装!$G$2 | 日機装報告書のファイル名 |
| `FilNameO` | オリエンタル!$G$2 | オリエンタル報告書のファイル名 |
| `FNameH` | 標準!$E$3 | 標準ファイル保存先フォルダパス |
| `FNameN` | 日機装!$H$1 | 日機装ファイル保存先フォルダパス |
| `FNameO` | オリエンタル!$G$1 | オリエンタルファイル保存先フォルダパス |
| `Fukumu` | 標準!$C$19 | 出荷/入荷「含む・含まない」の説明文 |
| `Hakkoubi` | 標準!$L$13 | 報告書発行日 |
| `HDay` | 標準!$P$1 | 標準用在庫基準日 |
| `Hizuke` | 日機装!$R$3 | 日機装用在庫基準日 |
| `HizukeO` | オリエンタル!$C$9 | オリエンタル用在庫日付 |
| `HSyuF` | 標準!$H$1 | 出荷フラグ（含む/含まない） |
| `ITDate` | ITパック!$C$9 | ITパック用在庫日付 |
| `ITend` | オリエンタル!$T$2 | ITパック選択終了行番号 |
| `ITPack` | オリエンタル!$T$1 | ITパック選択開始行番号 |
| `KaiName` | 抽出!$I$2 | DB から取得した顧客名 |
| `KaisyaCD` | 日機装!$E$1 | 日機装の会社コード |
| `KaisyaCD747` | オリエンタル!$D$1 | オリエンタル酵母会社コード（0747） |
| `KaisyaCDO` | オリエンタル!$D$2 | オリエンタル酵母会社コード（0226） |
| `KaisyaName` | 標準!$C$16 | 表示用顧客社名 |
| `KaisyaNo` | 標準!$E$1 | 標準用会社コード入力欄 |
| `Kpumoku` | オリエンタル!$C$11:$J$11 | オリエンタル列ヘッダー |
| `MaxH` | 標準!$Q$1 | 最大行管理値 |
| `NDay` | 日機装!$D$5 | 日機装用在庫基準日 |
| `NenGappi` | 標準!$C$18 | 在庫基準日表示 |
| `Nikki` | 日機装!$B$8:$H$199 | 日機装データ表示領域 |
| `Orient` | オリエンタル!$B$12:$K$204 | オリエンタルデータ表示領域 |
| `OrientIT` | ITパック!$C$12:$J$204 | ITパックデータ表示領域 |
| `RirwkiDataSuu` | 抽出!$X$3 | 在庫履歴データ件数 |
| `SeihinnTB` | 抽出!$B$6:$C$1002 | 製品マスタ参照テーブル |
| `SyukkaF` | 日機装!$E$2 | 日機装出荷フラグ |
| `SyukkaFO` | オリエンタル!$E$2 | オリエンタル出荷フラグ |
| `SyukkaR` | 抽出!$AL$6:$AM$100000 | 在庫履歴用出荷数集計テーブル |
| `SyukkaZ` | 抽出!$AI$6:$AJ$1005 | 在庫用出荷数集計テーブル |
| `SyuuDay` | オリエンタル!$W$1 | 収受日 |
| `Tnyu` | 標準!$H$2 | 入荷フラグ（含む/含まない） |
| `UnoMax` | 抽出!$AC$3 | 受付番号最大値（SQL WHERE 句に使用） |
| `UnoMin` | 抽出!$AC$2 | 受付番号最小値（SQL WHERE 句に使用） |
| `ZaiDatasuu` | 抽出!$I$3 | 在庫テーブルのデータ件数 |
| `Zaiko` | 抽出!$E$6:$N$205 | 在庫テーブルデータ格納領域 |
| `ZaikoHyou` | 標準!$C$21:$L$200 | 標準シートの在庫一覧表示領域 |
| `ZaikoR` | 抽出!$T$6:$AC$205 | 在庫履歴テーブルデータ格納領域 |

---

## 4. 数式一覧

### 抽出シート（代表的な数式パターン）

| セル | 数式 | 説明 |
|---|---|---|
| O6〜O205 | `=IF(E6="","",I6-R6)` | 在庫数算出（入荷数 − 現在庫数） |
| P6〜P205 | `=IF(G6="","",VLOOKUP(G6,SeihinnTB,2,FALSE))` | 製品コードから製品名をマスタ参照 |
| Q6〜Q205 | `=IF(ISERROR(VLOOKUP(E6,SyukkaZ,2,FALSE)),0,VLOOKUP(E6,SyukkaZ,2,FALSE))` | 在庫に加える出荷数（指定日以降の出荷分） |
| R6〜R205 | `=IF(ISERROR(...),...,I6-M6+Q6)` | 最終在庫数（入荷数 − 出荷済数 + 指定日以降出荷数） |
| AD6〜AD205 | `=X6-AG6` | 在庫履歴の在庫数 |
| AE6〜AE205 | `=IF(V6="","",VLOOKUP(V6,SeihinnTB,2,FALSE))` | 在庫履歴の製品名参照 |
| AF6〜AF205 | `=IF(ISERROR(VLOOKUP(T6,SyukkaR,2,FALSE)),0,VLOOKUP(T6,SyukkaR,2,FALSE))` | 在庫履歴用出荷数（指定日以前の出荷分） |
| AG6〜AG205 | `=X6-AF6` | 在庫履歴の最終在庫数 |

---

## 5. ボタン・マクロ対応

| シート | ボタンラベル | 割り当てマクロ | 動作概要 |
|---|---|---|---|
| 標準 | 集計 | `標準データ抽出` | DBから標準顧客の在庫・履歴を抽出し標準シートに表示 |
| 標準 | 記録 | `標準報告書記録処理` | 標準シートの内容をファイルに保存 |
| 日機装 | 集計 | `日機装抽出` | DBから日機装の在庫・履歴を抽出し日機装シートに表示 |
| 日機装 | 記録 | `日機装報告書記録処理` | 日機装シートの内容をファイルに保存 |
| オリエンタル | 集計 | `オリエンタル酵母` | DBからオリエンタル酵母の在庫を抽出し表示 |
| オリエンタル | IT選別 | `IT抽出` | ITパック対象品を選別してITパックシートに転記 |
| オリエンタル | 記録 | `オリエンタル報告書記録` | オリエンタルシートの内容をファイルに保存 |

---

## 6. VBA モジュール仕様

### 6.1 ThisWorkbook

#### `Workbook_Open()`

**処理概要**: ブック起動時に初期値を設定し、画面クリアを実行する。

**処理フロー**:
1. `HSyuF`（出荷フラグ）を「含まない」に設定
2. `Tnyu`（入荷フラグ）を「含む」に設定
3. `画面クリア` を呼び出してすべての入力・表示をリセット

```vba
Private Sub Workbook_Open()
    Range("HSyuF") = "含まない"
    Range("Tnyu") = "含む"
    Call 画面クリア
End Sub
```

#### `Workbook_BeforeClose(Cancel As Boolean)`

**処理概要**: ブックを閉じる前に「保存済み」フラグを立て、保存ダイアログを抑止する。

```vba
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Application.DisplayAlerts = False
    ActiveWorkbook.Saved = True
End Sub
```

---

### 6.2 在庫抽出設定

#### `データ抽出(mySyuF, myNyu, myKaisya, mydate)`

**処理概要**: ODBCを経由してOracle DBに接続し、指定顧客の在庫・在庫履歴・出荷履歴・製品マスタを「抽出」シートに格納する共通サブルーチン。

**処理フロー**:
1. 接続文字列を `DSN=ricdb;UID=ric;PWD=t6101` に設定
2. 製品マスタ（`sehmst`）を抽出 → `SeihinnTB` に格納
3. 在庫テーブル（`zaiko`）から対象顧客の在庫を抽出 → `Zaiko` に格納
4. 顧客名（`tokumst`）を取得 → `KaiName` に格納
5. 在庫履歴テーブル（`zaikor`）を抽出 → `ZaikoR` に格納（`DSN=ricdb;UID=rich` を使用）
6. 出荷履歴（`syukar`）から在庫加算用出荷数を抽出 → `SyukkaZ` に格納
7. 出荷履歴（`syukar`）から在庫履歴用出荷数を抽出 → `SyukkaR` に格納

```vba
Sub データ抽出(ByVal mySyuF, ByVal myNyu, ByVal myKaisya, ByVal mydate)
    Dim myDB As String
    myDB = "DSN=ricdb;UID=ric;PWD=t6101"
    ' ... 製品マスタ抽出
    ' ... 在庫抽出
    ' ... 在庫履歴抽出（DSN=ricdb;UID=rich;PWD=t6101）
    ' ... 出荷履歴抽出（在庫用・履歴用）
End Sub
```

#### `データ抽出サブ(myDB, strSql, mySH, myRow, myColumn, myF)`

**処理概要**: ADO+ODBCを使用してSQLを実行し、結果をワークシートの指定セルに貼り付ける汎用サブルーチン。

**処理フロー**:
1. `ADODB.Connection` で接続文字列を開く
2. `oraconn.Execute(strSql)` でレコードセット取得
3. `CopyFromRecordset` でワークシートに一括貼り付け
4. エラー時は `MsgBox` でエラー内容を表示

```vba
Sub データ抽出サブ(myDB As String, strSql As String, mySH As String, _
                   myRow As Single, myColumn As Single, myF As Integer)
    Dim oraconn As New ADODB.Connection
    Dim rs As ADODB.Recordset
    oraconn.ConnectionString = myDB
    oraconn.Open
    Set rs = oraconn.Execute(strSql)
    Worksheets(mySH).Cells(myRow + myF, myColumn).CopyFromRecordset rs
    rs.Close
End Sub
```

---

### 6.3 標準書

#### `標準データ抽出()`

**処理概要**: 標準顧客向けに在庫データを抽出し、標準シートの在庫一覧表に表示・ソートし印刷範囲を設定する。

**処理フロー**:
1. 会社コード・在庫基準日の入力チェック
2. 出荷/入荷フラグに応じて `Fukumu`（注記文言）を生成
3. `データ抽出` を呼び出して抽出シートにDBデータを格納
4. 在庫・履歴データを配列 `myData()` に格納
5. 標準シートの `ZaikoHyou` 領域に書き込み（受付番号でソート）
6. 合計行を追加
7. 印刷範囲を設定
8. 完了メッセージを表示

---

### 6.4 日機装

#### `日機装抽出()`

**処理概要**: 日機装専用の在庫データを抽出し、日機装シートに表示・ソートし印刷範囲を設定する。

**処理フロー**:
1. 日付入力チェック（`NDay`）
2. 出荷フラグ（`SyukkaF`）・顧客コード・基準日を取得
3. `データ抽出` を呼び出して抽出シートにDBデータを格納
4. 在庫・履歴データを配列 `myNikki()` に格納（入荷日/品名/受付番号/ケース数/本数/ロット番号）
5. 日機装シートの `Nikki` 領域に書き込み（受付番号でソート）
6. 印刷範囲を設定
7. 完了メッセージを表示

---

### 6.5 オリエンタル

#### `オリエンタル酵母()`

**処理概要**: オリエンタル酵母（会社コード0747・0226）の在庫データを抽出し、ITパック入力規則を設定する。

**処理フロー**:
1. 在庫日付（`HizukeO`）の入力チェック
2. 会社コード0747の在庫を抽出（`オリエンタル抽出` を呼び出し）
3. 会社コード0226の在庫を抽出
4. ITパック対象行に「IT」を選択するドロップダウンリストを設定

#### `IT抽出()`

**処理概要**: オリエンタルシートで「IT」を選択された行をITパックシートへ転記し、残りはオリエンタルシートに残す。

**処理フロー**:
1. `ITPack`〜`ITend` の行範囲でITパック対象を振り分け
2. 「IT」選択行を `myIT(1)` に、それ以外を `myIT(0)` に格納
3. オリエンタルシートをクリアしてIT以外を再表示
4. ITパックシートに転記
5. 両シートの印刷範囲を設定

---

### 6.6 画面初期化

#### `画面クリア()`

**処理概要**: すべてのシートの入力値・表示データを一括クリアする。

```vba
Sub 画面クリア()
    Range("Zaiko") = ""
    Range("ZaikoHyou") = ""
    Range("SeihinnTB") = ""
    Range("Orient") = ""
    Range("Nikki") = ""
    ' ... 他の名前付き範囲もクリア
    Sheets("標準").Select
    Range("NenGappi") = ""
    Range("Hakkoubi") = ""
    Range("FilNameH") = ""
End Sub
```

---

### 6.7 終了処理

#### `終了処理Sub()`

**処理概要**: ブックを閉じる。他にブックが開いていなければ Excel ごと終了する。

```vba
Sub 終了処理Sub()
    Application.DisplayAlerts = False
    If Application.Workbooks.Count = 1 Then
        Application.Quit
    Else
        ThisWorkbook.Close
    End If
End Sub
```

---

## 7. DB 接続・外部連携

### ODBC 接続設定

| DSN 名 | UID | PWD | 用途 |
|---|---|---|---|
| `ricdb` | `ric` | `t6101` | 製品マスタ・在庫テーブル取得 |
| `ricdb` | `rich` | `t6101` | 在庫履歴・出荷履歴取得 |

### 参照テーブル一覧

| テーブル名 | 主な用途 | 主要カラム |
|---|---|---|
| `sehmst` | 製品マスタ | `sehncd`（製品コード）, `seiname`（製品名）, `kaisyacd` |
| `zaiko` | 在庫 | `uno`（受付番号）, `kainame`, `sehncd`, `nyukabi`（入荷日）, `nyukasu`（入荷数）, `siteisn`（指定線量）, `pass`, `syouso`, `syukasu`（出荷済数）, `rotno2`（ロット番号）, `kaisyacd` |
| `zaikor` | 在庫履歴 | zaiko と同構造 + `syukabi`（出荷日） |
| `tokumst` | 顧客マスタ | `coname`（顧客名）, `kaisyacd` |
| `syukar` | 出荷履歴 | `uno`, `kaisyacd`, `syudate`（出荷日）, `syukasu`（出荷数） |

### 主要 SQL 文

```sql
-- 製品マスタ抽出
SELECT sehncd*1, trim(seiname)
FROM sehmst
WHERE kaisyacd = '<会社コード>'
ORDER BY sehncd

-- 在庫抽出（入荷日含む場合）
SELECT uno*1, trim(kainame), sehncd*1, nyukabi, nyukasu*1,
       siteisn, pass, syouso, syukasu*1, trim(rotno2)
FROM zaiko
WHERE nyukabi <= '<基準日>' AND kaisyacd = '<会社コード>'
ORDER BY uno

-- 在庫履歴抽出（出荷日含む・入荷日含む場合）
SELECT uno*1, trim(kainame), sehncd*1, nyukabi, nyukasu*1,
       siteisn, pass, syouso, syukasu*1, trim(rotno2)
FROM zaikor
WHERE syukabi >= '<基準日>'
  AND nyukabi <= '<基準日>'
  AND kaisyacd = '<会社コード>'
ORDER BY uno

-- 出荷履歴（在庫に加算する出荷数）
SELECT uno*1, sum(syukasu)
FROM syukar
WHERE uno >= '<UnoMin>' AND uno <= '<UnoMax>'
  AND kaisyacd = '<会社コード>'
  AND syudate >= '<基準日>'
GROUP BY uno
ORDER BY uno DESC

-- 顧客名取得
SELECT coname FROM tokumst WHERE kaisyacd = '<会社コード>'
```

### 外部ファイル連携

| ファイル名 | 用途 |
|---|---|
| `\\RNTSVR-FS\Sv_cup\営業部照射課\女子事務用\顧客在庫報告\報告データ\` | 在庫報告書（新規ブック）の保存先フォルダ |

---

## 8. データフロー

```
【標準顧客向け在庫報告フロー】
  標準シート: 会社コード・在庫基準日 入力
        ↓ [集計]ボタン → 標準データ抽出()
  在庫抽出設定.データ抽出() → ODBC(ricdb/ricdbh)
        ↓
  DB(sehmst): 製品マスタ → 抽出!SeihinnTB
  DB(zaiko): 在庫 → 抽出!Zaiko
  DB(zaikor): 在庫履歴 → 抽出!ZaikoR
  DB(syukar): 出荷数集計 → 抽出!SyukkaZ, SyukkaR
        ↓
  抽出シートの数式(O列〜R列/AD列〜AG列)で最終在庫数を計算
        ↓
  標準シート ZaikoHyou 領域に表示・受付番号でソート
        ↓ [記録]ボタン → 標準報告書記録処理()
  新規ブックに標準シートをコピー → 指定フォルダに名前を付けて保存

【日機装向け在庫報告フロー】
  日機装シート: 基準日・出荷フラグ 入力
        ↓ [集計]ボタン → 日機装抽出()
  （同様に DB 抽出 → Nikki 領域に表示）
        ↓ [記録]ボタン → 日機装報告書記録処理()

【オリエンタル酵母向け在庫報告フロー】
  オリエンタルシート: 在庫日付・会社コード 入力
        ↓ [集計]ボタン → オリエンタル酵母()
  会社コード0747 → Orient 上部に表示
  会社コード0226 → Orient 下部に表示、ITドロップダウン設定
        ↓ [IT選別]ボタン → IT抽出()
  「IT」選択行 → ITパックシートに転記
  非IT行 → オリエンタルシートに再整理
        ↓ [記録]ボタン → オリエンタル報告書記録()
```

---

## 9. セキュリティ注意事項

| 種別 | キーワード | 内容 |
|---|---|---|
| AutoExec | `Workbook_Open` | ブックを開いた際に自動的に `画面クリア` が実行される |
| AutoExec | `Workbook_BeforeClose` | ブックを閉じる際に保存ダイアログを抑止する |
| Suspicious | `Open` | ファイルを開く操作が含まれる（報告書保存時に新規ブックを開く） |
| Suspicious | `Call` | マクロ間の呼び出し（DLL呼び出しではなく内部モジュール呼び出し） |
| Suspicious | `MkDir` | フォルダが存在しない場合に自動作成する処理が含まれる |
| Suspicious | `ActiveWorkbook.SaveAs` | 報告書を指定フォルダに自動保存する（正常な業務処理） |
| Suspicious | `Windows` | ウィンドウ切り替え処理が含まれる（複数ブック操作時の正常な動作） |
| Suspicious | `Chr` | 改行文字（Chr(13)）を使用したメッセージ生成（難読化ではない） |
| Suspicious | Hex Strings | バイナリデータに16進エンコード文字列が検出された（xlsmの正常な構造） |
| Suspicious | Base64 Strings | Base64エンコード文字列が検出された（xlsmの正常な構造） |

> **注意**: DB接続文字列（`DSN=ricdb;UID=ric;PWD=t6101`）がVBAコード内にハードコーディングされています。パスワードが平文で埋め込まれているため、ファイルの配布・共有時には注意が必要です。
