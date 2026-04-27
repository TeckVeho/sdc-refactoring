# Ex単価検索 仕様書

> **ファイル種別**: .xlsm（マクロ付き）
> **用途**: 受付番号・会社コード・製品コードから単価を検索・表示し、DBへ単価および価格表ファイル情報を登録する
> **VBA プロジェクトサイズ**: 5モジュール（Sheet1.cls / ThisWorkbook.cls / SQL_Execution.bas / データベースR.bas / 表示.bas）
> **外部連携ファイル**: \\RNTSVR-FS\Sv_cup\営業部照射課\価格表\（フォルダ・ファイル直接参照）

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
Ex単価検索.xlsm
├── シート
│   └── Main（単価検索メイン画面）
├── VBA モジュール
│   ├── Sheet1.cls         – Worksheet イベント（Change / BeforeDoubleClick）
│   ├── ThisWorkbook.cls   – ブックイベント（Open / Activate / BeforeClose）
│   ├── SQL_Execution.bas  – DB 接続・SQL 実行共通ルーチン
│   ├── データベースR.bas  – DB 読み込み処理（受付番号・会社・製品・単価）
│   └── 表示.bas           – 表示処理（価格表ファイル・受付番号・会社・製品）
├── ボタン（5個）
│   ├── 終了
│   ├── ｺｰﾄﾞ検索
│   ├── ファイル登録
│   ├── 単価登録
│   └── 表示
└── 外部リンク: なし（VBA 内 UNC パス参照のみ）
```

---

## 2. シート詳細

### 2.1 Main

**目的**: 受付番号・会社コード・製品コードを入力して単価情報を検索・表示し、単価登録や価格表ファイルの登録・表示を行うメイン画面。

#### レイアウト構造

| セル範囲 | 名前付き範囲 | 内容 |
|---|---|---|
| B1 | — | バージョン識別子（`Ex190924`） |
| N2:O2 | `Kyou` | 本日日付（TODAY()数式） |
| B3 | `SyainName`（C3） | タイトル「単価検索」 |
| C3:D3 | `SyainName` | 社員名入力欄 |
| G3 | — | ラベル「ﾌｫﾙﾀﾞｰ名」 |
| I3 | — | ラベル「ﾌｧｲﾙ名」 |
| J3:O3 | `FailName` | ファイル名入力欄 |
| B4 | — | ヘッダー「受付番号」 |
| C4 | — | ヘッダー「会社ｺｰﾄﾞ」 |
| D4 | — | ヘッダー「製品ｺｰﾄﾞ」 |
| E4 | — | ヘッダー「単価」 |
| F4 | — | ヘッダー「単位」 |
| G4 | — | ヘッダー「価格登録日」 |
| H4 | — | ヘッダー「線量」 |
| I4 | — | ヘッダー「装置」 |
| J4 | — | ヘッダー「ﾊﾟｽ数」 |
| K4 | — | ヘッダー「幅cm」 |
| L4 | — | ヘッダー「高cm」 |
| M4 | — | ヘッダー「長cm」 |
| N4 | — | ヘッダー「重量」 |
| O4 | — | ヘッダー「入り数」 |
| T4 | `FaileDispName` | ヘッダー「ﾌｧｲﾙ表示ｵﾌﾞｼﾞｪｸﾄ名」 |
| U4 | `FolderName` | ヘッダー「フォルダ名」 |
| V4 | — | ヘッダー「単価」（区分ラベル列） |
| W4 | `TanniList` | ヘッダー「単位」（単位一覧） |
| X4 | `Ryaku` | ヘッダー「検索方法」 |
| B5 | `Uno` | 受付番号入力欄 |
| C5 | `KaiCd` | 会社コード入力欄 |
| D5 | `SeiCd` | 製品コード入力欄 |
| E5 | `Tannka` / `SeihinD` | 単価表示欄 |
| F5 | `Tanni` | 単位表示欄 |
| G5 | `TourokuBi` | 価格登録日表示欄 |
| H5 | `Dose` | 線量表示欄（`H8` も参照） |
| I5 | `Souti` | 装置表示欄 |
| J5 | `Pass` | パス数表示欄 |
| K5 | `Haba` | 幅cm表示欄 |
| L5 | `Takasa` | 高cm表示欄 |
| M5 | `Nagasa` | 長cm表示欄 |
| N5 | `JyuuRyou` | 重量表示欄 |
| O5 | `Irisuu` | 入り数表示欄 |
| U5:U14 | `FolderName` | フォルダ名一覧（マスタ） |
| V5 | — | 「登録単価」区分ラベル |
| W5 | — | 単位選択肢「箱」 |
| X5:X6 | `Ryaku` | 検索方法（先頭/含む） |
| B6 | — | ラベル「装置　　」 |
| C6:G6 | `KaiName` | 会社名表示欄 |
| H6:N6 | `SeiName` | 製品名表示欄 |
| T6 | — | ラベル「価格表記録ﾌｫﾙﾀﾞ」 |
| V6 | — | 「確定単価」区分ラベル |
| W6 | — | 単位選択肢「PL」 |
| B7 | `Syouso` | 装置種別（1,2,3号機 / EB） |
| C7:G7 | `KaiBikou` | 会社備考表示欄 |
| H7:N7 | `SeiBikou` | 製品備考表示欄 |
| T7 | `PathFolder` | 価格表記録フォルダパス（`d:aaa`） |
| V7 | — | 「単価未定」区分ラベル |
| W7 | — | 単位選択肢「m」 |
| B8 | — | ラベル「ファイル表示」 |
| C8 | `Hyouji` | ファイル表示設定（しない） |
| H8 | `FolDer` | フォルダ入力欄（検索開始入荷日計算に使用） |
| K8 | `Ryakusyou` | コード入力欄 |
| L8 | `Ryakusyou` | 略称フラグ |
| M8 | `MojiLike` | 文字検索方法（先頭） |
| T8 | — | ヘッダー「装置」（マスタ） |
| U8 | — | フォルダ名マスタ（たちつてと） |
| V8 | — | 「単価」区分ラベル |
| W8 | — | 単位選択肢「kg」 |
| T9 | — | 装置マスタ「EB」 |
| W9 | — | 単位選択肢「IB」 |
| T10 | — | 装置マスタ「1,2,3号機」 |
| W10 | — | 単位選択肢「ｻｲｸﾙ」 |
| K9:O2000 | `RyakuTB` | 略称テーブル領域 |
| T13 | — | ラベル「検索開始入荷日」 |
| T14 | — | 検索開始入荷日（TODAY()-H8 数式） |
| T15 | `NyukabiBef` | 入荷日文字列（YYYYMMDD 形式） |
| A1 | `Debugf` | デバッグフラグ（9の時処理スキップ） |

（行番号 16〜2000 は主にデータ領域・略称テーブル。B5〜O5 がメイン入出力行）

---

## 3. 名前付き範囲一覧

| 名前 | 参照先 | 説明 |
|---|---|---|
| `Debugf` | Main!$A$1 | デバッグフラグ（9でVBA処理スキップ） |
| `Dose` | Main!$H$5 | 線量表示欄 |
| `FaileDispName` | Main!$T$5 | ファイル表示オブジェクト名 |
| `FailName` | Main!$J$3 | ファイル名入力欄 |
| `FolDer` | Main!$H$3 | フォルダ名入力欄（フォルダパス計算に利用） |
| `FolderName` | Main!$U$5:$U$14 | フォルダ名一覧マスタ |
| `Haba` | Main!$K$5 | 幅cm |
| `Hyouji` | Main!$C$8 | ファイル自動表示設定（しない/する） |
| `Irisuu` | Main!$O$5 | 入り数 |
| `JyuuRyou` | Main!$N$5 | 重量 |
| `KaiBikou` | Main!$C$7 | 会社備考 |
| `KaiCd` | Main!$C$5 | 会社コード入力欄 |
| `KaiName` | Main!$C$6 | 会社名表示欄 |
| `Kyou` | Main!$N$2 | 本日日付 |
| `MojiLike` | Main!$M$8 | 文字検索方法（先頭/含む） |
| `Nagasa` | Main!$M$5 | 長さcm |
| `NyukabiBef` | Main!$T$15 | 検索開始入荷日（YYYYMMDD形式） |
| `Pass` | Main!$J$5 | パス数 |
| `PathFolder` | Main!$T$7 | 価格表記録フォルダルートパス |
| `Ryaku` | Main!$X$5:$X$6 | 検索方法（先頭/含む） |
| `Ryakusyou` | Main!$L$8 | 略称表示フラグ |
| `RyakuTB` | Main!$K$9:$O$2000 | 略称テーブル |
| `SeiBikou` | Main!$H$7 | 製品備考 |
| `SeiCd` | Main!$D$5 | 製品コード入力欄 |
| `SeihinD` | Main!$E$5:$O$5 | 製品データ行（単価〜入り数） |
| `SeiName` | Main!$H$6 | 製品名表示欄 |
| `Souti` | Main!$I$5 | 装置 |
| `SyainName` | Main!$C$3 | 登録者（社員名） |
| `Syouso` | Main!$B$7 | 装置種別（1,2,3号機 / EB） |
| `Takasa` | Main!$L$5 | 高さcm |
| `Tanni` | Main!$F$5 | 単位 |
| `TanniList` | Main!$W$5:$W$11 | 単位選択肢一覧 |
| `Tannka` | Main!$E$5 | 単価 |
| `TokuD` | Main!$B$6:$O$7 | 得意先データ行（会社名〜製品備考） |
| `TourokuBi` | Main!$G$5 | 価格登録日 |
| `Uno` | Main!$B$5 | 受付番号入力欄 |

---

## 4. 数式一覧

### Main シート

| セル | 数式 | 説明 |
|---|---|---|
| N2 | `=TODAY()` | 本日日付を表示 |
| T14 | `=TODAY()-H8` | 検索開始入荷日（本日 − H8の日数） |
| T15 | `=(YEAR(T14)&RIGHT("00"&MONTH(T14),2)&RIGHT("00"&DAY(T14),2))` | T14をYYYYMMDD形式の文字列に変換しSQL WHERE句に使用 |

---

## 5. ボタン・マクロ対応

| シート | ボタンラベル | 割り当てマクロ | 動作概要 |
|---|---|---|---|
| Main | 終了 | `Bookを閉じる` | ブックを保存せずに閉じる |
| Main | ｺｰﾄﾞ検索 | `会社コード表示` | 会社コード一覧を表示して選択 |
| Main | ファイル登録 | `単価_フォルダーファイル名更新` | フォルダ名・ファイル名をDBの ExSeihinJ に登録・更新 |
| Main | 単価登録 | `単価登録` | 入力中の単価・単位・備考を ExSeihinZ に INSERT/UPDATE |
| Main | 表示 | `価格表の表示` | UNC パスの価格表ファイル（Word/Excel）を WScript.Shell で開く |

---

## 6. VBA モジュール仕様

### 6.1 Sheet1.cls

#### `Worksheet_BeforeDoubleClick(Target, Cancel)`

**処理概要**: セル上でダブルクリックしたとき、K列（列11）8行目以降のセルをクリックした場合に会社コードをKaiCdセルにコピーする。

**処理フロー**:
1. ダブルクリックされたセルが 9行目以降 かつ K列（列11）かを確認
2. 条件に合致する場合、クリックセルの値を `Range("KaiCd")` にセット

```vba
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    Dim myKaiCD As String
    With Target
        If .Row > 8 And .Column = 11 Then
            Range("KaiCd") = Cells(.Row, .Column)
        End If
    End With
End Sub
```

#### `Worksheet_Change(Target)`

**処理概要**: セル値変更イベント。変更されたセルの行・列に応じて受付番号・会社コード・製品コードの連動表示処理を呼び出す。

**処理フロー**:
1. `Debugf = 9` なら処理スキップ
2. 変更セルが B5（受付番号）→ `受付番号表示` 呼び出し
3. 変更セルが C5（会社コード）→ `セル書式設定` + `会社データ表示` 呼び出し
4. 変更セルが D5（製品コード）かつ会社コード入力済み → `製品データ表示` 呼び出し
5. 変更セルが E5（単価）→ 数値書式 `\#,##0.00;\-#,##0.00` を適用

```vba
Private Sub Worksheet_Change(ByVal Target As Range)
    If Range("Debugf") = 9 Then Exit Sub
    Dim s As Integer, myF As Variant
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    With Target
        If .Row = 5 And .Column = 2 Then
            Call 受付番号表示
            Range("Uno").Select
        ElseIf .Row = 5 And .Column = 3 Then
            Call セル書式設定
            Call 会社データ表示
        ElseIf .Row = 5 And .Column = 4 Then
            If Range("KaiCd") <> "" Then Call 製品データ表示
            Range("SeiCd").Select
        ElseIf .Row = 5 And .Column = 5 Then
            Selection.NumberFormatLocal = "\#,##0.00;\-#,##0.00"
        End If
    End With
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub
```

---

### 6.2 ThisWorkbook.cls

#### `Workbook_Open()`

**処理概要**: ブックを開いたとき `Start_GO` を呼び出して初期化処理を実行する。

```vba
Private Sub Workbook_Open()
    Call Start_GO
End Sub
```

#### `Workbook_Activate()`

**処理概要**: ブックがアクティブになったとき、ウィンドウを最大化する。

```vba
Private Sub Workbook_Activate()
    ActiveWindow.WindowState = xlMaximized
End Sub
```

#### `Workbook_BeforeClose(Cancel)`

**処理概要**: ブックを閉じる前に、デバッグフラグが未設定なら画面解除処理を呼び出し、保存確認ダイアログを抑制して閉じる。

```vba
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Application.DisplayAlerts = False
    If Range("DebugF") = "" Then Call 画面解除
    ActiveWorkbook.Saved = True
End Sub
```

---

### 6.3 SQL_Execution.bas

**概要**: ADODB を使用した DB 接続・SQL 実行の共通ライブラリ。

#### `Open_oraconDB()`

**処理概要**: DSN `ricdb`（UID: ric / PWD: t6101）で Oracle DB に ADODB 接続を開く。

```vba
Sub Open_oraconDB()
    mpDSN = "DSN=ricdb;UID=ric;PWD=t6101"
    oraconn.ConnectionString = mpDSN
    oraconn.Open
    oraconn.CursorLocation = adUseClient
End Sub
```

#### `SQL_Exe(mySQL)`

**処理概要**: 任意の SQL 文を ADODB.Connection.Execute で実行し、結果を `rs`（Recordset）に格納する。エラー時は `mpErrDes` にメッセージをセット。

#### `SQL_INSERT_UPDATE(myTBL, myKey, myD(), myN)`

**処理概要**: 指定テーブルにキー条件でレコードの存在確認を行い、なければ INSERT、あれば UPDATE を実行する汎用関数。

**処理フロー**:
1. `SELECT COUNT(*) FROM myTBL WHERE myKey` で件数確認
2. 0件 → INSERT 文を構築・実行
3. 1件以上 → UPDATE 文を構築・実行
4. トランザクション COMMIT

#### `SQL_Delete(myTBL, myWhere)`

**処理概要**: 指定テーブルの指定条件レコードを DELETE する。

#### `Disp_Sheet(mySQL, mySH, myRow, myRecordCount, myColumn, myFieldCount, myF)`

**処理概要**: SQL を実行してシート `mySH` の `myRow` 行・`myColumn` 列以降にレコードセットを貼り付ける。`myF=1` の場合はフィールド名も出力する。

#### `Set_Array(mySQL, myData(), myRecordCount, myFldCount)`

**処理概要**: SQL を実行して結果を2次元配列 `myData(i, j)` に格納する（i: レコード番号、j: フィールド番号）。

---

### 6.4 データベースR.bas

#### `受付番号読込表示()`

**処理概要**: 受付番号（`mpUno`）を条件に在庫ビュー `zaiko_V` を検索し、受付番号・会社コード・製品コードをシートに表示する。

**処理フロー**:
1. `Syouso` が "EB" なら `syouso='4'`、それ以外は `syouso<>'4'` の条件を設定
2. `NyukabiBef` を入荷日下限として WHERE 句に追加
3. `zaiko_V` を SELECT して `Set_Array` で取得
4. 0件 → エラーメッセージ設定
5. 複数件 → 件数と先頭5件をエラーメッセージとして設定
6. 1件 → `Uno` / `KaiCd` / `SeiCd` に値をセット

**SQL**:
```sql
SELECT Uno, kaisyacd, sehncd, kainame
FROM zaiko_V
WHERE uno LIKE '%<mpUno>'
  AND syouso = '4'  -- EBの場合（それ以外は <>'4'）
  AND nyukabi > '<NyukabiBef>'
```

#### `価格ファイル表示(myKaiCD)`

**処理概要**: 会社コードに対応する価格表ファイル情報（フォルダ・ファイル名・備考・登録日等）を `ExSeihinJ`・`tokumst` から取得してシートに表示する。

**SQL**:
```sql
SELECT t.coname, e.folder, e.filename, e.kaibikou, e.toudate, e.touname
FROM ExSeihinJ e, tokumst t
WHERE e.kaisyacd(+) = t.kaisyacd
  AND t.kaisyacd = '<myKaiCD>'
```

#### `製品読込表示(myKaiCD, mySeiCD)`

**処理概要**: 会社コード・製品コードに対応する単価・線量・装置等の製品情報を DB から取得してシートに表示する。

---

### 6.5 表示.bas

#### `価格表の表示()`

**処理概要**: `FolDer` + `FailName` から価格表ファイルのフルパスを生成し、`WScript.Shell.Run` で外部ファイル（Word/Excel等）を開く。

**処理フロー**:
1. `FolDer` と `FailName` が両方入力済みかを確認
2. UNC パス `\\RNTSVR-FS\Sv_cup\営業部照射課\価格表\<FolDer>\<FailName>` を構築
3. `CreateObject("Wscript.Shell").Run FilePathName, 3` でファイルを開く
4. エラー発生時はエラー番号・説明をメッセージ表示

#### `受付番号表示()`

**処理概要**: 受付番号が入力されたとき、画面を初期化してから受付番号・会社・製品・価格表ファイルの情報を順次取得・表示する。

**処理フロー**:
1. `mpUno` = `Uno`（"-"除去）をセット
2. 関連セルを全クリア・表示オブジェクト消去
3. `受付番号読込表示` でDB照合
4. 会社コードから `価格ファイル表示` 呼び出し
5. 製品コードから `製品読込表示` 呼び出し
6. `Hyouji` が "しない" でなければ `価格表の表示` を呼び出す

#### `会社データ表示()`

**処理概要**: 会社コードが変更されたとき、関連セルをクリアして価格ファイル情報を再取得・表示する。

#### `製品データ表示()`

**処理概要**: 製品コードが変更されたとき、製品情報セルをクリアして `製品読込表示` を呼び出す。

---

## 7. DB 接続・外部連携

### ODBC 接続設定

| DSN 名 | UID | PWD | 用途 |
|---|---|---|---|
| `ricdb` | ric | t6101 | Oracle DB接続（在庫・製品・単価テーブル） |

### 参照テーブル一覧

| テーブル名 | 主な用途 | 主要カラム |
|---|---|---|
| `zaiko_V` | 在庫ビュー（受付番号検索） | uno, kaisyacd, sehncd, kainame, syouso, nyukabi |
| `ExSeihinJ` | 製品価格表ファイル情報 | kaisyacd, folder, filename, kaibikou, toudate, touname |
| `ExSeihinZ` | 製品単価マスタ | kaisyacd, sehncd, tanka, tani, keijyou, toudate, touname |
| `tokumst` | 得意先マスタ | kaisyacd, coname |

### 主要 SQL 文

```sql
-- 受付番号から在庫検索
SELECT Uno, kaisyacd, sehncd, kainame
FROM zaiko_V
WHERE uno LIKE '%<受付番号>'
  AND syouso = '4'  -- EB装置の場合
  AND nyukabi > '<YYYYMMDD>'

-- 会社コードから価格表ファイル情報取得
SELECT t.coname, e.folder, e.filename, e.kaibikou, e.toudate, e.touname
FROM ExSeihinJ e, tokumst t
WHERE e.kaisyacd(+) = t.kaisyacd
  AND t.kaisyacd = '<kaisyacd>'

-- 単価のINSERT/UPDATE（SQL_INSERT_UPDATEで動的生成）
INSERT INTO ExSeihinZ (kaisyacd, sehncd, tanka, tani, keijyou, toudate, touname)
VALUES ('<kaisyacd>', '<sehncd>', <tanka>, '<tani>', '<keijyou>', '<date>', '<name>')
-- または
UPDATE ExSeihinZ SET tanka=<tanka>, tani='<tani>', ... WHERE kaisyacd='<cd>' AND sehncd='<cd>'
```

### 外部ファイル連携

| ファイルパス | 用途 |
|---|---|
| `\\RNTSVR-FS\Sv_cup\営業部照射課\価格表\<フォルダ名>\<ファイル名>` | 価格表ファイルの直接表示（WScript.Shell.Run） |

---

## 8. データフロー

```
【受付番号入力フロー】
  B5（Uno）に受付番号を入力
       ↓ Worksheet_Change イベント
  受付番号表示() 呼び出し
       ↓
  DB（zaiko_V）: uno, kaisyacd, sehncd を取得
       ↓
  C5（KaiCd）/ D5（SeiCd）に表示
       ↓
  DB（ExSeihinJ + tokumst）: 価格ファイル情報を取得
       ↓
  H3（FolDer）/ J3（FailName）/ C6（KaiName）に表示
       ↓
  DB（ExSeihinZ）: 単価・単位・備考等を取得
       ↓
  E5〜O5（製品情報行）に表示
       ↓
  Hyouji = "しない" でなければ WScript.Shell で価格表ファイルを開く

【単価登録フロー】
  KaiCd / SeiCd / 単価（E5）/ 単位（F5）を入力
       ↓ 「単価登録」ボタン押下
  単価登録() 呼び出し
       ↓
  DB（ExSeihinZ）: SQL_INSERT_UPDATE で INSERT または UPDATE 実行
       ↓
  完了メッセージ表示

【ファイル登録フロー】
  FolDer / FailName を入力
       ↓ 「ファイル登録」ボタン押下
  単価_フォルダーファイル名更新() 呼び出し
       ↓
  ファイル名妥当性チェック（DB試験登録）
       ↓
  DB（ExSeihinJ）: SQL_INSERT_UPDATE で INSERT または UPDATE 実行
       ↓
  完了メッセージ表示
```

---

## 9. セキュリティ注意事項

olevba 解析結果より以下の警告が検出されています。

| 種別 | キーワード | 内容 |
|---|---|---|
| AutoExec | `Workbook_Open` | ブックを開いたとき自動的に `Start_GO` が実行される |
| AutoExec | `Workbook_Activate` | ブックがアクティブになるたびにウィンドウ最大化が実行される |
| AutoExec | `Workbook_BeforeClose` | ブックを閉じる前に自動で画面解除処理が実行される |
| AutoExec | `Worksheet_Change` | セル変更時に自動でDB照会処理が実行される |
| Suspicious | `Shell` / `Wscript.Shell` | `WScript.Shell.Run` で外部ファイルを実行する処理あり |
| Suspicious | `Run` | 任意のファイルパスを外部プログラムとして実行できる |
| Suspicious | `CreateObject` | `WScript.Shell` オブジェクトを動的生成して外部プロセス起動 |
| Suspicious | `Call` | Excel 4 マクロ形式の DLL 呼び出しの可能性 |
| Suspicious | `Chr` | 文字列難読化の可能性（Chr(13) はCRLFとして使用） |
| Suspicious | `Hex Strings` | VBA 内に16進数エンコード文字列が存在 |
| Suspicious | `Base64 Strings` | VBA 内にBase64エンコード文字列が存在（`Kyou` 変数名等） |

> **注意**: DSN=ricdb の接続文字列（UID/PWD）がVBAコード内にハードコードされています。パスワード `t6101` が平文で記述されているため、コードの漏洩に注意してください。
