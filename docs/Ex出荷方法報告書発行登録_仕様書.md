# Ex出荷方法報告書発行登録 仕様書

> **ファイル種別**: .xlsm（マクロ付き）
> **用途**: 取引先（業者）ごとの出荷方法（引取・混載便・保管品など）および報告書発行種別を Oracle DB（ExSeihinJ テーブル）に登録・管理する
> **VBA プロジェクトサイズ**: 8 モジュール（ThisWorkbook, Sheet1, 登録, 終了処理, デーた抽出, ユーティリティ, 画面クリア引取業者, SQL_Execution, 画面操作1）
> **外部連携**: Oracle DB（DSN=ricdb）/ ADODB 接続

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
Ex出荷方法報告書発行登録.xlsm
├── シート
│   └── 業者一覧          （メイン画面・業者別出荷方法入力・登録）
└── VBA モジュール
    ├── ThisWorkbook.cls    （Workbook_Open / BeforeClose）
    ├── Sheet1.cls          （Worksheet_Change イベント）
    ├── 登録.bas            （引取報告書登録処理）
    ├── 終了処理.bas        （ブック閉じる処理）
    ├── デーた抽出.bas      （DB から業者データ取得・表示）
    ├── ユーティリティ.bas  （イベント有効化・画面クリア）
    ├── 画面クリア引取業者.bas （データクリア処理）
    ├── SQL_Execution.bas   （ADODB 接続共通ライブラリ）
    └── 画面操作1.bas       （画面操作サンプル集・ユーティリティ）
```

---

## 2. シート詳細

### 2.1 業者一覧

**目的**: 取引先業者ごとの出荷方法と報告書発行種別を表示・編集し、DB に保存する。DB の `TOKUMST`・`ExSeihinJ`・`SEHMST` テーブルを参照して最大 1000 件の業者データを一覧表示する。

#### レイアウト構造

| セル範囲 | 名前付き範囲 | 内容 |
|---|---|---|
| A1 | `Debug` | デバッグフラグ |
| B1 | — | タイトル「出荷方法／報告書発行種別の登録」 |
| B4 | — | 列ヘッダー：会社コード |
| C4 | — | 列ヘッダー：略称 |
| D4 | — | 列ヘッダー：会社名 |
| E4 | — | 列ヘッダー：出荷方法（現在の登録値） |
| F4 | — | 列ヘッダー：出荷方法（DB 最新値） |
| G4 | — | 列ヘッダー：報告書発行種別（現在の登録値） |
| H4 | — | 列ヘッダー：報告書発行種別（DB 最新値） |
| I4 | — | 列ヘッダー：SEHMST の報告書要不要フラグ |
| K4 | — | 補助テーブルヘッダー：会社コード |
| L4 | — | 補助テーブルヘッダー：報告書発行フラグ |
| N4:O4 | — | 参照テーブルヘッダー（出荷方法種別一覧） |
| B5:H1005 | `HikitoriTB` | 業者データ本体（会社コード〜報告書発行種別） |
| K5:L1004 | `HouTB` | 報告書発行フラグ補助テーブル（DB から取得） |
| I5〜 | — | SEHMST 報告書要不要（数式による `HouTB` の VLOOKUP） |

#### 参照テーブル（N〜O列）

| セル | 内容 |
|---|---|
| N5 | 引取 |
| N6 | 混載便 |
| N7 | 保管品 |
| N8 | チャータ便 |
| N9 | 納品 |
| N10 | 品証扱い |
| O5 | 照射後Fax送信 |
| O6 | 出荷後Fax送信 |
| O7 | 照射後報告書発行 |
| O8 | 出荷後報告書発行 |

#### DB 読み込み列（行 5〜）

| 列 | ヘッダー | DB フィールド |
|---|---|---|
| B | 会社コード | `TOKUMST.kaisyacd` |
| C | 略称 | `TOKUMST.kairname` |
| D | 会社名 | `TOKUMST.coname` |
| E | 出荷方法（編集前） | `ExSeihinJ.hikitori`（DB 保存値） |
| F | 出荷方法（DB 最新） | `ExSeihinJ.hikitori` |
| G | 報告書発行種別（編集前） | `ExSeihinJ.housyube`（DB 保存値） |
| H | 報告書発行種別（DB 最新） | `ExSeihinJ.housyube` |
| I | 報告書要不要 | `SEHMST.max(syouho)`（数式経由） |
| K | 会社コード | `SEHMST.kaisyacd` |
| L | 報告書発行フラグ | `SEHMST.max(syouho)` |

---

## 3. 名前付き範囲一覧

| 名前 | 参照先 | 説明 |
|---|---|---|
| `Debug` | 業者一覧!$A$1 | デバッグフラグ |
| `HikitoriTB` | 業者一覧!$B$5:$H$1005 | 業者データ表示・編集領域 |
| `HouTB` | 業者一覧!$K$5:$L$1004 | 報告書発行フラグ補助テーブル |

---

## 4. 数式一覧

### 業者一覧シート

I 列（I5〜I1004）に同一パターンの数式が全 1000 行分設定されている。

| セル | 数式（代表） | 説明 |
|---|---|---|
| I5 | `=IF(ISERROR(VLOOKUP(B5,$K$5:$L$1004,2,FALSE)),"",VLOOKUP(B5,$K$5:$L$1004,2,FALSE))` | 会社コードを `HouTB` で検索し、報告書発行フラグを返す |
| I6〜I1004 | ～（I5 と同パターン、行番号のみ変化） | 各業者行の報告書要不要フラグ表示 |

> **補足**: `HouTB`（K列・L列）は DB の `SEHMST` から `max(syouho)` を会社コードごとに集計した値が書き込まれる。この値が 0 または空の場合、VBA が G・H列に「不要」と表示する。

---

## 5. ボタン・マクロ対応

ボタンオブジェクトは抽出スクリプトでは検出されなかった（シート保護またはメニュー経由の可能性あり）。

VBA コードでは以下のマクロにショートカット属性が定義されている：

| マクロ名 | ショートカット属性 | 処理概要 |
|---|---|---|
| `イベント有効` | `r\n14` | `EnableEvents = True` に強制復旧 |
| `画面消去処理` | `e\n14` | データクリア（HikitoriTB / HouTB） |

---

## 6. VBA モジュール仕様

### 6.1 ThisWorkbook.cls

#### `Workbook_Open()`

**処理概要**: ブック起動時にシート保護（UIのみ）を設定してから `業者名と引取抽出` を呼び出す。

**処理フロー**:
1. アクティブシートのシート保護を一時解除
2. `UserInterfaceOnly:=True` でシート保護を再設定
3. `業者名と引取抽出` を呼び出す

```vba
Private Sub Workbook_Open()
    ActiveSheet.Unprotect
    ActiveSheet.Protect UserInterfaceOnly:=True
    Call 業者名と引取抽出
End Sub
```

#### `Workbook_BeforeClose(Cancel As Boolean)`

**処理概要**: 確認なしでアラートを非表示にし、変更を保存済みとしてクローズする。

```vba
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Application.DisplayAlerts = False
    ActiveWorkbook.Saved = True
End Sub
```

---

### 6.2 Sheet1.cls（業者一覧シート）

#### `Worksheet_Change(Target As Range)`

**処理概要**: G 列（報告書発行種別）のセルが変更されたとき、製品仕様台帳の報告書発行フラグをチェックし、発行不要の場合は「不要」に戻す。

**処理フロー**:
1. 変更セルが行 5 以降かつ列 G（列7）の場合
2. 同行の列 I（列9）の値が 0 の場合 → 「報告書発行が不要」旨のエラーメッセージ表示
3. 変更セルを強制的に「不要」に上書き

```vba
Private Sub Worksheet_Change(ByVal Target As Range)
    Application.EnableEvents = False
    With Target
        If .Row > 4 And .Column = 7 Then
            If Cells(.Row, .Column + 2) = 0 Then
                Cells(.Row, .Column) = "不要"
                MsgBox "製品仕様台帳の報告書発行が不要になっています..."
            End If
        End If
    End With
    Application.EnableEvents = True
End Sub
```

---

### 6.3 登録.bas

#### `引取報告書登録()`

**処理概要**: 業者一覧シートの E・G 列（変更後）と F・H 列（DB 保存値）を比較し、差異があった行を `ExSeihinJ` テーブルに INSERT または UPDATE する。

**処理フロー**:
1. 業者一覧シートの行 5 から B 列が空になるまでループ
2. E列（出荷方法）≠ F列（DB保存値）または G列（報告書種別）≠ H列（DB保存値）の場合
3. `kaisyacd`、`hikitori`、`housyube` を引数に `SQL_INSERT_UPDATE` を呼び出す
4. 完了後「更新しました。」または異常メッセージを表示

```vba
Sub 引取報告書登録()
    ReDim myDa(1, 2)
    myDa(0, 0) = "kaisyacd"
    myDa(0, 1) = "hikitori"
    myDa(0, 2) = "housyube"
    myTBL = "ExSeihinJ"
    With Sheets("業者一覧")
        myRow = 5
        Do Until .Cells(myRow, 2) = ""
            If .Cells(myRow, 5) <> .Cells(myRow, 6) _
                Or .Cells(myRow, 7) <> .Cells(myRow, 8) Then
                myDa(1, 0) = "'" & .Cells(myRow, 2) & "'"
                myDa(1, 1) = "'" & .Cells(myRow, 5) & "'"
                myDa(1, 2) = "'" & .Cells(myRow, 7) & "'"
                myKey = myDa(0, 0) & "=" & myDa(1, 0)
                Call SQL_INSERT_UPDATE(myTBL, myKey, myDa(), 3)
            End If
            myRow = myRow + 1
        Loop
    End With
End Sub
```

---

### 6.4 終了処理.bas

#### `Bookを閉じる()`

**処理概要**: 保存確認ダイアログを非表示にしてブックを閉じる。最後のブックの場合は Excel ごと終了。

```vba
Sub Bookを閉じる()
    Application.DisplayAlerts = False
    If Application.Workbooks.Count = 1 Then
        Application.Quit
    Else
        ActiveWorkbook.Close
    End If
End Sub
```

---

### 6.5 デーた抽出.bas

#### `業者名と引取抽出()`

**処理概要**: DB から取引先マスタ（`TOKUMST`）と出荷方法登録（`ExSeihinJ`）を結合して業者一覧を表示し、`SEHMST` から報告書要不要フラグを取得してシートに書き込む。

**処理フロー**:
1. イベント無効化、`HikitoriTB`・`HouTB` をクリア
2. `TOKUMST` と `ExSeihinJ` の外部結合で会社コード・略称・会社名・出荷方法・報告書発行種別を 7 列取得 → 行 5 の B 列から書き込み
3. `SEHMST` から会社コード別に `max(syouho)` を集計 → 行 5 の K 列から書き込み
4. 各行の L 列（`max(syouho)`）が 0 または空の場合は G・H 列を「不要」に設定
5. 表計算を手動→自動に戻し、イベント再有効化

```sql
-- 業者一覧取得
SELECT t.kaisyacd, t.kairname, t.coname,
       s.hikitori, s.hikitori, s.housyube, s.housyube
FROM tokumst t, ExSeihinJ s
WHERE t.kaisyacd=s.kaisyacd(+) AND t.kaisyacd<'2000'
ORDER BY t.kaisyacd

-- 報告書要不要フラグ取得
SELECT kaisyacd, max(syouho)
FROM sehmst
GROUP BY kaisyacd
ORDER BY kaisyacd
```

#### `報告書不要表示()`

**処理概要**: `業者名と引取抽出` と同様のロジックで H 列が 0 の行の F・G 列を「不要」に設定する（旧版処理、現状は未使用）。

---

### 6.6 ユーティリティ.bas

| プロシージャ | 処理概要 |
|---|---|
| `イベント有効()` | `Application.EnableEvents = True` に強制復旧 |
| `画面クリア()` | `HikitoriTB` と `HouTB` を空白にクリア |

---

### 6.7 画面クリア引取業者.bas

#### `画面消去処理()`

**処理概要**: `HikitoriTB`・`HouTB` をクリアして E5 セルに移動する。

```vba
Sub 画面消去処理()
    Range("HikitoriTB") = ""
    Range("HouTB") = ""
    Range("E5").Select
End Sub
```

---

### 6.8 SQL_Execution.bas

ファイル1（ExRIC3線量不足報告書住電用）と同一の共通 ADODB ライブラリ。DB 接続文字列: `DSN=ricdb;UID=ric;PWD=t6101`

主なプロシージャ: `Open_oraconDB`、`SQL_Exe`、`SQL_INSERT_UPDATE`、`SQL_Delete`、`Disp_Sheet`、`Set_Array`

---

### 6.9 画面操作1.bas

**処理概要**: 画面操作・Excel 設定変更のサンプルコードを集めたユーティリティモジュール。実運用では直接使用しない。含まれる機能: 印刷パラメータ設定、ドライブ存在確認（`DriveSearch`）、シート保護パターン解説、複数列選択、リボン操作など。

---

## 7. DB 接続・外部連携

### ODBC 接続設定

| DSN 名 | UID | PWD | 用途 |
|---|---|---|---|
| `ricdb` | ric | t6101 | Oracle DB（照射管理システム）への接続 |

### 参照テーブル一覧

| テーブル名 | 主な用途 | 主要カラム |
|---|---|---|
| `TOKUMST` | 取引先マスタ | kaisyacd, kairname, coname |
| `ExSeihinJ` | 出荷方法・報告書発行種別登録 | kaisyacd, hikitori, housyube |
| `SEHMST` | 製品仕様台帳（報告書要不要フラグ） | kaisyacd, syouho |

### 主要 SQL 文

```sql
-- 業者一覧取得（起動時）
SELECT t.kaisyacd, t.kairname, t.coname,
       s.hikitori, s.hikitori, s.housyube, s.housyube
FROM tokumst t, ExSeihinJ s
WHERE t.kaisyacd=s.kaisyacd(+) AND t.kaisyacd<'2000'
ORDER BY t.kaisyacd

-- 報告書要不要フラグ取得（起動時）
SELECT kaisyacd, max(syouho)
FROM sehmst
GROUP BY kaisyacd
ORDER BY kaisyacd

-- 出荷方法・報告書発行種別の INSERT/UPDATE（登録時）
-- 件数 0 → INSERT
INSERT INTO ExSeihinJ (kaisyacd, hikitori, housyube)
VALUES ('[会社コード]', '[出荷方法]', '[報告書発行種別]')

-- 件数 1以上 → UPDATE
UPDATE ExSeihinJ
SET hikitori='[出荷方法]', housyube='[報告書発行種別]'
WHERE kaisyacd='[会社コード]'
```

---

## 8. データフロー

```
【起動フロー】
  Workbook_Open
       ↓
  シート保護（UIのみ）設定
       ↓
  業者名と引取抽出
       ↓
  DB(TOKUMST + ExSeihinJ): 業者一覧取得 → B〜H 列（行5〜）書き込み
       ↓
  DB(SEHMST): 報告書要不要フラグ取得 → K〜L 列（行5〜）書き込み
       ↓
  報告書不要（syouho=0 or 空）の業者は G・H 列を「不要」に設定
       ↓
  I 列の VLOOKUP 数式が HouTB を参照して報告書フラグ表示

【セル変更フロー】
  G 列（報告書発行種別）を編集
       ↓ Worksheet_Change
  I 列（syouho）が 0 の場合 → 「不要」に強制上書き + 警告

【登録フロー】
  引取報告書登録 マクロ実行
       ↓
  E列 ≠ F列 または G列 ≠ H列 の行を検出
       ↓
  DB(ExSeihinJ): 差分行を SQL_INSERT_UPDATE で更新
       ↓
  「更新しました。」メッセージ表示

【終了フロー】
  Bookを閉じる または Workbook_BeforeClose
       ↓
  アラート非表示、保存済みフラグ設定
       ↓
  ブック（または Excel）を閉じる
```

---

## 9. セキュリティ注意事項

| 種別 | キーワード | 内容 |
|---|---|---|
| AutoExec | `Workbook_Open` | 起動時に自動で DB 接続・データ取得が実行される |
| AutoExec | `Workbook_BeforeClose` | アラートなしで保存状態を偽装してクローズする |
| AutoExec | `Worksheet_Change` | セル変更で自動チェック・上書きが実行される |
| Suspicious | `Environ` | `COMPUTERNAME` 環境変数を読み取るコードあり（`画面操作1.bas`） |
| Suspicious | `CreateObject` | `Scripting.FileSystemObject` を動的生成（DriveSearch 関数） |
| Suspicious | `ExecuteExcel4Macro` | `SHOW.TOOLBAR("RIBBON",TRUE)` を Excel 4 Macro で実行（画面操作1.bas） |
| Suspicious | `Chr` | `Chr(13)` による改行（標準的な VBA 記述） |
| Suspicious | `Hex Strings` | VBA プロジェクト内に16進エンコード文字列が検出 |
| Suspicious | `Base64 Strings` | Base64 エンコード文字列が検出 |
| 認証情報 | `PWD=t6101` | DB パスワードが VBA コードにハードコードされている |
| 注意 | `Stop` | `SQL_Exe` でエラー発生時に `Stop` ステートメントが実行される（デバッグ用コードが残存） |
