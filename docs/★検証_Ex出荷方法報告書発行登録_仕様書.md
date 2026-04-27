# ★検証_Ex出荷方法報告書発行登録 仕様書

> **ファイル種別**: .xlsm（マクロ付き）
> **用途**: 得意先（業者）ごとの出荷方法および報告書発行種別を一覧表示・編集し、DBの ExSeihinJ テーブルに登録する
> **VBA プロジェクトサイズ**: 9モジュール（ThisWorkbook.cls / Sheet1.cls / 登録.bas / 終了処理.bas / デーた抽出.bas / ユーティリティ.bas / 画面クリア引取業者.bas / SQL_Execution.bas / 画面操作1.bas）

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
★検証_Ex出荷方法報告書発行登録.xlsm
├── シート
│   └── 業者一覧（出荷方法・報告書発行種別の一覧・編集画面）
├── VBA モジュール
│   ├── ThisWorkbook.cls       – ブックイベント（Open / BeforeClose）
│   ├── Sheet1.cls             – Worksheet_Change イベント
│   ├── 登録.bas               – 変更行のDBへの INSERT/UPDATE
│   ├── 終了処理.bas           – ブックを閉じる処理
│   ├── デーた抽出.bas         – DB からデータ取得してシートに展開
│   ├── ユーティリティ.bas     – イベント有効化・画面クリア
│   ├── 画面クリア引取業者.bas – 画面消去処理
│   ├── SQL_Execution.bas      – DB 接続・SQL 実行共通ルーチン
│   └── 画面操作1.bas          – 画面操作サンプル・ユーティリティ（参考モジュール）
├── ボタン（3個）
│   ├── 終了
│   ├── 変更を登録する
│   └── 出荷方法/報告書発行種別表示
└── 外部リンク: なし
```

---

## 2. シート詳細

### 2.1 業者一覧

**目的**: 全得意先（kaisyacd < 2000）の出荷方法と報告書発行種別を一覧表示する。現在値（E・G列）と DB 登録値（F・H列）を並べて差分を可視化し、変更を「変更を登録する」ボタンで DB に反映する。

#### レイアウト構造

| セル範囲 | 名前付き範囲 | 内容 |
|---|---|---|
| A1 | `Debug` | デバッグフラグ |
| B1 | — | タイトル「出荷方法／報告書発行種別の登録」 |
| B4 | — | ヘッダー「会社ｺｰﾄﾞ」 |
| C4 | — | ヘッダー「略称」 |
| D4 | — | ヘッダー「会社名」 |
| E4 | — | ヘッダー「出荷方法」（編集列） |
| F4 | — | ヘッダー「出荷方法」（DB参照列・変更検出用） |
| G4 | — | ヘッダー「報告書発行種別」（編集列） |
| H4 | — | ヘッダー「報告書発行種別」（DB参照列・変更検出用） |
| I4 | — | ヘッダー「Sehmstの報告書要不要」（VLOOKUP参照） |
| K4 | — | ヘッダー「会社ｺｰﾄﾞ」（sehmst集計結果） |
| L4 | — | ヘッダー「報告書発行フラグ」（sehmst集計結果） |
| N4 | — | ヘッダー「出荷方法」（選択肢マスタ） |
| O4 | — | ヘッダー「報告書発行種別内容」（選択肢マスタ） |
| B5:H1005 | `HikitoriTB` | 業者データ入力・表示テーブル（全行） |
| K5:L1004 | `HouTB` | sehmst 報告書要不要テーブル（VLOOKUP参照元） |
| I5:I1004 | — | Sehmst 報告書要不要（VLOOKUP数式） |
| N5 | — | 出荷方法選択肢「引取」 |
| N6 | — | 出荷方法選択肢「混載便」 |
| N7 | — | 出荷方法選択肢「保管品」 |
| N8 | — | 出荷方法選択肢「ﾁｬｰﾀ便」 |
| N9 | — | 出荷方法選択肢「納品」 |
| N10 | — | 出荷方法選択肢「品証扱い」 |
| N11 | — | 出荷方法選択肢「γ扱い」 |
| N12 | — | 出荷方法選択肢「営業扱い」 |
| N13 | — | 出荷方法選択肢「その他」 |
| O5 | — | 報告書発行種別「照射後Fax送信」 |
| O6 | — | 報告書発行種別「出荷後Fax送信」 |
| O7 | — | 報告書発行種別「照射後報告書発行」 |
| O8 | — | 報告書発行種別「出荷後報告書発行」 |

（行番号 1005〜1004 は空。データ行は B5:B が空になるまで続く）

#### データ列（B〜I列）

| 列 | ヘッダー | DB フィールド / 備考 |
|---|---|---|
| B | 会社ｺｰﾄﾞ | `tokumst.kaisyacd` |
| C | 略称 | `tokumst.kairname` |
| D | 会社名 | `tokumst.coname` |
| E | 出荷方法（現在値） | `ExSeihinJ.hikitori`（編集対象） |
| F | 出荷方法（DB値） | `ExSeihinJ.hikitori`（DB取得値・比較用） |
| G | 報告書発行種別（現在値） | `ExSeihinJ.housyube`（編集対象） |
| H | 報告書発行種別（DB値） | `ExSeihinJ.housyube`（DB取得値・比較用） |
| I | Sehmst報告書要不要 | VLOOKUP で K:L から取得（0=不要・1=要） |

---

## 3. 名前付き範囲一覧

| 名前 | 参照先 | 説明 |
|---|---|---|
| `Debug` | 業者一覧!$A$1 | デバッグフラグ |
| `HikitoriTB` | 業者一覧!$B$5:$H$1005 | 業者データ入力テーブル（会社コード〜報告書発行種別DB値） |
| `HouTB` | 業者一覧!$K$5:$L$1004 | sehmst より集計した報告書要不要テーブル（VLOOKUP参照元） |

---

## 4. 数式一覧

### 業者一覧シート

| セル範囲 | 数式（代表） | 説明 |
|---|---|---|
| I5:I1004 | `=IF(ISERROR(VLOOKUP(B5,$K$5:$L$1004,2,FALSE)),"",VLOOKUP(B5,$K$5:$L$1004,2,FALSE))` | 同行B列の会社コードで HouTB（K:L）を検索し、報告書発行フラグを表示。エラー時は空白。全1000行に同パターンで設定。 |

（行番号 5〜1004 はすべて同一パターン。I5〜I1004 の数式は省略し代表1行を記載）

---

## 5. ボタン・マクロ対応

| シート | ボタンラベル | 割り当てマクロ | 動作概要 |
|---|---|---|---|
| 業者一覧 | 終了 | `Bookを閉じる` | ブックを保存せずに閉じる |
| 業者一覧 | 変更を登録する | `引取報告書登録` | E列またはG列が変更された行のみ DB（ExSeihinJ）に hikitori/housyube を UPDATE |
| 業者一覧 | 出荷方法/報告書発行種別表示 | `業者名と引取抽出` | DB（tokumst + ExSeihinJ + sehmst）から最新データを取得してシートに再描画 |

---

## 6. VBA モジュール仕様

### 6.1 ThisWorkbook.cls

#### `Workbook_Open()`

**処理概要**: ブックを開いたとき、シート保護を一旦解除してUIのみ保護を再設定し、業者データをDBから取得してシートに展開する。

**処理フロー**:
1. `ActiveSheet.Unprotect` でシート保護を解除
2. `ActiveSheet.Protect UserInterfaceOnly:=True` でUIのみ保護を設定
3. `業者名と引取抽出` を呼び出してDBからデータ取得・表示

```vba
Private Sub Workbook_Open()
    ActiveSheet.Unprotect
    ActiveSheet.Protect UserInterfaceOnly:=True
    Call 業者名と引取抽出
End Sub
```

#### `Workbook_BeforeClose(Cancel)`

**処理概要**: ブックを閉じる前に保存確認ダイアログを抑制し、変更なしで閉じる。

```vba
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Application.DisplayAlerts = False
    ActiveWorkbook.Saved = True
End Sub
```

---

### 6.2 Sheet1.cls

#### `Worksheet_Change(Target)`

**処理概要**: G列（列7）の5行目以降が変更されたとき、同行のI列（報告書要不要フラグ）が0の場合に「報告書発行が不要」とメッセージを表示し、セル値を「不要」に書き戻す。

**処理フロー**:
1. `Application.EnableEvents = False` でイベント無効
2. 変更セルが 5行目以降 かつ G列（列7）かを確認
3. 同行の I列（列9）が 0 なら G列を「不要」に設定してメッセージ表示
4. `Application.EnableEvents = True` でイベント再有効

```vba
Private Sub Worksheet_Change(ByVal Target As Range)
    Application.EnableEvents = False
    With Target
        If .Row > 4 And .Column = 7 Then
            If Cells(.Row, .Column + 2) = 0 Then
                mpErrDes = "製品仕様台帳の報告書発行が不要になっています." & Chr(13) & Chr(13) _
                         & "製品仕様台帳に報告書発行要を登録してください。"
                Cells(.Row, .Column) = "不要"
                MsgBox mpErrDes
            End If
        End If
    End With
    Application.EnableEvents = True
End Sub
```

---

### 6.3 登録.bas

#### `引取報告書登録()`

**処理概要**: 業者一覧シートをスキャンし、出荷方法（E列≠F列）または報告書発行種別（G列≠H列）に差分がある行のみ `ExSeihinJ` テーブルに hikitori / housyube を INSERT/UPDATE する。

**処理フロー**:
1. 業者一覧シートの5行目からB列が空になるまでループ
2. E列 ≠ F列 または G列 ≠ H列 の行を検出
3. kaisyacd（B列）/ hikitori（E列）/ housyube（G列）を `SQL_INSERT_UPDATE` で DB 更新
4. エラーがあれば MsgBox 表示、なければ「更新しました。」を表示

```vba
Sub 引取報告書登録()
    Dim myRow As Integer
    Dim myTBL As String, myKey As String
    Dim myDa() As Variant, myN As Single
    ReDim myDa(1, 2)
    myDa(0, 0) = "kaisyacd": myDa(0, 1) = "hikitori": myDa(0, 2) = "housyube"
    myN = 3
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
                Call SQL_INSERT_UPDATE(myTBL, myKey, myDa(), myN)
            End If
            myRow = myRow + 1
        Loop
    End With
    If mpErrDes <> "" Then
        MsgBox "異常が発生しました" & Chr(13) & Chr(13) & mpErrDes
    Else
        MsgBox "更新しました。"
    End If
End Sub
```

---

### 6.4 終了処理.bas

#### `Bookを閉じる()`

**処理概要**: ブックを保存確認なしで閉じる。ブックが1つのみなら Excel を終了、複数なら当ブックのみ閉じる。

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

**処理概要**: DB の `tokumst`・`ExSeihinJ`・`sehmst` からデータを取得し、業者一覧シートの5行目以降に展開する。報告書不要（フラグ 0 または 空）の行は G・H 列を「不要」と表示する。

**処理フロー**:
1. `HikitoriTB` と `HouTB` をクリア
2. SQL 1: `tokumst JOIN ExSeihinJ` で会社コード・略称・会社名・出荷方法・報告書発行種別を取得 → B列〜H列に `Disp_Sheet` で貼り付け
3. SQL 2: `sehmst` で kaisyacd 別の報告書要不要最大値を集計 → K列〜L列に貼り付け
4. I列の VLOOKUP 結果が 0 または 空の行の G・H 列を「不要」に書き換え
5. 2000行超でエラー停止

**SQL**:
```sql
-- 会社コード・出荷方法・報告書発行種別を取得
SELECT t.kaisyacd, t.kairname, t.coname,
       s.hikitori, s.hikitori, s.housyube, s.housyube
FROM tokumst t, ExSeihinJ s
WHERE t.kaisyacd = s.kaisyacd(+)
  AND t.kaisyacd < '2000'
ORDER BY t.kaisyacd

-- sehmst より報告書要不要を集計
SELECT kaisyacd, MAX(syouho)
FROM sehmst
GROUP BY kaisyacd
ORDER BY kaisyacd
```

#### `報告書不要表示()`

**処理概要**: `HouTB`（K:L列）のフラグが 0 の行の E・F 列を「不要」に設定するサブ処理（補助的な処理。メインは `業者名と引取抽出` 内で実施）。

---

### 6.6 ユーティリティ.bas

#### `イベント有効()`

**処理概要**: `Application.EnableEvents = True` でイベントを強制有効化する（デバッグ用ショートカット）。

#### `画面クリア()`

**処理概要**: `HikitoriTB` と `HouTB` の名前付き範囲をクリアする。

---

### 6.7 画面クリア引取業者.bas

#### `画面消去処理()`

**処理概要**: `HikitoriTB` と `HouTB` をクリアし、カーソルを E5 に移動する。

```vba
Sub 画面消去処理()
    Range("HikitoriTB") = ""
    Range("HouTB") = ""
    Range("E5").Select
End Sub
```

---

### 6.8 SQL_Execution.bas

`Ex単価検索.xlsm` の `SQL_Execution.bas` と同一実装。DSN `ricdb`（UID: ric / PWD: t6101）での Oracle DB 接続・SQL 実行共通ルーチン（`Open_oraconDB` / `SQL_Exe` / `SQL_INSERT_UPDATE` / `SQL_Delete` / `Disp_Sheet` / `Set_Array`）。

---

### 6.9 画面操作1.bas

**概要**: Excel VBA の各種 API 操作サンプルコードを集めた参考モジュール（印刷パラメータ設定、シート保護、セル操作等）。プロダクションロジックではなく開発用リファレンスとして格納されている。

主な関数:
- `Reidai()` – Excel API 操作サンプル集（セル移動・ウィンドウ最大化・Zoom等）
- `印刷パラメータ設定(...)` – 指定シートの印刷余白・向き・用紙サイズを一括設定
- `DriveSearch(myDrive)` – FileSystemObject でドライブ存在確認
- `シート保護()` – 各種シート保護パターンのサンプル

---

## 7. DB 接続・外部連携

### ODBC 接続設定

| DSN 名 | UID | PWD | 用途 |
|---|---|---|---|
| `ricdb` | ric | t6101 | Oracle DB接続（得意先・出荷方法・報告書フラグ） |

### 参照テーブル一覧

| テーブル名 | 主な用途 | 主要カラム |
|---|---|---|
| `tokumst` | 得意先マスタ | kaisyacd, kairname, coname |
| `ExSeihinJ` | 出荷方法・報告書発行種別の登録先 | kaisyacd, hikitori, housyube |
| `sehmst` | 製品仕様台帳（報告書要不要フラグ） | kaisyacd, syouho |

### 主要 SQL 文

```sql
-- 業者一覧取得（tokumst + ExSeihinJ の外部結合）
SELECT t.kaisyacd, t.kairname, t.coname,
       s.hikitori, s.hikitori, s.housyube, s.housyube
FROM tokumst t, ExSeihinJ s
WHERE t.kaisyacd = s.kaisyacd(+)
  AND t.kaisyacd < '2000'
ORDER BY t.kaisyacd

-- 報告書要不要フラグ集計（sehmst）
SELECT kaisyacd, MAX(syouho)
FROM sehmst
GROUP BY kaisyacd
ORDER BY kaisyacd

-- 出荷方法・報告書発行種別の INSERT/UPDATE（引取報告書登録で動的生成）
UPDATE ExSeihinJ
SET hikitori = '<hikitori>', housyube = '<housyube>'
WHERE kaisyacd = '<kaisyacd>'
-- または（未登録時）
INSERT INTO ExSeihinJ (kaisyacd, hikitori, housyube)
VALUES ('<kaisyacd>', '<hikitori>', '<housyube>')
```

---

## 8. データフロー

```
【起動・データ表示フロー】
  ブックを開く（Workbook_Open）
       ↓ シート保護解除→UIのみ保護再設定
  業者名と引取抽出() 呼び出し
       ↓
  DB（tokumst + ExSeihinJ）: 全得意先の出荷方法・報告書発行種別を取得
       ↓ Disp_Sheet で B2:H列 に貼り付け
  DB（sehmst）: kaisyacd 別の報告書要不要集計値を取得
       ↓ Disp_Sheet で K:L 列に貼り付け
  I列 VLOOKUP でフラグを確認
       ↓ フラグ 0 または 空 → G列・H列を「不要」に書き換え
  業者一覧シートに表示完了

【編集フロー】
  G列（報告書発行種別）を変更
       ↓ Worksheet_Change イベント
  同行 I列（報告書要不要）が 0 → 「不要」に戻してメッセージ表示
  （問題なし → そのまま）

【登録フロー】
  「変更を登録する」ボタン押下
       ↓ 引取報告書登録() 呼び出し
  E列 ≠ F列 または G列 ≠ H列 の行を検出
       ↓
  DB（ExSeihinJ）: SQL_INSERT_UPDATE で hikitori / housyube を更新
       ↓
  完了メッセージ表示
```

---

## 9. セキュリティ注意事項

olevba 解析結果より以下の警告が検出されています。

| 種別 | キーワード | 内容 |
|---|---|---|
| AutoExec | `Workbook_Open` | ブックを開いたとき自動的に `業者名と引取抽出` が実行されDB接続が発生する |
| AutoExec | `Workbook_Activate` | ブックがアクティブになるたびに処理が実行される |
| AutoExec | `Workbook_BeforeClose` | ブックを閉じる前に自動処理が実行される |
| AutoExec | `Worksheet_Change` | セル変更時に自動でバリデーション処理が実行される |
| Suspicious | `Environ` | `画面操作1.bas` で `Environ("COMPUTERNAME")` を使用（PC名取得） |
| Suspicious | `CreateObject` | `画面操作1.bas` で `Scripting.FileSystemObject` を動的生成 |
| Suspicious | `ExecuteExcel4Macro` | `画面操作1.bas` でリボン操作に Excel 4 マクロを使用 |
| Suspicious | `Call` | Excel 4 マクロ形式の DLL 呼び出しの可能性 |
| Suspicious | `Chr` | 文字列内に `Chr(13)` 使用（CRLFとして使用） |
| Suspicious | `Hex Strings` | VBA 内に16進数エンコード文字列が存在 |
| Suspicious | `Base64 Strings` | VBA 内にBase64エンコード文字列が存在 |

> **注意**: DSN=ricdb の接続文字列（UID/PWD）が `SQL_Execution.bas` にハードコードされています。パスワード `t6101` が平文で記述されているため、コードの漏洩に注意してください。
> **注意**: `画面操作1.bas` は開発用サンプルコードを含む参考モジュールです。本番環境では不要なコードが含まれているため、定期的な棚卸しが推奨されます。
