# Ex会社別入荷集計 仕様書

> **ファイル種別**: .xlsm（マクロ付き）
> **用途**: 会社コードで取引先を指定し、製品別の入荷数量を年別・月別・前年対比で集計して一覧表示する
> **VBA プロジェクトサイズ**: 12 モジュール（ThisWorkbook, Sheet3, Sheet4, 会社集計, 会社製品Read, SQL_Execution, 共通変数, データ退避, 会社コード検索, 入荷集計終了, 画面クリア, Module1）
> **外部連携**: Oracle DB（DSN=ricdb、集計時は UID=rich で履歴データに接続）

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
Ex会社別入荷集計.xlsm
├── シート
│   ├── 設定      （条件入力・製品選択・集計実行）
│   └── 集計表    （集計結果表示・エクスポート用）
└── VBA モジュール
    ├── ThisWorkbook.cls    （Workbook_Open / BeforeClose）
    ├── Sheet3.cls          （設定シート：ダブルクリック/右クリック/Changeイベント）
    ├── Sheet4.cls          （空）
    ├── 会社集計.bas        （集計メイン処理）
    ├── 会社製品Read.bas    （会社名・製品リスト・入荷数取得）
    ├── SQL_Execution.bas   （ADODB 接続共通ライブラリ）
    ├── 共通変数.bas        （Public 変数宣言）
    ├── データ退避.bas      （集計表を新規ブックにコピー）
    ├── 会社コード検索.bas  （略称で会社コード検索）
    ├── 入荷集計終了.bas    （ブック終了処理）
    ├── 画面クリア.bas      （集計画面クリア）
    └── Module1.bas         （イベント復旧用デバッグコード）
```

---

## 2. シート詳細

### 2.1 設定

**目的**: 集計対象の会社・製品・期間を設定し、集計実行ボタンで `集計表` シートに結果を出力する。略称検索で会社コードを引き当てる機能もある。

#### レイアウト構造（設定エリア）

| セル範囲 | 名前付き範囲 | 内容 |
|---|---|---|
| C2 | — | タイトル「受付数量集計」 |
| C4 | — | ラベル「会社コード」 |
| D4 | `KaiCD` | 会社コード入力（変更で会社名・製品リスト自動取得） |
| E4 | `KaiName` | 会社名（DB から自動取得） |
| C5 | — | ラベル「集計方法」 |
| D5 | — | 集計方法選択（「年別」「月別」「前年対比」） |
| C6 | — | ラベル「期間から」（数式で動的変化） |
| D6 | `NenHajime` | 開始年 |
| F6 | `TukiHajime` | 開始月 |
| C7 | — | ラベル「期間まで」 |
| D7 | `NenOwari` | 終了年 |
| F7 | `TukiOwari` | 終了月 |
| N7 | `HouHou` | 集計方法コード（VLOOKUP 数式）：0=年別, 1=月別, 2=前年対比 |
| Q7 | `SDate` | 集計開始日（DATE 数式） |
| Q8 | `EDate` | 集計終了日（DATE 数式） |
| C8 | — | ラベル「集計製品数」 |
| D8 | `SyuukeiN` | 選択された製品数（COUNTIF で自動集計） |
| E8 | `Sentaku` | 一括選択方法（「全て選択」「有効のみ選択」「選択指定」） |
| K4 | `Ryakusyou` | 会社略称検索入力（検索ボタンで `RyakuTB` を更新） |
| K5 | `MojiLike` | 検索方式（「先頭より」「含む」） |

#### レイアウト構造（参照テーブル）

| セル範囲 | 内容 |
|---|---|
| O7:P9 | 集計方法コード変換テーブル（「年別」=0, 「月別」=1, 「前年対比」=2） |
| O8 | 「月別」 |
| O9 | 「前年対比」 |

#### 製品データ領域（行 10〜）

| 列 | 名前付き範囲 | 内容 |
|---|---|---|
| B | — | 有効/無効（数式: I列の値が `**********` なら「無効」） |
| C | — | 集計対象選択（`○` でチェック、ダブルクリックで切り替え） |
| D | — | 製品コード（`sehmst.sehncd`） |
| E | — | 線量（`sehmst.siteisn`） |
| F | — | 製品名（`sehmst.seiname`） |
| G〜I | — | 空（将来用またはDB追加データ用） |
| I | — | 有効判定値（`**********` = 無効） |
| N | — | 入荷数量（`zaikor` から取得後 VLOOKUP で表示） |
| O | — | 製品コード（入荷数テーブル） |
| P | — | 数量（入荷数テーブル） |

（行番号 11〜510 がデータ領域、行 511〜1100 は空）

#### 会社コード検索テーブル（K〜L 列）

| セル範囲 | 名前付き範囲 | 内容 |
|---|---|---|
| K11:L1100 | `RyakuTB` | 会社コード検索結果（K: 会社コード、L: 会社名） |

---

### 2.2 集計表

**目的**: `集計MAIN` 実行後に集計結果が書き込まれる出力シート。初期状態は空（A1 のみ）。

| 行・列 | 内容 |
|---|---|
| 1行目 | 会社名（`KaiName` の値） |
| 3行目 | ヘッダー（製品コード / 線量 / 製品名 / 年/月列.../ 合計または前年対比） |
| 4行目〜 | 製品ごとのデータ行（選択製品数分） |
| 最終行 | 全製品の合計行 |

---

## 3. 名前付き範囲一覧

| 名前 | 参照先 | 説明 |
|---|---|---|
| `EDate` | 設定!$Q$8 | 集計終了日（DATE 数式） |
| `Goukei` | 設定!$O$11:$P$510 | 入荷数量一時テーブル（DB から取得） |
| `HouHou` | 設定!$N$7 | 集計方法コード（VLOOKUP 結果: 0/1/2） |
| `KaiCD` | 設定!$D$4 | 会社コード |
| `KaiName` | 設定!$E$4 | 会社名 |
| `MojiLike` | 設定!$K$5 | 略称検索方式（「先頭より」「含む」） |
| `NenHajime` | 設定!$D$6 | 集計開始年 |
| `NenOwari` | 設定!$D$7 | 集計終了年 |
| `Ryakusyou` | 設定!$K$4 | 会社略称検索入力 |
| `RyakuTB` | 設定!$K$11:$L$1100 | 会社コード検索結果テーブル |
| `SDate` | 設定!$Q$7 | 集計開始日（DATE 数式） |
| `SeihinTB` | 設定!$C$11:$I$510 | 製品一覧データ領域 |
| `Sentaku` | 設定!$E$8 | 一括選択方法 |
| `SuryouTB` | 設定!$O$11:$Q$510 | 入荷数量テーブル（製品コード・数量） |
| `SyuukeiN` | 設定!$D$8 | 集計対象製品数（COUNTIF） |
| `TukiHajime` | 設定!$F$6 | 集計開始月 |
| `TukiOwari` | 設定!$F$7 | 集計終了月 |

---

## 4. 数式一覧

### 設定シート

| セル | 数式 | 説明 |
|---|---|---|
| C6 | `=IF(D5=O9,"集計年月","期間から")` | 集計方法が「前年対比」なら「集計年月」に変更 |
| G6 | `=IF(OR(D5=O8,D5=O9),"月","")` | 月別・前年対比の場合のみ「月」ラベル表示 |
| C7 | `=IF(D5=O9,"","期間まで")` | 前年対比の場合は「期間まで」を非表示 |
| E7 | `=IF(OR(D5="",D5=O9),"","年")` | 前年対比・未選択の場合は「年」ラベル非表示 |
| G7 | `=IF(D5=O8,"月","")` | 月別の場合のみ「月」ラベル表示 |
| N7 | `=IF(ISERROR(VLOOKUP(D5,O7:P9,2,FALSE)),-1,VLOOKUP(D5,O7:P9,2,FALSE))` | 集計方法文字列をコード（0/1/2）に変換（未設定は -1） |
| Q7 | `=DATE(NenHajime,TukiHajime,"1")` | 集計開始日を DATE 型で計算 |
| D8 | `=COUNTIF(C11:C510,"○")` | 「○」が付いた製品数を集計（`SyuukeiN`） |
| Q8 | `=DATE(NenOwari,TukiOwari,"1")` | 集計終了日を DATE 型で計算 |
| B11〜B510 | `=IF(I11="","",IF(I11="**********","無効","有効"))` | I列の値で有効/無効を判定（全行同パターン） |
| N11〜N510 | `=IF(ISERROR(VLOOKUP(D11,$O$11:$P$42,2,FALSE)),0,VLOOKUP(D11,$O$11:$P$42,2,FALSE))` | 製品コードで `Goukei` テーブルを検索して入荷数量を表示 |

> **B列有効判定**: I列の値が `**********` の製品は「無効」と表示され、集計時に有効のみ選択すると除外される。
> **N列入荷数量**: `入荷数読込` プロシージャ実行後に `Goukei`（O〜P列）の数値が VLOOKUP で参照される。

---

## 5. ボタン・マクロ対応

| シート | ボタンラベル | 割り当てマクロ | 動作概要 |
|---|---|---|---|
| 設定 | 集計実行 | `集計MAIN` | 入力チェック → 製品情報読み込み → 集計表に出力 |
| 設定 | 記録 | `新規ブックに退避` | 集計表を新規ブックにコピーして保存 |
| 設定 | 終了 | `入荷数集計表示終了` | ブック（または Excel）を終了 |
| 設定 | 検索 | `会社コード表示` | 略称で会社コードを検索して `RyakuTB` に表示 |

---

## 6. VBA モジュール仕様

### 6.1 ThisWorkbook.cls

#### `Workbook_Open()`

**処理概要**: ブック起動時に最大化・シート保護設定・画面クリアを行う。

```vba
Private Sub Workbook_Open()
    ActiveWindow.WindowState = xlMaximized
    Sheets("設定").Protect userinterfaceonly:=True
    ActiveSheet.EnableSelection = xlUnlockedCells
    Call 集計画面クリア
End Sub
```

#### `Workbook_BeforeClose(Cancel As Boolean)`

アラートなしで保存済みフラグを設定してサイレントクローズ。

---

### 6.2 Sheet3.cls（設定シート）

#### `Worksheet_BeforeDoubleClick(Target, Cancel)`

**処理概要**: 製品データ行（行11〜510、C列）のセルをダブルクリックすると「○」のオン/オフを切り替える。製品コードが空の行は操作不可。

#### `Worksheet_BeforeRightClick(Target, Cancel)`

**処理概要**: 会社コード検索テーブル（K〜L列、行11〜499）の右クリックで会社コードを `KaiCD` に転記する。

#### `Worksheet_Change(Target)`

**処理概要**: `KaiCD`（D4）変更 → 会社名と製品リストをDBから取得。集計方法（D5）変更 → 入力フィールドのロック状態を切り替える。

**処理フロー（KaiCD変更）**:
1. カーソルを砂時計に変更
2. `会社名読込` で会社名取得（`tokumst`）
3. 成功時に `製品名読込` で製品リスト取得（`sehmst`）
4. カーソルを元に戻す

**処理フロー（集計方法変更）**:
- 「年別」→ 開始月・終了月をロック、終了年をアンロック
- 「月別」→ 全フィールドをアンロック
- 「前年対比」→ 開始月をアンロック、終了年・終了月をロック

---

### 6.3 会社集計.bas

#### `集計MAIN()`

**処理概要**: 入力値を検証し、選択製品の情報を配列に読み込み、集計表シートを初期化して期間ヘッダーを生成後、`会社入荷数集計` を呼び出す。

**処理フロー**:
1. 入力チェック（会社コード・集計方法・期間・集計製品数）
2. `Sentaku` が「全て選択」または「有効のみ選択」の場合は一括で C 列に「○」を付与
3. `SeihinTB` から「○」の製品コード・線量・製品名を配列 `mmDA` に読み込み
4. 集計表をクリアし、製品名・ヘッダー行を書き込み
5. 集計方法に応じて年/月のヘッダー列を生成
6. `DSN=ricdb;UID=rich;PWD=t6101`（履歴DB）で `会社入荷数集計` を呼び出す

#### `会社入荷数集計()`

**処理概要**: 集計表のヘッダー列を順に処理し、各期間の入荷数量を取得・表示する。最後に行合計または前年対比を計算する。

**処理フロー**:
1. 集計表のヘッダー列（4列目以降）を空になるまでループ
2. 集計方法（年別/月別/前年対比）に応じて `uno` 条件文字列を生成
3. `入荷数読込` で DB から集計値を取得（`Goukei` テーブルに書き込み）
4. `集計表に表示` で `設定` シートの N 列（VLOOKUP 経由）から値を読み取り集計表に転記
5. 全期間処理後に「合計」列または「前年対比」列を追加

#### `集計表に表示(myCol)`

**処理概要**: `設定` シートの N 列（入荷数集計値）を読み取り、選択製品分を集計表の指定列に書き込む。製品別累計（`mpTotal`）を更新する。

---

### 6.4 会社製品Read.bas

#### `会社名読込(myAns)`

**処理概要**: `KaiCD` をゼロ埋め4桁に正規化し、`tokumst` から会社名を取得して `KaiName` に書き込む。

```sql
SELECT TRIM(coname) FROM tokumst@ricdb WHERE kaisyacd='[KaiCDゼロ埋め4桁]'
```

#### `製品名読込()`

**処理概要**: `sehmst` から製品コード・線量・製品名を取得し、`設定` シートの D11〜F510 に書き込む。

```sql
SELECT sehncd, TO_NUMBER(siteisn), TRIM(seiname), '', '', ric
FROM sehmst@ricdb
WHERE kaisyacd='[KaiCDゼロ埋め4桁]'
ORDER BY sehncd
```

#### `入荷数読込(myUno)`

**処理概要**: `zaikor` テーブルから会社・製品・期間で入荷数を集計し、`Goukei`（O11〜P510）に書き込む。

```sql
SELECT sehncd, SUM(nyukasu)
FROM zaikor
WHERE kaisyacd='[KaiCDゼロ埋め4桁]'
  [期間条件] [製品コード条件（IN句）]
GROUP BY sehncd ORDER BY sehncd
```

---

### 6.5 SQL_Execution.bas

同一 DB 接続ライブラリ（他ファイルと共通パターン）。ただし `集計MAIN` から呼び出す際は `mpDSN = "DSN=ricdb;UID=rich;PWD=t6101"` で履歴 DB に接続する。

---

### 6.6 会社コード検索.bas

#### `会社コード表示()`

**処理概要**: `Ryakusyou`（略称）と `MojiLike`（「先頭より」「含む」）に応じて `TOKUMST` を検索し、`RyakuTB` に会社コード・会社名一覧を表示する。

```sql
SELECT kaisyacd, RTRIM(TRANSLATE(coname,'　',' '))
FROM tokumst
WHERE kairname LIKE '[略称]%' AND kaisyacd<'2000'  -- 先頭より
-- または
WHERE kairname LIKE '%[略称]%' AND kaisyacd<'2000'  -- 含む
ORDER BY kairname
```

---

### 6.7 データ退避.bas

#### `新規ブックに退避()`

**処理概要**: 集計表が空でない場合、新規ブックを作成して「集計表」シートをコピーする。

```vba
Workbooks.Add
Sheets("集計表").Copy Before:=Workbooks(myNewBookName).Sheets("Sheet1")
```

---

### 6.8 画面クリア.bas

#### `集計画面クリア()`

**処理概要**: 集計表をクリアし、設定シートの全入力セルをリセットする。

リセット対象: `集計表` の全セル、`SeihinTB`、`SuryouTB`、`KaiCD`、`NenHajime`、`NenOwari`、`TukiHajime`、`TukiOwari`、`KaiName`、`RyakuTB`、`Ryakusyou`

---

### 6.9 共通変数.bas

```vba
Public mpSehnCD As String    ' 製品コード WHERE 句（IN 句形式）
Public mpTotal() As Single   ' 製品別累計配列
```

---

### 6.10 入荷集計終了.bas

#### `入荷数集計表示終了()`

最後のブックならば Excel ごと終了、そうでなければブックを閉じる。

---

## 7. DB 接続・外部連携

### ODBC 接続設定

| DSN 名 | UID | PWD | 用途 |
|---|---|---|---|
| `ricdb` | ric | t6101 | マスタデータ取得（会社名・製品名の読み込み時） |
| `ricdb` | rich | t6101 | **履歴データ取得（入荷数集計時）** ※他ファイルと異なる |

> **重要**: 集計処理（`集計MAIN`）では `UID=rich` の履歴データ用接続文字列を使用する。会社名・製品名の読み込みには `UID=ric` を使用。

### 参照テーブル一覧

| テーブル名 | 主な用途 | 主要カラム |
|---|---|---|
| `tokumst` | 会社マスタ（略称検索・会社名取得） | kaisyacd, coname, kairname |
| `sehmst` | 製品仕様台帳（製品リスト取得） | kaisyacd, sehncd, siteisn, seiname, ric |
| `zaikor` | 入荷実績（入荷数量集計）※履歴DB | kaisyacd, sehncd, nyukasu, uno |

### 主要 SQL 文

```sql
-- 会社名取得
SELECT TRIM(coname) FROM tokumst@ricdb WHERE kaisyacd='[4桁コード]'

-- 製品リスト取得
SELECT sehncd, TO_NUMBER(siteisn), TRIM(seiname), '', '', ric
FROM sehmst@ricdb WHERE kaisyacd='[4桁コード]' ORDER BY sehncd

-- 会社コード検索（先頭より）
SELECT kaisyacd, RTRIM(TRANSLATE(coname,'　',' '))
FROM tokumst
WHERE kairname LIKE '[略称]%' AND kaisyacd<'2000' ORDER BY kairname

-- 入荷数量集計（年別例）
SELECT sehncd, SUM(nyukasu)
FROM zaikor
WHERE kaisyacd='[4桁コード]'
  AND uno>'[年]000000' AND uno<'[年]999999'
  AND (sehncd='[コード1]' OR sehncd='[コード2]' OR ...)
GROUP BY sehncd ORDER BY sehncd

-- 入荷数量集計（月別例）
SELECT sehncd, SUM(nyukasu)
FROM zaikor
WHERE kaisyacd='[4桁コード]'
  AND uno>'[年月]0000' AND uno<='[年月]9999'
  AND (sehncd='[コード1]' OR ...)
GROUP BY sehncd ORDER BY sehncd
```

---

## 8. データフロー

```
【起動フロー】
  Workbook_Open
       ↓
  設定シートを最大化・保護
  集計画面クリア（全入力・集計表をリセット）

【会社コード入力フロー】
  KaiCD(D4) に会社コード入力
       ↓ Worksheet_Change(D4)
  DB(tokumst): 会社名取得 → KaiName に表示
       ↓
  DB(sehmst): 製品リスト取得 → SeihinTB(D11:F510) に書き込み
  B列数式が有効/無効を自動判定

【会社コード検索フロー】
  Ryakusyou に略称入力 → 検索ボタン（会社コード表示）
       ↓
  DB(tokumst): LIKE 検索 → RyakuTB(K11:L1100) に表示
       ↓
  K〜L列を右クリックで KaiCD に会社コードを転記 → 会社名・製品リスト自動取得

【製品選択フロー】
  設定シートの C 列をダブルクリックで「○」トグル
  または Sentaku で「全て選択」「有効のみ選択」
       ↓
  SyuukeiN(D8) の COUNTIF が自動更新

【集計実行フロー】
  集計実行ボタン（集計MAIN）
       ↓
  入力値バリデーション（会社コード・集計方法・期間・製品数）
       ↓
  「○」製品コードを配列 mmDA に読み込み、IN 句（mpSehnCD）を生成
       ↓
  集計表シートをクリア → 製品名リストと期間ヘッダーを書き込み
       ↓
  mpDSN = "DSN=ricdb;UID=rich;PWD=t6101"（履歴DB）に切り替え
       ↓
  各期間列ごとにループ:
    uno 条件文字列を生成（年別/月別/前年対比）
    DB(zaikor): 入荷数量 → Goukei(O11:P510) に書き込み
    N列の VLOOKUP 数式が入荷数量を参照
    集計表の該当列に値を転記
       ↓
  合計列（または前年対比列）を追加
  「集計終わりました。」メッセージ

【退避フロー】
  記録ボタン（新規ブックに退避）
       ↓
  集計表を新規ブックにコピー → 「退避しました。」メッセージ
```

---

## 9. セキュリティ注意事項

| 種別 | キーワード | 内容 |
|---|---|---|
| AutoExec | `Workbook_Open` | 起動時に `集計画面クリア` が自動実行される |
| AutoExec | `Workbook_BeforeClose` | アラートなしで保存済みフラグを設定してクローズ |
| AutoExec | `Worksheet_Change` | `KaiCD` 変更で自動的に DB 接続・データ取得が実行される |
| Suspicious | `Windows` | `Windows(myBookname).Activate` でブック間操作 |
| Suspicious | `Open` | ブック間の `Copy` 操作が含まれる |
| Suspicious | `Call` | Excel 4 Macro として検知 |
| Suspicious | `Hex Strings` | VBA プロジェクト内に16進エンコード文字列が検出 |
| Suspicious | `Base64 Strings` | Base64 エンコード文字列が検出 |
| 認証情報 | `PWD=t6101` | DB パスワードが VBA コードにハードコードされている |
| 認証情報 | `UID=rich` | 集計時は通常と異なる `rich` ユーザーで履歴 DB に接続（権限差異に注意） |
| デバッグ残存 | `Stop` | `SQL_Exe` エラー時に `Stop` ステートメントが実行される（デバッグ用コードが残存） |
| 注意 | `@ricdb` | `sehmst@ricdb`・`tokumst@ricdb` の DB リンク経由でアクセスしている |
