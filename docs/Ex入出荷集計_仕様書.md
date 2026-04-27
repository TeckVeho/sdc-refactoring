# Ex入出荷集計.xlsm 仕様書

> **ファイル種別**: マクロ付き Excel ファイル（.xlsm）  
> **用途**: ガンマ線照射装置（1号機・2号機・3号機）における製品の入荷・出荷・装置稼働時間を集計・管理する業務ツール  
> **VBA プロジェクトサイズ**: 164 KB  
> **外部連携ファイル**: `ExDBファイル表示.xlsm`（DB接続・Table登録シートを保持）

---

## 目次

1. [ファイル構成](#1-ファイル構成)
2. [シート詳細](#2-シート詳細)
   - [入荷状況シート](#21-入荷状況シート)
   - [出荷実績シート](#22-出荷実績シート)
   - [稼働時間シート](#23-稼働時間シート)
3. [名前付き範囲一覧](#3-名前付き範囲一覧)
4. [数式一覧](#4-数式一覧)
5. [ボタン・マクロ対応](#5-ボタンマクロ対応)
6. [VBA モジュール仕様](#6-vba-モジュール仕様)
   - [ThisWorkbook](#61-thisworkbook)
   - [Sheet1（出荷実績）](#62-sheet1出荷実績)
   - [Sheet2（入荷状況）](#63-sheet2入荷状況)
   - [BD_Read入出荷](#64-bd_read入出荷)
   - [集計処理](#65-集計処理)
   - [集計終了](#66-集計終了)
   - [表クリア](#67-表クリア)
   - [ExFunction](#68-exfunction)
   - [SQL_Execution](#69-sql_execution)
   - [装置稼働状況](#610-装置稼働状況)
   - [運転時間](#611-運転時間)
7. [DB 接続・SQL 仕様](#7-db-接続sql-仕様)
8. [外部ファイル連携](#8-外部ファイル連携)
9. [データフロー](#9-データフロー)
10. [セキュリティ注意事項](#10-セキュリティ注意事項)

---

## 1. ファイル構成

```
Ex入出荷集計.xlsm
├── シート
│   ├── 入荷状況    (116行 × 30列)
│   ├── 出荷実績    (801行 × 52列)
│   └── 稼働時間    (77行  × 15列)
├── VBA モジュール
│   ├── ThisWorkbook.cls
│   ├── Sheet1.cls           ← 出荷実績シートのイベント
│   ├── Sheet2.cls           ← 入荷状況シートのイベント
│   ├── Sheet3.cls           ← 稼働時間シート（空）
│   ├── BD_Read入出荷.bas
│   ├── 集計処理.bas
│   ├── 集計終了.bas
│   ├── 表クリア.bas
│   ├── ExFunction.bas
│   ├── SQL_Execution.bas
│   ├── 装置稼働状況.bas
│   └── 運転時間.bas
└── 外部リンク
    └── ExDBファイル表示.xlsm
        ├── 使い方
        ├── 項目TB
        ├── 抽出
        ├── 抽出結果
        └── Table登録  ← DSN・テーブル名等の設定値を保持
```

---

## 2. シート詳細

### 2.1 入荷状況シート

**目的**: 指定入荷日に各ガンマ線照射装置（1〜3号機）へ持ち込まれた製品の会社別入荷数と処理時間を一覧表示する。

#### レイアウト構造

| セル範囲 | 内容 |
|---|---|
| C3 | タイトル「ガンマ処理製品入荷状況」 |
| H3（`Nyuuakbi`） | 入荷日入力欄（日付を入力するとマクロ起動） |
| N5（`Today2V`相当）| 本日（TODAY()関数） |
| C5:D5 | 「No」（行ヘッダー） |
| C5:D5 | 「1号機」列グループヘッダー |
| E5:G5 | 「2号機」列グループヘッダー |
| H5:J5 | 「3号機」列グループヘッダー |
| C7 | 「会社名」 |
| D7 | 「入荷数」 |
| E7 | 「会社名」 |
| F7 | 「入荷数」 |
| G7 | 「要処理時間（時間）」 |
| H7 | 「会社名」 |
| I7 | 「入荷数」 |
| J7 | 「要処理時間（時間）」 |
| B8〜B27 | 行番号 1〜20 |
| B28 | 「他」（20件超分） |
| C29, E29, H29 | 「合計」行 |
| D29 | 1号機入荷数合計（`=SUM(D8:D28)`） |
| F29 | 2号機入荷数合計（`=SUM(F8:F28)`） |
| G29 | 2号機処理時間合計（`=SUM(G8:G28)`） |
| I29 | 3号機入荷数合計（`=SUM(I8:I28)`） |
| J29 | 3号機処理時間合計（`=SUM(J8:J28)`） |
| H30:J30 | 「1,2,3号機合計」（`=D29+F29+I29`） |

#### 線源パラメータ領域（M〜R列）

| セル | 名前 | 内容 |
|---|---|---|
| M9 | — | 「Ric2」 |
| N9 | `Ric2T` | Ric2 設定日（例: 2020/07/07） |
| O9 | `Ric2V` | Ric2 速度/M_time（例: 1208） |
| P9 | `Ric2HP` | Ric2 HP/PP（例: 1798） |
| Q9 | `Today2V` | Ric2 今日の速度（減衰計算値） |
| R9 | `TodayHP` | Ric2 今日の HP（比率計算値） |
| M10 | — | 「Ric3」 |
| N10 | `Ric3T` | Ric3 設定日（例: 2019/10/08） |
| O10 | `Ric3M` | Ric3 速度（例: 281） |
| P10 | `Ric3PP` | Ric3 PP（例: 295） |
| Q10 | `Today3M` | Ric3 今日の速度（減衰計算値） |
| R10 | `TodayPP` | Ric3 今日の PP（比率計算値） |

#### DB 読み込み列（T〜AC列）= `Zaiko` 範囲

| 列 | 名前 | DB フィールド |
|---|---|---|
| T | UNO | uno |
| U | KAISYACD | kaisyacd |
| V | SEHNCD | sehncd |
| W | SYOUSO | syouso（装置番号: 1/2/3） |
| X | KAINAME | kainame（会社名） |
| Y | NOUKI | nouki |
| Z | PASS | pass（パス値） |
| AA | NYUKABI | nyukabi（入荷日） |
| AB | NYUKASU | nyukasu（入荷数） |
| AC | INCNT | incnt（入り数） |

---

### 2.2 出荷実績シート

**目的**: 指定出荷日の各社への出荷数量を DB から取得・集計し、出荷数降順で一覧表示する。

#### レイアウト構造

| セル範囲 | 内容 |
|---|---|
| C1（`Syukkabi`） | 出荷日入力欄（日付を入力するとマクロ起動） |
| E1 | タイトル「ガンマ処理品出荷数量」 |
| D2 | 「出荷合計数」ラベル |
| E2（`SyukkaSuu`） | 出荷合計数（`=SUM(E4:E48)+SUM(J4:J49)`） |
| J2 | 「単位：箱」 |

#### 表示テーブル（`Hyou1` / `Hyou2`）

| 範囲 | 名前 | 内容 |
|---|---|---|
| B4:E48 | `Hyou1` | 出荷実績 第1グループ（最大45社） |
| G4:J49 | `Hyou2` | 出荷実績 第2グループ（最大45社） |

各グループの列構成:

| 列オフセット | 内容 |
|---|---|
| +0 | No（連番） |
| +1 | 会社コード |
| +2 | 会社名（VLOOKUP で引当） |
| +3 | 出荷数量 |

- 90社超の場合: 91行目に「その他」として残社合計を表示

#### DB 作業列（N〜S列）= `DataTB` 範囲

| 列 | 名前 | 内容 |
|---|---|---|
| N | — | VLOOKUP 数式列（会社名引当） |
| O | — | 会社コード（DB 生データ） |
| P（`DataN`） | — | レコード件数 |
| R | — | 会社コードマスタ |
| S | — | 会社名マスタ |

#### 登録会社マスタ（ファイル内確認分）

| 会社コード | 会社名 |
|---|---|
| — | （株）大協精工 |
| — | （株）ユーケンサイエンス |
| — | エルゴジャパン（株） |
| — | ＨＯＹＡ株式会社 |
| — | 株式会社　ナカニシ |
| — | （株）海老原ゴム商会 |
| — | （株）フタバ |

---

### 2.3 稼働時間シート

**目的**: 集計期間を指定し、1〜3号機の照射時間・停止時間・稼働率を DB から算出・表示する。

#### レイアウト構造

| セル範囲 | 名前 | 内容 |
|---|---|---|
| C3 | — | タイトル「装置稼働実績」 |
| C5 | — | 「集計期間年月日」ラベル |
| C6 | — | 「集計開始日」ラベル |
| D6 | — | 「集計終了日」ラベル |
| E6 | — | 「総時間」ラベル |
| C7（`KaisiDay`） | `Nengappi` の一部 | 集計開始日入力欄 |
| D7（`SyuuryouDay`） | `Nengappi` の一部 | 集計終了日入力欄 |
| E7（`TotalTime`） | — | 総時間（マクロで計算） |
| C8 | — | 「装置」ラベル |
| D8 | — | 「1号機 照射時間」 |
| E8 | — | 「2号機 コンベア運転時間」 |
| F8 | — | 「3号機 照射時間」 |
| C9 | — | 「照射時間」ラベル |
| C10 | — | 「停止時間」ラベル |
| D9:F9（`SyousyaTime` / `Kekka` の一部） | — | 各装置の照射時間（マクロで書き込み） |
| D10:F10（`Kekka` の一部） | — | 各装置の停止時間（マクロで書き込み） |
| C19 | — | 「稼働率（％）」ラベル |
| D19 | — | 1号機稼働率（`=IFERROR(D9/TotalTime*100,"")` ） |
| E19 | — | 2号機稼働率（`=IFERROR(E9/TotalTime*100,"")` ） |
| F19 | — | 3号機稼働率（`=IFERROR(F9/TotalTime*100,"")` ） |
| B49〜B77 | — | 行番号 1〜29（補助データ列） |

#### O列 固定設定値

| セル | 値 | 用途 |
|---|---|---|
| O11 | 10 | — |
| O12 | 79.4 | — |
| O13 | 72.4 | — |
| O14 | 48.4（`=O13-24`） | — |
| O15 | 24.4 | — |
| O16 | 5.5 | — |

---

## 3. 名前付き範囲一覧

| 名前 | 参照先 | 説明 |
|---|---|---|
| `Nyuuakbi` | 入荷状況!$H$3 | 入荷日入力セル（変更イベントのトリガー） |
| `Nyukabi` | 入荷状況!$P$5 | 日付→数値変換後の入荷日（YYYYMMDD） |
| `N_Hyouji` | 入荷状況!$C$8:$J$28 | 入荷状況表示領域（クリア・書き込み対象） |
| `Zaiko` | 入荷状況!$T$8:$AC$307 | DB から読み込んだ在庫生データ領域 |
| `Ric2T` | 入荷状況!$N$9 | Ric2 線源設定日 |
| `Ric2V` | 入荷状況!$O$9 | Ric2 速度/M_time（設定時） |
| `Ric2HP` | 入荷状況!$P$9 | Ric2 HP/PP（設定時） |
| `Today2V` | 入荷状況!$Q$9 | Ric2 速度（本日・減衰計算後） |
| `TodayHP` | 入荷状況!$R$9 | Ric2 HP（本日・比率計算後） |
| `Ric3T` | 入荷状況!$N$10 | Ric3 線源設定日 |
| `Ric3M` | 入荷状況!$O$10 | Ric3 速度/M_time（設定時） |
| `Ric3PP` | 入荷状況!$P$10 | Ric3 PP（設定時） |
| `Today3M` | 入荷状況!$Q$10 | Ric3 速度（本日・減衰計算後） |
| `TodayPP` | 入荷状況!$R$10 | Ric3 PP（本日・比率計算後） |
| `Syukkabi` | 出荷実績!$D$1 | 出荷日入力セル（変更イベントのトリガー） |
| `SyukkaSuu` | 出荷実績!$E$2 | 出荷合計数 |
| `Hyou1` | 出荷実績!$B$4:$E$48 | 出荷実績表示 第1グループ |
| `Hyou2` | 出荷実績!$G$4:$J$49 | 出荷実績表示 第2グループ |
| `SyukkaTB` | 出荷実績!$C$4:$E$48,出荷実績!$H$4:$J$49 | 出荷実績全表示領域（クリア対象） |
| `DataN` | 出荷実績!$P$1 | DB から取得したレコード件数 |
| `DataTB` | 出荷実績!$N$4:$S$801 | DB 生データ作業領域 |
| `Siki` | 出荷実績!$N$1 | VLOOKUP 数式テンプレートセル |
| `KaisiDay` | 稼働時間!$C$7 | 集計開始日 |
| `SyuuryouDay` | 稼働時間!$D$7 | 集計終了日 |
| `TotalTime` | 稼働時間!$E$7 | 集計期間の総時間（時間単位） |
| `Nengappi` | 稼働時間!$C$7:$D$7 | 集計期間（開始日〜終了日）（クリア対象） |
| `Kekka` | 稼働時間!$D$9:$F$10 | 照射・停止時間の計算結果（クリア対象） |
| `SyousyaTime` | 稼働時間!$D$9:$F$10 | 各装置の照射時間セル |
| `TableName` | [1]Table登録!$C$5 | 外部ファイルの接続テーブル名 |
| `TableNameDB` | [1]Table登録!$D$5 | 外部ファイルの接続 DB 名 |
| `Souti` | [1]Table登録!$G$5 | 外部ファイルの装置設定 |
| `Setumei` | [1]Table登録!$H$5 | 外部ファイルの説明 |
| `_xlnm.Database` | [1]Table登録!$F$5 | 外部ファイルの DB 設定 |

---

## 4. 数式一覧

### 入荷状況シート

| セル | 数式 | 説明 |
|---|---|---|
| N5 | `=TODAY()` | 本日の日付 |
| P5 | `=YEAR(H3) & RIGHT("00" & MONTH(H3),2) & RIGHT("00" & DAY(H3),2)` | 入荷日を YYYYMMDD 文字列に変換（DB クエリ用） |
| Q9 | `=ROUND(Ric2V*0.5^((N5-Ric2T)/1921)+0.0004,0)` | Ric2 の本日速度（放射線源の半減期減衰計算） |
| R9 | `=ROUND(Ric2HP*Today2V/Ric2V,0)` | Ric2 の本日 HP（速度比率から比例計算） |
| Q10 | `=ROUND(Ric3M/(0.5^((N5-Ric3T)/1921)+0.0004),0)` | Ric3 の本日速度（半減期減衰計算の逆数） |
| R10 | `=ROUND(Ric3PP*Ric3M/Today3M,0)` | Ric3 の本日 PP（速度比率から比例計算） |
| D29 | `=SUM(D8:D28)` | 1号機入荷数合計 |
| F29 | `=SUM(F8:F28)` | 2号機入荷数合計 |
| G29 | `=SUM(G8:G28)` | 2号機処理時間合計 |
| I29 | `=SUM(I8:I28)` | 3号機入荷数合計 |
| J29 | `=SUM(J8:J28)` | 3号機処理時間合計 |
| I30 | `=D29+F29+I29` | 1〜3号機入荷数総合計 |

> **放射線減衰計算の根拠**  
> `0.5^(経過日数/1921)` はコバルト60（Co-60）の半減期 1921 日（約 5.27 年）を使用した放射能減衰計算式。

### 出荷実績シート

| セル | 数式 | 説明 |
|---|---|---|
| N1 | `=TRIM(VLOOKUP(O1,$R$4:$S$1003,2,FALSE))` | 数式テンプレート（`Siki`）。マクロが N4〜N(N+3) にコピー |
| E2 | `=SUM(E4:E48)+SUM(J4:J49)` | Hyou1・Hyou2 の出荷数量合計 |
| N3 | `=TRIM(VLOOKUP(O3,$R$4:$S$1003,2,FALSE))` | 会社コード→会社名変換 |

### 稼働時間シート

| セル | 数式 | 説明 |
|---|---|---|
| O14 | `=O13-24` | 補助計算 |
| D19 | `=IFERROR(D9/TotalTime*100,"")` | 1号機稼働率（%） |
| E19 | `=IFERROR(E9/TotalTime*100,"")` | 2号機稼働率（%） |
| F19 | `=IFERROR(F9/TotalTime*100,"")` | 3号機稼働率（%） |

---

## 5. ボタン・マクロ対応

| シート | ボタンラベル | 割り当てマクロ | 動作 |
|---|---|---|---|
| 入荷状況 | 終了 | `終了` | ブックを閉じる（最後のブックなら Excel 終了） |
| 出荷実績 | 終了 | `終了` | 同上 |
| 稼働時間 | 集計 | `線源登録表示` | 集計期間の照射時間・停止時間・稼働率を DB から算出して表示 |

> ボタンはすべて非印刷設定（`PrintObject: False`）

---

## 6. VBA モジュール仕様

### 6.1 ThisWorkbook

```vba
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    ' 保存確認ダイアログを抑制し、変更なしとしてブックを閉じる
    Application.DisplayAlerts = False
    ActiveWorkbook.Saved = True
End Sub

Private Sub Workbook_Open()
    ' 全シートを「UI操作のみ許可」で保護
    Worksheets("稼働時間").Unprotect
    Worksheets("稼働時間").Protect UserInterfaceOnly:=True
    Worksheets("出荷実績").Unprotect
    Worksheets("出荷実績").Protect UserInterfaceOnly:=True
    Worksheets("入荷状況").Unprotect
    Worksheets("入荷状況").Protect UserInterfaceOnly:=True
    ' 全画面クリア後、入荷状況シートを表示
    Call 画面クリア
    Worksheets("入荷状況").Select
End Sub
```

---

### 6.2 Sheet1（出荷実績）

**トリガー**: `Syukkabi`（D1セル）が変更されたとき

```vba
Private Sub Worksheet_Change(ByVal Target As Range)
    If .Row = 1 And .Column = 4 Then  ' D1 = 出荷日セル
        If Range("Syukkabi") = "" Then
            Range("SyukkaTB") = ""    ' 出荷日が空なら表示クリア
        Else
            Call 出荷集計              ' 出荷日が入力されたら集計実行
        End If
    End If
End Sub
```

---

### 6.3 Sheet2（入荷状況）

**トリガー**: `Nyuuakbi`（H3セル）が変更されたとき

```vba
Private Sub Worksheet_Change(ByVal Target As Range)
    If .Row = 3 And .Column = 8 Then  ' H3 = 入荷日セル
        If Range("Nyuuakbi") = "" Then
            Range("N_Hyouji") = ""    ' 入荷日が空なら表示クリア
        Else
            Call 入荷集計              ' 入荷日が入力されたら集計実行
        End If
    End If
End Sub
```

---

### 6.4 BD_Read入出荷

#### `出荷履歴データ()`

| 項目 | 内容 |
|---|---|
| 接続 DSN | `ricdb`（UID: rich / PWD: t6101） |
| 参照テーブル | `syukar`（出荷記録）、`tokumst`（得意先マスタ） |
| 処理概要 | 指定出荷日（YYYYMMDD）で `syukar` から会社コード・出荷数を抽出 → `出荷実績` シート 4行目から書き込み → `tokumst` から会社コード・会社名を取得して R 列に書き込み → `Siki` セルの VLOOKUP 数式を N 列全体にコピー |

```sql
-- 出荷数取得
SELECT TO_NUMBER(kaisyacd), TO_NUMBER(syukasu)
FROM syukar
WHERE syudate = 'YYYYMMDD'
ORDER BY kaisyacd

-- 得意先マスタ取得
SELECT TO_NUMBER(kaisyacd), TRIM(coname)
FROM tokumst
WHERE kaisyacd < '2000'
ORDER BY kaisyacd
```

#### `Ric23HP_Zaiko(myDataN)`

| 項目 | 内容 |
|---|---|
| 接続 DSN | `ricdb`（UID: ric / PWD: t6101） および `ricdb`（UID: rich / PWD: t6101） |
| 参照テーブル | `ExKanriTB`（線源管理）、`zaiko`（在庫: 1号機）、`zaikor`（在庫: 3号機） |
| 処理概要 | 1. Ric2・Ric3 の線源情報を `ExKanriTB` から取得して N9:P10 に書き込み。2. 指定入荷日の在庫を `zaiko`（装置コード 1/2）・`zaikor`（装置コード 3）から取得して T8: 以降に書き込み。3. 装置順→会社順→製品順でソート。 |

```sql
-- 線源パラメータ取得（Ric2: sikibetu='4', Ric3: sikibetu='5'）
SELECT kousinn, ricvm, hppp
FROM ExKanriTB
WHERE sikibetu='4' OR sikibetu='5'
ORDER BY sikibetu

-- 在庫データ取得（1号機）
SELECT uno, kaisyacd, sehncd, syouso, kainame, nouki, pass, nyukabi, nyukasu, incnt
FROM zaiko
WHERE nyukabi = 'YYYYMMDD'

-- 在庫データ取得（3号機）
SELECT uno, kaisyacd, sehncd, syouso, kainame, nouki, pass, nyukabi, nyukasu, incnt
FROM zaikor
WHERE nyukabi = 'YYYYMMDD'
```

---

### 6.5 集計処理

#### `出荷集計()`

1. 出荷日が未入力ならエラーメッセージを表示して終了
2. 表示領域（`Hyou1`・`Hyou2`・`DataTB`）をクリア
3. `出荷履歴データ()` を呼び出して DB からデータ取得
4. 同一会社コードのデータを会社ごとに出荷数集計（バブルソート）
5. 出荷数降順でソート
6. 第1グループ（最大45社）→ B〜E列、第2グループ（最大45社）→ G〜J列 に表示
7. 90社超の場合: 91行目に「その他」として残社合計を表示
8. 集計後に `SyukkaSuu = 0` なら「出荷数量はありません」と表示

#### `入荷集計()`

1. 入荷日が未入力ならエラーメッセージを表示して終了
2. `Zaiko` 範囲・`N_Hyouji` 範囲をクリア
3. `Ric23HP_Zaiko()` を呼び出して DB データ取得
4. 本日の HP / PP を名前付き範囲から取得
5. `Zaiko` 領域（T〜AC列）のデータを装置コード（W列 = syouso）で振り分け:
   - `Case 1` → Ric1（1号機）: 会社コード・会社名・入荷数
   - `Case 2` → Ric2（2号機）: 同上 + HP時間計算（`RoundUp(入荷数/入り数/2,0)*pass/1日HP*24`）
   - `Case 3` → Ric3（3号機）: 同上 + PP時間計算（`RoundUp(入荷数/入り数,0)*pass/1日PP*24`）
6. 各装置ごとに会社単位で集計、HP/PP時間降順でソート（最大20社）
7. 20社超の場合: 「20番以降合計」行に表示
8. 入荷なしの場合: 「入荷はありません」と表示

##### HP/PP 時間計算ロジック

```
2号機 HP 処理時間 = ROUNDUP(入荷数 / 入り数 / 2, 0) × pass ÷ TodayHP × 24
3号機 PP 処理時間 = ROUNDUP(入荷数 / 入り数, 0) × pass ÷ TodayPP × 24
※ 入り数が 0 の場合: 2号機=4、3号機=32 をデフォルト値として使用
```

---

### 6.6 集計終了

```vba
Sub 終了()
    Application.DisplayAlerts = False
    If Application.Workbooks.Count = 1 Then
        Application.Quit      ' 最後のブックなら Excel 終了
    Else
        ThisWorkbook.Close    ' 他のブックが開いていればこのブックのみ閉じる
    End If
End Sub
```

---

### 6.7 表クリア

```vba
Sub 画面クリア()
    ' 全シートの入力・表示領域をクリア
    Range("Hyou1") = ""      ' 出荷実績 第1グループ
    Range("Hyou2") = ""      ' 出荷実績 第2グループ
    Range("DAtaTB") = ""     ' DB作業領域
    Range("Syukkabi") = ""   ' 出荷日
    Range("Zaiko") = ""      ' 在庫データ
    Range("N_Hyouji") = ""   ' 入荷表示領域
    Range("Nyuuakbi") = ""   ' 入荷日
    Range("Nengappi") = ""   ' 集計期間
    Range("Kekka") = ""      ' 稼働時間結果
    Range("TotalTime") = ""  ' 総時間
End Sub
```

---

### 6.8 ExFunction

#### `ExchengeDay(myDate, myType) As String`

日付値を指定フォーマットの文字列に変換する汎用関数。

| `myType` | 出力形式 | 例 |
|---|---|---|
| `"mmdd"` | `MMDD`（4桁） | `"0420"` |
| `"yyyymmdd"` | `YYYYMMDD`（8桁） | `"20120420"` |

- `myDate` が空文字の場合は空文字を返す
- 変換エラー時は先頭に `"?"` を付与して返す

#### `ExchengeDATE(myDate, myType)`

文字列形式の日付を表示用にフォーマット変換する関数。

| `myType` | 入力形式 | 出力形式 |
|---|---|---|
| `"mm/dd"` | `MMDD`（4桁） | `MM/DD` |
| `"yyyy/mm/dd"` | `YYYYMMDD`（8桁） | `YYYY/MM/DD` |

---

### 6.9 SQL_Execution

Oracle DB への ADO + ODBC 接続・SQL 実行を担う共通モジュール。

#### グローバル変数

| 変数名 | 型 | 説明 |
|---|---|---|
| `mpErrDes` | String | 直前の SQL エラー説明 |
| `mpDSN` | String | ODBC 接続文字列（DSN形式） |
| `oraconn` | ADODB.Connection | DB 接続オブジェクト |
| `rs` | ADODB.Recordset | レコードセット |

#### プロシージャ一覧

| プロシージャ | 引数 | 処理概要 |
|---|---|---|
| `Open_oraconDB()` | — | `mpDSN` を使って ODBC 接続を開く（`adUseClient`） |
| `SQL_Exe(mySQL)` | SQL文字列 | SQL を実行してレコードセットを `rs` に格納 |
| `SQL_INSERT_UPDATE(myTBL, myKey, myD(), myN)` | テーブル名・キー条件・データ配列・項目数 | COUNT(*) でレコード存在確認 → なければ INSERT、あれば UPDATE（トランザクション付き） |
| `SQL_Delete(myTBL, myWhere)` | テーブル名・WHERE句 | DELETE 実行（トランザクション付き） |
| `Disp_Sheet(mySQL, mySH, myRow, myRecordCount, myColumn, myFieldCount, myF)` | SQL・シート名・開始行・件数・開始列・フィールド数・ヘッダフラグ | SQL を実行し、結果を指定シートの指定セルに `CopyFromRecordset` で貼り付け |
| `Set_Array(mySQL, myData(), myRecordCount, myFldCount)` | SQL・データ配列・件数・フィールド数 | SQL を実行し、結果を 2次元配列 `myData(i,j)` に格納 |

---

### 6.10 装置稼働状況

#### `稼働状況(myR1, myR1t, myR2, myR2t, myR3, myR3t)`

各装置の現在の稼働状態を DB から取得する。

| 装置 | 参照テーブル | 判定方法 |
|---|---|---|
| 1号機 | `sengnr1` | 最新イベントコード（0:貯蔵中 / 1:照射中 / 2:昇降中 / 3:移動照射中 / E:PCストップ / S:PCスタート） |
| 2号機 | `kyouj2` | 最新記録時刻から現在時刻の差が5分未満なら「照射中」、それ以外「停止中」 |
| 3号機 | `sengnr3` | 最新イベントを 2進数変換し、ビット位置（1,5,14桁目）が「1」なら「照射中」 |

#### `ExHenkan(Number) As String`

10進数 → 2進数変換関数（例: `5` → `"101"`）。3号機のイベントフラグ解析に使用。

---

### 6.11 運転時間

#### `線源登録表示()`

稼働時間シートの「集計」ボタンに割り当てられたメインプロシージャ。

**入力バリデーション**:
1. `KaisiDay` が空なら処理終了
2. `KaisiDay` が日付型でなければエラー表示
3. `SyuuryouDay` が `KaisiDay` より前、または日付型でなければエラー表示
4. `SyuuryouDay` が未来日ならエラー表示
5. 集計範囲が6か月超ならエラー表示

**総時間の計算**:
- 終了日未指定または今日以降: `TotalTime = (Now() - KaisiDay) × 24`
- 終了日指定: `TotalTime = (SyuuryouDay + 1 - KaisiDay) × 24`

**1号機 照射時間計算** (テーブル: `sengnr1`):
1. 集計開始日より前の最大タイマー値（`myMinTime`）を取得
2. 集計終了日の最大タイマー値（`myMaxTime`）を取得
3. 範囲内の移動照射時間（timer < 24）を合算
4. 照射時間 = `myMaxTime - myMinTime + 移動照射時間合計`
5. 停止時間 = `TotalTime - 照射時間`

**2号機 運転時間計算** (テーブル: `kyouj2`):
1. 該当期間のデータ有無を確認
2. `REALSPD = 0`（停止）の時間区間を積算
3. 停止時間 = 積算値 + 日をまたいだ全日停止時間
4. 運転時間 = `TotalTime - 停止時間`

**3号機 照射時間計算** (テーブル: `sengnr3`):
1. 集計開始時の積算タイマー（`sekitime`）をビット解析で取得 → `myMinTime`
2. 集計終了時の積算タイマーを同様に取得 → `myMaxTime`
3. 照射時間 = `myMaxTime - myMinTime`
4. 停止時間 = `TotalTime - 照射時間`

#### `ChengTime(DTime) As Double`

`YYYYMMDD HHMMSS` 形式の文字列を Excel 日付シリアル値（小数点以下=時刻）に変換する。

---

## 7. DB 接続・SQL 仕様

### ODBC 接続設定

| DSN 名 | UID | PWD | 用途 |
|---|---|---|---|
| `ricdb` | `ric` | `t6101` | 在庫・線源管理・3号機センサーデータ |
| `ricdb` | `rich` | `t6101` | 出荷記録・3号機在庫（`zaikor`） |
| `ricdbh` | `ric` | `t6101` | 稼働時間集計（1号機・3号機センサーログ） |

### 参照テーブル一覧

| テーブル名 | 主な用途 | 主要カラム |
|---|---|---|
| `syukar` | 出荷記録 | kaisyacd（会社CD）, syudate（出荷日）, syukasu（出荷数） |
| `tokumst` | 得意先マスタ | kaisyacd（会社CD）, coname（会社名） |
| `zaiko` | 在庫（1・2号機） | uno, kaisyacd, sehncd, syouso, kainame, nouki, pass, nyukabi, nyukasu, incnt |
| `zaikor` | 在庫（3号機） | 同上 |
| `ExKanriTB` | 線源管理 | sikibetu（4=Ric2 / 5=Ric3）, kousinn（更新日）, ricvm（速度）, hppp（HP/PP） |
| `sengnr1` | 1号機センサーログ | sdate（日付）, stime（時刻）, timer（積算時間）, event（状態コード） |
| `kyouj2` | 2号機コンベアログ | rectime（記録日時: YYYYMMDD HHMMSS）, realspd（実速度） |
| `sengnr3` | 3号機センサーログ | sdate（日時）, sekitime（積算タイマー）, event（ビットフラグ） |

---

## 8. 外部ファイル連携

### ExDBファイル表示.xlsm

| 項目 | 内容 |
|---|---|
| ファイル名 | `ExDBファイル表示.xlsm` |
| 参照シート | `Table登録`（シートID: 4） |
| 参照セル | C5（`TableName`）、D5（`TableNameDB`）、F5（`_xlnm.Database`）、G5（`Souti`）、H5（`Setumei`） |
| その他シート | 使い方、項目TB、抽出、抽出結果（いずれも refreshError=true で未接続） |

> 実行時に `ExDBファイル表示.xlsm` が開かれていない場合、これらの名前付き範囲の参照はエラーになる可能性がある。

---

## 9. データフロー

```
【入荷集計フロー】
  入荷状況シート H3（入荷日）入力
       ↓ Worksheet_Change イベント
  入荷集計() 呼び出し
       ↓
  Ric23HP_Zaiko() → DB(ExKanriTB): 線源パラメータ取得 → N9:P10 書き込み
       ↓
  DB(zaiko / zaikor): 在庫データ取得 → T8:AC列 書き込み
       ↓
  装置コード(syouso)で振り分け → HP/PP時間計算
       ↓
  会社ごとに集計 → 処理量降順ソート
       ↓
  入荷状況シート C〜J列（N_Hyouji）に最大20社表示

【出荷集計フロー】
  出荷実績シート D1（出荷日）入力
       ↓ Worksheet_Change イベント
  出荷集計() 呼び出し
       ↓
  出荷履歴データ() → DB(syukar): 出荷データ取得 → N・O列 書き込み
       ↓
  DB(tokumst): 得意先マスタ取得 → R・S列 書き込み
       ↓
  VLOOKUP数式を N 列にコピー（会社名引当）
       ↓
  会社ごとに集計 → 出荷数降順ソート
       ↓
  出荷実績シート B〜J列（Hyou1・Hyou2）に最大90社表示

【稼働時間集計フロー】
  稼働時間シート C7（開始日）・D7（終了日）入力
       ↓「集計」ボタン押下
  線源登録表示() 呼び出し
       ↓ バリデーション
  総時間計算 → TotalTime 書き込み
       ↓
  1号機: DB(sengnr1) → タイマー差分計算 → 照射時間/停止時間 書き込み
  2号機: DB(kyouj2)  → 停止区間積算計算 → 運転時間/停止時間 書き込み
  3号機: DB(sengnr3) → 積算タイマー差分 → 照射時間/停止時間 書き込み
       ↓
  稼働率(%) = 照射時間 / TotalTime × 100 (数式で自動計算)
```

---

## 10. セキュリティ注意事項

`olevba` による解析で以下の警告が検出されています:

| 種別 | キーワード | 内容 |
|---|---|---|
| AutoExec | `Workbook_Open` | ブックを開いたときに自動実行 |
| AutoExec | `Workbook_BeforeClose` | ブックを閉じたときに自動実行 |
| AutoExec | `Worksheet_Change` | セル変更時に自動実行 |
| Suspicious | `Open` | ファイルを開く可能性 |
| Suspicious | `Chr` | 文字列の難読化に使用される可能性（エラーメッセージ用の `Chr(13)` として使用） |
| Suspicious | Hex Strings | Hex エンコード文字列の存在 |
| Suspicious | Base64 Strings | Base64 エンコード文字列の存在（`Siki` セルの VLOOKUP 式と推定） |
| **Critical** | VBA Stomping | **VBA ソースコードと P-code が不一致**（隠蔽されたコードが存在する可能性） |

> ⚠️ **VBA Stomping** が検出されています。これは VBA のソースコードとコンパイル済み P-code が意図的に異なっている状態です。セキュリティ上の懸念があるため、信頼できる環境でのみ使用してください。実際の動作は上記のソースコードとは異なる可能性があります。

---

*本ドキュメントは `Ex入出荷集計.xlsm` を `openpyxl` および `olevba 0.60.2` で解析して生成しました。*
