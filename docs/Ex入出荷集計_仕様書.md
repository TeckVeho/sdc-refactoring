# Ex入出荷集計 仕様書

> **対象ファイル**: Ex入出荷集計.xlsm
> **ファイル種別**: .xlsm（マクロ付き）
> **用途**: ガンマ線照射処理における入荷（1・2・3号機別）・出荷（会社別）の日次集計、および装置稼働時間の算出を行う業務集計ツール
> **VBA プロジェクト**: モジュール 12 本（.bas 7 / .cls 4 / .frm 0）
> **外部連携**: DSN=ricdb / ricdbh（Oracle）、DB接続先: ODBC経由
> **配置パス**: `D:\ExSys実行\`
> **解析日**: 2026-06-29（excel-to-md スキルによる自動解析）

---

## 1. ファイル構成

### 1.1 シート一覧

| ✓ | # | シート名 | codeName | 用途 | 最大行 | 最大列 | 印刷範囲 |
|---|---|----------|----------|------|--------|--------|----------|
| ✓ | 1 | 「入荷状況」 | Sheet2 | ガンマ処理製品の入荷集計（1・2・3号機別） | 116 | AD | `B3:J30` |
| ✓ | 2 | 「出荷実績」 | Sheet1 | ガンマ処理品の出荷数量集計（会社別） | 801 | AZ | `B1:J50` |
| ✓ | 3 | 「稼働時間」 | Sheet3 | 装置稼働実績（照射時間・停止時間・稼働率） | 77 | O | — |

### 1.2 フォーム一覧

本ファイルにユーザーフォーム（.frm）は存在しない。

### 1.3 モジュール一覧

| ✓ | # | モジュール名 | 種別 | プロシージャ数 | 概要 |
|---|---|-------------|------|---------------|------|
| ✓ | 1 | **ThisWorkbook** | .cls | 2 | ブック開閉時のシート保護制御・画面初期化 |
| ✓ | 2 | **Sheet1** | .cls | 1 | 「出荷実績」シート変更イベント（出荷集計トリガー） |
| ✓ | 3 | **Sheet2** | .cls | 1 | 「入荷状況」シート変更イベント（入荷集計トリガー） |
| | 4 | **Sheet3** | .cls | 0 | 「稼働時間」シート（イベントなし） |
| ✓ | 5 | **BD_Read入出荷** | .bas | 2 | 出荷履歴・入荷在庫データのDB読込 |
| ✓ | 6 | **ExFunction** | .bas | 2 | 日付変換ユーティリティ関数 |
| ✓ | 7 | **集計処理** | .bas | 2 | 出荷集計・入荷集計のメインロジック |
| ✓ | 8 | **集計終了** | .bas | 1 | ブック終了処理 |
| ✓ | 9 | **表クリア** | .bas | 1 | 全シートのデータクリア |
| ✓ | 10 | **SQL_Execution** | .bas | 6 | ADODB 経由のDB接続・SQL実行・シート出力汎用基盤 |
| ✓ | 11 | **装置稼働状況** | .bas | 2 | 1・2・3号機のリアルタイム稼働状況取得 |
| ✓ | 12 | **運転時間** | .bas | 3 | 装置稼働時間の集計計算 |

---

## 2. シート詳細

### 2.0 可視性一覧

| シート名 | 可視性 | VBA制御 |
|----------|--------|---------|
| 「入荷状況」 | 表示 | `Workbook_Open` で Protect/Unprotect 制御 |
| 「出荷実績」 | 表示 | `Workbook_Open` で Protect/Unprotect 制御 |
| 「稼働時間」 | 表示 | `Workbook_Open` で Protect/Unprotect 制御 |

全シートは `Workbook_Open` で `Unprotect` → `Protect UserInterfaceOnly:=True` を実行し、VBA からのセル書込みを許可しつつユーザーの手動編集を制限する。

### 2.1 「入荷状況」シート

**タイトル**: ｶﾞﾝﾏ処理製品入荷状況

#### レイアウト構造

| 領域 | セル範囲 | 内容 |
|------|----------|------|
| タイトル | `B3:C3` | "New" / "ｶﾞﾝﾏ処理製品入荷状況" |
| 入荷日入力 | `G3`（ラベル）/ `H3`（入力） | ユーザーが入荷日を入力する起点セル（名前: `Nyuuakbi`） |
| 識別子 | `K3` | "Ex5" |
| 装置ヘッダ | `B5:J5` | No / 1号機 / 2号機 / 3号機（結合セル） |
| 本日日付 | `M5:N5` | "本日" / `=TODAY()` |
| 入荷日フォーマット | `P5` | `=YEAR(H3)&RIGHT("00"&MONTH(H3),2)&RIGHT("00"&DAY(H3),2)`（名前: `Nyukabi`） |
| データヘッダ | `C7:J7` | 会社名 / 入荷数（×3装置分、2号機・3号機は要処理時間列あり） |
| 線源設定エリア | `M7:R10` | Ric2/Ric3の設定日・速度・HP/PP（設定値 + 今日の計算値） |
| データ表示 | `C8:J27` | No.1〜20（名前: `N_Hyouji`） |
| 20番以降 | `C28:J28` | "他"（20社超の場合の合算行） |
| 合計行 | `C29:J30` | 装置別合計 / 1,2,3号機総合計 |
| DB取得バッファ | `T8:AC307` | 在庫データ一時格納領域（名前: `Zaiko`、非表示列） |

#### 主要セル詳細

| ✓ | セル | 名前 | 内容 | 数式/備考 |
|---|------|------|------|-----------|
| ✓ | `H3` | Nyuuakbi | 入荷日入力 | ユーザー入力。変更時 `Worksheet_Change` が `入荷集計()` を起動 |
| ✓ | `N5` | — | 本日日付 | `=TODAY()` |
| ✓ | `P5` | Nyukabi | 入荷日（yyyymmdd） | `=YEAR(H3)&RIGHT("00"&MONTH(H3),2)&RIGHT("00"&DAY(H3),2)` |
| ✓ | `N9` | Ric2T | Ric2 設定日 | DB取得値（ExKanriTB） |
| ✓ | `O9` | Ric2V | Ric2 速度 | DB取得値（ExKanriTB） |
| ✓ | `P9` | Ric2HP | Ric2 HP/日 | DB取得値（ExKanriTB） |
| ✓ | `Q9` | Today2V | Ric2 今日の速度 | `=ROUND(Ric2V*0.5^((N5-Ric2T)/1921)+0.0004,0)` — Co-60 半減期補正 |
| ✓ | `R9` | TodayHP | Ric2 今日のHP | `=ROUND(Ric2HP*Today2V/Ric2V,0)` |
| ✓ | `N10` | Ric3T | Ric3 設定日 | DB取得値（ExKanriTB） |
| ✓ | `O10` | Ric3M | Ric3 M_time | DB取得値（ExKanriTB） |
| ✓ | `P10` | Ric3PP | Ric3 PP | DB取得値（ExKanriTB） |
| ✓ | `Q10` | Today3M | Ric3 今日のM_time | `=ROUND(Ric3M/(0.5^((N5-Ric3T)/1921)+0.0004),0)` — Co-60 半減期補正 |
| ✓ | `R10` | TodayPP | Ric3 今日のPP | `=ROUND(Ric3PP*Ric3M/Today3M,0)` |
| | `D29` | — | 1号機合計 | `=SUM(D8:D28)` |
| | `F29` | — | 2号機入荷数合計 | `=SUM(F8:F28)` |
| | `G29` | — | 2号機処理時間合計 | `=SUM(G8:G28)` |
| | `I29` | — | 3号機入荷数合計 | `=SUM(I8:I28)` |
| | `J29` | — | 3号機処理時間合計 | `=SUM(J8:J28)` |
| | `I30` | — | 全号機合計 | `=D29+F29+I29` |

#### 結合セル

| 範囲 | 内容 |
|------|------|
| `B5:B7` | No ヘッダ |
| `C5:D5` | 1号機 |
| `E5:G5` | 2号機 |
| `H5:J5` | 3号機 |
| `M7:M8` | 項目/装置 |
| `N7:P7` | 設定 |
| `Q7:R7` | 今日 |
| `I30:J30` | 1,2,3号機合計 |

### 2.2 「出荷実績」シート

**タイトル**: ｶﾞﾝﾏ処理品出荷数量

#### レイアウト構造

| 領域 | セル範囲 | 内容 |
|------|----------|------|
| 出荷日入力 | `C1`（ラベル）/ `D1`（入力） | ユーザーが出荷日を入力する起点セル（名前: `Syukkabi`） |
| タイトル | `E1` | "ｶﾞﾝﾏ処理品出荷数量" |
| 合計表示 | `D2:E2` | "出荷合計数" / SUM式 |
| 単位 | `J2` | "単位：箱" |
| 左テーブル | `B3:E48` | No / 会社ｺｰﾄﾞ / 会社名 / 出荷数量（名前: `Hyou1`、上位45社） |
| 右テーブル | `G3:J49` | No / 会社ｺｰﾄﾞ / 会社名 / 出荷数量（名前: `Hyou2`、46〜90位） |
| VLOOKUP式 | `N1`, `N3` | 会社名逆引き（`VLOOKUP` で `R:S` を参照） |
| データ件数 | `P1` | DataN（DB取得レコード数） |
| DB作業エリア | `N4:S801` | 会社コード・会社名ルックアップテーブル（名前: `DataTB`） |
| 出荷データ | `C4:E48, H4:J49` | 集計結果表示（名前: `SyukkaTB`） |

#### 主要セル詳細

| ✓ | セル | 名前 | 内容 | 数式/備考 |
|---|------|------|------|-----------|
| ✓ | `D1` | Syukkabi | 出荷日入力 | ユーザー入力。変更時 `Worksheet_Change` が `出荷集計()` を起動 |
| ✓ | `E2` | SyukkaSuu | 出荷合計数 | `=SUM(E4:E48)+SUM(J4:J49)` |
| ✓ | `N1` | Siki | VLOOKUP式 | `=TRIM(VLOOKUP(O1,$R$4:$S$1003,2,FALSE))`（会社名逆引き） |
| ✓ | `P1` | DataN | データ件数 | VBA から `myRecordCount` を代入 |
| | `A1` | Debug | デバッグ用 | — |

### 2.3 「稼働時間」シート

**タイトル**: 装置稼働実績

#### レイアウト構造

| 領域 | セル範囲 | 内容 |
|------|----------|------|
| タイトル | `C3:F3` | "装置稼働実績"（結合セル） |
| 期間ヘッダ | `C5:E5` | "集計期間年月日"（結合セル） |
| 期間入力 | `C6:E6` | 集計開始日 / 集計終了日 / 総時間（ヘッダ） |
| 期間値 | `C7:E7` | KaisiDay / SyuuryouDay / TotalTime |
| 装置ヘッダ | `C8:F8` | 装置 / 1号機照射時間 / 2号機コンベア運転時間 / 3号機照射時間 |
| 照射時間 | `D9:F9` | 各号機照射時間（VBA算出） |
| 停止時間 | `D10:F10` | 各号機停止時間（VBA算出） |
| 稼働率 | `D19:F19` | 各号機稼働率（数式） |
| フィルタ範囲 | `B48:E73` | `_FilterDatabase`（AutoFilter設定あり） |
| 番号列 | `B49:B77` | 1〜29（データ行番号） |

#### 主要セル詳細

| ✓ | セル | 名前 | 内容 | 数式/備考 |
|---|------|------|------|-----------|
| ✓ | `C7` | KaisiDay | 集計開始日 | ユーザー入力 |
| ✓ | `D7` | SyuuryouDay | 集計終了日 | ユーザー入力（空欄時は当日） |
| ✓ | `E7` | TotalTime | 総時間 | VBA が `(終了日 - 開始日) * 24` で算出 |
| ✓ | `D9:F9` | Kekka / SyousyaTime | 照射時間 | VBA `線源登録表示()` が算出 |
| ✓ | `D10:F10` | Kekka / SyousyaTime | 停止時間 | VBA `線源登録表示()` が算出 |
| | `D19` | — | 1号機稼働率 | `=IFERROR(D9/TotalTime*100,"")` |
| | `E19` | — | 2号機稼働率 | `=IFERROR(E9/TotalTime*100,"")` |
| | `F19` | — | 3号機稼働率 | `=IFERROR(F9/TotalTime*100,"")` |
| | `O14` | — | 補助計算 | `=O13-24` |

#### 結合セル

| 範囲 | 内容 |
|------|------|
| `C3:F3` | 装置稼働実績 |
| `C5:E5` | 集計期間年月日 |

---

## 3. 名前付き範囲一覧

| ✓ | 名前 | 参照先 | スコープ | 用途 |
|---|------|--------|----------|------|
| ✓ | Nyuuakbi | 入荷状況!`$H$3` | ブック | 入荷日入力セル |
| ✓ | Nyukabi | 入荷状況!`$P$5` | ブック | 入荷日（yyyymmdd形式） |
| ✓ | N_Hyouji | 入荷状況!`$C$8:$J$28` | ブック | 入荷表示エリア |
| ✓ | Zaiko | 入荷状況!`$T$8:$AC$307` | ブック | 在庫DBバッファ |
| ✓ | Ric2T | 入荷状況!`$N$9` | ブック | Ric2 線源設定日 |
| ✓ | Ric2V | 入荷状況!`$O$9` | ブック | Ric2 初期速度 |
| ✓ | Ric2HP | 入荷状況!`$P$9` | ブック | Ric2 初期HP/日 |
| ✓ | Ric3T | 入荷状況!`$N$10` | ブック | Ric3 線源設定日 |
| ✓ | Ric3M | 入荷状況!`$O$10` | ブック | Ric3 初期M_time |
| ✓ | Ric3PP | 入荷状況!`$P$10` | ブック | Ric3 初期PP |
| ✓ | Today2V | 入荷状況!`$Q$9` | ブック | Ric2 今日の速度（半減期補正） |
| ✓ | Today3M | 入荷状況!`$Q$10` | ブック | Ric3 今日のM_time（半減期補正） |
| ✓ | TodayHP | 入荷状況!`$R$9` | ブック | Ric2 今日のHP/日 |
| ✓ | TodayPP | 入荷状況!`$R$10` | ブック | Ric3 今日のPP/日 |
| ✓ | Syukkabi | 出荷実績!`$D$1` | ブック | 出荷日入力セル |
| ✓ | SyukkaSuu | 出荷実績!`$E$2` | ブック | 出荷合計数 |
| ✓ | Hyou1 | 出荷実績!`$B$4:$E$48` | ブック | 出荷表示テーブル（左） |
| ✓ | Hyou2 | 出荷実績!`$G$4:$J$49` | ブック | 出荷表示テーブル（右） |
| ✓ | SyukkaTB | 出荷実績!`$C$4:$E$48`,`$H$4:$J$49` | ブック | 出荷データクリア対象 |
| ✓ | DataN | 出荷実績!`$P$1` | ブック | DBレコード件数 |
| ✓ | DataTB | 出荷実績!`$N$4:$S$801` | ブック | DB作業テーブル |
| ✓ | Siki | 出荷実績!`$N$1` | ブック | VLOOKUP式（会社名逆引き） |
| | Debug | 出荷実績!`$A$1` | ブック | デバッグ用 |
| ✓ | KaisiDay | 稼働時間!`$C$7` | ブック | 集計開始日 |
| ✓ | SyuuryouDay | 稼働時間!`$D$7` | ブック | 集計終了日 |
| ✓ | TotalTime | 稼働時間!`$E$7` | ブック | 総時間 |
| ✓ | Kekka | 稼働時間!`$D$9:$F$10` | ブック | 照射/停止時間結果 |
| ✓ | SyousyaTime | 稼働時間!`$D$9:$F$10` | ブック | 照射時間（Kekka と同一範囲） |
| ✓ | Nengappi | 稼働時間!`$C$7:$D$7` | ブック | 集計開始/終了日ペア |
| | _xlnm._FilterDatabase | 稼働時間!`$B$48:$E$73` | ローカル | AutoFilter範囲 |
| | _xlnm.Database | [1]Table登録!`$F$5` | ブック | 外部参照（ExDBファイル表示.xlsm） |
| | TableName | [1]Table登録!`$C$5` | ブック | 外部参照テーブル名 |
| | TableNameDB | [1]Table登録!`$D$5` | ブック | 外部参照DBテーブル名 |
| | Souti | [1]Table登録!`$G$5` | ブック | 外部参照装置 |
| | Setumei | [1]Table登録!`$H$5` | ブック | 外部参照説明 |

---

## 4. 数式一覧

### 4.1 「入荷状況」シート

| セル | 数式 | 説明 |
|------|------|------|
| `N5` | `=TODAY()` | 本日日付 |
| `P5` | `=YEAR(H3)&RIGHT("00"&MONTH(H3),2)&RIGHT("00"&DAY(H3),2)` | 入荷日を yyyymmdd 形式に変換 |
| `Q9` | `=ROUND(Ric2V*0.5^((N5-Ric2T)/1921)+0.0004,0)` | Co-60 半減期（≒1921日）に基づく2号機の今日の処理速度 |
| `R9` | `=ROUND(Ric2HP*Today2V/Ric2V,0)` | 2号機の今日のHP/日（速度比で按分） |
| `Q10` | `=ROUND(Ric3M/(0.5^((N5-Ric3T)/1921)+0.0004),0)` | Co-60 半減期に基づく3号機の今日のM_time |
| `R10` | `=ROUND(Ric3PP*Ric3M/Today3M,0)` | 3号機の今日のPP/日 |
| `D29` | `=SUM(D8:D28)` | 1号機入荷数合計 |
| `F29` | `=SUM(F8:F28)` | 2号機入荷数合計 |
| `G29` | `=SUM(G8:G28)` | 2号機要処理時間合計 |
| `I29` | `=SUM(I8:I28)` | 3号機入荷数合計 |
| `J29` | `=SUM(J8:J28)` | 3号機要処理時間合計 |
| `I30` | `=D29+F29+I29` | 1,2,3号機入荷数総合計 |

### 4.2 「出荷実績」シート

| セル | 数式 | 説明 |
|------|------|------|
| `N1` | `=TRIM(VLOOKUP(O1,$R$4:$S$1003,2,FALSE))` | 会社コードから会社名を逆引き |
| `E2` | `=SUM(E4:E48)+SUM(J4:J49)` | 出荷合計数（左右テーブル合算） |
| `N3` | `=TRIM(VLOOKUP(O3,$R$4:$S$1003,2,FALSE))` | 会社名逆引き（コピー用テンプレート） |

### 4.3 「稼働時間」シート

| セル | 数式 | 説明 |
|------|------|------|
| `O14` | `=O13-24` | 補助計算 |
| `D19` | `=IFERROR(D9/TotalTime*100,"")` | 1号機稼働率（%） |
| `E19` | `=IFERROR(E9/TotalTime*100,"")` | 2号機稼働率（%） |
| `F19` | `=IFERROR(F9/TotalTime*100,"")` | 3号機稼働率（%） |

---

## 5. ボタン・マクロ対応

### 5.1 シート上ボタン

| ✓ | シート | ボタンテキスト | マクロ | VMLソース | 動作 |
|---|--------|--------------|--------|-----------|------|
| ✓ | 「入荷状況」 | 終了 | `[0]!終了` | vmlDrawing1.vml | ブックを閉じる（最後の1ブックならExcel終了） |
| ✓ | 「出荷実績」 | 終了 | `[0]!終了` | vmlDrawing2.vml | ブックを閉じる（最後の1ブックならExcel終了） |
| ✓ | 「稼働時間」 | 集計 | `[0]!線源登録表示` | vmlDrawing3.vml | 指定期間の装置稼働時間をDB集計・計算 |

### 5.2 フォームボタン

該当なし（ユーザーフォームは存在しない）。

### 5.3 CommandBar

VBA コード内に `CommandBar` の定義・操作は存在しない。

---

## 6. VBAモジュール仕様

### 6.0 全プロシージャ一覧

| ✓ | # | モジュール | プロシージャ名 | 種別 | 行数 | DB | 呼出元 |
|---|---|-----------|---------------|------|------|-----|--------|
| ✓ | 1 | **ThisWorkbook** | `Workbook_BeforeClose()` | Event | 4 | | （自動） |
| ✓ | 2 | **ThisWorkbook** | `Workbook_Open()` | Event | 9 | | （自動） |
| ✓ | 3 | **Sheet1** | `Worksheet_Change()` | Event | 13 | | （自動：「出荷実績」D1変更時） |
| ✓ | 4 | **Sheet2** | `Worksheet_Change()` | Event | 13 | | （自動：「入荷状況」H3変更時） |
| ✓ | 5 | **BD_Read入出荷** | `出荷履歴データ()` | Sub | 25 | ✓ | `出荷集計()` |
| ✓ | 6 | **BD_Read入出荷** | `Ric23HP_Zaiko()` | Sub | 52 | ✓ | `入荷集計()` |
| ✓ | 7 | **ExFunction** | `ExchengeDay()` | Function | 30 | | `出荷履歴データ()` |
| ✓ | 8 | **ExFunction** | `ExchengeDATE()` | Function | 18 | | （汎用、本ファイル内未使用の可能性あり） |
| ✓ | 9 | **集計処理** | `出荷集計()` | Sub | 100 | ✓ | `Sheet1.Worksheet_Change()` |
| ✓ | 10 | **集計処理** | `入荷集計()` | Sub | 165 | ✓ | `Sheet2.Worksheet_Change()` |
| ✓ | 11 | **集計終了** | `終了()` | Sub | 8 | | シートボタン「終了」 |
| ✓ | 12 | **表クリア** | `画面クリア()` | Sub | 16 | | `Workbook_Open()` |
| ✓ | 13 | **SQL_Execution** | `Open_oraconDB()` | Sub | 10 | ✓ | 各DB操作プロシージャ |
| ✓ | 14 | **SQL_Execution** | `SQL_Exe()` | Sub | 14 | ✓ | `Disp_Sheet()`, `Set_Array()`, `SQL_INSERT_UPDATE()` |
| ✓ | 15 | **SQL_Execution** | `SQL_INSERT_UPDATE()` | Sub | 38 | ✓ | （汎用：本ファイル内で直接呼出なし） |
| ✓ | 16 | **SQL_Execution** | `SQL_Delete()` | Sub | 16 | ✓ | （汎用：本ファイル内で直接呼出なし） |
| ✓ | 17 | **SQL_Execution** | `Disp_Sheet()` | Sub | 48 | ✓ | `出荷履歴データ()`, `Ric23HP_Zaiko()` |
| ✓ | 18 | **SQL_Execution** | `Set_Array()` | Sub | 28 | ✓ | `稼働状況()`, `線源登録表示()` |
| ✓ | 19 | **装置稼働状況** | `稼働状況()` | Sub | 58 | ✓ | （外部呼出用：EXメニュー等） |
| ✓ | 20 | **装置稼働状況** | `ExHenkan()` | Function | 9 | | `稼働状況()` |
| ✓ | 21 | **運転時間** | `線源登録表示()` | Sub | 166 | ✓ | シートボタン「集計」 |
| ✓ | 22 | **運転時間** | `ExHenkan()` | Function | 9 | | `線源登録表示()` |
| ✓ | 23 | **運転時間** | `ChengTime()` | Function | 3 | | `線源登録表示()` |

### 6.1 **ThisWorkbook** (ThisWorkbook.cls)

#### `Workbook_BeforeClose(Cancel As Boolean)`
- **種別**: Event
- **動作**: 保存確認ダイアログを抑制（`DisplayAlerts = False`）、変更を保存済みとしてマーク
- **呼出先**: なし

#### `Workbook_Open()`
- **種別**: Event
- **動作**:
  1. 全3シートの保護を解除 → `UserInterfaceOnly:=True` で再保護（VBA書込みを許可）
  2. `画面クリア()` を呼出してデータをリセット
  3. 「入荷状況」シートを選択（初期表示）
- **呼出先**: `画面クリア()`

### 6.2 **Sheet1** (Sheet1.cls → 「出荷実績」)

#### `Worksheet_Change(ByVal Target As Range)`
- **種別**: Event
- **トリガー条件**: 行1・列4（= `D1` = `Syukkabi`）が変更されたとき
- **動作**:
  - 空欄の場合: `SyukkaTB` をクリア
  - 値ありの場合: マウスカーソルを砂時計に変更 → `出荷集計()` を実行
- **呼出先**: `出荷集計()`

### 6.3 **Sheet2** (Sheet2.cls → 「入荷状況」)

#### `Worksheet_Change(ByVal Target As Range)`
- **種別**: Event
- **トリガー条件**: 行3・列8（= `H3` = `Nyuuakbi`）が変更されたとき
- **動作**:
  - 空欄の場合: `N_Hyouji` をクリア
  - 値ありの場合: マウスカーソルを砂時計に変更 → `入荷集計()` を実行
- **呼出先**: `入荷集計()`

### 6.4 **Sheet3** (Sheet3.cls → 「稼働時間」)

- コード本体なし（`Option Explicit` のみ）

### 6.5 **BD_Read入出荷** (BD_Read入出荷.bas)

#### `出荷履歴データ()`
- **種別**: Sub
- **動作**:
  1. `Syukkabi` を yyyymmdd に変換（`ExchengeDay()` 使用）
  2. DSN=ricdb（UID=rich）で syukar テーブルから出荷データを取得
  3. `Disp_Sheet()` で「出荷実績」シートの列15に出力
  4. DSN=ricdb（UID=ric）で tokumst テーブルから会社マスタを取得
  5. `Disp_Sheet()` で「出荷実績」シートの列18に出力
  6. `Siki` の VLOOKUP 式をデータ行分にコピー（`PasteSpecial` で数式貼付）
- **呼出先**: `ExchengeDay()`, `Disp_Sheet()`

#### `Ric23HP_Zaiko(myDataN As Single)`
- **種別**: Sub
- **動作**:
  1. DSN=ricdb（UID=ric）で ExKanriTB から Ric2/Ric3 の線源設定値を取得 → 「入荷状況」N9:P10 に出力
  2. zaiko テーブルから指定入荷日の在庫データを取得 → 「入荷状況」の Zaiko 領域に出力
  3. DSN=ricdb（UID=rich）で zaikor テーブルからも在庫データを追加取得
  4. Zaiko 領域を装置順→会社順→製品順でソート
  5. `myDataN` に合計レコード数を返す
- **呼出先**: `Disp_Sheet()`

### 6.6 **ExFunction** (ExFunction.bas)

#### `ExchengeDay(ByVal myDate, ByVal myType) As String`
- **種別**: Function
- **動作**: 日付値を指定フォーマット（"mmdd" or "yyyymmdd"）の文字列に変換
- **エラーハンドリング**: 型不一致（Err 13）時は先頭に "?" を付与

#### `ExchengeDATE(ByVal myDate, ByVal myType)`
- **種別**: Function
- **動作**: yyyymmdd/mmdd 文字列をスラッシュ区切りの日付文字列に変換

### 6.7 **集計処理** (集計処理.bas)

#### `出荷集計()`
- **種別**: Sub
- **動作**:
  1. `Syukkabi` 未入力チェック
  2. 表示エリア（`Hyou1`, `Hyou2`, `DataTB`）をクリア
  3. `出荷履歴データ()` で DB から出荷データ取得
  4. 同一会社コードの出荷数を合算（会社別集計）
  5. 出荷数の降順でバブルソート
  6. 上位90社を左右テーブルに表示（45社ずつ）、91社以降は「その他」に合算
  7. 出荷数0件の場合はメッセージ表示
- **呼出先**: `出荷履歴データ()`

#### `入荷集計()`
- **種別**: Sub
- **動作**:
  1. `Nyuuakbi` 未入力チェック
  2. Zaiko, N_Hyouji をクリア
  3. `Ric23HP_Zaiko()` で DB から在庫・線源データ取得
  4. `TodayHP` / `TodayPP` を取得（半減期補正後の処理能力）
  5. Zaiko バッファからレコードを読み取り、装置番号（1/2/3）で分類
  6. 2号機: 入荷数から HP（処理量）を計算 — `RoundUp(入荷数 / 入数 / 2, 0) * PASS / TodayHP * 24`
  7. 3号機: 入荷数から PP（処理量）を計算 — `RoundUp(入荷数 / 入数, 0) * PASS / TodayPP * 24`
  8. 装置ごとに同一会社で合算 → 処理量降順ソート
  9. 上位20社を表示、21社以降は「20番以降合計」に合算
  10. 入荷0件の場合はメッセージ表示
- **呼出先**: `Ric23HP_Zaiko()`

### 6.8 **集計終了** (集計終了.bas)

#### `終了()`
- **種別**: Sub
- **動作**: 開いているブックが1つだけなら `Application.Quit`、複数なら `ThisWorkbook.Close`
- **保存**: `DisplayAlerts = False` のため確認なし

### 6.9 **表クリア** (表クリア.bas)

#### `画面クリア()`
- **種別**: Sub
- **VB_Invoke_Func**: `e\n14`（Ctrl+Shift+E のショートカットキー）
- **動作**: 全シートのデータ範囲をクリア
  - 出荷画面: `Hyou1`, `Hyou2`, `DataTB`, `Syukkabi`
  - 入荷画面: `Zaiko`, `N_Hyouji`, `Nyuuakbi`
  - 稼働時間: `Nengappi`, `Kekka`, `TotalTime`
- **イベント制御**: クリア中は `EnableEvents = False` でイベント抑制

### 6.10 **SQL_Execution** (SQL_Execution.bas)

#### 公開変数

| 変数名 | 型 | 用途 |
|--------|-----|------|
| `mpErrDes` | String | エラーメッセージ格納 |
| `mpDSN` | String | ODBC 接続文字列 |
| `oraconn` | ADODB.Connection | DB接続オブジェクト |
| `rs` | ADODB.Recordset | レコードセット |

#### `Open_oraconDB()`
- **種別**: Sub
- **動作**: `mpDSN` の接続文字列で ADODB.Connection を Open、`CursorLocation = adUseClient`

#### `SQL_Exe(mySQL As String)`
- **種別**: Sub
- **動作**: `oraconn.Execute(mySQL)` でSQL文を実行し結果を `rs` に格納
- **エラー処理**: エラー時は `Stop`（デバッグモード突入）+ `Debug.Print`

#### `SQL_INSERT_UPDATE(myTBL, myKey, myD(), myN)`
- **種別**: Sub
- **動作**: キー一致レコードが存在すれば UPDATE、なければ INSERT を動的生成・実行
- **トランザクション**: `BeginTrans` → `CommitTrans`
- **備考**: 本ファイル内で直接呼出なし（共有モジュールとして他ブックから参照される可能性あり）

#### `SQL_Delete(myTBL, myWhere)`
- **種別**: Sub
- **動作**: `DELETE {table} {where}` を生成・実行
- **トランザクション**: `BeginTrans` → `CommitTrans`
- **備考**: 本ファイル内で直接呼出なし

#### `Disp_Sheet(mySQL, mySH, myRow, myRecordCount, myColumn, myFieldCount, myF)`
- **種別**: Sub
- **動作**:
  1. DB接続・SQL実行
  2. `myF = 1` の場合、フィールド名をヘッダ行に出力
  3. `CopyFromRecordset` でレコードセットをシートに一括出力
  4. `myRecordCount` / `myFieldCount` を返却

#### `Set_Array(mySQL, myData(), myRecordCount, myFldCount)`
- **種別**: Sub
- **動作**: SQL結果を二次元配列 `myData(レコードNo, フィールドNo)` に格納
- **用途**: 装置稼働状況・運転時間計算で使用

### 6.11 **装置稼働状況** (装置稼働状況.bas)

#### `稼働状況(myR1, myR1t, myR2, myR2t, myR3, myR3t)`
- **種別**: Sub
- **動作**: 1・2・3号機の現在の稼働状況をリアルタイム取得
  - **1号機**: sengnr1 の最新イベントから状態判定（貯蔵中/照射中/昇降中/移動照射中/PCストップ/PCスタート）
  - **2号機**: kyouj2 の最新レコード時刻と現在時刻の差が5分未満なら「照射中」、それ以外は「停止中」
  - **3号機**: sengnr3 の最新イベントのビットパターンから照射中/停止中を判定
- **引数**: 各号機の状態文字列（myR1/R2/R3）と時刻文字列（myR1t/R2t/R3t）を ByRef 返却
- **備考**: EXメニューなど外部ブックから呼出される設計

#### `ExHenkan(ByVal Number) As String`
- **種別**: Function
- **動作**: 10進数→2進数文字列変換（3号機のイベントビット解析用）

### 6.12 **運転時間** (運転時間.bas)

#### `線源登録表示()`
- **種別**: Sub
- **動作**: 指定期間の装置稼働時間を集計
  - **入力検証**: KaisiDay / SyuuryouDay の日付妥当性、6ヶ月上限チェック
  - **総時間算出**: `(終了日 - 開始日) * 24` → `TotalTime` に設定
  - **1号機**: sengnr1 のタイマー値差分 + 移動照射時間（timer < 24）を合算
  - **2号機**: kyouj2 の REALSPD = 0（停止）区間の時間を積算し停止時間を算出、照射時間 = 総時間 − 停止時間
  - **3号機**: sengnr3 の sekitime 差分から照射時間を算出（イベントビット判定あり）
  - **結果出力**: 各号機の照射時間・停止時間を `Kekka` 範囲に設定
- **DSN**: ricdbh（UID=ric）
- **呼出先**: `Set_Array()`, `ExHenkan()`

#### `ExHenkan(ByVal Number) As String`
- **種別**: Function
- **動作**: 10進数→2進数文字列変換（**装置稼働状況** モジュールと同名の重複定義）

#### `ChengTime(ByVal DTime) As Double`
- **種別**: Function
- **動作**: "yyyymmdd hhmmss" 形式の文字列を VBA の日付シリアル値に変換

---

## 7. ユーザーフォーム仕様

本ファイルにユーザーフォーム（.frm）は存在しない。
OLE ストリーム内にもフォーム定義は検出されなかった。

---

## 8. DB接続・外部連携

### 8.1 ODBC 接続

| # | DSN | UID | PWD | 使用モジュール | 備考 |
|---|-----|-----|-----|---------------|------|
| 1 | ricdb | ric | t6101 | **BD_Read入出荷**（ExKanriTB, tokumst） | — |
| 2 | ricdb | rich | t6101 | **BD_Read入出荷**（syukar, zaikor） | 2015/1/6 に DSN=ricdbh から変更 |
| 3 | ricdbh | ric | t6101 | **運転時間**（sengnr1, kyouj2, sengnr3） | 装置稼働データ用 |

- 接続方式: ADODB.Connection（ActiveX Data Objects）
- カーソル: `adUseClient`（クライアントサイドカーソル）
- DB種別: Oracle（`TO_NUMBER`, `TO_DATE`, `||` 演算子の使用から推定）

### 8.2 テーブル一覧

| ✓ | テーブル名 | 操作 | 使用プロシージャ | 用途 |
|---|-----------|------|-----------------|------|
| | syukar | SELECT | `出荷履歴データ()` | 出荷履歴（日次出荷データ） |
| | tokumst | SELECT | `出荷履歴データ()` | 得意先マスタ（会社コード⇔会社名） |
| | ExKanriTB | SELECT | `Ric23HP_Zaiko()` | 装置管理テーブル（線源設定値） |
| | zaiko | SELECT | `Ric23HP_Zaiko()` | 在庫テーブル（ricdb） |
| | zaikor | SELECT | `Ric23HP_Zaiko()` | 在庫テーブル（ricdb/rich） |
| | sengnr1 | SELECT | `稼働状況()`, `線源登録表示()` | 1号機線源制御ログ |
| | kyouj2 / KYOUJ2 | SELECT | `稼働状況()`, `線源登録表示()` | 2号機共用記録 |
| | sengnr3 | SELECT | `稼働状況()`, `線源登録表示()` | 3号機線源制御ログ |

全テーブルは SELECT のみ。`SQL_INSERT_UPDATE()` / `SQL_Delete()` は汎用プロシージャとして定義されているが、本ファイル内からの直接呼出はない。

### 8.3 SQL文一覧

#### **BD_Read入出荷** モジュール

| # | プロシージャ | SQL | DSN |
|---|------------|-----|-----|
| 1 | `出荷履歴データ()` | `SELECT TO_NUMBER(kaisyacd),TO_NUMBER(syukasu) FROM syukar WHERE syudate='{yyyymmdd}' ORDER BY kaisyacd` | ricdb/rich |
| 2 | `出荷履歴データ()` | `SELECT TO_NUMBER(kaisyacd),TRIM(coname) FROM tokumst WHERE kaisyacd<'2000' ORDER BY kaisyacd` | ricdb/ric |
| 3 | `Ric23HP_Zaiko()` | `SELECT kousinn,ricvm,hppp FROM ExKanriTB WHERE sikibetu='4' or sikibetu='5' ORDER BY sikibetu` | ricdb/ric |
| 4 | `Ric23HP_Zaiko()` | `SELECT uno,kaisyacd,sehncd,syouso,kainame,nouki,pass,nyukabi,nyukasu,incnt FROM zaiko WHERE nyukabi='{Nyukabi}'` | ricdb/ric |
| 5 | `Ric23HP_Zaiko()` | `SELECT uno,kaisyacd,sehncd,syouso,kainame,nouki,pass,nyukabi,nyukasu,incnt FROM zaikor WHERE nyukabi='{Nyukabi}'` | ricdb/rich |

#### **装置稼働状況** モジュール

| # | プロシージャ | SQL | DSN |
|---|------------|-----|-----|
| 6 | `稼働状況()` | `SELECT event,sdate,stime FROM sengnr1 WHERE sdate\|\|stime=(SELECT MAX(sdate\|\|stime) FROM sengnr1)` | （呼出元で設定） |
| 7 | `稼働状況()` | `SELECT MAX(rectime) FROM kyouj2` | （呼出元で設定） |
| 8 | `稼働状況()` | `SELECT event,sdate FROM sengnr3 WHERE sdate=(SELECT MAX(sdate) FROM sengnr3)` | （呼出元で設定） |

#### **運転時間** モジュール

| # | プロシージャ | SQL | DSN |
|---|------------|-----|-----|
| 9 | `線源登録表示()` | `SELECT MAX(sdate\|\|stime) FROM sengnr1 WHERE sdate\|\|stime<'{開始日}'` | ricdbh/ric |
| 10 | `線源登録表示()` | `SELECT MAX(sdate\|\|stime) FROM sengnr1 WHERE sdate\|\|stime<'{終了日}'` | ricdbh/ric |
| 11 | `線源登録表示()` | `SELECT sdate,stime,TO_NUMBER(timer),event FROM sengnr1 WHERE sdate\|\|stime>='{DS}' AND sdate\|\|stime<='{DE}' ORDER BY sdate\|\|stime` | ricdbh/ric |
| 12 | `線源登録表示()` | `SELECT SUM(TO_NUMBER(timer)) FROM sengnr1 WHERE TO_NUMBER(timer)<24 AND sdate\|\|stime>='{開始日}' AND sdate\|\|stime<'{終了日}'` | ricdbh/ric |
| 13 | `線源登録表示()` | `SELECT COUNT(*) FROM kyouj2 WHERE rectime>='{開始日}' AND rectime<'{終了日}'` | ricdbh/ric |
| 14 | `線源登録表示()` | `SELECT MAX(rectime) FROM KYOUJ2 WHERE rectime<'{開始日}'` | ricdbh/ric |
| 15 | `線源登録表示()` | `SELECT RECTIME,REALSPD FROM KYOUJ2 WHERE rectime>='{日時}' AND rectime<='{終了日}' ORDER BY rectime` | ricdbh/ric |
| 16 | `線源登録表示()` | `SELECT sdate,sekitime,event FROM sengnr3 WHERE sdate=(SELECT MAX(sdate) FROM sengnr3 WHERE sdate<TO_DATE('{日時}','yyyy/mm/dd hh24:mi:ss'))` | ricdbh/ric |

#### **SQL_Execution** 汎用プロシージャ

| # | プロシージャ | SQL パターン | 備考 |
|---|------------|-------------|------|
| 17 | `SQL_INSERT_UPDATE()` | `SELECT COUNT(*) FROM {table} WHERE {key}` → `INSERT INTO {table} (...)` or `UPDATE {table} SET ...` | 本ファイル内で直接呼出なし |
| 18 | `SQL_Delete()` | `DELETE {table} {where}` | 本ファイル内で直接呼出なし |

### 8.4 外部ファイル連携

| # | ファイル名 | 参照方式 | 参照先シート | 用途 |
|---|-----------|---------|-------------|------|
| 1 | ExDBファイル表示.xlsm | ExternalLink（[1]） | Table登録 | テーブル名・DB名・装置・説明の参照 |

**外部参照で使用される名前付き範囲:**

| 名前 | 参照先 |
|------|--------|
| _xlnm.Database | [1]Table登録!`$F$5` |
| TableName | [1]Table登録!`$C$5` |
| TableNameDB | [1]Table登録!`$D$5` |
| Souti | [1]Table登録!`$G$5` |
| Setumei | [1]Table登録!`$H$5` |

VBA コード内で `ExDBファイル表示.xlsm` を動的に開く処理は確認されない。名前付き範囲による静的な外部参照のみ。

---

## 9. データフロー

### 9.1 データフロー一覧

| ✓ | 方向 | 場所 | データ | 経由 | 出力先 |
|---|------|------|--------|------|--------|
| ✓ | 📊 → 🗄️ | 「入荷状況」`H3` | 入荷日 | `Sheet2.Worksheet_Change()` → `入荷集計()` → `Ric23HP_Zaiko()` | 🗄️ zaiko/zaikor → 📊「入荷状況」`C8:J28` |
| ✓ | 🗄️ → 📊 | ExKanriTB | Ric2/Ric3 線源設定値 | `Ric23HP_Zaiko()` → `Disp_Sheet()` | 📊「入荷状況」`N9:P10` |
| ✓ | 📊 → 📊 | 「入荷状況」`N9:P10` → `Q9:R10` | 今日の処理能力 | セル数式（Co-60半減期計算） | 📊「入荷状況」`Q9:R10` |
| ✓ | 📊 → 🗄️ | 「出荷実績」`D1` | 出荷日 | `Sheet1.Worksheet_Change()` → `出荷集計()` → `出荷履歴データ()` | 🗄️ syukar → 📊「出荷実績」`B4:J49` |
| ✓ | 🗄️ → 📊 | tokumst | 会社マスタ | `出荷履歴データ()` → `Disp_Sheet()` | 📊「出荷実績」`R4:S801` |
| ✓ | 📊 → 📊 | 「出荷実績」`R:S` → `N1,N3` | 会社名逆引き | VLOOKUP 式 | 📊「出荷実績」`N1:N列` |
| ✓ | 🖥️ → 🗄️ | 「稼働時間」`C7,D7` | 集計期間 | ボタン「集計」→ `線源登録表示()` | 🗄️ sengnr1/kyouj2/sengnr3 → 📊「稼働時間」`D9:F10` |
| ✓ | 📊 → 📊 | 「稼働時間」`D9:F10` → `D19:F19` | 稼働率 | セル数式 `=IFERROR(照射/TotalTime*100,"")` | 📊「稼働時間」`D19:F19` |
| ✓ | 🗄️ → 🖥️ | sengnr1/kyouj2/sengnr3 | 装置状態 | `稼働状況()` | 🖥️ 呼出元（EXメニュー等） |

### 9.2 データフローツリー図

```
📊 Ex入出荷集計.xlsm
├── 🖥️ ブック起動（Workbook_Open）
│   ├── 画面クリア() → 全データリセット
│   └── 「入荷状況」シート表示
│
├── 📊「入荷状況」
│   ├── 🖥️ ユーザー入力: H3（入荷日）
│   │   └── Sheet2.Worksheet_Change()
│   │       └── 入荷集計()
│   │           ├── Ric23HP_Zaiko()
│   │           │   ├── 🗄️ ExKanriTB → N9:P10（線源設定値）
│   │           │   ├── 🗄️ zaiko → T8:AC307（在庫データ ricdb/ric）
│   │           │   └── 🗄️ zaikor → T8:AC307（在庫データ ricdb/rich）
│   │           ├── 📊 Q9:R10 ← 半減期数式（今日の処理能力）
│   │           ├── 📊 装置別分類 → 会社集計 → ソート
│   │           └── 📊 C8:J28 ← 表示（1号機/2号機/3号機 × 上位20社）
│   └── 🖥️ ボタン「終了」→ 終了()
│
├── 📊「出荷実績」
│   ├── 🖥️ ユーザー入力: D1（出荷日）
│   │   └── Sheet1.Worksheet_Change()
│   │       └── 出荷集計()
│   │           ├── 出荷履歴データ()
│   │           │   ├── 🗄️ syukar → N4列（出荷データ ricdb/rich）
│   │           │   └── 🗄️ tokumst → R4:S列（会社マスタ ricdb/ric）
│   │           ├── 📊 N1:N列 ← VLOOKUP（会社名結合）
│   │           ├── 📊 会社集計 → 出荷数降順ソート
│   │           └── 📊 B4:J49 ← 表示（上位90社 + その他）
│   └── 🖥️ ボタン「終了」→ 終了()
│
├── 📊「稼働時間」
│   ├── 🖥️ ユーザー入力: C7（開始日）, D7（終了日）
│   │   └── 🖥️ ボタン「集計」→ 線源登録表示()
│   │       ├── 📊 E7 ← TotalTime 算出
│   │       ├── 🗄️ sengnr1 → 1号機照射時間（タイマー差分 + 移動照射）
│   │       ├── 🗄️ kyouj2 → 2号機停止時間（REALSPD=0 区間積算）
│   │       ├── 🗄️ sengnr3 → 3号機照射時間（sekitime差分 + ビット判定）
│   │       ├── 📊 D9:F9 ← 照射時間
│   │       ├── 📊 D10:F10 ← 停止時間
│   │       └── 📊 D19:F19 ← 稼働率（数式で自動計算）
│   └── 🖥️ ボタン「集計」（上記と同じ）
│
└── 📄 ExDBファイル表示.xlsm（外部参照 [1]）
    └── Table登録シート → テーブル名・DB名等の参照
```

---

## 10. セキュリティ注意事項

### 10.1 認証情報のハードコーディング

以下の接続文字列が VBA ソースにハードコードされている:

| モジュール | 接続文字列 |
|-----------|-----------|
| **BD_Read入出荷** | `DSN=ricdb;UID=rich;PWD=t6101` |
| **BD_Read入出荷** | `DSN=ricdb;UID=ric;PWD=t6101` |
| **運転時間** | `DSN=ricdbh;UID=ric;PWD=t6101` |

パスワード `t6101` が平文で埋め込まれている。ODBC DSN 設定側で認証を管理する方式への移行を推奨。

### 10.2 SQLインジェクションリスク

以下の箇所でユーザー入力値が文字列連結で SQL に埋め込まれる:

| プロシージャ | リスク箇所 |
|------------|-----------|
| `出荷履歴データ()` | `WHERE syudate='{mySyuDay}'` — `Syukkabi` セル値を `ExchengeDay()` で変換後に連結 |
| `Ric23HP_Zaiko()` | `WHERE nyukabi={myNyukaBi}` — `Nyukabi` セル値をクォートして連結 |
| `線源登録表示()` | `WHERE sdate||stime<'{myKaisiDay}'` — 集計開始日を連結 |

入力はシートセル経由であり外部からの直接入力は困難だが、バインド変数（パラメータクエリ）の使用を推奨。

### 10.3 エラーハンドリング

- `SQL_Exe()` 内でエラー時に `Stop`（デバッグモード突入）が記述されており、本番環境ではブレークが発生する可能性がある
- `On Error Resume Next` が多用されており、DB接続失敗時のエラーが握りつぶされる箇所がある

### 10.4 シート保護

- 全シートは `Protect UserInterfaceOnly:=True` で保護されているが、パスワードは設定されていない
- `Workbook_BeforeClose` で `DisplayAlerts = False` → `Saved = True` とされるため、未保存データがあっても警告なしで閉じられる

### 10.5 VBA Stomping 検出

olevba が VBA Stomping（VBA ソースコードと P-code の不一致）を検出している。意図的な改竄の可能性は低いが、コンパイル環境の差異により発生した可能性がある。P-code の整合性確認を推奨。

### 10.6 外部参照

`ExDBファイル表示.xlsm` への外部リンクが存在する。ファイルが移動・削除された場合、名前付き範囲（`TableName`, `TableNameDB` 等）の参照が `#REF!` エラーとなる。
