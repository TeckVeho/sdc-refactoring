# Ex照射実績表示2 仕様書

| 項目 | 内容 |
|------|------|
| 対象ファイル | Ex照射実績表示2.xlsm |
| 解析日 | 2026-06-29 |
| マクロ | あり（VBA） |
| ユーザーフォーム | なし |

---


> 本仕様書の表記ルール・章構成は [_仕様書共通ルール.md](./_仕様書共通ルール.md) を参照。

## 目次

1. [ファイル構成](#1-ファイル構成)
2. [シート詳細](#2-シート詳細)
3. [名前付き範囲一覧](#3-名前付き範囲一覧)
4. [数式一覧](#4-数式一覧)
5. [ボタン・マクロ対応](#5-ボタンマクロ対応)
6. [VBAモジュール仕様](#6-vbaモジュール仕様)
7. [ユーザーフォーム仕様](#7-ユーザーフォーム仕様)
8. [DB接続・外部連携](#8-db接続外部連携)
9. [データフロー](#9-データフロー)
10. [セキュリティ注意事項](#10-セキュリティ注意事項)

---

## 1. ファイル構成

### 1.1 シート一覧

| # | シート名 | codeName | 可視性 | 用途 | ✓ |
|---|----------|----------|--------|------|---|
| 1 | 「実績表示」 | Sheet2 | 表示 | 受付番号を入力し、照射実績を一覧表示する操作画面 | ✓ |
| 2 | 「実績」 | Sheet1 | 非表示 | DBから取得した在庫・実績データの一時格納領域 | |
| 3 | 「社員」 | Sheet3 | 非表示 | DBから取得した社員マスタの一時格納領域 | |

### 1.2 ユーザーフォーム一覧

ユーザーフォームは存在しない。

### 1.3 VBAモジュール一覧

| # | モジュール名 | 種別 | 行数 | 用途 | ✓ |
|---|-------------|------|------|------|---|
| 1 | **ThisWorkbook** | ThisWorkbook | 13 | ブック開閉時の初期化・終了処理 | ✓ |
| 2 | **Sheet1** | Sheet（「実績」） | 2 | 宣言のみ（未使用） | |
| 3 | **Sheet2** | Sheet（「実績表示」） | 15 | `Uno`セル変更イベントで検索起動 | ✓ |
| 4 | **Sheet3** | Sheet（「社員」） | 2 | 宣言のみ（未使用） | |
| 5 | **Utility** | 標準モジュール | 4 | イベント有効化ユーティリティ | |
| 6 | **画面クリア** | 標準モジュール | 71 | 画面・データ領域の初期化 | ✓ |
| 7 | **起動** | 標準モジュール | 23 | 検索メイン処理の制御 | ✓ |
| 8 | **表示** | 標準モジュール | 149 | 在庫・実績データの画面表示 | ✓ |
| 9 | **終了処理** | 標準モジュール | 11 | ブックを閉じる処理 | |
| 10 | **SQL_Execution** | 標準モジュール | 182 | DB接続・SQL実行・シート/配列格納 | ✓ |
| 11 | **在庫実績Read** | 標準モジュール | 111 | 在庫・実績データのSQL組み立てとDB読み込み | ✓ |
| 12 | **社員データRead** | 標準モジュール | 20 | 社員マスタのDB読み込み | ✓ |
| 13 | **共通変数** | 標準モジュール | 5 | Public変数宣言 | ✓ |

---

## 2. シート詳細

### 2.0 シート可視性一覧

| シート名 | 状態 | VBA制御 |
|----------|------|---------|
| 「実績表示」 | 表示（Visible） | `Protect UserInterfaceOnly:=True` でシート保護（VBA操作は許可） |
| 「実績」 | 非表示（Hidden） | VBAがDBデータを書き込むワーク領域 |
| 「社員」 | 非表示（Hidden） | VBAが社員マスタを書き込むワーク領域 |


### 2.0b 非表示行・列一覧

| シート | 非表示行 | 非表示列 |
|---|---|---|
| 「実績表示」 | 1, 6 | A |

### 2.1 「実績表示」シート（codeName: Sheet2）

ユーザーが直接操作する画面。受付番号を入力すると在庫情報と照射実績が自動表示される。

#### レイアウト構成

**ヘッダ部（行2〜5）：在庫情報表示**

| セル | 名前付き範囲 | 内容 | VBA参照 | ✓ |
|------|-------------|------|---------|---|
| `B2` | ラベル | "受付番号" | | |
| `C2` | `Kaisyacd` | 会社コード | `在庫表示()` で代入 | ✓ |
| `D2` | `Kainame` | 会社名 | `在庫表示()` で代入 | ✓ |
| `G2` | `Sehncd` | 製品コード | `在庫表示()` で代入 | ✓ |
| `H2` | `SName` | 製品名（指定線名） | `在庫表示()` で代入 | ✓ |
| `L2` | ラベル | "指定線量kGy" | | |
| `N2` | `Siteisn` | 指定線量 | `在庫表示()` で代入 | ✓ |
| `B3` | `Uno` | 受付番号（ユーザー入力セル） | `Worksheet_Change` トリガー | ✓ |
| `C3` | ラベル | "受付数量" | | |
| `D3` | `Nyukasu` | 受付数量 | `在庫表示()` で代入 | ✓ |
| `E3` | ラベル | "受付日" | | |
| `F3` | `Koudate` | 受付日 | `在庫表示()` で代入 | ✓ |
| `H3` | ラベル | "分類" | | |
| `I3` | `Tani` | 分類（一般品等） | `在庫表示()` で代入 | ✓ |
| `J3` | ラベル | "受付" | | |
| `K3` | `Kousncd` | 受付者名 | `在庫表示()` で代入 | ✓ |
| `M3` | ラベル | 荷姿表示（"Box入数"/"PL積載数"） | `在庫表示()` で条件代入 | ✓ |
| `N3` | `Incnt` | 入数 | `在庫表示()` で代入 | ✓ |
| `E4` | `Kagensn` | 管理点線量 下限 | `在庫表示()` で代入 | ✓ |
| `G4` | `Jyougsn` | 管理点線量 上限 | `在庫表示()` で代入 | ✓ |
| `I4` | `Labelcd` | ラベルコード | `在庫表示()` で代入 | ✓ |
| `K4` | `Pass` | 指定パス数 | `在庫表示()` で代入 | ✓ |
| `M4` | `Syouso` | 装置番号 | `在庫表示()` で代入 | ✓ |
| `D5` | `Syouzusu` | 照射済数 | `在庫表示()` で代入 | ✓ |
| `E5` | ラベル | "線量検査合格数" | | |
| `G5` | `Senkssu` | 線量検査合格数 | `在庫表示()` で代入 | ✓ |
| `I5` | `Syukasu` | 出荷済数 | `在庫表示()` で代入 | ✓ |
| `K5` | `Syukabi` | 出荷日 | `在庫表示()` で代入 | ✓ |
| `N5` | `Syukacd` | 出荷者コード | `在庫表示()` で代入 | ✓ |

**実績データ部（行7〜31）**

| 行 | 項目 | 内容 |
|----|------|------|
| 7 | RicNo | 照射番号（FRicNo / LRicNo） |
| 8 | 照射 | 通常照射 / 追加照射 |
| 9 | 照射数量 | 照射数量 |
| 10 | 開始日 | mm/dd形式 |
| 11 | 投入時刻 | dd hh:mm形式 |
| 12 | 載荷時刻 | dd hh:mm形式 |
| 13 | ハンガー番号 | ハンガー番号 |
| 14 | シャフリング | dd hh:mm形式（装置3の場合"--"） |
| 15 | 終了日 | mm/dd形式 |
| 16 | 脱荷時刻 | dd hh:mm形式 |
| 17 | 測定日 | mm/dd形式 |
| 18 | 線量計番号 | 線量計番号 |
| 19 | 線量計種類 | 線量計種類 |
| 20 | 測定器記号 | 測定器記号 |
| 21 | 素子厚(mm) | 素子厚 |
| 22 | ABS | ABS値 |
| 23 | 計算式 | 計算式 |
| 24 | 温度（℃） | 温度 |
| 25 | 測定線量(kGy) | 測定値（小数第1位切り捨て） |
| 26 | 出荷判定日 | mm/dd形式 |
| 27 | 計画者 | 作業者名 |
| 28 | 投入者 | 作業者名 |
| 29 | 測定者 | 作業者名 |
| 30 | 出荷判定者 | 作業者名 |
| 31 | 実パス数 | 実パス数 |

実績データは複数レコード（照射回数分）が列方向に並ぶ。各レコードは2列幅（C:D, E:F, G:H, ...）で配置される。

**パス時刻部（行32〜131）**

| セル範囲 | 名前付き範囲 | 内容 |
|----------|-------------|------|
| `B34:B131` | `PassKai` | パス回数番号（1, 2, 3, ...） |
| `C7:FZ131` | `JissekiTB` | 実績データ表示領域全体 |

行32にヘッダ「入室」「退室」、行33に「日時分」を表示。行34以降に各パスの入室・退室時刻を dd hh:mm 形式で表示。

### 2.2 「実績」シート（codeName: Sheet1）

DBから取得したデータの一時格納シート。ユーザーには非表示。

| セル範囲 | 名前付き範囲 | 内容 |
|----------|-------------|------|
| `A1` | — | "在庫製品" / "履歴製品" / "在庫／履歴共データ無" |
| `A3:T4` | `Zaiko` | 在庫データ（DB→シート格納領域、行3ヘッダ・行4データ） |
| `A6:U6` | — | 在庫データの加工行（数式で変換） |
| `A10:AD110` | `Jisseki` | 実績データ（DB→シート格納領域、行10ヘッダ・行11以降データ） |
| `U6` | `SeiName` | 製品名（数式参照用） |
| `AE6:AH110` | — | 作業者名変換（社員コード→氏名のVLOOKUP） |

#### 数式

| セル | 数式 | 用途 |
|------|------|------|
| `A6` | `=A4` | 在庫データ転記 |
| `B6` | `=B4` | 在庫データ転記 |
| `C6` | `=TRIM(C4)` | 会社名のトリム |
| `D6`〜`F6` | `=D4` 〜 `=F4` | 在庫データ転記 |
| `G6` | `=LEFT(G4,4)&"/"&MID(G4,5,2)&"/"&RIGHT(G4,2)` | 受付日の日付変換（YYYYMMDD→YYYY/MM/DD） |
| `H6` | `=VLOOKUP(H4,Bunnrui,2,FALSE)` | 分類コード→分類名変換 |
| `I6` | `=VLOOKUP(VALUE(I4),Syainn,2,FALSE)` | 受付者コード→氏名変換 |
| `J6`〜`U6` | `=J4` 〜 `=U4` | 在庫データ転記 |
| `S6` | `=LEFT(S4,4)&"/"&MID(S4,5,2)&"/"&RIGHT(S4,2)` | 出荷日の日付変換 |
| `T6` | `=IF(ISERROR(VLOOKUP(VALUE(T4),Syainn,2,FALSE)),"*",VLOOKUP(VALUE(T4),Syainn,2,FALSE))` | 出荷者コード→氏名変換（エラー時"*"） |
| `AE6:AH6` | `=IF(TRIM(X6)="0","---",TRIM(VLOOKUP(VALUE(X6),Syainn,2,FALSE)))` | 作業者コード→氏名変換（0の場合"---"） |
| `AE11:AH110` | 同上パターン（行11〜110に連番展開） | 実績データの作業者名変換 |

### 2.3 「社員」シート（codeName: Sheet3）

| セル範囲 | 名前付き範囲 | 内容 |
|----------|-------------|------|
| `B2:C110` | `Syainn` | 社員マスタ（B列:社員番号, C列:氏名） |
| `D1` | `ReadDate` | 読み込み日時 |
| `E3:F7` | `Bunnrui` | 分類マスタ（E列:コード, F列:名称） |

**分類マスタ（`Bunnrui`）内容：**

| コード | 分類名 |
|--------|--------|
| 1 | 医療機器 |
| 2 | 一般品 |
| 3 | 医薬品 |
| 4 | 試験品 |

---

## 3. 名前付き範囲一覧

| # | 名前 | 参照先 | 用途 | VBA参照 |
|---|------|--------|------|---------|
| 1 | `Bunnrui` | 社員!$E$3:$F$7 | 分類コード→名称変換テーブル | 数式から参照 |
| 2 | `Incnt` | 実績表示!$N$3 | 入数 | `GamenCls1()` `在庫表示()` |
| 3 | `Jisseki` | 実績!$A$10:$AD$110 | 実績データ格納領域 | `GamenCls1()` |
| 4 | `JissekiTB` | 実績表示!$C$7:$FZ$131 | 実績表示領域 | `GamenCls1()` `GamennCls2()` |
| 5 | `Jyougsn` | 実績表示!$G$4 | 管理点線量 上限 | `GamenCls1()` `GamennCls2()` `在庫表示()` |
| 6 | `Kagensn` | 実績表示!$E$4 | 管理点線量 下限 | `GamenCls1()` `GamennCls2()` `在庫表示()` |
| 7 | `Kainame` | 実績表示!$D$2 | 会社名 | `GamenCls1()` `GamennCls2()` `在庫表示()` |
| 8 | `Kaisyacd` | 実績表示!$C$2 | 会社コード | `GamenCls1()` `GamennCls2()` `在庫表示()` |
| 9 | `Koudate` | 実績表示!$F$3 | 受付日 | `GamenCls1()` `GamennCls2()` `在庫表示()` |
| 10 | `Kousncd` | 実績表示!$K$3 | 受付者コード | `GamenCls1()` `GamennCls2()` `在庫表示()` |
| 11 | `Labelcd` | 実績表示!$I$4 | ラベルコード | `GamenCls1()` `GamennCls2()` `在庫表示()` |
| 12 | `Nyukasu` | 実績表示!$D$3 | 受付数量 | `GamenCls1()` `GamennCls2()` `在庫表示()` |
| 13 | `Pass` | 実績表示!$K$4 | 指定パス数 | `GamenCls1()` `GamennCls2()` `在庫表示()` `実績表示()` |
| 14 | `PassKai` | 実績表示!$B$34:$B$131 | パス回数表示列 | `GamenCls1()` `GamennCls2()` |
| 15 | `ReadDate` | 社員!$D$1 | 社員データ読み込み日時 | `GamenCls1()` |
| 16 | `Sehncd` | 実績表示!$G$2 | 製品コード | `GamenCls1()` `GamennCls2()` `在庫表示()` |
| 17 | `SeiName` | 実績!$U$6 | 製品名（数式用） | 数式から参照 |
| 18 | `Senkssu` | 実績表示!$G$5 | 線量検査合格数 | `GamenCls1()` `GamennCls2()` `在庫表示()` |
| 19 | `Siteisn` | 実績表示!$N$2 | 指定線量 | `GamenCls1()` `GamennCls2()` `在庫表示()` |
| 20 | `SName` | 実績表示!$H$2 | 製品名表示 | `GamenCls1()` `GamennCls2()` `在庫表示()` |
| 21 | `Syainn` | 社員!$B$2:$C$110 | 社員マスタテーブル | `GamenCls1()` 、数式VLOOKUP |
| 22 | `Syouso` | 実績表示!$M$4 | 装置番号 | `GamenCls1()` `GamennCls2()` `在庫表示()` `実績表示()` |
| 23 | `Syouzusu` | 実績表示!$D$5 | 照射済数 | `GamenCls1()` `GamennCls2()` `在庫表示()` |
| 24 | `Syukabi` | 実績表示!$K$5 | 出荷日 | `GamenCls1()` `GamennCls2()` `在庫表示()` |
| 25 | `Syukacd` | 実績表示!$N$5 | 出荷者コード | `GamenCls1()` `GamennCls2()` `在庫表示()` |
| 26 | `Syukasu` | 実績表示!$I$5 | 出荷済数 | `GamenCls1()` `GamennCls2()` `在庫表示()` |
| 27 | `Tani` | 実績表示!$I$3 | 分類 | `GamenCls1()` `GamennCls2()` `在庫表示()` |
| 28 | `Uno` | 実績表示!$B$3 | 受付番号（入力セル） | `GamenCls1()` `Worksheet_Change` `在庫表示()` `在庫実績データ()` |
| 29 | `Zaiko` | 実績!$A$3:$T$4 | 在庫データ格納領域 | `GamenCls1()` |

### 3.1 データの入力規則（「実績表示」シート）

| セル | 名前付き範囲 | 種別 | 制約 | 用途 |
|------|-------------|------|------|------|
| `B3` | `Uno` | 整数 | > 1990010000 | 受付番号（10桁。入力値がこの範囲外の場合エラー） |
| `C3:D3`, `C5:D5`, `E3:N5` | — | 整数 | =9999（固定値のみ許可） | 内部制御用セル |

---

## 4. 数式一覧

数式は「実績」シートにのみ存在する。

### 「実績」シート

| セル | 数式 | 用途 |
|------|------|------|
| `A6` | `=A4` | 在庫UNO転記 |
| `B6` | `=B4` | 会社コード転記 |
| `C6` | `=TRIM(C4)` | 会社名（前後空白除去） |
| `D6` | `=D4` | 製品コード転記 |
| `E6` | `=E4` | 指定線名転記 |
| `F6` | `=F4` | 受付数量転記 |
| `G6` | `=LEFT(G4,4)&"/"&MID(G4,5,2)&"/"&RIGHT(G4,2)` | 受付日（YYYYMMDD→YYYY/MM/DD） |
| `H6` | `=VLOOKUP(H4,Bunnrui,2,FALSE)` | 分類コード→分類名 |
| `I6` | `=VLOOKUP(VALUE(I4),Syainn,2,FALSE)` | 受付者コード→氏名 |
| `J6`〜`R6` | `=J4` 〜 `=R4` | 在庫データ転記 |
| `S6` | `=LEFT(S4,4)&"/"&MID(S4,5,2)&"/"&RIGHT(S4,2)` | 出荷日（YYYYMMDD→YYYY/MM/DD） |
| `T6` | `=IF(ISERROR(VLOOKUP(VALUE(T4),Syainn,2,FALSE)),"*",VLOOKUP(VALUE(T4),Syainn,2,FALSE))` | 出荷者コード→氏名（エラー時"*"） |
| `U6` | `=U4` | 製品名転記（`SeiName`） |
| `AE6:AH6` | `=IF(TRIM(Xn)="0","---",TRIM(VLOOKUP(VALUE(Xn),Syainn,2,FALSE)))` | 作業者コード→氏名（列X〜AA, 0は"---"） |
| `AE11:AH110` | 同上パターン（100行分展開） | 実績データの作業者名一括変換 |

---

## 5. ボタン・マクロ対応

### 5.1 シート上のボタン

| # | シート | ボタンテキスト | 割り当てマクロ | 動作 | ✓ |
|---|--------|---------------|---------------|------|---|
| 1 | 「実績表示」 | 終了 | `Bookを閉じる()` | ブックを閉じる（上書き保存なし） | |

### 5.2 フォーム上のボタン

フォームは存在しない。

### 5.x ショートカットキー

| マクロ名 | ショートカット | 備考 |
|---|---|---|
| `GamenCls1()` | **Ctrl+E\N14** | |

### 5.3 CommandBar

CommandBarの定義は存在しない。

---

## 6. VBAモジュール仕様

### 6.0 全プロシージャ一覧

| # | モジュール | プロシージャ | 種別 | 行数 | 用途 | ✓ |
|---|-----------|-------------|------|------|------|---|
| 1 | **ThisWorkbook** | `Workbook_BeforeClose()` | Private Sub | 3 | 保存確認なしでブックを閉じる | |
| 2 | **ThisWorkbook** | `Workbook_Open()` | Private Sub | 7 | 📊シート保護設定 → `GamenCls1()` → `社員データ()` 呼び出し | ✓ |
| 3 | **Sheet2** | `Worksheet_Change()` | Private Sub | 12 | 📊`Uno`セル変更時、空なら画面クリア / 値ありなら`起動Main()` | ✓ |
| 4 | **Utility** | `aaaaa()` | Sub | 3 | イベント有効化（デバッグ用） | |
| 5 | **画面クリア** | `GamenCls1()` | Sub | 37 | 🖥️全名前付き範囲を空文字にクリア、PrintAreaリセット | ✓ |
| 6 | **画面クリア** | `GamennCls2()` | Sub | 30 | 🖥️実績表示領域とヘッダ（Uno以外）をクリア | ✓ |
| 7 | **起動** | `起動Main()` | Sub | 21 | 📄メイン制御：`社員データ()` → `在庫実績データ()` → `在庫表示()` → `実績表示()` | ✓ |
| 8 | **表示** | `在庫表示()` | Sub | 48 | 🖥️「実績」シートから在庫データを読み取り「実績表示」シートに埋め込み | ✓ |
| 9 | **表示** | `実績表示()` | Sub | 96 | 🖥️「実績」シートから実績データを読み取り「実績表示」シートに一覧表示 | ✓ |
| 10 | **終了処理** | `Bookを閉じる()` | Sub | 8 | 🖥️メッセージ非表示で閉じる（最後の1ブックならExcel終了） | |
| 11 | **SQL_Execution** | `Open_oraconDB()` | Sub | 9 | 🗄️ADODB.Connectionを使ったODBC接続 | ✓ |
| 12 | **SQL_Execution** | `SQL_Exe()` | Sub | 11 | 🗄️SQL文実行（Execute） | ✓ |
| 13 | **SQL_Execution** | `SQL_INSERT_UPDATE()` | Sub | 34 | 🗄️INSERT/UPDATE汎用処理（キー存在チェック付き） | ✓ |
| 14 | **SQL_Execution** | `SQL_Delete()` | Sub | 14 | 🗄️DELETE汎用処理 | ✓ |
| 15 | **SQL_Execution** | `Disp_Sheet()` | Sub | 43 | 🗄️SQL実行結果をシートに貼り付け | ✓ |
| 16 | **SQL_Execution** | `Set_Array()` | Sub | 27 | 🗄️SQL実行結果を配列に格納 | ✓ |
| 17 | **在庫実績Read** | `在庫実績データ()` | Sub | 107 | 🗄️在庫・実績データのSQL構築とDB取得 | ✓ |
| 18 | **社員データRead** | `社員データ()` | Sub | 16 | 🗄️社員マスタのDB取得 | ✓ |

### 6.1 **ThisWorkbook**（ThisWorkbook.cls）

#### `Workbook_BeforeClose(Cancel As Boolean)`
- 保存確認ダイアログを抑制し、未保存状態でブックを閉じる

#### `Workbook_Open()`
1. 📊「実績表示」シートの保護を一旦解除
2. 📊`UserInterfaceOnly:=True` で再保護（VBA操作は許可、手動操作はロック）
3. 📊カーソル移動を非保護セルに制限
4. `GamenCls1()` を呼び出し → 画面初期化
5. `社員データ()` を呼び出し → 社員マスタDB読み込み

### 6.2 **Sheet1**（Sheet1.cls / 「実績」シート）

`Option Explicit` のみ。イベント処理なし。

### 6.3 **Sheet2**（Sheet2.cls / 「実績表示」シート）

#### `Worksheet_Change(ByVal Target As Range)`
- トリガー条件：`B3`（=`Uno`）セルの変更
- 📊`Uno` が空 → `GamennCls2()` で表示クリア
- 📊`Uno` に値あり → `起動Main()` で検索実行
- イベント二重発火防止（`EnableEvents = False/True`）

### 6.4 **Sheet3**（Sheet3.cls / 「社員」シート）

`Option Explicit` のみ。イベント処理なし。

### 6.5 **Utility**（Utility.bas）

#### `aaaaa()`
- `Application.EnableEvents = True` でイベントを有効化
- デバッグ・復旧用ユーティリティ

### 6.6 **画面クリア**（画面クリア.bas）

#### `GamenCls1()`
- 🖥️全データ領域を初期化する完全クリア
- 対象：`Zaiko`, `Jisseki`, `Syainn`, `ReadDate`, `JissekiTB`, `PassKai`
- 対象：`Uno`, `Kaisyacd`, `Kainame`, `Sehncd`, `Siteisn`, `Nyukasu`, `Koudate`, `Tani`, `Kousncd`, `Incnt`, `Kagensn`, `Jyougsn`, `Labelcd`, `Pass`, `Syouso`, `Syouzusu`, `Senkssu`, `Syukasu`, `Syukabi`, `Syukacd`, `SName`
- 📊カーソルを`Uno`に移動、PrintAreaをリセット

#### `GamennCls2()`
- 🖥️表示領域のみクリア（`Uno`は保持）
- 対象：`JissekiTB`, `PassKai`, および`Uno`以外の全ヘッダセル

### 6.7 **起動**（起動.bas）

#### `起動Main()`
- 📄メイン制御プロシージャ
1. イベント無効化・画面更新停止
2. `社員データ()` 呼び出し
3. `在庫実績データ()` 呼び出し
4. エラーがあれば `MsgBox` 表示 → `GamennCls2()` でクリア
5. エラーなし → `GamennCls2()` → `在庫表示()` → `実績表示()` で画面再構築
6. 📊「実績表示」シートを選択、`Uno`にカーソル移動
7. イベント有効化・画面更新再開

### 6.8 **表示**（表示.bas）

#### `在庫表示()`
- 🖥️「実績」シート行6（加工済み在庫データ）から21項目を読み取り
- 「実績表示」シートの名前付き範囲に1項目ずつ代入
- 受付番号が0の場合はエラーメッセージ設定で中断
- 装置番号（`Syouso`）に応じてM3セルのラベルを切替：
  - 2 → "Box入数"
  - 3 → "PL積載数"
- Public変数 `mpPass` に指定パス数を格納

#### `実績表示()`
- 🖥️「実績」シートの実績データ（行11以降）を読み取り
- 各照射回に対しFRicNo〜出荷判定日（21項目）、作業者名（4名）、実パス数を取得
- 日時フィールドは `ddhhmmss` 形式から `dd hh:mm` 形式に変換
- `tuikaflg` により "通常照射" / "追加照射" を判定
- 装置3の場合シャフリング時刻を "--" 表示
- 測定値は `RoundDown` で小数第1位に切り捨て
- パス時刻は入室・退室を列ペアで表示（dd hh:mm形式）
- 2024-06-12改修：パス配列サイズを `mpPass + 5` に拡大（TS）

### 6.9 **終了処理**（終了処理.bas）

#### `Bookを閉じる()`
- 🖥️`DisplayAlerts = False` でメッセージ非表示
- ブックが1つだけ → `Application.Quit`（Excel終了）
- 複数ブック → `ActiveWorkbook.Close`（当該ブックのみ閉じる）

### 6.10 **SQL_Execution**（SQL_Execution.bas）

Public変数：
- `mpErrDes As String` — エラーメッセージ格納用
- `mpDSN As String` — ODBC接続文字列

モジュールレベル変数：
- `oraconn As New ADODB.Connection` — ADOコネクション
- `rs As ADODB.Recordset` — レコードセット

#### `Open_oraconDB()`
- 🗄️`mpDSN`に設定された接続文字列でADODB接続を開く
- `CursorLocation = adUseClient`（クライアントサイドカーソル）

#### `SQL_Exe(ByVal mySQL)`
- 🗄️`oraconn.Execute` でSQL文を実行
- エラー発生時は `mpErrDes` にエラーメッセージを格納、`Debug.Print` で出力、`Stop` で中断

#### `SQL_INSERT_UPDATE(ByVal myTBL, ByVal myKey, myD() As Variant, myN As Single)`
- 🗄️キー条件でSELECT COUNT(*)を実行
- 0件 → INSERT文を動的構築して実行
- 1件以上 → UPDATE文を動的構築して実行
- トランザクション制御（BeginTrans / CommitTrans）

#### `SQL_Delete(myTBL As String, ByVal myWhere)`
- 🗄️WHERE条件付きDELETE文を実行
- 空のWHERE条件では何もしない（安全対策）
- トランザクション制御あり

#### `Disp_Sheet(ByVal mySQL, ByVal mySH, ByVal myROW, myRecordCount As Single, ByVal myColumn, myFieldCount As Single, ByVal myF)`
- 🗄️SQL実行 → 結果をシートに貼り付け
- `myF = 1` の場合フィールド名をヘッダ行に出力
- `CopyFromRecordset` で一括転記
- レコード数・フィールド数を参照パラメータで返却

#### `Set_Array(ByVal mySQL, myData(), myRecordCount As Single, myFldCount As Single)`
- 🗄️SQL実行 → 結果を2次元配列 `myData(i, j)` に格納
- i: レコード番号、j: フィールド番号

### 6.11 **在庫実績Read**（在庫実績Read.bas）

#### `在庫実績データ()`
- 🗄️受付番号（`Uno`）をキーに在庫・実績データをDBから取得
- 接続先：`DSN=ricdb;UID=ric;PWD=t6101`

**処理フロー：**
1. 📊表計算を手動モードに切替（速度向上）
2. 🗄️装置判定：`SELECT TO_NUMBER(syouso) FROM zaiko WHERE uno='...'`
   - 装置1の場合 → "1号機データです。表示できません" でエラー終了
3. 🗄️在庫履歴の装置判定：`SELECT TO_NUMBER(syouso) FROM rich.zaikor WHERE uno='...'`
   - 装置1の場合 → 同上エラー
4. 🗄️在庫データ取得（① `RIC.ZAIKO` + `RIC.SEHMST`）
   - 21カラム（UNO〜製品名）をJOINで取得
   - 結果0件の場合 → ② 履歴在庫（`RICH.ZAIKOR` + `RIC.SEHMST`）にフォールバック
   - 「実績」シートA1に "在庫製品" / "履歴製品" を表示
5. 🗄️実績データ取得（③ `RIC.SYOUJ2` または `RICH.SYOUJR2`）
   - 29カラム（照射番号〜退室時刻）を取得
   - `ORDER BY fricno, senkno, tuikaflg`
   - レコード数を `mpDataCount` に格納
6. 📊表計算を自動モードに復帰

### 6.12 **社員データRead**（社員データRead.bas）

#### `社員データ()`
- 🗄️社員マスタを全件取得
- 接続先：`DSN=ricdb;UID=ric;PWD=t6101`
- SQL：`SELECT TO_NUMBER(shano), REPLACE(TRIM(shaname),'　','') FROM shainmst ORDER BY shano`
- 「社員」シートの`Syainn`範囲に格納

### 6.13 **共通変数**（共通変数.bas）

| 変数名 | 型 | 用途 |
|--------|-----|------|
| `mpDataCount` | Single | 実績レコード数 |
| `mpPass` | Integer | 指定パス数 |

---

## 7. ユーザーフォーム仕様

ユーザーフォームは存在しない。

---

## 8. DB接続・外部連携

### 8.1 ODBC接続

| 項目 | 値 |
|------|-----|
| DSN | `ricdb` |
| UID | `ric` |
| PWD | `t6101` |
| 方式 | ADODB.Connection（ADO + ODBC） |
| カーソル | クライアントサイド（`adUseClient`） |
| 設定箇所 | **在庫実績Read** `在庫実績データ()` / **社員データRead** `社員データ()` |

### 8.2 テーブル一覧

| # | スキーマ | テーブル名 | 用途 | 操作種別 | ✓ |
|---|---------|-----------|------|----------|---|
| 1 | RIC | `ZAIKO` | 在庫データ（製品・受付情報） | SELECT | |
| 2 | RIC | `SEHMST` | 製品マスタ（製品名取得用） | SELECT（JOIN） | |
| 3 | RIC | `SYOUJ2` | 照射実績データ（現行） | SELECT | |
| 4 | RICH | `ZAIKOR` | 在庫履歴データ（過去分） | SELECT | |
| 5 | RICH | `SYOUJR2` | 照射実績履歴データ（過去分） | SELECT | |
| 6 | — | `SHAINMST` | 社員マスタ | SELECT | |

※ このファイルではSELECT（参照）のみ。INSERT/UPDATE/DELETEは `SQL_INSERT_UPDATE()` / `SQL_Delete()` として汎用関数が実装されているが、本ファイル内では呼び出されていない。

### 8.3 SQL一覧

| # | 発行元 | SQL | 用途 |
|---|--------|-----|------|
| 1 | `在庫実績データ()` | `SELECT TO_NUMBER(syouso) FROM zaiko WHERE uno='...'` | 装置番号判定（1号機チェック） |
| 2 | `在庫実績データ()` | `SELECT TO_NUMBER(syouso) FROM rich.zaikor WHERE uno='...'` | 履歴の装置番号判定 |
| 3 | `在庫実績データ()` | `SELECT z.uno, z.kaisyacd, z.kainame, z.sehncd, z.siteisn, z.nyukasu, z.nyukabi, TO_NUMBER(z.tani), z.kousncd, z.incnt, z.kagensn, z.jyougsn, z.labelcd, z.pass, z.syouso, z.syouzusu, z.senkssu, z.syukasu, z.syukabi, z.syukacd, TRIM(h.seiname) FROM RIC.ZAIKO z, ric.sehmst h WHERE uno='...' AND z.kaisyacd=h.kaisyacd AND z.sehncd=h.sehncd` | 在庫データ取得 |
| 4 | `在庫実績データ()` | 同上（FROM句を `RICH.ZAIKOR z, ric.sehmst h` に変更） | 在庫履歴データ取得 |
| 5 | `在庫実績データ()` | `SELECT uno, fricno, lricno, tuikaflg, suryou*1, kmmdd, tonyutime, saikatime, hangno, syafutime, smmdd, dattktime, sesdate, senkno, senksyu, toridcd, atusa*1, sokutti*1, keisask, ondok, sokutsn*1, sezhnsu, jituno*1, keikacd, tonyucd, sokutcd, syuhncd, nyutime, taitime FROM RIC.SYOUJ2 WHERE uno='...' ORDER BY fricno, senkno, tuikaflg` | 実績データ取得 |
| 6 | `在庫実績データ()` | 同上（FROM句を `RICH.SYOUJR2` に変更） | 実績履歴データ取得 |
| 7 | `社員データ()` | `SELECT TO_NUMBER(shano), REPLACE(TRIM(shaname),'　','') FROM shainmst ORDER BY shano` | 社員マスタ全件取得 |

### 8.4 外部ファイル参照

外部ファイルへの参照は存在しない。

---

## 9. データフロー

### 9.1 データフローテーブル

| # | ステップ | 場所 | データソース | データ先 | 処理内容 |
|---|---------|------|-------------|---------|---------|
| 1 | ブック起動 | 📄`Workbook_Open` | — | — | シート保護設定 |
| 2 | 画面初期化 | 🖥️`GamenCls1` | — | 「実績表示」「実績」「社員」全範囲 | 全名前付き範囲をクリア |
| 3 | 社員マスタ読込 | 🗄️`社員データ` | DB `SHAINMST` | 「社員」`Syainn` | 社員番号＋氏名を全件取得 |
| 4 | 受付番号入力 | 📊`Worksheet_Change` | ユーザー入力 `B3` | — | `Uno`セル変更を検知 |
| 5 | メイン制御 | 📄`起動Main` | — | — | 処理フロー制御 |
| 6 | 装置判定 | 🗄️`在庫実績データ` | DB `ZAIKO` / `ZAIKOR` | `mpErrDes` | 1号機チェック |
| 7 | 在庫取得 | 🗄️`在庫実績データ` | DB `ZAIKO`+`SEHMST` | 「実績」`Zaiko` | 在庫21項目をJOIN取得 |
| 8 | 在庫フォールバック | 🗄️`在庫実績データ` | DB `ZAIKOR`+`SEHMST` | 「実績」`Zaiko` | 在庫なし→履歴検索 |
| 9 | 実績取得 | 🗄️`在庫実績データ` | DB `SYOUJ2` / `SYOUJR2` | 「実績」`Jisseki` | 実績29項目を取得 |
| 10 | 在庫表示 | 🖥️`在庫表示` | 「実績」行6 | 「実績表示」ヘッダ部 | 名前付き範囲に代入 |
| 11 | 数式変換 | 📊数式 | 「実績」行4 | 「実績」行6 | 日付・コード→名称変換 |
| 12 | 実績表示 | 🖥️`実績表示` | 「実績」行11〜 | 「実績表示」`JissekiTB` | 照射回ごとに列展開 |
| 13 | 作業者名変換 | 📊数式 | 「実績」X〜AA列 | 「実績」AE〜AH列 | VLOOKUP(`Syainn`)で氏名変換 |

### 9.2 データフローツリー図

```
📊 ユーザー操作
└── 受付番号入力 → 「実績表示」B3 (Uno)
    │
    ├── 📄 Worksheet_Change (Sheet2)
    │   ├── [Uno空] → 🖥️ GamennCls2() → 表示クリア
    │   └── [Uno値あり] → 📄 起動Main()
    │       │
    │       ├── 🗄️ 社員データ()
    │       │   └── DB SHAINMST → 「社員」Syainn
    │       │
    │       ├── 🗄️ 在庫実績データ()
    │       │   ├── DB ZAIKO → 装置判定（1号機チェック）
    │       │   ├── DB ZAIKOR → 装置判定（履歴1号機チェック）
    │       │   ├── DB ZAIKO + SEHMST → 「実績」Zaiko（行3-4）
    │       │   │   └── [0件] → DB ZAIKOR + SEHMST → 「実績」Zaiko（履歴）
    │       │   └── DB SYOUJ2 → 「実績」Jisseki（行10-110）
    │       │       └── [在庫が履歴] → DB SYOUJR2 → 「実績」Jisseki
    │       │
    │       ├── [エラー] → MsgBox → 🖥️ GamennCls2()
    │       │
    │       └── [正常]
    │           ├── 🖥️ GamennCls2() → 表示領域クリア
    │           ├── 🖥️ 在庫表示()
    │           │   └── 「実績」行6 → 「実績表示」ヘッダ部（名前付き範囲）
    │           └── 🖥️ 実績表示()
    │               ├── 「実績」行11〜 → 「実績表示」行7〜31（照射実績）
    │               └── パス時刻分解 → 「実績表示」行34〜（入退室時刻）
    │
    └── 📊 ブック起動時
        ├── 📄 Workbook_Open()
        │   ├── シート保護設定
        │   ├── 🖥️ GamenCls1() → 全画面クリア
        │   └── 🗄️ 社員データ() → 社員マスタ読込
        │
        └── 📊「終了」ボタン
            └── 🖥️ Bookを閉じる() → ブック終了
```

---

## 10. セキュリティ注意事項

| # | 項目 | 内容 | 対策 |
|---|------|------|------|
| 1 | DB認証情報のハードコード | `DSN=ricdb;UID=ric;PWD=t6101` がVBAソースに直書き | 環境変数・設定ファイル等への外部化を推奨 |
| 2 | SQLインジェクション | `Uno` の値をそのまま文字列結合でSQL構築（`WHERE uno='..." & myUno & "'`） | パラメータ化クエリの使用を推奨 |
| 3 | エラー処理での `Stop` | `SQL_Exe()` 内でエラー時に `Stop` が実行される（開発環境ではブレーク、実行環境では実行時エラー） | 本番環境では `Stop` を除去しログ出力に変更を推奨 |
| 4 | シート保護 | `UserInterfaceOnly:=True` でUI操作は制限されるが、VBAパスワード保護なし | VBAプロジェクトのパスワード保護を推奨 |
| 5 | `On Error Resume Next` の多用 | `Open_oraconDB()` / `SQL_INSERT_UPDATE()` 等で広範なエラー無視 | エラーハンドリングの適切な範囲制限を推奨 |
| 6 | `DisplayAlerts = False` | `Workbook_BeforeClose` / `Bookを閉じる()` で変更の保存確認を抑制 | 意図的な設計だが、データ消失リスクあり |
