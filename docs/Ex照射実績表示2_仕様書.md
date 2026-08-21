# Ex照射実績表示2 仕様書

> **ファイル種別**: .xlsm（マクロ付き）
> **用途**: 受付番号を入力し、照射実績（在庫情報・照射データ）を一覧表示する検索ツール
> **VBA プロジェクト**: モジュール 13 本（.bas 9 / .cls 4 / .frm 0）
> **外部連携**: DSN=ricdb（Oracle）
> **解析日**: 2026-06-29
| 用途 | 照射実績データの表示（EXメニューの1ファイルとして照射管理システムを補完） |

---


## 凡例（本仕様書の表記ルール）

本仕様書では、保守時の判別を容易にするため、以下の表記ルールを使用します。

### 用語規約

| 用語 | 意味 |
| --- | --- |
| EXメニュー | ExRicSys フォルダに配置される VBA ファイル群（Ex*.xlsm）の総称 |
| `ExRicSys` | 配置フォルダ名。初出時は〈EXメニュー配置フォルダ〉と注記 |

### 表記規則

| 種別 | 表記 | 例 |
| --- | --- | --- |
| モジュール（.bas / .cls） | **太字** | **スタート処理.bas** |
| ユーザーフォーム（.frm） | **太字** | **SoutiSenntaku.frm** |
| プロシージャ / イベント | `コード体()` | `生産情報開始処理()` |
| シート名 | 「」 | 「未処理品一覧」 |
| セル参照 | `コード体` | `$W$3` |
| 名前付き範囲 | `コード体` | `Amari` |
| DB テーブル / カラム | `コード体` | `ExKeikakuX` / `kakunin` |
| ユーザー操作 | （操作名） | （処理方法実行 Click） |
| 主要マーク | ✓ | ✓ = 保守時に最初に確認すべき項目 |

### データフロー 場所マーク（9章）

9章のデータフロー（テーブル・ツリー図）では、処理が行われる場所を以下のアイコンで区別します。

| アイコン | 種別 | 意味 |
| --- | --- | --- |
| 📊 | シート操作 | ワークシート上のセル書込み・読取り・表示変更 |
| 🖥️ | 画面操作 | ユーザーフォーム（.frm）の表示・入力・操作 |
| 🗄️ | DB操作 | Oracle DB への SELECT / INSERT / UPDATE / DELETE |
| 📄 | VBA内部処理 | 変数計算・条件分岐など、画面・シートに直接関与しない処理 |

### ✓（主要マーク）の判定基準

✓ は **保守時に最初に確認すべき項目** を示します。
判定基準は対象の種類ごとに以下のとおりです。

| 章 | 対象 | ✓ の判定基準 |
| --- | --- | --- |
| 1.1 | シート | ユーザーが直接操作する、または VBA が動的に表示/非表示を切り替える |
| 1.2 | ユーザーフォーム | ユーザー入力を受け付ける、または業務フローの起点となる |
| 1.3 / 6.0 | VBA モジュール | ① ユーザー操作の起点 ② DB I/O を含む ③ 他モジュールから呼び出される ④ コード行数上位 25%　のいずれか |
| 2 | セル / 名前付き範囲 | VBA から `Range()` で代入または参照され、業務ロジックに直結する |
| 3 | 名前付き範囲 | VBA から `Range()` で代入または参照され、業務ロジックに直結する |
| 5 | ボタン / コントロール | DB 更新・画面遷移・計算実行など副作用のある操作を起動する |
| 6.0（全プロシージャ） | プロシージャ | ① ユーザー操作の起点（Click イベント等） ② DB I/O を実行 ③ 他モジュールから呼び出される Public　のいずれか |
| 7 | フォームコントロール | ユーザー入力を受け付ける、またはイベントで業務処理を起動する |
| 8.2 | DB テーブル | INSERT / UPDATE / DELETE の対象（参照のみのテーブルは ✓ なし） |

---

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

> ✓ = ユーザーが直接操作する、または VBA が動的に表示/非表示を切り替えるシート


| ✓ | No | シート名 | 最大行 | 最大列 | 保存時 Visible | VBA による動的切替 |
| --- | --- | --- | --- | --- | --- | --- |
| ✓ | 1 | 実績表示 | — | — | visible | — |
|  | 2 | 実績 | — | — | hidden | 保存時非表示。VBA が DB データを書き込む |
|  | 3 | 社員 | — | — | hidden | 保存時非表示。VBA が社員マスタを書き込む |

### 1.2 ユーザーフォーム一覧

なし。ユーザーフォームは存在しない。

### 1.3 VBA モジュール一覧

> ✓ = ユーザー操作の起点 / DB I/O を含む / 他モジュールから呼び出される / コード行数上位 25%


| ✓ | No | モジュール | 種別 | プロシージャ数 | 主な役割 |
| --- | --- | --- | --- | --- | --- |
| ✓ | 1 | **ThisWorkbook** | ThisWorkbook | 2 | ブック開閉時の初期化・終了処理 |
|  | 2 | **Sheet1** | Sheet（「実績」） | 0 | 宣言のみ（未使用） |
| ✓ | 3 | **Sheet2** | Sheet（「実績表示」） | 1 | `Uno`セル変更イベントで検索起動 |
|  | 4 | **Sheet3** | Sheet（「社員」） | 0 | 宣言のみ（未使用） |
|  | 5 | **Utility** | 標準モジュール | 1 | イベント有効化ユーティリティ |
| ✓ | 6 | **画面クリア** | 標準モジュール | 2 | 画面・データ領域の初期化 |
| ✓ | 7 | **起動** | 標準モジュール | 1 | 検索メイン処理の制御 |
| ✓ | 8 | **表示** | 標準モジュール | 2 | 在庫・実績データの画面表示 |
|  | 9 | **終了処理** | 標準モジュール | 1 | ブックを閉じる処理 |
| ✓ | 10 | **SQL_Execution** | 標準モジュール | 6 | DB接続・SQL実行・シート/配列格納 |
| ✓ | 11 | **在庫実績Read** | 標準モジュール | 1 | 在庫・実績データのSQL組み立てとDB読み込み |
| ✓ | 12 | **社員データRead** | 標準モジュール | 1 | 社員マスタのDB読み込み |
| ✓ | 13 | **共通変数** | 標準モジュール | 0 | Public変数宣言 |

---

## 2. シート詳細

### 2.0 シート可視性一覧


| No | シート | VBA による非表示化 | 表示するタイミング | 非表示にするタイミング | 制御プロシージャ |
| --- | --- | --- | --- | --- | --- |
| 1 | 実績表示 | — | 常時表示 | — | — |
| 2 | 実績 | あり（保存時 Hidden） | VBA が在庫・実績を書き込むとき | 保存時は非表示のまま | `在庫実績データ()` |
| 3 | 社員 | あり（保存時 Hidden） | VBA が社員マスタを書き込むとき | 保存時は非表示のまま | `社員データ()` |


> 以下の各シートのレイアウト構造表における ✓ = VBA から `Range()` で代入または参照され、業務ロジックに直結するセル

### 2.0 b 非表示行・列一覧

| シート | 非表示行 | 非表示列 |
| --- | --- | --- |
| 実績表示 | 1, 6 | A |

### 2.1 実績表示（codeName: Sheet2）

**目的**: ユーザーが直接操作する画面。受付番号を入力すると在庫情報と照射実績が自動表示される。

#### 非表示行・列

なし。

#### レイアウト構成

**ヘッダ部（行2〜5）：在庫情報表示**

| ✓ | No | セル | 名前付き範囲 | 内容 | VBA参照 |
| --- | --- | --- | --- | --- | --- |
|  | 1 | `B2` | ラベル | "受付番号" |  |
| ✓ | 2 | `C2` | `Kaisyacd` | 会社コード | `在庫表示()` で代入 |
| ✓ | 3 | `D2` | `Kainame` | 会社名 | `在庫表示()` で代入 |
| ✓ | 4 | `G2` | `Sehncd` | 製品コード | `在庫表示()` で代入 |
| ✓ | 5 | `H2` | `SName` | 製品名（指定線名） | `在庫表示()` で代入 |
|  | 6 | `L2` | ラベル | "指定線量kGy" |  |
| ✓ | 7 | `N2` | `Siteisn` | 指定線量 | `在庫表示()` で代入 |
| ✓ | 8 | `B3` | `Uno` | 受付番号（ユーザー入力セル） | `Worksheet_Change` トリガー |
|  | 9 | `C3` | ラベル | "受付数量" |  |
| ✓ | 10 | `D3` | `Nyukasu` | 受付数量 | `在庫表示()` で代入 |
|  | 11 | `E3` | ラベル | "受付日" |  |
| ✓ | 12 | `F3` | `Koudate` | 受付日 | `在庫表示()` で代入 |
|  | 13 | `H3` | ラベル | "分類" |  |
| ✓ | 14 | `I3` | `Tani` | 分類（一般品等） | `在庫表示()` で代入 |
|  | 15 | `J3` | ラベル | "受付" |  |
| ✓ | 16 | `K3` | `Kousncd` | 受付者名 | `在庫表示()` で代入 |
| ✓ | 17 | `M3` | ラベル | 荷姿表示（"Box入数"/"PL積載数"） | `在庫表示()` で条件代入 |
| ✓ | 18 | `N3` | `Incnt` | 入数 | `在庫表示()` で代入 |
| ✓ | 19 | `E4` | `Kagensn` | 管理点線量 下限 | `在庫表示()` で代入 |
| ✓ | 20 | `G4` | `Jyougsn` | 管理点線量 上限 | `在庫表示()` で代入 |
| ✓ | 21 | `I4` | `Labelcd` | ラベルコード | `在庫表示()` で代入 |
| ✓ | 22 | `K4` | `Pass` | 指定パス数 | `在庫表示()` で代入 |
| ✓ | 23 | `M4` | `Syouso` | 装置番号 | `在庫表示()` で代入 |
| ✓ | 24 | `D5` | `Syouzusu` | 照射済数 | `在庫表示()` で代入 |
|  | 25 | `E5` | ラベル | "線量検査合格数" |  |
| ✓ | 26 | `G5` | `Senkssu` | 線量検査合格数 | `在庫表示()` で代入 |
| ✓ | 27 | `I5` | `Syukasu` | 出荷済数 | `在庫表示()` で代入 |
| ✓ | 28 | `K5` | `Syukabi` | 出荷日 | `在庫表示()` で代入 |
| ✓ | 29 | `N5` | `Syukacd` | 出荷者コード | `在庫表示()` で代入 |

**実績データ部（行7〜31）**

| 行 | 項目 | 内容 |
| --- | --- | --- |
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

| ✓ | セル範囲 | 名前付き範囲 | 内容 |
| --- | --- | --- | --- |
|  | `B34:B131` | `PassKai` | パス回数番号（1, 2, 3, ...） |
|  | `C7:FZ131` | `JissekiTB` | 実績データ表示領域全体 |

行32にヘッダ「入室」「退室」、行33に「日時分」を表示。行34以降に各パスの入室・退室時刻を dd hh:mm 形式で表示。

### 2.2 実績（codeName: Sheet1）

**目的**: DBから取得したデータの一時格納シート。ユーザーには非表示。

#### 非表示行・列

なし。

| ✓ | セル範囲 | 名前付き範囲 | 内容 |
| --- | --- | --- | --- |
|  | `A1` | — | "在庫製品" / "履歴製品" / "在庫／履歴共データ無" |
|  | `A3:T4` | `Zaiko` | 在庫データ（DB→シート格納領域、行3ヘッダ・行4データ） |
|  | `A6:U6` | — | 在庫データの加工行（数式で変換） |
|  | `A10:AD110` | `Jisseki` | 実績データ（DB→シート格納領域、行10ヘッダ・行11以降データ） |
|  | `U6` | `SeiName` | 製品名（数式参照用） |
|  | `AE6:AH110` | — | 作業者名変換（社員コード→氏名のVLOOKUP） |

#### 数式

| セル | 数式 | 用途 |
| --- | --- | --- |
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

### 2.3 社員（codeName: Sheet3）

**目的**: 社員マスタ（社員番号・氏名）をDBから読み込んで格納するマスタシート。

#### 非表示行・列

なし。

| ✓ | セル範囲 | 名前付き範囲 | 内容 |
| --- | --- | --- | --- |
|  | `B2:C110` | `Syainn` | 社員マスタ（B列:社員番号, C列:氏名） |
|  | `D1` | `ReadDate` | 読み込み日時 |
|  | `E3:F7` | `Bunnrui` | 分類マスタ（E列:コード, F列:名称） |

**分類マスタ（`Bunnrui`）内容：**

| コード | 分類名 |
| --- | --- |
| 1 | 医療機器 |
| 2 | 一般品 |
| 3 | 医薬品 |
| 4 | 試験品 |

---

## 3. 名前付き範囲一覧

> ✓ = VBA から `Range()` で代入または参照され、業務ロジックに直結する名前付き範囲

| ✓ | No | 名前 | 参照先 | 業務的意味 |
| --- | --- | --- | --- | --- |
|  | 1 | `Bunnrui` | 社員!$E$3:$F$7 | 分類コード→名称変換テーブル |
|  | 2 | `Incnt` | 実績表示!$N$3 | 入数 |
|  | 3 | `Jisseki` | 実績!$A$10:$AD$110 | 実績データ格納領域 |
|  | 4 | `JissekiTB` | 実績表示!$C$7:$FZ$131 | 実績表示領域 |
|  | 5 | `Jyougsn` | 実績表示!$G$4 | 管理点線量 上限 |
|  | 6 | `Kagensn` | 実績表示!$E$4 | 管理点線量 下限 |
|  | 7 | `Kainame` | 実績表示!$D$2 | 会社名 |
|  | 8 | `Kaisyacd` | 実績表示!$C$2 | 会社コード |
|  | 9 | `Koudate` | 実績表示!$F$3 | 受付日 |
|  | 10 | `Kousncd` | 実績表示!$K$3 | 受付者コード |
|  | 11 | `Labelcd` | 実績表示!$I$4 | ラベルコード |
|  | 12 | `Nyukasu` | 実績表示!$D$3 | 受付数量 |
|  | 13 | `Pass` | 実績表示!$K$4 | 指定パス数 |
|  | 14 | `PassKai` | 実績表示!$B$34:$B$131 | パス回数表示列 |
|  | 15 | `ReadDate` | 社員!$D$1 | 社員データ読み込み日時 |
|  | 16 | `Sehncd` | 実績表示!$G$2 | 製品コード |
|  | 17 | `SeiName` | 実績!$U$6 | 製品名（数式用） |
|  | 18 | `Senkssu` | 実績表示!$G$5 | 線量検査合格数 |
|  | 19 | `Siteisn` | 実績表示!$N$2 | 指定線量 |
|  | 20 | `SName` | 実績表示!$H$2 | 製品名表示 |
|  | 21 | `Syainn` | 社員!$B$2:$C$110 | 社員マスタテーブル |
|  | 22 | `Syouso` | 実績表示!$M$4 | 装置番号 |
|  | 23 | `Syouzusu` | 実績表示!$D$5 | 照射済数 |
|  | 24 | `Syukabi` | 実績表示!$K$5 | 出荷日 |
|  | 25 | `Syukacd` | 実績表示!$N$5 | 出荷者コード |
|  | 26 | `Syukasu` | 実績表示!$I$5 | 出荷済数 |
|  | 27 | `Tani` | 実績表示!$I$3 | 分類 |
|  | 28 | `Uno` | 実績表示!$B$3 | 受付番号（入力セル） |
|  | 29 | `Zaiko` | 実績!$A$3:$T$4 | 在庫データ格納領域 |

### 3.1 データの入力規則（実績表示）

| セル | 名前付き範囲 | 種別 | 制約 | 用途 |
| --- | --- | --- | --- | --- |
| `B3` | `Uno` | 整数 | > 1990010000 | 受付番号（10桁。入力値がこの範囲外の場合エラー） |
| `C3:D3`, `C5:D5`, `E3:N5` | — | 整数 | =9999（固定値のみ許可） | 内部制御用セル |

---

## 4. 数式一覧

数式は「実績」シートにのみ存在する。

### 実績

| セル | 数式 | 説明 |
| --- | --- | --- |
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

> ✓ = DB 更新・画面遷移・計算実行など副作用のある操作を起動するボタン

### 5.1 シート上のボタン（Form Control）

| ✓ | No | シート | ボタンラベル | 割り当てマクロ | 動作概要 |
| --- | --- | --- | --- | --- | --- |
|  | 1 | 実績表示 | 終了 | `Bookを閉じる()` | ブックを閉じる（上書き保存なし） |

### 5.2 ユーザーフォーム上のボタン（サマリ）

なし。

### 5.3 ショートカットキー

| No | マクロ名 | ショートカット | 処理概要 |
| --- | --- | --- | --- |
| 1 | `GamenCls1()` | **Ctrl+E** | 全名前付き範囲をクリアし初期状態に戻す |

### 5.4 CommandBar

なし。CommandBarの定義は存在しない。

---

## 6. VBAモジュール仕様

### 6.0 全プロシージャ一覧

> ✓ = ユーザー操作の起点（Click イベント等） / DB I/O を実行 / 他モジュールから呼び出される Public


| ✓ | No | モジュール | プロシージャ | スコープ | 種別 | 概要 |
| --- | --- | --- | --- | --- | --- | --- |
|  | 1 | **ThisWorkbook** | `Workbook_BeforeClose()` | Public | Private Sub | 保存確認なしでブックを閉じる |
| ✓ | 2 | **ThisWorkbook** | `Workbook_Open()` | Public | Private Sub | 📊シート保護設定 → `GamenCls1()` → `社員データ()` 呼び出し |
| ✓ | 3 | **Sheet2**（「実績表示」） | `Worksheet_Change()` | Public | Private Sub | 📊`Uno`セル変更時、空なら画面クリア / 値ありなら`起動Main()` |
|  | 4 | **Utility** | `aaaaa()` | Public | Sub | イベント有効化（デバッグ用） |
| ✓ | 5 | **画面クリア** | `GamenCls1()` | Public | Sub | 🖥️全名前付き範囲を空文字にクリア、PrintAreaリセット |
| ✓ | 6 | **画面クリア** | `GamennCls2()` | Public | Sub | 🖥️実績表示領域とヘッダ（Uno以外）をクリア |
| ✓ | 7 | **起動** | `起動Main()` | Public | Sub | 📄メイン制御：`社員データ()` → `在庫実績データ()` → `在庫表示()` → `実績表示()` |
| ✓ | 8 | **表示** | `在庫表示()` | Public | Sub | 🖥️「実績」シートから在庫データを読み取り「実績表示」シートに埋め込み |
| ✓ | 9 | **表示** | `実績表示()` | Public | Sub | 🖥️「実績」シートから実績データを読み取り「実績表示」シートに一覧表示 |
|  | 10 | **終了処理** | `Bookを閉じる()` | Public | Sub | 🖥️メッセージ非表示で閉じる（最後の1ブックならExcel終了） |
| ✓ | 11 | **SQL_Execution** | `Open_oraconDB()` | Public | Sub | 🗄️ADODB.Connectionを使ったODBC接続 |
| ✓ | 12 | **SQL_Execution** | `SQL_Exe()` | Public | Sub | 🗄️SQL文実行（Execute） |
| ✓ | 13 | **SQL_Execution** | `SQL_INSERT_UPDATE()` | Public | Sub | 🗄️INSERT/UPDATE汎用処理（キー存在チェック付き） |
| ✓ | 14 | **SQL_Execution** | `SQL_Delete()` | Public | Sub | 🗄️DELETE汎用処理 |
| ✓ | 15 | **SQL_Execution** | `Disp_Sheet()` | Public | Sub | 🗄️SQL実行結果をシートに貼り付け |
| ✓ | 16 | **SQL_Execution** | `Set_Array()` | Public | Sub | 🗄️SQL実行結果を配列に格納 |
| ✓ | 17 | **在庫実績Read** | `在庫実績データ()` | Public | Sub | 🗄️在庫・実績データのSQL構築とDB取得 |
| ✓ | 18 | **社員データRead** | `社員データ()` | Public | Sub | 🗄️社員マスタのDB取得 |

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
| --- | --- | --- |
| `mpDataCount` | Single | 実績レコード数 |
| `mpPass` | Integer | 指定パス数 |

---

## 7. ユーザーフォーム仕様

ユーザーフォームは存在しない。

---

## 8. DB 接続・外部連携

### 8.1 ODBC 接続情報

| No | DSN | UID | PWD | 用途 | 使用箇所 |
| --- | --- | --- | --- | --- | --- |
| 1 | ricdb | ric | t6101 | メインDB（RIC スキーマ：在庫・製品マスタ・照射実績・社員マスタ） | **在庫実績Read** `在庫実績データ()`, **社員データRead** `社員データ()` |
| 2 | ricdb | rich | t6101 | 履歴DB（rich スキーマ：在庫履歴・照射実績履歴） | **在庫実績Read** `在庫実績データ()`（`rich.zaikor`, `RICH.SYOUJR2`） |

### 8.2 テーブル一覧

> ✓ = INSERT / UPDATE / DELETE の対象テーブル（参照のみのテーブルは ✓ なし）

| ✓ | No | スキーマ | テーブル名 | 区分 | 用途 |
| --- | --- | --- | --- | --- | --- |
|  | 1 | RIC | `ZAIKO` | 参照 | 在庫データ（製品・受付情報） |
|  | 2 | RIC | `SEHMST` | 参照 | 製品マスタ（製品名取得用） |
|  | 3 | RIC | `SYOUJ2` | 参照 | 照射実績データ（現行） |
|  | 4 | rich | `ZAIKOR` | 参照 | 在庫履歴データ（過去分） |
|  | 5 | rich | `SYOUJR2` | 参照 | 照射実績履歴データ（過去分） |
|  | 6 | RIC | `SHAINMST` | 参照 | 社員マスタ |

全テーブルが SELECT のみのため ✓ 該当なし。`SQL_INSERT_UPDATE()` / `SQL_Delete()` は汎用関数として実装されているが、本ファイル内では呼び出されていない。

### 8.3 SQL 一覧

#### 8.3.1 装置番号判定（1号機チェック）（`在庫実績データ()` / **在庫実績Read.bas**）

```sql
SELECT TO_NUMBER(syouso) FROM zaiko WHERE uno='...'
```

#### 8.3.2 履歴の装置番号判定（`在庫実績データ()` / **在庫実績Read.bas**）

```sql
SELECT TO_NUMBER(syouso) FROM rich.zaikor WHERE uno='...'
```

#### 8.3.3 在庫データ取得（`在庫実績データ()` / **在庫実績Read.bas**）

```sql
SELECT z.uno, z.kaisyacd, z.kainame, z.sehncd, z.siteisn, z.nyukasu, z.nyukabi, TO_NUMBER(z.tani), z.kousncd, z.incnt, z.kagensn, z.jyougsn, z.labelcd, z.pass, z.syouso, z.syouzusu, z.senkssu, z.syukasu, z.syukabi, z.syukacd, TRIM(h.seiname) FROM RIC.ZAIKO z, ric.sehmst h WHERE uno='...' AND z.kaisyacd=h.kaisyacd AND z.sehncd=h.sehncd
```

#### 8.3.4 在庫履歴データ取得（`在庫実績データ()` / **在庫実績Read.bas**）

```sql
SELECT z.uno, z.kaisyacd, z.kainame, z.sehncd, z.siteisn, z.nyukasu, z.nyukabi, TO_NUMBER(z.tani), z.kousncd, z.incnt, z.kagensn, z.jyougsn, z.labelcd, z.pass, z.syouso, z.syouzusu, z.senkssu, z.syukasu, z.syukabi, z.syukacd, TRIM(h.seiname) FROM RICH.ZAIKOR z, ric.sehmst h WHERE uno='...' AND z.kaisyacd=h.kaisyacd AND z.sehncd=h.sehncd
```

#### 8.3.5 実績データ取得（`在庫実績データ()` / **在庫実績Read.bas**）

```sql
SELECT uno, fricno, lricno, tuikaflg, suryou*1, kmmdd, tonyutime, saikatime, hangno, syafutime, smmdd, dattktime, sesdate, senkno, senksyu, toridcd, atusa*1, sokutti*1, keisask, ondok, sokutsn*1, sezhnsu, jituno*1, keikacd, tonyucd, sokutcd, syuhncd, nyutime, taitime FROM RIC.SYOUJ2 WHERE uno='...' ORDER BY fricno, senkno, tuikaflg
```

#### 8.3.6 実績履歴データ取得（`在庫実績データ()` / **在庫実績Read.bas**）

```sql
SELECT uno, fricno, lricno, tuikaflg, suryou*1, kmmdd, tonyutime, saikatime, hangno, syafutime, smmdd, dattktime, sesdate, senkno, senksyu, toridcd, atusa*1, sokutti*1, keisask, ondok, sokutsn*1, sezhnsu, jituno*1, keikacd, tonyucd, sokutcd, syuhncd, nyutime, taitime FROM RICH.SYOUJR2 WHERE uno='...' ORDER BY fricno, senkno, tuikaflg
```

#### 8.3.7 社員マスタ全件取得（`社員データ()` / **社員データRead.bas**）

```sql
SELECT TO_NUMBER(shano), REPLACE(TRIM(shaname),'　','') FROM shainmst ORDER BY shano
```

### 8.4 外部ファイル参照

なし。外部ファイルへの参照は存在しない。

---

## 9. データフロー

各フローは「起点 → 処理 → 結果」の粒度で記述する。

### 9.1 データフローテーブル

| No | ステップ | 場所 | データソース | データ先 | 処理内容 |
| --- | --- | --- | --- | --- | --- |
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


| No | カテゴリ | 内容 | リスク |
| --- | --- | --- | --- |
| 1 | 認証情報ハードコード | DSN=`ricdb`, UID=`ric`, PWD=`t6101` が VBA ソースに平文記載 | 中：VBAエディタで閲覧可能 |
| 2 | SQLインジェクション | `Uno` の値を文字列連結で SQL 構築（`WHERE uno='..." & myUno & "'`） | 中：文字列連結によるSQL構築 |
| 3 | エラーハンドリング | `SQL_Exe()` 内でエラー時に `Stop` が実行される | 中：本番でブレークする可能性 |
| 4 | シート保護 | `UserInterfaceOnly:=True`。VBA パスワード保護なし | 低：VBAからは無制限アクセス |
| 5 | `On Error Resume Next` の多用 | `Open_oraconDB()` / `SQL_INSERT_UPDATE()` 等で広範なエラー無視 | 中：エラーハンドリングの適切な範囲制限を推奨 |
| 6 | `DisplayAlerts = False` | `Workbook_BeforeClose` / `Bookを閉じる()` で変更の保存確認を抑制 | 低：意図的な設計だが、データ消失リスクあり |

## スコープ外（本仕様書に含まないもの）

- セル書式（色・罫線・フォント）
- 条件付き書式、グラフ・画像、印刷設定

必要な場合は Excel 画面のスクリーンショットで補完してください。
