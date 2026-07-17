# Ex出荷方法報告書発行登録 仕様書

> **対象ファイル**: Ex出荷方法報告書発行登録.xlsm
> **ファイル種別**: .xlsm（マクロ付き）
> **用途**: 取引先（業者）ごとの出荷方法（引取・混載便・保管品など）および報告書発行種別を Oracle DB（`ExSeihinJ` テーブル）に登録・管理する
> **VBA プロジェクト**: モジュール 9 本（.bas 6 / .cls 2 / .frm 0）
> **外部連携**: DSN=ricdb（Oracle DB）
> **解析日**: 2026-06-29（excel-to-md スキルによる自動解析）

---

## 凡例（本仕様書の表記ルール）

### 用語規約

| 用語 | 意味 |
|---|---|
| EXメニュー | ExRicSys フォルダに配置される VBA ファイル群（Ex*.xlsm）の総称 |
| `ExRicSys` | 配置フォルダ名。初出時は〈EXメニュー配置フォルダ〉と注記 |

### 表記規則

| 種別 | 表記 | 例 |
|---|---|---|
| モジュール（.bas / .cls） | **太字** | **登録.bas** |
| プロシージャ / イベント | `コード体()` | `引取報告書登録()` |
| シート名 | 「」 | 「業者一覧」 |
| セル参照 | `コード体` | `$B$5` |
| 名前付き範囲 | `コード体` | `HikitoriTB` |
| DB テーブル / カラム | `コード体` | `ExSeihinJ` / `hikitori` |
| ユーザー操作 | （操作名） | （変更を登録する Click） |
| 主要マーク | ✓ | ✓ = 保守時に最初に確認すべき項目 |

### データフロー 場所マーク（9章）

| アイコン | 種別 | 意味 |
|---|---|---|
| 📊 | シート操作 | ワークシート上のセル書込み・読取り・表示変更 |
| 🖥️ | 画面操作 | ユーザーフォーム（.frm）の表示・入力・操作 |
| 🗄️ | DB操作 | DB への SELECT / INSERT / UPDATE / DELETE |
| 📄 | VBA内部処理 | 変数計算・条件分岐など、画面・シートに直接関与しない処理 |

### ✓（主要マーク）の判定基準

| 章 | 対象 | ✓ の判定基準 |
|---|---|---|
| 1.1 | シート | ユーザーが直接操作する、または VBA が動的に表示/非表示を切り替える |
| 1.3 / 6.0 | VBA モジュール | ① ユーザー操作の起点 ② DB I/O を含む ③ 他モジュールから呼び出される ④ コード行数上位 25% のいずれか |
| 2 | セル / 名前付き範囲 | VBA から `Range()` で代入または参照され、業務ロジックに直結する |
| 3 | 名前付き範囲 | VBA から `Range()` で代入または参照され、業務ロジックに直結する |
| 5 | ボタン / コントロール | DB 更新・画面遷移・計算実行など副作用のある操作を起動する |
| 6.0（全プロシージャ） | プロシージャ | ① ユーザー操作の起点（Click イベント等） ② DB I/O を実行 ③ 他モジュールから呼び出される Public のいずれか |
| 8.2 | DB テーブル | INSERT / UPDATE / DELETE の対象（参照のみのテーブルは ✓ なし） |

---

## 目次

1. [ファイル構成](#1-ファイル構成)
2. [シート詳細](#2-シート詳細)
3. [名前付き範囲一覧](#3-名前付き範囲一覧)
4. [数式一覧](#4-数式一覧)
5. [ボタン・マクロ対応](#5-ボタンマクロ対応)
6. [VBA モジュール仕様](#6-vba-モジュール仕様)
7. [ユーザーフォーム仕様](#7-ユーザーフォーム仕様)
8. [DB 接続・外部連携](#8-db-接続外部連携)
9. [データフロー](#9-データフロー)
10. [セキュリティ注意事項](#10-セキュリティ注意事項)

---

## 1. ファイル構成

```
Ex出荷方法報告書発行登録.xlsm
├── シート
│   └── 業者一覧（出荷方法・報告書発行種別の一覧・編集画面）
├── VBA モジュール（9本）
│   ├── ThisWorkbook.cls       – ブックイベント（Open / BeforeClose）
│   ├── Sheet1.cls             – 「業者一覧」Worksheet_Change イベント
│   ├── 登録.bas               – 変更行の DB INSERT/UPDATE
│   ├── 終了処理.bas           – ブックを閉じる処理
│   ├── デーた抽出.bas         – DB からデータ取得・シート展開
│   ├── ユーティリティ.bas     – イベント有効化・画面クリア
│   ├── 画面クリア引取業者.bas – 画面消去処理
│   ├── SQL_Execution.bas      – ADODB 接続共通ライブラリ
│   └── 画面操作1.bas          – 画面操作サンプル集（参考モジュール）
├── ボタン（3個）
│   ├── 出荷方法/報告書発行種別表示 → 業者名と引取抽出()
│   ├── 変更を登録する             → 引取報告書登録()
│   └── 終了                       → Bookを閉じる()
├── 描画オブジェクト
│   └── テキストボックス「CLS」（J1 付近、用途不明）
└── 外部リンク: なし
```

### 1.1 シート一覧（全 1 件）

> ✓ = ユーザーが直接操作する、または VBA が動的に表示/非表示を切り替えるシート

| ✓ | シート名 | 最大行 | 最大列 | 保存時 Visible | VBA による動的切替 |
|---|---|---|---|---|---|
| ✓ | 「業者一覧」 | 1004 | 17 (Q) | visible | — |

### 1.3 VBA モジュール一覧（全 9 件）

> ✓ = ユーザー操作の起点 / DB I/O を含む / 他モジュールから呼び出される / コード行数上位 25%
> 全モジュールを列挙し、✓ 基準に該当するものにマークを付ける。

| ✓ | モジュール | 種別 | プロシージャ数 | 主な役割 |
|---|---|---|---|---|
| ✓ | **ThisWorkbook.cls** | .cls | 2 | 起動/終了イベント |
| ✓ | **Sheet1.cls** | .cls | 1 | 「業者一覧」シートイベント |
| ✓ | **登録.bas** | .bas | 1 | 引取報告書登録（DB 更新） |
|   | **終了処理.bas** | .bas | 1 | ブックを閉じる処理 |
| ✓ | **デーた抽出.bas** | .bas | 2 | DB からデータ取得・シート展開 |
|   | **ユーティリティ.bas** | .bas | 2 | イベント有効化・画面クリア |
|   | **画面クリア引取業者.bas** | .bas | 1 | 画面消去処理 |
| ✓ | **SQL_Execution.bas** | .bas | 6 | ADODB 接続共通ライブラリ |
|   | **画面操作1.bas** | .bas | 8 | 画面操作サンプル集（参考モジュール） |

---

## 2. シート詳細

### 2.0 シート可視性一覧

| シート | VBA による非表示化 | 表示するタイミング | 非表示にするタイミング | 制御プロシージャ |
|---|---|---|---|---|
| 「業者一覧」 | — | — | — | — |

> 以下の各シートのレイアウト構造表における ✓ = VBA から `Range()` で代入または参照され、業務ロジックに直結するセル

### 2.1 業者一覧

**目的**: 全得意先（`kaisyacd` < 2000）の出荷方法と報告書発行種別を一覧表示する。現在値（E・G列）と DB 登録値（F・H列）を並べて差分を可視化し、変更を「変更を登録する」ボタンで DB に反映する。

#### 非表示行・列

| 種類 | 対象 | 備考 |
|---|---|---|
| 非表示行 | なし | — |
| 非表示列 | F, H, K, L, M, N, O | F・H: DB 保存値（比較用、ユーザーには非表示）。K〜L: `sehmst` 集計テーブル（VLOOKUP参照元）。M: 空列（区切り）。N〜O: 出荷方法・報告書発行種別のコード変換テーブル |

#### レイアウト構造

```
行 1: タイトル行（B1 = "出荷方法／報告書発行種別の登録"）
行 2〜3: 空行
行 4: 列見出し行
行 5〜1004/1005: データ本体（VBA が `HikitoriTB` / `HouTB` に DB データを書込み）
```

| ✓ | セル | 名前付き範囲 | 種別 | 実態（値/数式/VBA代入） | 業務的意味 |
|---|---|---|---|---|---|
|   | `$A$1` | `Debug` | 設定値 | 空欄 | デバッグフラグ |
|   | `$B$1` | — | 設定値 | "出荷方法／報告書発行種別の登録" | 画面タイトル |
|   | `$B$4` | — | 設定値 | "会社コード" | 列ヘッダー |
|   | `$C$4` | — | 設定値 | "略称" | 列ヘッダー |
|   | `$D$4` | — | 設定値 | "会社名" | 列ヘッダー |
|   | `$E$4` | — | 設定値 | "出荷方法" | 列ヘッダー（編集列） |
|   | `$F$4` | — | 設定値 | "出荷方法" | 列ヘッダー（DB値・非表示列） |
|   | `$G$4` | — | 設定値 | "報告書発行種別" | 列ヘッダー（編集列） |
|   | `$H$4` | — | 設定値 | "報告書発行種別" | 列ヘッダー（DB値・非表示列） |
|   | `$I$4` | — | 設定値 | "Sehmstの報告書要不要" | 列ヘッダー（VLOOKUP参照列） |
|   | `$K$4` | — | 設定値 | "会社コード" | 補助テーブルヘッダー（非表示列） |
|   | `$L$4` | — | 設定値 | "報告書発行フラグ" | 補助テーブルヘッダー（非表示列） |
|   | `$N$4` | — | 設定値 | "出荷方法" | コード変換テーブルヘッダー（非表示列） |
|   | `$O$4` | — | 設定値 | "報告書発行種別内容" | コード変換テーブルヘッダー（非表示列） |
| ✓ | `$B$5:$H$1005` | `HikitoriTB` | VBA代入 | `業者名と引取抽出()` で DB データ書込み | 業者データ入力・表示テーブル |
| ✓ | `$K$5:$L$1004` | `HouTB` | VBA代入 | `業者名と引取抽出()` で `sehmst` 集計データ書込み | 報告書要不要フラグ参照テーブル（VLOOKUP 参照元） |
| ✓ | `$I$5:$I$1004` | — | 数式 | `=IF(ISERROR(VLOOKUP(B5,$K$5:$L$1004,2,FALSE)),"",VLOOKUP(...))` | 同行B列の会社コードで `HouTB` を検索し報告書発行フラグ表示 |
|   | `$N$5:$N$13` | — | 設定値 | 引取 / 混載便 / 保管品 / チャータ便 / 納品 / 品証扱い / γ扱い / 営業扱い / その他 | 出荷方法の選択肢マスタ（9種） |
|   | `$O$5:$O$8` | — | 設定値 | 照射後Fax送信 / 出荷後Fax送信 / 照射後報告書発行 / 出荷後報告書発行 | 報告書発行種別の選択肢マスタ（4種） |

#### DB 読込列（行 5〜）

| 列 | ヘッダー | DB フィールド | 備考 |
|---|---|---|---|
| B | 会社コード | `tokumst.kaisyacd` | — |
| C | 略称 | `tokumst.kairname` | — |
| D | 会社名 | `tokumst.coname` | — |
| E | 出荷方法（現在値） | `ExSeihinJ.hikitori` | ユーザー編集対象 |
| F | 出荷方法（DB値） | `ExSeihinJ.hikitori` | 非表示・変更検出用 |
| G | 報告書発行種別（現在値） | `ExSeihinJ.housyube` | ユーザー編集対象 |
| H | 報告書発行種別（DB値） | `ExSeihinJ.housyube` | 非表示・変更検出用 |
| I | Sehmst報告書要不要 | VLOOKUP で K:L から取得 | 0=不要・それ以外=要 |
| K | 会社コード | `sehmst.kaisyacd` | 非表示・HouTB |
| L | 報告書発行フラグ | `sehmst.max(syouho)` | 非表示・HouTB |

#### コード変換テーブル（N5:O13）

| N列（出荷方法） | O列（報告書発行種別） |
|---|---|
| 引取 | 照射後Fax送信 |
| 混載便 | 出荷後Fax送信 |
| 保管品 | 照射後報告書発行 |
| チャータ便 | 出荷後報告書発行 |
| 納品 | — |
| 品証扱い | — |
| γ扱い | — |
| 営業扱い | — |
| その他 | — |

---

## 3. 名前付き範囲一覧

全 **3** 件。

> ✓ = VBA から `Range()` で代入または参照され、業務ロジックに直結する名前付き範囲
> 全件を列挙し、✓ 基準に該当するものにマークを付ける。

| ✓ | 名前 | 参照先 | 業務的意味 |
|---|---|---|---|
| ✓ | `HikitoriTB` | 業者一覧!$B$5:$H$1005 | 業者データ入力テーブル（会社コード〜報告書発行種別DB値）。`業者名と引取抽出()` / `画面消去処理()` / `画面クリア()` でクリア・書込み |
| ✓ | `HouTB` | 業者一覧!$K$5:$L$1004 | `sehmst` より集計した報告書要不要テーブル（I列 VLOOKUP の参照元）。`業者名と引取抽出()` / `画面消去処理()` / `画面クリア()` でクリア・書込み |
|   | `Debug` | 業者一覧!$A$1 | デバッグフラグ（VBA からの参照なし） |


### 3.1 データの入力規則

| シート | セル | 種別 | 制約 |
|---|---|---|---|
| 「業者一覧」 | `B5:D700` | 整数 | =99999 |
| 「業者一覧」 | `E5:E1004` | リスト | リスト: `$N$5:$N$13` |
| 「業者一覧」 | `G5:G1004` | リスト | リスト: `$O$5:$O$9` |


---

---

## 4. 数式一覧

| シート | 数式件数 | 備考 |
|---|---|---|
| 「業者一覧」 | 1000 | I5:I1004 に同一パターンの VLOOKUP 数式 |

### 4.1 業者一覧

| セル | 数式 | 説明 |
|---|---|---|
| `$I$5` | `=IF(ISERROR(VLOOKUP(B5,$K$5:$L$1004,2,FALSE)),"",VLOOKUP(B5,$K$5:$L$1004,2,FALSE))` | 同行 B 列の会社コードで `HouTB`（K:L列）を検索し、報告書発行フラグを表示。エラー時は空白 |
| `$I$6`〜`$I$1004` | 同上（行番号のみ変化） | I5 と同一パターンで全 1000 行に展開 |

> **補足**: `$K$5:$L$1004`（名前付き範囲 `HouTB` に相当）を固定参照しているため、`HouTB` への書込みが即座に I 列全体へ反映される。

---

## 5. ボタン・マクロ対応

### 5.1 シート上のボタン（Form Control）

> ✓ = DB 更新・画面遷移・計算実行など副作用のある操作を起動するボタン

| ✓ | シート | ボタンラベル | 割り当てマクロ | 動作概要 |
|---|---|---|---|---|
| ✓ | 「業者一覧」 | 出荷方法/報告書発行種別表示 | `業者名と引取抽出()` | DB（`tokumst` + `ExSeihinJ` + `sehmst`）から最新データを取得してシートに再描画 |
| ✓ | 「業者一覧」 | 変更を登録する | `引取報告書登録()` | E列またはG列が変更された行のみ DB（`ExSeihinJ`）に `hikitori` / `housyube` を UPDATE/INSERT |
|   | 「業者一覧」 | 終了 | `Bookを閉じる()` | ブックを保存せずに閉じる |

### 5.2 ショートカットキー

| マクロ名 | ショートカット | 処理概要 |
|---|---|---|
| `イベント有効()` | `r` | `Application.EnableEvents = True` に強制復旧（デバッグ用） |
| `画面消去処理()` | `e` | `HikitoriTB` / `HouTB` をクリアして `$E$5` に移動 |

### 5.3 起動元マッピング（全プロシージャ × 呼出元）

全プロシージャの起動元を整理し、呼出元のないプロシージャを明示する。

| プロシージャ | モジュール | 起動元 | 起動種別 |
|---|---|---|---|
| `Workbook_Open()` | **ThisWorkbook.cls** | Excel 自動イベント（ブック起動時） | イベント |
| `Workbook_BeforeClose()` | **ThisWorkbook.cls** | Excel 自動イベント（ブック終了時） | イベント |
| `Worksheet_Change()` | **Sheet1.cls** | Excel 自動イベント（G列セル変更時） | イベント |
| `業者名と引取抽出()` | **デーた抽出.bas** | ① ボタン【出荷方法/報告書発行種別表示】 ② `Workbook_Open()` から Call | ボタン + 内部Call |
| `引取報告書登録()` | **登録.bas** | ボタン【変更を登録する】 | ボタン |
| `Bookを閉じる()` | **終了処理.bas** | ボタン【終了】 | ボタン |
| `イベント有効()` | **ユーティリティ.bas** | ショートカットキー `r` | ショートカット |
| `画面消去処理()` | **画面クリア引取業者.bas** | ショートカットキー `e` | ショートカット |
| `Open_oraconDB()` | **SQL_Execution.bas** | `SQL_INSERT_UPDATE()` / `SQL_Delete()` / `Disp_Sheet()` / `Set_Array()` から内部Call | 内部Call |
| `SQL_Exe()` | **SQL_Execution.bas** | `SQL_INSERT_UPDATE()` / `SQL_Delete()` / `Disp_Sheet()` / `Set_Array()` から内部Call | 内部Call |
| `SQL_INSERT_UPDATE()` | **SQL_Execution.bas** | `引取報告書登録()` から Call | 内部Call |
| `Disp_Sheet()` | **SQL_Execution.bas** | `業者名と引取抽出()` から Call | 内部Call |
| **`報告書不要表示()`** | **デーた抽出.bas** | **呼出元なし**（旧版の残存コード。`業者名と引取抽出()` 内にインラインで同等処理が実装されており、本プロシージャは使用されていない。列参照が異なる点から旧レイアウト時代のコードと推定） | **⚠ 孤立** |
| **`画面クリア()`** | **ユーティリティ.bas** | **呼出元なし**（`業者名と引取抽出()` 内で `Range("HikitoriTB") = ""` / `Range("HouTB") = ""` を直接実行しており、本プロシージャは呼び出されていない） | **⚠ 孤立** |
| **`SQL_Delete()`** | **SQL_Execution.bas** | **本ファイルからの呼出元なし**（共通ライブラリとして定義されているが、本ファイルの業務フローに DELETE 処理は存在しない） | **⚠ 共通ライブラリ（未使用）** |
| **`Set_Array()`** | **SQL_Execution.bas** | **本ファイルからの呼出元なし**（共通ライブラリとして定義されているが、本ファイルでは使用されていない） | **⚠ 共通ライブラリ（未使用）** |
| `Reidai()` 他 8件 | **画面操作1.bas** | なし（開発サンプル集、プロダクションでは不使用） | 参考コード |

> **孤立プロシージャのまとめ**:
> - `報告書不要表示()` — 旧版の残存コード。列参照（F・G列に「不要」設定）が現行版（G・H列に「不要」設定）と異なり、旧レイアウト時代の処理と推定される
> - `画面クリア()` — `業者名と引取抽出()` 内で同等処理がインライン実行されており、重複コードとして残存
> - `SQL_Delete()` / `Set_Array()` — 共通ライブラリ（`SQL_Execution.bas`）の一部。他の EXメニューファイルでは使用される可能性があるが、本ファイル内では不使用

---

## 6. VBA モジュール仕様

### 6.0 全プロシージャ一覧

全 **24** 件。

> ✓ = ユーザー操作の起点（Click イベント等） / DB I/O を実行 / 他モジュールから呼び出される Public
> 全プロシージャを列挙し、✓ 基準に該当するものにマークを付ける。

| ✓ | モジュール | プロシージャ | 種別 | 概要 |
|---|---|---|---|---|
| ✓ | **ThisWorkbook.cls** | `Workbook_Open()` | Event | シート保護設定後、`業者名と引取抽出()` を呼び出してデータ表示 |
| ✓ | **ThisWorkbook.cls** | `Workbook_BeforeClose()` | Event | アラート非表示で保存済みフラグ設定 |
| ✓ | **Sheet1.cls** | `Worksheet_Change()` | Event | G列変更時に報告書要不要をチェックし、不要なら強制上書き |
| ✓ | **登録.bas** | `引取報告書登録()` | Sub | E/G列の差分行を `ExSeihinJ` に INSERT/UPDATE |
| ✓ | **終了処理.bas** | `Bookを閉じる()` | Sub | 保存確認なしでブックを閉じる |
| ✓ | **デーた抽出.bas** | `業者名と引取抽出()` | Sub | DB からデータ取得して「業者一覧」シートに展開 |
|   | **デーた抽出.bas** | `報告書不要表示()` | Sub | 旧版の報告書不要設定処理（⚠ 呼出元なし・孤立コード） |
|   | **ユーティリティ.bas** | `イベント有効()` | Sub | `EnableEvents = True` 強制復旧（ショートカット `r`） |
|   | **ユーティリティ.bas** | `画面クリア()` | Sub | `HikitoriTB` / `HouTB` クリア（⚠ 呼出元なし・孤立コード） |
|   | **画面クリア引取業者.bas** | `画面消去処理()` | Sub | `HikitoriTB` / `HouTB` クリア + E5 移動（ショートカット `e`） |
| ✓ | **SQL_Execution.bas** | `Open_oraconDB()` | Sub | ADODB で Oracle DB に接続 |
| ✓ | **SQL_Execution.bas** | `SQL_Exe()` | Sub | SQL 文を Execute で実行 |
| ✓ | **SQL_Execution.bas** | `SQL_INSERT_UPDATE()` | Sub | 件数チェック → INSERT or UPDATE を動的生成・実行 |
|   | **SQL_Execution.bas** | `SQL_Delete()` | Sub | DELETE 文を生成・実行（本ファイルでは未使用） |
| ✓ | **SQL_Execution.bas** | `Disp_Sheet()` | Sub | SQL 結果をシートに CopyFromRecordset で貼り付け |
|   | **SQL_Execution.bas** | `Set_Array()` | Sub | SQL 結果を配列に格納（本ファイルでは未使用） |
|   | **画面操作1.bas** | `Reidai()` | Sub | Excel API 操作サンプル集 |
|   | **画面操作1.bas** | `印刷パラメータ設定()` | Sub | 印刷余白・用紙設定サンプル |
|   | **画面操作1.bas** | `DriveSearch()` | Function | ドライブ存在確認（FileSystemObject） |
|   | **画面操作1.bas** | `シート保護()` | Sub | シート保護パターンサンプル |
|   | **画面操作1.bas** | `複数列の選択()` | Sub | 列選択サンプル |
|   | **画面操作1.bas** | `セル名の定義_削除()` | Sub | 名前付き範囲の定義・削除サンプル |
|   | **画面操作1.bas** | `リボン操作()` | Sub | リボン表示/非表示サンプル（Excel 4 Macro） |
|   | **画面操作1.bas** | `セルのコピー_値の貼り付け()` | Sub | セルコピー・値貼り付けサンプル |

---

### 6.1 ThisWorkbook.cls

#### `Workbook_Open()`

**処理概要**: ブック起動時にシート保護（UIのみ）を設定してから `業者名と引取抽出()` を呼び出す。

**処理フロー**:
1. `ActiveSheet.Unprotect` でシート保護を解除
2. `ActiveSheet.Protect UserInterfaceOnly:=True` でUIのみ保護を再設定
3. `業者名と引取抽出()` を呼び出してDBからデータ取得・表示

```vba
Private Sub Workbook_Open()
    ActiveSheet.Unprotect
    ActiveSheet.Protect UserInterfaceOnly:=True
    Call 業者名と引取抽出
End Sub
```

#### `Workbook_BeforeClose(Cancel As Boolean)`

**処理概要**: ブックを閉じる前に保存確認ダイアログを抑制し、変更なしで閉じる。

```vba
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Application.DisplayAlerts = False
    ActiveWorkbook.Saved = True
End Sub
```

---

### 6.2 Sheet1.cls（「業者一覧」シートイベント）

#### `Worksheet_Change(Target As Range)`

**処理概要**: G列（列7）の5行目以降が変更されたとき、同行のI列（報告書要不要フラグ）が0の場合に「報告書発行が不要」とメッセージを表示し、セル値を「不要」に書き戻す。

**処理フロー**:
1. `Application.EnableEvents = False` でイベント無効化
2. 変更セルが 5行目以降 かつ G列（列7）かを確認
3. 同行の I列（列9）が 0 なら G列を「不要」に強制設定してメッセージ表示
4. `Application.EnableEvents = True` でイベント再有効化

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

#### Public 変数（モジュールレベル）

```vba
Public mpApMotoPath As String
Public mpTxMotoPath As String
Public mpApSakiPath As String
Public mpTxSakiPath As String
Public Const mpFnameSyu As String = "\引取業者.txt"
```

> **注意**: これらの Public 変数・定数はすべて **本ファイル内で未使用**。他ファイルからのコピー残存と推定される。

#### `引取報告書登録()`

**処理概要**: 「業者一覧」シートをスキャンし、出荷方法（E列≠F列）または報告書発行種別（G列≠H列）に差分がある行のみ `ExSeihinJ` テーブルに `hikitori` / `housyube` を INSERT/UPDATE する。

**処理フロー**:
1. カラム配列を定義: `kaisyacd`, `hikitori`, `housyube`
2. 「業者一覧」シートの 5行目から B列が空になるまでループ
3. E列 ≠ F列 または G列 ≠ H列 の行を検出
4. `kaisyacd`（B列）/ `hikitori`（E列）/ `housyube`（G列）を引数に `SQL_INSERT_UPDATE()` を呼び出し
5. エラーがあれば異常メッセージ、なければ「更新しました。」を表示

```vba
Sub 引取報告書登録()
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

**処理概要**: 保存確認ダイアログを非表示にしてブックを閉じる。最後のブックの場合は Excel ごと終了する。

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

**処理概要**: DB の `tokumst`・`ExSeihinJ`・`sehmst` からデータを取得し、「業者一覧」シートの 5行目以降に展開する。報告書不要（フラグ 0 または空）の行は G・H列を「不要」と表示する。

**処理フロー**:
1. `HikitoriTB`・`HouTB` を直接クリア（`Range("HikitoriTB") = ""`）
2. SQL 1: `tokumst` と `ExSeihinJ` の外部結合で会社コード・略称・会社名・出荷方法・報告書発行種別を 7列取得 → `Disp_Sheet()` で B列〜H列に貼り付け
3. SQL 2: `sehmst` から `kaisyacd` 別の `max(syouho)` を集計 → `Disp_Sheet()` で K列〜L列に貼り付け
4. 表計算を手動モードに切替（速度向上）
5. I列の VLOOKUP 結果（`sehmst` フラグ）が 0 または空の行の G・H列を「不要」に書き換え
6. 2000行超でエラー停止
7. 表計算を自動モードに復旧、イベント再有効化

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

#### `報告書不要表示()` ⚠ 孤立コード

**処理概要**: H列（列8）が 0 の行の F・G列（列6・7）を「不要」に設定する。

> **⚠ 注意**: 本プロシージャは呼出元が存在しない孤立コードです。
> `業者名と引取抽出()` 内にインラインで同等処理が実装されていますが、
> 列参照が異なります（本プロシージャは H→F・G列、現行版は I→G・H列）。
> 旧レイアウト時代の残存コードと推定されます。
> また、ループ上限が 1000 に達すると `Stop` で VBA エディタが開く（デバッグ用コード残存）。

```vba
Sub 報告書不要表示()
    Dim i As Integer
    With Sheets("業者一覧")
        i = 5
        Do Until .Cells(i, 2) = ""
            If .Cells(i, 8) = 0 Then
                .Cells(i, 6) = "不要"
                .Cells(i, 7) = "不要"
            End If
            i = i + 1
            If i > 1000 Then Stop
        Loop
    End With
End Sub
```

---

### 6.6 ユーティリティ.bas

| ✓ | プロシージャ | 処理概要 |
|---|---|---|
|   | `イベント有効()` | `Application.EnableEvents = True` に強制復旧（ショートカット `r`） |
|   | `画面クリア()` | `HikitoriTB` と `HouTB` を空白にクリア（⚠ 呼出元なし。`業者名と引取抽出()` 内で同等処理がインライン実行されている） |

---

### 6.7 画面クリア引取業者.bas

#### `画面消去処理()`

**処理概要**: `HikitoriTB`・`HouTB` をクリアして E5 セルに移動する。ショートカットキー `e` で起動。

```vba
Sub 画面消去処理()
    Range("HikitoriTB") = ""
    Range("HouTB") = ""
    Range("E5").Select
End Sub
```

> **補足**: コメントアウトされたシート保護の解除・再設定コードが残存している（`ActiveSheet.Unprotect` / `ActiveSheet.Protect`）。

---

### 6.8 SQL_Execution.bas

ADODB による Oracle DB 接続・SQL 実行の共通ライブラリモジュール。他の EXメニューファイルでも同一実装が使用されている。

**接続文字列**: `DSN=ricdb;UID=ric;PWD=t6101`

| ✓ | プロシージャ | 処理概要 |
|---|---|---|
| ✓ | `Open_oraconDB()` | ADODB で Oracle DB に接続（`adUseClient` カーソル） |
| ✓ | `SQL_Exe()` | SQL 文を `Execute` で実行。エラー時に `Stop` ステートメント（デバッグ用残存） |
| ✓ | `SQL_INSERT_UPDATE()` | キーで件数チェック → 0件なら INSERT、1件以上なら UPDATE を動的生成・実行。トランザクション制御あり |
|   | `SQL_Delete()` | DELETE 文を生成・実行（本ファイルでは未使用） |
| ✓ | `Disp_Sheet()` | SQL 結果を `CopyFromRecordset` で指定シート・指定行列に貼り付け。ヘッダー出力オプションあり |
|   | `Set_Array()` | SQL 結果を VBA 配列に格納（本ファイルでは未使用） |

---

### 6.9 画面操作1.bas

**概要**: Excel VBA の各種 API 操作サンプルコードを集めた参考モジュール。プロダクションロジックではなく開発用リファレンスとして格納されている。

| プロシージャ | 処理概要 |
|---|---|
| `Reidai()` | セル移動・ウィンドウ最大化・Zoom・イベント制御・画面更新制御のサンプル集 |
| `印刷パラメータ設定()` | 指定シートの印刷余白・向き・用紙サイズを一括設定 |
| `DriveSearch()` | `FileSystemObject` でドライブ存在確認 |
| `シート保護()` | 各種シート保護パターンのサンプル |
| `複数列の選択()` | 複数列選択サンプル |
| `セル名の定義_削除()` | 名前付き範囲の定義・削除サンプル |
| `リボン操作()` | `ExecuteExcel4Macro` でリボン表示制御 |
| `セルのコピー_値の貼り付け()` | セルコピー・値貼り付けサンプル |

---

## 7. ユーザーフォーム仕様

該当なし（本ファイルにユーザーフォームは存在しない）。

---

## 8. DB 接続・外部連携

### 8.1 ODBC 接続設定

| DSN 名 | UID | PWD | 用途 |
|---|---|---|---|
| `ricdb` | ric | t6101 | Oracle DB 接続（照射管理システム） |

### 8.2 テーブル一覧（参照/更新区分付き）

> ✓ = INSERT / UPDATE / DELETE の対象テーブル（参照のみのテーブルは ✓ なし）

| ✓ | テーブル名 | 区分 | 主な用途 | キー列 | 参照/更新列 |
|---|---|---|---|---|---|
| ✓ | `ExSeihinJ` | **参照＋更新** | 出荷方法・報告書発行種別の登録先 | `kaisyacd` | 参照: `hikitori`, `housyube`（`業者名と引取抽出()` で SELECT）。更新: `hikitori`, `housyube`（`引取報告書登録()` で INSERT/UPDATE） |
|   | `tokumst` | 参照 | 得意先マスタ | `kaisyacd` | `kaisyacd`, `kairname`, `coname` |
|   | `sehmst` | 参照 | 製品仕様台帳（報告書要不要フラグ） | `kaisyacd` | `kaisyacd`, `syouho`（MAX集計） |

### 8.3 SQL（全件）

```sql
-- 業者一覧取得（tokumst + ExSeihinJ の外部結合）
-- 呼出元: 業者名と引取抽出() / デーた抽出.bas
SELECT t.kaisyacd, t.kairname, t.coname,
       s.hikitori, s.hikitori, s.housyube, s.housyube
FROM tokumst t, ExSeihinJ s
WHERE t.kaisyacd = s.kaisyacd(+)
  AND t.kaisyacd < '2000'
ORDER BY t.kaisyacd

-- 報告書要不要フラグ集計（sehmst）
-- 呼出元: 業者名と引取抽出() / デーた抽出.bas
SELECT kaisyacd, MAX(syouho)
FROM sehmst
GROUP BY kaisyacd
ORDER BY kaisyacd

-- 件数チェック（INSERT/UPDATE 判定用）
-- 呼出元: SQL_INSERT_UPDATE() / SQL_Execution.bas
SELECT COUNT(*) FROM ExSeihinJ WHERE kaisyacd = '<kaisyacd>'

-- 出荷方法・報告書発行種別の UPDATE
-- 呼出元: SQL_INSERT_UPDATE() / SQL_Execution.bas（引取報告書登録() 経由）
UPDATE ExSeihinJ
SET hikitori = '<hikitori>', housyube = '<housyube>'
WHERE kaisyacd = '<kaisyacd>'

-- 出荷方法・報告書発行種別の INSERT（未登録時）
-- 呼出元: SQL_INSERT_UPDATE() / SQL_Execution.bas（引取報告書登録() 経由）
INSERT INTO ExSeihinJ (kaisyacd, hikitori, housyube)
VALUES ('<kaisyacd>', '<hikitori>', '<housyube>')
```

---

## 9. データフロー

### 9.1 起動・データ表示フロー

| # | 起点 | 処理 | 結果 |
|---|---|---|---|
| 1 | 📄 ブックを開く（`Workbook_Open`） | 📄 `Workbook_Open()` / **ThisWorkbook.cls** | シート保護解除 → UIのみ保護再設定 |
| 2 | 上記 | 📄 `業者名と引取抽出()` を Call / **デーた抽出.bas** | データ取得処理を開始 |
| 3 | 上記 | 📊 `Range("HikitoriTB") = ""` / `Range("HouTB") = ""` | 既存データをクリア |
| 4 | 上記 | 🗄️ SQL実行: `tokumst` + `ExSeihinJ` 外部結合 | 全得意先の出荷方法・報告書発行種別を取得 |
| 5 | 上記 | 📊 `Disp_Sheet()` で B〜H列に貼り付け / **SQL_Execution.bas** | 業者一覧テーブルにデータ展開 |
| 6 | 上記 | 🗄️ SQL実行: `sehmst` から `MAX(syouho)` 集計 | 報告書要不要フラグを取得 |
| 7 | 上記 | 📊 `Disp_Sheet()` で K〜L列に貼り付け | `HouTB` にフラグデータ展開 |
| 8 | 上記 | 📊 I列 VLOOKUP が自動計算 → フラグ 0 or 空の行の G・H列を「不要」に設定 | 報告書不要の業者は編集不可表示 |

#### ツリー図（補助）

```
（ブックを開く）
└─ 📄 Workbook_Open                           [ThisWorkbook.cls]
   ├─ 📊 シート保護解除 → UIのみ保護再設定
   └─ 📄 業者名と引取抽出                     [デーた抽出.bas]
      ├─ 📊 HikitoriTB / HouTB クリア
      ├─ 🗄️ SELECT tokumst + ExSeihinJ        [SQL_Execution.bas]
      │   └─ 📊 Disp_Sheet → B〜H列に貼り付け
      ├─ 🗄️ SELECT sehmst MAX(syouho)         [SQL_Execution.bas]
      │   └─ 📊 Disp_Sheet → K〜L列に貼り付け
      └─ 📊 I列フラグ 0/空 → G・H列「不要」設定
```

### 9.2 セル変更フロー

| # | 起点 | 処理 | 結果 |
|---|---|---|---|
| 1 | 📊 G列（報告書発行種別）を変更 | 📄 `Worksheet_Change()` / **Sheet1.cls** | イベント無効化 → 変更セル検査 |
| 2 | [条件: 同行 I列 = 0] | 📊 G列を「不要」に強制上書き | セル値を「不要」に設定 |
| 3 | 上記 | 🖥️ MsgBox「製品仕様台帳の報告書発行が不要…」 | 警告メッセージ表示 |
| 4 | [条件: 同行 I列 ≠ 0] | — | そのまま編集値を保持 |

#### ツリー図（補助）

```
（G列セル変更）
└─ 📄 Worksheet_Change                        [Sheet1.cls]
   └─ [I列 = 0 の場合]
      ├─ 📊 G列 → 「不要」に上書き
      └─ 🖥️ MsgBox 警告表示
```

### 9.3 登録フロー

| # | 起点 | 処理 | 結果 |
|---|---|---|---|
| 1 | 📊（変更を登録する ボタン Click） | 📄 `引取報告書登録()` / **登録.bas** | 差分検出ループ開始 |
| 2 | 上記 | 📊 E列 ≠ F列 または G列 ≠ H列 の行を検出 | 変更行を特定 |
| 3 | 上記 | 🗄️ `SQL_INSERT_UPDATE()` / **SQL_Execution.bas** | `ExSeihinJ` に `hikitori` / `housyube` を UPDATE（既存行）or INSERT（新規行） |
| 4 | 上記 | 🖥️ MsgBox「更新しました。」or「異常が発生しました」 | 完了メッセージ表示 |

#### ツリー図（補助）

```
（変更を登録する ボタン Click）
└─ 📄 引取報告書登録                           [登録.bas]
   ├─ 📊 行5〜 ループ: E≠F or G≠H を検出
   ├─ 🗄️ SQL_INSERT_UPDATE                    [SQL_Execution.bas]
   │   ├─ 🗄️ SELECT COUNT(*) → 件数チェック
   │   └─ [件数 0] 🗄️ INSERT / [件数 1+] 🗄️ UPDATE
   └─ 🖥️ MsgBox 結果表示
```

### 9.4 終了フロー

| # | 起点 | 処理 | 結果 |
|---|---|---|---|
| 1 | 📊（終了 ボタン Click） | 📄 `Bookを閉じる()` / **終了処理.bas** | アラート非表示 |
| 2 | 上記 | 📄 ブック数判定 | 1冊なら `Application.Quit`、複数なら `ActiveWorkbook.Close` |
| 3 | — | 📄 `Workbook_BeforeClose()` / **ThisWorkbook.cls** | `Saved = True` で保存確認スキップ |

#### ツリー図（補助）

```
（終了 ボタン Click）
└─ 📄 Bookを閉じる                            [終了処理.bas]
   ├─ 📄 DisplayAlerts = False
   └─ [Workbooks.Count = 1] 📄 Application.Quit
      [Workbooks.Count > 1] 📄 ActiveWorkbook.Close
         └─ 📄 Workbook_BeforeClose            [ThisWorkbook.cls]
            └─ 📄 Saved = True
```

---

## 10. セキュリティ注意事項

olevba 解析結果:

| 種別 | キーワード | 内容 |
|---|---|---|
| AutoExec | `Workbook_Open` | ブックを開くと自動的に `業者名と引取抽出()` が実行され DB 接続が発生する |
| AutoExec | `Workbook_Activate` | ブックがアクティブになるたびに処理が実行される（**画面操作1.bas** のサンプルコード内） |
| AutoExec | `Workbook_BeforeClose` | ブックを閉じる前にアラート非表示・保存済みフラグ設定が自動実行される |
| AutoExec | `Worksheet_Change` | G列セル変更時に自動でバリデーション処理が実行される |
| Suspicious | `Environ` | **画面操作1.bas** で `Environ("COMPUTERNAME")` を使用（PC名取得） |
| Suspicious | `Open` | ファイルオープン操作の可能性（**登録.bas** の定数 `mpFnameSyu` に関連、ただし実際のファイルI/Oコードは未使用） |
| Suspicious | `Call` | Excel 4 マクロ形式の DLL 呼び出しの可能性 |
| Suspicious | `CreateObject` | **画面操作1.bas** で `Scripting.FileSystemObject` を動的生成 |
| Suspicious | `ExecuteExcel4Macro` | **画面操作1.bas** でリボン操作に Excel 4 マクロを使用 |
| Suspicious | `Windows` | アプリケーションウィンドウ列挙の可能性 |
| Suspicious | `Chr` | 文字列内に `Chr(13)` 使用（CRLF として使用） |
| Suspicious | `Hex Strings` | VBA 内に 16進数エンコード文字列が存在 |
| Suspicious | `Base64 Strings` | VBA 内に Base64 エンコード文字列が存在 |

> **注意**: DSN=ricdb の接続文字列（UID/PWD）が **SQL_Execution.bas** にハードコードされています。パスワード `t6101` が平文で記述されているため、コードの漏洩に注意してください。

> **注意**: **SQL_Execution.bas** の `SQL_Exe()` 内にデバッグ用の `Stop` ステートメントが残存しています。エラー発生時に VBA エディタが開くため、本番環境では削除を推奨します。

> **注意**: **デーた抽出.bas** の `報告書不要表示()` 内にもデバッグ用の `Stop` ステートメントが残存しています（ループ上限 1000 到達時）。

> **注意**: **登録.bas** に `Public Const mpFnameSyu As String = "\引取業者.txt"` および複数の Public パス変数が定義されていますが、本ファイル内では一切使用されていません。他ファイルからのコピー残存と推定されます。

> **注意**: **画面操作1.bas** は開発用サンプルコードを含む参考モジュールです。本番環境では不要なコードが含まれているため、定期的な棚卸しが推奨されます。

---

## スコープ外（本仕様書に含まないもの）

- セル書式（色・罫線・フォント）
- 条件付き書式、グラフ・画像、印刷設定

必要な場合は Excel 画面のスクリーンショットで補完してください。
