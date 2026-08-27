# ExRic3詰替作業 仕様書

> **ファイル種別**: .xlsm（マクロ付き）
> **用途**: 照射済み製品の積替え（詰替え）作業管理。DBから積替え対象製品を抽出・一覧表示し、積替え要/不要フラグの登録更新を行う
> **VBA プロジェクト**: モジュール 12 本（.bas 8 / .cls 4 / .frm 0）
> **外部連携**: DSN=ricdb（Oracle）
> **解析日**: 2026-06-29

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
| ✓ | 1 | 積替品 | 235 | 15 | visible | — |
| ✓ | 2 | 積替TB | 10000 | 12 | visible | — |
|  | 3 | WorkTB | 133 | 9 | visible | — |

### 1.2 ユーザーフォーム一覧

なし

### 1.3 VBA モジュール一覧

> ✓ = ユーザー操作の起点 / DB I/O を含む / 他モジュールから呼び出される / コード行数上位 25%


| ✓ | No | モジュール | 種別 | プロシージャ数 | 主な役割 |
| --- | --- | --- | --- | --- | --- |
| ✓ | 1 | **ThisWorkbook** | .cls | 2 | ブック開閉時イベント（初期化・画面クリア・積替品抽出呼び出し） |
| ✓ | 2 | **SQL_Execution** | .bas | 6 | DB接続・SQL実行・シート転記・配列格納の汎用ルーチン |
| ✓ | 3 | **Ex画面クリア** | .bas | 2 | 名前付き範囲のクリア・印刷範囲初期化・シート遷移 |
| ✓ | 4 | **Ex詰替品更新** | .bas | 1 | 積替TB上の変更差分をDBテーブル`ExSeihinZ`へINSERT/UPDATE |
| ✓ | 5 | **Ex積替え品抽出** | .bas | 2 | DBから積替え対象品を抽出しWorkTB経由で積替品シートに加工表示 |
| ✓ | 6 | **Ex積替品表示** | .bas | 1 | 積替TBシートに全製品マスタ＋積替えフラグを表示 |
|  | 7 | **ExFunction** | .bas | 1 | 日付フォーマット変換ユーティリティ |
|  | 8 | **印刷範囲** | .bas | 1 | 積替品シートの印刷範囲設定 |
|  | 9 | **終了処理** | .bas | 1 | ブックを閉じる（上書きなし） |
|  | 10 | **Sheet4** | .cls | 0 | 「積替TB」シートモジュール（コードなし） |
|  | 11 | **Sheet5** | .cls | 0 | 「積替品」シートモジュール（コードなし） |
|  | 12 | **Sheet6** | .cls | 0 | 「WorkTB」シートモジュール（コードなし） |

---

## 2. シート詳細

### 2.0 シート可視性一覧


| No | シート | VBA による非表示化 | 表示するタイミング | 非表示にするタイミング | 制御プロシージャ |
| --- | --- | --- | --- | --- | --- |
| 1 | 積替品 | — | — | — | — |
| 2 | 積替TB | — | — | — | — |
| 3 | WorkTB | — | — | — | — |


> 以下の各シートのレイアウト構造表における ✓ = VBA から `Range()` で代入または参照され、業務ロジックに直結するセル

### 2.0 b 非表示行・列一覧

| シート | 非表示行 | 非表示列 |
| --- | --- | --- |
| 積替品 | 4 | C |
| 積替TB | 2〜3 | B〜C, J |

### 2.1 積替品

**目的**: 積替え品の照射状況を一覧表示し、印刷する。

#### 非表示行・列

なし。

#### ヘッダ領域（行1〜4）

| ✓ | No | セル | 内容 | 備考 |
| --- | --- | --- | --- | --- |
|  | 1 | `B1` | "New" | ※推論: バージョンラベル |
|  | 2 | `B3` | "積替え品一覧" | シートタイトル |
| ✓ | 3 | `F3` | 32 | 印刷範囲行数（名前付き範囲`Innsatu`） |
|  | 4 | `G3:H3` | "まで印刷　Max" | 結合セル |
|  | 5 | `I3` | `=MAX(B6:B235)` | 現在のNo最大値 |
|  | 6 | `L3` | `=NOW()` | 現在日時 |
|  | 7 | `M3` | "Ex4" | ※推論: システム識別子 |

#### データ領域ヘッダ（行5）

| ✓ | No | 列 | ヘッダ | 業務的意味 |
| --- | --- | --- | --- | --- |
|  | 1 | B | No | 連番 |
| ✓ | 2 | D | 最終積替え日 | 出荷日-1日で算出 |
| ✓ | 3 | E | 照射状況 | i/0/C→未、1→中、2→済 に変換 |
|  | 4 | F | 線量計番号 | 下4桁表示 |
|  | 5 | G | 受付番号 | 下4桁表示 |
|  | 6 | H | 会社名 | 「株式会社」除去済み |
|  | 7 | I | 納期 | mm/dd形式 |
| ✓ | 8 | J | 出荷日 | DB値をDate型変換 |
|  | 9 | K | パス数 | 照射パス数 |
|  | 10 | L | 備考 | DB bikou1 |
|  | 11 | M | チェック | 作業確認用 |

#### データ領域

| ✓ | No | 範囲 | 内容 |
| --- | --- | --- | --- |
| ✓ | 1 | `B6:M235` | 積替え品一覧データ（名前付き範囲`TumikaeTB`） |

### 2.2 積替TB

**目的**: 積替え対象の製品マスタを表示・編集する。

#### 非表示行・列

なし。

#### ヘッダ領域

| ✓ | No | セル | 内容 | 備考 |
| --- | --- | --- | --- | --- |
|  | 1 | `E1` | "積み替え製品テーブル" | シートタイトル |
| ✓ | 2 | `D3` | `=COUNTA(SeihinnSuu)` | 登録製品数 |
|  | 3 | `D4` | "登録製品毎に積替え品登録してください" | 操作案内 |

#### データ領域ヘッダ（行5）

| ✓ | No | 列 | ヘッダ | 業務的意味 |
| --- | --- | --- | --- | --- |
| ✓ | 1 | D | 会社コード | 4桁ゼロ埋め |
| ✓ | 2 | E | 製品コード | 3桁ゼロ埋め |
|  | 3 | F | 会社コード&製品コード | 結合キー |
|  | 4 | G | 会社名 | tokumst.coname |
|  | 5 | H | 製品名 | sehmst.seiname |
| ✓ | 6 | I | 詰替え要不要 | ユーザー編集対象（"1"=要） |
| ✓ | 7 | J | ExSeihin Kaisyacd | DB現在値（比較用） |

#### データ領域

| ✓ | No | 範囲 | 内容 |
| --- | --- | --- | --- |
| ✓ | 1 | `D6:J10000` | 製品マスタデータ（名前付き範囲`TumikaeHinn`） |

#### 数式

| セル | 数式 | 業務的意味 |
| --- | --- | --- |
| `I3` | `=IF(ISERROR(VLOOKUP(#REF!,TumiFlg,2,FALSE)),"",VLOOKUP(#REF!,TumiFlg,2,FALSE))` | ※参照エラー状態（`TumiFlg`は未定義） |
| `J3` | `=IF(ISERROR(VLOOKUP(#REF!,TumiFlg,2,FALSE)),"",VLOOKUP(#REF!,TumiFlg,2,FALSE))` | 同上 |

### 2.3 WorkTB

**目的**: DB抽出結果を一時格納する作業用シート。ユーザーは直接操作しない。

#### 非表示行・列

なし。

#### セル

| セル | 内容 | 備考 |
| --- | --- | --- |
| `I1` | 1 | ※推論: フラグまたはカウンタ |

VBAからは名前付き範囲`Work`(`A1:H201`)および`Wtb`(`A1:H250`)で参照される。

---

## 3. 名前付き範囲一覧

> ✓ = VBA から `Range()` で代入または参照され、業務ロジックに直結する名前付き範囲

| ✓ | No | 名前 | 参照先 | 業務的意味 |
| --- | --- | --- | --- | --- |
|  | 1 | `DebugFlg` | 積替品!$A$1 | デバッグフラグ |
| ✓ | 2 | `Innsatu` | 積替品!$F$3 | 印刷範囲行数（デフォルト32） |
|  | 3 | `SeihinnSuu` | 積替TB!$D$6:$D$1005 | 登録製品数カウント用範囲 |
| ✓ | 4 | `SeiKennsuu` | 積替TB!$D$3 | 登録製品件数（COUNTA結果） |
|  | 5 | `TumeTB` | 積替TB!$G$6:$I$1005 | 会社名・製品名・積替えフラグの表示範囲 |
| ✓ | 6 | `TumikaeHinn` | 積替TB!$D$6:$J$10000 | 積替え製品マスタ全データ範囲 |
| ✓ | 7 | `TumikaeTB` | 積替品!$B$6:$M$235 | 積替品一覧表示データ範囲 |
| ✓ | 8 | `Work` | WorkTB!$A$1:$H$201 | ワーク領域（DB抽出一時格納） |
| ✓ | 9 | `Wtb` | WorkTB!$A$1:$H$250 | ワーク領域拡張（クリア用） |


### 3.1 データの入力規則

| シート | セル | 種別 | 制約 | 用途 |
| --- | --- | --- | --- | --- |
| 積替品 | `A2` | 整数 | =999 |  |
| 積替品 | `F3` | 整数 | 1〜230 |  |
| 積替TB | `I6:I1005` | 整数 | =1 |  |
| 積替TB | `D6:H1005` | 整数 | =9999 |  |


---

---

## 4. 数式一覧

| No | シート | セル | 数式 | 説明 |
| --- | --- | --- | --- | --- |
| 1 | 積替品 | `I3` | `=MAX(B6:B235)` | 一覧の最大No表示 |
| 2 | 積替品 | `L3` | `=NOW()` | 現在日時の表示 |
| 3 | 積替TB | `D3` | `=COUNTA(SeihinnSuu)` | 登録製品件数のカウント |
| 4 | 積替TB | `I3` | `=IF(ISERROR(VLOOKUP(#REF!,TumiFlg,2,FALSE)),"",VLOOKUP(#REF!,TumiFlg,2,FALSE))` | 参照エラー状態 |
| 5 | 積替TB | `J3` | `=IF(ISERROR(VLOOKUP(#REF!,TumiFlg,2,FALSE)),"",VLOOKUP(#REF!,TumiFlg,2,FALSE))` | 参照エラー状態 |

---

## 5. ボタン・マクロ対応

> ✓ = DB 更新・画面遷移・計算実行など副作用のある操作を起動するボタン

### 5.1 シート上のボタン

#### 積替品（vmlDrawing1.vml）

| ✓ | No | シート | ボタンラベル | 割り当てマクロ | 動作概要 |
| --- | --- | --- | --- | --- | --- |
| ✓ | 1 | 積替品 | 現状の積替品の照射状況 | `TumikaeHinn()` | DBから積替え対象品を抽出し一覧表示 |
| ✓ | 2 | 積替品 | 積替する製品の登録 | `積替製品TB表示()` | 積替TBシートに製品マスタを表示し遷移 |
|  | 3 | 積替品 | 画面を閉じる | `Bookを閉じる()` | ブックを保存せず閉じる |
|  | 4 | 積替品 | ←印刷範囲設定は…（描画図形） | `InsatuHanni()` | 名前付き範囲 `Innsatu` の値に基づき印刷範囲を `$B$6:$M$n` に設定（デフォルト32行） |

#### 積替TB（vmlDrawing2.vml）

| ✓ | No | シート | ボタンラベル | 割り当てマクロ | 動作概要 |
| --- | --- | --- | --- | --- | --- |
|  | 1 | 積替TB | 戻る | `Modori()` | 積替品シートに戻る |
| ✓ | 2 | 積替TB | 登録内容更新 | `詰替品データ更新()` | 変更行をDBへINSERT/UPDATE |
| ✓ | 3 | 積替TB | データ表示製品検索 | `積替製品TB表示()` | 製品マスタを再取得表示 |

### 5.2 ショートカットキー

| No | マクロ名 | ショートカット | 処理概要 |
| --- | --- | --- | --- |
| 1 | `画面クリア3詰替()` | **Ctrl+E** | 画面データクリア・初期状態復帰 |

### 5.3 ユーザーフォーム上のボタン（サマリ）

なし。

### 5.4 CommandBar

コメントアウト済み（`Application.CommandBars("Worksheet Menu Bar").Enabled = True`）。現在は無効。

---

## 6. VBAモジュール仕様

### 6.0 全プロシージャ一覧

> ✓ = ユーザー操作の起点（Click イベント等） / DB I/O を実行 / 他モジュールから呼び出される Public


| ✓ | No | モジュール | プロシージャ | スコープ | 種別 | 概要 |
| --- | --- | --- | --- | --- | --- | --- |
|  | 1 | **ThisWorkbook** | `Workbook_BeforeClose()` | Private | Event | 保存ダイアログを抑止して閉じる |
| ✓ | 2 | **ThisWorkbook** | `Workbook_Open()` | Private | Event | 起動時に画面クリアと積替品抽出を実行 |
| ✓ | 3 | **SQL_Execution** | `Open_oraconDB()` | Public | Sub | ODBC で DB 接続を開く |
| ✓ | 4 | **SQL_Execution** | `SQL_Exe()` | Public | Sub | SQL 文を Execute で実行 |
| ✓ | 5 | **SQL_Execution** | `SQL_INSERT_UPDATE()` | Public | Sub | キー存在チェック付き INSERT/UPDATE |
| ✓ | 6 | **SQL_Execution** | `SQL_Delete()` | Public | Sub | WHERE 条件で DELETE |
| ✓ | 7 | **SQL_Execution** | `Disp_Sheet()` | Public | Sub | SQL 結果をシートに転記 |
| ✓ | 8 | **SQL_Execution** | `Set_Array()` | Public | Sub | SQL 結果を配列に格納 |
| ✓ | 9 | **Ex画面クリア** | `画面クリア3詰替()` | Public | Sub | 画面データクリア・初期状態復帰 |
|  | 10 | **Ex画面クリア** | `modori()` | Public | Sub | 積替品シートに戻る |
| ✓ | 11 | **Ex詰替品更新** | `詰替品データ更新()` | Public | Sub | 変更行を ExSeihinZ へ INSERT/UPDATE |
| ✓ | 12 | **Ex積替え品抽出** | `TumikaeHinn()` | Public | Sub | DB から積替え対象品を抽出し一覧表示 |
| ✓ | 13 | **Ex積替え品抽出** | `DataKakou()` | Public | Sub | 抽出データを加工してシートに表示 |
| ✓ | 14 | **Ex積替品表示** | `積替製品TB表示()` | Public | Sub | 製品マスタと積替えフラグを積替TBに表示 |
|  | 15 | **ExFunction** | `ExchengeDATE()` | Public | Function | 日付フォーマット変換 |
|  | 16 | **印刷範囲** | `InsatuHanni()` | Public | Sub | 印刷範囲を設定 |
|  | 17 | **終了処理** | `Bookを閉じる()` | Public | Sub | ブックを保存せず閉じる |

### 6.1 **ThisWorkbook**（ThisWorkbook.cls）

#### `Workbook_BeforeClose(Cancel As Boolean)`
- DisplayAlertsをFalseにし保存済みフラグを設定して無警告で閉じる

#### `Workbook_Open()`
- ブック保護を解除
- ウィンドウ最大化
- 「積替品」シートを選択
- `画面クリア3詰替()` を呼び出し（全範囲クリア）
- `TumikaeHinn()` を呼び出し（DB抽出→一覧表示）

### 6.2 **SQL_Execution**（SQL_Execution.bas）

モジュール変数:
- `mpErrDes As String` — エラー記述（Public）
- `oraconn As New ADODB.Connection` — DB接続オブジェクト
- `rs As ADODB.Recordset` — レコードセット

#### `Open_oraconDB()`
- ODBC DSN `ricdb` でDB接続を確立
- CursorLocationを`adUseClient`に設定

#### `SQL_Exe(mySQL As String)`
- 引数のSQL文をExecuteで実行し結果を`rs`に格納
- エラー発生時は`mpErrDes`にエラー内容を記録

#### `SQL_INSERT_UPDATE(myTBL, myKey, myD(), myN)`
- テーブル`myTBL`に対し、`myKey`条件でレコード存在チェック
- 存在しない場合: INSERT文を動的生成し実行
- 存在する場合: UPDATE文を動的生成し実行
- トランザクション制御あり（BeginTrans→CommitTrans）

#### `SQL_Delete(myTBL, myWhere)`
- テーブル`myTBL`から`myWhere`条件でDELETE実行
- トランザクション制御あり

#### `Disp_Sheet(mySQL, mySH, myRow, myRecordCount, myColumn, myFieldCount, myF)`
- SQL実行結果をシート`mySH`の`myRow`行`myColumn`列から転記
- `myF=1`の場合はフィールド名ヘッダも出力
- `CopyFromRecordset`でデータを一括貼り付け
- 出力パラメータ: `myRecordCount`（レコード数）, `myFieldCount`（フィールド数）

#### `Set_Array(mySQL, myData(), myRecordCount, myFldCount)`
- SQL実行結果を二次元配列`myData(i,j)`に格納
- i=レコード番号, j=フィールド番号

### 6.3 **Ex画面クリア**（Ex画面クリア.bas）

#### `画面クリア3詰替()`
- `TumikaeHinn`, `TumikaeTB`, `Work` をクリア
- `Innsatu` を32にリセット
- `InsatuHanni()` を呼び出し

#### `modori()`
- 「積替品」シートを選択（画面遷移）

### 6.4 **Ex詰替品更新**（Ex詰替品更新.bas）

#### `詰替品データ更新()`
- 「積替TB」シートの`TumikaeHinn`範囲をループ
- I列（ユーザー編集値）とJ列（DB現在値）を比較
- 差分がある行のみ`ExSeihinZ`テーブルへINSERT/UPDATE
- キー: `kaisyacd`（4桁ゼロ埋め）＋ `sehncd`（3桁ゼロ埋め）
- 更新カラム: `tumikae`（積替えフラグ）
- 更新後にJ列をI列の値で上書き（差分解消）

### 6.5 **Ex積替え品抽出**（Ex積替え品抽出.bas）

#### `TumikaeHinn()`
- `Wtb`範囲をクリア
- SQL: `syoukj3`, `zaiko`, `ExKeikakuX`, `ExSeihinZ` を結合
- 条件: `ExSeihinZ.tumikae='1'`（積替え対象のみ）
- `Disp_Sheet()`でWorkTBに転記
- `DataKakou()`を呼び出しデータ加工

#### `DataKakou(myDataN)`
- WorkTBからデータを配列に読み込み
- 加工処理:
  - 出荷日→最終積替え日（出荷日-1日）
  - 照射状況コード変換（i/0/C→未、1→中、2→済）
  - 線量計番号→下4桁
  - 受付番号→下4桁
  - 会社名→「株式会社」除去
  - 納期→mm/dd形式
- 加工後データを「積替品」シート`B6:M`に書き込み
- ソート: 照射状況（昇順）→ 最終積替え日（昇順）→ 線量計番号（昇順）

### 6.6 **Ex積替品表示**（Ex積替品表示.bas）

#### `積替製品TB表示()`
- 「積替TB」シートを選択し`TumikaeHinn`をクリア
- SQL: `sehmst`, `tokumst`, `ExSeihinZ` を結合
- 条件: `sehmst.syouso='3'` AND `sehmst.ric<>'**********'`
- 取得項目: 会社コード、製品コード、結合キー、会社名、製品名、積替えフラグ×2
- `Disp_Sheet()`で積替TBシートD6から転記

### 6.7 **ExFunction**（ExFunction.bas）

#### `ExchengeDATE(ByVal myDate, ByVal myType)`
- `myType="mm/dd"`: 数値を"MM/DD"形式に変換
- `myType="yyyy/mm/dd"`: 8桁数値を"YYYY/MM/DD"形式に変換
- 数値0またはブランクの場合はTrimのみ

### 6.8 **印刷範囲**（印刷範囲.bas）

#### `InsatuHanni()`
- 名前付き範囲`Innsatu`の値に基づき印刷範囲を`$B$6:$M$n`に設定
- 32以外の場合はメッセージボックスで通知

### 6.9 **終了処理**（終了処理.bas）

#### `Bookを閉じる()`
- DisplayAlertsをFalseに設定
- 開いているブックが1つなら`Application.Quit`
- 複数なら`ActiveWorkbook.Close`

### 6.10 **Sheet4**（Sheet4.cls）

- 「積替TB」のシートモジュール。コードなし（`Option Explicit`のみ）

### 6.11 **Sheet5**（Sheet5.cls）

- 「積替品」のシートモジュール。コードなし（`Option Explicit`のみ）

### 6.12 **Sheet6**（Sheet6.cls）

- 「WorkTB」のシートモジュール。コードなし（`Option Explicit`のみ）

---

## 7. ユーザーフォーム仕様

なし（本ブックにユーザーフォームは含まれない）

---

## 8. DB 接続・外部連携

### 8.1 ODBC 接続設定

| DSN 名 | UID | PWD | 用途 |
| --- | --- | --- | --- |
| `ricdb` | `ric` | `t6101` | 照射管理システムDB — 積替フラグの更新と照射状況の参照 |

### 8.2 テーブル一覧（参照/更新区分付き）

> ✓ = INSERT / UPDATE / DELETE の対象テーブル（参照のみのテーブルは ✓ なし）

| ✓ | No | テーブル名 | 区分 | 主な用途 | キー列 | 参照/更新列 |
| --- | --- | --- | --- | --- | --- | --- |
| ✓ | 1 | `ExSeihinZ` | **参照＋更新** | 製品別積替えフラグ管理 | `kaisyacd` + `sehncd` | 更新: `kaisyacd`, `sehncd`, `tumikae`（`詰替品データ更新()` から INSERT/UPDATE） |
|  | 2 | `syoukj3` | 参照 | 照射工程管理（照射状況・線量計・受付番号・会社名） | `uno1` | `syoush_f`, `sykno`, `uno1`, `kainame1` |
|  | 3 | `zaiko` | 参照 | 在庫管理（受付番号・会社コード・製品コード・納期・パス数） | `uno` | `nouki`, `pass`, `kaisyacd`, `sehncd` |
|  | 4 | `ExKeikakuX` | 参照 | 計画管理（出荷日・備考） | `uno` | `syukkabi`, `bikou1` |
|  | 5 | `tokumst` | 参照 | 得意先マスタ（会社名） | `kaisyacd` | `coname` |
|  | 6 | `sehmst` | 参照 | 製品マスタ（製品名・照射所・会社コード） | `kaisyacd` + `sehncd` | `seiname`, `syouso`, `ric` |

> **「キー列」の定義**: JOIN 条件または UPDATE/DELETE の WHERE 句で使用される列を示す。

### 8.3 SQL 一覧

#### 8.3.1 積替え品照射状況抽出（`TumikaeHinn()` / **Ex積替え品抽出.bas**）

```sql
SELECT s.syoush_f, s.sykno, s.uno1, s.kainame1,
       z.nouki, k.syukkabi, z.pass, k.bikou1
FROM syoukj3 s, zaiko z, ExKeikakuX k, ExSeihinz e
WHERE s.uno1 = z.uno
  AND s.uno1 = k.uno(+)
  AND z.kaisyacd = e.kaisyacd(+)
  AND z.sehncd = e.sehncd(+)
  AND e.tumikae = '1'
ORDER BY s.syoush_f DESC, s.sykno
```

#### 8.3.2 積替製品マスタ取得（`積替製品TB表示()` / **Ex積替品表示.bas**）

```sql
SELECT s.kaisyacd, s.sehncd, s.kaisyacd || s.sehncd,
       TRIM(t.coname), TRIM(s.seiname), e.tumikae, e.tumikae
FROM tokumst t, sehmst s, ExSeihinZ e
WHERE s.kaisyacd = t.kaisyacd(+)
  AND s.syouso = '3'
  AND s.kaisyacd = e.kaisyacd(+)
  AND s.sehncd = e.sehncd(+)
  AND s.ric <> '**********'
ORDER BY s.kaisyacd, s.sehncd
```

#### 8.3.3 INSERT/UPDATE（`詰替品データ更新()` / **Ex詰替品更新.bas**）

```sql
-- 存在チェック
SELECT COUNT(*) FROM ExSeihinZ WHERE kaisyacd='XXXX' AND sehncd='XXX'

-- INSERT（レコードなしの場合）
INSERT INTO ExSeihinZ (kaisyacd, sehncd, tumikae)
VALUES ('XXXX', 'XXX', '1')

-- UPDATE（レコードありの場合）
UPDATE ExSeihinZ SET kaisyacd='XXXX', sehncd='XXX', tumikae='1'
WHERE kaisyacd='XXXX' AND sehncd='XXX'
```

### 8.4 外部ファイル

なし（外部リンク・外部ファイル参照は検出されず）

---

## 9. データフロー

各フローは「起点 → 処理 → 結果」の粒度で記述する。

### 9.1 データフローテーブル

| No | 起点 | → | 終点 | トリガー | 内容 |
| --- | --- | --- | --- | --- | --- |
| 1 | 🗄️ DB (syoukj3, zaiko, ExKeikakuX, ExSeihinZ) | → | 📊「WorkTB」 | ブック起動 / ボタン「現状の積替品の照射状況」 | 積替え対象品の照射状況データ |
| 2 | 📊「WorkTB」 | → | 📄 VBA配列 (myData) | `DataKakou()` | ワーク読み込み＋加工処理 |
| 3 | 📄 VBA配列 | → | 📊「積替品」B6:M235 | `DataKakou()` | 加工済み積替え品一覧の表示 |
| 4 | 🗄️ DB (sehmst, tokumst, ExSeihinZ) | → | 📊「積替TB」D6:J | ボタン「積替する製品の登録」/ ボタン「データ表示製品検索」 | 全製品マスタ＋積替えフラグ |
| 5 | 📊「積替TB」I列（ユーザー編集） | → | 🗄️ DB ExSeihinZ | ボタン「登録内容更新」 | 差分行の積替えフラグ更新 |

### 9.2 データフローツリー図

```
📂 ExRic3詰替作業
├── 🖥️ ブック起動 (Workbook_Open)
│   ├── 📄 画面クリア3詰替() ─── 📊「積替品」「積替TB」「WorkTB」全クリア
│   └── 📄 TumikaeHinn()
│       ├── 🗄️ SELECT syoukj3+zaiko+ExKeikakuX+ExSeihinZ
│       ├── 📊「WorkTB」← DB結果転記 (Disp_Sheet)
│       └── 📄 DataKakou()
│           ├── 📊「WorkTB」→ 配列読み込み
│           ├── 📄 データ加工（日付変換・コード変換・文字列整形）
│           └── 📊「積替品」B6:M235 ← 加工済みデータ書き込み＋ソート
│
├── 🖥️ ボタン「積替する製品の登録」
│   └── 📄 積替製品TB表示()
│       ├── 🗄️ SELECT sehmst+tokumst+ExSeihinZ
│       └── 📊「積替TB」D6:J ← 製品マスタ転記 (Disp_Sheet)
│
├── 🖥️ ボタン「登録内容更新」
│   └── 📄 詰替品データ更新()
│       ├── 📊「積替TB」I列/J列 比較ループ
│       ├── 🗄️ INSERT/UPDATE ExSeihinZ (差分行のみ)
│       └── 📊「積替TB」J列 ← I列で上書き（同期）
│
├── 🖥️ ボタン「現状の積替品の照射状況」
│   └── 📄 TumikaeHinn() ─── （上記と同一フロー）
│
├── 🖥️ ボタン「戻る」
│   └── 📄 modori() ─── 📊「積替品」シート選択
│
└── 🖥️ ボタン「画面を閉じる」
    └── 📄 Bookを閉じる() ─── 🖥️ ブック終了（保存なし）
```

---

## 10. セキュリティ注意事項


| No | カテゴリ | 内容 | リスク |
| --- | --- | --- | --- |
| 1 | 認証情報ハードコード | DSN=`ricdb`, UID=`ric`, PWD=`t6101` が **SQL_Execution** モジュールに平文記載 | 中：VBAエディタで閲覧可能 |
| 2 | SQLインジェクション | `SQL_INSERT_UPDATE()`はセル値を文字列連結でSQL構築。シングルクォートのエスケープなし | 中：文字列連結によるSQL構築 |
| 3 | エラーハンドリング | `On Error Resume Next` の多用によりDB操作失敗が無視される可能性 | 中：サイレント障害の可能性 |
| 4 | トランザクション | `SQL_INSERT_UPDATE()`内で1件ずつCommitするため、ループ途中のエラーで不整合が発生する可能性 | 中：サイレント障害の可能性 |
| 5 | ブック保護 | `Workbook_Open`で`Protect Structure:=False`を実行しブック保護を解除 | 低：VBAからは無制限アクセス |
| 6 | 保存なし終了 | `Workbook_BeforeClose`で`Saved=True`を強制設定し変更破棄を無警告化 | 低：変更破棄が無警告 |

---

## スコープ外（本仕様書に含まないもの）

- セル書式（色・罫線・フォント）
- 条件付き書式、グラフ・画像、印刷設定

必要な場合は Excel 画面のスクリーンショットで補完してください。
