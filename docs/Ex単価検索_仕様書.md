# Ex単価検索 仕様書

> **対象ファイル**: Ex単価検索.xlsm
> **ファイル種別**: .xlsm（マクロ付き）
> **用途**: 受付番号・会社コード・製品コードに基づく単価検索、価格表ファイル表示、単価登録/更新
> **VBA プロジェクト**: モジュール 11 本（.bas 9 / .cls 2 / .frm 0）
> **外部連携**: DSN=ricdb（ODBC）、DB接続先 IP: DSN経由
> **解析日**: 2026-06-29（excel-to-md スキルによる自動解析）

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
| モジュール（.bas / .cls） | **太字** | **開始処理.bas** |
| プロシージャ / イベント | `コード体()` | `Start_GO()` |
| シート名 | 「」 | 「Main」 |
| セル参照 | `コード体` | `$B$5` |
| 名前付き範囲 | `コード体` | `Uno` |
| DB テーブル / カラム | `コード体` | `ExSeihinZ` / `tanka` |
| ユーザー操作 | （操作名） | （単価登録ボタン Click） |
| 主要マーク | ✓ | ✓ = 保守時に最初に確認すべき項目 |

### データフロー 場所マーク（9章）

9章のデータフロー（テーブル・ツリー図）では、処理が行われる場所を以下のアイコンで区別します。

| アイコン | 種別 | 意味 |
| --- | --- | --- |
| 📊 | シート操作 | ワークシート上のセル書込み・読取り・表示変更 |
| 🖥️ | 画面操作 | ユーザーフォーム（.frm）の表示・入力・操作 |
| 🗄️ | DB操作 | DB への SELECT / INSERT / UPDATE / DELETE |
| 📄 | VBA内部処理 | 変数計算・条件分岐など、画面・シートに直接関与しない処理 |

### ✓（主要マーク）の判定基準

✓ は **保守時に最初に確認すべき項目** を示します。
判定基準は対象の種類ごとに以下のとおりです。

| 章 | 対象 | ✓ の判定基準 |
| --- | --- | --- |
| 1.1 | シート | ユーザーが直接操作する、または VBA が動的に表示/非表示を切り替える |
| 1.3 / 6.0 | VBA モジュール | ① ユーザー操作の起点 ② DB I/O を含む ③ 他モジュールから呼び出される ④ コード行数上位 25% のいずれか |
| 2 | セル / 名前付き範囲 | VBA から `Range()` で代入または参照され、業務ロジックに直結する |
| 3 | 名前付き範囲 | VBA から `Range()` で代入または参照され、業務ロジックに直結する |
| 5 | ボタン | DB 更新・画面遷移・計算実行など副作用のある操作を起動する |
| 6.0 | プロシージャ | ① ユーザー操作の起点（Click イベント等） ② DB I/O を実行 ③ 他モジュールから呼び出される Public のいずれか |
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

### 1.1 シート一覧

> ✓ = ユーザーが直接操作する、または VBA が動的に表示/非表示を切り替えるシート

| ✓ | No | シート名 | 最大行 | 最大列 | 保存時 Visible | VBA による動的切替 |
| --- | --- | --- | --- | --- | --- | --- |
| ✓ | 1 | Main | 2000 | 25 (Y) | visible | — |

### 1.2 ユーザーフォーム一覧

なし。本ファイルにユーザーフォームは含まれない。

### 1.3 VBA モジュール一覧

> ✓ = ユーザー操作の起点 / DB I/O を含む / 他モジュールから呼び出される / コード行数上位 25%
> 全モジュールを列挙し、✓ 基準に該当するものにマークを付ける。

| ✓ | No | モジュール | 種別 | プロシージャ数 | 主な役割 |
| --- | --- | --- | --- | --- | --- |
| ✓ | 1 | **Sheet1.cls** | .cls | 2 | 「Main」シートイベント |
| ✓ | 2 | **ThisWorkbook.cls** | .cls | 3 | 起動/終了イベント |
| ✓ | 3 | **SQL_Execution.bas** | .bas | 6 | DB接続・SQL実行共通処理 |
| ✓ | 4 | **データベースR.bas** | .bas | 3 | 受付番号・会社・製品のDB読込 |
| ✓ | 5 | **表示.bas** | .bas | 4 | 受付番号表示・会社データ表示・製品データ表示・価格表ファイル表示 |
| ✓ | 6 | **開始処理.bas** | .bas | 2 | BookOpen時の認証・画面初期化 |
| ✓ | 7 | **単価_ファイル登録処理.bas** | .bas | 2 | フォルダー/ファイル名登録・単価登録 |
| ✓ | 8 | **会社コード検索.bas** | .bas | 1 | 略称による会社コード検索 |
|  | 9 | **画面クリア.bas** | .bas | 5 | 画面消去・Object消去 |
|  | 10 | **画面設定解除.bas** | .bas | 2 | CommandBar表示制御・フルスクリーン設定 |
|  | 11 | **ユーティリティ.bas** | .bas | 2 | イベント有効化・セル書式設定 |
|  | 12 | **共通変数.bas** | .bas | 0 | Public変数宣言（mpTankaF, mpUno） |
|  | 13 | **終了処理.bas** | .bas | 1 | ブック閉じる処理 |

> **Sheet*.cls の記載ルール**: Sheet1.cls は「Main」シートに対応する（`sheetPr codeName="Sheet1"` と `sheet name="Main"` の突合による）。

---

## 2. シート詳細

### 2.0 シート可視性一覧

| No | シート | VBA による非表示化 | 表示するタイミング | 非表示にするタイミング | 制御プロシージャ |
| --- | --- | --- | --- | --- | --- |
| 1 | Main | — | — | — | — |

> 以下の各シートのレイアウト構造表における ✓ = VBA から `Range()` で代入または参照され、業務ロジックに直結するセル

### 2.1 Main

**目的**: 単価検索の操作画面。受付番号・会社コード・製品コード入力による単価情報表示と、単価登録・価格表ファイル登録を行う。

#### 非表示行・列

| 種類 | 対象 | 備考 |
| --- | --- | --- |
| 非表示列 | T〜X列（5列） | 設定値・リスト格納領域（PathFolder, FolderName, TanniList, Ryaku 等） |

#### レイアウト構造

```
行 1: システム情報行（バージョン）
行 2: 日付表示
行 3: タイトル・社員名・フォルダー名・ファイル名
行 4: 列見出し行（受付番号〜入り数）
行 5: データ入力行（検索条件・結果表示）
行 6〜7: 会社名・装置選択・備考表示
行 8: ファイル表示設定・コード検索条件
行 9〜2000: コード検索結果一覧（RyakuTB）
```

| ✓ | No | セル | 名前付き範囲 | 種別 | 実態（値/数式/VBA代入） | 業務的意味 |
| --- | --- | --- | --- | --- | --- | --- |
|  | 1 | `$A$1` | `Debugf` | 設定値 | デバッグフラグ（9でイベント無効） | デバッグモード制御 |
|  | 2 | `$B$1` | — | 設定値 | "Ex190924" | バージョン識別子 |
|  | 3 | `$N$2` | `Kyou` | 数式 | `=TODAY()` | 当日日付（登録日に使用） |
| ✓ | 4 | `$B$5` | `Uno` | VBA代入 | `受付番号読込表示()` で書込み | 受付番号（入力＋DB取得） |
| ✓ | 5 | `$C$5` | `KaiCd` | VBA代入 | `受付番号読込表示()` / ユーザー入力 | 会社コード |
| ✓ | 6 | `$D$5` | `SeiCd` | VBA代入 | `受付番号読込表示()` / ユーザー入力 | 製品コード |
| ✓ | 7 | `$E$5` | `Tannka` | VBA代入 | `製品読込表示()` で書込み | 単価 |
| ✓ | 8 | `$F$5` | `Tanni` | VBA代入 | `製品読込表示()` で書込み | 単位（箱/PL/m/kg等） |
| ✓ | 9 | `$G$5` | `TourokuBi` | VBA代入 | `製品読込表示()` で書込み | 価格登録日 |
| ✓ | 10 | `$H$5` | `Dose` | VBA代入 | `製品読込表示()` で書込み | 線量（指定線量） |
| ✓ | 11 | `$I$5` | `Souti` | VBA代入 | `製品読込表示()` で書込み | 装置 |
| ✓ | 12 | `$J$5` | `Pass` | VBA代入 | `製品読込表示()` で書込み | パス数 |
|  | 13 | `$K$5` | `Haba` | VBA代入 | `製品読込表示()` で書込み | 幅(cm) |
|  | 14 | `$L$5` | `Takasa` | VBA代入 | `製品読込表示()` で書込み | 高さ(cm) |
|  | 15 | `$M$5` | `Nagasa` | VBA代入 | `製品読込表示()` で書込み | 長さ(cm) |
|  | 16 | `$N$5` | `JyuuRyou` | VBA代入 | `製品読込表示()` で書込み | 重量 |
|  | 17 | `$O$5` | `Irisuu` | VBA代入 | `製品読込表示()` で書込み | 入り数 |
| ✓ | 18 | `$C$6` | `KaiName` | VBA代入 | `価格ファイル表示()` で書込み | 会社名 |
|  | 19 | `$H$6` | `SeiName` | VBA代入 | `製品読込表示()` で書込み | 製品名 |
| ✓ | 20 | `$B$7` | `Syouso` | VBA代入 | `Start_GO()` で "1,2,3号機" 設定 | 装置区分（EB/1,2,3号機） |
| ✓ | 21 | `$C$7` | `KaiBikou` | VBA代入 | `価格ファイル表示()` で書込み | 会社備考 |
|  | 22 | `$H$7` | `SeiBikou` | VBA代入 | `製品読込表示()` で書込み | 製品備考（計上方法） |
| ✓ | 23 | `$C$8` | `Hyouji` | 手動設定 | "しない" / 空欄 | ファイル自動表示ON/OFF |
|  | 24 | `$K$8` | — | 設定値 | "コード" | コード検索ラベル |
|  | 25 | `$L$8` | `Ryakusyou` | 手動設定 | ユーザー入力（カタカナ略称） | 会社略称検索文字 |
|  | 26 | `$M$8` | `MojiLike` | 手動設定 | "先頭" / "含む" | 検索方法（前方一致/部分一致） |
| ✓ | 27 | `$C$3` | `SyainName` | VBA代入 | `Start_GO()` で社員名設定 | ログイン社員名 |
| ✓ | 28 | `$H$3` | `FolDer` | VBA代入 | `価格ファイル表示()` で書込み | 価格表フォルダー名 |
| ✓ | 29 | `$J$3` | `FailName` | VBA代入 | `価格ファイル表示()` で書込み | 価格表ファイル名 |
| ✓ | 30 | `$E$5:$O$5` | `SeihinD` | VBA代入 | 製品データ領域（一括消去用） | 製品情報表示範囲 |
| ✓ | 31 | `$B$6:$O$7` | `TokuD` | VBA代入 | 得意先データ領域 | 得意先情報表示範囲 |
| ✓ | 32 | `$K$9:$O$2000` | `RyakuTB` | DB読込 | `Disp_Sheet()` でDB結果を書込み | 会社コード検索結果一覧 |
|  | 33 | `$T$5` | `FaileDispName` | VBA代入 | 表示中Objectシェイプ名 | 価格表表示Object管理用 |
|  | 34 | `$T$7` | `PathFolder` | VBA代入 | `ReadPath()` で設定ファイルから読込 | 価格表基本パス |
|  | 35 | `$T$14` | — | 数式 | `=TODAY()-H8` | 検索開始入荷日（日付計算） |
|  | 36 | `$T$15` | `NyukabiBef` | 数式 | `=(YEAR(T14)&RIGHT("00"&MONTH(T14),2)&RIGHT("00"&DAY(T14),2))` | 検索開始入荷日（文字列YYYYMMDD形式） |
|  | 37 | `$U$5:$U$14` | `FolderName` | 設定値 | あいうえお〜その他 | フォルダー名リスト（データ検証用） |
|  | 38 | `$W$5:$W$11` | `TanniList` | 設定値 | 箱/PL/m/kg/IB/サイクル 等 | 単位リスト（データ検証用） |
|  | 39 | `$X$5:$X$6` | `Ryaku` | 設定値 | 先頭/含む | 検索方法リスト |

---

## 3. 名前付き範囲一覧


> ✓ = VBA から `Range()` で代入または参照され、業務ロジックに直結する名前付き範囲
> 全件を列挙し、✓ 基準に該当するものにマークを付ける。

| ✓ | No | 名前 | 参照先 | 業務的意味 |
| --- | --- | --- | --- | --- |
|  | 1 | `Debugf` | Main!$A$1 | デバッグモードフラグ（9=イベント無効） |
| ✓ | 2 | `Uno` | Main!$B$5 | 受付番号（検索キー） |
| ✓ | 3 | `KaiCd` | Main!$C$5 | 会社コード（検索キー） |
| ✓ | 4 | `SeiCd` | Main!$D$5 | 製品コード（検索キー） |
| ✓ | 5 | `Tannka` | Main!$E$5 | 単価（表示・登録対象） |
| ✓ | 6 | `Tanni` | Main!$F$5 | 単位（登録対象） |
| ✓ | 7 | `TourokuBi` | Main!$G$5 | 価格登録日 |
| ✓ | 8 | `Dose` | Main!$H$5 | 指定線量 |
| ✓ | 9 | `Souti` | Main!$I$5 | 装置 |
| ✓ | 10 | `Pass` | Main!$J$5 | パス数 |
|  | 11 | `Haba` | Main!$K$5 | 幅(cm) |
|  | 12 | `Takasa` | Main!$L$5 | 高さ(cm) |
|  | 13 | `Nagasa` | Main!$M$5 | 長さ(cm) |
|  | 14 | `JyuuRyou` | Main!$N$5 | 重量 |
|  | 15 | `Irisuu` | Main!$O$5 | 入り数 |
| ✓ | 16 | `KaiName` | Main!$C$6 | 会社名 |
| ✓ | 17 | `SeiName` | Main!$H$6 | 製品名 |
| ✓ | 18 | `Syouso` | Main!$B$7 | 装置区分（EB/1,2,3号機）、在庫検索条件 |
| ✓ | 19 | `KaiBikou` | Main!$C$7 | 会社備考（登録対象） |
| ✓ | 20 | `SeiBikou` | Main!$H$7 | 製品備考＝計上方法（登録対象） |
| ✓ | 21 | `Hyouji` | Main!$C$8 | ファイル自動表示制御（"しない"=非表示） |
|  | 22 | `Ryakusyou` | Main!$L$8 | 会社略称検索入力文字 |
|  | 23 | `MojiLike` | Main!$M$8 | 検索方法（"先頭"/"含む"） |
| ✓ | 24 | `SyainName` | Main!$C$3 | ログイン社員名（登録者として使用） |
| ✓ | 25 | `FolDer` | Main!$H$3 | 価格表フォルダー名 |
| ✓ | 26 | `FailName` | Main!$J$3 | 価格表ファイル名 |
| ✓ | 27 | `Kyou` | Main!$N$2 | 当日日付（=TODAY()、登録日に使用） |
| ✓ | 28 | `SeihinD` | Main!$E$5:$O$5 | 製品データ領域（一括消去用） |
| ✓ | 29 | `TokuD` | Main!$B$6:$O$7 | 得意先データ領域 |
| ✓ | 30 | `RyakuTB` | Main!$K$9:$O$2000 | 会社コード検索結果一覧表示エリア |
|  | 31 | `FaileDispName` | Main!$T$5 | 表示中の価格表Objectシェイプ名 |
|  | 32 | `PathFolder` | Main!$T$7 | 価格表記録先フォルダーパス |
|  | 33 | `NyukabiBef` | Main!$T$15 | 検索開始入荷日（YYYYMMDD文字列） |
|  | 34 | `FolderName` | Main!$U$5:$U$14 | フォルダー名リスト（入力規則用） |
|  | 35 | `TanniList` | Main!$W$5:$W$11 | 単位リスト（入力規則用） |
|  | 36 | `Ryaku` | Main!$X$5:$X$6 | 検索方法リスト（先頭/含む） |


### 3.1 データの入力規則

| シート | セル | 種別 | 制約 | 用途 |
| --- | --- | --- | --- | --- |
| Main | `B5` | テキスト長 | 10〜11 |  |
| Main | `J3` | カスタム | <=LENB(J3)<=50 |  |
| Main | `C8` | リスト | リスト: `"する,しない"` |  |
| Main | `B7` | リスト | リスト: `$T$9:$T$10` |  |
| Main | `M8` | リスト | リスト: `Ryaku` |  |
| Main | `F5` | リスト | リスト: `TanniList` |  |
| Main | `H3` | リスト | リスト: `FolderName` |  |


---

---

## 4. 数式一覧

| シート | 数式件数 | 備考 |
| --- | --- | --- |
| Main | 3 | 日付計算・検索条件生成 |

### 4.1 Main

| セル | 数式 | 説明 |
| --- | --- | --- |
| `$N$2` | `=TODAY()` | 当日日付（`Kyou`）。登録日として使用 |
| `$T$14` | `=TODAY()-H8` | 検索開始入荷日の計算。H8は未使用（値なし）のため実質TODAY() |
| `$T$15` | `=(YEAR(T14)&RIGHT("00"&MONTH(T14),2)&RIGHT("00"&DAY(T14),2))` | T14をYYYYMMDD文字列に変換。`NyukabiBef` として在庫検索WHERE句に使用 |

---

## 5. ボタン・マクロ対応

> ✓ = DB 更新・画面遷移・計算実行など副作用のある操作を起動するボタン

### 5.1 シート上のボタン（Form Control）

| ✓ | No | シート | ボタンラベル | 割り当てマクロ | 動作概要 |
| --- | --- | --- | --- | --- | --- |
|  | 1 | Main | 終了 | `Bookを閉じる()` | ブックを閉じる（上書き保存なし） |
| ✓ | 2 | Main | コード検索 | `会社コード表示()` | 略称でDB検索し会社コード一覧をシートに表示 |
| ✓ | 3 | Main | ファイル登録 | `単価_フォルダーファイル名更新()` | `ExSeihinJ` にフォルダー名・ファイル名をINSERT/UPDATE |
| ✓ | 4 | Main | 単価登録 | `単価登録()` | `ExSeihinZ` に単価・単位・備考をINSERT/UPDATE |
| ✓ | 5 | Main | 表示 | `価格表の表示()` | 登録済み価格表ファイルを外部アプリで開く |

### 5.2 ショートカットキー

| No | マクロ名 | ショートカット | 処理概要 |
| --- | --- | --- | --- |
| 1 | `AllCls()` | **Ctrl+E** | 全データクリア（一覧・検索条件を初期化） |
| 2 | `画面初期化()` | **Ctrl+E** | 画面初期化（DB再読込を伴うクリア） |

### 5.3 ユーザーフォーム上のボタン

なし。

### 5.4 CommandBar に動的追加されるボタン

なし。`画面設定()` で既存CommandBarの表示/非表示を制御するのみ。

---

## 6. VBA モジュール仕様

### 6.0 全プロシージャ一覧


> ✓ = ユーザー操作の起点（Click イベント等） / DB I/O を実行 / 他モジュールから呼び出される Public
> 全プロシージャを列挙し、✓ 基準に該当するものにマークを付ける。

| ✓ | No | モジュール | プロシージャ | スコープ | 種別 | 概要 |
| --- | --- | --- | --- | --- | --- | --- |
| ✓ | 1 | **Sheet1**（「Main」） | `Worksheet_BeforeDoubleClick()` | Private | Event | K列9行以降のダブルクリックで会社コードを `KaiCd` に代入 |
| ✓ | 2 | **Sheet1**（「Main」） | `Worksheet_Change()` | Private | Event | B5/C5/D5/E5セル変更時に対応する表示処理を呼び出し |
| ✓ | 3 | **ThisWorkbook.cls** | `Workbook_Open()` | Private | Event | `Start_GO()` を呼び出し（起動処理） |
| ✓ | 4 | **ThisWorkbook.cls** | `Workbook_BeforeClose()` | Private | Event | `画面解除()` 呼び出し後、保存せず終了 |
|  | 5 | **ThisWorkbook.cls** | `Workbook_Activate()` | Private | Event | ウィンドウ最大化 |
| ✓ | 6 | **SQL_Execution.bas** | `Open_oraconDB()` | Public | Sub | ADO+ODBC でDB接続（DSN=ricdb） |
| ✓ | 7 | **SQL_Execution.bas** | `SQL_Exe()` | Public | Sub | SQL文を Execute で実行しレコードセット取得 |
| ✓ | 8 | **SQL_Execution.bas** | `SQL_INSERT_UPDATE()` | Public | Sub | 既存チェック→INSERT/UPDATE 汎用処理 |
| ✓ | 9 | **SQL_Execution.bas** | `SQL_Delete()` | Public | Sub | DELETE 汎用処理 |
| ✓ | 10 | **SQL_Execution.bas** | `Disp_Sheet()` | Public | Sub | SQL結果をシートに直接書込み（CopyFromRecordset） |
| ✓ | 11 | **SQL_Execution.bas** | `Set_Array()` | Public | Sub | SQL結果を配列に格納 |
| ✓ | 12 | **データベースR.bas** | `受付番号読込表示()` | Public | Sub | `zaiko_V` から受付番号で会社コード・製品コード取得 |
| ✓ | 13 | **データベースR.bas** | `価格ファイル表示()` | Public | Sub | `ExSeihinJ`+`tokumst` から会社名・フォルダー・ファイル名取得 |
| ✓ | 14 | **データベースR.bas** | `製品読込表示()` | Public | Sub | `ExSeihinZ`+`sehmst` から単価・製品情報取得 |
| ✓ | 15 | **表示.bas** | `価格表の表示()` | Public | Sub | WScript.Shell で価格表ファイルを外部アプリ起動 |
| ✓ | 16 | **表示.bas** | `受付番号表示()` | Public | Sub | 受付番号変更時の全体制御（消去→DB読込→表示） |
| ✓ | 17 | **表示.bas** | `会社データ表示()` | Public | Sub | 会社コード変更時の全体制御（消去→DB読込→表示） |
| ✓ | 18 | **表示.bas** | `製品データ表示()` | Public | Sub | 製品コード変更時のDB読込・表示 |
| ✓ | 19 | **開始処理.bas** | `Start_GO()` | Public | Sub | 起動時処理（シート保護・認証・初期化） |
| ✓ | 20 | **開始処理.bas** | `ReadPath()` | Public | Sub | `C:\ExSys実行\価格表Path.txt` から基本パス読込 |
| ✓ | 21 | **単価_ファイル登録処理.bas** | `単価_フォルダーファイル名更新()` | Public | Sub | `ExSeihinJ` にフォルダー名・ファイル名をINSERT/UPDATE |
| ✓ | 22 | **単価_ファイル登録処理.bas** | `単価登録()` | Public | Sub | `ExSeihinZ` に単価・単位・備考をINSERT/UPDATE |
| ✓ | 23 | **会社コード検索.bas** | `会社コード表示()` | Public | Sub | `tokumst` から略称検索し結果をシートに表示 |
| ✓ | 24 | **画面クリア.bas** | `AllCls()` | Public | Sub | 全画面消去（社員名含む） |
| ✓ | 25 | **画面クリア.bas** | `画面消去()` | Public | Sub | 入力/表示セルの初期化（起動時・操作時に呼出し） |
|  | 26 | **画面クリア.bas** | `表示Object消去()` | Public | Sub | 価格表表示Objectの削除 |
|  | 27 | **画面クリア.bas** | `画面初期化()` | Public | Sub | Debug用（受付番号クリア） |
|  | 28 | **画面クリア.bas** | `aaa()` | Public | Sub | Debug用（セル移動） |
|  | 29 | **画面設定解除.bas** | `画面設定()` | Public | Sub | イベント有効化・CommandBar制御・シート見出し非表示 |
|  | 30 | **画面設定解除.bas** | `画面解除()` | Public | Sub | フルスクリーン解除・CommandBar復元・シート見出し表示 |
|  | 31 | **ユーティリティ.bas** | `イベント有効()` | Public | Sub | Application.EnableEvents = True |
|  | 32 | **ユーティリティ.bas** | `セル書式設定()` | Public | Sub | C5セルのフォント・罫線・背景色設定 |
|  | 33 | **終了処理.bas** | `Bookを閉じる()` | Public | Sub | DisplayAlerts=False でブック閉じる/アプリ終了 |

---

### 6.1 ThisWorkbook.cls

ブックの起動・終了イベントを受け、後続処理へ委譲する。

| No | プロシージャ | 処理フロー | 主要SQL |
| --- | --- | --- | --- |
| 1 | `Workbook_Open()` | `Start_GO()` を呼び出し（起動処理） | なし |
| 2 | `Workbook_BeforeClose()` | `画面解除()` 呼び出し後、保存せず終了 | なし |
| 3 | `Workbook_Activate()` | ウィンドウ最大化 | なし |

## 7. ユーザーフォーム仕様

なし。本ファイルにユーザーフォームは含まれない。

---

## 8. DB 接続・外部連携

### 8.1 ODBC 接続設定

| DSN 名 | UID | PWD | 用途 |
| --- | --- | --- | --- |
| `ricdb` | `ric` | `t6101` | 照射管理システムDB接続 |

> **DB サーバー IP**: DSN（ricdb）経由で接続。IPアドレスはDSN設定に依存。

### 8.2 テーブル一覧（参照/更新区分付き）

> ✓ = INSERT / UPDATE / DELETE の対象テーブル（参照のみのテーブルは ✓ なし）

| ✓ | No | テーブル名 | 区分 | 主な用途 | キー列 | 参照/更新列 |
| --- | --- | --- | --- | --- | --- | --- |
| ✓ | 1 | `ExSeihinJ` | **参照＋更新** | 会社別価格表ファイル情報管理 | `kaisyacd` | 参照: `coname`(tokumst結合), `folder`, `filename`, `kaibikou`, `toudate`, `touname` / 更新: `kaisyacd`, `folder`, `filename`, `kaibikou`, `toudate`, `touname`（`単価_フォルダーファイル名更新()` からINSERT/UPDATE） |
| ✓ | 2 | `ExSeihinZ` | **参照＋更新** | 会社別・製品別単価情報管理 | `kaisyacd` + `sehncd` | 参照: `tanka`, `tani`, `toudate`, `keijyou` / 更新: `kaisyacd`, `sehncd`, `tanka`, `tani`, `keijyou`, `toudate`, `touname`（`単価登録()` からINSERT/UPDATE） |
|  | 3 | `zaiko_V` | 参照 | 在庫ビューから受付番号→会社・製品コード取得 | `uno` | 参照: `Uno`, `kaisyacd`, `sehncd`, `kainame` |
|  | 4 | `tokumst` | 参照 | 得意先マスタ（会社名・略称検索） | `kaisyacd` | 参照: `coname`, `kairname` |
|  | 5 | `sehmst` | 参照 | 製品マスタ（製品名・寸法・線量等） | `kaisyacd` + `sehncd` | 参照: `siteisn`, `syouso`, `pass`, `haba`, `takasa`, `nagasa`, `jyury`, `incnt`, `seiname` |
|  | 6 | `shainmst` | 参照 | 社員マスタ（認証・社員名取得） | `shano` | 参照: `shaname`, `kshika` |

> **「キー列」の定義**: JOIN 条件または UPDATE/DELETE の WHERE 句で使用される列を示す。

### 8.3 SQL 一覧

#### 8.3.1 受付番号による在庫検索（`受付番号読込表示()` / **データベースR.bas**）

```sql
SELECT Uno, kaisyacd, sehncd, kainame
FROM zaiko_V
WHERE uno LIKE '%<受付番号>' AND syouso<>'4' AND nyukabi>'<NyukabiBef>'
```

#### 8.3.2 会社情報・価格ファイル取得（`価格ファイル表示()` / **データベースR.bas**）

```sql
SELECT t.coname, e.folder, e.filename, e.kaibikou, e.toudate, e.touname
FROM ExSeihinJ e, tokumst t
WHERE e.kaisyacd(+) = t.kaisyacd AND t.kaisyacd = '<会社コード>'
```

#### 8.3.3 製品情報取得（`製品読込表示()` / **データベースR.bas**）

```sql
SELECT e.tanka, e.tani, e.toudate, s.siteisn, s.syouso, s.pass, s.haba,
       s.takasa, s.nagasa, s.jyury, s.incnt, s.seiname, e.keijyou
FROM ExSeihinZ e, sehmst s
WHERE s.kaisyacd = e.kaisyacd(+) AND s.sehncd = e.sehncd(+)
  AND s.kaisyacd = '<会社コード>' AND s.sehncd = '<製品コード>'
```

#### 8.3.4 会社コード略称検索（`会社コード表示()` / **会社コード検索.bas**）

```sql
SELECT kaisyacd, RTRIM(TRANSLATE(coname, '　', ' '))
FROM tokumst
WHERE kairname LIKE '<略称>%' AND kaisyacd < '2000'
ORDER BY kairname
```

#### 8.3.5 社員認証（`Start_GO()` / **開始処理.bas**）

```sql
SELECT TRIM(REPLACE(shaname, '　', ' '))
FROM shainmst
WHERE shano = '<社員番号>' AND kshika = '1'
```

#### 8.3.6 単価登録（`単価登録()` / **単価_ファイル登録処理.bas**）

```sql
SELECT COUNT(*) FROM ExSeihinZ WHERE kaisyacd='<会社>' AND sehncd='<製品>'
INSERT INTO ExSeihinZ (kaisyacd, sehncd, tanka, tani, keijyou, toudate, touname)
VALUES ('<会社>', '<製品>', <単価>, '<単位>', '<備考>', '<日付>', '<社員名>')
UPDATE ExSeihinZ SET tanka=<単価>, tani='<単位>', keijyou='<備考>',
       toudate='<日付>', touname='<社員名>'
WHERE kaisyacd='<会社>' AND sehncd='<製品>'
```

#### 8.3.7 ファイル登録（`単価_フォルダーファイル名更新()` / **単価_ファイル登録処理.bas**）

```sql
-- INSERT/UPDATE ExSeihinJ
-- キー: kaisyacd
-- 列: kaisyacd, folder, filename, kaibikou, toudate, touname
```

### 8.4 外部ファイル連携

| ファイル | パス | ファイル名 | 処理 | 備考 |
| --- | --- | --- | --- | --- |
| 価格表パス設定 | `C:\ExSys実行\` | `価格表Path.txt` | `Open ... For Input` | 起動時にフォルダーパスを読込 |
| 価格表ファイル | `\\RNTSVR-FS\Sv_cup\営業部照射課\価格表\<フォルダ>\` | DB登録ファイル名 | WScript.Shell で表示 | 外部アプリで開く |

#### 価格表Path.txt 詳細

- **出力元プロシージャ**: `ReadPath()` / **開始処理.bas**
- **パス生成ロジック**（VBAソースより）:
  ```
  Open "C:\ExSys実行\価格表Path.txt" For Input As #myFNo
  Input #myFNo, myPathName
  Range("PathFolder") = myPathName
  ```
- **出力例**: `C:\ExSys実行\価格表Path.txt`（内容はフォルダーパス文字列）
- **読込内容**: 価格表記録先フォルダーの基本パス（`PathFolder` に格納）
- **ファイル不在時**: エラーメッセージを表示し、ユーザーに「価格表Path作成」の実行を案内

#### 価格表ファイル 詳細

- **出力元プロシージャ**: `価格表の表示()` / **表示.bas**
- **パス生成ロジック**（VBAソースより）:
  ```
  FilePathName = "\\RNTSVR-FS\Sv_cup\営業部照射課\価格表\" & Range("FolDer") & "\" & Range("FailName")
  ```
- **出力例**: `\\RNTSVR-FS\Sv_cup\営業部照射課\価格表\あいうえお\見積書.xlsx`
- **読込内容**: 価格表ファイル（Word/Excel等）を外部アプリケーションで開いて表示
- **備考**: `FolDer` と `FailName` は `ExSeihinJ` テーブルから取得。ファイル形式はDB登録値に依存

---

## 9. データフロー

各フローは「起点 → 処理 → 結果」の粒度で記述する。

### 9.1 起動フロー

| No | 起点 | 処理 | 結果 |
| --- | --- | --- | --- |
| 1 | 📄 ブックOpen | 📄 `Workbook_Open()` / **ThisWorkbook.cls** → `Start_GO()` | 起動処理開始 |
| 2 | 上記 | 📊 シート保護設定・ウィンドウ最大化 | UI初期化 |
| 3 | 上記 | 📄 `画面消去()` / **画面クリア.bas** | 全セル初期化 |
| 4 | 上記 | 📄 `ReadPath()` / **開始処理.bas** | `C:\ExSys実行\価格表Path.txt` 読込 → `PathFolder` 設定 |
| 5 | 上記 | 🖥️ InputBox「社員番号入力」 | ユーザー認証開始 |
| 6 | 上記 | 🗄️ `Set_Array()` / **SQL_Execution.bas** | `shainmst` から社員名取得（`kshika='1'` 条件） |
| 7 | 上記 | 📊 `SyainName` に社員名書込み | 認証完了・操作可能状態 |

#### ツリー図（補助）

```
（ブックOpen）
└─ 📄 Workbook_Open              [ThisWorkbook.cls]
   └─ 📄 Start_GO                [開始処理.bas]
      ├─ 📊 シート保護設定
      ├─ 📄 画面消去              [画面クリア.bas]
      ├─ 📄 ReadPath              [開始処理.bas]（価格表Path.txt読込）
      │   └─ [ファイル不在] → エラー表示 → 終了
      ├─ 🖥️ InputBox（社員番号入力）
      ├─ 🗄️ Set_Array             [SQL_Execution.bas]（shainmst SELECT）
      │   └─ [認証失敗] → 再入力 or 終了
      ├─ 📊 画面設定              [画面設定解除.bas]
      └─ 📊 SyainName 書込み → KaiCd.Select
```

### 9.2 受付番号検索フロー

| No | 起点 | 処理 | 結果 |
| --- | --- | --- | --- |
| 1 | 📊 `$B$5`（Uno）セル値変更 | 📄 `Worksheet_Change()` / **Sheet1.cls** | イベント検知 |
| 2 | 上記 | 📄 `受付番号表示()` / **表示.bas** | 消去→DB検索開始 |
| 3 | 上記 | 🗄️ `受付番号読込表示()` / **データベースR.bas** | `zaiko_V` から会社・製品コード取得 |
| 4 | 上記 | 📊 `Uno`, `KaiCd`, `SeiCd` に書込み | シートに表示 |
| 5 | 上記 | 🗄️ `価格ファイル表示()` / **データベースR.bas** | `ExSeihinJ`+`tokumst` から会社名・ファイル情報取得 |
| 6 | 上記 | 📊 `KaiName`, `FolDer`, `FailName`, `KaiBikou` 書込み | 会社情報表示 |
| 7 | 上記 | 🗄️ `製品読込表示()` / **データベースR.bas** | `ExSeihinZ`+`sehmst` から単価・製品情報取得 |
| 8 | 上記 | 📊 `Tannka`, `Tanni`, `Dose`, `Souti` 等に書込み | 製品情報表示 |
| 9 | [条件] `Hyouji`≠"しない" | 📄 `価格表の表示()` / **表示.bas** | 価格表ファイル外部表示 |

#### ツリー図（補助）

```
（$B$5 セル値変更）
└─ 📄 Worksheet_Change           [Sheet1.cls]
   └─ 📄 受付番号表示            [表示.bas]
      ├─ 📊 関連セル消去
      ├─ 🗄️ 受付番号読込表示     [データベースR.bas]（zaiko_V SELECT）
      │   └─ [0件] → エラー / [複数件] → 一覧表示
      ├─ 📊 Uno, KaiCd, SeiCd 書込み
      ├─ 🗄️ 価格ファイル表示     [データベースR.bas]（ExSeihinJ+tokumst SELECT）
      ├─ 📊 KaiName, FolDer, FailName 書込み
      ├─ 🗄️ 製品読込表示         [データベースR.bas]（ExSeihinZ+sehmst SELECT）
      ├─ 📊 Tannka, Tanni, Dose 等書込み
      └─ [Hyouji≠"しない"]
          └─ 📄 価格表の表示     [表示.bas]（WScript.Shell で外部ファイル表示）
```

### 9.3 単価登録フロー

| No | 起点 | 処理 | 結果 |
| --- | --- | --- | --- |
| 1 | 📊（単価登録ボタン Click） | 📄 `単価登録()` / **単価_ファイル登録処理.bas** | 登録処理開始 |
| 2 | 上記 | 📄 入力チェック（KaiCd, SeiCd 必須） | 未入力ならエラー表示して終了 |
| 3 | 上記 | 📊 `Tannka`, `Tanni`, `SeiBikou`, `Kyou`, `SyainName` 読込 | 登録データ準備 |
| 4 | 上記 | 🗄️ `SQL_INSERT_UPDATE()` / **SQL_Execution.bas** | `ExSeihinZ` に INSERT or UPDATE |
| 5 | 上記 | 🖥️ MsgBox「単価と備考を登録しました。」 | 完了通知 |

#### ツリー図（補助）

```
（単価登録ボタン Click）
└─ 📄 単価登録                   [単価_ファイル登録処理.bas]
   ├─ 📄 入力チェック（KaiCd, SeiCd 必須）
   │   └─ [未入力] → 🖥️ MsgBox エラー → 終了
   ├─ 📊 シートから値読込（Tannka, Tanni, SeiBikou, Kyou, SyainName）
   ├─ 🗄️ SQL_INSERT_UPDATE       [SQL_Execution.bas]
   │   ├─ 🗄️ SELECT COUNT(*) → 既存チェック
   │   ├─ [0件] 🗄️ INSERT INTO ExSeihinZ
   │   └─ [1件以上] 🗄️ UPDATE ExSeihinZ
   └─ 🖥️ MsgBox 完了通知
```

### 9.4 ファイル登録フロー

| No | 起点 | 処理 | 結果 |
| --- | --- | --- | --- |
| 1 | 📊（ファイル登録ボタン Click） | 📄 `単価_フォルダーファイル名更新()` / **単価_ファイル登録処理.bas** | 登録処理開始 |
| 2 | 上記 | 📄 入力チェック（KaiCd, FolDer, FailName 必須） | 未入力ならエラー表示して終了 |
| 3 | 上記 | 📄 拡張子チェック・ファイル名適正チェック | 不正ならエラー表示して終了 |
| 4 | 上記 | 🗄️ `SQL_INSERT_UPDATE()` テスト登録（kaisyacd='0000'） | ファイル名のDB登録可否検証 |
| 5 | 上記 | 📄 Dir() でファイル存在チェック | 存在しない場合は上書き確認 |
| 6 | 上記 | 🗄️ `SQL_INSERT_UPDATE()` / **SQL_Execution.bas** | `ExSeihinJ` に INSERT or UPDATE |
| 7 | 上記 | 🖥️ MsgBox「フォルダ名とファイル名を登録しました。」 | 完了通知 |

#### ツリー図（補助）

```
（ファイル登録ボタン Click）
└─ 📄 単価_フォルダーファイル名更新  [単価_ファイル登録処理.bas]
   ├─ 📄 入力チェック（KaiCd, FolDer, FailName 必須）
   │   └─ [未入力] → 🖥️ MsgBox エラー → 終了
   ├─ 📄 拡張子チェック
   ├─ 🗄️ SQL_INSERT_UPDATE（テスト: kaisyacd='0000'）
   │   └─ [エラー] → 🖥️ MsgBox → 終了
   ├─ 🗄️ SELECT filename（適正チェック）
   │   └─ [文字化け] → 🖥️ MsgBox → 終了
   ├─ 📄 Dir() ファイル存在チェック
   │   └─ [不在] → 🖥️ MsgBox 上書き確認
   ├─ 📊 シートから値読込（KaiCd, FolDer, FailName, KaiBikou, Kyou, SyainName）
   ├─ 🗄️ SQL_INSERT_UPDATE       [SQL_Execution.bas]
   │   ├─ [0件] 🗄️ INSERT INTO ExSeihinJ
   │   └─ [1件以上] 🗄️ UPDATE ExSeihinJ
   └─ 🖥️ MsgBox 完了通知
```

### 9.5 会社コード検索フロー

| No | 起点 | 処理 | 結果 |
| --- | --- | --- | --- |
| 1 | 📊（コード検索ボタン Click） | 📄 `会社コード表示()` / **会社コード検索.bas** | 検索処理開始 |
| 2 | 上記 | 📊 `Ryakusyou` セル読込 | 検索文字取得 |
| 3 | 上記 | 📄 検索条件組立（MojiLike="含む"→部分一致 / "先頭"→前方一致） | SQL WHERE句生成 |
| 4 | 上記 | 🗄️ `Disp_Sheet()` / **SQL_Execution.bas** | `tokumst` から検索→結果をシートK9〜O列に書込み |
| 5 | 上記 | 📊 `RyakuTB` 領域に会社コード・会社名一覧表示 | ユーザーがダブルクリックで選択可能 |

#### ツリー図（補助）

```
（ｺｰﾄﾞ検索ボタン Click）
└─ 📄 会社コード表示             [会社コード検索.bas]
   ├─ 📊 Ryakusyou 読込（検索文字）
   ├─ 📄 WHERE句生成（MojiLike による分岐）
   ├─ 📊 RyakuTB 消去
   ├─ 🗄️ Disp_Sheet             [SQL_Execution.bas]（tokumst SELECT）
   └─ 📊 K9〜O列に結果書込み
       └─ 📊 Worksheet_BeforeDoubleClick で KaiCd に代入可能
```

---

## 10. セキュリティ注意事項


| No | カテゴリ | 内容 | リスク |
| --- | --- | --- | --- |
| 1 | 認証情報ハードコード | DSN=`ricdb`, UID=`ric`, PWD=`t6101` が VBA ソースに平文記載 | 中：VBAエディタで閲覧可能 |
| 2 | ファイル操作 | `価格表Path.txt` を読込 | 中：パス改ざん時の読込先変更 |
| 3 | 外部プログラム実行 | `Wscript.Shell.Run` で価格表ファイルを表示 | 中：外部プログラムの起動 |

---

## スコープ外（本仕様書に含まないもの）

- セル書式（色・罫線・フォント）
- 条件付き書式、グラフ・画像、印刷設定

必要な場合は Excel 画面のスクリーンショットで補完してください。
