# Exカレンダー 仕様書

> **ファイル種別**: .xlsm（マクロ付き）
> **用途**: 指定月から12か月分のカレンダーを表示し、会社休祭日をDBテーブル `ExYasumiX` と連携して管理する（EXメニューの1ファイルとして照射管理システムを補完）
> **VBA プロジェクト**: モジュール 7 本（.bas 5 / .cls 2 / .frm 0）
> **外部連携**: DSN=ricdb（Oracle DB）
> **解析日**: 2026-08-20（表記統一）

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
| ✓ | 1 | カレンダー | — | — | visible | — |


### 1.2 ユーザーフォーム一覧

> ✓ = ユーザー入力を受け付ける、または業務フローの起点となるフォーム

なし。

### 1.3 VBA モジュール一覧

> ✓ = ユーザー操作の起点 / DB I/O を含む / 他モジュールから呼び出される / コード行数上位 25%

| ✓ | No | モジュール | 種別 | プロシージャ数 | 主な役割 |
| --- | --- | --- | --- | --- | --- |
| ✓ | 1 | **ThisWorkbook.cls** | .cls | 2 | 起動/終了イベント |
| ✓ | 2 | **Sheet1.cls** | .cls | 2 | 「カレンダー」シートイベント |
| ✓ | 3 | **Ex休み読込.bas** | .bas | — | DBから休日データを読み込みシートに貼付け |
| ✓ | 4 | **Ex休み更新.bas** | .bas | — | シートの休日をDBに書き戻し |
| ✓ | 5 | **SQL_Execution.bas** | .bas | — | ADO/ODBC DB接続・SQL実行基盤 |
|  | 6 | **終了処理.bas** | .bas | — | ブック終了処理 |
|  | 7 | **初期化.bas** | .bas | — | クリア処理 |

---

## 2. シート詳細

### 2.0 シート可視性一覧

| No | シート | VBA による非表示化 | 表示するタイミング | 非表示にするタイミング | 制御プロシージャ |
| --- | --- | --- | --- | --- | --- |
| 1 | カレンダー | — | — | — | — |

> 以下の各シートのレイアウト構造表における ✓ = VBA から `Range()` で代入または参照され、業務ロジックに直結するセル

### 2.1 カレンダー

**目的**: `HyoujiBi`（G1）に表示月を入力すると、その月を起点に12か月分のカレンダーが横方向に生成される。各日付に対しDBから休祭日フラグを読み込み表示する。右クリックで会社独自の休日を手動追加・削除でき、「休日登録」ボタンでDBに反映する。

#### 非表示行・列

なし。

#### カラム構成（3列 × 12か月）

カレンダーは D列から AM列まで、3列1セットで12か月分が横に展開される。

| 月番号 | 日付列（Col1） | 休日フラグ列（Col2） | 更新前値列（Col3） |
| --- | --- | --- | --- |
| 1か月目 | D | E | F |
| 2か月目 | G | H | I |
| 3か月目 | J | K | L |
| 4か月目 | M | N | O |
| 5か月目 | P | Q | R |
| 6か月目 | S | T | U |
| 7か月目 | V | W | X |
| 8か月目 | Y | Z | AA |
| 9か月目 | AB | AC | AD |
| 10か月目 | AE | AF | AG |
| 11か月目 | AH | AI | AJ |
| 12か月目 | AK | AL | AM |

#### 行構成

| 行 | 内容 |
| --- | --- |
| 1 | ヘッダー（D1: ラベル「検索開始日」, G1:H1: `HyoujiBi`入力セル, J1: 操作説明, AL1: システムID「Ex5」） |
| 2 | `Honnjitu`（AQ2: 本日日付） |
| 3 | （空） |
| 4 | 休日フラグテンプレート行（VLOOKUPで `YasumiTB` を参照）。`休日読込` でこの行の数式を行7〜37にコピー・値貼付け |
| 5 | （空） |
| 6 | 各月の月初日（日付データ: D列=`HyoujiBi`、以降 EDATE で+1か月） |
| 7 | 各月の1日（INT変換した月初日） |
| 8〜37 | 各月の2日〜31日（+1日ずつ加算、月をまたいだら空） |
| 6〜37 | AP列以降 `YasumiTB`: DBから読み込んだ休日テーブル（KYUUJITU1, FLG） |


#### 主要セル

| ✓ | No | セル | 名前付き範囲 | 種別 | 実態（値/数式/VBA代入） | 業務的意味 |
| --- | --- | --- | --- | --- | --- | --- |
| ✓ | 1 | `$G$1` | `HyoujiBi` | 入力 | 表示開始月（YYYY/M/1 に正規化） | カレンダー表示の起点月 |
|  | 2 | `$AQ$2` | `Honnjitu` | VBA代入 | 本日日付 | 当日判定・休日更新の基準日 |
|  | 3 | `$A$1` | `Debug` | 手動設定 | 空=通常 / 非空=右クリック処理スキップ | デバッグモード切替 |
| ✓ | 4 | `$AQ$7:$AR$372` | `YasumiTB` | VBA代入 | `休日読込()` が DB から書込み | 休日テーブル（VLOOKUP 範囲） |
| ✓ | 5 | E7:F37 ほか12か月分 | `YasumiDay` | VBA代入 / 入力 | 休日フラグと更新前値 | 変更検出と一括クリア |
| ✓ | 6 | `$D$6` | — | 数式 | `=HyoujiBi` | 1か月目の月初日 |

#### DB読み込み領域

| 列 | 行範囲 | 内容 |
| --- | --- | --- |
| AP | 6（ヘッダー "No"） | 番号 |
| AQ | 6〜372 | `KYUUJITU1`（休日日付: Excelシリアル値） |
| AR | 6〜372 | `FLG`（フラグ文字列。`休` で始まるもの） |

---

## 3. 名前付き範囲一覧

> ✓ = VBA から `Range()` で代入または参照され、業務ロジックに直結する名前付き範囲

| ✓ | No | 名前 | 参照先 | 業務的意味 |
| --- | --- | --- | --- | --- |
|  | 1 | `Debug` | カレンダー!$A$1 | デバッグモード（空=通常動作、非空=右クリック処理をスキップ） |
|  | 2 | `Honnjitu` | カレンダー!$AQ$2 | 本日の日付 |
|  | 3 | `HyoujiBi` | カレンダー!$G$1 | カレンダー表示開始月入力セル（YYYY/M/1 形式に正規化される） |
|  | 4 | `HyujiS` | カレンダー!$D$6 | 検索開始日付（D6 = `HyoujiBi` をコピー: `=IF(HyoujiBi="","",HyoujiBi)`） |
|  | 5 | `YasumiDay` | E7:F37, H7:I37, K7:L37 ... AM7:AM37（12か月分） | 休日フラグ + 更新前値セル群（変更検出と一括クリア用） |
|  | 6 | `YasumiTB` | カレンダー!$AQ$7:$AR$372 | DBから読み込んだ休日テーブル（VLOOKUPのルックアップ範囲） |

---


### 3.1 データの入力規則

なし。

## 4. 数式一覧

### カレンダーシート

すべての数式は3列単位 × 12か月で繰り返されるため、代表パターンを示す。

#### 行6: 月初日計算

| セル（代表） | 数式 | 説明 |
| --- | --- | --- |
| D6 | `=IF(HyoujiBi="","",HyoujiBi)` | 1か月目の月初日（`HyoujiBi` と同値） |
| G6 | `=IF(HyujiS="","",EDATE(D6,1))` | 2か月目の月初日（前月+1か月） |
| J6〜AK6 | `=IF(HyujiS="","",EDATE(<前月>,1))` | 3〜12か月目（同パターン） |

#### 行4: 休日フラグテンプレート（VLOOKUP）

| セル（代表） | 数式 | 説明 |
| --- | --- | --- |
| E4, F4 | `=IF(ISERROR(VLOOKUP(D4,YasumiTB,2,FALSE)),"",LEFT(VLOOKUP(D4,YasumiTB,2,FALSE),1))` | D列の日付で `YasumiTB` を検索し、FLGの先頭1文字（「休」等）を取得 |
| H4, I4〜 | 同パターン（参照列が3列ずつシフト） | — |

この行が `休日読込` で行7〜37に数式→値貼付けされる。

#### 行7: 月初日整数化

| セル（代表） | 数式 | 説明 |
| --- | --- | --- |
| D7 | `=IF(HyujiS="","",INT(HyujiS))` | 月初日をシリアル整数値に変換 |
| G7〜AK7 | `=IF(HyujiS="","",INT(<月6セル>))` | 同パターン |

#### 行8〜37: 日付の連続生成

| セル（代表） | 数式 | 説明 |
| --- | --- | --- |
| D8 | `=IF(D7="","",IF(MONTH(D7)=MONTH(D7+1),D7+1,""))` | 前日+1日。翌日が別月なら空（月末処理） |
| D9〜D37 | 同パターン（D8→D9、D9→D10…） | 最大31日分 |
| G8〜AK37 | 同パターン | 各月列で繰り返し |

---

## 5. ボタン・マクロ対応

> ✓ = DB 更新・画面遷移・計算実行など副作用のある操作を起動するボタン

### 5.1 シート上のボタン（Form Control）

| ✓ | No | シート | ボタンラベル | 割り当てマクロ | 動作概要 |
| --- | --- | --- | --- | --- | --- |
|  | 1 | カレンダー | 休日登録 | `YasumiKousinn()` | シートの休日フラグ変更をDBテーブル `ExYasumiX` に書き込む |
|  | 2 | カレンダー | 終了 | `Bookを閉じる()` | ブックを保存せずに閉じる |

### 5.2 ショートカットキー

| No | マクロ名 | ショートカット | 処理概要 |
| --- | --- | --- | --- |
| 1 | `クリア()` | **Ctrl+E** | カレンダーの休日データ（`HyoujiBi`・`YasumiDay`・`YasumiTB`）をクリアし初期状態に戻す |

### 5.3 ユーザーフォーム上のボタン（サマリ）

なし。

### 5.4 CommandBar に動的追加されるボタン

なし。

## 6. VBA モジュール仕様

### 6.0 全プロシージャ一覧

> ✓ = ユーザー操作の起点（Click イベント等） / DB I/O を実行 / 他モジュールから呼び出される Public

| ✓ | No | モジュール | プロシージャ | スコープ | 種別 | 概要 |
| --- | --- | --- | --- | --- | --- | --- |
| ✓ | 1 | **ThisWorkbook.cls** | `Workbook_Open()` | Private | Event | ウィンドウ最大化・シート保護・休日データクリア・ズーム調整 |
| ✓ | 2 | **ThisWorkbook.cls** | `Workbook_BeforeClose()` | Private | Event | 保存ダイアログ抑止 |
| ✓ | 3 | **Sheet1.cls** | `Worksheet_BeforeRightClick()` | Private | Event | 右クリックで休日フラグ切替（条件付き書式により色表示） |
| ✓ | 4 | **Ex休み読込.bas** | `休日読込()` | Public | Sub | DBテーブル`ExYasumiX`から休日データを読込みシートに反映 |
| ✓ | 5 | **Ex休み更新.bas** | `YasumiKousinn()` | Public | Sub | シートの休日フラグ変更をDBテーブル`ExYasumiX`に書き込む |
| ✓ | 6 | **SQL_Execution.bas** | `Open_oraconDB()` | Public | Sub | Oracle DB接続（ODBC） |
| ✓ | 7 | **SQL_Execution.bas** | `SQL_Exe()` | Public | Sub | SQL実行 |
| ✓ | 8 | **SQL_Execution.bas** | `SQL_INSERT_UPDATE()` | Public | Sub | INSERT/UPDATE汎用処理 |
| ✓ | 9 | **SQL_Execution.bas** | `SQL_Delete()` | Public | Sub | DELETE汎用処理 |
| ✓ | 10 | **SQL_Execution.bas** | `Disp_Sheet()` | Public | Sub | SQL結果をシートに出力 |
| ✓ | 11 | **SQL_Execution.bas** | `Set_Array()` | Public | Sub | SQL結果を配列に格納 |
|  | 12 | **終了処理.bas** | `Bookを閉じる()` | Public | Sub | ブックを閉じる（最後の1つならExcel終了） |
| ✓ | 13 | **初期化.bas** | `クリア()` | Public | Sub | カレンダー休日データクリア（ショートカット Ctrl+E） |

### 6.1 ThisWorkbook.cls

#### `Workbook_Open()`

**処理概要**: ブック起動時の初期化。ウィンドウ最大化、シート保護設定、休日データのクリア、ズーム調整を行う。

**処理フロー**:
1. ウィンドウを最大化
2. カレンダーシートを選択
3. シート保護を解除後、再設定（UIのみ許可）
4. 編集可能セルをアンロックセルのみに制限
5. `YasumiDay`, `YasumiTB` をクリア
6. A1:AN37 を選択してウィンドウに合わせてズーム
7. `HyoujiBi` セルにフォーカス

```vba
Private Sub Workbook_Open()
    ActiveWindow.WindowState = xlMaximized
    Worksheets("ｶﾚﾝﾀﾞｰ").Select
    ActiveSheet.Unprotect
    ActiveSheet.Protect UserInterfaceOnly:=True
    ActiveSheet.EnableSelection = xlUnlockedCells
    Range("YasumiDay") = ""
    Range("YasumiTB") = ""
    Range("A1:AN37").Select
    ActiveWindow.Zoom = True
    Range("HyoujiBi").Select
End Sub
```

#### `Workbook_BeforeClose(Cancel As Boolean)`

**処理概要**: 保存ダイアログを抑制してブックを閉じる。

---

### 6.2 Sheet1.cls

#### `Worksheet_BeforeRightClick(ByVal Target As Range, Cancel As Boolean)`

**処理概要**: カレンダー日付セルの右クリックで休日（「休」）のトグル登録を行う。

**処理フロー**:
1. `Debug` セルが空でない場合は処理をスキップ
2. クリック位置が日付表示行（行7〜37）かつ休日フラグ列（列5,8,11...など3の倍数+2列目）かを判定
3. 左隣の日付セルが空の場合は何もしない（その日が存在しない）
4. 「休」が既に入力されていれば空に戻す（トグルOFF）
5. 空であれば「休」を入力（トグルON）
6. 右クリックメニュー表示をキャンセル（`Cancel = True`）

```vba
Private Sub Worksheet_BeforeRightClick(ByVal Target As Range, Cancel As Boolean)
    If Range("Debug") <> "" Then Exit Sub
    With Target
        If (.Row > 6 And .Row < 38) And _
           (.Column > 4 And .Column < 39 And .Column - Int(.Column / 3) * 3 = 2) Then
            If Worksheets("ｶﾚﾝﾀﾞｰ").Cells(.Row, .Column - 1) = "" Then
                Worksheets("ｶﾚﾝﾀﾞｰ").Cells(.Row, .Column) = ""
            Else
                If Worksheets("ｶﾚﾝﾀﾞｰ").Cells(.Row, .Column) = "休" Then
                    Worksheets("ｶﾚﾝﾀﾞｰ").Cells(.Row, .Column) = ""
                Else
                    Worksheets("ｶﾚﾝﾀﾞｰ").Cells(.Row, .Column) = "休"
                End If
            End If
        End If
    End With
    Cancel = True
End Sub
```

#### `Worksheet_Change(ByVal Target As Range)`

**処理概要**: `HyoujiBi` セルの変更時に、入力日付を月初日に正規化して `休日読込` を呼び出す。

**処理フロー**:
1. イベントを無効化
2. 変更セルが `HyoujiBi` の場合:
   - 空の場合: `YasumiDay`, `YasumiTB` をクリア
   - 入力あり: `YYYY/M/1` 形式に正規化後、`休日読込` を呼び出す
3. イベントを再有効化

---

### 6.3 Ex休み読込.bas

#### `休日読込()`

**処理概要**: DBから休祭日データを取得し、カレンダーシートに反映する。

**処理フロー**:
1. `HyujiS`（検索開始日）以降の `FLG LIKE '休%'` の休日レコードをDBから取得
2. 取得データを `YasumiTB`（AQ7:AR372）に上書き
3. 自動計算を有効化
4. テンプレート行4の数式（E4:F4）を行7〜37の各日付列に数式→値の順で貼付け（12か月分）
5. 画面更新を再開

```sql
SELECT KYUUJITU1, flg FROM ExYasumiX
WHERE FLG LIKE '休%' AND KYUUJITU1 >= <HyujiS>
ORDER BY KYUUJITU1
```

---

### 6.4 Ex休み更新.bas

#### `YasumiKousinn()`

**処理概要**: シートの休日フラグ（Col2）が前回値（Col3）から変更されている行を検出し、DBテーブル `ExYasumiX` をINSERT/UPDATEする。

**処理フロー**:
1. 更新確認ダイアログ（いいえで中止）
2. 12か月分のデータ列を走査（Step 3: D, G, J...列）
3. 各月の日付行（行7〜）を日付セルが空になるまでループ
4. Col2（休日フラグ）≠ Col3（前回値）の場合:
   - `kyuujitu1`: 日付シリアル値, `flg`: 休日フラグ, `kousinn`: 更新日時
   - `SQL_INSERT_UPDATE` で `ExYasumiX` をINSERT/UPDATE
   - Col3にCol2の値をコピー（比較基準を更新）
5. 「更新しました」メッセージ表示

---

### 6.5 SQL_Execution.bas

**処理概要**: ADO/ODBC DB接続・SQL実行の共通基盤（他ファイルと同構造）。DSN=ricdb, UID=ric, PWD=t6101 でOracleに接続。

---

### 6.6 終了処理.bas

#### `Bookを閉じる()`

**処理概要**: ブック数が1の場合はExcelを終了、複数の場合は当ブックのみを閉じる。

---

### 6.7 初期化.bas

#### `クリア()`

**処理概要**: `HyoujiBi`, `YasumiDay`, `YasumiTB` を空にしてカレンダーを初期化する。

```vba
Sub クリア()
    Range("HyoujiBi") = ""
    Range("YasumiDay") = ""
    Range("YasumiTB") = ""
End Sub
```

---

## 7. ユーザーフォーム仕様

なし。（ユーザーフォームなし）。

## 8. DB 接続・外部連携

### 8.1 ODBC 接続設定

| DSN 名 | UID | PWD | 用途 |
| --- | --- | --- | --- |
| `ricdb` | `ric` | `t6101` | 照射管理システムDB — 休日テーブルの読み書き |

### 8.2 テーブル一覧（参照/更新区分付き）

> ✓ = INSERT / UPDATE / DELETE の対象テーブル（参照のみのテーブルは ✓ なし）

| ✓ | No | テーブル名 | 区分 | 主な用途 | キー列 | 参照/更新列 |
| --- | --- | --- | --- | --- | --- | --- |
| ✓ | 1 | `ExYasumiX` | **参照＋更新** | 休祭日テーブル（※「マスタ」は推論。VBAメッセージ「休祭日を更新しますか？」より休祭日の用途は確認済み） | `KYUUJITU1` | 参照: `KYUUJITU1`, `FLG`。更新: `kyuujitu1`, `flg`, `kousinn`（`YasumiKousinn` から INSERT/UPDATE） |

> **「キー列」の定義**: JOIN 条件または UPDATE/DELETE の WHERE 句で使用される列を示す。

### 8.3 SQL 一覧

#### 8.3.1 休日読み込み（`休日読込()` / **Ex休み読込.bas**）

```sql
SELECT KYUUJITU1, flg FROM ExYasumiX
WHERE FLG LIKE '休%' AND KYUUJITU1 >= <HyujiS>
ORDER BY KYUUJITU1
```

#### 8.3.2 休日更新（`YasumiKousinn()` / **Ex休み更新.bas**）

```sql
SELECT COUNT(*) FROM ExYasumiX WHERE kyuujitu1=<date_serial>
INSERT INTO ExYasumiX (kyuujitu1, flg, kousinn) VALUES(<date>, '<flg>', <timestamp>)
UPDATE ExYasumiX SET flg='<flg>', kousinn=<timestamp> WHERE kyuujitu1=<date_serial>
```

---


### 8.4 外部ファイル連携

なし。

## 9. データフロー

各フローは「起点 → 処理 → 結果」の粒度で記述する。

### 9.1 起動フロー

| No | 起点 | 処理 | 結果 |
| --- | --- | --- | --- |
| 1 | 📄 ブックOpen | 📄 `Workbook_Open()` / **ThisWorkbook.cls** | 起動処理開始 |
| 2 | 上記 | 📊 「ｶﾚﾝﾀﾞｰ」選択・保護設定（UIのみ許可） | シート操作可能な保護状態 |
| 3 | 上記 | 📊 `YasumiDay` / `YasumiTB` クリア | 休日領域を初期化 |
| 4 | 上記 | 📊 ズーム設定 → `HyoujiBi` にフォーカス | 年月入力待ち |

#### ツリー図（補助）

```
（ブックOpen）
└─ 📄 Workbook_Open              [ThisWorkbook.cls]
   ├─ 📊 「ｶﾚﾝﾀﾞｰ」Select + Protect（UIのみ）
   ├─ 📊 YasumiDay / YasumiTB クリア
   └─ 📊 Zoom → HyoujiBi.Select
```

### 9.2 カレンダー表示フロー

| No | 起点 | 処理 | 結果 |
| --- | --- | --- | --- |
| 1 | 📊 `HyoujiBi`（G1）に年月入力 | 📄 `Worksheet_Change()` / **Sheet1.cls** | 変更を検知 |
| 2 | 上記 | 📄 `YYYY/M/1` に正規化 | 月初日に変換 |
| 3 | 上記 | 🗄️ `休日読込()` / **Ex休み読込.bas** | `ExYasumiX` から休日を SELECT |
| 4 | 上記 | 📊 `YasumiTB`（AQ7:AR372）に書込み | 休日データを展開 |
| 5 | 上記 | 📊 行4テンプレートを行7〜37へ数式→値コピー（12か月分） | 各日付セルに「休」または空 |

#### ツリー図（補助）

```
（HyoujiBi に年月入力）
└─ 📄 Worksheet_Change           [Sheet1.cls]
   ├─ 📄 YYYY/M/1 に正規化
   └─ 🗄️ 休日読込                 [Ex休み読込.bas]
      ├─ 🗄️ SELECT ExYasumiX（FLG LIKE '休%'）
      ├─ 📊 YasumiTB（AQ7:AR372）に書込み
      └─ 📊 行4 VLOOKUP を行7〜37へ数式→値コピー
```

### 9.3 手動休日入力フロー

| No | 起点 | 処理 | 結果 |
| --- | --- | --- | --- |
| 1 | 📊 日付セルを右クリック | 📄 `Worksheet_BeforeRightClick()` / **Sheet1.cls** | 「休」をトグル（ON/OFF） |
| 2 | 📊（休日登録 Click） | 📄 `YasumiKousinn()` / **Ex休み更新.bas** | 変更検出を開始 |
| 3 | 上記 | 📄 Col2 ≠ Col3 の行を検出 | 更新対象を抽出 |
| 4 | 上記 | 🗄️ `SQL_INSERT_UPDATE()` / **SQL_Execution.bas** | `ExYasumiX` を INSERT/UPDATE |
| 5 | 上記 | 🖥️ MsgBox「更新しました」 | 完了通知 |

#### ツリー図（補助）

```
（日付セルを右クリック）
└─ 📄 Worksheet_BeforeRightClick [Sheet1.cls]
   └─ 📊 「休」トグル
      └─ 📊（休日登録 Click）
         └─ 📄 YasumiKousinn      [Ex休み更新.bas]
            ├─ 📄 Col2 ≠ Col3 を検出
            ├─ 🗄️ SQL_INSERT_UPDATE [SQL_Execution.bas]（ExYasumiX）
            └─ 🖥️ MsgBox 完了通知
```

---

## 10. セキュリティ注意事項


| No | カテゴリ | 内容 | リスク |
| --- | --- | --- | --- |
| 1 | 認証情報ハードコード | DSN=`ricdb`, UID=`ric`, PWD=`t6101` が **SQL_Execution** モジュールに平文記載 | 中：VBAエディタで閲覧可能 |

## スコープ外（本仕様書に含まないもの）

- セル書式（色・罫線・フォント）
- 条件付き書式、グラフ・画像、印刷設定

必要な場合は Excel 画面のスクリーンショットで補完してください。
