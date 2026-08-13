# ExGM_EB会社コード変換TB 仕様書

> **ファイル種別**: .xlsm（マクロ付き）
> **用途**: ガンマシステムとEBシステムの顧客コードを紐付け登録・管理する変換テーブルメンテナンス画面（EXメニューの1ファイルとして照射管理システムを補完）
> **VBA プロジェクト**: 9モジュール（ThisWorkbook, Sheet1, 登録, リンク解消, クリア, 該当表示, 共通変数, 全データ表示, SQL_Execution）
> **外部連携**: なし（DB直接接続）

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
| ✓ | 1 | 「登録」 | — | — | visible | — |

### 1.3 VBA モジュール一覧

> ✓ = ユーザー操作の起点 / DB I/O を含む / 他モジュールから呼び出される / コード行数上位 25%

| ✓ | No | モジュール | 種別 | プロシージャ数 | 主な役割 |
| --- | --- | --- | --- | --- | --- |
| ✓ | 1 | **ThisWorkbook.cls** | .cls | 2 | 起動/終了イベント |
| ✓ | 2 | **Sheet1.cls** | .cls | 1 | 「登録」シートイベント |
|  | 3 | **共通変数.bas** | .bas | — | Public 変数定義 |
| ✓ | 4 | **SQL_Execution.bas** | .bas | — | ADO/ODBC DB接続・SQL実行基盤 |
| ✓ | 5 | **登録.bas** | .bas | — | 登録更新処理 |
| ✓ | 6 | **リンク解消.bas** | .bas | — | ペア削除処理 |
|  | 7 | **クリア.bas** | .bas | — | 画面クリア処理 |
| ✓ | 8 | **該当表示.bas** | .bas | — | 会社名表示・既存紐付チェック |
| ✓ | 9 | **全データ表示.bas** | .bas | — | 一覧表示処理 |

---

## 2. シート詳細


### 2.0b 非表示行・列一覧

| シート | 非表示行 | 非表示列 |
| --- | --- | --- |
| 「登録」 | 4, 8 | なし |

### 2.0 シート可視性一覧

> **ブック設定**: showSheetTabs: 非表示、showHorizontalScroll: 非表示

| No | シート名 | 可視性 |
| --- | --- | --- |
| 1 | 「登録」 | 表示 |

### 2.1 登録

**目的**: ガンマ顧客コードとEB顧客コードの紐付けを新規登録・更新・削除し、登録済み一覧を表示する。

#### レイアウト構造

| セル範囲 | 名前付き範囲 | 内容 |
| --- | --- | --- |
| A1 | — | システム識別子（`Ex4`） |
| B2 | — | タイトル「顧客コード変換テーブル登録」 |
| B5 | — | 入力ヘッダー「ガンマ顧客コード」 |
| C5 | — | 入力ヘッダー「EB顧客コード」 |
| D5 | — | 入力ヘッダー「会社名」 |
| E5 | — | 入力ヘッダー「住所１」 |
| F5 | — | 入力ヘッダー「住所2」 |
| B6 | `GMTNo` / `Touroku`(B6:F6) | ガンマ顧客コード入力欄 |
| C6 | `EBTNo` / `Touroku`(B6:F6) | EB顧客コード入力欄 |
| D6 | `KaiName` / `Touroku`(B6:F6) | 会社名表示欄（DB参照） |
| E6 | `Jyuusyo1` / `Touroku`(B6:F6) | 住所１表示欄（DB参照） |
| F6 | `Jyuusyo2` / `Touroku`(B6:F6) | 住所２表示欄（DB参照） |
| B7 | — | セクションラベル「登録済一覧」 |
| B9 | — | 一覧ヘッダー「ガンマ顧客コード」 |
| C9 | — | 一覧ヘッダー「EB顧客コード」 |
| D9 | — | 一覧ヘッダー「会社名」 |
| E9 | — | 一覧ヘッダー「住所１」 |
| F9 | — | 一覧ヘッダー「住所2」 |
| B10:F193 | `AllData` | DB取得データ表示領域（最大184件） |

（行番号 8 は空）

#### DB読み込み列

| 列 | ヘッダー名 | DB フィールド |
| --- | --- | --- |
| B | ガンマ顧客コード | `kcdcnvmst.kaisyacd` |
| C | EB顧客コード | `kcdcnvmst.ebkaisyacd` |
| D | 会社名 | `tokumst.coname` |
| E | 住所１ | `tokumst.jyus1` |
| F | 住所２ | `tokumst.jyus2` |

---

## 3. 名前付き範囲一覧

| No | 名前 | 参照先 | 説明 |
| --- | --- | --- | --- |
| 1 | `AllData` | 登録!$B$10:$F$193 | DB取得データ一覧表示領域（最大184行） |
| 2 | `EBTNo` | 登録!$C$6 | EB顧客コード入力セル |
| 3 | `GMTNo` | 登録!$B$6 | ガンマ顧客コード入力セル |
| 4 | `Jyuusyo1` | 登録!$E$6 | 住所１表示セル |
| 5 | `Jyuusyo2` | 登録!$F$6 | 住所２表示セル |
| 6 | `KaiName` | 登録!$D$6 | 会社名表示セル |
| 7 | `Touroku` | 登録!$B$6:$F$6 | 入力行全体（一括クリア用） |


### 3.1 データの入力規則

| シート | セル | 種別 | 制約 | 用途 |
| --- | --- | --- | --- | --- |
| 「登録」 | `B9:F193` | 整数 | =5180630 |  |


---

---

## 4. 数式一覧

数式セルなし（全データはVBAによりDB取得・シート書き込み）。

---

## 5. ボタン・マクロ対応

| シート | ボタンラベル | 割り当てマクロ | 動作概要 |
| --- | --- | --- | --- |
| 登録 | 登録 | `登録更新` | 入力したガンマ/EBコードをDBの`kcdcnvmst`テーブルに登録または更新する |
| 登録 | リンク解消 | `PairDelete` | 入力したガンマ/EBコードの紐付けをDBから削除する |
| 登録 | 終了 | `Bookを閉じる` | ブックを保存せずに閉じる（他のブックがなければExcel終了） |

### 5.1 ショートカットキー

| No | マクロ名 | ショートカット | 処理概要 |
| --- | --- | --- | --- |
| 1 | `画面クリア()` | **Ctrl+E** | 入力欄・一覧領域をクリアし初期状態に戻す |

---

## 6. VBA モジュール仕様

### 6.0 全プロシージャ一覧

| ✓ | No | モジュール | プロシージャ | スコープ | 種別 | 概要 |
| --- | --- | --- | --- | --- | --- | --- |
| ✓ | 1 | **ThisWorkbook.cls** | `Workbook_Open()` | Private | Event | シート保護・画面クリア・一覧表示・フィルター有効化 |
| ✓ | 2 | **ThisWorkbook.cls** | `Workbook_BeforeClose()` | Private | Event | 保存ダイアログ抑止 |
| ✓ | 3 | **Sheet1.cls** | `Worksheet_Change()` | Private | Event | セル変更時に会社名DB検索・既存紐付チェック |
|  | 4 | **共通変数.bas** | （宣言のみ） | — | — | Public変数定義（mpGMKaicd, mpEBKaicd, mpKaiName） |
| ✓ | 5 | **SQL_Execution.bas** | `Open_oraconDB()` | Public | Sub | Oracle DB接続（ODBC） |
| ✓ | 6 | **SQL_Execution.bas** | `SQL_Exe()` | Public | Sub | SQL実行 |
| ✓ | 7 | **SQL_Execution.bas** | `SQL_INSERT_UPDATE()` | Public | Sub | INSERT/UPDATE汎用処理 |
| ✓ | 8 | **SQL_Execution.bas** | `SQL_Delete()` | Public | Sub | DELETE汎用処理 |
| ✓ | 9 | **SQL_Execution.bas** | `Disp_Sheet()` | Public | Sub | SQL結果をシートに出力 |
| ✓ | 10 | **SQL_Execution.bas** | `Set_Array()` | Public | Sub | SQL結果を配列に格納 |
| ✓ | 11 | **登録.bas** | `登録更新()` | Public | Sub | ガンマ/EBコード組み合わせをDBに登録・更新 |
| ✓ | 12 | **リンク解消.bas** | `PairDelete()` | Public | Sub | ガンマ/EBコード紐付けをDBから削除 |
| ✓ | 13 | **クリア.bas** | `画面クリア()` | Public | Sub | 入力欄・一覧領域クリア（ショートカット Ctrl+E） |
|  | 14 | **クリア.bas** | `EventOn()` | Public | Sub | イベント強制再有効化（デバッグ用） |
| ✓ | 15 | **該当表示.bas** | `会社名表示()` | Public | Sub | ガンマコードで会社名・住所をDB検索表示 |
| ✓ | 16 | **該当表示.bas** | `既存紐付チェック()` | Public | Sub | EBコードの既存リンク有無を確認 |
| ✓ | 17 | **全データ表示.bas** | `AllHyouji()` | Public | Sub | DB登録済みデータを一覧表示 |

### 6.1 ThisWorkbook.cls

#### `Workbook_Open()`

**処理概要**: ブックを開いた際の初期化処理。シート保護設定・画面クリア・一覧表示・オートフィルター有効化を行う。

**処理フロー**:
1. シート保護を一旦解除
2. フィルタリング・ソート・描画オブジェクトを許可しながらシート保護を再設定（UI操作のみ許可）
3. `画面クリア` を呼び出し入力欄・一覧領域をクリア
4. `AllHyouji` を呼び出しDB登録済みデータを一覧表示
5. B9セルでオートフィルターを有効化
6. B6セルにフォーカスを移動

```vba
Private Sub Workbook_Open()
    ActiveSheet.Unprotect
    ActiveSheet.Protect UserInterfaceOnly:=True _
        , DrawingObjects:=True _
        , AllowFiltering:=True _
        , AllowSorting:=True
    Call 画面クリア
    Call AllHyouji
    Range("B9").AutoFilter
    If ActiveSheet.AutoFilterMode = False Then Range("B9").AutoFilter
    Range("B6").Select
End Sub
```

#### `Workbook_BeforeClose(Cancel As Boolean)`

**処理概要**: ブック閉じる前に保存ダイアログを抑制する。

```vba
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Application.DisplayAlerts = False
    ActiveWorkbook.Saved = True
End Sub
```

---

### 6.2 Sheet1.cls

#### `Worksheet_Change(ByVal Target As Range)`

**処理概要**: 入力セル変更イベント。ガンマコード変更時は会社名をDB検索表示、EBコード変更時は既存紐付チェックを実行する。

**処理フロー**:
1. イベントを無効化して再帰防止
2. 変更セルが B6（ガンマ顧客コード）の場合:
   - 空の場合: `Touroku`範囲をクリア
   - 入力あり: 4桁ゼロ埋め後、`会社名表示` を呼び出す
3. 変更セルが C6（EB顧客コード）かつ入力ありの場合:
   - 3桁ゼロ埋め後、ガンマコード未入力ならエラーメッセージ
   - ガンマコードあり: `既存紐付チェック` を呼び出す
4. イベントを再有効化

```vba
Private Sub Worksheet_Change(ByVal Target As Range)
    With Target
        Application.EnableEvents = False
        If .Row = 6 And .Column = 2 Then
            If Range("GMTNo") = "" Then
                Range("Touroku") = ""
            Else
                Range("GMTNo") = Right("0000" & Range("GMTNo"), 4)
                Call 会社名表示
            End If
        ElseIf .Row = 6 And .Column = 3 And Range("EBTNo") <> "" Then
            Range("EBTNo") = Right("000" & Range("EBTNo"), 3)
            If Range("GMTNo") = "" Then
                MsgBox "最初にｶﾞﾝﾏ顧客ｺｰﾄﾞを指定してください。"
                Range("EBTNo") = ""
                Range("GMTNo").Select
            Else
                Call 既存紐付チェック
                Range("EBTNo").Select
            End If
        End If
        Application.EnableEvents = True
    End With
End Sub
```

---

### 6.3 共通変数.bas

**処理概要**: 全モジュールで共有するPublicグローバル変数を定義する。

```vba
Public mpGMKaicd As String  'ｶﾞﾝﾏの会社ｺｰﾄﾞ
Public mpEBKaicd As String  'EBの会社ｺｰﾄﾞ
Public mpKaiName As String  '会社名
```

---

### 6.4 SQL_Execution.bas

**処理概要**: ADO + ODBC を使用したDB接続・SQL実行の共通基盤モジュール。SELECT/INSERT/UPDATE/DELETE・シート貼り付け・配列取得の各処理を提供する。

#### `Open_oraconDB()`

**処理概要**: DSN `ricdb` に ADO接続を開く。

```vba
Sub Open_oraconDB()
    oraconn.ConnectionString = "DSN=ricdb;UID=ric;PWD=t6101"
    oraconn.Open
    oraconn.CursorLocation = adUseClient
End Sub
```

#### `SQL_Exe(mySQL As String)`

**処理概要**: SQL文を実行しレコードセット `rs` に結果を格納する。

#### `SQL_INSERT_UPDATE(myTBL, myKey, myD(), myN)`

**処理概要**: 指定テーブルのキー条件でレコード存在を確認し、0件ならINSERT、1件以上ならUPDATEを実行する。トランザクションで保護。

**処理フロー**:
1. `Open_oraconDB` で接続
2. `BeginTrans` 開始
3. `SELECT COUNT(*)` でレコード存在確認
4. 0件: INSERT文生成・実行
5. 1件以上: UPDATE文生成・実行
6. `CommitTrans` でコミット

#### `SQL_Delete(myTBL, myWhere)`

**処理概要**: 指定テーブルから条件に合致するレコードを削除する。トランザクション保護付き。

#### `Disp_Sheet(mySQL, mySH, myRow, myRecordCount, myColumn, myFieldCount, myF)`

**処理概要**: SQLを実行しシート `mySH` の `myRow` 行・`myColumn` 列から結果を貼り付ける。`myF=1` の場合はフィールド名も1行目に出力。

#### `Set_Array(mySQL, myData(), myRecordCount, myFldCount)`

**処理概要**: SQLを実行し結果を2次元配列 `myData(i,j)`（i: レコードNo, j: フィールドNo）に格納する。

---

### 6.5 登録.bas

#### `登録更新()`

**処理概要**: 入力されたガンマ/EBコードの組み合わせをDBテーブル `kcdcnvmst` に登録または更新する。

**処理フロー**:
1. ガンマコード（`GMTNo`）・EBコード（`EBTNo`）の入力チェック
2. ①既にガンマ+EBの組み合わせが登録済みかSELECTで確認 → 登録済みならエラー終了
3. ②ガンマコードが別EBとリンク済みかSELECTで確認
4. ③EBコードが別ガンマとリンク済みかSELECTで確認
5. 両方リンク済みの場合: 登録不可メッセージ
6. ガンマのみリンク済みの場合: 既存リンク削除後に新規登録
7. EBのみリンク済みの場合: 既存リンク削除後に新規登録
8. 新規の場合: 確認ダイアログ後に新規INSERT
9. `SQL_INSERT_UPDATE` で `kcdcnvmst` へ登録
10. `AllHyouji` で一覧を再表示

```sql
-- ①既存組み合わせ確認
SELECT ebkaisyacd FROM kcdcnvmst
WHERE kaisyacd='<gmcd>' AND ebkaisyacd='<ebcd>'

-- ②ガンマコードの既存リンク確認
SELECT ebkaisyacd,kaisyacd FROM kcdcnvmst
WHERE kaisyacd='<gmcd>'

-- ③EBコードの既存リンク確認
SELECT kaisyacd,ebkaisyacd FROM kcdcnvmst
WHERE ebkaisyacd='<ebcd>'
```

---

### 6.6 リンク解消.bas

#### `PairDelete()`

**処理概要**: 入力されたガンマ/EBコードの紐付けレコードをDBから削除する。

**処理フロー**:
1. ガンマコード・EBコードの入力チェック（未入力ならエラー）
2. `kcdcnvmst` から該当ペアをSELECTで存在確認
3. 存在しない場合: 「該当データなし」メッセージ
4. 存在する場合: 確認ダイアログ後に `SQL_Delete` で削除
5. `Touroku` をクリアし、`AllHyouji` で一覧再表示

```sql
SELECT kaisyacd FROM kcdcnvmst
WHERE kaisyacd='<gmcd>' AND ebkaisyacd='<ebcd>'

DELETE kcdcnvmst WHERE kaisyacd='<gmcd>' AND ebkaisyacd='<ebcd>'
```

---

### 6.7 クリア.bas

#### `画面クリア()`

**処理概要**: 入力欄（`Touroku`）と一覧データ領域（`AllData`）をクリアする。

```vba
Sub 画面クリア()
    Range("Touroku") = ""
    Range("AllData") = ""
End Sub
```

#### `EventOn()`

**処理概要**: イベントを強制的に再有効化する（デバッグ用）。

---

### 6.8 該当表示.bas

#### `会社名表示()`

**処理概要**: 入力されたガンマ顧客コードでDBを検索し、会社名・住所・既存EBコードを入力行に表示する。

**処理フロー**:
1. `GMTNo` が空なら何もしない
2. 入力欄（`KaiName`, `Jyuusyo1`, `Jyuusyo2`, `EBTNo`）をクリア
3. ガンマコードを4桁ゼロ埋めに正規化
4. `tokumst`（顧客マスタ）と`kcdcnvmst`（変換テーブル）を結合してDB検索
5. 0件: 「照射管理システムに未登録」エラー
6. 複数件: 「管理者に連絡」エラー
7. 1件: `KaiName`, `Jyuusyo1`, `Jyuusyo2`, `EBTNo` に表示

```sql
SELECT g.coname, g.jyus1, g.jyus2, e.ebkaisyacd
FROM tokumst g, kcdcnvmst e
WHERE g.kaisyacd='<gmcd>' AND g.kaisyacd=e.kaisyacd(+)
```

#### `既存紐付チェック()`

**処理概要**: 入力されたEBコードが既に別のガンマコードとリンクされていないか確認し、リンク済みの場合はエラーを表示する。

```sql
SELECT g.kaisyacd, g.coname FROM tokumst g, kcdcnvmst e
WHERE e.ebkaisyacd='<ebcd>' AND e.kaisyacd=g.kaisyacd(+)
```

---

### 6.9 全データ表示.bas

#### `AllHyouji()`

**処理概要**: DBから全変換テーブルデータを取得し、シートのB10以降に一覧表示する。

```sql
SELECT k.kaisyacd, k.ebkaisyacd, t.coname, t.jyus1, t.jyus2
FROM tokumst t, kcdcnvmst k
WHERE t.kaisyacd=k.kaisyacd
ORDER BY k.kaisyacd
```

---

## 7. ユーザーフォーム仕様

該当なし（ユーザーフォームなし）。

## 8. DB 接続・外部連携

### ODBC 接続設定

| DSN 名 | UID | PWD | 用途 |
| --- | --- | --- | --- |
| `ricdb` | `ric` | `t6101` | 照射管理システムDB（Oracle想定） |

### 参照テーブル一覧

| No | テーブル名 | 主な用途 | 主要カラム |
| --- | --- | --- | --- |
| 1 | `kcdcnvmst` | ガンマ/EB顧客コード変換テーブル | `kaisyacd`（ガンマ顧客CD）, `ebkaisyacd`（EB顧客CD） |
| 2 | `tokumst` | 顧客マスタ（※推論: テーブル名 TOKU+MST から推定。VBAには「顧客コード」の文言あり、「顧客マスタ」の記載なし） | `kaisyacd`, `coname`（会社名）, `jyus1`（住所1）, `jyus2`（住所2） |

### 外部ファイル連携

なし

---

## 9. データフロー

各フローは「起点 → 処理 → 結果」の粒度で記述する。
```
【起動フロー】
  Workbook_Open
       ↓
  シート保護設定（UIのみ許可）
       ↓
  画面クリア（Touroku, AllData をクリア）
       ↓
  AllHyouji: DB(kcdcnvmst + tokumst) → B10以降に一覧表示
       ↓
  オートフィルター有効化 → B6にフォーカス

【登録フロー】
  B6にガンマ顧客コード入力
       ↓ Worksheet_Change
  会社名表示: DB(tokumst + kcdcnvmst) → D6/E6/F6/C6に表示
       ↓
  C6にEB顧客コード入力
       ↓ Worksheet_Change
  既存紐付チェック: DB(kcdcnvmst + tokumst) → 重複確認
       ↓
  「登録」ボタン押下
       ↓ 登録更新()
  各種チェックSQL（①②③）
       ↓
  SQL_INSERT_UPDATE: kcdcnvmst にINSERT or UPDATE
       ↓
  AllHyouji: 一覧再表示

【リンク解消フロー】
  「リンク解消」ボタン押下
       ↓ PairDelete()
  存在確認SQL
       ↓
  SQL_Delete: kcdcnvmst からDELETE
       ↓
  AllHyouji: 一覧再表示
```

---

## 10. セキュリティ注意事項



| No | カテゴリ | 内容 | リスク |
| --- | --- | --- | --- |
| 1 | AutoExec | Workbook_Open / BeforeClose 等が自動実行される | 低：意図通りの起動・終了処理 |
| 2 | ファイル操作 / DB接続 | ADO接続の `oraconn.Open` が該当。ファイル操作ではなくDB接続 | 低：DB 接続またはファイルオープン |
| 3 | 注意 | 接続文字列にUID/PWDをハードコード: `DSN=ricdb;UID=ric;PWD=t6101` | 低：olevba 検出項目 |
## スコープ外（本仕様書に含まないもの）

- セル書式（色・罫線・フォント）
- 条件付き書式、グラフ・画像、印刷設定

必要な場合は Excel 画面のスクリーンショットで補完してください。
