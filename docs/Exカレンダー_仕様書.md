# Exカレンダー 仕様書

> **ファイル種別**: .xlsm（マクロ付き）
> **用途**: 指定月から12か月分のカレンダーを表示し、会社休祭日をDBテーブル `ExYasumiX` と連携して管理する
> **VBA プロジェクトサイズ**: 7モジュール（ThisWorkbook, Sheet1, Ex休み読込, Ex休み更新, SQL_Execution, 終了処理, 初期化）
> **外部連携ファイル**: なし（DB直接接続）

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
Exカレンダー.xlsm
├── シート
│   └── ｶﾚﾝﾀﾞｰ（12か月カレンダー表示・休日管理）
└── VBA モジュール
    ├── ThisWorkbook.cls    （Workbook_Open / Workbook_BeforeClose）
    ├── Sheet1.cls          （Worksheet_BeforeRightClick / Worksheet_Change）
    ├── Ex休み読込.bas      （DBから休日データを読み込みシートに貼付け）
    ├── Ex休み更新.bas      （シートの休日をDBに書き戻し）
    ├── SQL_Execution.bas   （ADO/ODBC DB接続・SQL実行基盤）
    ├── 終了処理.bas        （ブック終了処理）
    └── 初期化.bas          （クリア処理）
```

---

## 2. シート詳細

### 2.1 ｶﾚﾝﾀﾞｰ

**目的**: `HyoujiBi`（G1）に表示月を入力すると、その月を起点に12か月分のカレンダーが横方向に生成される。各日付に対しDBから休祭日フラグを読み込み表示する。右クリックで会社独自の休日を手動追加・削除でき、「休日登録」ボタンでDBに反映する。

#### カラム構成（3列 × 12か月）

カレンダーは D列から AM列まで、3列1セットで12か月分が横に展開される。

| 月番号 | 日付列（Col1）| 休日フラグ列（Col2）| 更新前値列（Col3）|
|---|---|---|---|
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
|---|---|
| 1 | ヘッダー（D1: ラベル「検索開始日」, G1:H1: `HyoujiBi`入力セル, J1: 操作説明, AL1: システムID「Ex5」） |
| 2 | `Honnjitu`（AQ2: 本日日付） |
| 3 | （空） |
| 4 | 休日フラグテンプレート行（VLOOKUPで `YasumiTB` を参照）。`休日読込` でこの行の数式を行7〜37にコピー・値貼付け |
| 5 | （空） |
| 6 | 各月の月初日（日付データ: D列=`HyoujiBi`、以降 EDATE で+1か月） |
| 7 | 各月の1日（INT変換した月初日） |
| 8〜37 | 各月の2日〜31日（+1日ずつ加算、月をまたいだら空） |
| 6〜37 | AP列以降 `YasumiTB`: DBから読み込んだ休日テーブル（KYUUJITU1, FLG） |

#### DB読み込み領域

| 列 | 行範囲 | 内容 |
|---|---|---|
| AP | 6（ヘッダー "No"） | 番号 |
| AQ | 6〜372 | `KYUUJITU1`（休日日付: Excelシリアル値） |
| AR | 6〜372 | `FLG`（フラグ文字列。`休` で始まるもの） |

---

## 3. 名前付き範囲一覧

| 名前 | 参照先 | 説明 |
|---|---|---|
| `Debug` | ｶﾚﾝﾀﾞｰ!$A$1 | デバッグモード（空=通常動作、非空=右クリック処理をスキップ） |
| `Honnjitu` | ｶﾚﾝﾀﾞｰ!$AQ$2 | 本日の日付 |
| `HyoujiBi` | ｶﾚﾝﾀﾞｰ!$G$1 | カレンダー表示開始月入力セル（YYYY/M/1 形式に正規化される） |
| `HyujiS` | ｶﾚﾝﾀﾞｰ!$D$6 | 検索開始日付（D6 = `HyoujiBi` をコピー: `=IF(HyoujiBi="","",HyoujiBi)`） |
| `YasumiDay` | E7:F37, H7:I37, K7:L37 ... AM7:AM37（12か月分） | 休日フラグ + 更新前値セル群（変更検出と一括クリア用） |
| `YasumiTB` | ｶﾚﾝﾀﾞｰ!$AQ$7:$AR$372 | DBから読み込んだ休日テーブル（VLOOKUPのルックアップ範囲） |

---

## 4. 数式一覧

### ｶﾚﾝﾀﾞｰシート

すべての数式は3列単位 × 12か月で繰り返されるため、代表パターンを示す。

#### 行6: 月初日計算

| セル（代表） | 数式 | 説明 |
|---|---|---|
| D6 | `=IF(HyoujiBi="","",HyoujiBi)` | 1か月目の月初日（`HyoujiBi` と同値） |
| G6 | `=IF(HyujiS="","",EDATE(D6,1))` | 2か月目の月初日（前月+1か月） |
| J6〜AK6 | `=IF(HyujiS="","",EDATE(<前月>,1))` | 3〜12か月目（同パターン） |

#### 行4: 休日フラグテンプレート（VLOOKUP）

| セル（代表） | 数式 | 説明 |
|---|---|---|
| E4, F4 | `=IF(ISERROR(VLOOKUP(D4,YasumiTB,2,FALSE)),"",LEFT(VLOOKUP(D4,YasumiTB,2,FALSE),1))` | D列の日付で `YasumiTB` を検索し、FLGの先頭1文字（「休」等）を取得 |
| H4, I4〜 | 同パターン（参照列が3列ずつシフト） | — |

この行が `休日読込` で行7〜37に数式→値貼付けされる。

#### 行7: 月初日整数化

| セル（代表） | 数式 | 説明 |
|---|---|---|
| D7 | `=IF(HyujiS="","",INT(HyujiS))` | 月初日をシリアル整数値に変換 |
| G7〜AK7 | `=IF(HyujiS="","",INT(<月6セル>))` | 同パターン |

#### 行8〜37: 日付の連続生成

| セル（代表） | 数式 | 説明 |
|---|---|---|
| D8 | `=IF(D7="","",IF(MONTH(D7)=MONTH(D7+1),D7+1,""))` | 前日+1日。翌日が別月なら空（月末処理） |
| D9〜D37 | 同パターン（D8→D9、D9→D10…） | 最大31日分 |
| G8〜AK37 | 同パターン | 各月列で繰り返し |

---

## 5. ボタン・マクロ対応

| シート | ボタンラベル | 割り当てマクロ | 動作概要 |
|---|---|---|---|
| ｶﾚﾝﾀﾞｰ | 休日登録 | `YasumiKousinn` | シートの休日フラグ変更をDBテーブル `ExYasumiX` に書き込む |
| ｶﾚﾝﾀﾞｰ | 終了 | `Bookを閉じる` | ブックを保存せずに閉じる |

---

## 6. VBA モジュール仕様

### 6.1 ThisWorkbook.cls

#### `Workbook_Open()`

**処理概要**: ブック起動時の初期化。ウィンドウ最大化、シート保護設定、休日データのクリア、ズーム調整を行う。

**処理フロー**:
1. ウィンドウを最大化
2. ｶﾚﾝﾀﾞｰシートを選択
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

## 7. DB 接続・外部連携

### ODBC 接続設定

| DSN 名 | UID | PWD | 用途 |
|---|---|---|---|
| `ricdb` | `ric` | `t6101` | 休日テーブルの読み書き |

### 参照テーブル一覧

| テーブル名 | 主な用途 | 主要カラム |
|---|---|---|
| `ExYasumiX` | 休祭日マスタ | `KYUUJITU1`（日付: Excelシリアル値）, `FLG`（フラグ: `休` で始まる文字列）, `kousinn`（更新日時） |

### 主要 SQL 文

```sql
-- 休日読み込み（月初以降の全休日取得）
SELECT KYUUJITU1, flg FROM ExYasumiX
WHERE FLG LIKE '休%' AND KYUUJITU1 >= <HyujiS>
ORDER BY KYUUJITU1

-- 休日更新（YasumiKousinn から SQL_INSERT_UPDATE 経由）
-- 存在チェック
SELECT COUNT(*) FROM ExYasumiX WHERE kyuujitu1=<date_serial>
-- 新規
INSERT INTO ExYasumiX (kyuujitu1, flg, kousinn) VALUES(<date>, '<flg>', <timestamp>)
-- 更新
UPDATE ExYasumiX SET flg='<flg>', kousinn=<timestamp> WHERE kyuujitu1=<date_serial>
```

---

## 8. データフロー

```
【起動フロー】
  Workbook_Open
       ↓
  ｶﾚﾝﾀﾞｰシート表示 + 保護設定
       ↓
  YasumiDay / YasumiTB クリア
       ↓
  ズーム設定 → HyoujiBi にフォーカス

【カレンダー表示フロー】
  HyoujiBi（G1）に年月入力
       ↓ Worksheet_Change
  YYYY/M/1 に正規化
       ↓
  休日読込(): DB(ExYasumiX) → YasumiTB（AQ7:AR372）に書き込み
       ↓
  行4テンプレート（VLOOKUP数式）を行7〜37に数式→値コピー
  （12か月分 × E:F列 / H:I列... のペア）
       ↓
  各日付セルに「休」または空が設定されたカレンダーが完成

【手動休日入力フロー】
  カレンダー日付セルを右クリック
       ↓ Worksheet_BeforeRightClick
  「休」のトグル（ON/OFF）
       ↓
  「休日登録」ボタン押下
       ↓ YasumiKousinn()
  変更セル（Col2 ≠ Col3）を検出
       ↓
  SQL_INSERT_UPDATE: ExYasumiX に更新
```

---

## 9. セキュリティ注意事項

| 種別 | キーワード | 内容 |
|---|---|---|
| AutoExec | `Workbook_Open` | ブックオープン時に自動実行。シート保護設定・DB読み込みを行う |
| AutoExec | `Workbook_BeforeClose` | ブック終了時に自動実行。保存ダイアログ抑制 |
| AutoExec | `Worksheet_Change` | 表示月変更時に自動実行。DB接続が発生する |
| Suspicious | `Open` | ADO接続の `oraconn.Open` が該当 |
| Suspicious | `Call` | 各サブルーチン呼び出しで使用 |
| Suspicious | Hex Strings / Base64 Strings | VBAバイナリ内のエンコード。実際のエンコード処理なし |
| 注意 | DB認証情報 | 接続文字列にUID/PWDをハードコード: `DSN=ricdb;UID=ric;PWD=t6101` |
