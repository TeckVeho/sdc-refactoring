# ExJMM60φ記入用紙 仕様書

> **ファイル種別**: .xltm（マクロ付きテンプレート）
> **用途**: 1号機照射の JMM 60φPE 丸棒線量測定記録用テンプレート。DB から受付情報を読み込み、測定値入力→線量計算→報告書印刷→ファイル保存を行う。
> **VBA プロジェクトサイズ**: 134 KB
> **外部連携ファイル**: `D:\JMM線量\<年度>\<受付番号>JMM60.xlsm`（保存先）
> **備考**: 元ファイルは `.xltm`（Excel マクロ有効テンプレート）。内部構造は `.xlsm` と同一の Open XML 形式。

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
ExJMM60φ記入用紙.xltm
├── シート
│   ├── 使用方法（操作手順書）
│   ├── 測定値（メイン入力・計算シート）
│   └── 報告書（印刷用報告書）
├── VBA モジュール
│   ├── ThisWorkbook（Workbook_Open / Workbook_BeforeClose）
│   ├── Sheet2（Worksheet_Change / Worksheet_SelectionChange）
│   ├── 印刷範囲.bas
│   ├── 計算コード.bas
│   ├── カーソル移動.bas
│   ├── 初期設定.bas
│   ├── SQL_Execution.bas
│   ├── フォルダー作成.bas
│   ├── 線量計算.bas
│   ├── 終了処理.bas
│   ├── 共通変数.bas
│   └── StartForm.frm（UserForm）
└── 外部連携
    └── DB: DSN=ricdb（Oracle ODBC）
```

---

## 2. シート詳細

### 2.1 使用方法

**目的**: オペレーター向け操作マニュアル。用紙準備・データ保存・測定値入力の手順を記載。

#### レイアウト構造

| セル範囲 | 名前付き範囲 | 内容 |
|---|---|---|
| `C1` | — | タイトル「JMM PE 丸棒線量記録方法」 |
| `B3:D30` | — | Ⅰ〜Ⅳ の操作手順説明（印刷範囲設定・保存・測定入力）|

（30 行 × 4 列、ほぼ説明テキストのみ）

---

### 2.2 測定値

**目的**: 受付番号・照射条件・測定データの入力シート。セル変更時に自動で線量計算を実行する。

#### レイアウト構造

| セル範囲 | 名前付き範囲 | 内容 |
|---|---|---|
| `E1` | — | タイトル「1 号機線量測定記録（60φPE 丸棒用）」 |
| `B2:C2` | `UkeHed` | ヘッダー「受付番号」（マージ B2:C2） |
| `B3:C3` | `Uno` | 受付番号入力欄（10 桁、マージ） |
| `D3` | `Honnsuu` | 受付本数 |
| `E3` | `Dose` | 指定線量 |
| `F3` | `SokuKiki` | 測定器 |
| `G3` | `SenKind` | 線量計種類 |
| `H3` | `keicord` | 計算式コード |
| `I3` | `Kagenn` | 下限線量 |
| `J3` | `Jyogen` | 上限線量 |
| `K3` | `ReadFlg` | 前回記録フラグ |
| `N1` | — | 「記録先フォルダー」ラベル |
| `O1` | `KirokuFolder` | 保存先フォルダーパス（`D:\JMM線量\`） |
| `O2` | `HyoujiKeta` | 小数点以下桁数（規定値 1） |
| `O3` | `Nenndo` | 年度（数式で自動計算） |
| `O4` | `KaisyaCD` | 会社コード（規定値 0725） |
| `A4:K4` | — | 列ヘッダー（バッチ / 照射月日 / 照射時間 / 測定温度 / RIC-No / 線量計番号 / 厚さ / ABS / 測定線量 / 備考 / 計算式コード） |
| `A5` | `Bach` | バッチ番号（1 から始まるシリアル） |
| `E5` | `RicNo` | RIC-No（受付本数が 1 の場合は空） |
| `G5:K544` | `DataTB` | 測定データ本体（最大 540 行） |
| `O5:S8` | `SikiPara` / `SikiTep` | 計算式パラメータ（DB から取得） |
| `Q6:Q8` | `SenKindTB` | 線量計種類テーブル（DB から取得） |
| `S6:S7` | `KeiCode` | 有効計算式コードリスト（DB から取得） |
| `P14:AI30` | `keisuu` | 計算係数テーブル |

---

### 2.3 報告書

**目的**: 客先提出用の線量測定結果報告書。測定値シートの値を参照して自動生成される印刷用フォーマット。

#### レイアウト構造

| セル範囲 | 名前付き範囲 | 内容 |
|---|---|---|
| `B2` | — | 「表１」ラベル |
| `C2` | — | タイトル「京セラメディカル様 60φPE 丸棒線量測定結果」 |
| `I2` | — | 発行元「ラジエ工業㈱」 |
| `D5` | — | 受付番号（`=測定値!B3`） |
| `B7:I159` | — | バッチ別測定結果（照射日・照射時間・RIC-No・各段線量）を繰り返しブロックで表示 |

（報告書シート: 160 行 × 33 列、多数のマージセルあり）

---

## 3. 名前付き範囲一覧

| 名前 | 参照先 | 説明 |
|---|---|---|
| `Bach` | `測定値!$A$5` | バッチ番号先頭セル |
| `DataTB` | `測定値!$G$5:$K$544` | 測定データ本体テーブル（厚さ〜計算式コード） |
| `Debug` | `測定値!$A$1` | デバッグ用セル |
| `Dose` | `測定値!$E$3` | 指定線量 |
| `Honnsuu` | `測定値!$D$3` | 受付本数 |
| `HyoujiKeta` | `測定値!$O$2` | 表示小数点以下桁数 |
| `Jyogen` | `測定値!$J$3` | 上限線量 |
| `Kagenn` | `測定値!$I$3` | 下限線量 |
| `KaisyaCD` | `測定値!$O$4` | 会社コード |
| `KeiCode` | `測定値!$S$6:$S$7` | 有効計算式コードリスト（DB から動的に更新） |
| `keicord` | `測定値!$H$3` | 現在選択中の計算式コード |
| `keicordS` | `測定値!$U$6` | 前回計算式コード（変化検知用） |
| `keisuu` | `測定値!$P$14:$AI$30` | 計算係数テーブル |
| `KirokuFolder` | `測定値!$O$1` | ファイル保存先フォルダーパス |
| `Nenndo` | `測定値!$O$3` | 年度（数式で自動計算） |
| `ReadFlg` | `測定値!$K$3` | 既存記録読込フラグ |
| `RicNo` | `測定値!$E$5` | RIC-No |
| `SenKind` | `測定値!$G$3` | 線量計種類 |
| `SenKindTB` | `測定値!$Q$6:$Q$8` | 線量計種類テーブル（DB から動的に更新） |
| `senkusyu` | `測定値!$S$5:$S$6` | 線種区分（未使用またはワーク） |
| `SikiPara` | `測定値!$X$6:$AD$10` | 計算式パラメータ全体（5 行 × 7 列） |
| `SikiTep` | `測定値!$X$6:$AD$8` | 計算式パラメータ（温度範囲部分） |
| `SokuKiki` | `測定値!$F$3` | 測定器 |
| `UkeHed` | `測定値!$B$3:$K$3` | 受付情報入力行ヘッダー |
| `Uno` | `測定値!$B$3` | 受付番号（10 桁） |

---

## 4. 数式一覧

### 測定値シート（代表）

| セル | 数式 | 説明 |
|---|---|---|
| `O3` | `=IFERROR(IF(MID(Uno,5,2)*1<4,(LEFT(Uno,4)*1-1) & "年度\",LEFT(Uno,4)& "年度\"),"")` | 受付番号の年月から年度フォルダー名を生成。4月以降は同年、3月以前は前年度扱い |
| `F5` | `=IF(OR(D3=1,E5=0),"","A"&E5&"-0")` | 受付本数が 1 または RicNo が 0 なら空。それ以外は `A{RicNo}-0` 形式の線量計番号 |
| `F6` | `=IF(OR(D3=1,E5=0),"","A"&E5&"-1")` | 同上、末尾が `-1` |

（測定値シートは全 659 件の数式あり。F 列の線量計番号自動生成パターンが大半）

### 報告書シート（代表）

| セル | 数式 | 説明 |
|---|---|---|
| `D5` | `=測定値!B3` | 受付番号を測定値シートから参照 |
| `D7` | `=IF(D11=0,"",測定値!E5)` | バッチ 1 の照射日。D11 が 0（データなし）なら空 |
| `E7` | `=IF(E11=0,"",D7+1)` | バッチ 2 の照射日（バッチ 1 の翌日） |

（報告書シートは全 901 件の数式あり。バッチごとの繰り返しパターン）

---

## 5. ボタン・マクロ対応

| シート | ボタンラベル | 割り当てマクロ | 動作概要 |
|---|---|---|---|
| 測定値 | `ｶｰｿﾙ右` | `CarsolRight` | Enter キーでカーソルを右方向に移動するよう設定 |
| 測定値 | `保存` | `フォルダー作成と保存` | 年度フォルダーを作成し、受付番号をファイル名として保存 |
| 測定値 | `終了` | `Bookを閉じる` | アラート非表示で上書きせずブックを閉じる |

---

## 6. VBA モジュール仕様

### 6.1 ThisWorkbook

#### `Workbook_Open()`

**処理概要**: ブック起動時に `初期化` を呼び出す。

**処理フロー**:
1. `Call 初期化`

#### `Workbook_BeforeClose(Cancel As Boolean)`

**処理概要**: 閉じる前に未保存チェックを行う。

**処理フロー**:
1. `mpHozon = False` かつ受付番号（`Uno`）が空でなければ確認ダイアログを表示
2. 「はい」の場合 `フォルダー作成と保存` を呼び出し、アラートを非表示にして閉じる

---

### 6.2 Sheet2（測定値シートのイベント）

#### `Worksheet_Change(ByVal Target As Range)`

**処理概要**: セル変更時に印刷範囲設定・計算式パラメータ取得・線量計算を自動実行。

**処理フロー**:
1. D3（受付本数）変更時 → `印刷範囲の設定` を呼び出し、本数が 1 なら RicNo をクリア
2. H3（計算式コード）変更時 → `計算パラメータ` を呼び出す
3. 5 行目以降の H 列（ABS 値）変更時 → 測定器・線量計種類・計算コードの入力確認後 `線量` 関数を呼び出して計算結果を書き込む

#### `Worksheet_SelectionChange(ByVal Target As Range)`

**処理概要**: 測定値入力行で厚さを前行からコピー。

**処理フロー**:
1. 6 行目以降の G 列（厚さ）が空で前行が空でない場合、前行の厚さを自動コピーし、H 列（ABS 入力列）へカーソルを移動

---

### 6.3 印刷範囲.bas

#### `印刷範囲の設定(ByVal myDataN)`

**処理概要**: 受付本数に応じて測定値シートの印刷範囲を動的に設定。

**処理フロー**:
1. 本数から行数を計算: `$A$5:$J${ Int((myDataN-1)/6)*36 + 40 }`
2. 印刷タイトル行を `$1:$4` に固定

---

### 6.4 計算コード.bas

#### `有効計算式コード()`

**処理概要**: DB から有効な計算式コードを取得し `KeiCode` 範囲に書き込む。

```vba
mySQL = "SELECT keisask FROM keicode WHERE yflg1='1' AND SUBSTR(keisask,2,1)='1' ORDER BY keisask"
Call Disp_Sheet(mySQL, "測定値", myRow, myRecordCount, myCol, myFldCount, 0)
```

#### `計算パラメータ()`

**処理概要**: 選択した計算式コードの係数（温度範囲・補正係数など）を DB から取得し `SikiPara` に書き込む。

**処理フロー**:
1. コードが前回と変わっていない場合はスキップ
2. `keicode` テーブルから温度 1〜3 の係数、温度補正係数、誤差修正係数を 5 行分取得
3. 取得後に `SikiTep` を有効/無効フラグの降順・温度昇順でソート

---

### 6.5 カーソル移動.bas

#### `CarsolRight()`

**処理概要**: Enter キーでカーソルを右方向に移動するよう Excel 設定を変更。

---

### 6.6 初期設定.bas

#### `初期化()`

**処理概要**: ブック起動時の初期設定。シート保護・カーソル方向設定・DB からマスタデータ読み込み・スタートフォーム表示。

**処理フロー**:
1. 測定値シートを選択し、UI のみシート保護を設定
2. Enter 後のカーソル方向を右に設定
3. DB から `senkind` テーブルを参照し `SenKindTB` に線量計種類を書き込む
4. `有効計算式コード` を呼び出す
5. 新規（受付番号が空かつ ReadFlg が空）の場合は `StartForm` を表示

---

### 6.7 SQL_Execution.bas

**処理概要**: ADO+ODBC を使用した汎用 DB 操作モジュール。

主要プロシージャ:
- `Open_oraconDB()` — DB 接続を開く
- `SQL_Exe(mySQL)` — SELECT 文を実行しレコードセットを取得
- `SQL_INSERT_UPDATE(myTBL, myKey, myD(), myN)` — INSERT または UPDATE を実行
- `SQL_DELETE(myTBL, myWhere)` — DELETE を実行
- `Disp_Sheet(mySQL, sheetName, ...)` — SQL 結果をシートに書き込む
- `Set_Array(mySQL, myData(), ...)` — SQL 結果を配列に格納

---

### 6.8 フォルダー作成.bas

#### `フォルダー作成と保存()`

**処理概要**: 年度フォルダーを作成し、受付番号ベースのファイル名で `SaveAs` する。

**処理フロー**:
1. `KirokuFolder` + 年度フォルダーのパスを組み立て
2. `MkDir` でフォルダーを作成（存在する場合はスキップ）
3. `ActiveWorkbook.SaveAs` でファイルを保存、`mpHozon = True` を設定

---

### 6.9 線量計算.bas

#### `線量(myR, myC)`

**処理概要**: ABS 値と温度から多項式計算で線量を算出し、結果セルに書き込む。

**処理フロー**:
1. `SikiPara` から有効な計算式を判定（最大 3 式）
2. 各式で `ABS` 値に 4 次多項式を適用し線量を算出
3. 温度補正係数（`SikiPara` 4 行目）を掛け合わせ
4. 誤差修正係数（`SikiPara` 5 行目）で最終補正
5. `HyoujiKeta` の桁数で `RoundDown` して結果を書き込む

```vba
線量 = .Cells(myRow+4, myCol) _
    + .Cells(myRow+4, myCol+1) * 線量 _
    + .Cells(myRow+4, myCol+2) * 線量^2 _
    + .Cells(myRow+4, myCol+3) * 線量^3 _
    + .Cells(myRow+4, myCol+3) * 線量^4
```

---

### 6.10 終了処理.bas

#### `Bookを閉じる()`

**処理概要**: アラート非表示のまま上書きせずにブックを閉じる。アプリケーション起動中の最後のブックの場合は Excel ごと終了。

---

### 6.11 共通変数.bas

```vba
Public mpSikiN As Integer      ' 線量計算式有効式数
Public mpDose(2) As Double     ' 線量計算式毎の線量
Public mpOndo(2) As Single     ' 線量計算式の校正温度
Public mpYFLG(4) As Integer    ' 式の有効無効配列（0-4）
Public mpOndHose As Double     ' 温度補正係数
Public mpHozon As Boolean      ' True:保存実行済み
```

---

### 6.12 StartForm.frm（UserForm）

**処理概要**: 起動時に表示されるダイアログ。受付番号入力フォーム。

コントロール: `TextUno`（受付番号入力）, `CommandButton1`（OK）, `CommandButton2`（キャンセル・終了）

#### `CommandButton1_Click()`

**処理フロー**:
1. `zaiko` テーブルから受付番号・会社コードで在庫データを取得
2. 既存記録ファイルが存在する場合はそのファイルを開いて本フォームを閉じる
3. 記録がない場合は各入力欄（Uno, Honnsuu, Dose, Kagenn, Jyogen）に DB 値をセット

```sql
SELECT nyukasu, siteisn, kagensn, jyougsn, nyukabi
FROM zaiko
WHERE uno='<受付番号>' AND syouso='1' AND kaisyacd='<会社コード>'
```

---

## 7. DB 接続・外部連携

### ODBC 接続設定

| DSN 名 | UID | PWD | 用途 |
|---|---|---|---|
| `ricdb` | `ric` | `t6101` | メイン実績 DB（Oracle） |
| `ricdb` | `rich` | `t6101` | 代替接続（コメントアウト） |

### 参照テーブル一覧

| テーブル名 | 主な用途 | 主要カラム |
|---|---|---|
| `keicode` | 計算式コードマスタ | `keisask`, `yflg1〜5`, `consta1〜5`, `constb1〜5`, ..., `temper1〜3` |
| `senkind` | 線量計種類マスタ | `sensyu` |
| `zaiko` | 在庫テーブル（受付情報） | `uno`, `kaisyacd`, `nyukasu`, `siteisn`, `kagensn`, `jyougsn`, `nyukabi`, `syouso` |

### 主要 SQL 文

```sql
-- 有効な計算式コードを取得
SELECT keisask FROM keicode WHERE yflg1='1' AND SUBSTR(keisask,2,1)='1' ORDER BY keisask

-- 計算式パラメータ（温度 1 の係数）
SELECT TO_NUMBER(conste1), TO_NUMBER(constd1), TO_NUMBER(constc1),
       TO_NUMBER(constb1), TO_NUMBER(consta1), TO_NUMBER(temper1), TO_NUMBER(yflg1)
FROM keicode WHERE keisask='<計算式コード>'

-- 線量計種類マスタ
SELECT sensyu FROM senkind

-- 在庫データ取得（スタートフォーム）
SELECT nyukasu, siteisn, kagensn, jyougsn, nyukabi
FROM zaiko WHERE uno='<受付番号>' AND syouso='1' AND kaisyacd='<会社コード>'
```

### 外部ファイル連携

| ファイル名パターン | 説明 |
|---|---|
| `D:\JMM線量\<年度>年度\<受付番号>JMM60.xlsm` | 測定記録の保存先。`フォルダー作成と保存` で生成される |

---

## 8. データフロー

```
【起動フロー】
  ブックを開く（Workbook_Open）
       ↓ 初期化
  DB(senkind): 線量計種類取得 → SenKindTB へ書き込み
       ↓
  DB(keicode): 有効計算式コード取得 → KeiCode へ書き込み
       ↓
  StartForm 表示（受付番号入力）
       ↓ CommandButton1_Click
  DB(zaiko): 受付番号で在庫データ取得
       ↓
  既存ファイルあり → SaveAs されたファイルを Workbooks.Open で開く
  新規             → Uno, Honnsuu, Dose 等の入力欄をセット

【測定値入力フロー】
  受付本数（D3）変更
       ↓ Worksheet_Change
  印刷範囲を本数に応じて自動設定
       ↓
  計算式コード（H3）変更
       ↓ Worksheet_Change
  DB(keicode): 選択コードのパラメータ取得 → SikiPara へ書き込み

  ABS 値（H列）入力
       ↓ Worksheet_Change
  線量計算（多項式 + 温度補正 + 誤差修正）
       ↓
  測定線量（I列）・計算式コード（K列）へ書き込み

【保存フロー】
  「保存」ボタン クリック
       ↓ フォルダー作成と保存
  KirokuFolder + 年度 + 受付番号 でパスを組み立て
       ↓ MkDir（フォルダーがない場合）
  ActiveWorkbook.SaveAs でファイル保存
       ↓
  mpHozon = True

【終了フロー】
  「終了」ボタン または ×ボタン
       ↓ Workbook_BeforeClose
  mpHozon = False かつ Uno が空でない場合 → 保存確認ダイアログ
       ↓ 「はい」
  フォルダー作成と保存 → Bookを閉じる
```

---

## 9. セキュリティ注意事項

| 種別 | キーワード | 内容 |
|---|---|---|
| AutoExec | `Workbook_Open` | ブックを開くと自動的に `初期化` が実行され、DB 接続・マスタ取得・UserForm 表示が走る |
| AutoExec | `Workbook_BeforeClose` | ブックを閉じる前に自動実行。保存処理を行う可能性がある |
| AutoExec | `CommandButton1_Click` | UserForm のボタン操作で DB クエリとファイル操作が実行される |
| AutoExec | `Worksheet_Change` | 測定値シートのセル変更ごとに DB クエリ・計算が自動実行される |
| Suspicious | `Open` / `Write` / `Output` | ファイル操作系ステートメントを使用。`SaveAs` でローカルファイルに書き込む |
| Suspicious | `MkDir` | ローカルディスク（`D:\JMM線量\`）にディレクトリを作成する |
| Suspicious | `Call` | 内部マクロの呼び出しに使用。Excel 4 Macro XLM ではなく通常の VBA Call |
| Suspicious | `Chr` | ダイアログの改行に `Chr(13)` を使用（難読化ではなく通常用途） |
| Suspicious | `Hex Strings` / `Base64 Strings` | olevba が検知。フォームの UI 設定バイナリに由来する可能性が高い |
| Suspicious | VBA Stomping | P-code と VBA ソースの差異を検知。旧バージョンの Excel で保存した可能性あり |
| 認証情報 | `DSN=ricdb;UID=ric;PWD=t6101` | Oracle DB の接続文字列がコード内にハードコードされている。**パスワードの管理に注意** |
