# ExMenu 仕様書

> **ファイル種別**: .xlsm（マクロ付き）
> **用途**: 照射管理システムの業務アプリケーション起動メニュー。最大30件のExcelアプリを登録・管理し、ボタンクリックで起動する
> **VBA プロジェクトサイズ**: 10モジュール（ThisWorkbook, Sheet3/Sheet5, メニュー選択実行, 共通変数と開始処理, SQL_Execution, ファイル登録_取込, お知らせメッセージ, Debug用, SikakuKakuninn[UserForm]）
> **外部連携ファイル**: `ExAprReadPath.txt`（サーバーパス設定）、`ExMenuからのお知らせ.txt`（お知らせ）

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
ExMenu.xlsm
├── シート
│   ├── メニュー  （アプリ起動メニュー画面 - ボタン30個）
│   └── AprList   （登録アプリ一覧・管理テーブル）
└── VBA モジュール
    ├── ThisWorkbook.cls        （Workbook_Open / BeforeClose / BeforeSave）
    ├── Sheet3.cls              （空）
    ├── Sheet5.cls              （空）
    ├── 共通変数と開始処理.bas  （Public変数定義・初期設定・MenuGetPath）
    ├── メニュー選択実行.bas    （選択1〜30・二重起動チェックと起動）
    ├── SQL_Execution.bas       （ADO/ODBC DB接続・SQL実行基盤）
    ├── ファイル登録_取込.bas   （メニュー登録Main・メニュー取込・メニュー登録）
    ├── お知らせメッセージ.bas  （メッセージ・お知らせ関数）
    ├── Debug用.bas             （AllCls・DellMenu）
    └── SikakuKakuninn.frm     （資格確認UserForm）
└── 外部連携ファイル
    ├── ExAprReadPath.txt       （サーバー元パス・コピー先パスの設定）
    └── ExMenuからのお知らせ.txt（運用通知テキスト）
```

---

## 2. シート詳細

### 2.1 メニュー

**目的**: 業務アプリ起動のメイン画面。30個のメニューボタン（Menu1〜Menu30）が配置され、クリックするとアプリが起動する。ヘッダー・グリッド・スクロールバー・タブはすべて非表示にして専用メニュー画面として機能する。

#### レイアウト構造

| セル/オブジェクト | 内容 |
|---|---|
| A1 | `Debug`（デバッグ用。空=通常動作） |
| Menu1〜Menu30 | メニューボタン（Shapeオブジェクト）。`AprList` の登録内容が動的にラベルと OnAction に設定される |

ボタンのラベルが `-----` の場合: サーバー未接続PCまたはファイル未登録（クリック不可ではないが起動は失敗する）

---

### 2.2 AprList

**目的**: 起動アプリの登録・管理テーブル。DB（ExSEIHINJ）から読み込んだ情報（I〜N列）と管理者が直接編集する情報（C〜H列）を比較し、変更検出を行う。

#### レイアウト構造

| セル範囲 | 名前付き範囲 | 内容 |
|---|---|---|
| B2 | — | ラベル「ﾌｧｲﾙ一覧」 |
| F2 | `AprSuu` | 登録ファイル数（30） |
| B3:O3 | — | 列ヘッダー行 |
| B4:B33 | — | メニューNo（1〜30） |
| C4:C33 | `PGList`の先頭列 | Path名（アプリの格納フォルダパス）（50B以内） |
| D4:D33 | — | ファイル名（50B以内） |
| E4:E33 | — | メニュー表示文字（100B以内） |
| F4:F33 | — | ソフト種別（20B以内） |
| G4:G33 | — | 登録/非表示日（12B以内） |
| H4:H33 | — | 登録者（20B以内） |
| I4:I33 | — | DB元 Path名（`メニュー取込` で書き込み） |
| J4:J33 | — | DB元 ファイル名 |
| K4:K33 | — | DB元 メニュー表示文字 |
| L4:L33 | — | DB元 ソフト種別 |
| M4:M33 | — | DB元 登録日 |
| N4:N33 | — | DB元 登録者 |
| O4:O33 | — | 変更有無フラグ（True/False。C〜H列がDB値と異なる場合 False） |

※ AprList シートはデータが文字化けしている場合がある（Shift-JIS⇔UTF-8の混在によるOpenPyXL読み取り問題）

---

## 3. 名前付き範囲一覧

| 名前 | 参照先 | 説明 |
|---|---|---|
| `AprSuu` | AprList!$F$2 | 登録可能アプリ数（30） |
| `Debug` | メニュー!$A$1 | デバッグモードフラグ（空=通常、`*`=非表示モード） |
| `PGList` | AprList!$C$4:$N$33 | アプリ一覧データ範囲（VBAはRow/Column参照でCell操作） |

---

## 4. 数式一覧

数式セルなし（すべての計算・表示はVBAで処理）。

---

## 5. ボタン・マクロ対応

### AprListシート

| シート | ボタンラベル | 割り当てマクロ | 動作概要 |
|---|---|---|---|
| AprList | ﾌｧｲﾙ登録 | `メニュー登録Main` | 変更ありの行をDBのExSEIHINJテーブルに登録・更新する |
| AprList | メニュー画面 | `BackMenu` | メニューシートに戻る |

### メニューシート（動的設定）

| シート | ボタン名 | 割り当てマクロ | 動作概要 |
|---|---|---|---|
| メニュー | Menu1 | `選択1` | AprListのNo.1のファイルを起動 |
| メニュー | Menu2〜Menu30 | `選択2`〜`選択30` | 同様（1〜30） |

---

## 6. VBA モジュール仕様

### 6.1 ThisWorkbook.cls

#### `Workbook_Open()`

**処理概要**: ブック起動時に `初期設定` を呼び出す。

#### `Workbook_BeforeClose(Cancel As Boolean)`

**処理概要**: 閉じる際にToolbarを元に戻し、数式バーを表示に戻す。

```vba
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    With Application
        .CommandBars("Standard").Visible = True
        .CommandBars("Formatting").Visible = True
        .DisplayFormulaBar = True
    End With
    ThisWorkbook.Saved = True
End Sub
```

---

### 6.2 共通変数と開始処理.bas

#### 共通変数

```vba
Public mpAprD(30, 2)    '(メニューNo, 0=Path/1=ファイル名/2=表示文字)
Public mpSyaName As String  '社員名（資格確認後に設定）
```

#### `初期設定()`

**処理概要**: 起動時の全初期化処理。バージョン確認→メニュー取込→ファイル同期→メニュー画面構築→お知らせ表示。

**処理フロー**:
1. ウィンドウ最大化、AprListシート保護設定
2. `MenuGetPath` でメニューのバージョン確認（元ファイルと本ファイルの更新日時比較）
3. バージョン不一致の場合: ユーザーに継続確認（Noの場合ブックを閉じる）
4. `メニュー取込` でDB（ExSEIHINJ）からAprListシートへデータ読み込み
5. AprListの各行（最大30件）を確認:
   - PC名先頭3文字が「RIC」以外かつサーバーパス以外: 表示文字を `-----` に
   - ファイルが見つからない場合: 表示文字を `-----` に
   - 元ファイルと本ファイルの更新日時が異なる場合: `FileCopy` でサーバーからローカルにコピー
6. メニューシートでグリッド・スクロールバー等を非表示
7. メニューボタン（Menu1〜Menu30）にラベルと OnAction（`選択1`〜`選択30`）を設定
8. `メッセージ` でお知らせを表示

```vba
If myPcName <> "RIC" And mpAprD(i, 0) <> "C:\ラジエ工業\ExRicSys\" Then
    mpAprD(i, 2) = "-----"  '非RICサーバー接続PCでは利用不可表示
```

#### `MenuGetPath() As Boolean`

**処理概要**: `ExAprReadPath.txt` から元パス・先パスを読み込み、`ExMenu.xlsm` の更新日時を比較してバージョン一致を確認する。

```vba
Function MenuGetPath() As Boolean
    Open ThisWorkbook.Path & "\" & "ExAprReadPath.txt" For Input As #myFno
        Input #myFno, mpAprMotoPath, mpAprSakiPath
    Close #myFno
    If FileDateTime(mpAprMotoPath & ThisWorkbook.Name) <> FileDateTime(mpAprSakiPath & ThisWorkbook.Name) Then
        MenuGetPath = False  'バージョン違い
    Else
        MenuGetPath = True
    End If
End Function
```

---

### 6.3 メニュー選択実行.bas

#### `選択1()` 〜 `選択30()`

**処理概要**: 各メニューボタンから呼び出されるスタブ。番号を引数に `二重起動チェックと起動` を呼ぶ。

#### `二重起動チェックと起動(ByVal myMenuNo)`

**処理概要**: 選択されたメニュー番号のファイルが既に開かれていないか確認し、未起動の場合は `Workbooks.Open` で起動する。

**処理フロー**:
1. AprListの `PGList` 範囲からファイル名一覧（`mpAprD(*,1)`）を再読み込み
2. 開いている全ブックに対し、ファイル名が一致するものがあれば「既に開いています」メッセージを表示して該当ブックをアクティブ化して終了
3. 一致しない場合: `Workbooks.Open ThisWorkbook.Path & "\" & mpAprD(myMenuNo, 1)` でファイルを起動
4. ファイル未見つかりエラー（1004）はサイレント無視

```vba
Workbooks.Open ThisWorkbook.Path & "\" & mpAprD(myMenuNo, 1)
```

---

### 6.4 SQL_Execution.bas

**処理概要**: ADO/ODBC DB接続・SQL実行の共通基盤（他ファイルと同構造）。

接続文字列はハードコード: `DSN=ricdb;UID=ric;PWD=t6101`

---

### 6.5 ファイル登録_取込.bas

#### `メニュー取込()` （呼出元: `初期設定`）

**処理概要**: DBテーブル `ExSeihinj`（KAISYACD>'9000'）からアプリ一覧を取得し、AprListシートのI列〜N列（`PGList`の列+6〜列+11）に書き込む。

```sql
SELECT TO_NUMBER(SUBSTR(KAISYACD,2)), FOLDER, FILENAME, KAIBIKOU,
       HIKITORI, TOUDATE, TOUNAME,
       FOLDER, FILENAME, KAIBIKOU, HIKITORI, TOUDATE, TOUNAME
FROM ExSeihinj
WHERE KAISYACD>'9000' ORDER BY KAISYACD
```

#### `メニュー登録Main()` （呼出元: 「ﾌｧｲﾙ登録」ボタン）

**処理概要**: 変更ありの行を検出してDBに登録・更新する。

**処理フロー**:
1. 確認ダイアログ（Noで中止）
2. 社員名が未設定の場合: `SikakuKakuninn`（資格確認UserForm）を表示
3. AprListの各行を走査（`PGList` 範囲）:
   - `変更有無`（O列）= True の行はスキップ
   - Path名またはファイル名が空: メニューから消去確認ダイアログ → Yesなら削除してDB更新
   - ファイルが存在しない場合: エラーメッセージで中止
   - 正常: `メニュー登録` でDB更新
4. 更新件数を表示。更新あれば再起動確認

#### `メニュー登録(myRow)` （呼出元: `メニュー登録Main`）

**処理概要**: 1行分のデータを `SQL_INSERT_UPDATE` で `ExSEIHINJ` テーブルにINSERT/UPDATEする。

| DBカラム | 内容 | 上限 |
|---|---|---|
| `KAISYACD` | コード（9001〜9030: メニューNo+9000） | — |
| `FOLDER` | アプリの格納フォルダパス | 50B |
| `FILENAME` | ファイル名 | 50B |
| `KAIBIKOU` | メニュー表示文字 | 100B |
| `HIKITORI` | ソフト種別 | 20B |
| `TOUDATE` | 登録日（Excelシリアル整数） | — |
| `TOUNAME` | 登録者（社員名） | — |

#### `BackMenu()` / `AddSoft()`

**処理概要**: メニューシートへ戻る / AprListのC4にフォーカス移動。

---

### 6.6 お知らせメッセージ.bas

#### `メッセージ()`

**処理概要**: `ExAprReadPath.txt` からサーバーパスを読み込み、`お知らせ` 関数を呼ぶ。

#### `お知らせ(myAprMotoPath, myAprSakiPath) As Boolean`

**処理概要**: サーバー元パスの `ExMenuからのお知らせ.txt` と、ローカルコピーを比較し、差異があればメッセージを表示する。

**処理フロー**:
1. 元ファイルの更新日時が10日以上前の場合: スキップ（`お知らせ = True` で正常終了）
2. 元ファイルを読み込み（行ごとに配列格納）
3. ローカルコピーを読み込み
4. 行数または内容が異なる場合: メッセージを表示し、「次回もお知らせ?」→ Noならローカルコピーを更新して差分を消去
5. ファイル未存在時: 空ファイルを自動作成して続行

---

### 6.7 SikakuKakuninn.frm（UserForm）

**処理概要**: ファイル登録操作に必要な資格（顧客・製品登録資格）を確認する認証フォーム。

| コントロール | 種別 | 動作 |
|---|---|---|
| `SyainnNo` | テキストボックス | 社員番号入力 |
| `PassWord` | テキストボックス | パスワード入力 |
| `Kakuninn` | コマンドボタン | 「確認」- DB照合して資格チェック |
| `CommandButton1` | コマンドボタン | 「キャンセル」 |

```sql
SELECT TRIM(shaname) FROM SHAINMST
WHERE shano='<社員番号>' AND shask='<パスワード>'
AND hshika='1' AND (cshika='2' OR cshika='3')
```

資格条件: `hshika='1'`（有効）かつ `cshika='2'` または `'3'`（顧客または製品登録資格者）

---

### 6.8 Debug用.bas

#### `AllCls()` / `DellMenu()`

**処理概要**: デバッグ用。メニューボタンのラベルをクリア、OnAction を削除する（開発・テスト時に使用）。

---

## 7. DB 接続・外部連携

### ODBC 接続設定

| DSN 名 | UID | PWD | 用途 |
|---|---|---|---|
| `ricdb` | `ric` | `t6101` | アプリ一覧・社員マスタの読み書き |

### 参照テーブル一覧

| テーブル名 | 主な用途 | 主要カラム |
|---|---|---|
| `ExSeihinj` / `ExSEIHINJ` | アプリ登録一覧 | `KAISYACD`（9001〜9030）, `FOLDER`, `FILENAME`, `KAIBIKOU`（表示名）, `HIKITORI`（種別）, `TOUDATE`, `TOUNAME` |
| `SHAINMST` | 社員マスタ（資格確認用） | `shano`（社員番号）, `shask`（パスワード）, `hshika`（有効フラグ）, `cshika`（資格種別）, `shaname`（社員名） |

### 外部ファイル連携

| ファイル名 | 格納場所 | 用途 |
|---|---|---|
| `ExAprReadPath.txt` | `ThisWorkbook.Path\` | 元サーバーパス・コピー先パスを CSV 形式で格納 |
| `ExMenuからのお知らせ.txt` | 元サーバーパス / コピー先パス | 運用通知テキスト。更新差分があれば起動時に表示 |
| 各アプリ（.xlsm等） | `ThisWorkbook.Path\` | 起動対象のアプリファイル（初期設定でサーバーから自動コピー） |

### 主要 SQL 文

```sql
-- アプリ一覧読み込み（メニュー取込）
SELECT TO_NUMBER(SUBSTR(KAISYACD,2)), FOLDER, FILENAME, KAIBIKOU,
       HIKITORI, TOUDATE, TOUNAME,
       FOLDER, FILENAME, KAIBIKOU, HIKITORI, TOUDATE, TOUNAME
FROM ExSeihinj
WHERE KAISYACD>'9000' ORDER BY KAISYACD

-- アプリ登録（メニュー登録）
SELECT COUNT(*) FROM ExSEIHINJ WHERE KAISYACD='<9001〜9030>'
INSERT INTO ExSEIHINJ (KAISYACD, FOLDER, FILENAME, KAIBIKOU, HIKITORI, TOUDATE, TOUNAME)
  VALUES('<cd>', '<path>', '<file>', '<label>', '<type>', '<date>', '<name>')
UPDATE ExSEIHINJ SET FOLDER='<path>', ... WHERE KAISYACD='<cd>'

-- 資格確認（SikakuKakuninn）
SELECT TRIM(shaname) FROM SHAINMST
WHERE shano='<社員番号>' AND shask='<パスワード>'
AND hshika='1' AND (cshika='2' OR cshika='3')
```

---

## 8. データフロー

```
【起動フロー】
  Workbook_Open → 初期設定()
       ↓
  MenuGetPath(): ExAprReadPath.txt 読み込み
       ↓ バージョン確認（ExMenu.xlsmの更新日時比較）
       ↓
  メニュー取込(): DB(ExSeihinj) → AprList I〜N列に書き込み
       ↓
  AprList C〜H列を走査（各アプリ最大30件）:
    - 非RICサーバー: 表示文字 = "-----"
    - ファイル未存在: 表示文字 = "-----"
    - サーバーとローカルの更新日時が異なる: FileCopy でコピー
       ↓
  メニュー画面設定（グリッド/スクロールバー等を非表示）
       ↓
  Menu1〜Menu30 に表示文字・OnAction を設定
       ↓
  メッセージ(): ExMenuからのお知らせ.txt の差分確認・表示

【アプリ起動フロー】
  メニューボタン（Menu1〜Menu30）クリック
       ↓ 選択N()
  二重起動チェックと起動(N):
    既に開いていればアクティブ化
       ↓ 未起動の場合
  Workbooks.Open ThisWorkbook.Path & "\" & <ファイル名>

【アプリ登録フロー（AprListシート）】
  管理者がAprListのC〜H列を直接編集
       ↓
  「ﾌｧｲﾙ登録」ボタン → メニュー登録Main()
       ↓
  SikakuKakuninn（資格確認UserForm）表示
       ↓ DB(SHAINMST) で社員番号・パスワード・資格確認
       ↓
  変更ありの行を走査
       ↓
  メニュー登録(): SQL_INSERT_UPDATE → DB(ExSEIHINJ) に更新
       ↓
  更新完了 → 初期設定()を再実行してメニュー再構築
```

---

## 9. セキュリティ注意事項

| 種別 | キーワード | 内容 |
|---|---|---|
| AutoExec | `Workbook_Open` | ブックオープン時に自動実行。DB接続・ファイルコピー・外部ファイル読み込みが発生 |
| AutoExec | `Workbook_BeforeClose` | ブック終了時にツールバーを元に戻す |
| AutoExec | `CommandButton1_Click` | UserFormのボタンクリックイベント |
| Suspicious | `Open` | `ExAprReadPath.txt`・`ExMenuからのお知らせ.txt` の読み書きに使用 |
| Suspicious | `Write` | `ExMenuからのお知らせ.txt` の更新時に使用（ユーザー操作に基づく） |
| Suspicious | `Output` | テキストファイル書き込みモードで使用 |
| Suspicious | `FileCopy` | サーバーからローカルへのアプリ自動コピーに使用 |
| Suspicious | `Call` | 各サブルーチン呼び出しで使用 |
| Suspicious | `Chr` | `Chr(13)` による改行文字生成（メッセージ整形） |
| Suspicious | Hex Strings / Base64 Strings | VBAバイナリ内のエンコード。実際のエンコード処理なし |
| IOC | `163.59.144.156` | 旧サーバーIPアドレス（コメントアウト済み）。現在は `C:\ラジエ工業\ExRicSys\` を使用 |
| 注意 | DB認証情報 | 接続文字列にUID/PWDをハードコード: `DSN=ricdb;UID=ric;PWD=t6101` |
| 注意 | ファイル自動コピー | 起動時にサーバーから最大30ファイルを自動でローカルにコピーする動作あり |
| 注意 | 社員パスワード | `SHAINMST.shask` カラムにパスワードをSQL直結で照合。ハッシュ化等の対策は不明 |
