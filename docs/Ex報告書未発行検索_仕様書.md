# Ex報告書未発行検索 仕様書

> **ファイル種別**: .xlsm（マクロ付き）
> **用途**: 照射管理DBから報告書が未発行・未送信の在庫品および出荷品を検索して一覧表示し、FaxフラグやロットNo・表示設定を変更して `ExKeikakuX` テーブルに登録する管理ツール
> **VBA プロジェクトサイズ**: 約893KB
> **外部連携**: DSN=ricdb（Oracle DB）、DBリンク `rich.zaikor@ricdb`、`rich.syoujr1@ricdb`、`rich.syoujr2@ricdb`

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
Ex報告書未発行検索.xlsm
├── シート
│   ├── 報告書一覧   （検索条件入力・報告書未発行品の一覧表示・Fax/表示設定変更、最大10000行×28列）
│   ├── 設定         （検索SQL条件文字列・FaxリストSQL・発行テーブル格納、最大2655行×20列）
│   └── Sheet1       （受付番号リスト参照用シート、25行×6列）
│
├── VBA モジュール
│   ├── ThisWorkbook.cls          （Workbook_Open / Workbook_BeforeClose）
│   ├── Sheet1.cls                （右クリックイベント：Fax/表示設定変更、Worksheet_Change）
│   ├── Sheet2.cls / Sheet3.cls   （空）
│   ├── データ表示_変更登録.bas   （報告書データ表示・Fax設定・最終完了日・変更登録・未登録検出処理）
│   ├── SQL_Execution.bas         （ODBC接続・SQL実行共通処理）
│   ├── スタート処理.bas          （起動時初期化・古いデータ削除・未登録検出）
│   ├── 共通変数.bas              （mpSaveFlg 変数定義）
│   ├── オートフィルター.bas      （オートフィルター設定/解除）
│   ├── 画面クリア.bas            （画面表示クリア・表示設定）
│   ├── 終了処理.bas              （ブック終了処理）
│   ├── ソート.bas                （検索結果のソート）
│   ├── 印刷範囲.bas              （印刷範囲設定）
│   ├── DB不要データ削除.bas       （ExKeikakuXの古いデータ削除）
│   ├── 高さの自動調整.bas        （行高さの自動調整）
│   ├── 画面拡大.bas              （ズーム設定）
│   └── ロギング記録.bas          （ログ書込処理 ※全コメントアウト済） 
│
└── 外部DB連携
    ├── DSN=ricdb（Oracle DB）
    ├── rich.zaikor@ricdb（DBリンク：在庫履歴）
    ├── rich.syoujr1@ricdb（DBリンク：1号機照射履歴）
    └── rich.syoujr2@ricdb（DBリンク：2,3号機照射履歴）
```

---

## 2. シート詳細

### 2.1 報告書一覧

**目的**: メインの操作画面。検索条件を入力して「検索」ボタンを押すと、DBから報告書未発行・未送信品を抽出して一覧表示する。右クリックで各行のFaxフラグ・非表示設定を変更できる。

#### レイアウト構造（ヘッダー・条件入力領域）

| セル範囲 | 名前付き範囲 | 内容 |
|---|---|---|
| A1 | `Debug` | デバッグフラグ（空白=通常動作） |
| B1 | — | 固定値「20181225」（最古受付番号基準） |
| B4 | — | ラベル「検索開始受付番号」 |
| C4 | — | ラベル「会社ｺｰﾄﾞ」 |
| D4 | — | ラベル「最終出荷日」 |
| E4 | — | ラベル「合計件数」 |
| F4 | — | ラベル「在庫品」 |
| G4 | — | ラベル「出荷品」 |
| H4 | — | ラベル「表示設定」 |
| I4 | — | ラベル「Fax送信」 |
| W4 | — | ラベル「実績件数」 |
| X4 | `JisekiSu` | 実績件数（VBAが設定） |
| B5 | `Uno` | 検索開始受付番号入力欄 |
| C5 | `KaisyaCD` | 会社コード入力欄 |
| D5 | `Syukkabi` | 最終出荷日入力欄 |
| E5 | `MiKensu` | 合計件数（VBAが設定） |
| F5 | — | 在庫品件数（数式） |
| G5 | — | 出荷品件数（数式） |
| H5 | `HyoujiFlg` | 表示設定（非表示除く/全て/非表示のみ） |
| I5 | `FaxHyouji` | Fax送信設定（未送信/送信済/空=全て） |
| B8 | — | ラベル「計算式行」（数式テンプレート行） |
| M8 | `KanDateC` | 最終完了日計算式テンプレート |
| P8 | `UmuSiki` | 変更有無判定式テンプレート |
| Q8 | `HenkouSu` | Fax/表示変更件数合計 |
| R8 | `RotSiki` | ロット番号変更有無式テンプレート |
| S8 | `HenkouSuRot` | ロット番号変更件数合計 |

#### レイアウト構造（データ表示領域）

| セル範囲 | 名前付き範囲 | 内容 |
|---|---|---|
| B10:R10 | — | 列ヘッダー行（受付番号/顧客ｺｰﾄﾞ/顧客名/製品ｺｰﾄﾞ/製品名/ロット番号/在庫品・出荷品/報告書発行種別/Fax/報告書発行フラグ/最終完了日/表示設定/表示変更前/Fax変更有無/変更前ロット番号/ロット番号変更有無） |
| W10:X10 | — | 列ヘッダー（受付番号・最終完了日：集計用） |
| B11:X10011 | `ItiranTB` | 報告書一覧データ表示領域（VBAが書き込み） |
| W11:X10011 | `JissekiTB` | 最終完了日集計テーブル |

（行番号 9 は空白。行番号 10012〜10014 は空白）

#### 列詳細（データ行 B11以降）

| 列 | ヘッダー | 内容 |
|---|---|---|
| B | 受付番号 | DB `uno` |
| C | 顧客ｺｰﾄﾞ | DB `kaisyacd` |
| D | 顧客名 | DB `kainame` |
| E | 製品ｺｰﾄﾞ | DB `sehncd` |
| F | 製品名 | DB `seiname` |
| G | ロット番号 | DB `rotno2`（変更可） |
| H | 在庫品・出荷品 | 在庫品 / 出荷品 の区別 |
| I | 報告書発行種別 | DB `housyube` |
| J | Fax（右クリックで変更） | 未/済（右クリックで切替） |
| K | Fax（表示コピー） | J列のコピー |
| L | 報告書発行フラグ | DB `syinji`（0=未発行/1=発行済） |
| M | 最終完了日 | 数式式により計算（mmdd形式） |
| N | 表示設定 | 非/空（右クリックで切替） |
| O | 表示変更前 | 変更前の値 |
| P | Fax表示設定変更有無 | 数式（1=変更あり） |
| Q | 変更前ロット番号 | 変更前の値 |
| R | ロット番号変更有無 | 数式（1=変更あり） |
| W | 受付番号（集計用） | `最終完了日` 関数用 |
| X | 最終完了日（集計用） | DB `max(smmdd)` |

---

### 2.2 設定

**目的**: 検索に使用するSQL条件文字列・報告書発行テーブル・Faxリスト等を格納する設定シート。

#### レイアウト構造（主要）

| セル範囲 | 名前付き範囲 | 内容 |
|---|---|---|
| B2 | — | ラベル「検索テーブル」 |
| C2 | — | ラベル「最終出荷日」 |
| C3 | `SyukkaDay` | 出荷日をYYYYMMDD形式に変換する数式 |
| B3〜B7 | `HouSybetu` | 報告書種別リスト |
| B10:C12 | `FaxSettei` | Faxフラグ一覧（未送信/送信済/全て） |
| C10 | — | SQL条件「AND e.houfax='未'」 |
| C11 | — | SQL条件「AND e.houfax='済'」 |
| D10 | `SQLfax` | Fax検索条件（数式） |
| B15〜B17 | `HyoujiHou` | 表示方法リスト（未発行/発行済/全て） |
| B20:C22 | `HyoujiLst` | 表示設定リスト（非表示除く/全て/非表示のみ） |
| C20 | — | SQL条件「AND e.hyouji IS NULL」 |
| C22 | — | SQL条件「AND e.hyouji='非'」 |
| D20 | `SQLHyouji` | 表示検索条件（数式） |
| B25:B26 | `FaxList` | Faxリスト |
| C25 | `HyoujiList` | 表示設定リスト |
| J5:Q50000 | `HakkouTB` | 報告書発行テーブル（VBAが書き込み） |

---

### 2.3 Sheet1

**目的**: 受付番号リスト（参照用）。25行×6列。B列とD列に受付番号が格納されている。

---

## 3. 名前付き範囲一覧

| 名前 | 参照先 | 説明 |
|---|---|---|
| `Debug` | 報告書一覧!$A$1 | デバッグフラグ |
| `FaxHyouji` | 報告書一覧!$I$5 | Fax送信表示設定（未送信/送信済/全て） |
| `FaxList` | 設定!$B$25:$B$26 | Faxリスト（未/済/不） |
| `FaxSettei` | 設定!$B$10:$B$12 | Fax設定リスト（表示用） |
| `HakkouTB` | 設定!$J$5:$Q$50000 | 報告書発行テーブル（未登録検出処理で使用） |
| `HenkouSu` | 報告書一覧!$Q$8 | Fax/表示変更件数の合計 |
| `HenkouSuRot` | 報告書一覧!$S$8 | ロット番号変更件数の合計 |
| `HouSybetu` | 設定!$B$3:$B$7 | 報告書発行種別リスト |
| `HyoujiFlg` | 報告書一覧!$H$5 | 表示設定フラグ入力欄 |
| `HyoujiHou` | 設定!$B$15:$B$17 | 表示方法リスト |
| `HyoujiList` | 設定!$C$25 | 表示設定リスト参照 |
| `HyoujiLst` | 設定!$B$20:$B$22 | 表示設定一覧 |
| `ItiranTB` | 報告書一覧!$B$11:$X$10011 | 報告書一覧データ表示領域 |
| `JisekiSu` | 報告書一覧!$X$4 | 実績件数（最終完了日件数） |
| `JissekiTB` | 報告書一覧!$W$11:$X$10011 | 最終完了日集計テーブル |
| `KaisyaCD` | 報告書一覧!$C$5 | 会社コード検索条件入力欄 |
| `KanDateC` | 報告書一覧!$M$8 | 最終完了日計算式テンプレート（行8） |
| `MiKensu` | 報告書一覧!$E$5 | 合計件数 |
| `RotSiki` | 報告書一覧!$R$8 | ロット番号変更有無判定式テンプレート |
| `SQLfax` | 設定!$D$10 | Fax検索SQL条件文字列（数式） |
| `SQLHyouji` | 設定!$D$20 | 表示設定SQL条件文字列（数式） |
| `Syukkabi` | 報告書一覧!$D$5 | 最終出荷日検索条件入力欄 |
| `SyukkaDay` | 設定!$C$3 | 出荷日のYYYYMMDD形式（数式） |
| `UmuSiki` | 報告書一覧!$P$8 | Fax/表示設定変更有無判定式テンプレート |
| `Uno` | 報告書一覧!$B$5 | 検索開始受付番号入力欄 |

---

## 4. 数式一覧

### 報告書一覧シート

| セル | 数式 | 説明 |
|---|---|---|
| F5 | `=COUNTIF(H10:H10011,F4)` | 在庫品件数（「在庫品」の件数） |
| G5 | `=COUNTIF(H11:H10012,G4)` | 出荷品件数（「出荷品」の件数） |
| M8 | `=IF(ISERROR(VLOOKUP(B8,JissekiTB,2,FALSE)),"",VLOOKUP(B8,JissekiTB,2,FALSE))` | 最終完了日取得（受付番号で集計テーブルを参照）※テンプレート |
| P8 | `=IF(OR(J8<>K8,N8<>O8),1,"")` | Fax/表示設定変更有無（変更があれば1）※テンプレート |
| Q8 | `=SUM(P11:P10011)` | Fax/表示変更件数合計 |
| R8 | `=IF(G8<>Q8,1,"")` | ロット番号変更有無※テンプレート |
| S8 | `=SUM(R11:R10011)` | ロット番号変更件数合計 |

### 設定シート

| セル | 数式 | 説明 |
|---|---|---|
| C3 | `=IF(Syukkabi="","",YEAR(Syukkabi)&RIGHT("00"&MONTH(Syukkabi),2)&RIGHT("00"&DAY(Syukkabi),2))` | 出荷日入力値をYYYYMMDD形式に変換 |
| D10 | `=IF(FaxHyouji="","",VLOOKUP(FaxHyouji,設定!B10:C12,2,FALSE))` | FaxフラグのSQL条件文を動的に生成 |
| D20 | `=IF(HyoujiFlg=設定!B21,"",VLOOKUP(HyoujiFlg,設定!$B$20:$C$22,2,FALSE))` | 表示設定のSQL条件文を動的に生成 |

---

## 5. ボタン・マクロ対応

| シート | ボタンラベル | 割り当てマクロ | 動作概要 |
|---|---|---|---|
| 報告書一覧 | 検索 | `報告書データ表示` | 条件に合う報告書未発行品をDBから検索して一覧表示 |
| 報告書一覧 | 変更登録 | `変更登録` | Faxフラグ・表示設定・ロットNoの変更をDBに登録 |
| 報告書一覧 | 印刷 | `印刷行` | 印刷行数を指定して印刷範囲を設定 |
| 報告書一覧 | 終了 | `Bookを閉じる` | ブックを閉じる（未登録確認あり） |

**右クリック操作（Sheet1）**:
- J列（Fax欄）: 「未」→「済」/「済」→「未」の切替
- N列（表示設定欄）: 「非」/ 空（表示）の切替

---

## 6. VBA モジュール仕様

### 6.1 ThisWorkbook

#### `Workbook_Open()`

**処理概要**: ブック起動時に `Bookスタート` を呼び出す。

#### `Workbook_BeforeClose(Cancel As Boolean)`

**処理概要**: `mpSaveFlg = False`（変更登録未処理）の場合に確認ダイアログを表示する。「いいえ」を選択すると終了をキャンセルする。

---

### 6.2 Sheet1

#### `Worksheet_BeforeRightClick(Target, Cancel)`

**処理概要**: データ行（11行以降）の右クリックイベントを処理する。
- J列（Fax欄）: 「未」かつ報告書発行済（L列=1）なら「済」に変更。「済」なら「未」に戻す。
- N列（表示設定欄）: 空なら「非」に設定、「非」なら空に戻す。

`Debug` フラグが空の場合、コンテキストメニューは非表示にする。

---

### 6.3 データ表示_変更登録

#### `報告書データ表示()`

**処理概要**: 検索条件を検証後、`データ読込` → `最終完了日` → `Sort` → 印刷範囲設定の順で処理を実行する。

**処理フロー**:
1. `Uno`（受付番号）の入力チェック（2014050000以上必須）
2. `ItiranTB` をクリア
3. `データ読込()` を呼び出してDBからデータ取得
4. `最終完了日()` を呼び出して完了日を付加
5. `Sort()` でソート（在庫品/出荷品降順→顧客コード昇順→受付番号昇順）
6. 行高自動調整・オートフィルター設定・印刷範囲設定

#### `データ読込()`

**処理概要**: 4パターン（在庫未発行・出荷未発行・在庫発行済Fax未送信・出荷発行済Fax未送信）のSQLを `UNION ALL` で結合し、`ItiranTB` 領域に結果を書き込む。

**主要 SQL**:
```sql
-- 在庫品（報告書未発行）
SELECT TO_NUMBER(z.uno), z.kaisyacd, z.kainame, z.sehncd, RTRIM(s.seiname),
       RTRIM(z.rotno2), '在庫品', SJ.housyube, e.houfax, e.houfax,
       z.syinji, '', e.hyouji, e.hyouji, '', RTRIM(z.rotno2)
FROM zaiko z, sehmst s, ExKeikakuX e, ExSeihinJ SJ
WHERE TO_NUMBER(z.syouzusu) = TO_NUMBER(z.nyukasu)
  AND z.kaisyacd = s.kaisyacd AND z.sehncd = s.sehncd
  AND s.syouho <> '0'
  AND z.uno > '<検索開始受付番号>'
  AND z.uno = e.uno(+)
  AND z.kaisyacd = SJ.kaisyacd(+)
  AND z.syinji = '0'
  AND (SJ.housyube = '照射後Fax送信' OR SJ.housyube = '照射後報告書発行' OR SJ.housyube = '不要')
  [AND e.hyouji IS NULL]  -- 非表示除く時
  [AND e.houfax = '未']   -- 未送信のみ表示時

UNION ALL

-- 出荷品（報告書未発行）
SELECT ... FROM rich.zaikor@ricdb z, sehmst s, ExKeikakuX e, ExSeihinJ SJ
WHERE TO_NUMBER(z.syukasu) = TO_NUMBER(z.nyukasu) AND ...

UNION ALL

-- 在庫品（報告書発行済・Fax未送信）
... AND z.syinji = '1' AND e.houfax = '未' ...

UNION ALL

-- 出荷品（報告書発行済・Fax未送信）
... FROM rich.zaikor@ricdb z, ...
```

#### `Fax送信()`

**処理概要**: DBから取得後、Fax列（J列）が空の行について `housyube` をチェックし、Faxを含む種別なら「未」、含まない場合は「不」を設定する。

#### `最終完了日()`

**処理概要**: 1号機・2号機/3号機の照射記録（在庫・履歴）から最終完了日（照射yymmdd）を取得して `JissekiTB`（W/X列）に格納し、`KanDateC`（M列）の数式をコピーして値化する。

**主要 SQL**:
```sql
-- 1号機在庫の最終照射日
SELECT TO_NUMBER(uno), max(smmdd) FROM syouj1
WHERE dflg = '1' AND TO_NUMBER(syoiti) > 0
GROUP BY uno

UNION ALL

-- 1号機履歴の最終照射日
SELECT TO_NUMBER(uno), max(smmdd) FROM rich.syoujr1@ricdb
WHERE dflg = '1' AND Uno > '<検索開始受付番号>'
GROUP BY uno

UNION ALL

-- 2,3号機在庫の最終照射日
SELECT TO_NUMBER(uno), max(smmdd) FROM syouj2
WHERE syouflg = '1'
GROUP BY uno

UNION ALL

-- 2,3号機履歴の最終照射日
SELECT TO_NUMBER(uno), max(smmdd) FROM rich.syoujr2@ricdb
WHERE Uno > '<検索開始受付番号>'
GROUP BY uno
```

#### `変更登録()`

**処理概要**: `HenkouSu`（Fax/表示変更件数）と `HenkouSuRot`（ロット番号変更件数）に基づいてDBに変更を登録する。

**処理フロー**:
1. 「報告書発行フラグ=1（発行済）」の行を非表示に設定するか確認
2. `UmuSiki`（変更有無式）をP列、`RotSiki`（ロット番号変更有無式）をR列にコピー・値化
3. P列=1（Fax/表示変更あり）の行について `ExKeikakuX` を UPDATE
4. R列=1（ロット番号変更あり）の行について `zaiko`（在庫品）または `rich.zaikor@ricdb`（出荷品）を UPDATE

```sql
-- Fax/表示設定変更
UPDATE ExKeikakuX SET houfax='<Fax値>', hyouji='<表示設定>'
WHERE uno='<受付番号>'

-- ロット番号変更（在庫品）
UPDATE zaiko SET rotno2='<ロット番号>'
WHERE uno='<受付番号>'

-- ロット番号変更（出荷品）
UPDATE rich.zaikor@ricdb SET rotno2='<ロット番号>'
WHERE uno='<受付番号>'
```

#### `未登録検出在庫変更登録()`

**処理概要**: `ExKeikakuX` に登録されていない受付番号を検出して自動登録する。また `zaiko` と `ExKeikakuX` で会社/製品コードが異なる（変更された）レコードを更新する。

---

### 6.4 SQL_Execution

`Open_oraconDB`、`SQL_Exe`、`SQL_INSERT_UPDATE`、`SQL_Delete`、`Disp_Sheet`、`Set_Array` の各共通サブルーチン（他ファイルと同様の実装。接続文字列: `DSN=ricdb;UID=ric;PWD=t6101`）。

---

### 6.5 スタート処理

#### `Bookスタート()`

**処理概要**: 起動時の初期化処理。

**処理フロー**:
1. `mpSaveFlg = False` に設定
2. 「報告書一覧」シートのシート保護を解除
3. `画面消去()` で表示をクリア
4. `画面拡大設定()` でズームを自動設定
5. `古いデータ削除()` で `ExKeikakuX` の不要レコードを削除（50000件以上の場合）
6. `Uno` を「1か月前の年月0000」に自動設定
7. `未登録検出在庫変更登録()` で未登録受付番号を自動登録

---

### 6.6 DB不要データ削除

#### `古いデータ削除()`

**処理概要**: `ExKeikakuX` のレコード数が50000件以上の場合、`ZAIKOK_V`（在庫ビュー）に存在しない3年以上前のレコードを削除する。

```sql
SELECT COUNT(*) FROM ExkeikakuX WHERE uno > '1999010000'
-- 50000件以上の場合:
DELETE ExkeikakuX
WHERE NOT EXISTS (SELECT NULL FROM ZAIKOK_V WHERE ExkeikakuX.uno = ZAIKOK_V.uno)
  AND uno > '1999010000'
  AND uno < '<3年前の年月0000>'
```

---

### 6.7 ロギング記録（コメントアウト済）

**概要**: 変更登録・未登録検出時にログファイル（`D:\ExSysLogging\<yyyymm>H.Txt`）に書き込む処理が実装されていたが、現在はすべてコメントアウトされている。ログパスは `D:\ExSysLogging`。

---

## 7. DB 接続・外部連携

### ODBC 接続設定

| DSN 名 | UID | PWD | 用途 |
|---|---|---|---|
| `ricdb` | `ric` | `t6101` | 在庫・計画・製品マスタへのメインアクセス |

### DBリンク

| リンク名 | 説明 |
|---|---|
| `rich.zaikor@ricdb` | 在庫履歴テーブル（`zaikor`）へのDBリンク |
| `rich.syoujr1@ricdb` | 1号機照射履歴テーブル（`syoujr1`）へのDBリンク |
| `rich.syoujr2@ricdb` | 2,3号機照射履歴テーブル（`syoujr2`）へのDBリンク |

### 参照テーブル一覧

| テーブル名 | 主な用途 | 主要カラム |
|---|---|---|
| `zaiko` | 在庫 | `uno`, `kaisyacd`, `kainame`, `sehncd`, `rotno2`, `syinji`（0=未発行/1=発行済）, `syouzusu`, `nyukasu` |
| `zaikor`（DBリンク） | 在庫履歴 | `uno`, `kaisyacd`, `sehncd`, `rotno2`, `syinji`, `syukasu`, `nyukasu` |
| `sehmst` | 製品マスタ | `kaisyacd`, `sehncd`, `seiname`, `syouho`（報告書要否） |
| `ExKeikakuX` | 計画テーブル | `uno`, `houfax`（Faxフラグ）, `hyouji`（表示設定）, `housyube`（発行種別）, `kaisyacd`, `sehncd` |
| `ExSeihinJ` | 製品情報 | `kaisyacd`, `housyube`（報告書発行種別） |
| `syouj1` | 1号機照射記録（在庫） | `uno`, `smmdd`（照射日mmdd）, `dflg`, `syoiti` |
| `syoujr1`（DBリンク） | 1号機照射履歴 | `uno`, `smmdd`, `dflg` |
| `syouj2` | 2,3号機照射記録（在庫） | `uno`, `smmdd`, `syouflg` |
| `syoujr2`（DBリンク） | 2,3号機照射履歴 | `uno`, `smmdd` |
| `ZAIKOK_V` | 在庫ビュー（削除判定用） | `uno` |

---

## 8. データフロー

```
【起動フロー】
  Workbook_Open() → Bookスタート()
        ↓
  画面消去（ItiranTB クリア）
  古いデータ削除（ExKeikakuX の不要レコード削除）
  Uno = 1か月前の受付番号に自動設定
  未登録検出在庫変更登録(): DB(zaiko/zaikor) → ExKeikakuX に自動登録

【検索フロー】
  報告書一覧シート: 受付番号・会社コード・出荷日・表示設定・Fax条件 入力
        ↓ [検索]ボタン → 報告書データ表示()
  データ読込():
    DB(zaiko/zaikor/sehmst/ExKeikakuX/ExSeihinJ):
      4つのSQL を UNION ALL で結合 → ItiranTB に書き込み
    Fax送信(): J/K列の Fax フラグを自動設定
        ↓
  最終完了日():
    DB(syouj1/syoujr1/syouj2/syoujr2): UNION ALL で最終照射日を取得
    → JissekiTB(W/X列) に格納
    → KanDateC 数式をM列にコピー・値化
        ↓
  Sort(): H列(在庫品/出荷品)降順→C列(顧客CD)昇順→B列(受付番号)昇順でソート
  行高自動調整・オートフィルター・印刷範囲設定

【変更登録フロー】
  報告書一覧: J列(Fax)右クリックで未/済切替、N列(表示設定)右クリックで非/空切替
  G列(ロット番号)を直接編集で変更
        ↓ [変更登録]ボタン → 変更登録()
  発行済（L列=1）の行を非表示に設定するか確認
  UmuSiki/RotSikii をP/R列にコピー・値化
  P列=1: DB(ExKeikakuX) に houfax・hyouji を UPDATE
  R列=1: DB(zaiko または rich.zaikor@ricdb) に rotno2 を UPDATE
  mpSaveFlg = True
```

---

## 9. セキュリティ注意事項

| 種別 | キーワード | 内容 |
|---|---|---|
| AutoExec | `Workbook_Open` | 起動時に `Bookスタート()` が自動実行される |
| AutoExec | `Workbook_BeforeClose` | 変更未登録時は終了確認ダイアログを表示する |
| AutoExec | `Worksheet_Change` | セル変更イベントが設定されているが現在は `Exit Sub`（無効） |
| Suspicious | `Open` | ロギング処理でファイルを開く操作（現在はコメントアウト済） |
| Suspicious | `Print #` | ロギングファイルへの書き込み（現在はコメントアウト済） |
| Suspicious | `Call` | モジュール間の内部呼び出し |
| Suspicious | `MkDir` | ログフォルダ `D:\ExSysLogging` を作成する処理（コメントアウト済） |
| Suspicious | `Chr` | Chr(13)（改行）を使用したメッセージ生成 |
| Suspicious | Hex/Base64 Strings | xlsm 構造上の正常なバイナリデータ |

> **注意点**:
> 1. DB接続文字列（`DSN=ricdb;UID=ric;PWD=t6101`）がVBAコード内に平文でハードコーディングされています。
> 2. DBリンク `rich.zaikor@ricdb` 等により、メインDB以外のサーバーにも接続します。
> 3. ロギング処理（`D:\ExSysLogging` へのファイル書込）は現在コメントアウトされており無効ですが、コードは残存しています。
