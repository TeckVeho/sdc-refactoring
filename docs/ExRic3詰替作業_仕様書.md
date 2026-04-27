# ExRic3詰替作業 仕様書

> **ファイル種別**: .xlsm（マクロ付き）
> **用途**: ガンマ照射における積替え品（詰替え対象製品）の照射状況確認・積替え要否設定・管理
> **VBA プロジェクトサイズ**: 標準モジュール7、クラスモジュール4
> **外部連携**: Oracle DB（DSN=ricdb）、テーブル syoukj3/zaiko/ExKeikakuX/ExSeihinZ/tokumst/sehmst

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
ExRic3詰替作業.xlsm
├── シート
│   ├── 積替品     ← メイン画面（235行×15列）積替え品一覧表示
│   ├── 積替TB     ← 積替え製品テーブル（10000行×12列）製品マスター
│   └── WorkTB     ← DB取得ワーク（133行×9列）一時作業テーブル
└── VBA モジュール
    ├── ThisWorkbook.cls      ← Workbook_Open/BeforeClose
    ├── Sheet4.cls            ← （空）
    ├── Sheet5.cls            ← （空）
    ├── Sheet6.cls            ← （空）
    ├── SQL_Execution.bas     ← Oracle DB ADODB接続・SQL実行共通ルーチン
    ├── Ex積替え品抽出.bas    ← DBから積替え品データを取得・表示
    ├── Ex積替品表示.bas      ← 積替製品テーブルをDBから取得・表示
    ├── Ex詰替品更新.bas      ← 積替え要否フラグをDBに更新
    ├── Ex画面クリア.bas      ← 画面クリア・印刷範囲初期化
    ├── ExFunction.bas        ← 日付変換ユーティリティ
    ├── 印刷範囲.bas          ← 印刷範囲設定
    └── 終了処理.bas          ← ブックを閉じる処理
```

---

## 2. シート詳細

### 2.1 積替品（メイン画面）

**目的**: 積替え対象品の照射状況（未/中/済）・納期・受付番号等を一覧表示する操作画面。

#### レイアウト構造

| セル範囲 | 名前付き範囲 | 内容 |
|---|---|---|
| A1 | `DebugFlg` | デバッグフラグ |
| B3 | — | 「積替え品一覧」タイトル |
| F3 | `Innsatu` | 印刷範囲行数（初期値=32） |
| G3 | — | 「まで印刷 Max」 |
| I3 | — | 最大件数（`=MAX(B6:B235)`） |
| L3 | — | 現在日時（`=NOW()`） |
| M3 | — | バージョン表示（「Ex4」） |
| B5 | — | 「No」ヘッダー |
| D5 | — | 「最終積替え日」ヘッダー |
| E5 | — | 「照射状況」ヘッダー |
| F5 | — | 「線量計番号」ヘッダー |
| G5 | — | 「受付番号」ヘッダー |
| H5 | — | 「会社名」ヘッダー |
| I5 | — | 「納期」ヘッダー |
| J5 | — | 「出荷日」ヘッダー |
| K5 | — | 「パス数」ヘッダー |
| L5 | — | 「備考」ヘッダー |
| M5 | — | 「チェック」ヘッダー |
| B6〜M235 | `TumikaeTB` | 積替え品データ表示エリア（最大230行） |

#### 照射状況コード変換

| DBコード | 表示値 | 意味 |
|---|---|---|
| `i` | 未 | 未照射 |
| `0` | 未 | 未照射 |
| `C` | 未 | 未照射 |
| `1` | 中 | 照射中 |
| `2` | 済 | 照射済み |
| その他 | ?? | 不明 |

---

### 2.2 積替TB（積替え製品テーブル）

**目的**: 積替え対象かどうかを設定する製品マスターテーブル。DBから読み込んだ製品一覧に対し詰替え要否フラグを設定する。

#### レイアウト構造

| セル範囲 | 名前付き範囲 | 内容 |
|---|---|---|
| E1 | — | 「積み替え製品テーブル」タイトル |
| D3 | `SeiKennsuu` | 登録製品数（`=COUNTA(SeihinnSuu)`） |
| D4 | — | 操作ガイドメッセージ |
| D5 | — | 「会社コード」ヘッダー |
| E5 | — | 「製品コード」ヘッダー |
| F5 | — | 「会社コード＆製品コード」ヘッダー |
| G5 | — | 「会社名」ヘッダー |
| H5 | — | 「製品名」ヘッダー |
| I5 | — | 「詰替え要不要」ヘッダー（編集可） |
| J5 | — | `ExSeihinZ.kaisyacd` ヘッダー（元の値保存用） |
| D6〜J10000 | `TumikaeHinn` | 製品データエリア |
| D6〜D1005 | `SeihinnSuu` | 製品数カウント用列 |
| G6〜I1005 | `TumeTB` | 会社名・製品名・要否フラグ列 |

---

### 2.3 WorkTB（作業テーブル）

**目的**: DBからデータを一時取得するワークシート。`積替品` シートへの変換処理前に使用する。

| セル範囲 | 名前付き範囲 | 内容 |
|---|---|---|
| A1〜H201 | `Work` | ワークデータエリア |
| A1〜H250 | `Wtb` | 拡張ワークエリア |

DBから取得されるフィールド順序（列）: syoush_f, sykno, uno1, kainame1, nouki, syukkabi, pass, bikou1

---

## 3. 名前付き範囲一覧

| 名前 | 参照先 | 説明 |
|---|---|---|
| `DebugFlg` | 積替品!$A$1 | デバッグフラグ |
| `Innsatu` | 積替品!$F$3 | 印刷範囲行数設定 |
| `SeihinnSuu` | 積替TB!$D$6:$D$1005 | 製品カウント用列 |
| `SeiKennsuu` | 積替TB!$D$3 | 登録製品数 |
| `TumeTB` | 積替TB!$G$6:$I$1005 | 会社名・製品名・要否列 |
| `TumikaeHinn` | 積替TB!$D$6:$J$10000 | 製品テーブル全体 |
| `TumikaeTB` | 積替品!$B$6:$M$235 | 積替品表示エリア |
| `Work` | WorkTB!$A$1:$H$201 | ワークデータエリア |
| `Wtb` | WorkTB!$A$1:$H$250 | 拡張ワークエリア |

---

## 4. 数式一覧

### 積替品シート

| セル | 数式 | 説明 |
|---|---|---|
| I3 | `=MAX(B6:B235)` | 表示件数の最大値 |
| L3 | `=NOW()` | 現在日時 |

### 積替TBシート

| セル | 数式 | 説明 |
|---|---|---|
| D3 | `=COUNTA(SeihinnSuu)` | 登録製品総数 |
| I3 | `=IF(ISERROR(VLOOKUP(#REF!,TumiFlg,2,FALSE)),"",VLOOKUP(...))` | 積替フラグ参照（参照エラーあり、開発中の痕跡） |
| J3 | `=IF(ISERROR(VLOOKUP(#REF!,TumiFlg,2,FALSE)),"",VLOOKUP(...))` | 同上 |

---

## 5. ボタン・マクロ対応

| シート | ボタンラベル | 割り当てマクロ | 動作概要 |
|---|---|---|---|
| 積替品 | 現状の積替品の照射状況 | `TumikaeHinn` | DBから積替え対象品を取得し一覧表示・照射状況表示 |
| 積替品 | 積替する製品の登録 | `積替製品TB表示` | 積替TBシートに遷移し製品マスターをDB読込表示 |
| 積替品 | 画面を閉じる | `Bookを閉じる` | ブックを保存せずに閉じる |
| 積替TB | 戻る | `Modori` | 積替品シートに戻る |
| 積替TB | 登録内容更新 | `詰替品データ更新` | 積替TBシートの詰替え要否変更をDBに保存 |
| 積替TB | データ表示製品検索 | `積替製品TB表示` | 積替製品テーブルをDBから再取得 |

---

## 6. VBA モジュール仕様

### 6.1 ThisWorkbook.cls

#### `Workbook_Open()`

**処理概要**: ブック起動時に画面クリア後、DBから積替え品データを取得・表示する。

**処理フロー**:
1. ブック構造保護解除
2. ウィンドウ最大化
3. `積替品` シートを選択
4. `画面クリア3詰替` → 全エリアをクリアし印刷範囲初期化
5. `TumikaeHinn` → DBから積替え品データを取得・表示

#### `Workbook_BeforeClose()`

**処理概要**: アラート非表示で保存済みフラグをセットし、変更なしで閉じる。

---

### 6.2 Ex積替え品抽出.bas

#### `TumikaeHinn()`

**処理概要**: Oracle DBからRIC-3積替え品の照射状況データを取得し、加工して積替品シートに表示する。

**処理フロー**:
1. `Wtb` エリアをクリア
2. `syoukj3`, `zaiko`, `ExKeikakuX`, `ExSeihinZ` をJOINしてSQLを実行
3. 結果を `WorkTB` シートの `Work` エリアに格納
4. `DataKakou()` → データを加工して `積替品` シートに表示

#### `DataKakou(myDataN)`

**処理概要**: WorkTBのデータを加工（日付フォーマット変換・照射状況コード変換・文字列クリーン）し積替品シートに表示後、ソートを実行する。

**処理フロー**:
1. WorkTBシートからデータを配列に読み込み
2. 各行の加工:
   - 出荷日（syukkabi）をExcel日付値に変換し、最終積替え日 = 出荷日 − 1日に設定
   - 照射状況コード（i/0/C→未、1→中、2→済）に変換
   - 線量計番号を右4桁に切り出し
   - 受付番号を右4桁に切り出し
   - 会社名から「株式会社」を削除
   - 納期を mm/dd 形式に変換
3. 積替品シートの `TumikaeTB` エリアにデータを書き込み
4. `E6`（照射状況）、`D6`（積替日）、`F6`（線量計番号）キーでソート

---

### 6.3 Ex積替品表示.bas

#### `積替製品TB表示()`

**処理概要**: tokumst, sehmst, ExSeihinZ テーブルをJOINしてRIC-3対象製品の一覧を取得し、積替TBシートに表示する。

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

---

### 6.4 Ex詰替品更新.bas

#### `詰替品データ更新()`

**処理概要**: 積替TBシートで変更された詰替え要否フラグ（I列）を Oracle DB の `ExSeihinZ.tumikae` に更新する。

**処理フロー**:
1. `SeiKennsuu` で登録製品数確認
2. 確認ダイアログ表示
3. 積替TBシートの `TumikaeHinn` エリアをループ
4. I列（新値）≠ J列（元値）の場合のみDBを更新
5. `ExSeihinZ` テーブルを `kaisyacd` と `sehncd` をキーにINSERT/UPDATE
6. J列（元値）をI列（新値）に更新
7. 「積替品データ記録更新しました」メッセージ

---

### 6.5 Ex画面クリア.bas

#### `画面クリア3詰替()`

- `TumikaeHinn`、`TumikaeTB`、`Work` の各エリアをクリア
- `Innsatu` セルに 32 をセット
- `InsatuHanni` → 印刷範囲を設定

#### `modori()`

- `積替品` シートへ戻る

---

### 6.6 SQL_Execution.bas

Oracle DB への接続・SQL実行共通ルーチン（Ex1号機照射計画と同一構成）。

| メソッド | 説明 |
|---|---|
| `Open_oraconDB()` | DSN=ricdb で ADODB 接続を開く |
| `SQL_Exe(mySQL)` | SQL文を実行 |
| `SQL_INSERT_UPDATE(myTBL, myKey, myD(), myN)` | INSERT または UPDATE を汎用実行 |
| `SQL_Delete(myTBL, myWhere)` | DELETE を実行 |
| `Disp_Sheet(...)` | SQL結果をシートに貼り付け |
| `Set_Array(...)` | SQL結果を2次元配列に格納 |

---

## 7. DB 接続・外部連携

### ODBC 接続設定

| DSN 名 | UID | PWD | 用途 |
|---|---|---|---|
| `ricdb` | ric | t6101 | 積替え品・製品マスター取得・更新 |

### 参照テーブル一覧

| テーブル名 | 主な用途 | 主要カラム |
|---|---|---|
| `syoukj3` | RIC-3積替え品状態テーブル | syoush_f（照射状況）, sykno（線量計番号）, uno1（受付番号）, kainame1（会社名） |
| `zaiko` | 在庫情報 | uno, nouki（納期）, pass（パス数） |
| `ExKeikakuX` | 照射予定情報 | uno, syukkabi（出荷日）, bikou1（備考） |
| `ExSeihinZ` | 製品詰替え設定 | kaisyacd（会社コード）, sehncd（製品コード）, tumikae（積替え要否フラグ） |
| `tokumst` | 会社マスター | kaisyacd, coname（会社名） |
| `sehmst` | 製品マスター | kaisyacd, sehncd, seiname（製品名）, syouso（照射装置コード）, ric |

### 主要 SQL 文

```sql
-- 積替え品照射状況取得（起動時）
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

```sql
-- RIC-3製品マスター取得
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

```sql
-- 積替え要否フラグ更新
-- (INSERT/UPDATE に変換されて実行)
UPDATE ExSeihinZ
SET tumikae = '要否フラグ'
WHERE kaisyacd = '会社コード' AND sehncd = '製品コード'
```

---

## 8. データフロー

```
【起動フロー】
  ブック起動 (Workbook_Open)
       ↓
  画面クリア3詰替: 全エリアクリア・印刷範囲初期化
       ↓
  TumikaeHinn: DBから積替え品データ取得
    (syoukj3/zaiko/ExKeikakuX/ExSeihinZ JOIN → e.tumikae='1')
       ↓ WorkTBシートに一時格納
  DataKakou: データ加工（日付変換・状況コード変換・文字列クリーン）
       ↓
  積替品シートに表示・ソート（照射状況/積替日/線量計番号順）

【積替製品テーブル表示フロー】
  「積替する製品の登録」ボタンクリック
       ↓
  積替製品TB表示: tokumst/sehmst/ExSeihinZからSELECT
       ↓
  積替TBシートのTumikaeHinnエリアに表示
       ↓
  I列（詰替え要不要）を手動編集

【積替え要否更新フロー】
  「登録内容更新」ボタンクリック
       ↓
  詰替品データ更新:
    積替TBシートをループ
    I列（新値）≠ J列（元値）の行を抽出
       ↓
  ExSeihinZ テーブルへINSERT/UPDATE（kaisyacd/sehncd キー）
       ↓
  J列（元値）を更新して差分リセット
       ↓
  「積替品データ記録更新しました」
```

---

## 9. セキュリティ注意事項

| 種別 | キーワード | 内容 |
|---|---|---|
| AutoExec | `Workbook_Open` | ブック起動時に自動実行。DB接続・データ取得が行われる |
| AutoExec | `Workbook_BeforeClose` | 終了時に自動実行 |
| Suspicious | `Open` | ファイルオープン操作が含まれる可能性 |
| Suspicious | `Call` | DLL呼び出しの可能性 |
| Suspicious | `Windows` | アプリケーションウィンドウ列挙の可能性 |
| Suspicious | Hex Strings | 16進数エンコード文字列が検出された |
| Suspicious | Base64 Strings | Base64エンコード文字列が検出された |

> **重大な注意**: VBAコード中にDB接続パスワードが平文で記載されている。
> ```
> oraconn.ConnectionString = "DSN=ricdb;UID=ric;PWD=t6101"
> ```
> 積替TBシートの数式 `=IF(ISERROR(VLOOKUP(#REF!,TumiFlg,2,FALSE))...)` は `#REF!` エラーを含む。`TumiFlg` という名前付き範囲が削除または未定義の状態であり、修正が必要。
