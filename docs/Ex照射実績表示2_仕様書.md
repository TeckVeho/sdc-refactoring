# Ex照射実績表示2 仕様書

> **ファイル種別**: .xlsm（マクロ付き）
> **用途**: 受付番号を入力して在庫または履歴在庫データを DB から取得し、照射実績（照射番号・照射位置・時刻・線量等）をパス回数分のマトリクス形式で一覧表示する照射実績照会ツール。2号機・3号機対象（1号機データは表示不可）。
> **VBA プロジェクトサイズ**: 10モジュール（Utility.bas / ThisWorkbook.cls / Sheet1〜3.cls / 画面クリア.bas / 起動.bas / 表示.bas / SQL_Execution.bas / 在庫実績Read.bas / 社員データRead.bas / 共通変数.bas）

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
Ex照射実績表示2.xlsm
├── シート
│   ├── 実績表示 （受付番号入力・照射実績表示メイン画面 131行×184列）
│   ├── 実績     （DB取得一時格納シート 110行×34列）
│   └── 社員     （社員マスタ一時格納シート 110行×6列）
├── VBA モジュール
│   ├── ThisWorkbook.cls    – ブックイベント（Open / BeforeClose）
│   ├── Sheet1.cls          – 実績シートイベント（空）
│   ├── Sheet2.cls          – 実績表示シートイベント（Worksheet_Change）
│   ├── Sheet3.cls          – 社員シートイベント（空）
│   ├── Utility.bas         – イベント有効化ユーティリティ
│   ├── 画面クリア.bas      – 画面クリア処理（GamenCls1 / GamennCls2）
│   ├── 起動.bas            – 検索起動メイン（起動Main）
│   ├── 表示.bas            – 在庫・実績データのシート表示処理
│   ├── SQL_Execution.bas   – DB 接続・SQL 実行共通ルーチン
│   ├── 在庫実績Read.bas    – 在庫・実績データ DB 取得処理
│   ├── 社員データRead.bas  – 社員マスタ DB 取得処理
│   └── 共通変数.bas        – 公開変数（mpDataCount / mpPass）
├── ボタン（1個）
│   └── 終了
└── 外部リンク: なし
```

---

## 2. シート詳細

### 2.1 実績表示

**目的**: 受付番号を入力して照射実績を表示するメイン画面。ヘッダー行（2〜6行目）に製品情報、7行目以降に各照射パスの実績データをパス回数（最大93パス）×実績件数のマトリクスで表示する。

#### レイアウト構造（ヘッダー部 B2〜N5）

| セル / 範囲 | 名前付き範囲 | 内容 |
|---|---|---|
| B2 | — | ラベル「受付番号」 |
| B3 | `Uno` | 受付番号入力欄（バーコードスキャン対応） |
| C2 | `Kaisyacd` | 会社コード |
| D2 | `Kainame` | 会社名（「〜様」付き） |
| G2 | `Sehncd` | 製品コード |
| H2 | `SName` | 製品名 |
| L2 | — | ラベル「指定線量kGy」 |
| N2 | `Siteisn` | 指定線量 |
| D3 | `Nyukasu` | 受付数量 |
| E3 | — | ラベル「受付日」 |
| F3 | `Koudate` | 受付日 |
| I3 | `Tani` | 分類（一般品等） |
| K3 | `Kousncd` | 担当者名 |
| N3 | — | Box入数またはPL積載数（装置種別により変化） |
| E4 | `Kagensn` | 管理点線量合格下限 |
| G4 | `Jyougsn` | 管理点線量合格上限 |
| I4 | `Labelcd` | ラベル有無 |
| K4 | `Pass` | 指定パス数 |
| M4 | `Syouso` | 装置（2=2号機 / 3=3号機） |
| D5 | `Syouzusu` | 照射済数 |
| G5 | `Senkssu` | 線量検査合格数 |
| I5 | `Syukasu` | 出荷済数 |
| K5 | `Syukabi` | 出荷日 |
| N5 | `Syukacd` | 出荷者名 |
| M3 | — | 「PL積載数」または「Box入数」（装置コードで切替） |

#### データ表示エリア（B7〜FZ131）

| セル / 範囲 | 名前付き範囲 | 内容 |
|---|---|---|
| B7 | — | 「ＲｉｃＮｏ」ラベル |
| C7〜FZ7 | `JissekiTB`（一部） | 各実績件数の最初の行（FRicNo / LRicNo） |
| C8〜C131 | — | 各パス行の照射種別（追加照射 / 通常照射）他 |
| B34:B131 | `PassKai` | パス番号（1〜N）の列 |

**データ配置規則**: 実績件数 j ごとに列を2本使用（奇数列: 実績値、偶数列: 付加情報）。最大 92 件 × 2列 = 184列。

### 2.2 実績

**目的**: DB から取得した在庫データ・照射実績データの一時格納シート。VBA によって上書きされる。

#### 在庫エリア（A3:T6）

| 行 | 内容 |
|---|---|
| 3行 | DBフィールド名ヘッダー（SELECT文の各フィールド名） |
| 4行 | DB取得値（`Disp_Sheet` で貼り付け） |
| 6行 | 加工済み表示値（数式による変換後の値） |

##### 行3のヘッダー（在庫）
`UNO`, `KAISYACD`, `KAINAME`, `SEHNCD`, `SITEISN`, `NYUKASU`, `NYUKABI`, `TANI`, `KOUSNCD`, `INCNT`, `KAGENSN`, `JYOUGSN`, `LABELCD`, `PASS`, `SYOUSO`, `SYOUZUSU`, `SENKSSU`, `SYUKASU`, `SYUKABI`, `SYUKACD`

#### 実績エリア（A10:AD110）

| 範囲 | 名前付き範囲 | 内容 |
|---|---|---|
| A10:AD110 | `Jisseki` | 照射実績データ（DB取得値。最大100件） |

##### 実績フィールド（A列〜AD列）
`UNO`, `FRICNO`, `LRICNO`, `TUIKAFLG`, `SURYOU`, `KMMDD`, `TONYUTIME`, `SAIKATIME`, `HANGNO`, `SYAFUTIME`, `SMMDD`, `DATTKTIME`, `SESDATE`, `SENKNO`, `SENKSYU`, `TORIDCD`, `ATUSA`, `SOKUTTI`, `KEISASK`, `ONDOK`, `SOKUTSN`, `SEZHNSU`, `JITUNO`, `KEIKACD`, `TONYUCD`, `SOKUTCD`, `SYUHNCD`, `NYUTIME`, `TAITIME`（列29）、作業者ID列（X〜AA）

### 2.3 社員

**目的**: 社員マスタを格納するシート。社員番号→社員名の VLOOKUP 参照元として使用。

| 範囲 | 名前付き範囲 | 内容 |
|---|---|---|
| B2:C110 | `Syainn` | 社員マスタ（B: 社員番号 / C: 社員名） |
| D1 | `ReadDate` | 社員データ読込日時 |
| E2:F7 | `Bunnrui` | 分類コードマスタ（E: コード / F: 分類名） |

#### 分類マスタ（Bunnrui）

| コード | 分類名 |
|---|---|
| 1 | 医療機器 |
| 2 | 一般品 |
| 3 | 医薬品 |
| 4 | 試験品 |
| 5 | その他（推定） |

---

## 3. 名前付き範囲一覧

| 名前 | 参照先 | 説明 |
|---|---|---|
| `Uno` | 実績表示!$B$3 | 受付番号入力欄 |
| `Kaisyacd` | 実績表示!$C$2 | 会社コード |
| `Kainame` | 実績表示!$D$2 | 会社名 |
| `Sehncd` | 実績表示!$G$2 | 製品コード |
| `SName` | 実績表示!$H$2 | 製品名 |
| `Siteisn` | 実績表示!$N$2 | 指定線量 |
| `Nyukasu` | 実績表示!$D$3 | 受付数量 |
| `Koudate` | 実績表示!$F$3 | 受付日 |
| `Tani` | 実績表示!$I$3 | 分類名 |
| `Kousncd` | 実績表示!$K$3 | 担当者名 |
| `Incnt` | 実績表示!$N$3 | Box入数またはPL積載数 |
| `Kagensn` | 実績表示!$E$4 | 下限線量 |
| `Jyougsn` | 実績表示!$G$4 | 上限線量 |
| `Labelcd` | 実績表示!$I$4 | ラベルコード |
| `Pass` | 実績表示!$K$4 | 指定パス数 |
| `Syouso` | 実績表示!$M$4 | 装置コード |
| `Syouzusu` | 実績表示!$D$5 | 照射済数 |
| `Senkssu` | 実績表示!$G$5 | 線量検査合格数 |
| `Syukasu` | 実績表示!$I$5 | 出荷済数 |
| `Syukabi` | 実績表示!$K$5 | 出荷日 |
| `Syukacd` | 実績表示!$N$5 | 出荷者名 |
| `JissekiTB` | 実績表示!$C$7:$FZ$131 | 実績表示テーブル（画面クリア対象） |
| `PassKai` | 実績表示!$B$34:$B$131 | パス番号列 |
| `Zaiko` | 実績!$A$3:$T$4 | 在庫データ格納範囲 |
| `Jisseki` | 実績!$A$10:$AD$110 | 実績データ格納範囲 |
| `SeiName` | 実績!$U$6 | 製品名（変換済み） |
| `Syainn` | 社員!$B$2:$C$110 | 社員番号・社員名テーブル |
| `ReadDate` | 社員!$D$1 | 社員データ読込日 |
| `Bunnrui` | 社員!$E$3:$F$7 | 分類コードマスタ |

---

## 4. 数式一覧

### 実績シート（6行目：加工済み表示行）

6行目は4行目（DB取得値）を変換して表示する。VBA の `在庫表示()` はこの6行目からデータを読み取る。

| セル | 数式 | 説明 |
|---|---|---|
| A6 | `=A4` | 受付番号そのまま |
| B6 | `=B4` | 会社コードそのまま |
| C6 | `=TRIM(C4)` | 会社名のトリミング |
| D6〜F6 | `=D4`〜`=F4` | そのまま |
| G6 | `=LEFT(G4,4)&"/"&MID(G4,5,2)&"/"&RIGHT(G4,2)` | 受付日 YYYYMMDD → YYYY/MM/DD 変換 |
| H6 | `=VLOOKUP(H4,Bunnrui,2,FALSE)` | 分類コード → 分類名（医療機器・一般品等） |
| I6 | `=VLOOKUP(VALUE(I4),Syainn,2,FALSE)` | 担当者社員番号 → 担当者名 |
| J6〜R6 | `=J4`〜`=R4` | そのまま |
| S6 | `=LEFT(S4,4)&"/"&MID(S4,5,2)&"/"&RIGHT(S4,2)` | 出荷日 YYYYMMDD → YYYY/MM/DD 変換 |
| T6 | `=IF(ISERROR(VLOOKUP(VALUE(T4),Syainn,2,FALSE)),"*",VLOOKUP(VALUE(T4),Syainn,2,FALSE))` | 出荷者社員番号 → 出荷者名（エラー時は"*"） |
| U6 | `=U4` | 製品名そのまま |
| AE6〜AH6 | `=IF(TRIM(X6)="0","---",TRIM(VLOOKUP(VALUE(X6),Syainn,2,FALSE)))` | 実績作業者1〜4: 社員番号0の場合は"---"、それ以外は社員名 |

（A〜AH の各行について同パターンが実績件数行数分（10〜110行目）繰り返し設定されている）

---

## 5. ボタン・マクロ対応

| シート | ボタンラベル | 割り当てマクロ | 動作概要 |
|---|---|---|---|
| 実績表示 | 終了 | `Bookを閉じる` | ブックを保存せずに閉じる（ブックが1つのみなら Excel 終了） |

---

## 6. VBA モジュール仕様

### 6.1 ThisWorkbook.cls

#### `Workbook_Open()`

**処理概要**: ブックを開いたとき、実績表示シートを UIのみ保護に設定し、画面クリアと社員データ読み込みを実行する。

```vba
Private Sub Workbook_Open()
    Sheets("実績表示").Unprotect
    Sheets("実績表示").Protect UserInterfaceOnly:=True
    Sheets("実績表示").EnableSelection = xlUnlockedCells
    Call GamenCls1
    Call 社員データ
End Sub
```

#### `Workbook_BeforeClose(Cancel)`

**処理概要**: ブックを閉じる前に保存確認を抑制する。

---

### 6.2 Sheet2.cls（実績表示シートイベント）

#### `Worksheet_Change(Target)`

**処理概要**: B3 セル（受付番号）が変更されたとき、受付番号が空なら `GamennCls2` でデータクリア、値があれば `起動Main` で検索・表示を実行する。

```vba
Private Sub Worksheet_Change(ByVal Target As Range)
    If Target.Row = 3 And Target.Column = 2 Then
        Application.EnableEvents = False
        If Range("Uno") = "" Then
            Call GamennCls2
        Else
            Call 起動Main
        End If
        Range("Uno").Select
        Application.EnableEvents = True
    End If
End Sub
```

---

### 6.3 画面クリア.bas

#### `GamenCls1()`

**処理概要**: 全シート（実績・社員・実績表示）の全名前付き範囲をクリアし、受付番号等すべての表示セルを空白に初期化する。

クリア対象:
- 実績シート: `Zaiko`・`Jisseki`
- 社員シート: `Syainn`・`ReadDate`
- 実績表示シート: `JissekiTB`・`PassKai`・各ヘッダー項目（Uno〜Sname）

#### `GamennCls2()`

**処理概要**: 受付番号（Uno）以外の表示セルおよびデータテーブルをクリアする（受付番号入力フィールドは保持）。

---

### 6.4 起動.bas

#### `起動Main()`

**処理概要**: 受付番号変更時の検索・表示メイン処理。

**処理フロー**:
1. `社員データ` で社員マスタを更新（毎回）
2. `在庫実績データ` で受付番号の在庫・実績データを DB から取得
3. エラーあり → エラーメッセージ表示 + `GamennCls2` でクリア
4. エラーなし → `GamennCls2` でデータエリアをクリア → `在庫表示` → `実績表示`

---

### 6.5 表示.bas

#### `在庫表示()`

**処理概要**: 実績シートの6行目（加工済み在庫データ）を読み取り、実績表示シートの各名前付き範囲（Uno〜Sname）に値をセットする。装置コード（Syouso）が2なら「Box入数」、3なら「PL積載数」のラベルを表示。

#### `実績表示()`

**処理概要**: 実績シートの10行目以降（DB取得実績データ）を読み取り、実績表示シートのマトリクス形式（行=パス番号、列=実績件数）に展開して表示する。

**表示変換ルール**:

| 実績フィールド（列） | 変換処理 |
|---|---|
| TUIKAFLG（列3） | 1→「追加照射」、それ以外→「通常照射」 |
| KMMDD / SMMDD / SEZHNSU 等（日付mmdd系） | `MM/DD` 形式に変換 |
| TONYUTIME / SAIKATIME / DATTKTIME 等（時刻hhmmss系） | `HH MM:SS` 形式に変換 |
| 入退室時刻（NYUTIME / TAITIME） | パス回数分（最大mpPass回）を6文字ずつ分解して `DD HH:MM` 形式に変換 |
| SOKUTSN（列20・測定値） | `RoundDown(..., 1)` で小数1桁に丸め |
| 装置=3（3号機）のシャッフリング列 | `"--"` を表示 |

パス番号（PassKai: B34〜B131）はパス数分（1〜n）をシートに書き込む。

---

### 6.6 SQL_Execution.bas

`Ex単価検索.xlsm` と同一の実装。DSN `ricdb`（UID: ric / PWD: t6101）での Oracle DB 接続・SQL 実行共通ルーチン。

本ファイルでは `Disp_Sheet` と `Set_Array` を使用してシートへの貼り付けと配列取得を行う。

---

### 6.7 在庫実績Read.bas

#### `在庫実績データ()`

**処理概要**: 受付番号（UNo）に対応する在庫データおよび照射実績データを DB から取得して実績シートに格納する。在庫がなければ履歴テーブルを参照する。1号機データ（syouso=1）は表示不可としてエラーメッセージを設定する。

**処理フロー**:
1. 受付番号で `zaiko` を照合して syouso=1 なら 1号機エラー
2. 受付番号で `rich.zaikor` を照合して syouso=1 なら 1号機エラー
3. 在庫SQL（`RIC.ZAIKO + ric.sehmst`）を実行して実績シート行3〜に貼り付け
4. 0件 → 履歴 SQL（`RICH.ZAIKOR + ric.sehmst`）に切替
5. 複数件 → エラー
6. 実績SQL（`RIC.SYOUJ2` または `RICH.SYOUJR2`）を実行して実績シート行10〜に貼り付け
7. `mpDataCount` に件数をセット

**在庫 SQL**:
```sql
SELECT z.uno, z.kaisyacd, z.kainame, z.sehncd, z.siteisn, z.nyukasu, z.nyukabi,
       TO_NUMBER(z.tani), z.kousncd, z.incnt, z.kagensn, z.jyougsn, z.labelcd, z.pass,
       z.syouso, z.syouzusu, z.senkssu, z.syukasu, z.syukabi, z.syukacd,
       TRIM(h.seiname)
FROM RIC.ZAIKO z, ric.sehmst h
WHERE uno = '<受付番号>'
  AND z.kaisyacd = h.kaisyacd
  AND z.sehncd = h.sehncd
```

**実績 SQL（在庫の場合 RIC.SYOUJ2、履歴の場合 RICH.SYOUJR2）**:
```sql
SELECT uno, fricno, lricno, tuikaflg, suryou*1, kmmdd, tonyutime, saikatime,
       hangno, syafutime, smmdd, dattktime, sesdate, senkno, senksyu, toridcd,
       atusa*1, sokutti*1, keisask, ondok, sokutsn*1, sezhnsu, jituno*1,
       keikacd, tonyucd, sokutcd, syuhncd, nyutime, taitime
FROM RIC.SYOUJ2
WHERE uno = '<受付番号>'
ORDER BY fricno, senkno, tuikaflg
```

---

### 6.8 社員データRead.bas

#### `社員データ()`

**処理概要**: `shainmst` から全社員データを取得して社員シートの `Syainn` 範囲に書き込む。

```sql
SELECT TO_NUMBER(shano), REPLACE(TRIM(shaname), '　', '')
FROM shainmst
ORDER BY shano
```

---

### 6.9 共通変数.bas

```vba
Public mpDataCount As Single  ' 実績データ件数
Public mpPass As Integer      ' 指定パス数（在庫表示で設定）
```

---

## 7. DB 接続・外部連携

### ODBC 接続設定

| DSN 名 | UID | PWD | 用途 |
|---|---|---|---|
| `ricdb` | ric | t6101 | Oracle DB接続（在庫・実績・社員マスタ） |

### 参照テーブル一覧

| テーブル名 | 主な用途 | 主要カラム |
|---|---|---|
| `RIC.ZAIKO` / `zaiko` | 在庫テーブル | uno, kaisyacd, kainame, sehncd, siteisn, pass, syouso, syouzusu, senkssu, syukasu, syukabi, syukacd |
| `RICH.ZAIKOR` | 在庫履歴テーブル | 同上（在庫から出荷済みのもの） |
| `ric.sehmst` | 製品仕様マスタ（製品名取得） | kaisyacd, sehncd, seiname |
| `RIC.SYOUJ2` | 照射実績テーブル（在庫品） | uno, fricno, lricno, tuikaflg, suryou, kmmdd, tonyutime, saikatime, smmdd, dattktime, sesdate, senkno, sokutsn, nyutime, taitime 等 |
| `RICH.SYOUJR2` | 照射実績履歴テーブル（出荷済み品） | 同上 |
| `shainmst` | 社員マスタ | shano, shaname |

---

## 8. データフロー

```
【起動フロー】
  ブックを開く（Workbook_Open）
       ↓
  実績表示シートを UIのみ保護
  GamenCls1() – 全シートの全表示データをクリア
  社員データ()
       ↓ DB（shainmst）: 全社員データ取得 → 社員シートに貼り付け
  実績表示シートに移動

【受付番号入力フロー】
  B3（Uno）に受付番号を入力（バーコードスキャン）
       ↓ Sheet2.Worksheet_Change
  起動Main() 呼び出し
  ┌─ 社員データ() – 社員マスタ再取得
  │        ↓
  └─ 在庫実績データ()
       ↓
       1号機チェック（zaiko / zaikor の syouso = 1 → エラー）
       ↓
       DB（RIC.ZAIKO + sehmst）: 在庫データ取得 → 実績シート 3〜4行
       └ 0件 → DB（RICH.ZAIKOR + sehmst）: 履歴在庫取得 → 実績シート 3〜4行
       ↓
       DB（RIC.SYOUJ2 または RICH.SYOUJR2）: 実績データ取得 → 実績シート 10〜110行
       ↓
  GamennCls2() – 実績表示エリアをクリア
  在庫表示()
       ↓ 実績シート6行目（数式加工済み）から実績表示シートのヘッダー欄に転記
  実績表示()
       ↓ 実績シート10行目以降から JissekiTB マトリクスに展開
         ┌ 日付(mmdd) → MM/DD 変換
         ├ 時刻(hhmmss) → HH MM:SS 変換
         ├ 入退室時刻(ddhhmmのパス数連結文字列) → DD HH:MM × パス数
         ├ 測定値 → 小数1桁丸め
         └ 3号機のシャッフリング列 → "--" 表示
  実績表示シートに表示完了
```

---

## 9. セキュリティ注意事項

| 種別 | キーワード | 内容 |
|---|---|---|
| AutoExec | `Workbook_Open` | 起動時に自動で社員データ DB 接続・取得が実行される |
| AutoExec | `Workbook_BeforeClose` | 閉じる前に保存フラグが自動操作される |
| AutoExec | `Worksheet_Change` | 受付番号変更時に自動で DB 照会処理が実行される |
| Suspicious | `Call` | Excel 4 マクロ形式の DLL 呼び出しの可能性 |
| Suspicious | `Hex Strings` | VBA 内に16進数エンコード文字列が存在 |
| Suspicious | `Base64 Strings` | VBA 内にBase64エンコード文字列が存在 |

> **注意**: DSN=ricdb の接続文字列（UID: ric / PWD: t6101）が `在庫実績Read.bas` の `在庫実績データ()` および `社員データRead.bas` の `社員データ()` にハードコードされています。パスワードが平文で記述されているため、コードの漏洩に注意してください。
> **注意**: `実績表示` シートは184列まで使用しており、照射実績件数が92件を超えた場合に列が不足する可能性があります。
