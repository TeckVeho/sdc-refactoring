# ExRIC2線量不足報告書 仕様書

> **ファイル種別**: .xlsm（マクロ付き）
> **用途**: ２号機線量不足照射管理番号に対して追加照射指示書（工程確認・追加照射指示書）を作成・印刷する
> **VBA プロジェクトサイズ**: 9 モジュール（ThisWorkbook, Sheet1, Sheet2, Sheet5, Sheet6, 初期化, データ抽出SQL, SQL_Execution, 印刷, 改定履歴）
> **外部連携**: Oracle DB（DSN=ricdb）/ ADODB 接続

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
ExRIC2線量不足報告書.xlsm
├── シート
│   ├── 報告書            （メイン帳票・入力/表示）
│   ├── 報告書用元ﾃﾞｰﾀ    （VLOOKUP 参照用データ領域）
│   ├── 試験成績書用紙    （テスト仕様書／結果報告書）
│   └── 改訂履歴          （様式改訂ログ）
└── VBA モジュール
    ├── ThisWorkbook.cls    （Workbook_Open / BeforeClose / BeforePrint）
    ├── Sheet1.cls          （Worksheet_Change イベント）
    ├── Sheet2.cls          （空）
    ├── Sheet5.cls          （空）
    ├── Sheet6.cls          （空）
    ├── 初期化.bas          （画面初期設定 / クリア / CTRCls）
    ├── データ抽出SQL.bas   （DB よりデータ取得・シート書き込み）
    ├── SQL_Execution.bas   （ADODB 接続共通ライブラリ）
    ├── 印刷.bas            （印刷範囲の設定）
    └── 改定履歴.bas        （サーバー保存処理）
```

---

## 2. シート詳細

### 2.1 報告書

**目的**: 線量不足と判定された照射管理番号（線量計番号）を選択し、追加照射に必要な測定データ・照射計画データを DB から取得して帳票を完成させ、印刷する。RIC3版と比較して受付データが最大 8 件（RIC3は5件）に拡張されており、2パス単位の追加パス数計算を採用している。

#### レイアウト構造

| セル範囲 | 名前付き範囲 | 内容 |
|---|---|---|
| A1 | `Debug` | デバッグフラグ |
| B1 | — | 様式番号（数式で改訂版数を動的生成） |
| B2 | — | タイトル「２号機線量不足報告書（工程確認・追加照射指示書）」 |
| B7 | `SenkNo` | 線量計番号（ドロップダウン入力） |
| G7 | `SokDate` | 測定日 |
| K7 | `SoktCD` | 測定者コード |
| O7 | `SaiSoku` | 再測定者 |
| S7 | `SokutSN` | 測定線量 |
| W7 | `Tani` | 製品目的 |
| AA7 | `Pass` | 指定パス数 |
| AD7 | `SikiKodo` | 計算式コード |
| AG7 | `Sosi` | 素子異常 |
| AJ7 | `Atusa` | 素子厚さ |
| G7:AN7 | `Data1` | 測定データ行 |
| C9:AN16 | `Data2` | 照射計画データ（最大8受付分） |
| C9〜C16 | `UkeNo1`〜`UkeNo8` | 受付番号（1〜8回目、1次データ） |
| C22〜C29 | `UkeNo11`〜`UkeNo88` | 受付番号（1〜8回目、2次データ / 追加照射行） |
| G9〜G16 | `RicNoS1`〜`RicNoS8` | 開始 RIC-NO |
| K9〜K16 | `RicNoE1`〜`RicNoE8` | 終了 RIC-NO |
| O9〜O16 | `SiteiSN1`〜`SiteiSN8` | 指定線量 |
| S9〜S16 | `Kagen1`〜`Kagen8` | 下限線量 |
| W9〜W16 | `Jyoug1`〜`Jyoug8` | 上限線量 |
| AA9〜AA16 | `JituP1`〜`JituP8` | 実パス数 |
| AD9〜AD16 | `KaiCD1`〜`KaiCD8` | 会社コード |
| AG9〜AG16 | `KaiName1`〜`KaiName8` | 会社名 |
| B9 | — | 「Box1」（印刷種別チェックボックス） |
| B13 | — | 「Box0」（印刷種別チェックボックス） |
| G19 | — | 線量計番号（`SenkNo` 参照） |
| G20 | `SEnKind` | 線量計種類 |
| O20 | `AtusaT` | 線量計厚さ |
| W20 | `Keisask` | 計算式コード |
| C22〜C29 | — | 追加照射データ行（UkeNoN参照） |
| AA22〜AA29 | — | 追加パス数計算（2パス単位の ROUNDUP） |
| BB3:BC6 | `mokuteki` / `MokuCD` | 製品目的コード・説明テーブル |
| BB10:BB12 | `SenSyu` | 線量計種類テーブル（DB から取得） |
| BB25 | `TuikaMin` | 線量計高・低閾値 |
| BC10 | `Innsatu` | 印刷種別文字列（「新規印刷」） |
| BE3:BE5 | `SebkNoTB` | 線量不足線量計番号リスト（最大3件） |
| BF3 | — | `COUNTA(SebkNoTB)` で件数表示 |
| BF5:BF35 | `SyainTB` | 社員テーブル（DB から取得） |
| BG3 | — | 最大件数（=30） |

（行番号 17〜19 は主にラベル行、行 30〜46 は追加データ領域・チェック用）

#### 追加照射データ列ヘッダー（行 21）

| 列 | 内容 |
|---|---|
| C | 受付番号 |
| G | 開始 RIC-NO |
| K | 終了 RIC-NO |
| O | 線量不足量（下限線量 − 測定線量） |
| S | 線量不足量コピー |
| W | 上限線量からの余裕 |
| AA | 追加パス数（2パス単位） |
| AD | 会社名 |

---

### 2.2 報告書用元ﾃﾞｰﾀ

**目的**: `Data` 名前付き範囲（`A1:AQ30`）として定義された VLOOKUP 参照用の元データ領域。VBA が DB から取得したデータを一時格納し、`報告書` シートの行36以降の数式が参照する。

| セル範囲 | 名前付き範囲 | 内容 |
|---|---|---|
| A1:AQ30 | `Data` | VLOOKUP ルックアップテーブル（VBA が書き込む） |

---

### 2.3 試験成績書用紙

**目的**: 「２号機線量不足報告書」プログラムのテスト仕様書と結果報告書（RIC3版と同構成）。

---

### 2.4 改訂履歴

**目的**: 様式 G2-12 の改訂ログを管理する。

| セル範囲 | 名前付き範囲 | 内容 |
|---|---|---|
| A4 | `YousikiName` | 様式名（２号機線量不足報告書） |
| D4 | `YousikiNo` | 様式番号（G2-12） |
| A6:D105 | `Rireki` | 改訂履歴データ本体 |

---

## 3. 名前付き範囲一覧

RIC3版との主な相違: 受付件数が 8件分（1〜8 + 11〜88）、`SebkNoTB` が範囲型（リスト）、`Data` が別シート参照、`Innsatu` が追加。

| 名前 | 参照先 | 説明 |
|---|---|---|
| `Atusa` | 報告書!$AJ$7 | 素子厚さ |
| `AtusaT` | 報告書!$O$20 | 線量計厚さ（RIC3は$O$16） |
| `Data` | 報告書用元ﾃﾞｰﾀ!$A$1:$AQ$30 | VLOOKUP 参照元データ（別シート） |
| `Data1` | 報告書!$G$7:$AN$7 | 測定データ行 |
| `Data2` | 報告書!$C$9:$AN$16 | 照射計画データ（8件分） |
| `Debug` | 報告書!$A$1 | デバッグフラグ |
| `Innsatu` | 報告書!$BC$10 | 印刷種別文字列（新規印刷） |
| `JituP1`〜`JituP8` | 報告書!$AA$9〜$AA$16 | 実パス数（1〜8） |
| `Jyoug1`〜`Jyoug8` | 報告書!$W$9〜$W$16 | 上限線量（1〜8） |
| `Kagen1`〜`Kagen8` | 報告書!$S$9〜$S$16 | 下限線量（1〜8） |
| `KaiCD1`〜`KaiCD8` | 報告書!$AD$9〜$AD$16 | 会社コード（1〜8） |
| `KaiName1`〜`KaiName8` | 報告書!$AG$9〜$AG$16 | 会社名（1〜8） |
| `Keisask` | 報告書!$W$20 | 計算式コード（RIC3は$W$16） |
| `MokuCD` | 報告書!$BB$3 | 製品目的コード先頭セル（RIC3は$BC$3） |
| `mokuteki` | 報告書!$BB$3:$BC$6 | 製品目的テーブル |
| `Pass` | 報告書!$AA$7 | 指定パス数 |
| `RicNoE1`〜`RicNoE8` | 報告書!$K$9〜$K$16 | 終了 RIC-NO（1〜8） |
| `RicNoS1`〜`RicNoS8` | 報告書!$G$9〜$G$16 | 開始 RIC-NO（1〜8） |
| `Rireki` | 改訂履歴!$A$6:$D$105 | 改訂履歴データ |
| `SaiSoku` | 報告書!$O$7 | 再測定者 |
| `SebkNoTB` | 報告書!$BE$3:$BE$5 | 線量不足線量計番号リスト（最大3件、動的リサイズ） |
| `SEnKind` | 報告書!$G$20 | 線量計種類 |
| `SenkNo` | 報告書!$B$7 | 線量計番号（入力セル） |
| `SenSyu` | 報告書!$BB$10:$BB$12 | 線量計種類テーブル |
| `SikiKodo` | 報告書!$AD$7 | 計算式コード（測定データ） |
| `SiteiSN1`〜`SiteiSN8` | 報告書!$O$9〜$O$16 | 指定線量（1〜8） |
| `SokDate` | 報告書!$G$7 | 測定日 |
| `SoktCD` | 報告書!$K$7 | 測定者コード |
| `SokutSN` | 報告書!$S$7 | 測定線量 |
| `Sosi` | 報告書!$AG$7 | 素子異常 |
| `SyainTB` | 報告書!$BF$5:$BF$35 | 社員テーブル |
| `Tani` | 報告書!$W$7 | 照射目的 |
| `TuikaMin` | 報告書!$BB$25 | 線量計高・低閾値（RIC3は$BC$21） |
| `UkeNo1`〜`UkeNo8` | 報告書!$C$9〜$C$16 | 受付番号（1次） |
| `UkeNo11`〜`UkeNo88` | 報告書!$C$22〜$C$29 | 受付番号（2次・追加照射行） |
| `YousikiName` | 改訂履歴!$A$4 | 様式名 |
| `YousikiNo` | 改訂履歴!$D$4 | 様式番号 |

---

## 4. 数式一覧

### 報告書シート

| セル | 数式 | 説明 |
|---|---|---|
| B1 | `="様式 G2-12("&MAX(改訂履歴!A6:A105)&")"` | 改訂履歴から最新版数を動的表示 |
| M3 | `=IF(BF3>BG3,"追加照射データが多過ぎるため管理者に連絡","")` | 件数超過警告 |
| BF3 | `=COUNTA(SebkNoTB)` | 線量不足線量計件数を `SebkNoTB` 範囲から集計 |
| G19 | `=SenkNo` | 選択中の線量計番号を表示 |
| C22〜C29 | `=IF(UkeNoN="","",UkeNoN)` | 追加照射行の受付番号（空白制御、UkeNo1〜8参照） |
| G22〜G29 | `=RicNoSN` | 開始 RIC-NO 参照 |
| K22〜K29 | `=RicNoEN` | 終了 RIC-NO 参照 |
| O22〜O29 | `=IF(UkeNo11="","",KagenN-SokutSN)` | 線量不足量（下限線量 − 測定線量） |
| S22〜S29 | `=IF(ON="","",ON)` | 線量不足量コピー |
| W22〜W29 | `=IF(UkeNo11="","",JyougN-SokutSN)` | 上限線量からの余裕 |
| **AA22** | `=IF(UkeNo11="","",ROUNDUP((O22/(SokutSN/JituP1))/2,0)*2)` | **追加パス数を2パス単位で切り上げ計算** |
| AA23〜AA29 | ～（AA22 と同パターン） | 追加パス数2〜8 |
| AD22〜AD29 | `=KaiNameN` | 会社名参照 |
| 行36以降 | `=IF(ISERROR(VLOOKUP(SenkNo,Data,列番号,FALSE)),"",VLOOKUP(SenkNo,Data,列番号,FALSE))` | `Data`（報告書用元データ）からの VLOOKUP |

> **追加パス数計算の差異（RIC2 vs RIC3）**:
> - RIC2: `ROUNDUP((O22/(SokutSN/JituP1))/2,0)*2` → **2パス単位**での切り上げ（ガンマ照射の特性上、2パス単位が必要）
> - RIC3（一般）: `ROUNDUP((O18/($S$7/AA9)),0)` → 単純な1パス単位切り上げ
> - RIC3（住電用）: `IF(ROUNDUP(...)*...-O18<2,...+1,ROUNDUP(...))` → 端数2未満の場合に+1パス

---

## 5. ボタン・マクロ対応

ボタンオブジェクトは抽出スクリプトでは検出されなかった。

VBA コードでショートカット属性付きマクロ：

| マクロ名 | ショートカット属性 | 処理概要 |
|---|---|---|
| `画面初期設定` | `s\n14` | 起動時の初期化・DB取得 |
| `クリア` | `e\n14` | 入力フォームクリア |
| `CTRCls` | `e\n14` | `SenkNo` クリア後にクリア呼出 |

---

## 6. VBA モジュール仕様

### 6.1 ThisWorkbook.cls

RIC3版と同一の構成（`Workbook_Open` → `画面初期設定`、`BeforeClose` → サイレントクローズ、`BeforePrint` → 未入力チェック）。

---

### 6.2 Sheet1.cls（報告書シート）

#### `Worksheet_Change(Target As Range)`

RIC3版と同一の構成（線量計種類変更 → `計算式コード取得`、照射管理番号変更 → `線量不足データ抽出`）。

---

### 6.3 初期化.bas

#### `画面初期設定()`

**処理フロー**:
1. イベント無効化
2. `クリア` でシートリセット
3. `線量不足線量計` で線量計番号リスト取得（`SebkNoTB` に書き込み）
4. `社員データ` で社員テーブル取得
5. `線量計種類` で線量計種類取得
6. `SenkNo` セル選択・イベント再有効化

> **RIC3版との差異**: `クリア` 内で `Debug` フラグチェックし、空の場合のみシート保護を設定する。

#### `クリア()`

`Data1`, `Data2`, `SEnKind`, `AtusaT`, `Keisask`, `SenSyu`, `SebkNoTB`, `SyainTB` をクリア後、A1:AO1 範囲に合わせてズーム設定。

---

### 6.4 データ抽出SQL.bas

#### `線量不足線量計()`

**処理概要**: 線量不足フラグが立っている線量計番号を DB から抽出し、`SebkNoTB` セル範囲（BE3〜BE5）に書き込む。

> **RIC3版との主な差異**:
> - 参照テーブル: `SYOUJ2` + `SYOUK2`（RIC3は `SYOUKJ3`）
> - 結合条件: `s.senkno=sk.senkno`（RIC3は `s.UNO = sk.UNO1`）
> - 追加条件: `sk.tuikaflg='0'`（RIC3は `sk.SYOUSH_F='2'`）
> - 結果書き込み先: `SebkNoTB` に線量計番号リストを直接書き込み（RIC3は SenkNo のドロップダウン）

```sql
-- 線量不足線量計番号抽出（RIC2用）
SELECT DISTINCT s.senkno
FROM syouj2 s, syouk2 sk
WHERE s.senkno=sk.senkno AND s.syouflg='1'
  AND (s.senhflg='2' OR s.senhflg='4') AND s.syouso='2'
  AND sk.tuikaflg='0'
```

#### `社員データ()`

```sql
SELECT TRIM(SHANAME)
FROM SHAINMST
WHERE HSHIKA='1' AND LSHIKA='1'
ORDER BY SHANO
```

> RIC3版との差異: `shano>'1000'` 条件なし

#### `線量計種類()`

RIC3版と同一。

#### `計算式コード取得(SenKind)`

RIC3版とほぼ同一だが、コードフィルタが `"2"` を識別する（RIC2対応）。

```vba
If Mid(myData(i, 0), 2, 1) = "2" Then   '例:X2F->2:Ric2
```

#### `線量不足データ抽出()`

**処理概要**: 選択された線量計番号に対応する測定実績と照射計画を DB から取得してシートに書き込む。

> **RIC3版との主な差異**:
> - `SenkNo` を右4桁ゼロ埋め（`Right("0000" & Range("SenkNo"), 4)`）で正規化
> - 照射計画テーブル: `SYOUK2`（RIC3は `SYOUKJ3`）
> - 受付データ件数: 最大 8件（RIC3は5件）
> - 照射計画の WHERE: `TUIKAFLG='0'`（RIC3は `SYOUSH_F='2'`）
> - 追加回数（`Tuika`）を書き込まない

```sql
-- 測定実績取得（RIC2用）
SELECT S.SESDATE, sha.shaname,
       TO_NUMBER(S.SOKUTSN), TO_NUMBER(za.tani), TO_NUMBER(S.PASS),
       S.KEISASK, TO_NUMBER(S.ATUSA), TO_NUMBER(S.JITUNO), s.kmmdd
FROM SHAINMST SHA, RIC.SYOUJ2 S, RIC.ZAIKO ZA
WHERE s.sokutcd=sha.shano AND za.uno=s.uno
  AND (s.senhflg='2' OR s.senhflg='4')
  AND s.senkno='[SenkNoゼロ埋め4桁]'

-- 照射計画取得（RIC2用、最大8件）
SELECT TO_NUMBER(UNO1), FRICNO1, LRICNO1, SITEISN1, KAGENSN1, JYOUGSN1, kaisyacd1, KAINAME1,
       ...(UNO8まで同パターン)
FROM RIC.SYOUK2
WHERE SENKNO = '[SenkNoゼロ埋め4桁]' AND TUIKAFLG='0'
```

---

### 6.5 SQL_Execution.bas / 印刷.bas / 改定履歴.bas

RIC3版と同一。DB 接続文字列: `DSN=ricdb;UID=ric;PWD=t6101`

---

## 7. DB 接続・外部連携

### ODBC 接続設定

| DSN 名 | UID | PWD | 用途 |
|---|---|---|---|
| `ricdb` | ric | t6101 | Oracle DB（照射管理システム） |

### 参照テーブル一覧

| テーブル名 | 主な用途 | 主要カラム |
|---|---|---|
| `RIC.SYOUJ2` | 照射実績（線量不足判定） | SENKNO, SENHFLG, SOKUTSN, SESDATE, SOKUTCD, SYOUSO, KMMDD |
| `RIC.SYOUK2` | 照射計画（最大8受付）※RIC2固有テーブル | SENKNO, UNO1〜8, FRICNO1〜8, KAGENSN1〜8, KAINAME1〜8, TUIKAFLG |
| `SHAINMST` | 社員マスタ | SHANO, SHANAME, HSHIKA, LSHIKA |
| `RIC.ZAIKO` | 在庫マスタ（製品目的コード） | UNO, TANI |
| `SENKIND` | 線量計種類マスタ | SENSYU, KKODE, DISPNO |
| `KEICODE` | 計算式コードマスタ（RIC2: `Mid(KEISASK,2,1)="2"`） | KEISASK, YFLG1 |
| `ExSeihinJ` | ファイル保存先テーブル | FILENAME, FOLDER |

> **RIC3との DB テーブル差異**: `SYOUK2`（RIC2用計画テーブル）vs `SYOUKJ3`（RIC3用計画テーブル）

---

## 8. データフロー

```
【起動フロー】
  Workbook_Open → 画面初期設定
       ↓
  DB(SYOUJ2 + SYOUK2): 線量不足線量計番号抽出 → SebkNoTB(BE3:BE5)に書き込み
  DB(SHAINMST): 社員テーブル → SyainTB(BF5:BF35)
  DB(SENKIND): 線量計種類 → SenSyu

【データ選択フロー】
  SenkNo 入力（BF3 の COUNTA で件数更新）
       ↓ Worksheet_Change(B7)
  線量不足データ抽出
       ↓
  SenkNo をゼロ埋め4桁に変換
       ↓
  DB(SYOUJ2 + SHAINMST + ZAIKO): 測定実績 → Data1(G7:AN7)
  DB(SYOUK2): 照射計画（最大8件、TUIKAFLG='0'）→ Data2(C9:AN16)
       ↓
  AA22〜AA29 に 2パス単位の ROUNDUP で追加パス数計算

【線量計種類変更フロー】
  SEnKind(G20) 変更 → Worksheet_Change(G20)
       ↓
  DB(SENKIND): KKODE 取得
  DB(KEICODE): Mid(KEISASK,2,1)="2" でフィルタ → Keisask 設定

【印刷フロー】
  Workbook_BeforePrint
       ↓
  未入力チェック（SaiSoku / Sosi / SEnKind / AtusaT）
       ↓
  線量計種類チェック（TuikaMin 基準の高/低判定）
       ↓
  印刷範囲設定 → 印刷実行

【保存フロー】
  ブック保存 → DB(ExSeihinJ)でフォルダ取得 → SaveAs でサーバー保存
```

---

## 9. セキュリティ注意事項

| 種別 | キーワード | 内容 |
|---|---|---|
| AutoExec | `Workbook_Open` | 起動時に自動で DB 接続・データ取得が実行される |
| AutoExec | `Workbook_BeforeClose` | `Debug` セルが空の場合はアラートなしにクローズ |
| AutoExec | `Worksheet_Change` | セル変更（B7・G20）で自動的に SQL が実行される |
| Suspicious | `Open` / `SaveAs` | `ブック保存` でサーバーパスへの書き込みが実行される |
| Suspicious | `Call` | Excel 4 Macro として検知 |
| Suspicious | `Chr` | `Chr(13)` による改行（標準的な VBA 記述） |
| Suspicious | `Hex Strings` | VBA プロジェクト内に16進エンコード文字列が検出 |
| Suspicious | `Base64 Strings` | Base64 エンコード文字列が検出 |
| 認証情報 | `PWD=t6101` | DB パスワードが VBA コードにハードコードされている |
| 業務リスク | `Debug` フラグ | `Debug` セルに値がある場合、シート保護が無効になる |
