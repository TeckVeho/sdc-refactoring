# ExRIC3線量不足報告書 仕様書

> **ファイル種別**: .xlsm（マクロ付き）
> **用途**: ３号機線量不足照射管理番号に対して追加照射指示書（工程確認・追加照射指示書）を作成・印刷する（一般向け、住友電工専用版とは別）
> **VBA プロジェクトサイズ**: 10 モジュール（ThisWorkbook, Sheet1, Sheet2, Sheet5, 初期化, データ抽出SQL, SQL_Execution, 印刷, Debug用, 改定履歴）
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
ExRIC3線量不足報告書.xlsm
├── シート
│   ├── 報告書          （メイン帳票・入力/表示）
│   ├── 試験成績書用紙  （テスト仕様書／結果報告書）
│   └── 改訂履歴        （様式改訂ログ）
└── VBA モジュール
    ├── ThisWorkbook.cls  （Workbook_Open / BeforeClose / BeforePrint）
    ├── Sheet1.cls        （Worksheet_Change イベント）
    ├── Sheet2.cls        （空）
    ├── Sheet5.cls        （空）
    ├── 初期化.bas        （画面初期設定 / クリア / CTRCls）
    ├── データ抽出SQL.bas （DB よりデータ取得・シート書き込み）
    ├── SQL_Execution.bas （ADODB 接続共通ライブラリ）
    ├── 印刷.bas          （印刷範囲の設定）
    ├── Debug用.bas       （デバッグ用イベント復旧）
    └── 改定履歴.bas      （サーバー保存処理）
```

---

## 2. シート詳細

### 2.1 報告書

**目的**: 線量不足と判定された照射管理番号を選択し、追加照射に必要な測定データ・照射計画データを DB から取得して一般向け帳票を完成させ、印刷する。

> **住電用との相違点**: タイトルが「３号機線量不足報告書（工程確認・追加照射指示書）」（住電用は「…住友電工専用…」）。追加パス数計算式（AA18〜AA22）が単純な `ROUNDUP` のみ（住電用は端数 2 未満判定あり）。`SyainTB` 範囲が `$BG$5:$BG$36`（住電用は `$BG$5:$BG$39`）。

#### レイアウト構造

| セル範囲 | 名前付き範囲 | 内容 |
|---|---|---|
| A1 | `Debug` | デバッグフラグ（通常空白） |
| B1 | — | 様式番号（数式で改訂版数を動的生成） |
| B2 | — | タイトル「３号機線量不足報告書（工程確認・追加照射指示書）」 |
| E3 | `Tuika` | 追加照射回目 |
| B7 | `SenkNo` | 照射管理番号（ドロップダウン入力） |
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
| B9:AN13 | `Data2` | 照射計画データ（最大5受付分） |
| B9〜B13 | `UkeNo1`〜`UkeNo5` | 受付番号（1〜5） |
| G9〜G13 | `RicNoS1`〜`RicNoS5` | 開始 RIC-NO |
| K9〜K13 | `RicNoE1`〜`RicNoE5` | 終了 RIC-NO |
| O9〜O13 | `SiteiSN1`〜`SiteiSN5` | 指定線量 |
| S9〜S13 | `Kagen1`〜`Kagen5` | 下限線量 |
| W9〜W13 | `Jyoug1`〜`Jyoug5` | 上限線量 |
| AA9〜AA13 | `JituP1`〜`JituP5` | 実パス数 |
| AD9〜AD13 | `KaiCD1`〜`KaiCD5` | 会社コード |
| AG9〜AG13 | `KaiName1`〜`KaiName5` | 会社名 |
| G15 | — | 線量計番号（`SenkNo` の下4桁） |
| G16 | `SEnKind` | 線量計種類 |
| O16 | `AtusaT` | 線量計厚さ |
| W16 | `Keisask` | 計算式コード |
| B18〜B22 | — | 追加照射データ行（受付番号参照） |
| AA18〜AA22 | — | 追加パス数計算（単純 ROUNDUP） |
| M24 | — | 測定データ不適切警告 |
| S24 | — | 追加パス数不一致警告 |
| BC3:BD6 | `mokuteki` / `MokuCD` | 製品目的コード・説明テーブル |
| BC10:BC12 | `SenSyu` | 線量計種類テーブル（DB から取得） |
| BC21 | `TuikaMin` | 線量計高・低閾値 |
| BG3 | `SebkNoTB` | 線量不足データ件数 |
| BG5:BG36 | `SyainTB` | 社員テーブル（DB から取得） |

---

### 2.2 試験成績書用紙

**目的**: 「３号機線量不足報告書」プログラムのテスト仕様書と結果報告書を兼ねた内部資料シート（ExRIC3線量不足報告書住電用と同構成）。

---

### 2.3 改訂履歴

**目的**: 様式 G3-08 の改訂ログを管理する。

| セル範囲 | 名前付き範囲 | 内容 |
|---|---|---|
| A4 | `YousikiName` | 様式名「３号機線量不足報告書」 |
| D4 | `YousikiNo` | 様式番号「G03-08」 |
| A6:D112 | `Rireki` | 改訂履歴データ本体 |

---

## 3. 名前付き範囲一覧

住電用と同一の名前付き範囲を持つ。`SyainTB` の参照先のみ異なる。

| 名前 | 参照先 | 説明 |
|---|---|---|
| `Atusa` | 報告書!$AJ$7 | 素子厚さ |
| `AtusaT` | 報告書!$O$16 | 線量計厚さ |
| `Data1` | 報告書!$G$7:$AN$7 | 測定データ行 |
| `Data2` | 報告書!$B$9:$AN$13 | 照射計画データ（5件分） |
| `Debug` | 報告書!$A$1 | デバッグフラグ |
| `JituP1`〜`JituP5` | 報告書!$AA$9〜$AA$13 | 実パス数（1〜5） |
| `Jyoug1`〜`Jyoug5` | 報告書!$W$9〜$W$13 | 上限線量（1〜5） |
| `Kagen1`〜`Kagen5` | 報告書!$S$9〜$S$13 | 下限線量（1〜5） |
| `KaiCD1`〜`KaiCD5` | 報告書!$AD$9〜$AD$13 | 会社コード（1〜5） |
| `KaiName1`〜`KaiName5` | 報告書!$AG$9〜$AG$13 | 会社名（1〜5） |
| `Keisask` | 報告書!$W$16 | 計算式コード |
| `MokuCD` | 報告書!$BC$3 | 製品目的コード先頭セル |
| `MokuSeu` | 報告書!#REF! | 参照エラー（要確認） |
| `mokuteki` | 報告書!$BC$3:$BD$6 | 製品目的テーブル |
| `Pass` | 報告書!$AA$7 | 指定パス数 |
| `RicNoE1`〜`RicNoE5` | 報告書!$K$9〜$K$13 | 終了 RIC-NO（1〜5） |
| `RicNoS1`〜`RicNoS5` | 報告書!$G$9〜$G$13 | 開始 RIC-NO（1〜5） |
| `Rireki` | 改訂履歴!$A$6:$D$112 | 改訂履歴データ |
| `SaiSoku` | 報告書!$O$7 | 再測定者 |
| `SebkNoTB` | 報告書!$BG$3 | 線量不足データ件数 |
| `SEnKind` | 報告書!$G$16 | 線量計種類 |
| `SenkNo` | 報告書!$B$7 | 照射管理番号 |
| `SenSyu` | 報告書!$BC$10:$BC$12 | 線量計種類テーブル |
| `SikiKodo` | 報告書!$AD$7 | 計算式コード（測定データ） |
| `SiteiSN1`〜`SiteiSN5` | 報告書!$O$9〜$O$13 | 指定線量（1〜5） |
| `SokDate` | 報告書!$G$7 | 測定日 |
| `SoktCD` | 報告書!$K$7 | 測定者コード |
| `SokutSN` | 報告書!$S$7 | 測定線量 |
| `Sosi` | 報告書!$AG$7 | 素子異常 |
| `SyainTB` | 報告書!$BG$5:$BG$36 | 社員テーブル（※住電用は $BG$39 まで） |
| `Tani` | 報告書!$W$7 | 照射目的 |
| `Tuika` | 報告書!$E$3 | 追加照射回目 |
| `TuikaMin` | 報告書!$BC$21 | 線量計高・低閾値 |
| `UkeNo1`〜`UkeNo5` | 報告書!$B$9〜$B$13 | 受付番号（1〜5） |
| `YousikiName` | 改訂履歴!$A$4 | 様式名 |
| `YousikiNo` | 改訂履歴!$D$4 | 様式番号 |

---

## 4. 数式一覧

### 報告書シート

住電用ファイルとほぼ同一の数式構成。**追加パス数計算（AA18〜AA22）のみ異なる**。

| セル | 数式 | 説明 |
|---|---|---|
| B1 | `="様式 G3-08("&MAX(改訂履歴!A6:A112)&")"` | 最新版数を動的表示 |
| M3 | `=IF(SebkNoTB>BH3,"追加照射データが多過ぎるため管理者に連絡","")` | 件数超過警告 |
| G15 | `=RIGHT(SenkNo,4)` | 線量計番号（下4桁） |
| B18〜B22 | `=IF(UkeNoN="","",UkeNoN)` | 受付番号（空白制御） |
| G18〜G22 | `=RicNoSN` | 開始 RIC-NO 参照 |
| K18〜K22 | `=RicNoEN` | 終了 RIC-NO 参照 |
| O18〜O22 | `=IF(BN="","",KagenN-SokutSN)` | 線量不足量（下限線量 − 測定線量） |
| S18〜S22 | `=ON` | 線量不足量コピー |
| W18〜W22 | `=IF(BN="","",WN-$S$7)` | 上限線量からの余裕 |
| **AA18** | `=IF(B18="","",ROUNDUP((O18/($S$7/AA9)),0))` | **追加パス数計算（単純切り上げ）** ※住電用は端数判定あり |
| AA19〜AA22 | ～（AA18 と同パターン、行番号変化） | 追加パス数2〜5 |
| AD18〜AD22 | `=KaiNameN` | 会社名参照 |
| M24 | `=IF(OR(AA18<0,...,AA22<0),"測定データが不適切","")` | 負値チェック |
| S24 | `=IF(MAX(AA18:AC22)<>MIN(AA18:AC22),"追加ﾊﾟｽ数異なるため課長に報告のこと","")` | パス数統一確認 |
| BF34 | `=#REF!` | 参照エラー（要確認） |
| A36〜AX36 | `=IF(ISERROR(VLOOKUP(SenkNo,Data,列番号,FALSE)),"",VLOOKUP(SenkNo,Data,列番号,FALSE))` | データ配列から列別に値を検索 |
| AQ39 | `=IF(ISERROR(VLOOKUP(TEXT(SenkNo,"0000"),Data,AQ37,FALSE)),"",VLOOKUP(TEXT(SenkNo,"0000"),Data,AQ37,FALSE))` | ゼロ埋め4桁で VLOOKUP |

> **住電用との計算式差異**: 住電用の AA18 は `ROUNDUP(...)+1` の条件を持つが、本ファイルは単純な `ROUNDUP` のみ。これは追加パス数に1パス余裕を加算するかどうかの業務差異を反映している。

---

## 5. ボタン・マクロ対応

住電用と同一。ボタンオブジェクトは検出されず。VBA コードにショートカット属性付きマクロあり。

| マクロ名 | ショートカット属性 | 処理概要 |
|---|---|---|
| `画面初期設定` | `\n14` | 起動時の初期化・DB取得 |
| `クリア` | `e\n14` | 入力フォームクリア |
| `CTRCls` | `e\n14` | `SenkNo` クリア後にクリア呼出 |

---

## 6. VBA モジュール仕様

VBA の全モジュール名・構成・処理ロジックは `ExRIC3線量不足報告書住電用.xlsm` と同一。以下に差分のみ記載する。

### 住電用との主な差異

| 項目 | 本ファイル | 住電用 |
|---|---|---|
| タイトル（B2） | 「３号機線量不足報告書（工程確認・追加照射指示書）」 | 「３号機線量不足報告書　住友電工専用（工程確認・追加照射指示書）」 |
| 追加パス数計算 | `ROUNDUP((O18/($S$7/AA9)),0)` のみ | `IF(ROUNDUP(...)*...-O18<2,...+1,ROUNDUP(...))` で端数 2 未満の場合に +1 |
| `SyainTB` 範囲 | `$BG$5:$BG$36` | `$BG$5:$BG$39` |

> 各モジュールの詳細な仕様は [`ExRIC3線量不足報告書住電用_仕様書.md`](./ExRIC3線量不足報告書住電用_仕様書.md) を参照のこと。

### VBA モジュール一覧

| モジュール | 主要プロシージャ | 処理概要 |
|---|---|---|
| `ThisWorkbook.cls` | `Workbook_Open`, `Workbook_BeforeClose`, `Workbook_BeforePrint` | 起動・終了・印刷前処理 |
| `Sheet1.cls` | `Worksheet_Change` | セル変更イベント（B7, G16） |
| `初期化.bas` | `画面初期設定`, `クリア`, `CTRCls` | 起動時初期化とフォームクリア |
| `データ抽出SQL.bas` | `線量不足線量計`, `社員データ`, `線量計種類`, `計算式コード取得`, `線量不足データ抽出`, `製品目的` | DB データ取得処理 |
| `SQL_Execution.bas` | `Open_oraconDB`, `SQL_Exe`, `SQL_INSERT_UPDATE`, `SQL_Delete`, `Disp_Sheet`, `Set_Array` | ADODB 接続共通ライブラリ |
| `印刷.bas` | `印刷範囲の設定` | 印刷範囲 B1:AN30 に設定 |
| `Debug用.bas` | `DebugCom` | イベント復旧用 |
| `改定履歴.bas` | `ブック保存` | サーバーパスへの `SaveAs` |

---

## 7. DB 接続・外部連携

住電用と同一の DB 接続設定・参照テーブル構成。

### ODBC 接続設定

| DSN 名 | UID | PWD | 用途 |
|---|---|---|---|
| `ricdb` | ric | t6101 | Oracle DB（照射管理システム） |

### 参照テーブル一覧

| テーブル名 | 主な用途 | 主要カラム |
|---|---|---|
| `RIC.SYOUJ2` | 照射実績（線量不足判定） | UNO, SENKNO, SENHFLG, SOKUTSN, SESDATE, TUIKAFLG |
| `RIC.SYOUKJ3` | 照射計画（最大5受付） | SYKNO, UNO1〜5, KAGENSN1〜5, KAINAME1〜5 |
| `SHAINMST` | 社員マスタ | SHANO, SHANAME, HSHIKA, LSHIKA |
| `RIC.ZAIKO` | 在庫マスタ（製品目的コード） | UNO, TANI |
| `SENKIND` | 線量計種類マスタ | SENSYU, KKODE, DISPNO |
| `KEICODE` | 計算式コードマスタ | KEISASK, YFLG1 |
| `ExSeihinJ` | ファイル保存先テーブル | FILENAME, FOLDER |

---

## 8. データフロー

住電用と同一のデータフロー。詳細は [`ExRIC3線量不足報告書住電用_仕様書.md`](./ExRIC3線量不足報告書住電用_仕様書.md) の「8. データフロー」を参照。

```
【起動フロー】
  Workbook_Open → 画面初期設定
       ↓
  DB(SYOUJ2/SYOUKJ3): 線量不足照射管理番号 → SenkNo ドロップダウン
  DB(SHAINMST): 社員テーブル → SyainTB (BG5:BG36)
  DB(SENKIND): 線量計種類 → SenSyu

【データ選択フロー】
  SenkNo 選択 → Worksheet_Change → 線量不足データ抽出
       ↓
  DB(SYOUJ2 + SHAINMST + ZAIKO): 測定実績 → Data1
  DB(SYOUKJ3): 照射計画 → Data2
       ↓
  AA18〜AA22 に単純 ROUNDUP で追加パス数計算

【印刷フロー】
  Workbook_BeforePrint → 未入力チェック → 線量計種類チェック → 印刷範囲設定(B1:AN30)
```

---

## 9. セキュリティ注意事項

住電用と同一の注意事項が適用される。

| 種別 | キーワード | 内容 |
|---|---|---|
| AutoExec | `Workbook_Open` | 起動時に自動で DB 接続・データ取得が実行される |
| AutoExec | `Workbook_BeforeClose` | `Debug` セルが空の場合はアラートなしにクローズ |
| AutoExec | `Worksheet_Change` | セル変更（B7・G16）で自動的に SQL が実行される |
| Suspicious | `Open` / `SaveAs` | `ブック保存` でサーバーパスに書き込む |
| Suspicious | `Call` | Excel 4 Macro として検知 |
| Suspicious | `Chr` | `Chr(13)` による改行（標準的な VBA 記述） |
| Suspicious | `Hex Strings` | VBA プロジェクト内に16進エンコード文字列が検出 |
| Suspicious | `Base64 Strings` | Base64 エンコード文字列が検出 |
| 認証情報 | `PWD=t6101` | DB パスワードが VBA コードにハードコードされている |
| 参照エラー | `MokuSeu = #REF!` | 名前付き範囲 `MokuSeu` が壊れている |
| 参照エラー | `BF34 = #REF!` | 報告書シートの BF34 セルに参照エラーがある |
