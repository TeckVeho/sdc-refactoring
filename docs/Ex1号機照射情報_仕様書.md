# Ex1号機照射情報 仕様書

> **ファイル種別**: .xlsm（マクロ付き）
> **用途**: 1号機の固定照射中製品について、照射完了までの残時間・完了予想日時を DB からリアルタイムに取得してシートに一覧表示する照射情報モニタリングツール
> **VBA プロジェクトサイズ**: 10モジュール（ThisWorkbook.cls / Sheet1〜5.cls / Ex抽出処理.bas / SQL_Execution.bas / Ex画面クリア.bas / Ex終了照射情報.bas / Ex出荷日情報読込Ric1.bas / FunctionR1.bas / GetPathRic1Jyou.bas）
> **外部連携ファイル**: `ExAprReadPath.txt`（ブックと同一フォルダ、パス設定ファイル）

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
Ex1号機照射情報.xlsm
├── シート
│   ├── 完了予定時間 （表示用メインシート）
│   ├── 照射ﾃﾞｰﾀ     （DB取得一時格納シート）
│   ├── 線源情報      （線源状態 DB取得一時格納シート）
│   └── 出荷日情報    （ExKeikakuX から出荷日を格納するシート）
├── VBA モジュール
│   ├── ThisWorkbook.cls          – ブックイベント（Open / BeforeClose）
│   ├── Sheet1〜5.cls             – シートイベント（実装なし・空）
│   ├── Ex抽出処理.bas            – メイン更新処理（Kousinn / Yomikomi / SenGenn）
│   ├── SQL_Execution.bas         – DB 接続・SQL 実行共通ルーチン
│   ├── Ex画面クリア.bas          – 画面クリア処理（GamennCls）
│   ├── Ex終了照射情報.bas        – ブック終了処理
│   ├── Ex出荷日情報読込Ric1.bas  – 出荷日情報読み込み（SyukabiRead）
│   ├── FunctionR1.bas            – 日付変換ユーティリティ関数
│   └── GetPathRic1Jyou.bas       – パス設定ファイル読み込み（ExAprReadPath.txt）
├── ボタン（2個）
│   ├── 終了
│   └── 最新情報
└── 外部ファイル: ExAprReadPath.txt（パス設定）
```

---

## 2. シート詳細

### 2.1 完了予定時間

**目的**: 1号機で照射中の製品の完了までの残時間・完了予想日時・出荷日を一覧表示するメインモニタリング画面。VBA によってリアルタイムに更新される。

#### レイアウト構造

| セル範囲 | 名前付き範囲 | 内容 |
|---|---|---|
| A1 | — | バージョン識別子（`ExM`） |
| D2 | — | タイトル「1号機固定照射情報」 |
| F2:G2 | `Jyoutai`（一部） | 線源状態表示（照射中 / 貯蔵中 / 昇降中 / 移動照射中） |
| H2:J2 | — | 移動照射中メッセージ表示欄（`=IF(F2="移動照射中","完了予想時刻は担当者に確認のこと","")` 数式） |
| F3:G3 | — | 更新日時表示（`Range("F3") = myNow` でVBAから書き込み） |
| C4 | — | ヘッダー「No」 |
| D4 | — | ヘッダー「受付番号」 |
| E4 | — | ヘッダー「会社名」 |
| F4 | — | ヘッダー「指定線量」 |
| G4 | — | ヘッダー「数量」 |
| H4 | — | ヘッダー「照射位置」 |
| I4 | — | ヘッダー「完了までの時間」 |
| J4 | — | ヘッダー「完了予想日時」 |
| K4 | — | ヘッダー「出荷日」 |
| L4 | — | ヘッダー「備考」 |
| C5:J50 | `Hyouji` | データ表示テーブル（C:No / D:受付番号 / E:会社名 / F:線量 / G:数量 / H:照射位置 / I:完了までの時間 / J:完了予想日時） |
| D5:K50 | `SortTB` | ソート対象テーブル（I列:完了までの時間 で昇順ソート） |
| K5:K50 | — | 出荷日（VLOOKUP: SyukkabiTB 第2列） |
| L5:L50 | — | 備考（VLOOKUP: SyukkabiTB 第3列） |

（行番号 51〜以降は空）

### 2.2 照射ﾃﾞｰﾀ

**目的**: `syouk1` テーブルから取得した照射中製品データの一時格納シート。ヘッダー行（1行目）のみ静的定義、データ行は VBA で書き込まれる。

#### DB フィールド対応

| 列 | ヘッダー | DB フィールド | 説明 |
|---|---|---|---|
| A | UNO | syouk1.UNO | 受付番号 |
| B | SYONO | syouk1.SYONO | 照射番号 |
| C | KAINAME | syouk1.KAINAME | 会社名 |
| D | SITEISN | syouk1.SITEISN | 指定線量 |
| E | SYOSUU | syouk1.SYOSUU | 数量 |
| F | SYOICHI | syouk1.SYOICHI | 照射位置コード |
| G | SYOTIME | syouk1.SYOTIME | 照射時間 |
| H | HANSUU | syouk1.HANSUU | 反数 |
| I | STIMER | syouk1.STIMER | 開始タイマー値 |
| J | KTIMER | syouk1.KTIMER | 完了タイマー値 |
| K | SENRITU | syouk1.SENRITU | 線率 |
| L | SYOSTAT | syouk1.SYOSTAT | 照射状態（1:開始/2:中断/3:完了/4:取消/5:再開/6:修正） |
| M | CTIMER | syouk1.CTIMER | 現在タイマー値 |
| N | ZHANSUU | syouk1.ZHANSUU | 残数 |
| O | HTIMER | syouk1.HTIMER | 完了タイマー値（HTIMER） |
| P | SLOTNO | syouk1.SLOTNO | スロット番号 |
| Q | SDATE | syouk1.SDATE | 開始日付 |
| R | EDATE | syouk1.EDATE | 終了日付 |
| S | UPDFLG | syouk1.UPDFLG | 更新フラグ |
| T | SYOKIND | syouk1.SYOKIND | 照射種別（'2' 以外を対象） |
| U | BIKOU | syouk1.BIKOU | 備考 |

### 2.3 線源情報

**目的**: `RIC.SENGNR1` テーブルから最新の線源状態（日付・時刻・タイマー値・イベント状態）を格納するシート。VBA で毎回上書きされる。

| 列 | ヘッダー | DB フィールド | 説明 |
|---|---|---|---|
| A | SDATE | SENGNR1.SDATE | 日付（YYYYMMDD形式） |
| B | STIME | SENGNR1.STIME | 時刻（HHMMSS形式） |
| C | TIMER | SENGNR1.TIMER | タイマー値（照射経過時間） |
| D | EVENT | SENGNR1.EVENT | 線源状態（0:昇降中 / 1:照射中 / 2:貯蔵中 / 3:移動照射中） |

### 2.4 出荷日情報

**目的**: `ExKeikakuX` テーブルから1号機（souti='1'）の出荷日・備考・出荷方法を格納するシート。完了予定時間シートの VLOOKUP 参照元として使用。

| 列 | ヘッダー | DB フィールド / 内容 |
|---|---|---|
| A | 受付番号 | ExKeikakuX.uno（数値変換済み） |
| B | 出荷日 | ExKeikakuX.syukkabi（trim後） |
| C | 備考 | ExKeikakuX.bikou1 |
| D | 出荷方法 | ExKeikakuX.syuhouhou |
| G | 受付番号 | （内部参照用の第2受付番号列） |
| H | 出荷日 | （内部参照用の第2出荷日列） |
| I | 備考 | （内部参照用の第2備考列） |

（行番号 2〜4491 にデータ格納。実データ行数はDB件数に依存）

---

## 3. 名前付き範囲一覧

| 名前 | 参照先 | 説明 |
|---|---|---|
| `ExKei` | 出荷日情報!$A$2:$D$1000 | 出荷日情報テーブル（第1グループ） |
| `Hyouji` | 完了予定時間!$C$5:$J$50 | 照射データ表示領域（画面クリア対象） |
| `Jyoutai` | 完了予定時間!$F$2:$G$3 | 線源状態・更新日時表示エリア（画面クリア対象） |
| `SenngennTB` | 線源情報!$A$2:$D$145 | 線源情報テーブル（画面クリア対象） |
| `SortTB` | 完了予定時間!$D$5:$K$50 | ソート対象テーブル（完了時間でソート） |
| `SyouTB` | 照射ﾃﾞｰﾀ!$A$2:$U$26 | 照射データテーブル（画面クリア対象） |
| `SyukkabiTB` | 出荷日情報!$A$2:$D$1000 | 出荷日 VLOOKUP 参照テーブル |
| `DataInp` | #REF! | 参照エラー（未使用・廃止済み） |
| `DebugFlg` | #REF! | 参照エラー（未使用・廃止済み） |
| `InpTbl` | #REF! | 参照エラー（未使用・廃止済み） |
| `Misyousya` | #REF! | 参照エラー（未使用・廃止済み） |
| `NowTime` | #REF! | 参照エラー（未使用・廃止済み） |
| `Nuru` | #REF! | 参照エラー（未使用・廃止済み） |
| `SetTbl` | #REF! | 参照エラー（未使用・廃止済み） |
| `TyuudannInp` | #REF! | 参照エラー（未使用・廃止済み） |

---

## 4. 数式一覧

### 完了予定時間シート

| セル | 数式 | 説明 |
|---|---|---|
| H2 | `=IF(F2="移動照射中","完了予想時刻は担当者に確認のこと","")` | 線源が移動照射中の場合に完了予想時刻は担当者確認のメッセージを表示 |
| K5:K50 | `=IF(ISERROR(VLOOKUP(D5,SyukkabiTB,2,FALSE)),"",VLOOKUP(D5,SyukkabiTB,2,FALSE))` | 受付番号で SyukkabiTB を検索し出荷日を表示（エラー時は空白）。K5〜K50 の46行に同パターン |
| L5:L50 | `=IF(ISERROR(VLOOKUP(D5,SyukkabiTB,3,FALSE)),"",VLOOKUP(D5,SyukkabiTB,3,FALSE))` | 受付番号で SyukkabiTB を検索し備考を表示（エラー時は空白）。L5〜L50 の46行に同パターン |

（K5:K50 / L5:L50 は全行同一パターンのため代表1行を記載）

---

## 5. ボタン・マクロ対応

| シート | ボタンラベル | 割り当てマクロ | 動作概要 |
|---|---|---|---|
| 完了予定時間 | 最新情報 | `Kousinn` | DB から照射データ・線源情報・出荷日情報を取得し、完了までの残時間・完了予想日時を計算してシートを更新する |
| 完了予定時間 | 終了 | `照射情報終了処理` | ブックを保存せずに閉じる（ブックが1つのみなら Excel 終了） |

---

## 6. VBA モジュール仕様

### 6.1 ThisWorkbook.cls

#### `Workbook_Open()`

**処理概要**: ブックを開いたとき、シート保護を解除・UIのみ保護を再設定し、画面クリアを実行する（自動更新はコメントアウト済み）。

```vba
Private Sub Workbook_Open()
    ActiveSheet.Unprotect
    ActiveSheet.Protect UserInterfaceOnly:=True
    GamennCls
    'Call Kousinn  ' 自動更新は無効化されている
End Sub
```

#### `Workbook_BeforeClose(Cancel)`

**処理概要**: ブックを閉じる前に変更フラグを True にセットして保存確認を抑制する。

```vba
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    ThisWorkbook.Saved = True
End Sub
```

---

### 6.2 Ex抽出処理.bas

#### `Kousinn()`

**処理概要**: 照射管理DBから照射データ・線源情報・出荷日情報を一括取得し、完了予想時間を計算して「完了予定時間」シートに更新するメイン処理。

**処理フロー**:
1. `SyukabiRead` で出荷日情報を読み込む
2. 照射ﾃﾞｰﾀシートに `Yomikomi` で `syouk1` から照射データを取得
3. 照射位置コード（11〜32）を「南ｺﾝ」「南固」「北ｺﾝ」「北固」「特１」「特２」に変換
4. 線源情報シートに `SenGenn` で `SENGNR1` から最新線源状態を取得
5. 線源タイマー値・日時から完了までの残時間を計算（最小残時間品を基準に他は1.2倍補正）
6. 線源照射中の場合は経過時間を補正して残時間を更新
7. 中断中（SYOSTAT=2）の行は残時間を "-----"、完了日時を "中断中" に設定
8. 移動照射中（EVENT=3）かつ固定照射種別（SYOKIND=1）の行は "------" 表示
9. 完了予定時間シートの `Hyouji` 範囲に書き込み
10. `SortTB` 範囲を I列（完了までの時間）で昇順ソート

**照射位置コード変換**:

| コード | 表示 |
|---|---|
| 11 | 南ｺﾝ |
| 12 | 南固 |
| 21 | 北ｺﾝ |
| 22 | 北固 |
| 31 | 特１ |
| 32 | 特２ |
| その他 | ?? |

```vba
Sub Kousinn()
    ' ... 照射データ取得・変換・計算・表示処理
    Call SyukabiRead
    Call Yomikomi(mySyouSuu)
    ' 照射位置コード変換
    ' 線源情報取得
    Call SenGenn
    ' 残時間計算・完了日時計算
    ' 完了予定時間シートへの書き込み
    ' SortTB を I列で昇順ソート
End Sub
```

#### `Yomikomi(mySuu)`

**処理概要**: `syouk1` テーブルから未照射品データを「照射ﾃﾞｰﾀ」シートに取得する。

**SQL**:
```sql
SELECT UNO, SYONO, KAINAME, SITEISN, SYOSUU, SYOICHI, SYOTIME, HANSUU,
       STIMER, KTIMER, SENRITU, SYOSTAT, CTIMER, ZHANSUU, HTIMER, SLOTNO,
       SDATE, EDATE, UPDFLG, SYOKIND, BIKOU
FROM syouk1
WHERE syokind <> '2'
ORDER BY uno
```

#### `SenGenn()`

**処理概要**: `RIC.SENGNR1` テーブルから最新日付の最新時刻レコード（線源状態）を「線源情報」シートに取得する。

**SQL**:
```sql
SELECT SENGNR1.SDATE, SENGNR1.STIME, SENGNR1.TIMER, SENGNR1.EVENT
FROM RIC.SENGNR1
WHERE SENGNR1.SDATE IN (
    SELECT MAX(SENGNR1.SDATE) FROM RIC.SENGNR1
)
ORDER BY SENGNR1.STIME DESC
```

---

### 6.3 SQL_Execution.bas

`Ex単価検索.xlsm` の `SQL_Execution.bas` と同一実装。DSN `ricdb`（UID: ric / PWD: t6101）での Oracle DB 接続・SQL 実行共通ルーチン（`Open_oraconDB` / `SQL_Exe` / `SQL_INSERT_UPDATE` / `SQL_Delete` / `Disp_Sheet` / `Set_Array`）。

本ファイルでは主に `Disp_Sheet` を使用してシートへのレコードセット貼り付けを行う。

---

### 6.4 Ex画面クリア.bas

#### `GamennCls()`

**処理概要**: 完了予定時間・線源情報・照射ﾃﾞｰﾀシートの表示領域をすべてクリアする。

```vba
Sub GamennCls()
    Sheets("完了予定時間").Select
    Range("Hyouji") = ""
    Range("ExKei") = ""
    Range("SenngennTB") = ""
    Range("SyouTB") = ""
    Range("Jyoutai") = ""
End Sub
```

---

### 6.5 Ex終了照射情報.bas

#### `照射情報終了処理()`

**処理概要**: ブックを保存せずに閉じる。ブックが1冊のみなら Excel を終了する。

```vba
Sub 照射情報終了処理()
    If Workbooks.Count <= 1 Then Application.Quit
    ThisWorkbook.Close savechanges:=False
End Sub
```

---

### 6.6 Ex出荷日情報読込Ric1.bas

#### `SyukabiRead()`

**処理概要**: `ExKeikakuX` テーブルから1号機（souti='1'）の出荷日・備考・出荷方法を取得し、「出荷日情報」シートの2行目から書き込む。

**SQL**:
```sql
SELECT uno * 1, TRIM(syukkabi), bikou1, syuhouhou
FROM ExKeikakuX
WHERE souti = '1'
ORDER BY uno
```

---

### 6.7 FunctionR1.bas

#### `ExchengeDATE(myDate, myType)` 関数

**処理概要**: 日付文字列を指定フォーマットに変換する。

| myType | 入力形式 | 出力形式 |
|---|---|---|
| `"mm/dd"` | 4桁数値（MMDD） | `MM/DD` |
| `"yyyy/mm/dd"` | 8桁数値（YYYYMMDD） | `YYYY/MM/DD` |
| （数値ゼロ） | 任意 | そのまま（trim） |

#### `ExchengeDay(myDate, myType)` 関数

**処理概要**: 日付値を指定フォーマットの文字列に変換する。

| myType | 出力形式 |
|---|---|
| `"mmdd"` | `MMDD` |
| `"yyyymmdd"` | `YYYYMMDD` |

---

### 6.8 GetPathRic1Jyou.bas

#### `Ric1JyouGetPathX()`

**処理概要**: ブックと同一フォルダの `ExAprReadPath.txt` を読み込み、`mpAprMotoPath`・`mpAprSakiPath`・`mpDB` の3変数にパス情報を設定する。

```vba
Sub Ric1JyouGetPathX()
    Dim myFno As Double
    myFno = FreeFile
    Open ActiveWorkbook.Path & "\" & "ExAprReadPath.txt" For Input As #myFno
        Input #myFno, mpAprMotoPath, mpAprSakiPath, mpDB
    Close #myFno
End Sub
```

---

## 7. DB 接続・外部連携

### ODBC 接続設定

| DSN 名 | UID | PWD | 用途 |
|---|---|---|---|
| `ricdb` | ric | t6101 | Oracle DB接続（照射データ・線源情報・出荷日情報） |

### 参照テーブル一覧

| テーブル名 | 主な用途 | 主要カラム |
|---|---|---|
| `syouk1` | 1号機の照射中製品データ | UNO, KAINAME, SITEISN, SYOSUU, SYOICHI, HTIMER, SYOSTAT, SYOKIND |
| `RIC.SENGNR1` | 線源状態情報 | SDATE, STIME, TIMER, EVENT |
| `ExKeikakuX` | 出荷計画情報（出荷日・備考） | uno, syukkabi, bikou1, syuhouhou, souti |

### 外部ファイル連携

| ファイル名 | 場所 | 内容 |
|---|---|---|
| `ExAprReadPath.txt` | ブックと同一フォルダ | mpAprMotoPath, mpAprSakiPath, mpDB の3値をカンマ区切りで記録 |

---

## 8. データフロー

```
【起動フロー】
  ブックを開く（Workbook_Open）
       ↓
  GamennCls() – 全表示セルをクリア
  （Kousinn 自動呼び出しはコメントアウト済み）

【最新情報更新フロー】
  「最新情報」ボタン押下
       ↓ Kousinn() 呼び出し
  SyukabiRead()
       ↓ DB（ExKeikakuX）: souti='1' の出荷日情報を取得
       ↓ 出荷日情報シート（A2:D列）に書き込み
  Yomikomi()
       ↓ DB（syouk1）: syokind<>'2' の照射中データを取得
       ↓ 照射ﾃﾞｰﾀシート（A2:U列）に書き込み
  SenGenn()
       ↓ DB（RIC.SENGNR1）: 最新日付の最新時刻レコードを取得
       ↓ 線源情報シート（A1:D2）に書き込み
  照射位置コード変換（11→南ｺﾝ, 12→南固, 21→北ｺﾝ, 22→北固, 31→特１, 32→特２）
       ↓
  残時間計算（完了タイマー値 - 現在タイマー値）
  最小残時間を基準に他の品は1.2倍補正
  線源照射中（EVENT=1）の場合は経過時間を差し引いて補正
       ↓
  完了予定時間シート Hyouji 範囲に書き込み
       ↓
  SortTB を I列（完了までの時間）で昇順ソート
       ↓
  完了予定時間シートに表示完了

【VLOOKUP 出荷日表示フロー】
  完了予定時間シートの K5:K50 / L5:L50
       ↓ VLOOKUP(受付番号, SyukkabiTB, 2/3, FALSE)
  出荷日情報シートから出荷日・備考を取得して表示
```

---

## 9. セキュリティ注意事項

olevba 解析結果より以下の警告が検出されています。

| 種別 | キーワード | 内容 |
|---|---|---|
| AutoExec | `Workbook_Open` | ブックを開いたとき自動的に `GamennCls` が実行される |
| AutoExec | `Workbook_BeforeClose` | ブックを閉じる前に保存フラグ操作が自動実行される |
| Suspicious | `Open` | `GetPathRic1Jyou.bas` で `ExAprReadPath.txt` をファイルオープン |
| Suspicious | `Call` | Excel 4 マクロ形式の DLL 呼び出しの可能性 |
| Suspicious | `Hex Strings` | VBA 内に16進数エンコード文字列が存在 |
| Suspicious | `Base64 Strings` | VBA 内にBase64エンコード文字列が存在 |

> **注意**: DSN=ricdb の接続文字列（UID/PWD）が `SQL_Execution.bas` にハードコードされています。パスワード `t6101` が平文で記述されているため、コードの漏洩に注意してください。
> **注意**: `#REF!` の名前付き範囲が7件残存しています。削除またはシート・セルとの再紐付けが必要です。
