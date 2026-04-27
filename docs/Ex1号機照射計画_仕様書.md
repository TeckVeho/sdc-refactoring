# Ex1号機照射計画 仕様書

> **ファイル種別**: .xlsm（マクロ付き）
> **用途**: ガンマ線照射装置（1号機）における照射計画の立案・表示・記録・管理
> **VBA プロジェクトサイズ**: 19モジュール（標準モジュール14、クラスモジュール5）
> **外部連携**: Oracle DB（DSN=ricdb）、テーブル ExR1Keikaku1/ExR1Keikaku2/syouk1/zaiko/ExkeikakuX 等

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
Ex1号機照射計画.xlsm
├── シート
│   ├── 照射計画          ← メイン画面（177行×293列）
│   ├── 未計画品一覧      ← 未照射在庫一覧（204行×14列）
│   ├── 設定値            ← セット場所・表示文字の設定テーブル（26行×14列）
│   └── 使用方法          ← 操作説明（27行×4列）
└── VBA モジュール
    ├── Ex共通変数定義.bas        ← グローバル変数・定数定義
    ├── ThisWorkbook.cls          ← Workbook_Open / Workbook_BeforeClose
    ├── Sheet4.cls                ← 照射計画シートイベント（ダブルクリック・変更）
    ├── Sheet5.cls                ← 未計画品一覧シートイベント
    ├── Sheet1.cls                ← （空）
    ├── Sheet2.cls                ← （空）
    ├── Ex印刷設定.bas            ← 印刷範囲・改ページ設定
    ├── Ex最新データ表示起動.bas  ← 画面全体の再表示起動
    ├── 中断・空き設定.bas        ← 中断/空き設定・取消
    ├── Exサブルーチン.bas        ← セット場所ソート
    ├── Ex計画表示.bas            ← DBから計画データ読込・表示
    ├── Ex計画記録.bas            ← 計画データをDBへ記録
    ├── Ex最新照射中表示.bas      ← 照射中製品をDBから取得・表示
    ├── Ex計画エリア整え.bas      ← 計画エリアの罫線・書式整形
    ├── Ex未照射品.bas            ← 未照射在庫をDBから取得・表示
    ├── SQL_Execution.bas         ← ADODB接続・SQL実行共通ルーチン
    ├── ExSubFunc.bas             ← 日付変換・文字列クリーン等ユーティリティ
    ├── Ex照射予定指定削除.bas    ← 予定行削除
    └── ErrCheck.bas              ← エラーメッセージ表示
```

---

## 2. シート詳細

### 2.1 照射計画

**目的**: 1号機の照射計画を時系列グリッド（3時間単位）で管理・表示する。照射中・計画中製品の状態をリアルタイムにDB連携して可視化する。

#### レイアウト構造

| セル範囲 | 名前付き範囲 | 内容 |
|---|---|---|
| A1 | `DebugFlg` | デバッグフラグ（通常空白） |
| P2 | `MSGArea` | 照射状態メッセージ（貯蔵中/固定照射中/昇降中/移動照射中/PCスタート/PC停止中/装置停止中） |
| W2 | — | 「高速回転数(rpm)」ヘッダー |
| AB2 | `MoterRPM` | 高速回転数（rpm）表示 |
| AD2 | — | 「処理箱数(L:40cm換算)」ヘッダー |
| AI2 | `Hakosuu` | 処理箱数 |
| AK2 | — | 「処理時間(h)」ヘッダー |
| AO2 | — | 処理時間計算（数式） |
| F4〜KG4 | — | 時刻ヘッダー行（0〜216の3時間刻み、216列分） |
| F5〜KG5 | — | 列連番行（1〜293） |
| A19〜U19 | — | 製品情報列ヘッダー（No/移動・中断・空/未計画数/場所ソート用/受付番号/会社名/線量/入荷数/未照射数/納期/出荷日/備考/セット数/照射時間/セット場所/管理番号/残時間/照射/メモ） |
| V19〜KG19 | `NowTime` (V19) | 日付ヘッダー行（8列=1日単位で日付を表示） |
| V20〜KG20 | — | 時刻ヘッダー行（0,3,6,9…21 の繰り返し） |
| B21 | — | 移動照射ありなし行 |
| F21 | `Jikoku` | 時刻補正値 |
| A23〜U72 | `SotrArea`(D〜IC) | 製品行エリア（最大50行） |
| V23〜IC72 | `KeikakuAria` | 計画表示エリア（時系列グリッド、V列〜IC列） |
| E23〜KG72 | `HyoujiAria` | 全表示エリア |
| A77〜A84 | `SetIti` | セット場所管理配列 |
| B73 | `IdouUMu` | 移動照射有無フラグ行 |

（行番号 5〜18、73〜76、85〜177 は設定・非表示行）

#### 製品情報列（行19のヘッダー対応）

| 列 | ヘッダー | 用途 |
|---|---|---|
| A | No | 行番号 |
| B | 移動/中断/空有無 | DB記録有無フラグ |
| D | 未計画数 | 未照射残数 |
| E | 場所ソート用 | セット場所によるソートキー |
| F | Ful受付番号 | 完全受付番号 |
| G | 受付No | 受付番号（4桁） |
| H | 会社名 | 顧客会社名 |
| I | 線量 | 指定線量 |
| J | 入荷数 | 入荷箱数 |
| K | 未照射数 | 照射未完了数 |
| L | 納期 | 納期（mm/dd形式） |
| M | 出荷日 | 出荷予定日 |
| N | 備考 | 備考欄 |
| O | セット数 | 照射セット数 |
| P | 照射時間 | 1回の照射時間（h） |
| Q | セット場所 | 照射位置（北固/南固/北コン/南コン/特１/特２/移動） |
| R | 管理番号 | DB管理番号（keikakuno） |
| S | 残時間 | 残照射時間 |
| T | 照射 | 照射状態（中/予定/中断） |
| U | メモ | メモ欄 |

---

### 2.2 未計画品一覧

**目的**: 照射未完了の在庫品および予約品を一覧表示し、照射計画への登録を行う画面。

#### レイアウト構造

| セル範囲 | 名前付き範囲 | 内容 |
|---|---|---|
| E2 | — | 「未照射品一覧」タイトル |
| B4〜M4 | — | ヘッダー行（No./受付No/会社名/製品名/線量/入荷数/未照射数/納期/出荷日/備考） |
| C6〜M105 | `Misyousya` | データ表示エリア（最大100行） |
| E6〜E105 | `MisyoriUno` | 受付番号列 |

---

### 2.3 設定値

**目的**: セット場所コード・表示文字・照射時間等の設定値マスターテーブル。

#### レイアウト構造

| セル範囲 | 名前付き範囲 | 内容 |
|---|---|---|
| B3〜E3 | — | ヘッダー（セット場所/No:セット場所(ソート用)/DB登録値） |
| B4〜M13 | — | セット場所定義テーブル（10行） |
| C4〜D13 | `SetTbl` | セット場所変換テーブル（No→DB値） |
| B4〜C13 | `BayoTB` | 場所名→Noテーブル |
| G3〜H5 | — | 計画表示文字定義 |
| H4 | `HyoujiMoji1` | 計画中文字（`｜`） |
| H5 | `Hyoujimoji2` | 計画完了文字（`↑`） |
| I4 | — | 照射中文字 |
| J4 | `HyoujiMoji3` | 照射中文字（`｜`） |
| J5 | `HyoujiMoji4` | 照射完了文字（`↑`） |
| K4 | `IdouMoji` | 移動列表示文字（`移動`） |
| L4 | `TyuudanMoji` | 中断文字（`中断`） |
| M4 | `AkiMoji` | 空き文字（`空`） |
| H9 | `KadouJikann` | 1日の照射時間（`20`時間） |
| H6 | `KeiIdou` | 計画移動時文字（`・`） |

#### セット場所テーブル

| セット場所名 | ソートNo | DB登録値 | 照射中文字 |
|---|---|---|---|
| 北固 | 11 | 22 | 北固照射中 |
| 北コン | 12 | 21 | 北コン照射中 |
| 南固 | 21 | 12 | 南固照射中 |
| 南コン | 22 | 11 | 南コン照射中 |
| 移動 | 31 | 00 | 移動照射 |
| 特１ | 40 | 31 | 特１照射中 |
| 特2 | 41 | 32 | 特2照射中 |
| 特殊 | 42 | 40 | 特殊位置 |
| 未定 | 50 | 80 | — |
| 予約 | 99 | 99 | — |

---

### 2.4 使用方法

**目的**: 操作手順の説明シート（印刷・参照用）。

| 番号 | 操作 | 説明概要 |
|---|---|---|
| 1 | 起動 | 「最新データ」クリックで最新の照射中・計画予定・未計画品が表示される |
| 2 | 計画登録 | 未計画品一覧でダブルクリック→セット数・照射時間・場所入力→開始日時セルWクリック→「計画記録」実行 |
| 3 | 計画記録 | 中断・空・移動日時がDBに記録される |
| 4 | 印刷設定 | 印刷日数を入力→プリンター選択→印刷プレビュー |
| 5 | 予定行削除 | 削除したい予定行を選択し「予定行削除」をクリック |
| 6 | 中断設定 | 中断したい行・日時範囲を選択し「中断設定」をクリック（以降自動的に右に移動） |
| 7 | 空設定 | 空にしたい範囲を選択し「空設定」をクリック |
| 8 | 中断/空取消 | 取消したい範囲を選択し「中断/空取消」をクリック |

---

## 3. 名前付き範囲一覧

| 名前 | 参照先 | 説明 |
|---|---|---|
| `AkiMoji` | 設定値!$M$4 | 空き文字（`空`） |
| `BayoTB` | 設定値!$C$4:$D$13 | セット場所→No変換テーブル |
| `DebugFlg` | 照射計画!$A$1 | デバッグフラグ |
| `Hakosuu` | 照射計画!$AI$2 | 処理箱数 |
| `HyoujiAria` | 照射計画!$E$23:$KG$72 | 全表示エリア |
| `HyoujiMoji1` | 設定値!$H$4 | 計画中記号（`｜`） |
| `Hyoujimoji2` | 設定値!$H$5 | 計画完了記号（`↑`） |
| `HyoujiMoji3` | 設定値!$J$4 | 照射中記号（`｜`） |
| `HyoujiMoji4` | 設定値!$J$5 | 照射完了記号（`↑`） |
| `IdouMoji` | 設定値!$K$4 | 移動表示文字（`移動`） |
| `IdouUMu` | 照射計画!$B$73 | 移動照射有無フラグ行 |
| `Jikoku` | 照射計画!$F$21 | 時刻補正値 |
| `KadouJikann` | 設定値!$H$9 | 1日の照射稼働時間（20h） |
| `KeiIdou` | 設定値!$H$6 | 計画移動文字（`・`） |
| `KeikakuAria` | 照射計画!$V$23:$IC$72 | 計画グリッド表示エリア |
| `MisyoriUno` | 未計画品一覧!$E$6:$E$105 | 受付番号列 |
| `Misyousya` | 未計画品一覧!$C$6:$M$105 | 未計画品データエリア |
| `MoterRPM` | 照射計画!$AB$2 | モーター回転数 |
| `MSGArea` | 照射計画!$P$2 | 照射状態メッセージエリア |
| `NewData` | 照射計画!$C$3 | 新データフラグ |
| `NowTime` | 照射計画!$V$19 | 現在日時（`=NOW()`） |
| `SetIti` | 照射計画!$A$77:$A$84 | セット場所管理配列 |
| `SetTbl` | 設定値!$B$4:$C$13 | セット場所変換テーブル |
| `SotrArea` | 照射計画!$D$23:$IC$72 | ソート対象エリア |
| `TeisaiMoto` | 照射計画!$ID$20:$IK$72 | 体裁整形元データ |
| `TyuudanMoji` | 設定値!$L$4 | 中断文字（`中断`） |

---

## 4. 数式一覧

### 照射計画シート

| セル | 数式 | 説明 |
|---|---|---|
| AO2 | `=IF(OR(AB2="",AI2=""),"",(AI2*40+1200)/(AB2*11.324))` | 処理時間計算（箱数と回転数から算出）。`(箱数×40+1200) ÷ (回転数×11.324)` |
| V19 | `=NOW()` | 現在日時（NowTime名前付き範囲） |
| AD19 | `=INT(NowTime)+1` | 翌日日付（1日目） |
| AL19 | `=AD19+1` | 2日後日付 |
| AT19 | `=AL19+1` | 3日後（以降8列=1日単位で連続） |
| BB19〜IL19 | `=前列+1` | 翌日付の連続計算（8列間隔で日付が進む） |

> **補足**: 照射計画グリッドは1日=8列（3時間×8=24時間）で構成。V列が当日0時、W列が当日3時…という3時間刻み表示。

---

## 5. ボタン・マクロ対応

| シート | ボタンラベル | 割り当てマクロ | 動作概要 |
|---|---|---|---|
| 照射計画 | 中断/空 取消 | `中断空き取消` | 選択セル範囲の中断・空設定を取消し左にシフト |
| 照射計画 | 空 設定 | `空き設定` | 選択セル範囲に「空」を設定し右にシフト |
| 照射計画 | 中断 設定 | `中断設定` | 選択セル範囲に「中断」を設定し右にシフト |
| 照射計画 | 予定行 削除 | `予定行削除` | 選択した予定行をDBから削除しクリア |
| 照射計画 | 印刷 設定 | `Innsatu` | 印刷日数入力→改ページ設定→印刷プレビュー |
| 照射計画 | 計画 記録 | `Ex計画記録.Ex計画記録` | 画面の計画データをDBに保存 |
| 照射計画 | 最新 データ | `Ex表示起動` | DBから最新データを取得し画面を再描画 |
| 未計画品一覧 | 選択 | `計画指定` | 選択行を照射計画シートへ転送 |

---

## 6. VBA モジュール仕様

### 6.1 Ex共通変数定義.bas

グローバル変数・定数の宣言モジュール。全モジュールから参照される。

| 変数/定数 | 型 | 内容 |
|---|---|---|
| `mpPUSyousuu` | Single | 照射中データ数 |
| `mpSyouMode` | String | 照射モード（E/S/1/2/3） |
| `mpIdou(230)` | Array | 移動停止データ（移動停止該当日時は1） |
| `mpNow` | Variant | 現在日時 |
| `mpExKeikaku1` | Const | 計画表示部DBテーブル名（`"ExR1Keikaku1"`） |
| `mpExKeikaku2` | Const | 計画データ部DBテーブル名（`"ExR1Keikaku2"`） |
| `mpTyudanN` | Integer | 中断指定した数 |
| `mpTorikesiN` | Integer | 中断取消した数 |
| `mpYukou` | Const | 有効フラグ値（`'2'`） |

---

### 6.2 ThisWorkbook.cls

#### `Workbook_Open()`

**処理概要**: ブック起動時に照射計画の最新データを表示する。

**処理フロー**:
1. ブック構造保護・ウィンドウ保護を解除
2. 数式バーを表示
3. `mpTyudanN`、`mpTorikesiN` を0にリセット
4. `Ex表示起動` を呼び出しDB最新データ表示

```vba
Private Sub Workbook_Open()
    ActiveWorkbook.Protect Structure:=False, Windows:=False
    Application.DisplayFormulaBar = True
    mpTyudanN = 0
    mpTorikesiN = 0
    Call Ex表示起動
End Sub
```

#### `Workbook_BeforeClose(Cancel As Boolean)`

**処理概要**: 未記録行がある場合、終了前に確認メッセージを表示する。

**処理フロー**:
1. 未記録行または未取消の中断がある場合、確認ダイアログ表示
2. キャンセル選択時は終了を中断
3. 問題なければアラート無効化・保存済みフラグをセットして終了

```vba
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    If mpTyudanN <> mpTorikesiN Or Kiroku > 0 Then
        mpErrDes = "記録されていない行があります" & Chr(13) & "このまま終了してよいですが"
        If MsgBox(mpErrDes, vbYesNo) = vbNo Then
            Cancel = True
            Exit Sub
        End If
    End If
    Application.DisplayAlerts = False
    ActiveWorkbook.Saved = True
End Sub
```

---

### 6.3 Ex最新データ表示起動.bas

#### `Ex表示起動()`

**処理概要**: 照射計画シートの全データをDBから最新状態に更新・表示する。

**処理フロー**:
1. 未記録行確認・ユーザー確認
2. カウンターリセット
3. 計算・イベント・画面更新を一時停止（高速化）
4. シート保護設定
5. `ExAriaCLS` → 表示エリアクリア
6. `Ex未照射品表示` → 未照射在庫をDBから読込
7. `Ex照射中製品表示` → 照射中製品をDBから読込
8. `Ex計画表示Main` → 計画データをDBから読込
9. `MSGArea` に照射状態を表示
10. `計画画面整形`、`SetBasyoSort` で表示整理
11. 計算・イベント・画面更新を再開

---

### 6.4 中断・空き設定.bas

#### `中断設定()` / `空き設定()`

**処理概要**: 選択行に中断または空き文字を挿入し、後続の計画を右にシフトする。

**処理フロー**:
1. アクティブセル行・列・選択列数を取得
2. 範囲バリデーション（行23〜72、列20〜293）
3. 選択範囲に中断/空文字を挿入（右シフト）
4. 後続の「空」文字セルを入れ替え（空が詰まらないよう整理）
5. `計画画面整形` で再描画

#### `中断空き取消()`

**処理概要**: 選択した中断/空セルを削除し後続データを左にシフトする。

---

### 6.5 Ex計画表示.bas

#### `Ex計画表示Main()`

**処理概要**: DBから計画データを読み込み照射計画シートに表示する。

**処理フロー**:
1. `Ex計画データ表示` → ExR1Keikaku1からSQLで計画行を読み込む
2. `照射中の中断` → ExR1Keikaku2から中断データを読み込み挿入
3. `移動照射反映` → 移動照射中の補正処理
4. `計画画面整形` → 罫線・書式の再設定

#### `Ex計画データ表示()`

DB `ExR1Keikaku1` から以下のSQLを実行し照射計画シートに書き込む：

```sql
SELECT UNO, SUBSTR(uno,7,4), KAINAME, SITEISN, NYUKASU, MISYOUSU,
       trim(NOUKI), trim(SYUKKABI), BIKOU, SETSUU, SYOZI, TRIM(SYOITI),
       KEIKAKUNO, ZANJIKANN, SYOUJYO, MENO
FROM ExR1Keikaku1
WHERE keikakuno > 1000
  AND keieday >= <現在日時>
  AND kakuninn = '2'
  AND kainame IS NOT NULL
ORDER BY keikakuno
```

---

### 6.6 Ex計画記録.bas

#### `Ex計画記録()`

**処理概要**: 画面上の照射計画データ（照射中・予定行）をDBに保存する。

**処理フロー**:
1. `ExR1Keikaku1`、`ExR1Keikaku2` の有効データをDELETE
2. 行23〜72をループ（移動/中断/空有無フラグ>0の行のみ）
3. 照射中行（T列="中"）: 計画記号列をExR1Keikaku2にINSERT/UPDATE
4. 予定行（T列="予定"）: 計画記号列をExR1Keikaku2にINSERT/UPDATE
5. 各行をExR1Keikaku1にINSERT/UPDATE（管理番号、受付番号、会社名、線量、箱数、セット数、照射時間、セット場所、計画開始/終了日時 等18フィールド）
6. 「記録終了しました」メッセージ表示

---

### 6.7 Ex最新照射中表示.bas

#### `Ex照射中製品表示()`

**処理概要**: 照射中製品の残照射時間を計算し照射計画シートに表示する。

**処理フロー**:
1. 場所テーブル(`BayoTB`)読込
2. `syouk1`、`zaiko`、`ExkeikakuX`、`ExR1Keikaku1`をJOINしてSQLで照射中製品取得
3. `sengnr1` テーブルから直近の照射開始イベント取得
4. 残照射時間を計算（固定照射中/移動停止中の差分を補正）
5. 残時間0以下の製品を除外
6. 照射計画シートに表示、計画線（`｜`、`↑`）をグリッドに記入

---

### 6.8 SQL_Execution.bas

#### `Open_oraconDB()`

**処理概要**: Oracle DBへADODB接続を開く。

```vba
oraconn.ConnectionString = "DSN=ricdb;UID=ric;PWD=t6101"
oraconn.Open
```

#### `SQL_INSERT_UPDATE(myTBL, myKey, myD(), myN)`

**処理概要**: 指定テーブルに対し、キー一致ならUPDATE、なければINSERTを実行する汎用メソッド。

#### `Disp_Sheet(mySQL, mySH, myRow, …)`

**処理概要**: SQL実行結果をシート指定セルへ貼り付ける汎用表示ルーチン。`CopyFromRecordset` を使用。

#### `Set_Array(mySQL, myData(), …)`

**処理概要**: SQL実行結果を2次元配列（レコード数×フィールド数）に格納する汎用ルーチン。

---

### 6.9 ExSubFunc.bas

| 関数 | 引数 | 説明 |
|---|---|---|
| `ExchengeDATE(myUno, myDate, myType)` | 受付番号, 日付, 形式 | DB日付（mmdd形式）を mm/dd または yyyy/mm/dd に変換 |
| `ExchengeDay(myDate, myType)` | 日付, 形式 | Excel日付値をmmdd/yyyymmdd文字列に変換 |
| `Kiroku()` | — | 未記録予定行が存在するか確認。存在する行Noを返す |
| `CleanChar(strData)` | 文字列 | 制御コード除去（漢字のAsc<0は保持） |

---

## 7. DB 接続・外部連携

### ODBC 接続設定

| DSN 名 | UID | PWD | 用途 |
|---|---|---|---|
| `ricdb` | ric | t6101 | 照射管理DBメイン接続 |

> **注意**: 接続文字列がVBAコードに平文で記載されている（セキュリティ注意事項参照）。

### 参照テーブル一覧

| テーブル名 | 主な用途 | 主要カラム |
|---|---|---|
| `ExR1Keikaku1` | 照射計画ヘッダー（計画表示部） | keikakuno, uno, kainame, siteisn, nyukasu, misyousu, nouki, syukkabi, bikou, setsuu, syozi, syoiti, keikakuno, zanjikann, syoujyo, meno, keisday, keieday, kakuninn |
| `ExR1Keikaku2` | 照射計画詳細（日時・記号列） | keikakuno, edano, keiday, keihr, keikaku, yuukou |
| `syouk1` | 照射中製品情報 | uno, kainame, siteisn, syosuu, syotime, syoichi, slotno, htimer, syokind, syostat |
| `zaiko` | 在庫情報 | uno, kainame, siteisn, nyukasu, misyousu, nouki, syouso |
| `ExkeikakuX` | 照射予定情報 | uno, syukkabi, bikou1 |
| `sehmst` | 製品マスター | kaisyacd, sehncd, seiname |
| `sengnr1` | 線源利用データ | event, sdate, stime, timer |
| `ExYoyakuX` | 予約データ | yoyakuno, uno, kainame, siteisn, nyukasu, misyousu, nouki, syukkabi, bikou, souti, yuukou |

### 主要 SQL 文

```sql
-- 照射計画データ取得（ExR1Keikaku1）
SELECT UNO, SUBSTR(uno,7,4), KAINAME, SITEISN, NYUKASU, MISYOUSU,
       trim(NOUKI), trim(SYUKKABI), BIKOU, SETSUU, SYOZI, TRIM(SYOITI),
       KEIKAKUNO, ZANJIKANN, SYOUJYO, MENO
FROM ExR1Keikaku1
WHERE keikakuno > 1000
  AND keieday >= :現在日時
  AND kakuninn = '2'
  AND kainame IS NOT NULL
ORDER BY keikakuno
```

```sql
-- 照射中製品取得
SELECT s.uno, SUBSTR(s.uno,7), s.kainame,
       TO_NUMBER(s.siteisn), TO_NUMBER(z.nyukasu), TO_NUMBER(z.misyousu),
       z.nouki, e.syukkabi, e.bikou1,
       TO_NUMBER(s.syosuu), TO_NUMBER(s.syotime), s.syoichi, s.slotno,
       s.htimer, s.syokind, s.syostat, k.meno
FROM syouk1 s, zaiko z, ExkeikakuX e, ExR1Keikaku1 k
WHERE s.uno = z.uno
  AND s.uno = e.uno(+)
  AND s.syokind <> '2'
  AND TO_NUMBER(s.slotno) = k.keikakuno(+)
ORDER BY s.uno
```

```sql
-- 直近の照射開始イベント取得
SELECT event, sdate, stime, timer FROM sengnr1
WHERE (event='1' OR event='3')
  AND sdate >= :2日前
ORDER BY sdate DESC, stime DESC
```

```sql
-- 未照射在庫取得
SELECT z.uno, SUBSTR(z.uno,7,4), REPLACE(TRIM(z.kainame),'株式会社',''),
       TRIM(s.seiname), TO_NUMBER(z.siteisn), TO_NUMBER(z.nyukasu),
       TO_NUMBER(z.misyousu), z.nouki, TRIM(e.syukkabi), e.bikou1
FROM zaiko z, ExkeikakuX e, sehmst s
WHERE z.uno = e.uno(+)
  AND z.misyousu*1 > 0
  AND z.syouso = '1'
  AND z.kaisyacd = s.kaisyacd
  AND z.sehncd = s.sehncd
ORDER BY z.uno
```

---

## 8. データフロー

```
【起動・最新データ表示フロー】
  ブック起動 (Workbook_Open)
       ↓
  Ex表示起動 呼び出し
       ↓
  ExAriaCLS: 表示エリアをクリア
       ↓
  Ex未照射品表示: zaiko/ExkeikakuX/sehmstからSELECT → 未計画品一覧シートに書込み
       ↓
  Ex照射中製品表示: syouk1/zaiko/ExkeikakuX/ExR1Keikaku1からSELECT
       → 残照射時間計算 → 照射計画シートに照射中行・計画線を書込み
       ↓
  Ex計画表示Main: ExR1Keikaku1/ExR1Keikaku2からSELECT
       → 計画行・中断データを書込み
       ↓
  計画画面整形: 罫線・条件付き書式・ソートを適用
       ↓
  照射計画シート表示完了

【計画記録フロー】
  「計画記録」ボタンクリック
       ↓
  ExR1Keikaku1/ExR1Keikaku2の有効データをDELETE
       ↓
  照射計画シート 行23〜72 をループ
    ├─ 照射中行（T列="中"）→ ExR1Keikaku2にINSERT/UPDATE（中断/空の日時・記号）
    └─ 予定行（T列="予定"）→ ExR1Keikaku2にINSERT/UPDATE（計画日時・記号）
       ↓
  ExR1Keikaku1にINSERT/UPDATE（各行ヘッダー情報・管理番号）
       ↓
  「記録終了しました」表示

【計画登録フロー（未計画品→照射計画）】
  未計画品一覧シートで行をダブルクリック or「選択」ボタン
       ↓
  計画指定: 選択行データ（受付番号/会社名/線量/入荷数 等）を照射計画シートに転送
       ↓
  照射計画シートで セット数・照射時間・セット場所 を手動入力
       ↓
  開始日時のグリッドセルをダブルクリック（Sheet4.Worksheet_BeforeDoubleClick）
       ↓
  記号記入: 照射時間に応じた「｜」「↑」記号をグリッドに記入
       ↓
  「計画記録」で保存
```

---

## 9. セキュリティ注意事項

| 種別 | キーワード | 内容 |
|---|---|---|
| AutoExec | `Workbook_Open` | ブック起動時に自動実行。DB接続・データ取得が行われる |
| AutoExec | `Workbook_BeforeClose` | ブック終了時に自動実行。未記録チェックが行われる |
| AutoExec | `Worksheet_Change` | シート変更時に自動実行（セット場所変更でソートキー更新） |
| Suspicious | `Open` | ファイルオープン操作が含まれる可能性 |
| Suspicious | `Call` | DLLをExcel 4 Macros経由で呼び出す可能性 |
| Suspicious | `Windows` | アプリケーションウィンドウ列挙の可能性 |
| Suspicious | `Chr` | 特定文字列の難読化に使用されている可能性（`Chr(13)` = 改行は正常用途） |
| Suspicious | Hex Strings | 16進数エンコード文字列が検出された |
| Suspicious | Base64 Strings | Base64エンコード文字列が検出された（`0CN` → `MENO`） |

> **重大な注意**: VBAコード中にDB接続パスワードが平文で記載されている。
> ```
> oraconn.ConnectionString = "DSN=ricdb;UID=ric;PWD=t6101"
> ```
> 本番環境では接続情報を外部設定ファイルや環境変数から取得する仕組みへの変更を強く推奨する。
