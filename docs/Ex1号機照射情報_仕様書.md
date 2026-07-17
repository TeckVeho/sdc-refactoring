# Ex1号機照射情報.xlsm 仕様書

> **ファイルパス**: `xlms/Ex1号機照射情報.xlsm`
> **種別**: OpenXML マクロ有効ブック (.xlsm)
> **用途**: 1号機固定照射の照射状況モニタリング — 照射管理データベース (SYOUK1) から未照射品の照射データを取得し、完了予想時刻・出荷日情報を一覧表示するダッシュボード
> **VBA**: あり（ユーザーフォームなし）
> **本書の目的**: 保守・改修時のリファレンスドキュメント

---


## 凡例（本仕様書の表記ルール）

本仕様書では、保守時の判別を容易にするため、以下の表記ルールを使用します。

| 種別 | 表記 | 例 |
|---|---|---|
| モジュール（.bas / .cls） | **太字** | **スタート処理.bas** |
| ユーザーフォーム（.frm） | **太字** | **SoutiSenntaku.frm** |
| プロシージャ / イベント | `コード体()` | `生産情報開始処理()` |
| シート名 | 「」 | 「未処理品一覧」 |
| セル参照 | `コード体` | `$W$3` |
| 名前付き範囲 | `コード体` | `Amari` |
| DB テーブル / カラム | `コード体` | `ExKeikakuX` / `kakunin` |
| ユーザー操作 | （操作名） | （処理方法実行 Click） |
| 主要マーク | ✓ | ✓ = 保守時に最初に確認すべき項目 |

### データフロー 場所マーク（9章）

9章のデータフロー（テーブル・ツリー図）では、処理が行われる場所を以下のアイコンで区別します。

| アイコン | 種別 | 意味 |
|---|---|---|
| 📊 | シート操作 | ワークシート上のセル書込み・読取り・表示変更 |
| 🖥️ | 画面操作 | ユーザーフォーム（.frm）の表示・入力・操作 |
| 🗄️ | DB操作 | Oracle DB への SELECT / INSERT / UPDATE / DELETE |
| 📄 | VBA内部処理 | 変数計算・条件分岐など、画面・シートに直接関与しない処理 |

### ✓（主要マーク）の判定基準

✓ は **保守時に最初に確認すべき項目** を示します。
判定基準は対象の種類ごとに以下のとおりです。

| 章 | 対象 | ✓ の判定基準 |
|---|---|---|
| 1.1 | シート | ユーザーが直接操作する、または VBA が動的に表示/非表示を切り替える |
| 1.2 | ユーザーフォーム | ユーザー入力を受け付ける、または業務フローの起点となる |
| 1.3 / 6.0 | VBA モジュール | ① ユーザー操作の起点 ② DB I/O を含む ③ 他モジュールから呼び出される ④ コード行数上位 25%　のいずれか |
| 2 | セル / 名前付き範囲 | VBA から `Range()` で代入または参照され、業務ロジックに直結する |
| 3 | 名前付き範囲 | VBA から `Range()` で代入または参照され、業務ロジックに直結する |
| 5 | ボタン / コントロール | DB 更新・画面遷移・計算実行など副作用のある操作を起動する |
| 6.0（全プロシージャ） | プロシージャ | ① ユーザー操作の起点（Click イベント等） ② DB I/O を実行 ③ 他モジュールから呼び出される Public　のいずれか |
| 7 | フォームコントロール | ユーザー入力を受け付ける、またはイベントで業務処理を起動する |
| 8.2 | DB テーブル | INSERT / UPDATE / DELETE の対象（参照のみのテーブルは ✓ なし） |

---

## 目次

1. [ファイル構成](#1-ファイル構成)
2. [シート詳細](#2-シート詳細)
3. [名前付き範囲一覧](#3-名前付き範囲一覧)
4. [数式一覧](#4-数式一覧)
5. [ボタン・マクロ対応](#5-ボタンマクロ対応)
6. [VBAモジュール仕様](#6-vbaモジュール仕様)
7. [ユーザーフォーム仕様](#7-ユーザーフォーム仕様)
8. [DB接続・外部連携](#8-db接続外部連携)
9. [データフロー](#9-データフロー)
10. [セキュリティ注意事項](#10-セキュリティ注意事項)

---

## 1. ファイル構成

### 1.1 シート一覧

| ✓ | # | シート名 | タブ名 | 用途 | 最大行 | 最大列 |
|---|---|----------|--------|------|--------|--------|
| ✓ | 1 | 「完了予定時間」 | 完了予定時間 | メイン画面 — 照射中製品の完了予想時刻を一覧表示 | 50 | O (15) |
|   | 2 | 「照射データ」 | 照射データ | DB (SYOUK1) から取得した照射データの中間格納 | 1 | U (21) |
|   | 3 | 「線源情報」 | 線源情報 | DB (SENGNR1) から取得した線源タイマー情報の中間格納 | 1 | D (4) |
|   | 4 | 「出荷日情報」 | 出荷日情報 | DB (ExKeikakuX) から取得した出荷日マスタ | 4491 | J (10) |

### 1.2 フォーム一覧

本ブックにユーザーフォーム (UserForm) は存在しない。

### 1.3 モジュール一覧

| ✓ | モジュール名 | 種別 | OLE ストリーム | 用途 |
|---|-------------|------|---------------|------|
| ✓ | **ThisWorkbook** | .cls | VBA/ThisWorkbook | ブック Open / Close イベント処理 |
|   | **Sheet1** | .cls | VBA/Sheet1 | 「出荷日情報」シートクラス（コード空） |
|   | **Sheet2** | .cls | VBA/Sheet2 | 「照射データ」シートクラス（コード空） |
|   | **Sheet4** | .cls | VBA/Sheet4 | 「完了予定時間」シートクラス（コード空） |
|   | **Sheet5** | .cls | VBA/Sheet5 | 「線源情報」シートクラス（コード空） |
| ✓ | **Ex抽出処理** | .bas | VBA/Ex抽出処理 | メインロジック — DB から照射データ・線源情報を取得し完了予想時刻を算出 |
| ✓ | **SQL_Execution** | .bas | VBA/SQL_Execution | ADO+ODBC によるDB接続・SQL実行ユーティリティ |
|   | **Ex画面クリア** | .bas | VBA/Ex画面クリア | 画面初期化処理 |
|   | **Ex終了照射情報** | .bas | VBA/Ex終了照射情報 | ブック終了処理 |
|   | **Ex出荷日情報読込Ric1** | .bas | VBA/Ex出荷日情報読込Ric1 | 出荷日情報の DB 読込 |
|   | **FunctionR1** | .bas | VBA/FunctionR1 | 日付変換ユーティリティ関数 |
|   | **GetPathRic1Jyou** | .bas | VBA/GetPathRic1Jyou | 外部パス設定ファイル読込 |

---

## 2. シート詳細

### 2.0 可視性一覧

| シート名 | 可視性 |
|----------|--------|
| 「完了予定時間」 | Visible |
| 「照射データ」 | Visible |
| 「線源情報」 | Visible |
| 「出荷日情報」 | Visible |

VBA による動的な Visible 制御コードは存在しない。

### 2.1 「完了予定時間」シート

メイン画面。1号機固定照射中の製品一覧と完了予想時刻を表示する。

#### レイアウト構造

| 行 | 列 | 内容 |
|----|----|------|
| A1 | A | `ExM` — EXメニュー識別子 |
| D2 | D | タイトル「1号機固定照射情報」 |
| F2:G2 | F–G | 線源状態（照射中 / 貯蔵中 / 昇降中 / 移動照射中） |
| H2:J2 | H–J | 移動照射中の場合の注意メッセージ（数式） |
| F3:G3 | F–G | 現在時刻 |
| 4行目 | C–L | ヘッダー行 |
| 5–50行 | C–L | データ行（最大46製品分） |

#### ヘッダー列定義（4行目）

| ✓ | 列 | ヘッダー | 内容 |
|---|----|----------|------|
| ✓ | C | No | 連番 |
| ✓ | D | 受付番号 | 照射対象の受付番号 |
| ✓ | E | 会社名 | 顧客名 |
| ✓ | F | 指定線量 | kGy 単位 |
| ✓ | G | 数量 | 照射数量 |
| ✓ | H | 照射位置 | 南コン / 南固 / 北コン / 北固 / 特１ / 特２ |
| ✓ | I | 完了までの時間 | 残時間（時間） |
| ✓ | J | 完了予想日時 | 日時形式 |
|   | K | 出荷日 | VLOOKUPで出荷日情報から取得 |
|   | L | 備考 | VLOOKUPで出荷日情報から取得 |

#### セル結合

| 範囲 | 用途 |
|------|------|
| F2:G2 | 線源状態表示 |
| F3:G3 | 現在時刻表示 |
| H2:J2 | 移動照射注意メッセージ |

### 2.2 「照射データ」シート

DB テーブル `SYOUK1` から取得した照射管理データの中間格納シート。VBA の `Yomikomi()` により書き込まれる。

#### カラム定義（1行目ヘッダー）

| ✓ | 列 | フィールド名 | 内容 |
|---|----|-------------|------|
| ✓ | A | UNO | 受付番号 |
|   | B | SYONO | 照射番号 |
| ✓ | C | KAINAME | 会社名 |
| ✓ | D | SITEISN | 指定線量 |
| ✓ | E | SYOSUU | 照射数量 |
| ✓ | F | SYOICHI | 照射位置 |
| ✓ | G | SYOTIME | 照射時間 |
|   | H | HANSUU | 反転回数 |
|   | I | STIMER | 開始タイマー |
|   | J | KTIMER | 完了タイマー |
|   | K | SENRITU | 線量率 |
| ✓ | L | SYOSTAT | 照射状態（1:開始, 2:中断, 3:完了, 4:取消, 5:再開, 6:修正） |
|   | M | CTIMER | 現在タイマー |
|   | N | ZHANSUU | 残反転回数 |
| ✓ | O | HTIMER | 完了タイマー値 |
|   | P | SLOTNO | スロット番号 |
|   | Q | SDATE | 開始日 |
|   | R | EDATE | 終了日 |
|   | S | UPDFLG | 更新フラグ |
| ✓ | T | SYOKIND | 照射種別（2号機除外フィルタに使用） |
|   | U | BIKOU | 備考 |

### 2.3 「線源情報」シート

DB テーブル `SENGNR1` から取得した線源のタイマー情報の中間格納シート。VBA の `SenGenn()` により書き込まれる。

#### カラム定義（1行目ヘッダー）

| ✓ | 列 | フィールド名 | 内容 |
|---|----|-------------|------|
| ✓ | A | SDATE | 日付 (yyyymmdd) |
| ✓ | B | STIME | 時刻 (hhmmss) |
| ✓ | C | TIMER | タイマー値 |
| ✓ | D | EVENT | イベント種別（0:昇降中, 1:照射中, 2:貯蔵中, 3:移動照射） |

### 2.4 「出荷日情報」シート

DB テーブル `ExKeikakuX` から取得した出荷日マスタ。VBA の `SyukabiRead()` により書き込まれる。

#### カラム定義（1行目ヘッダー）

| ✓ | 列 | フィールド名 | 内容 |
|---|----|-------------|------|
| ✓ | A | 受付番号 | 受付番号（数値×1 変換） |
| ✓ | B | 出荷日 | 出荷予定日 |
|   | C | 備考 | 備考欄 |
| ✓ | D | 出荷方法 | 出荷方法（混載便 等） |
|   | G | 受付番号 | （重複列 — 用途不明） |
|   | H | 出荷日 | （重複列） |
|   | I | 備考 | （重複列） |
|   | J | 出荷方法 | （重複列） |

データ行は A1001 行以降に実データが格納されている（約3,491件分）。

---

## 3. 名前付き範囲一覧

| ✓ | 名前 | 参照先 | 用途 |
|---|------|--------|------|
| ✓ | `Hyouji` | `完了予定時間!$C$5:$J$50` | メイン一覧表示エリア — 画面クリア対象 |
| ✓ | `SortTB` | `完了予定時間!$D$5:$K$50` | ソート対象範囲（完了予想日時 I列 昇順） |
| ✓ | `Jyoutai` | `完了予定時間!$F$2:$G$3` | 線源状態・現在時刻表示エリア |
| ✓ | `SyouTB` | `照射ﾃﾞｰﾀ!$A$2:$U$26` | 照射データ格納範囲 — 画面クリア対象 |
| ✓ | `SenngennTB` | `線源情報!$A$2:$D$145` | 線源情報格納範囲 — 画面クリア対象 |
| ✓ | `ExKei` | `出荷日情報!$A$2:$D$1000` | 出荷日情報格納範囲 — 画面クリア対象 |
| ✓ | `SyukkabiTB` | `出荷日情報!$A$2:$D$1000` | VLOOKUP 検索範囲（出荷日・備考取得用） |
|   | `DataInp` | `#REF!` | 無効参照（未使用） |
|   | `DebugFlg` | `#REF!` | 無効参照（未使用） |
|   | `InpTbl` | `#REF!` | 無効参照（未使用） |
|   | `Misyousya` | `#REF!` | 無効参照（未使用） |
|   | `NowTime` | `#REF!` | 無効参照（未使用） |
|   | `Nuru` | `#REF!` | 無効参照（未使用） |
|   | `SetTbl` | `#REF!` | 無効参照（未使用） |
|   | `TyuudannInp` | `#REF!` | 無効参照（未使用） |

`#REF!` の名前付き範囲は、過去にシートやセル範囲が削除された際に残った参照切れ。現在のコードでは使用されていない。

---

## 4. 数式一覧

数式は「完了予定時間」シートのみに存在する。全93件。

| ✓ | セル | 数式 | 用途 |
|---|------|------|------|
| ✓ | H2 | `=IF(F2="移動照射中","完了予想時刻は担当者に確認のこと","")` | 移動照射時の注意表示 |
| ✓ | K5 | `=IF(ISERROR(VLOOKUP(D5,SyukkabiTB,2,FALSE)),"",VLOOKUP(D5,SyukkabiTB,2,FALSE))` | 受付番号から出荷日を取得 |
| ✓ | L5 | `=IF(ISERROR(VLOOKUP(D5,SyukkabiTB,3,FALSE)),"",VLOOKUP(D5,SyukkabiTB,3,FALSE))` | 受付番号から備考を取得 |
|   | K6–K50 | K5 と同パターン（行番号のみ変化） | 出荷日取得（6–50行分） |
|   | L6–L50 | L5 と同パターン（行番号のみ変化） | 備考取得（6–50行分） |

**数式パターン**: K列は `SyukkabiTB` の第2列（出荷日）、L列は第3列（備考）を VLOOKUP で完全一致検索する。`ISERROR` でデータ不在時は空文字を返す。

---

## 5. ボタン・マクロ対応

### 5.1 シート上ボタン

「完了予定時間」シート上に2つの VML ボタンが配置されている。

| ✓ | ボタン名 | マクロ | 用途 |
|---|----------|--------|------|
| ✓ | 「最新情報」 | `[0]!Kousinn` | DB から最新の照射データ・線源情報・出荷日情報を取得し画面を更新 |
| ✓ | 「終了」 | `[0]!照射情報終了処理` | ブックを閉じる（保存なし） |

### 5.2 フォームボタン

ユーザーフォームが存在しないため、フォームボタンなし。

### 5.3 CommandBar

本ブックに CommandBar のカスタマイズコードは存在しない。

---

## 6. VBAモジュール仕様

### 6.0 全プロシージャ一覧

| ✓ | モジュール | プロシージャ名 | 種別 | 用途 |
|---|-----------|---------------|------|------|
| ✓ | **ThisWorkbook** | `Workbook_BeforeClose()` | Event | ブック閉じる前に保存済みフラグを設定 |
| ✓ | **ThisWorkbook** | `Workbook_Open()` | Event | ブック起動時にシート保護解除→再設定・画面クリア実行 |
|   | **Sheet1** | （空） | — | コードなし |
|   | **Sheet2** | （空） | — | コードなし |
|   | **Sheet4** | （空） | — | コードなし |
|   | **Sheet5** | （空） | — | コードなし |
| ✓ | **Ex抽出処理** | `Kousinn()` | Sub | メイン処理 — 出荷日読込→照射データ取得→線源情報取得→完了予想時刻計算→画面表示 |
| ✓ | **Ex抽出処理** | `Yomikomi()` | Sub | SYOUK1 テーブルから照射データを取得し「照射データ」シートに書込 |
| ✓ | **Ex抽出処理** | `SenGenn()` | Sub | SENGNR1 テーブルから直近の線源タイマー情報を取得し「線源情報」シートに書込 |
| ✓ | **SQL_Execution** | `Open_oraconDB()` | Sub | ADO+ODBC による Oracle DB 接続確立 |
| ✓ | **SQL_Execution** | `SQL_Exe()` | Sub | SQL 文を Execute で実行しレコードセットを取得 |
|   | **SQL_Execution** | `SQL_INSERT_UPDATE()` | Sub | INSERT/UPDATE の汎用実行（本ブックでは未使用） |
|   | **SQL_Execution** | `SQL_Delete()` | Sub | DELETE の汎用実行（本ブックでは未使用） |
| ✓ | **SQL_Execution** | `Disp_Sheet()` | Sub | SQL 結果をシートに直接書込み（CopyFromRecordset） |
| ✓ | **SQL_Execution** | `Set_Array()` | Sub | SQL 結果を配列に格納 |
| ✓ | **Ex画面クリア** | `GamennCls()` | Sub | 画面クリア — 名前付き範囲 Hyouji, ExKei, SenngennTB, SyouTB, Jyoutai を空白化 |
| ✓ | **Ex終了照射情報** | `照射情報終了処理()` | Sub | ブック終了 — 最後のブックなら Application.Quit |
| ✓ | **Ex出荷日情報読込Ric1** | `SyukabiRead()` | Sub | ExKeikakuX テーブルから出荷日情報を取得し「出荷日情報」シートに書込 |
| ✓ | **FunctionR1** | `ExchengeDATE()` | Function | 数値日付を "mm/dd" or "yyyy/mm/dd" 形式に変換 |
| ✓ | **FunctionR1** | `ExchengeDay()` | Function | 日付値を "mmdd" or "yyyymmdd" 数値文字列に変換 |
| ✓ | **GetPathRic1Jyou** | `Ric1JyouGetPathX()` | Sub | ExAprReadPath.txt からパス情報（mpAprMotoPath, mpAprSakiPath, mpDB）を読込 |

### 6.1 **ThisWorkbook** (.cls)

#### `Workbook_BeforeClose(Cancel As Boolean)` — Event

```vb
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    ThisWorkbook.Saved = True
End Sub
```

保存ダイアログを抑止して閉じる。

#### `Workbook_Open()` — Event

```vb
Private Sub Workbook_Open()
    ActiveSheet.Unprotect
    ActiveSheet.Protect UserInterfaceOnly:=True
    GamennCls
End Sub
```

1. アクティブシートの保護を解除
2. UserInterfaceOnly モードでシート保護を再設定（VBA からのセル操作を許可）
3. `GamennCls()` を呼び出して画面クリア

コメントアウトされた `Call Kousinn` があるが、現在は起動時の自動更新は無効化されている。

### 6.2 **Ex抽出処理** (.bas)

本ブックの中核ロジック。

#### `Kousinn()` — Sub (メイン処理)

ショートカットキー: なし（VB_Invoke_Func = " \n14"）

**処理フロー**:

1. シート保護を解除
2. `SyukabiRead()` で出荷日情報を DB から読込
3. 画面更新を停止
4. 「照射データ」シートへ移動
5. `Yomikomi()` で照射データを DB (SYOUK1) から取得
6. 照射データ件数を取得し配列 `mySyo()` に格納
7. 照射位置コードを日本語ラベルに変換:
   - `11` → 南コン, `12` → 南固, `21` → 北コン, `22` → 北固, `31` → 特１, `32` → 特２
8. 現在時刻を取得
9. `SenGenn()` で線源情報を DB (SENGNR1) から取得
10. 線源タイマー日付・時刻を Date/Time 型に変換
11. 残時間計算（完了タイマー値 − 現在タイマー値）
12. 最小残時間を検出
13. 完了予想日時を計算:
    - 最小残時間の製品: 線源日時 + 残時間 / 24
    - その他: 線源日時 + 残時間 / 24 × 1.2（1日20時間照射の補正）
14. 照射中（EVENT=1）の場合、経過時間を差し引いて残時間を補正
15. 中断中（SYOSTAT=2）の場合は `-----` / `中断中` を表示
16. 移動照射中（EVENT=3）かつ固定照射中製品（SYOKIND=1）は完了予想を `------` 表示
17. 完了予想日時 (I5列) の昇順でソート
18. シート保護を再設定
19. 画面更新を再開

#### `Yomikomi(mySuu As Single)` — Sub

照射データを DB から取得する。

```sql
SELECT UNO,SYONO,KAINAME,SITEISN,SYOSUU,SYOICHI,SYOTIME,HANSUU,
       STIMER,KTIMER,SENRITU,SYOSTAT,CTIMER,ZHANSUU,HTIMER,SLOTNO,
       SDATE,EDATE,UPDFLG,SYOKIND,BIKOU
FROM syouk1 WHERE syokind<>'2' ORDER BY uno
```

- WHERE 条件: `syokind<>'2'`（2号機データを除外）
- `Disp_Sheet()` 経由で「照射データ」シートに書込

#### `SenGenn()` — Sub

直近の線源タイマー情報を DB から取得する。

```sql
SELECT SENGNR1.SDATE, SENGNR1.STIME, SENGNR1.TIMER, SENGNR1.EVENT
FROM RIC.SENGNR1 SENGNR1
WHERE SENGNR1.SDATE IN (SELECT MAX(SENGNR1.SDATE) FROM RIC.SENGNR1)
ORDER BY SENGNR1.STIME DESC
```

- 最新日付のレコードを時刻降順で取得（先頭行が最新）
- `Disp_Sheet()` 経由で「線源情報」シートに書込（ヘッダー付き myF=1）

### 6.3 **SQL_Execution** (.bas)

ADO + ODBC による DB 接続・SQL 実行のユーティリティモジュール。EXメニューの他のブックと共通構成。

#### Public 変数

| 変数名 | 型 | 用途 |
|--------|-----|------|
| `mpErrDes` | String | エラーメッセージ格納 |
| `oraconn` | ADODB.Connection | DB 接続オブジェクト |
| `rs` | ADODB.Recordset | レコードセットオブジェクト |

#### `Open_oraconDB()` — Sub

```vb
oraconn.ConnectionString = "DSN=ricdb;UID=ric;PWD=t6101"
oraconn.Open
oraconn.CursorLocation = adUseClient
```

ODBC DSN `ricdb` を使用して Oracle DB に接続する。

#### `SQL_Exe(mySQL As String)` — Sub

Execute メソッドで SQL を実行し、結果を `rs` レコードセットに格納する。エラー発生時は `mpErrDes` にエラー内容を設定。

#### `SQL_INSERT_UPDATE(myTBL, myKey, myD(), myN)` — Sub

汎用 INSERT / UPDATE 処理。キー存在チェック→INSERT or UPDATE。本ブックでは直接呼び出されていない（他ブックとの共通モジュール）。

#### `SQL_Delete(myTBL, myWhere)` — Sub

汎用 DELETE 処理。本ブックでは直接呼び出されていない。

#### `Disp_Sheet(mySQL, mySH, myRow, myRecordCount, myColumn, myFieldCount, myF)` — Sub

SQL 実行結果をシート `mySH` に直接書き込む。

- `myF=1`: ヘッダー行を出力
- `myF=0`: データ行のみ
- `CopyFromRecordset` でバルク書き込み

#### `Set_Array(mySQL, myData(), myRecordCount, myFldCount)` — Sub

SQL 実行結果を二次元配列 `myData(i, j)` に格納する。本ブックでは未使用（共通モジュール）。

### 6.4 **Ex画面クリア** (.bas)

#### `GamennCls()` — Sub

ショートカットキー: Ctrl+E（VB_Invoke_Func = "e\n14"）

「完了予定時間」シートを選択し、以下の名前付き範囲を空白化:

| 名前付き範囲 | 対象 |
|-------------|------|
| `Hyouji` | 完了予定時間!C5:J50（メイン一覧） |
| `ExKei` | 出荷日情報!A2:D1000（出荷日データ） |
| `SenngennTB` | 線源情報!A2:D145（線源タイマー） |
| `SyouTB` | 照射データ!A2:U26（照射データ） |
| `Jyoutai` | 完了予定時間!F2:G3（線源状態） |

### 6.5 **Ex終了照射情報** (.bas)

#### `照射情報終了処理()` — Sub

```vb
Sub 照射情報終了処理()
    If Workbooks.Count <= 1 Then Application.Quit
    ThisWorkbook.Close savechanges:=False
End Sub
```

開いているブックが1つだけなら Excel ごと終了。複数なら当該ブックのみ閉じる（保存なし）。

### 6.6 **Ex出荷日情報読込Ric1** (.bas)

#### `SyukabiRead()` — Sub

出荷日情報を DB から取得する。

```sql
SELECT uno*1, trim(syukkabi), bikou1, syuhouhou
FROM ExKeikakuX
WHERE souti='1'
ORDER BY uno
```

- WHERE 条件: `souti='1'`（1号機装置のデータのみ）
- `Disp_Sheet()` 経由で「出荷日情報」シート 2行目以降に書込
- 書込後 `xlCalculationAutomatic` で表計算を自動に設定

### 6.7 **FunctionR1** (.bas)

#### `ExchengeDATE(myDate, myType)` — Function

| myType | 入力例 | 出力例 |
|--------|--------|--------|
| `"mm/dd"` | `1225` | `12/25` |
| `"yyyy/mm/dd"` | `20240101` | `2024/01/01` |

数値が0の場合はTrim済み文字列をそのまま返す。

#### `ExchengeDay(myDate, myType)` — Function

| myType | 入力例 | 出力例 |
|--------|--------|--------|
| `"mmdd"` | Date値 | `0101` |
| `"yyyymmdd"` | Date値 | `20240101` |

### 6.8 **GetPathRic1Jyou** (.bas)

#### `Ric1JyouGetPathX()` — Sub

```vb
Sub Ric1JyouGetPathX()
    Dim myFno As Double
    myFno = FreeFile
    Open ActiveWorkbook.Path & "\" & "ExAprReadPath.txt" For Input As #myFno
        Input #myFno, mpAprMotoPath, mpAprSakiPath, mpDB
    Close #myFno
End Sub
```

ブックと同じフォルダの `ExAprReadPath.txt` から3つの設定値を読み込む:

| 変数 | 用途 |
|------|------|
| `mpAprMotoPath` | 元パス（Public変数 — 他モジュールで定義） |
| `mpAprSakiPath` | 先パス |
| `mpDB` | データベース識別子 |

現在の `Kousinn()` フローでは `Ric1JyouGetPathX()` は直接呼ばれていない。

---

## 7. ユーザーフォーム仕様

本ブックにユーザーフォーム (UserForm) は存在しない。

---

## 8. DB接続・外部連携

### 8.1 DB接続情報

| 項目 | 値 |
|------|-----|
| 接続方式 | ADO + ODBC |
| DSN | `ricdb` |
| UID | `ric` |
| PWD | `t6101` |
| DB種別 | Oracle（SQL構文から推定: `RIC.SENGNR1` スキーマ修飾） |

### 8.2 参照テーブル一覧

| ✓ | テーブル名 | 区分 | アクセス種別 | 用途 |
|---|-----------|------|-------------|------|
| ✓ | `SYOUK1` | 参照 (SELECT) | R | 照射管理データ（未照射品一覧） |
| ✓ | `RIC.SENGNR1` | 参照 (SELECT) | R | 線源タイマー情報（直近データ取得） |
| ✓ | `ExKeikakuX` | 参照 (SELECT) | R | 出荷計画（出荷日・備考・出荷方法） |

本ブックは参照専用であり、DB への INSERT / UPDATE / DELETE は行わない。

### 8.3 SQL文一覧

| # | モジュール | SQL 概要 | テーブル |
|---|-----------|---------|----------|
| 1 | **Ex抽出処理** `Yomikomi()` | 未照射品の照射データ全件取得（syokind<>'2'） | SYOUK1 |
| 2 | **Ex抽出処理** `SenGenn()` | 最新日付の線源タイマー情報取得 | RIC.SENGNR1 |
| 3 | **Ex出荷日情報読込Ric1** `SyukabiRead()` | 1号機の出荷日情報取得（souti='1'） | ExKeikakuX |

### 8.4 外部ファイル参照

| ファイル名 | 読込モジュール | 用途 |
|-----------|--------------|------|
| `ExAprReadPath.txt` | **GetPathRic1Jyou** `Ric1JyouGetPathX()` | パス設定ファイル（ブックと同フォルダ） |

---

## 9. データフロー

### 9.1 データフローテーブル

| ✓ | # | 起点 | データ内容 | 経由 | 終点 | 方向 |
|---|---|------|-----------|------|------|------|
| ✓ | 1 | 🗄️ DB: `SYOUK1` | 照射管理データ（受付番号・会社名・線量・数量・位置等） | 📊 VBA `Yomikomi()` → `Disp_Sheet()` | 📊「照射データ」シート | DB → Excel |
| ✓ | 2 | 🗄️ DB: `RIC.SENGNR1` | 線源タイマー情報（日付・時刻・タイマー値・状態） | 📊 VBA `SenGenn()` → `Disp_Sheet()` | 📊「線源情報」シート | DB → Excel |
| ✓ | 3 | 🗄️ DB: `ExKeikakuX` | 出荷日情報（受付番号・出荷日・備考・出荷方法） | 📊 VBA `SyukabiRead()` → `Disp_Sheet()` | 📊「出荷日情報」シート | DB → Excel |
| ✓ | 4 | 📊「照射データ」シート + 📊「線源情報」シート | 残時間・完了予想日時の計算結果 | 📊 VBA `Kousinn()` 算出ロジック | 📊「完了予定時間」シート | Excel 内部 |
| ✓ | 5 | 📊「出荷日情報」シート | 出荷日・備考 | 📊 VLOOKUP 数式 (`SyukkabiTB`) | 📊「完了予定時間」K–L列 | Excel 内部 |
| ✓ | 6 | 📄 `ExAprReadPath.txt` | パス設定値 | 📊 VBA `Ric1JyouGetPathX()` | VBA Public 変数 | ファイル → VBA |
| ✓ | 7 | 🖥️ ユーザー操作 | 「最新情報」ボタン押下 | 📊 VBA `Kousinn()` | 全シート更新 | UI → VBA |

### 9.2 データフローツリー図

```
🖥️ ユーザー操作
 └─ [最新情報] ボタン押下
     └─ 📊 VBA: Kousinn()
         │
         ├─ 📊 VBA: SyukabiRead()
         │   └─ 🗄️ DB: ExKeikakuX (SELECT ... WHERE souti='1')
         │       └─ 📊「出荷日情報」シート (Disp_Sheet)
         │           └─ 📊「完了予定時間」K–L列 (VLOOKUP: SyukkabiTB)
         │
         ├─ 📊 VBA: Yomikomi()
         │   └─ 🗄️ DB: SYOUK1 (SELECT ... WHERE syokind<>'2')
         │       └─ 📊「照射ﾃﾞｰﾀ」シート (Disp_Sheet)
         │           └─ 📊 VBA: mySyo() 配列
         │               ├─ 照射位置コード変換 (11→南ｺﾝ 等)
         │               └─ 完了予想日時計算
         │
         ├─ 📊 VBA: SenGenn()
         │   └─ 🗄️ DB: RIC.SENGNR1 (SELECT MAX(SDATE) ... ORDER BY STIME DESC)
         │       └─ 📊「線源情報」シート (Disp_Sheet)
         │           └─ 📊 VBA: mySdate, myStime, myTimer, myEvent
         │               └─ 残時間補正計算
         │
         └─ 📊「完了予定時間」シート
             ├─ C–J列: No・受付番号・会社名・線量・数量・位置・残時間・完了予想日時
             ├─ F2: 線源状態 (照射中/貯蔵中/昇降中/移動照射中)
             ├─ F3: 現在時刻
             └─ Sort: 完了予想日時 (I5) 昇順

🖥️ ユーザー操作
 └─ [終了] ボタン押下
     └─ 📊 VBA: 照射情報終了処理()
         └─ ブック閉じる (保存なし)

📊 自動実行 (Workbook_Open)
 └─ シート保護設定 (UserInterfaceOnly)
 └─ GamennCls() — 画面クリア
```

---

## 10. セキュリティ注意事項

### 10.1 DB 認証情報のハードコード

**SQL_Execution** モジュール内に DB 接続文字列がハードコードされている:

```vb
oraconn.ConnectionString = "DSN=ricdb;UID=ric;PWD=t6101"
```

- **リスク**: UID / PWD がソースコードに平文で記載されている
- **推奨**: 接続文字列を外部設定ファイル（既存の `ExAprReadPath.txt` 等）に移動するか、Windows 認証への切替を検討

### 10.2 シート保護

- `Workbook_Open` で `ActiveSheet.Protect UserInterfaceOnly:=True` を設定
- パスワードなしの保護のため、ユーザーが手動でシート保護を解除可能
- 保護の目的は誤操作防止であり、セキュリティ対策としては不十分

### 10.3 エラーハンドリング

- **SQL_Execution** の `SQL_Exe()`: `On Error Resume Next` でエラーを無視し、`mpErrDes` に格納するのみ。致命的な DB エラーが暗黙に無視される可能性がある
- **Disp_Sheet()**: `On Error Resume Next` が広範囲に適用され、CopyFromRecordset のエラーが検出されにくい
- **Open_oraconDB()**: 接続エラーが `On Error Resume Next` で無視される。接続失敗時に後続 SQL が実行されると予期しない動作になる

### 10.4 VBA Stomping

olevba の分析では VBA Stomping（P-code とソースコードの不一致）は検出されていない。

### 10.5 無効な名前付き範囲

`#REF!` を参照する名前付き範囲が8個存在する（`DataInp`, `DebugFlg`, `InpTbl`, `Misyousya`, `NowTime`, `Nuru`, `SetTbl`, `TyuudannInp`）。現在のコードでは使用されていないが、過去のシート構成変更の残骸であり、クリーンアップを推奨。

### 10.6 外部ファイル読込

`GetPathRic1Jyou` モジュールが `ExAprReadPath.txt` を Open 文で読み込む。ファイルのアクセス権限やパス改ざんに対する検証は行われていない。

### 10.7 コードネーム⇔シート名対応

| コードネーム | workbook.xml シート名 | worksheet XML |
|-------------|----------------------|---------------|
| Sheet4 | 「完了予定時間」 | xl/worksheets/sheet1.xml |
| Sheet2 | 「照射データ」 | xl/worksheets/sheet2.xml |
| Sheet5 | 「線源情報」 | xl/worksheets/sheet3.xml |
| Sheet1 | 「出荷日情報」 | xl/worksheets/sheet4.xml |

VBA コード内ではシート名による参照（`Sheets("完了予定時間")` 等）を使用しており、コードネームによるアクセスは行われていない。
