# Ex出荷方法報告書発行登録 仕様書

> **ファイル種別**: .xlsm（マクロ付き）
> **用途**: 得意先（業者）ごとの出荷方法（引取・混載便等）と報告書発行種別（Fax送信・報告書発行等）をDB（ExSeihinJ テーブル）に登録・更新する管理ツール
> **VBA プロジェクト**: モジュール 9 本（.bas 7 / .cls 2 / .frm 0）
> **外部連携**: DSN=ricdb（Oracle DB）、UID=ric
> **解析日**: 2026-06-18（excel-to-md スキルによる自動解析）

---

## 凡例（本仕様書の表記ルール）

本仕様書では、保守時の判別を容易にするため、以下の表記ルールを使用します。

| 種別 | 表記 | 例 |
|---|---|---|
| モジュール（.bas / .cls） | **太字** | **登録.bas** |
| プロシージャ / イベント | `コード体()` | `引取報告書登録()` |
| シート名 | 「」 | 「業者一覧」 |
| セル参照 | `コード体` | `$B$5` |
| 名前付き範囲 | `コード体` | `HikitoriTB` |
| DB テーブル / カラム | `コード体` | `ExSeihinJ` / `hikitori` |
| ユーザー操作 | （操作名） | （登録実行） |
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
| 1.3 / 6.0 | VBA モジュール | ① ユーザー操作の起点 ② DB I/O を含む ③ 他モジュールから呼び出される ④ コード行数上位 25% のいずれか |
| 2 | セル / 名前付き範囲 | VBA から `Range()` で代入または参照され、業務ロジックに直結する |
| 3 | 名前付き範囲 | VBA から `Range()` で代入または参照され、業務ロジックに直結する |
| 5 | ボタン / ショートカット | DB 更新・画面遷移・計算実行など副作用のある操作を起動する |
| 6.0 | プロシージャ | ① ユーザー操作の起点（イベント等） ② DB I/O を実行 ③ 他モジュールから呼び出される Public のいずれか |
| 8.2 | DB テーブル | INSERT / UPDATE / DELETE の対象（参照のみのテーブルは ✓ なし） |

---

## 目次

1. [ファイル構成](#1-ファイル構成)
2. [シート詳細](#2-シート詳細)
3. [名前付き範囲一覧](#3-名前付き範囲一覧)
4. [数式一覧](#4-数式一覧)
5. [ボタン・マクロ対応](#5-ボタンマクロ対応)
6. [VBA モジュール仕様](#6-vba-モジュール仕様)
7. [ユーザーフォーム仕様](#7-ユーザーフォーム仕様)
8. [DB 接続・外部連携](#8-db-接続外部連携)
9. [データフロー](#9-データフロー)
10. [セキュリティ注意事項](#10-セキュリティ注意事項)

---

## 1. ファイル構成

### 1.1 シート一覧（全 1 件）

> ✓ = ユーザーが直接操作する、または VBA が動的に表示/非表示を切り替えるシート

| ✓ | シート名 | 最大行 | 最大列 | 保存時 Visible | VBA による動的切替 |
|---|---|---|---|---|---|
| ✓ | 業者一覧 | 1004 | 17 (Q) | visible | — |

### 1.2 ユーザーフォーム一覧

本ファイルにユーザーフォームは存在しない。

### 1.3 VBA モジュール一覧（全 9 件）

> ✓ = ① ユーザー操作起点 ② DB I/O 含む ③ 他モジュールから呼出 ④ コード行数上位 25% のいずれか

| ✓ | モジュール名 | 種別 | 行数 | 概要 | ✓ 理由 |
|---|---|---|---|---|---|
| ✓ | **ThisWorkbook.cls** | .cls | 16 | ブックOpen時にシート保護設定＋DB抽出呼出、Close時に保存なし終了 | ① |
| ✓ | **Sheet1.cls** | .cls | 20 | セル変更イベントでG列（報告書発行種別）変更時にsehmstの報告書フラグを検証 | ① |
| ✓ | **登録.bas** | .bas | 46 | 「業者一覧」シートの変更データをExSeihinJにINSERT/UPDATE | ② |
| | **終了処理.bas** | .bas | 12 | ブック終了処理（上書きなしでClose） | — |
| ✓ | **デーた抽出.bas** | .bas | 66 | tokumst・ExSeihinJ・sehmstからデータ抽出し「業者一覧」に展開 | ②④ |
| | **ユーティリティ.bas** | .bas | 10 | イベント有効化(Ctrl+R)・画面クリアのユーティリティ | — |
| | **画面クリア引取業者.bas** | .bas | 12 | 名前付き範囲のクリアとセル選択(Ctrl+E) | — |
| ✓ | **SQL_Execution.bas** | .bas | 183 | ADODB接続・SQL実行・INSERT/UPDATE/DELETE・シートへのRecordset展開 | ②③④ |
| ✓ | **画面操作1.bas** | .bas | 216 | 印刷パラメータ設定・ドライブ検索・シート保護・各種UI操作ユーティリティ集 | ④ |

---

## 2. シート詳細

### 2.0 シート可視性一覧

| シート名 | 保存時状態 | VBA切替 | 備考 |
|---|---|---|---|
| 業者一覧 | visible | — | 唯一のシート。ユーザーが直接操作する |

### 2.1 「業者一覧」シート

**タイトル行**

| ✓ | セル | 値 | 業務的意味 |
|---|---|---|---|
| | `B1` | 出荷方法／報告書発行種別の登録 | 画面タイトル |

**ヘッダー行（4行目）**

| ✓ | 列 | セル | ヘッダー名 | 業務的意味 |
|---|---|---|---|---|
| ✓ | B | `B4` | 会社ｺｰﾄﾞ | `tokumst.kaisyacd`（得意先コード、4桁未満） |
| | C | `C4` | 略称 | `tokumst.kairname`（得意先略称） |
| | D | `D4` | 会社名 | `tokumst.coname`（得意先正式名称） |
| ✓ | E | `E4` | 出荷方法 | `ExSeihinJ.hikitori`（DB登録済の現在値） |
| ✓ | F | `F4` | 出荷方法 | DB値のコピー（E列と比較して変更検知に使用） |
| ✓ | G | `G4` | 報告書発行種別 | `ExSeihinJ.housyube`（DB登録済の現在値） |
| ✓ | H | `H4` | 報告書発行種別 | DB値のコピー（G列と比較して変更検知に使用） |
| ✓ | I | `I4` | Sehmstの報告書要不要 | VLOOKUP数式でK:L列から報告書発行フラグを参照 |
| | K | `K4` | 会社ｺｰﾄﾞ | sehmst抽出結果のキー列（VBAが書込み） |
| | L | `L4` | 報告書発行フラグ | sehmst抽出結果のフラグ値（VBAが書込み） |

**データ領域**

| 領域 | 行範囲 | 列範囲 | 説明 |
|---|---|---|---|
| 業者データ | 5〜1004行 | B〜I列 | `HikitoriTB` 範囲。VBAがDB抽出データを展開 |
| sehmst参照 | 5〜1004行 | K〜L列 | `HouTB` 範囲。sehmstの報告書フラグ参照用 |

**凡例エリア（N〜O列）**

出荷方法と報告書発行種別の選択肢マスタ。

| ✓ | セル | 出荷方法 | 報告書発行種別内容 |
|---|---|---|---|
| | `N5`/`O5` | 引取 | 照射後Fax送信 |
| | `N6`/`O6` | 混載便 | 出荷後Fax送信 |
| | `N7`/`O7` | 保管品 | 照射後報告書発行 |
| | `N8`/`O8` | ﾁｬｰﾀ便 | 出荷後報告書発行 |
| | `N9` | 納品 | — |
| | `N10` | 品証扱い | — |
| | `N11` | γ扱い | — |
| | `N12` | 営業扱い | — |
| | `N13` | その他 | — |

---

## 3. 名前付き範囲一覧（全 3 件）

> ✓ = VBA から `Range()` で代入または参照され、業務ロジックに直結する

| ✓ | 名前 | 参照先 | 業務的意味 |
|---|---|---|---|
| ✓ | `HikitoriTB` | `'業者一覧'!$B$5:$H$1005` | 業者データ本体領域。`業者名と引取抽出()`・`画面消去処理()`・`画面クリア()`で Range("HikitoriTB")="" としてクリアされる。`Disp_Sheet()`で tokumst+ExSeihinJ のSELECT結果が書き込まれる |
| ✓ | `HouTB` | `'業者一覧'!$K$5:$L$1004` | sehmst報告書フラグ参照領域。`業者名と引取抽出()`・`画面消去処理()`・`画面クリア()`でクリアされる。sehmstのSELECT結果が書き込まれ、I列のVLOOKUP数式が参照する |
| | `Debug` | `'業者一覧'!$A$1` | デバッグ用セル。VBAからの直接参照は確認されない |

---

## 4. 数式一覧

### 4.0 シート別サマリ

| シート名 | 数式セル数 | 数式パターン数 | 備考 |
|---|---|---|---|
| 業者一覧 | 1000 | 1 | I5:I1004 に同一パターンのVLOOKUP |

### 4.1 「業者一覧」シートの数式

| ✓ | セル範囲 | 数式パターン | 業務的意味 |
|---|---|---|---|
| ✓ | `I5:I1004` | `=IF(ISERROR(VLOOKUP(B5,$K$5:$L$1004,2,FALSE)),"",VLOOKUP(B5,$K$5:$L$1004,2,FALSE))` | B列の会社コードでK:L列（sehmst抽出結果）を検索し、報告書発行フラグ（`max(syouho)`）を表示。VLOOKUP失敗時は空文字。Worksheet_ChangeイベントでG列変更時にI列=0なら「不要」に上書きされる |

---

## 5. ボタン・マクロ対応

### 5.1 シート上ボタン

シート上にマクロ割当済のボタン（フォームコントロール/ActiveXコントロール）は検出されなかった。

> **注記**: workbook.xml に GoogleSheetsCustomDataVersion2 の拡張データが存在する。Google Sheets 経由で保存された可能性があり、元のExcelファイルに存在していたフォームコントロールのボタンが失われている可能性がある。`引取報告書登録()`等の Public Sub はボタンから呼び出される設計と推定される。

検出されたオブジェクト:

| 種別 | 名前 | 位置 | マクロ割当 | 備考 |
|---|---|---|---|---|
| テキストボックス | Shape 3 | J列付近・1行目 | なし | テキスト "CLS" を表示。マクロ未割当 |

### 5.2 キーボードショートカット

VBA の `Attribute VB_ProcData.VB_Invoke_Func` で定義されたショートカットキー。

| ✓ | ショートカット | 呼出プロシージャ | モジュール | 効果 |
|---|---|---|---|---|
| | Ctrl+R | `イベント有効()` | **ユーティリティ.bas** | `Application.EnableEvents = True` でイベント処理を再有効化 |
| | Ctrl+E | `画面消去処理()` | **画面クリア引取業者.bas** | `HikitoriTB` と `HouTB` をクリアし `E5` セルを選択 |

### 5.3 推定されるボタン割当（VBAから推定）

以下のプロシージャはユーザー操作起点として設計されており、元ファイルではボタンに割り当てられていたと推定される。

| プロシージャ | モジュール | 推定される操作 |
|---|---|---|
| `引取報告書登録()` | **登録.bas** | 変更データのDB登録（INSERT/UPDATE）実行 |
| `Bookを閉じる()` | **終了処理.bas** | ブックを閉じる |
| `業者名と引取抽出()` | **デーた抽出.bas** | DBからデータ再取得（通常はWorkbook_Openから自動呼出） |

---

## 6. VBA モジュール仕様

### 6.0 全プロシージャ一覧（全 20 件）

> ✓ = ① ユーザー操作の起点 ② DB I/O を実行 ③ 他モジュールから呼び出される Public のいずれか

| ✓ | プロシージャ名 | モジュール | 種別 | スコープ | 概要 | ✓ 理由 |
|---|---|---|---|---|---|---|
| ✓ | `Workbook_BeforeClose()` | **ThisWorkbook.cls** | Event | Private | 警告非表示・保存済みフラグをTrueにして終了 | ① |
| ✓ | `Workbook_Open()` | **ThisWorkbook.cls** | Event | Private | シート保護解除→UIのみ保護設定→`業者名と引取抽出()`呼出 | ①③ |
| ✓ | `Worksheet_Change()` | **Sheet1.cls** | Event | Private | G列変更時にI列(sehmst報告書フラグ)=0なら「不要」に上書きしMsgBox警告 | ① |
| ✓ | `引取報告書登録()` | **登録.bas** | Sub | Public | 「業者一覧」の E列vsF列・G列vsH列 を比較し、変更があれば `SQL_INSERT_UPDATE()` で `ExSeihinJ` を更新 | ②③ |
| ✓ | `業者名と引取抽出()` | **デーた抽出.bas** | Sub | Public | tokumst+ExSeihinJ を外部結合SELECTし `HikitoriTB` へ展開、sehmst集計を `HouTB` へ展開、報告書不要判定処理 | ②③ |
| | `報告書不要表示()` | **デーた抽出.bas** | Sub | Public | sehmstフラグ=0の業者に「不要」を設定（`業者名と引取抽出()`内で同等処理済のため予備的） | — |
| | `Bookを閉じる()` | **終了処理.bas** | Sub | Public | ワークブック数1なら `Application.Quit`、複数なら `ActiveWorkbook.Close` | — |
| | `イベント有効()` | **ユーティリティ.bas** | Sub | Public | `Application.EnableEvents = True`（Ctrl+R） | — |
| | `画面クリア()` | **ユーティリティ.bas** | Sub | Public | `HikitoriTB` と `HouTB` を空文字クリア | — |
| | `画面消去処理()` | **画面クリア引取業者.bas** | Sub | Public | `HikitoriTB`・`HouTB` クリア後 `E5` を選択（Ctrl+E） | — |
| ✓ | `Open_oraconDB()` | **SQL_Execution.bas** | Sub | Public | DSN=ricdb で ADODB.Connection を Open | ②③ |
| ✓ | `SQL_Exe()` | **SQL_Execution.bas** | Sub | Public | `oraconn.Execute(mySQL)` で SQL 実行。エラー時は `mpErrDes` にセット | ②③ |
| ✓ | `SQL_INSERT_UPDATE()` | **SQL_Execution.bas** | Sub | Public | SELECT COUNT→件数0ならINSERT、>0ならUPDATE。トランザクション付き | ②③ |
| ✓ | `SQL_Delete()` | **SQL_Execution.bas** | Sub | Public | DELETE文をトランザクション付きで実行 | ②③ |
| ✓ | `Disp_Sheet()` | **SQL_Execution.bas** | Sub | Public | SQL実行→Recordsetをシートに `CopyFromRecordset` で展開 | ②③ |
| ✓ | `Set_Array()` | **SQL_Execution.bas** | Sub | Public | SQL実行→Recordsetを2次元配列 `myData()` に格納 | ②③ |
| | `Reidai()` | **画面操作1.bas** | Sub | Public | VBA操作の例題コード集（セル移動・ウィンドウ操作・保護設定等） | — |
| | `印刷パラメータ設定()` | **画面操作1.bas** | Sub | Public | シートの印刷設定（余白・向き・倍率等）をパラメータ指定で一括設定 | — |
| | `DriveSearch()` | **画面操作1.bas** | Function | Public | 指定ドライブの存在チェック（FSO使用） | — |
| | `シート保護()` | **画面操作1.bas** | Sub | Public | シート保護の解除→再設定（UIのみ保護、カーソル制限等） | — |
| | `複数列の選択()` | **画面操作1.bas** | Sub | Public | 複数列選択の例題コード | — |
| | `セル名の定義_削除()` | **画面操作1.bas** | Sub | Public | セル名定義・削除の例題コード | — |
| | `リボン操作()` | **画面操作1.bas** | Sub | Public | リボン表示制御の例題コード | — |
| | `セルのコピー_値の貼り付け()` | **画面操作1.bas** | Sub | Public | セルコピー・値貼り付けの例題コード | — |

### 6.1 **ThisWorkbook.cls**

```vb
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Application.DisplayAlerts = False
    ActiveWorkbook.Saved = True
End Sub

Private Sub Workbook_Open()
    ActiveSheet.Unprotect
    ActiveSheet.Protect UserInterfaceOnly:=True
    Call 業者名と引取抽出
End Sub
```

- `Workbook_Open`: ブック起動時にシート保護をUI限定で設定し、`業者名と引取抽出()` を呼び出してDBからデータを取得する
- `Workbook_BeforeClose`: 警告を非表示にし、保存なしで閉じる

### 6.2 **Sheet1.cls**（「業者一覧」シート）

```vb
Private Sub Worksheet_Change(ByVal Target As Range)
    Application.EnableEvents = False
    With Target
        If .Row > 4 And .Column = 7 Then
            If Cells(.Row, .Column + 2) = 0 Then
                Cells(.Row, .Column) = "不要"
                MsgBox "製品仕様台帳の報告書発行が不要になっています..."
            End If
        End If
    End With
    Application.EnableEvents = True
End Sub
```

- G列（報告書発行種別）が変更されたとき、I列（sehmst報告書フラグ）=0 なら「不要」に強制上書きし警告メッセージを表示
- 5行目以降のデータ行のみ対象

### 6.3 **登録.bas**

```vb
Sub 引取報告書登録()
    ' E列(出荷方法)とF列(DB値)、G列(報告書種別)とH列(DB値)を比較
    ' 差異があれば ExSeihinJ に kaisyacd をキーとして INSERT or UPDATE
    ' カラム: kaisyacd, hikitori, housyube
End Sub
```

- テーブル `ExSeihinJ` に対し、`kaisyacd`（会社コード）をキーとして `hikitori`（出荷方法）・`housyube`（報告書発行種別）を登録
- 変更検知: E列≠F列 または G列≠H列 の行のみ更新対象
- `SQL_INSERT_UPDATE()` を呼び出し、レコードが無ければ INSERT、あれば UPDATE

### 6.4 **デーた抽出.bas**

```vb
Sub 業者名と引取抽出()
    ' 1. HikitoriTB / HouTB をクリア
    ' 2. tokumst LEFT JOIN ExSeihinJ で業者情報取得 → B5〜H列に展開
    ' 3. sehmst から kaisyacd別 max(syouho) 取得 → K5〜L列に展開
    ' 4. I列(sehmst報告書フラグ)=0 or 空 の行に「不要」を設定
End Sub
```

- SQL1: `SELECT t.kaisyacd,t.kairname,t.coname,s.hikitori,s.hikitori,s.housyube,s.housyube FROM tokumst t,ExSeihinJ s WHERE t.kaisyacd=s.kaisyacd(+) and t.kaisyacd<'2000' ORDER BY t.kaisyacd`
  - Oracle外部結合 `(+)` を使用。会社コード 2000 未満の得意先が対象
  - hikitori と housyube がそれぞれ2回SELECTされる（E列=現在値、F列=比較用DB値）
- SQL2: `SELECT kaisyacd,max(syouho) FROM sehmst GROUP BY kaisyacd ORDER BY kaisyacd`
  - 製品マスタの報告書発行要否フラグの最大値を取得
- 報告書不要判定: sehmstフラグ=0 or 空 → G列・H列に「不要」を設定
- ループ上限: 2000行（超過時はエラーメッセージで処理停止）

### 6.5 **SQL_Execution.bas**

本ファイルの DB 接続基盤モジュール。全 DB アクセスはこのモジュールを経由する。

**モジュールレベル変数**

| 変数名 | 型 | スコープ | 用途 |
|---|---|---|---|
| `mpErrDes` | String | Public | エラーメッセージ格納。各呼出元でチェックされる |
| `mpDSN` | String | Public | DSN接続文字列 |
| `oraconn` | ADODB.Connection | Private | DB接続オブジェクト |
| `rs` | ADODB.Recordset | Private | レコードセット |

**プロシージャ詳細**

| プロシージャ | 引数 | 処理概要 |
|---|---|---|
| `Open_oraconDB()` | なし | DSN=ricdb;UID=ric;PWD=t6101 で接続。CursorLocation=adUseClient |
| `SQL_Exe()` | mySQL | `oraconn.Execute(mySQL)` 実行。エラー時 `mpErrDes` にセット |
| `SQL_INSERT_UPDATE()` | myTBL, myKey, myD(), myN | SELECT COUNT→0:INSERT / >0:UPDATE。BeginTrans〜CommitTrans |
| `SQL_Delete()` | myTBL, myWhere | DELETE文をトランザクション付きで実行 |
| `Disp_Sheet()` | mySQL, mySH, myRow, ... | SQL実行→`CopyFromRecordset`でシートに展開 |
| `Set_Array()` | mySQL, myData(), ... | SQL実行→2次元配列に格納 |

### 6.6 **終了処理.bas**

```vb
Sub Bookを閉じる()
    Application.DisplayAlerts = False
    If Application.Workbooks.Count = 1 Then
        Application.Quit
    Else
        ActiveWorkbook.Close
    End If
End Sub
```

### 6.7 **ユーティリティ.bas**

| プロシージャ | ショートカット | 処理 |
|---|---|---|
| `イベント有効()` | Ctrl+R | `Application.EnableEvents = True` |
| `画面クリア()` | — | `Range("HikitoriTB") = ""` / `Range("HouTB") = ""` |

### 6.8 **画面クリア引取業者.bas**

```vb
Sub 画面消去処理()  ' Ctrl+E
    Range("HikitoriTB") = ""
    Range("HouTB") = ""
    Range("E5").Select
End Sub
```

### 6.9 **画面操作1.bas**

VBA 操作のユーティリティ・例題コード集。本ブックの業務ロジックには直接関与しないが、`印刷パラメータ設定()` と `DriveSearch()` は汎用関数として他ブックから参照される可能性がある。

| プロシージャ | 種別 | 概要 |
|---|---|---|
| `Reidai()` | Sub | VBA操作例題（セル移動・ウィンドウ操作・保護設定等のサンプルコード集） |
| `印刷パラメータ設定()` | Sub | 印刷設定をパラメータで一括指定（余白・向き・倍率・A4・1ページ収め） |
| `DriveSearch()` | Function | FSO でドライブ存在チェック。戻り値 Boolean |
| `シート保護()` | Sub | シート保護の解除→UIのみ保護→カーソル制限設定 |
| `複数列の選択()` | Sub | 複数列選択の例題 |
| `セル名の定義_削除()` | Sub | セル名定義・削除の例題 |
| `リボン操作()` | Sub | リボン表示制御（ExecuteExcel4Macro使用） |
| `セルのコピー_値の貼り付け()` | Sub | セルコピー・値貼り付けの例題 |

---

## 7. ユーザーフォーム仕様

本ファイルにユーザーフォーム（.frm）は存在しない。

---

## 8. DB 接続・外部連携

### 8.1 ODBC 接続情報

| 項目 | 値 |
|---|---|
| 接続方式 | ADO + ODBC（ADODB.Connection） |
| DSN | ricdb |
| UID | ric |
| PWD | t6101 |
| CursorLocation | adUseClient |
| トランザクション | BeginTrans / CommitTrans（INSERT/UPDATE/DELETE時） |

### 8.2 テーブル一覧

| ✓ | テーブル名 | 操作種別 | 用途 |
|---|---|---|---|
| ✓ | `ExSeihinJ` | SELECT / INSERT / UPDATE | 得意先別の出荷方法（hikitori）・報告書発行種別（housyube）を管理。登録処理の更新対象 |
| | `tokumst` | SELECT | 得意先マスタ。会社コード・略称・会社名の取得元 |
| | `sehmst` | SELECT | 製品マスタ。報告書発行要否フラグ（syouho）の集計元 |

### 8.3 SQL 一覧

| # | 実行元 | 操作 | SQL |
|---|---|---|---|
| 1 | `業者名と引取抽出()` | SELECT | `SELECT t.kaisyacd,t.kairname,t.coname,s.hikitori,s.hikitori,s.housyube,s.housyube FROM tokumst t,ExSeihinJ s WHERE t.kaisyacd=s.kaisyacd(+) and t.kaisyacd<'2000' ORDER BY t.kaisyacd` |
| 2 | `業者名と引取抽出()` | SELECT | `SELECT kaisyacd,max(syouho) FROM sehmst GROUP BY kaisyacd ORDER BY kaisyacd` |
| 3 | `SQL_INSERT_UPDATE()` | SELECT | `SELECT COUNT(*) FROM {テーブル} WHERE {キー条件}` |
| 4 | `SQL_INSERT_UPDATE()` | INSERT | `INSERT INTO {テーブル} ({列名...}) VALUES({値...})` |
| 5 | `SQL_INSERT_UPDATE()` | UPDATE | `UPDATE {テーブル} SET {列名=値...} WHERE {キー条件}` |
| 6 | `SQL_Delete()` | DELETE | `DELETE {テーブル} {WHERE条件}` |

> SQL #3〜#6 は汎用プロシージャのため、呼出元のパラメータにより対象テーブル・条件が変わる。本ブックでは `ExSeihinJ` テーブルに対して `kaisyacd` をキーとした INSERT/UPDATE のみ実行される。

### 8.4 外部ファイル参照

| 変数 | 定義 | 備考 |
|---|---|---|
| `mpFnameSyu` | `"\引取業者.txt"` | **登録.bas** で定数定義されているが、コード内で参照箇所なし。過去の実装残存と推定 |
| `mpApMotoPath` 等 | Public String | **登録.bas** で宣言されているが、コード内で使用されていない。パス関連の未使用変数 |

---

## 9. データフロー

### 9.1 データフローテーブル

| # | トリガー | 処理 | 場所 | 入力 | 出力 | 備考 |
|---|---|---|---|---|---|---|
| 1 | ブックOpen | `Workbook_Open()` | 📄 | — | — | シート保護設定後、`業者名と引取抽出()` を Call |
| 2 | #1 から呼出 | `業者名と引取抽出()` 前半 | 🗄️ | `tokumst` + `ExSeihinJ` | `HikitoriTB`（B5:H列） | Oracle外部結合で得意先+出荷方法を取得し展開 |
| 3 | #2 に続く | `業者名と引取抽出()` 後半 | 🗄️ | `sehmst` | `HouTB`（K5:L列） | kaisyacd別 max(syouho) を取得し展開 |
| 4 | #3 に続く | 報告書不要判定 | 📊 | I列（sehmstフラグ） | G列・H列 | I列=0 or 空 の行に「不要」を設定 |
| 5 | ユーザー操作 | E列・G列を編集 | 📊 | ユーザー入力 | E列・G列 | 出荷方法・報告書発行種別を手動変更 |
| 6 | #5 でG列変更時 | `Worksheet_Change()` | 📄 | G列・I列 | G列 | I列=0 なら「不要」に戻しMsgBox警告 |
| 7 | （登録実行） | `引取報告書登録()` | 📄 | E列vsF列、G列vsH列 | — | 変更行を検出 |
| 8 | #7 の変更行ごと | `SQL_INSERT_UPDATE()` | 🗄️ | kaisyacd, hikitori, housyube | `ExSeihinJ` | 既存レコードならUPDATE、なければINSERT |
| 9 | Ctrl+E | `画面消去処理()` | 📊 | — | `HikitoriTB`・`HouTB` | 全データクリア |
| 10 | ブックClose | `Workbook_BeforeClose()` | 📄 | — | — | 保存なしで終了 |

### 9.2 データフローツリー図

```
📄 Workbook_Open()
├── 📊 ActiveSheet.Unprotect → Protect (UIのみ)
└── 🗄️ 業者名と引取抽出()
    ├── 📊 Range("HikitoriTB") = "" / Range("HouTB") = "" … クリア
    ├── 🗄️ SELECT tokumst LEFT JOIN ExSeihinJ
    │   └── 📊 Disp_Sheet() → HikitoriTB (B5:H列) に展開
    ├── 🗄️ SELECT sehmst GROUP BY kaisyacd
    │   └── 📊 Disp_Sheet() → HouTB (K5:L列) に展開
    └── 📊 報告書不要判定ループ
        └── 📊 I列=0 or 空 → G列・H列 = "不要"

📊 ユーザー操作: E列(出荷方法)・G列(報告書発行種別) を編集
└── 📄 Worksheet_Change() [G列変更時のみ]
    └── 📄 I列(sehmst報告書フラグ) = 0 ?
        ├── Yes → 📊 G列 = "不要" + MsgBox警告
        └── No  → 変更を受け入れ

📄 引取報告書登録() [ユーザー操作起点]
├── 📄 業者一覧ループ: E列≠F列 or G列≠H列 を検出
│   └── 🗄️ SQL_INSERT_UPDATE("ExSeihinJ", "kaisyacd='{会社CD}'")
│       ├── 🗄️ SELECT COUNT(*) → 0件: INSERT
│       └── 🗄️ SELECT COUNT(*) → 1件以上: UPDATE
└── 📊 MsgBox("更新しました" or エラーメッセージ)

📄 Bookを閉じる() / Workbook_BeforeClose()
└── 📄 保存なしで終了
```

---

## 10. セキュリティ注意事項

### olevba 検出結果

| 種別 | キーワード | 説明 | リスク評価 |
|---|---|---|---|
| AutoExec | `Workbook_Open` | ブック起動時に自動実行 | 低：DB接続＋データ取得のみ |
| AutoExec | `Workbook_Activate` | ブックアクティブ時に自動実行 | 低：VBAコード内の例題参照 |
| AutoExec | `Workbook_BeforeClose` | ブック終了時に自動実行 | 低：保存なし終了のみ |
| AutoExec | `Worksheet_Change` | セル変更時に自動実行 | 低：入力検証のみ |
| Suspicious | `Environ` | 環境変数の読取り | 低：**画面操作1.bas** の例題コード内（`Environ("COMPUTERNAME")`）。業務ロジックでは未使用 |
| Suspicious | `Open` | ファイルオープン | 低：ADODB接続の `oraconn.Open` |
| Suspicious | `Call` | DLL呼出の可能性 | 低：VBA内部のプロシージャ呼出（`Call SQL_INSERT_UPDATE` 等） |
| Suspicious | `CreateObject` | OLEオブジェクト生成 | 低：**画面操作1.bas** の `CreateObject("Scripting.FileSystemObject")`。ドライブ存在チェック用 |
| Suspicious | `ExecuteExcel4Macro` | Excel4マクロ実行 | 中：**画面操作1.bas** のリボン操作例題。業務ロジックでは未使用だが、Excel4マクロは潜在的リスク |
| Suspicious | `Windows` | ウィンドウ列挙 | 低：`ActiveWindow.WindowState` 等の例題コード |
| Suspicious | `Chr` | 文字列難読化の可能性 | 低：`Chr(13)`（改行）をMsgBox内で使用 |
| Suspicious | Hex Strings | 16進文字列検出 | 低：自動検出によるもの。意図的な難読化ではない |
| Suspicious | Base64 Strings | Base64文字列検出 | 低：自動検出によるもの。意図的な難読化ではない |

### その他の注意事項

- **DB認証情報のハードコード**: `SQL_Execution.bas` 内に DSN・UID・PWD が平文で記述されている（`mpDSN = "DSN=ricdb;UID=ric;PWD=t6101"`）
- **On Error Resume Next の多用**: `SQL_Execution.bas` の各プロシージャで `On Error Resume Next` が使用されており、エラーが暗黙的に無視される可能性がある
- **Stop ステートメント**: `SQL_Exe()` 内に `Stop` が残存しており、本番環境でエラー発生時にデバッグモードに入る可能性がある。`デーた抽出.bas` の `報告書不要表示()` にも `If i > 1000 Then Stop` が存在
- **Google Sheets 保存痕跡**: workbook.xml に `GoogleSheetsCustomDataVersion2` が存在し、フォームコントロール（ボタン等）が失われている可能性がある
- **未使用変数・定数**: `mpApMotoPath`, `mpTxMotoPath`, `mpApSakiPath`, `mpTxSakiPath`, `mpFnameSyu` が宣言のみで未使用
