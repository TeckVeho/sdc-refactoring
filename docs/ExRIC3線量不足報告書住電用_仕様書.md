# ExRIC3線量不足報告書住電用 仕様書

> **ファイル種別**: .xlsm（マクロ付き）
> **用途**: ３号機線量不足照射管理番号に対して追加照射指示書（工程確認・追加照射指示書）を住友電工専用フォーマットで作成・印刷する
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
ExRIC3線量不足報告書住電用.xlsm
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

**目的**: 線量不足と判定された照射管理番号を選択し、追加照射に必要な測定データ・照射計画データを DB から取得して帳票を完成させ、印刷する。

#### レイアウト構造

| セル範囲 | 名前付き範囲 | 内容 |
|---|---|---|
| A1 | `Debug` | デバッグフラグ（通常空白） |
| B1 | — | 様式番号（数式で改訂版数を動的生成） |
| B2 | — | タイトル「３号機線量不足報告書　住友電工専用（工程確認・追加照射指示書）」 |
| E3 | `Tuika` | 追加照射回目 |
| B7 | `SenkNo` | 照射管理番号（ドロップダウン入力、選択でデータ抽出トリガー） |
| G7 | `SokDate` | 測定日 |
| K7 | `SoktCD` | 測定者コード |
| O7 | `SaiSoku` | 再測定者 |
| S7 | `SokutSN` | 測定線量 |
| W7 | `Tani` | 製品目的（単位） |
| AA7 | `Pass` | 指定パス数 |
| AD7 | `SikiKodo` | 計算式コード |
| AG7 | `Sosi` | 素子異常 |
| AJ7 | `Atusa` | 素子厚さ |
| G7:AN7 | `Data1` | 測定データ行（１行分） |
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
| G16 | `SEnKind` | 線量計種類（入力でDB照合） |
| O16 | `AtusaT` | 線量計厚さ |
| W16 | `Keisask` | 計算式コード |
| B18〜B22 | — | 追加照射データ行（受付番号参照） |
| AA18〜AA22 | — | 追加パス数計算（ROUNDUP 数式） |
| M24 | — | 測定データ不適切警告 |
| S24 | — | 追加パス数不一致警告 |
| B30 | — | 注記「本紙は照射計画作成後作業指図書に貼付すること」 |
| BC3:BD6 | `mokuteki` / `MokuCD` | 製品目的コード・説明テーブル |
| BC10:BC12 | `SenSyu` | 線量計種類テーブル（DB から取得） |
| BC21 | `TuikaMin` | 線量計高・低閾値（=5） |
| BG3 | `SebkNoTB` | 線量不足データ件数 |
| BG5:BG39 | `SyainTB` | 社員テーブル（DB から取得） |
| BH3 | — | 最大件数（=30） |

（行番号 31〜33、37〜38 は主に空または内部作業領域）

#### 追加照射データ列ヘッダー（行17）

| 列 | ヘッダー | 内容 |
|---|---|---|
| B | 受付番号 | `UkeNoN` 参照 |
| G | 開始 RIC-NO | `RicNoSN` 参照 |
| K | 終了 RIC-NO | `RicNoEN` 参照 |
| O | 指定線量 | `SiteiSNN` 参照 |
| S | 下限線量 | — |
| W | 上限線量 | — |
| AA | 追加パス数 | ROUNDUP 計算 |
| AD | 会社名 | `KaiNameN` 参照 |

---

### 2.2 試験成績書用紙

**目的**: 「３号機線量不足報告書」プログラムのテスト仕様書と結果報告書を兼ねた内部資料シート。実運用には使用しない。

#### レイアウト構造

| セル範囲 | 内容 |
|---|---|
| F2 | タイトル：テスト仕様書タイトル |
| A10〜 | 機能追加テスト項目 |
| A12:AD12 | ヘッダー行（No / 検査項目 / 検査手順 / 確認方法 / 判定） |
| A15〜 | テストケース行（起動時・画面表示・データ入力・印刷 等） |

（行番号 2〜9 は空または見出し、60 行目以降は空）

---

### 2.3 改訂履歴

**目的**: 様式 G3-08 の改訂ログを管理する。VBA の `ブック保存` プロシージャ内でも参照される。

#### レイアウト構造

| セル範囲 | 名前付き範囲 | 内容 |
|---|---|---|
| A1 | — | 「様式 G3-08の改訂履歴」 |
| A3:D3 | — | 列ヘッダー（様式名 / 様式番号） |
| A4 | `YousikiName` | 様式名「３号機線量不足報告書」 |
| D4 | `YousikiNo` | 様式番号「G03-08」 |
| A5:D5 | — | ログヘッダー（版数 / 制定改定日 / 改訂箇所・理由 / 登録者） |
| A6:D112 | `Rireki` | 改訂履歴データ本体（版数0〜N） |

主な改訂履歴（抜粋）：

| 版数 | 改定日 | 改訂内容 |
|---|---|---|
| 0 | 1998-10-21 | 新規制定 |
| 1 | 1999/3/4 | 見直し改訂 |
| 2 | 2001/12/3 | 測定線量記入欄を設けた |
| 3 | 2002/10/22 | 追加製造判定を出荷判定に変更 |
| 4 | 2005/06/15 | γ照射課責任者を一般品と医療機器に分けた |

---

## 3. 名前付き範囲一覧

| 名前 | 参照先 | 説明 |
|---|---|---|
| `Atusa` | 報告書!$AJ$7 | 素子厚さ（ヘッダー部） |
| `AtusaT` | 報告書!$O$16 | 線量計厚さ（追加照射データ部） |
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
| `MokuSeu` | 報告書!#REF! | 参照エラー（削除済み可能性あり） |
| `mokuteki` | 報告書!$BC$3:$BD$6 | 製品目的テーブル（コード・説明） |
| `Pass` | 報告書!$AA$7 | 指定パス数 |
| `RicNoE1`〜`RicNoE5` | 報告書!$K$9〜$K$13 | 終了 RIC-NO（1〜5） |
| `RicNoS1`〜`RicNoS5` | 報告書!$G$9〜$G$13 | 開始 RIC-NO（1〜5） |
| `Rireki` | 改訂履歴!$A$6:$D$112 | 改訂履歴データ本体 |
| `SaiSoku` | 報告書!$O$7 | 再測定者 |
| `SebkNoTB` | 報告書!$BG$3 | 線量不足データ件数 |
| `SEnKind` | 報告書!$G$16 | 線量計種類 |
| `SenkNo` | 報告書!$B$7 | 照射管理番号（入力セル、ドロップダウン） |
| `SenSyu` | 報告書!$BC$10:$BC$12 | 線量計種類テーブル |
| `SikiKodo` | 報告書!$AD$7 | 計算式コード（測定データ） |
| `SiteiSN1`〜`SiteiSN5` | 報告書!$O$9〜$O$13 | 指定線量（1〜5） |
| `SokDate` | 報告書!$G$7 | 測定日 |
| `SoktCD` | 報告書!$K$7 | 測定者コード |
| `SokutSN` | 報告書!$S$7 | 測定線量 |
| `Sosi` | 報告書!$AG$7 | 素子異常 |
| `SyainTB` | 報告書!$BG$5:$BG$39 | 社員テーブル（DB取得後に動的リサイズ） |
| `Tani` | 報告書!$W$7 | 照射目的（単位） |
| `Tuika` | 報告書!$E$3 | 追加照射回目 |
| `TuikaMin` | 報告書!$BC$21 | 線量計高・低閾値（5） |
| `UkeNo1`〜`UkeNo5` | 報告書!$B$9〜$B$13 | 受付番号（1〜5） |
| `YousikiName` | 改訂履歴!$A$4 | 様式名 |
| `YousikiNo` | 改訂履歴!$D$4 | 様式番号 |

---

## 4. 数式一覧

### 報告書シート

| セル | 数式 | 説明 |
|---|---|---|
| B1 | `="様式 G3-08("&MAX(改訂履歴!A6:A112)&")"` | 改訂履歴から最新版数を動的に表示 |
| M3 | `=IF(SebkNoTB>BH3,"追加照射データが多過ぎるため管理者に連絡","")` | データ件数超過警告 |
| G15 | `=RIGHT(SenkNo,4)` | 照射管理番号の下4桁（線量計番号） |
| B18 | `=IF(UkeNo1="","",UkeNo1)` | 受付番号1（空白制御） |
| G18 | `=RicNoS1` | 開始 RIC-NO 参照 |
| K18 | `=RicNoE1` | 終了 RIC-NO 参照 |
| O18 | `=IF(B18="","",Kagen1-SokutSN)` | 線量不足量（下限線量 − 測定線量） |
| S18 | `=O18` | 線量不足量コピー |
| W18 | `=IF(B18="","",W9-$S$7)` | 上限線量からの余裕 |
| AA18 | `=IF(B18="","",IF(ROUNDUP(O18/($S$7/AA9),0)*$S$7/AA9-O18<2,ROUNDUP(O18/($S$7/AA9),0)+1,ROUNDUP(O18/($S$7/AA9),0)))` | 追加パス数計算（端数処理付き） |
| AD18 | `=KaiName1` | 会社名参照 |
| B19〜B22 | ～（B18と同パターン、UkeNo2〜5） | 受付番号2〜5 |
| O19〜O22 | ～（O18と同パターン、Kagen2〜5） | 線量不足量2〜5 |
| AA19〜AA22 | ～（AA18と同パターン） | 追加パス数2〜5 |
| M24 | `=IF(OR(AA18<0,AA19<0,AA20<0,AA21<0,AA22<0),"測定データが不適切","")` | パス数負値チェック |
| S24 | `=IF(MAX(AA18:AC22)<>MIN(AA18:AC22),"追加ﾊﾟｽ数異なるため課長に報告のこと","")` | パス数統一確認 |
| BF34 | `=#REF!` | 参照エラー（要確認） |
| A36〜AX36 | `=IF(ISERROR(VLOOKUP(SenkNo,Data,列番号,FALSE)),"",VLOOKUP(SenkNo,Data,列番号,FALSE))` | `SenkNo` をキーとしてデータ配列から列別に値を検索（37列分） |
| AQ39 | `=IF(ISERROR(VLOOKUP(TEXT(SenkNo,"0000"),Data,AQ37,FALSE)),"",VLOOKUP(TEXT(SenkNo,"0000"),Data,AQ37,FALSE))` | `SenkNo` をゼロ埋め4桁テキストに変換して VLOOKUP |

> **補足**: AA列の追加パス数計算式は「測定線量1パス分を基準に必要追加パスを切り上げ、余裕が2未満なら+1パス」するロジック。

---

## 5. ボタン・マクロ対応

ボタンオブジェクトは抽出スクリプトでは検出されなかった（シート保護により非表示またはコマンドバー経由の可能性あり）。

VBA コードでは以下のマクロがキーボードショートカット属性（`VB_ProcData.VB_Invoke_Func`）付きで定義されている：

| マクロ名 | ショートカット属性 | 処理概要 |
|---|---|---|
| `画面初期設定` | `\n14` | 起動時の初期化・DB取得 |
| `クリア` | `e\n14` | 入力フォームクリア |
| `CTRCls` | `e\n14` | `SenkNo` クリア後に`クリア`呼出 |

---

## 6. VBA モジュール仕様

### 6.1 ThisWorkbook.cls

#### `Workbook_Open()`

**処理概要**: ブック起動時に `画面初期設定` を呼び出す。

**処理フロー**:
1. `画面初期設定` プロシージャを呼び出す

```vba
Private Sub Workbook_Open()
    Call 画面初期設定
End Sub
```

---

#### `Workbook_BeforeClose(Cancel As Boolean)`

**処理概要**: `Debug` セルが空の場合は確認なしで閉じる。値がある場合はYes/No確認を表示する。

**処理フロー**:
1. `Debug` セルが空かチェック
2. 空 → アラート無効化 + `Saved=True` でサイレントクローズ
3. 値あり → 「終了しますか」ダイアログ表示、Noでキャンセル

```vba
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    If Range("Debug") = "" Then
        Application.DisplayAlerts = False
        ActiveWorkbook.Saved = True
    Else
        If MsgBox("終了しますか", vbYesNo) = vbNo Then Cancel = True
    End If
End Sub
```

---

#### `Workbook_BeforePrint(Cancel As Boolean)`

**処理概要**: 「報告書」シート印刷前に未入力チェックと線量計種類チェックを実施する。

**処理フロー**:
1. 対象シートが「報告書」か確認
2. `未入力チェック` で必須項目（再測定者・素子異常欄・線量計種類・厚さ）を検証
3. 未入力あり → エラーメッセージと「印刷しますか？」確認
4. 全項目入力済み → `線量計種類チェック` で線量計の高/低種類と測定値の整合確認
5. `印刷範囲の設定` を呼び出す

---

### 6.2 Sheet1.cls（報告書シート）

#### `Worksheet_Change(Target As Range)`

**処理概要**: 「線量計種類」（G16）または「照射管理番号」（B7）が変更された際にDBからデータを取得する。

**処理フロー**:
1. イベント無効化
2. 変更セルが `G16`（線量計種類）の場合 → `計算式コード取得` 呼び出し
3. 変更セルが `B7`（照射管理番号）の場合 → `線量不足データ抽出` 呼び出し
4. イベント再有効化

```vba
Private Sub Worksheet_Change(ByVal Target As Range)
    Application.EnableEvents = False
    With Target
        If .Row = 16 And .Column = 7 Then   '線量計種類
            If Cells(.Row, .Column) <> "" Then
                Call 計算式コード取得(Cells(.Row, .Column))
            Else
                Range("Keisask") = ""
            End If
        End If
        If .Row = 7 And .Column = 2 Then    '照射管理番号
            If Cells(.Row, .Column) <> "" Then
                Call 線量不足データ抽出
            End If
        End If
    End With
    Application.EnableEvents = True
End Sub
```

---

### 6.3 初期化.bas

#### `画面初期設定()`

**処理概要**: ブック起動時に画面をクリアし、DB から線量不足データ・社員データ・線量計種類を取得してドロップダウンリストを構築する。

**処理フロー**:
1. イベント無効化
2. `クリア` でシートをリセット
3. `線量不足線量計` でDB から照射管理番号リスト取得
4. `社員データ` で有資格社員一覧取得
5. `線量計種類` で使用線量計種類取得
6. `SenkNo` セル選択
7. シート保護（UIのみ）を設定しイベント再有効化

#### `クリア()`

**処理概要**: 報告書シートの入力エリアをクリアし、バリデーションを削除してズームリセットする。

**処理フロー**:
1. 「報告書」シートをアクティブ化・シート保護解除
2. `Data1`, `Data2`, `SEnKind`, `AtusaT`, `Keisask`, `SenSyu`, `SyainTB`, `Tuika`, `SebkNoTB`, `SenkNo` の入力検証・値をクリア
3. A1:AO1 範囲に合わせてズーム設定

#### `CTRCls()`

**処理概要**: `SenkNo` をクリアして `クリア` を呼び出す。

---

### 6.4 データ抽出SQL.bas

#### `線量不足線量計()`

**処理概要**: 線量不足フラグが立っている全照射管理番号を DB から抽出し、`SenkNo` セルのドロップダウンリストに設定する。

**処理フロー**:
1. `SYOUJ2` と `SYOUKJ3` を結合して照射完了かつ線量不足の管理番号を取得
2. 各管理番号について追加照射有効かを `SYOUJ2` から確認
3. 有効な管理番号でドロップダウンリストを構築（`SenkNo.Validation`）
4. 件数を `SebkNoTB` に設定

```sql
-- 線量不足の照射管理番号抽出
SELECT DISTINCT sk.SYKNO, sk.uno1, sk.uno2, sk.uno3, sk.uno4, sk.uno5
FROM RIC.SYOUJ2 s, RIC.SYOUKJ3 sk
WHERE s.UNO = sk.UNO1 AND s.FRICNO = sk.FRICNO1
  AND sk.SYOUSH_F='2'
  AND (s.SENHFLG='2' OR s.SENHFLG='4') AND s.syouflg='1'
  AND s.seshflg='1'
ORDER BY sk.sykno
```

#### `社員データ()`

**処理概要**: 線量測定有資格社員の氏名一覧を取得し `SyainTB` に書き込む。

```sql
SELECT TRIM(SHANAME)
FROM SHAINMST
WHERE HSHIKA='1' AND LSHIKA='1' AND shano>'1000'
ORDER BY SHANO
```

#### `線量計種類()`

**処理概要**: 使用可能な線量計種類を `senkind` テーブルから取得し `SenSyu` に書き込む。

```sql
SELECT sensyu FROM senkind ORDER BY dispno
```

#### `計算式コード取得(SenKind)`

**処理概要**: 選択された線量計種類に対応する計算式コードを `SENKIND` → `KEICODE` テーブルから取得し `Keisask` に設定する。

**処理フロー**:
1. `SENKIND` テーブルから `KKODE`（計算式コード記号）取得
2. `KEICODE` テーブルから有効フラグ=1 の計算式コード取得
3. コードの2文字目が "3"（RIC3対応）かつ先頭がKKODEと一致するものを `Keisask` に設定

```sql
SELECT KKODE FROM SENKIND WHERE sensyu='[SenKind]'
SELECT KEISASK FROM KEICODE WHERE YFLG1='1'
```

#### `線量不足データ抽出()`

**処理概要**: 選択された照射管理番号（`SenkNo`）に対応する測定実績データと照射計画データを DB から取得してシートに書き込む。

**処理フロー**:
1. `SYOUJ2`, `SHAINMST`, `ZAIKO` を結合して測定データ取得（測定日・測定者・測定線量・製品目的・パス数・計算コード・厚さ・実パス・追加フラグ）
2. シートの `Data1` エリアをクリア
3. 測定データをシートに書き込み
4. `SYOUKJ3` から照射計画データ（最大5受付分）取得
5. 線量不足（下限線量 > 測定線量）の受付データのみ `Data2` エリアに書き込み

```sql
-- 測定実績取得
SELECT S.SESDATE, sha.shaname,
       TO_NUMBER(S.SOKUTSN), TO_NUMBER(za.tani), TO_NUMBER(S.PASS),
       S.KEISASK, TO_NUMBER(S.ATUSA), TO_NUMBER(S.JITUNO),
       TO_NUMBER(TUIKAFLG)
FROM SHAINMST SHA, RIC.SYOUJ2 S, RIC.ZAIKO ZA
WHERE s.sokutcd=sha.shano AND za.uno=s.uno
  AND (s.senhflg='2' OR s.senhflg='4')
  AND s.senkno='[SenkNo 下4桁]'
ORDER BY s.tuikaflg DESC

-- 照射計画取得
SELECT TO_NUMBER(UNO1),FRICNO1,LRICNO1,SITEISN1,KAGENSN1,JYOUGSN1,kaisyacd1,KAINAME1,
       TO_NUMBER(UNO2),...（UNO5 まで同パターン）
FROM RIC.SYOUKJ3
WHERE SYKNO = '[SenkNo]' AND SYOUSH_F='2'
```

#### `製品目的(myMokNo) As String`

**処理概要**: `MokuCD` テーブル（`BC3:BD6`）からコードに対応する製品目的文字列を検索して返す。

---

### 6.5 SQL_Execution.bas

**処理概要**: ADODB を使用した Oracle DB 接続の共通ライブラリ。全 SQL 実行はこのモジュールを経由する。

#### `Open_oraconDB()`

DB 接続文字列: `DSN=ricdb;UID=ric;PWD=t6101`（ODBC 経由）

```vba
Sub Open_oraconDB()
    mpDSN = "DSN=ricdb;UID=ric;PWD=t6101"
    oraconn.ConnectionString = mpDSN
    oraconn.Open
    oraconn.CursorLocation = adUseClient
End Sub
```

#### `SQL_Exe(mySql)`

SQL 文を `oraconn.Execute` で実行。エラー時は `MsgBox` でエラー内容表示後 `End` で強制終了。

#### `SQL_INSERT_UPDATE(myTBL, myKey, myD(), mYN)`

キー条件で件数チェックし、0件なら `INSERT`、1件以上なら `UPDATE` を実行。トランザクション付き。

#### `SQL_Delete(myTBL, myWhere)`

指定テーブル・WHERE条件で `DELETE` を実行。トランザクション付き。

#### `Disp_Sheet(mySql, mySH, myRow, myRecordCount, myColumn, myFieldCount, myF)`

SQL を実行してレコードセットを指定シートの指定行・列から `CopyFromRecordset` で書き込む。

#### `Set_Array(mySql, myData(), myRecordCount, myFldCount)`

SQL を実行してレコードセットを `myData(i, j)` の2次元配列に格納する。

---

### 6.6 印刷.bas

#### `印刷範囲の設定()`

**処理概要**: 「報告書」シートの印刷範囲を A4 1ページ（`$B$1:$AN$30`）に設定する。

```vba
Sub 印刷範囲の設定()
    With Worksheets("報告書").PageSetup
        .Zoom = False
        .PrintArea = "$B$1:$AN$30"
        .FitToPagesWide = 1
        .FitToPagesTall = 1
    End With
End Sub
```

---

### 6.7 改定履歴.bas

#### `ブック保存()`

**処理概要**: DB の `ExSeihinJ` テーブルからブック名に対応する保存フォルダを取得し、サーバー上の指定パスに `SaveAs` で保存する。

**処理フロー**:
1. `ExSeihinJ` テーブルから `filename=ThisWorkbook.Name` に対応する `folder` パスを取得
2. 「保存しますか？」確認
3. `CTRCls` でフォームクリア
4. `SaveAs` でサーバーパスに保存

```sql
SELECT folder FROM ExSeihinJ WHERE filename='[ThisWorkbook.Name]'
```

---

### 6.8 Debug用.bas

#### `DebugCom()`

**処理概要**: `EnableEvents = True` に戻す緊急復旧用プロシージャ。

---

## 7. DB 接続・外部連携

### ODBC 接続設定

| DSN 名 | UID | PWD | 用途 |
|---|---|---|---|
| `ricdb` | ric | t6101 | Oracle DB（照射管理システム）への接続 |

### 参照テーブル一覧

| テーブル名 | 主な用途 | 主要カラム |
|---|---|---|
| `RIC.SYOUJ2` | 照射実績（線量不足判定含む） | UNO, SENKNO, SENHFLG, SOKUTSN, SESDATE, SOKUTCD, TUIKAFLG |
| `RIC.SYOUKJ3` | 照射計画（最大5受付） | SYKNO, UNO1〜5, FRICNO1〜5, LRICNO1〜5, SITEISN1〜5, KAGENSN1〜5, KAINAME1〜5 |
| `SHAINMST` | 社員マスタ | SHANO, SHANAME, HSHIKA, LSHIKA |
| `RIC.ZAIKO` | 在庫マスタ（製品目的コード） | UNO, TANI |
| `SENKIND` | 線量計種類マスタ | SENSYU, KKODE, DISPNO |
| `KEICODE` | 計算式コードマスタ | KEISASK, YFLG1 |
| `ExSeihinJ` | ファイル保存先テーブル | FILENAME, FOLDER |

### 主要 SQL 文

```sql
-- 線量不足照射管理番号の抽出（線量不足線量計）
SELECT DISTINCT sk.SYKNO, sk.uno1, sk.uno2, sk.uno3, sk.uno4, sk.uno5
FROM RIC.SYOUJ2 s, RIC.SYOUKJ3 sk
WHERE s.UNO = sk.UNO1 AND s.FRICNO = sk.FRICNO1
  AND sk.SYOUSH_F='2'
  AND (s.SENHFLG='2' OR s.SENHFLG='4') AND s.syouflg='1'
  AND s.seshflg='1'
ORDER BY sk.sykno

-- 測定実績データ取得（線量不足データ抽出）
SELECT S.SESDATE, sha.shaname, TO_NUMBER(S.SOKUTSN),
       TO_NUMBER(za.tani), TO_NUMBER(S.PASS),
       S.KEISASK, TO_NUMBER(S.ATUSA), TO_NUMBER(S.JITUNO),
       TO_NUMBER(TUIKAFLG)
FROM SHAINMST SHA, RIC.SYOUJ2 S, RIC.ZAIKO ZA
WHERE s.sokutcd=sha.shano AND za.uno=s.uno
  AND (s.senhflg='2' OR s.senhflg='4')
  AND s.senkno='[SenkNo下4桁]'
ORDER BY s.tuikaflg DESC

-- 照射計画取得（線量不足データ抽出）
SELECT TO_NUMBER(UNO1), FRICNO1, LRICNO1, SITEISN1, KAGENSN1, JYOUGSN1, kaisyacd1, KAINAME1,
       ...(UNO5まで同パターン)
FROM RIC.SYOUKJ3
WHERE SYKNO = '[SenkNo]' AND SYOUSH_F='2'

-- ブック保存先フォルダ取得
SELECT folder FROM ExSeihinJ WHERE filename='[FileName]'
```

---

## 8. データフロー

```
【起動フロー】
  Workbook_Open
       ↓
  画面初期設定
       ↓
  クリア（シートリセット）
       ↓
  DB(SYOUJ2/SYOUKJ3): 線量不足照射管理番号取得 → SenkNo ドロップダウン設定
       ↓
  DB(SHAINMST): 有資格社員取得 → SyainTB 書き込み
       ↓
  DB(SENKIND): 線量計種類取得 → SenSyu 書き込み
       ↓
  シート保護（UI のみ）設定

【データ選択フロー】
  SenkNo ドロップダウンで照射管理番号選択
       ↓ Worksheet_Change(B7)
  線量不足データ抽出
       ↓
  DB(SYOUJ2 + SHAINMST + ZAIKO): 測定実績取得 → Data1 書き込み
       ↓
  DB(SYOUKJ3): 照射計画取得 → Data2 書き込み（線量不足分のみ）
       ↓
  シートの数式が追加パス数・線量不足量を自動計算

【線量計種類変更フロー】
  SEnKind(G16) に線量計種類入力
       ↓ Worksheet_Change(G16)
  計算式コード取得
       ↓
  DB(SENKIND + KEICODE): 計算式コード → Keisask 設定

【印刷フロー】
  Workbook_BeforePrint
       ↓
  未入力チェック（SaiSoku / Sosi / SEnKind / AtusaT）
       ↓ NG → 確認ダイアログ
  線量計種類チェック（高/低線量計と測定値の整合確認）
       ↓
  印刷範囲の設定（B1:AN30 に固定）
       ↓
  印刷実行

【保存フロー】
  ブック保存（改定履歴.bas）
       ↓
  DB(ExSeihinJ): 保存先フォルダ取得
       ↓
  CTRCls でフォームクリア
       ↓
  SaveAs でサーバーパスに上書き保存
```

---

## 9. セキュリティ注意事項

| 種別 | キーワード | 内容 |
|---|---|---|
| AutoExec | `Workbook_Open` | ブックを開くと自動で `画面初期設定` が実行される。DB 接続が即座に行われる。 |
| AutoExec | `Workbook_BeforeClose` | `Debug` セルが空の場合はアラートなしに保存状態を偽装してクローズする。 |
| AutoExec | `Worksheet_Change` | セル変更（B7・G16）で自動的に SQL が実行される。 |
| Suspicious | `Open` / `SaveAs` | `ブック保存` プロシージャでサーバー上のパスにファイルを書き込む。 |
| Suspicious | `Call` | 外部 DLL 呼び出しと同等の動作をする Excel 4 Macro として検知された。 |
| Suspicious | `Chr` | `Chr(13)` による改行文字の動的生成。可読性低下だが悪意なし（標準的な VBA 記述）。 |
| Suspicious | `Hex Strings` | VBA プロジェクト内に16進エンコード文字列が検出。内容精査が望ましい。 |
| Suspicious | `Base64 Strings` | Base64 エンコード文字列が検出。内容精査が望ましい。 |
| 認証情報 | `PWD=t6101` | DB パスワードが VBA コードにハードコードされている。運用環境では変更・管理が必要。 |
| 参照エラー | `MokuSeu = #REF!` | 名前付き範囲 `MokuSeu` が壊れている。動作に影響がないか確認が必要。 |
| 参照エラー | `BF34 = #REF!` | 報告書シートの BF34 セルに参照エラーがある。 |
