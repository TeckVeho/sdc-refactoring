# ExMenu 仕様書

> **対象ファイル**: ExMenu.xlsm
> **ファイル種別**: .xlsm（マクロ付き）
> **用途**: EXメニュー（照射管理業務 VBA アプリ群）のメイン起動画面。最大 30 件の Excel アプリを登録・管理し、ボタンクリックで起動する
> **VBA プロジェクト**: モジュール 10 本（.bas 6 / .cls 3 / .frm 1）
> **外部連携**: DSN=`ricdb`（Oracle DB）、DB 接続先 IP: 163.59.144.156（コメント上の旧パス参照あり。現在は `C:\ラジエ工業\ExRicSys\`〈EXメニュー配置フォルダ〉を使用）
> **解析日**: 2026-06-24（excel-to-md スキルによる自動解析）

---

## 凡例（本仕様書の表記ルール）

本仕様書では、保守時の判別を容易にするため、以下の表記ルールを使用します。

### 用語規約

| 用語 | 意味 |
| --- | --- |
| EXメニュー | ExRicSys フォルダに配置される VBA ファイル群（Ex*.xlsm）の総称 |
| `ExRicSys` | 配置フォルダ名。初出時は〈EXメニュー配置フォルダ〉と注記 |

### 表記規則

| 種別 | 表記 | 例 |
| --- | --- | --- |
| モジュール（.bas / .cls） | **太字** | **共通変数と開始処理.bas** |
| ユーザーフォーム（.frm） | **太字** | **SikakuKakuninn.frm** |
| プロシージャ / イベント | `コード体()` | `初期設定()` ※Markdown 制約上、斜体ではなくコード体で統一 |
| シート名 | 「」 | 「メニュー」 |
| セル参照 | `コード体` | `$F$2` |
| 名前付き範囲 | `コード体` | `PGList` |
| DB テーブル / カラム | `コード体` | `ExSEIHINJ` / `KAISYACD` |
| ユーザー操作 | （操作名） | （ファイル登録 Click） |
| 主要マーク | ✓ | ✓ = 保守時に最初に確認すべき項目 |

### データフロー 場所マーク（9章）

9章のデータフロー（テーブル・ツリー図）では、処理が行われる場所を以下のアイコンで区別します。

| アイコン | 種別 | 意味 |
| --- | --- | --- |
| 📊 | シート操作 | ワークシート上のセル書込み・読取り・表示変更 |
| 🖥️ | 画面操作 | ユーザーフォーム（.frm）の表示・入力・操作 |
| 🗄️ | DB操作 | DB への SELECT / INSERT / UPDATE / DELETE |
| 📄 | VBA内部処理 | 変数計算・条件分岐など、画面・シートに直接関与しない処理 |

### ✓（主要マーク）の判定基準

✓ は **保守時に最初に確認すべき項目** を示します。
判定基準は対象の種類ごとに以下のとおりです。

| 章 | 対象 | ✓ の判定基準 |
| --- | --- | --- |
| 1.1 | シート | ユーザーが直接操作する、または VBA が動的に表示/非表示を切り替える |
| 1.2 | ユーザーフォーム | ユーザー入力を受け付ける、または業務フローの起点となる |
| 1.3 / 6.0 | VBA モジュール | ① ユーザー操作の起点 ② DB I/O を含む ③ 他モジュールから呼び出される ④ コード行数上位 25% のいずれか |
| 2 | セル / 名前付き範囲 | VBA から `Range()` で代入または参照され、業務ロジックに直結する |
| 3 | 名前付き範囲 | VBA から `Range()` で代入または参照され、業務ロジックに直結する |
| 5 | ボタン / コントロール | DB 更新・画面遷移・計算実行など副作用のある操作を起動する |
| 6.0（全プロシージャ） | プロシージャ | ① ユーザー操作の起点（Click イベント等） ② DB I/O を実行 ③ 他モジュールから呼び出される Public のいずれか |
| 7 | フォームコントロール | ユーザー入力を受け付ける、またはイベントで業務処理を起動する |
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

### 1.1 シート一覧

> ✓ = ユーザーが直接操作する、または VBA が動的に表示/非表示を切り替えるシート

| ✓ | No | シート名 | 最大行 | 最大列 | 保存時 Visible | VBA による動的切替 |
| --- | --- | --- | --- | --- | --- | --- |
| ✓ | 1 | メニュー | 24 | 13 (M) | visible | —（シート自体の Visible は変更しない。ActiveWindow の見出し/グリッド等を非表示化） |
| ✓ | 2 | AprList | 38 | 15 (O) | visible | `AddSoft()` で「メニュー」から画面遷移（Select）/ `BackMenu()` で「メニュー」へ戻る |

### 1.2 ユーザーフォーム一覧

> ✓ = ユーザー入力を受け付ける、または業務フローの起点となるフォーム

| ✓ | No | フォーム名 | コントロール数 | 主用途 |
| --- | --- | --- | --- | --- |
| ✓ | 1 | **SikakuKakuninn.frm** | 6 | メニュー登録時の資格確認（社員番号/パスワード入力） |

### 1.3 VBA モジュール一覧

> ✓ = ユーザー操作の起点 / DB I/O を含む / 他モジュールから呼び出される / コード行数上位 25%

| ✓ | No | モジュール | 種別 | プロシージャ数 | 主な役割 |
| --- | --- | --- | --- | --- | --- |
| ✓ | 1 | **ThisWorkbook.cls** | .cls | 3 | 起動/終了/保存イベント |
|  | 2 | **Sheet3.cls** | .cls | 0 | 「メニュー」シート（コードなし） |
|  | 3 | **Sheet5.cls** | .cls | 0 | 「AprList」シート（コードなし） |
| ✓ | 4 | **共通変数と開始処理.bas** | .bas | 2 | 起動時初期化・パス取得 |
| ✓ | 5 | **メニュー選択実行.bas** | .bas | 31 | Menu1〜30 起動・二重起動チェック |
| ✓ | 6 | **SQL_Execution.bas** | .bas | 6 | ADO/ODBC 接続・SQL 実行基盤 |
| ✓ | 7 | **ファイル登録_取込.bas** | .bas | 5 | DB 取込・メニュー登録・画面遷移 |
| ✓ | 8 | **お知らせメッセージ.bas** | .bas | 2 | 起動時お知らせ表示 |
|  | 9 | **Debug用.bas** | .bas | 2 | デバッグ用メニュークリア |
| ✓ | 10 | **SikakuKakuninn.frm** | .frm | 5 | 資格確認 UserForm |

---

## 2. シート詳細

### 2.0 シート可視性一覧

| No | シート | VBA による非表示化 | 表示するタイミング | 非表示にするタイミング | 制御プロシージャ |
| --- | --- | --- | --- | --- | --- |
| 1 | メニュー | シート自体は常時 visible。ActiveWindow の見出し・スクロールバー・タブ・グリッドを非表示 | `初期設定()` 実行後 | `AllCls()` 実行時（Debug 用） | **共通変数と開始処理.bas** / **Debug用.bas** |
| 2 | AprList | シート自体は常時 visible だが通常操作時は「メニュー」が前面 | 「メニュー」シートの【ファイルの追加・消去】クリック時（`AddSoft()`） | `BackMenu()` で「メニュー」シートへ戻る | **ファイル登録_取込.bas**（`AddSoft()` / `BackMenu()`） |

> 以下の各シートのレイアウト構造表における ✓ = VBA から `Range()` で代入または参照され、業務ロジックに直結するセル

### 2.1 メニュー

**目的**: 業務アプリ起動のメイン画面。Shape オブジェクト Menu1〜Menu30 に登録アプリ名を表示し、クリックで各アプリを起動する。ヘッダー・グリッド・スクロールバー・タブは VBA で非表示にし、専用メニュー画面として機能する。

#### 非表示行・列

なし。（行/列の非表示ではなく、ActiveWindow の UI 要素を非表示化）

#### レイアウト構造

```
行 1: デバッグフラグ・日時表示
行 2: メニューバージョン表示
オブジェクト Menu1〜Menu30: アプリ起動ボタン（Shape。VBA がラベルと OnAction を動的設定）
```

| ✓ | No | セル / オブジェクト | 名前付き範囲 | 種別 | 実態（値/数式/VBA代入） | 業務的意味 |
| --- | --- | --- | --- | --- | --- | --- |
| ✓ | 1 | `A1` | `Debug` | 手動設定 | 空欄 or `*` | デバッグモードフラグ（空=通常、`*`=非表示モード。`AllCls` で参照） |
|  | 2 | `B2` | — | 設定値 | `V3M` | メニューバージョン表示 |
|  | 3 | `K1` | — | 数式 | `=NOW()` | 現在日時表示 |
| ✓ | 4 | Menu1〜Menu30 | — | VBA代入 | `初期設定()` で `mpAprD(i,2)` を Characters.Text に設定 | 各スロットのアプリ表示名。`-----` は利用不可 |
| ✓ | 5 | Menu1〜Menu30 | — | VBA代入 | `初期設定()` で OnAction=`選択N` を設定 | クリック時に `選択N()` を実行 |

> ボタンのラベルが `-----` の場合: 非 RIC サーバー接続 PC、またはファイル未登録/未存在（クリックしても起動失敗またはサイレント無視）

### 2.2 AprList

**目的**: 起動アプリの登録・管理テーブル。DB（ExSeihinj）から読み込んだ情報（I〜N 列）と管理者が直接編集する情報（C〜H 列）を比較し、変更検出（O 列）を行う。

#### 非表示行・列

なし。

#### レイアウト構造

```
行 2: タイトル・登録ファイル数
行 3: 列ヘッダー（B3:O3）
行 4〜33: メニューNo 1〜30 の登録データ（30 行固定）
行 37〜38: B37:B38 マージセル（備考/フッター領域）
```

| ✓ | No | セル | 名前付き範囲 | 種別 | 実態（値/数式/VBA代入） | 業務的意味 |
| --- | --- | --- | --- | --- | --- | --- |
|  | 1 | `B2` | — | 設定値 | `ﾌｧｲﾙ一覧` | シートタイトル |
| ✓ | 2 | `F2` | `AprSuu` | 数式 | `=COUNTA(D4:D33)` | 登録ファイル数（ファイル名が入っている行数） |
|  | 3 | `B3:O3` | — | 設定値 | 列ヘッダー | 各列の意味を示す見出し |
|  | 4 | `B4:B33` | — | 設定値 | 1〜30 | メニュー番号 |
| ✓ | 5 | `C4:C33` | `PGList` 先頭列 | 手動設定 / DB読込 | Path 名（50B 以内） | アプリ格納フォルダパス |
| ✓ | 6 | `D4:D33` | — | 手動設定 / DB読込 | ファイル名（50B 以内） | 起動対象 Excel ファイル名 |
| ✓ | 7 | `E4:E33` | — | 手動設定 / DB読込 | メニュー表示文字（100B 以内） | メニューボタンに表示する名称 |
| ✓ | 8 | `F4:F33` | — | 手動設定 / DB読込 | ソフト種別（20B 以内） | アプリ分類 |
| ✓ | 9 | `G4:G33` | — | 手動設定 / VBA代入 | 登録/非表示日 | `メニュー登録()` で `Now` を書込み |
| ✓ | 10 | `H4:H33` | — | 手動設定 / VBA代入 | 登録者（20B 以内） | 資格確認後の社員名 |
| ✓ | 11 | `I4:I33` | — | DB読込 | Path 名 | `メニュー取込()` が ExSeihinj から書込み |
| ✓ | 12 | `J4:J33` | — | DB読込 | ファイル名 | 同上 |
| ✓ | 13 | `K4:K33` | — | DB読込 | メニュー表示文字 | 同上 |
| ✓ | 14 | `L4:L33` | — | DB読込 | ソフト種別 | 同上 |
| ✓ | 15 | `M4:M33` | — | DB読込 | 登録日 | 同上 |
| ✓ | 16 | `N4:N33` | — | DB読込 | 登録者 | 同上 |
| ✓ | 17 | `O4:O33` | — | 数式 | `=IF(C& D& E& F <> I& J& K& L, FALSE, TRUE)` | 変更有無。False=変更あり（C〜F と I〜L の比較） |

> OpenPyXL 読取時、AprList の日本語セル値が文字化けする場合がある（Shift-JIS/UTF-8 混在）。実ファイル上の表示を正とする。

---

## 3. 名前付き範囲一覧


> ✓ = VBA から `Range()` で代入または参照され、業務ロジックに直結する名前付き範囲

| ✓ | No | 名前 | 参照先 | 業務的意味 |
| --- | --- | --- | --- | --- |
| ✓ | 1 | `AprSuu` | AprList!$F$2 | 登録可能アプリ数の表示セル（`=COUNTA(D4:D33)`） |
| ✓ | 2 | `Debug` | メニュー!$A$1 | デバッグモードフラグ。`AllCls` / `Workbook_BeforeSave`(TEST) で参照 |
| ✓ | 3 | `PGList` | AprList!$C$4:$N$33 | アプリ一覧データ範囲。`初期設定` / `メニュー取込` / `メニュー登録Main` / `二重起動チェックと起動` が Row/Column オフセットで参照 |

---


### 3.1 データの入力規則

なし。

## 4. 数式一覧

| シート | 数式件数 | 備考 |
| --- | --- | --- |
| AprList | 31 | 登録件数 1 + 変更検出 30 |
| メニュー | 1 | 現在日時表示 |

### 4.1 AprList

| セル | 数式 | 説明 |
| --- | --- | --- |
| `F2` | `=COUNTA(D4:D33)` | ファイル名（D 列）が入力されている行数をカウント |
| `O4` | `=IF(C4&D4&E4&F4<>I4&J4&K4&L4,FALSE,TRUE)` | C〜F 列（編集値）と I〜L 列（DB 値）の差分検出。False=変更あり |
| `O5` | `=IF(C5&D5&E5&F5<>I5&J5&K5&L5,FALSE,TRUE)` | C〜F 列（編集値）と I〜L 列（DB 値）の差分検出。False=変更あり |
| `O6` | `=IF(C6&D6&E6&F6<>I6&J6&K6&L6,FALSE,TRUE)` | C〜F 列（編集値）と I〜L 列（DB 値）の差分検出。False=変更あり |
| `O7` | `=IF(C7&D7&E7&F7<>I7&J7&K7&L7,FALSE,TRUE)` | C〜F 列（編集値）と I〜L 列（DB 値）の差分検出。False=変更あり |
| `O8` | `=IF(C8&D8&E8&F8<>I8&J8&K8&L8,FALSE,TRUE)` | C〜F 列（編集値）と I〜L 列（DB 値）の差分検出。False=変更あり |
| `O9` | `=IF(C9&D9&E9&F9<>I9&J9&K9&L9,FALSE,TRUE)` | C〜F 列（編集値）と I〜L 列（DB 値）の差分検出。False=変更あり |
| `O10` | `=IF(C10&D10&E10&F10<>I10&J10&K10&L10,FALSE,TRUE)` | C〜F 列（編集値）と I〜L 列（DB 値）の差分検出。False=変更あり |
| `O11` | `=IF(C11&D11&E11&F11<>I11&J11&K11&L11,FALSE,TRUE)` | C〜F 列（編集値）と I〜L 列（DB 値）の差分検出。False=変更あり |
| `O12` | `=IF(C12&D12&E12&F12<>I12&J12&K12&L12,FALSE,TRUE)` | C〜F 列（編集値）と I〜L 列（DB 値）の差分検出。False=変更あり |
| `O13` | `=IF(C13&D13&E13&F13<>I13&J13&K13&L13,FALSE,TRUE)` | C〜F 列（編集値）と I〜L 列（DB 値）の差分検出。False=変更あり |
| `O14` | `=IF(C14&D14&E14&F14<>I14&J14&K14&L14,FALSE,TRUE)` | C〜F 列（編集値）と I〜L 列（DB 値）の差分検出。False=変更あり |
| `O15` | `=IF(C15&D15&E15&F15<>I15&J15&K15&L15,FALSE,TRUE)` | C〜F 列（編集値）と I〜L 列（DB 値）の差分検出。False=変更あり |
| `O16` | `=IF(C16&D16&E16&F16<>I16&J16&K16&L16,FALSE,TRUE)` | C〜F 列（編集値）と I〜L 列（DB 値）の差分検出。False=変更あり |
| `O17` | `=IF(C17&D17&E17&F17<>I17&J17&K17&L17,FALSE,TRUE)` | C〜F 列（編集値）と I〜L 列（DB 値）の差分検出。False=変更あり |
| `O18` | `=IF(C18&D18&E18&F18<>I18&J18&K18&L18,FALSE,TRUE)` | C〜F 列（編集値）と I〜L 列（DB 値）の差分検出。False=変更あり |
| `O19` | `=IF(C19&D19&E19&F19<>I19&J19&K19&L19,FALSE,TRUE)` | C〜F 列（編集値）と I〜L 列（DB 値）の差分検出。False=変更あり |
| `O20` | `=IF(C20&D20&E20&F20<>I20&J20&K20&L20,FALSE,TRUE)` | C〜F 列（編集値）と I〜L 列（DB 値）の差分検出。False=変更あり |
| `O21` | `=IF(C21&D21&E21&F21<>I21&J21&K21&L21,FALSE,TRUE)` | C〜F 列（編集値）と I〜L 列（DB 値）の差分検出。False=変更あり |
| `O22` | `=IF(C22&D22&E22&F22<>I22&J22&K22&L22,FALSE,TRUE)` | C〜F 列（編集値）と I〜L 列（DB 値）の差分検出。False=変更あり |
| `O23` | `=IF(C23&D23&E23&F23<>I23&J23&K23&L23,FALSE,TRUE)` | C〜F 列（編集値）と I〜L 列（DB 値）の差分検出。False=変更あり |
| `O24` | `=IF(C24&D24&E24&F24<>I24&J24&K24&L24,FALSE,TRUE)` | C〜F 列（編集値）と I〜L 列（DB 値）の差分検出。False=変更あり |
| `O25` | `=IF(C25&D25&E25&F25<>I25&J25&K25&L25,FALSE,TRUE)` | C〜F 列（編集値）と I〜L 列（DB 値）の差分検出。False=変更あり |
| `O26` | `=IF(C26&D26&E26&F26<>I26&J26&K26&L26,FALSE,TRUE)` | C〜F 列（編集値）と I〜L 列（DB 値）の差分検出。False=変更あり |
| `O27` | `=IF(C27&D27&E27&F27<>I27&J27&K27&L27,FALSE,TRUE)` | C〜F 列（編集値）と I〜L 列（DB 値）の差分検出。False=変更あり |
| `O28` | `=IF(C28&D28&E28&F28<>I28&J28&K28&L28,FALSE,TRUE)` | C〜F 列（編集値）と I〜L 列（DB 値）の差分検出。False=変更あり |
| `O29` | `=IF(C29&D29&E29&F29<>I29&J29&K29&L29,FALSE,TRUE)` | C〜F 列（編集値）と I〜L 列（DB 値）の差分検出。False=変更あり |
| `O30` | `=IF(C30&D30&E30&F30<>I30&J30&K30&L30,FALSE,TRUE)` | C〜F 列（編集値）と I〜L 列（DB 値）の差分検出。False=変更あり |
| `O31` | `=IF(C31&D31&E31&F31<>I31&J31&K31&L31,FALSE,TRUE)` | C〜F 列（編集値）と I〜L 列（DB 値）の差分検出。False=変更あり |
| `O32` | `=IF(C32&D32&E32&F32<>I32&J32&K32&L32,FALSE,TRUE)` | C〜F 列（編集値）と I〜L 列（DB 値）の差分検出。False=変更あり |
| `O33` | `=IF(C33&D33&E33&F33<>I33&J33&K33&L33,FALSE,TRUE)` | C〜F 列（編集値）と I〜L 列（DB 値）の差分検出。False=変更あり |

### 4.2 メニュー

| セル | 数式 | 説明 |
| --- | --- | --- |
| `K1` | `=NOW()` | 現在日時を表示 |

---

## 5. ボタン・マクロ対応

> ✓ = DB 更新・画面遷移・計算実行など副作用のある操作を起動するボタン

### 5.1 シート上のボタン（Form Control / Shape）

#### AprList（Form Control）

| ✓ | No | シート | ボタンラベル | 割り当てマクロ | 動作概要 |
| --- | --- | --- | --- | --- | --- |
| ✓ | 1 | AprList | ファイル登録 | `メニュー登録Main()` | 変更あり行を ExSEIHINJ に登録/消去 |
| ✓ | 2 | AprList | メニュー画面 | `BackMenu()` | 「メニュー」シートへ戻る |

#### メニュー（Shape / FormControl）

##### 【ファイルの追加・消去】ボタン

| ✓ | No | シート | ボタンラベル | 割り当てマクロ | 動作概要 |
| --- | --- | --- | --- | --- | --- |
| ✓ | 1 | メニュー | ファイルの追加・消去 | `AddSoft()` | 「AprList」シートへ画面遷移し、ファイル登録・管理画面を表示 |

##### Menu1〜Menu30（Shape）

`初期設定()` 実行時に VBA が Characters.Text と OnAction を動的設定する。

| ✓ | No | シート | ボタンラベル | 割り当てマクロ | 動作概要 |
| --- | --- | --- | --- | --- | --- |
| ✓ | 1 | メニュー | Menu1 | `選択1()` | AprList No.1 のアプリを起動（`二重起動チェックと起動`） |
| ✓ | 2 | メニュー | Menu2 | `選択2()` | AprList No.2 のアプリを起動（`二重起動チェックと起動`） |
| ✓ | 3 | メニュー | Menu3 | `選択3()` | AprList No.3 のアプリを起動（`二重起動チェックと起動`） |
| ✓ | 4 | メニュー | Menu4 | `選択4()` | AprList No.4 のアプリを起動（`二重起動チェックと起動`） |
| ✓ | 5 | メニュー | Menu5 | `選択5()` | AprList No.5 のアプリを起動（`二重起動チェックと起動`） |
| ✓ | 6 | メニュー | Menu6 | `選択6()` | AprList No.6 のアプリを起動（`二重起動チェックと起動`） |
| ✓ | 7 | メニュー | Menu7 | `選択7()` | AprList No.7 のアプリを起動（`二重起動チェックと起動`） |
| ✓ | 8 | メニュー | Menu8 | `選択8()` | AprList No.8 のアプリを起動（`二重起動チェックと起動`） |
| ✓ | 9 | メニュー | Menu9 | `選択9()` | AprList No.9 のアプリを起動（`二重起動チェックと起動`） |
| ✓ | 10 | メニュー | Menu10 | `選択10()` | AprList No.10 のアプリを起動（`二重起動チェックと起動`） |
| ✓ | 11 | メニュー | Menu11 | `選択11()` | AprList No.11 のアプリを起動（`二重起動チェックと起動`） |
| ✓ | 12 | メニュー | Menu12 | `選択12()` | AprList No.12 のアプリを起動（`二重起動チェックと起動`） |
| ✓ | 13 | メニュー | Menu13 | `選択13()` | AprList No.13 のアプリを起動（`二重起動チェックと起動`） |
| ✓ | 14 | メニュー | Menu14 | `選択14()` | AprList No.14 のアプリを起動（`二重起動チェックと起動`） |
| ✓ | 15 | メニュー | Menu15 | `選択15()` | AprList No.15 のアプリを起動（`二重起動チェックと起動`） |
| ✓ | 16 | メニュー | Menu16 | `選択16()` | AprList No.16 のアプリを起動（`二重起動チェックと起動`） |
| ✓ | 17 | メニュー | Menu17 | `選択17()` | AprList No.17 のアプリを起動（`二重起動チェックと起動`） |
| ✓ | 18 | メニュー | Menu18 | `選択18()` | AprList No.18 のアプリを起動（`二重起動チェックと起動`） |
| ✓ | 19 | メニュー | Menu19 | `選択19()` | AprList No.19 のアプリを起動（`二重起動チェックと起動`） |
| ✓ | 20 | メニュー | Menu20 | `選択20()` | AprList No.20 のアプリを起動（`二重起動チェックと起動`） |
| ✓ | 21 | メニュー | Menu21 | `選択21()` | AprList No.21 のアプリを起動（`二重起動チェックと起動`） |
| ✓ | 22 | メニュー | Menu22 | `選択22()` | AprList No.22 のアプリを起動（`二重起動チェックと起動`） |
| ✓ | 23 | メニュー | Menu23 | `選択23()` | AprList No.23 のアプリを起動（`二重起動チェックと起動`） |
| ✓ | 24 | メニュー | Menu24 | `選択24()` | AprList No.24 のアプリを起動（`二重起動チェックと起動`） |
| ✓ | 25 | メニュー | Menu25 | `選択25()` | AprList No.25 のアプリを起動（`二重起動チェックと起動`） |
| ✓ | 26 | メニュー | Menu26 | `選択26()` | AprList No.26 のアプリを起動（`二重起動チェックと起動`） |
| ✓ | 27 | メニュー | Menu27 | `選択27()` | AprList No.27 のアプリを起動（`二重起動チェックと起動`） |
| ✓ | 28 | メニュー | Menu28 | `選択28()` | AprList No.28 のアプリを起動（`二重起動チェックと起動`） |
| ✓ | 29 | メニュー | Menu29 | `選択29()` | AprList No.29 のアプリを起動（`二重起動チェックと起動`） |
| ✓ | 30 | メニュー | Menu30 | `選択30()` | AprList No.30 のアプリを起動（`二重起動チェックと起動`） |

### 5.2 ショートカットキー

なし。

### 5.3 ユーザーフォーム上のボタン

| ✓ | No | フォーム | コントロール | キャプション | イベント | 呼び出すプロシージャ | 動作概要 |
| --- | --- | --- | --- | --- | --- | --- | --- |
| ✓ | 1 | **SikakuKakuninn.frm** | Kakuninn | 確認 | Click | `Kakuninn_Click()` → `SyainName()` | SHAINMST で資格照合。成功時 `mpSyaName` 設定してフォームを閉じる |
| ✓ | 2 | **SikakuKakuninn.frm** | CommandButton1 | キャンセル | Click | `CommandButton1_Click()` | フォームを閉じる（登録中止） |

---

### 5.4 CommandBar に動的追加されるボタン

なし。

## 6. VBA モジュール仕様

### 6.0 全プロシージャ一覧


> ✓ = ユーザー操作の起点（Click イベント等） / DB I/O を実行 / 他モジュールから呼び出される Public

| ✓ | No | モジュール | プロシージャ | スコープ | 種別 | 概要 |
| --- | --- | --- | --- | --- | --- | --- |
| ✓ | 1 | **ThisWorkbook.cls** | `Workbook_Open()` | Private | Event | ブック起動時に `初期設定()` を呼び出す |
| ✓ | 2 | **ThisWorkbook.cls** | `Workbook_BeforeClose()` | Private | Event | 終了時に Standard/Formatting ツールバーと数式バーを表示に戻す |
|  | 3 | **ThisWorkbook.cls** | `Workbook_BeforeSave()` | Private | Event | 保存抑止ロジックはコメントアウト済み（TEST） |
|  | 4 | **Sheet3**（「メニュー」） | — | — | — | 「メニュー」シート（コードなし） |
|  | 5 | **Sheet5**（「AprList」） | — | — | — | 「AprList」シート（コードなし） |
| ✓ | 6 | **メニュー選択実行.bas** | `選択1()` | Public | Sub | Menu1 クリック時。`二重起動チェックと起動(1)` を呼び出す |
| ✓ | 7 | **メニュー選択実行.bas** | `選択2()` | Public | Sub | Menu2 クリック時。`二重起動チェックと起動(2)` を呼び出す |
| ✓ | 8 | **メニュー選択実行.bas** | `選択3()` | Public | Sub | Menu3 クリック時。`二重起動チェックと起動(3)` を呼び出す |
| ✓ | 9 | **メニュー選択実行.bas** | `選択4()` | Public | Sub | Menu4 クリック時。`二重起動チェックと起動(4)` を呼び出す |
| ✓ | 10 | **メニュー選択実行.bas** | `選択5()` | Public | Sub | Menu5 クリック時。`二重起動チェックと起動(5)` を呼び出す |
| ✓ | 11 | **メニュー選択実行.bas** | `選択6()` | Public | Sub | Menu6 クリック時。`二重起動チェックと起動(6)` を呼び出す |
| ✓ | 12 | **メニュー選択実行.bas** | `選択7()` | Public | Sub | Menu7 クリック時。`二重起動チェックと起動(7)` を呼び出す |
| ✓ | 13 | **メニュー選択実行.bas** | `選択8()` | Public | Sub | Menu8 クリック時。`二重起動チェックと起動(8)` を呼び出す |
| ✓ | 14 | **メニュー選択実行.bas** | `選択9()` | Public | Sub | Menu9 クリック時。`二重起動チェックと起動(9)` を呼び出す |
| ✓ | 15 | **メニュー選択実行.bas** | `選択10()` | Public | Sub | Menu10 クリック時。`二重起動チェックと起動(10)` を呼び出す |
| ✓ | 16 | **メニュー選択実行.bas** | `選択11()` | Public | Sub | Menu11 クリック時。`二重起動チェックと起動(11)` を呼び出す |
| ✓ | 17 | **メニュー選択実行.bas** | `選択12()` | Public | Sub | Menu12 クリック時。`二重起動チェックと起動(12)` を呼び出す |
| ✓ | 18 | **メニュー選択実行.bas** | `選択13()` | Public | Sub | Menu13 クリック時。`二重起動チェックと起動(13)` を呼び出す |
| ✓ | 19 | **メニュー選択実行.bas** | `選択14()` | Public | Sub | Menu14 クリック時。`二重起動チェックと起動(14)` を呼び出す |
| ✓ | 20 | **メニュー選択実行.bas** | `選択15()` | Public | Sub | Menu15 クリック時。`二重起動チェックと起動(15)` を呼び出す |
| ✓ | 21 | **メニュー選択実行.bas** | `選択16()` | Public | Sub | Menu16 クリック時。`二重起動チェックと起動(16)` を呼び出す |
| ✓ | 22 | **メニュー選択実行.bas** | `選択17()` | Public | Sub | Menu17 クリック時。`二重起動チェックと起動(17)` を呼び出す |
| ✓ | 23 | **メニュー選択実行.bas** | `選択18()` | Public | Sub | Menu18 クリック時。`二重起動チェックと起動(18)` を呼び出す |
| ✓ | 24 | **メニュー選択実行.bas** | `選択19()` | Public | Sub | Menu19 クリック時。`二重起動チェックと起動(19)` を呼び出す |
| ✓ | 25 | **メニュー選択実行.bas** | `選択20()` | Public | Sub | Menu20 クリック時。`二重起動チェックと起動(20)` を呼び出す |
| ✓ | 26 | **メニュー選択実行.bas** | `選択21()` | Public | Sub | Menu21 クリック時。`二重起動チェックと起動(21)` を呼び出す |
| ✓ | 27 | **メニュー選択実行.bas** | `選択22()` | Public | Sub | Menu22 クリック時。`二重起動チェックと起動(22)` を呼び出す |
| ✓ | 28 | **メニュー選択実行.bas** | `選択23()` | Public | Sub | Menu23 クリック時。`二重起動チェックと起動(23)` を呼び出す |
| ✓ | 29 | **メニュー選択実行.bas** | `選択24()` | Public | Sub | Menu24 クリック時。`二重起動チェックと起動(24)` を呼び出す |
| ✓ | 30 | **メニュー選択実行.bas** | `選択25()` | Public | Sub | Menu25 クリック時。`二重起動チェックと起動(25)` を呼び出す |
| ✓ | 31 | **メニュー選択実行.bas** | `選択26()` | Public | Sub | Menu26 クリック時。`二重起動チェックと起動(26)` を呼び出す |
| ✓ | 32 | **メニュー選択実行.bas** | `選択27()` | Public | Sub | Menu27 クリック時。`二重起動チェックと起動(27)` を呼び出す |
| ✓ | 33 | **メニュー選択実行.bas** | `選択28()` | Public | Sub | Menu28 クリック時。`二重起動チェックと起動(28)` を呼び出す |
| ✓ | 34 | **メニュー選択実行.bas** | `選択29()` | Public | Sub | Menu29 クリック時。`二重起動チェックと起動(29)` を呼び出す |
| ✓ | 35 | **メニュー選択実行.bas** | `選択30()` | Public | Sub | Menu30 クリック時。`二重起動チェックと起動(30)` を呼び出す |
| ✓ | 36 | **メニュー選択実行.bas** | `二重起動チェックと起動()` | Public | Sub | 二重起動チェック後、`Workbooks.Open` でアプリを起動 |
|  | 37 | **Debug用.bas** | `AllCls()` | Public | Sub | デバッグ用。メニュー表示を解除し `PGList` をクリア |
|  | 38 | **Debug用.bas** | `DellMenu()` | Public | Sub | Menu1〜Menu30 のラベルと OnAction を削除 |
| ✓ | 39 | **共通変数と開始処理.bas** | `初期設定()` | Public | Sub | 起動時の全初期化（バージョン確認・取込・同期・メニュー構築・お知らせ） |
| ✓ | 40 | **共通変数と開始処理.bas** | `MenuGetPath()` | Public | Function | `ExAprReadPath.txt` からパス読込み、ExMenu.xlsm の更新日時を比較 |
| ✓ | 41 | **SQL_Execution.bas** | `Open_oraconDB()` | Public | Sub | ADO/ODBC で ricdb に接続 |
| ✓ | 42 | **SQL_Execution.bas** | `SQL_Exe()` | Public | Sub | SQL 文を Execute 実行 |
| ✓ | 43 | **SQL_Execution.bas** | `SQL_INSERT_UPDATE()` | Public | Sub | COUNT 判定後 INSERT または UPDATE |
|  | 44 | **SQL_Execution.bas** | `SQL_Delete()` | Public | Sub | DELETE 実行（本ブックからは未呼出し） |
| ✓ | 45 | **SQL_Execution.bas** | `Disp_Sheet()` | Public | Sub | SQL 結果を指定シートに書込み |
| ✓ | 46 | **SQL_Execution.bas** | `Set_Array()` | Public | Sub | SQL 結果を配列に格納 |
| ✓ | 47 | **ファイル登録_取込.bas** | `AddSoft()` | Public | Sub | 【ファイルの追加・消去】から呼出し。「AprList」シートを表示し C4 セルにフォーカス |
| ✓ | 48 | **ファイル登録_取込.bas** | `BackMenu()` | Public | Sub | 「メニュー」シートへ戻る |
| ✓ | 49 | **ファイル登録_取込.bas** | `メニュー登録Main()` | Public | Sub | 変更行の DB 登録・消去（資格確認後） |
| ✓ | 50 | **ファイル登録_取込.bas** | `メニュー取込()` | Public | Sub | DB(ExSeihinj) から AprList I〜N 列へ読込み |
| ✓ | 51 | **ファイル登録_取込.bas** | `メニュー登録()` | Public | Sub | 1 行分を ExSEIHINJ に INSERT/UPDATE |
| ✓ | 52 | **お知らせメッセージ.bas** | `メッセージ()` | Public | Sub | お知らせファイル表示処理の入口 |
| ✓ | 53 | **お知らせメッセージ.bas** | `お知らせ()` | Public | Function | サーバー/ローカルのお知らせ txt を比較・表示 |
| ✓ | 54 | **SikakuKakuninn.frm** | `CommandButton1_Click()` | Private | Event | 資格確認フォームをキャンセルして閉じる |
| ✓ | 55 | **SikakuKakuninn.frm** | `Kakuninn_Click()` | Private | Event | 社員番号/パスワードで DB 資格照合 |
|  | 56 | **SikakuKakuninn.frm** | `UserForm_Activate()` | Private | Event | 入力欄クリア、`mpErrDes` に「資格確認」を設定 |
|  | 57 | **SikakuKakuninn.frm** | `UserForm_QueryClose()` | Private | Event | フォームを Unload |
| ✓ | 58 | **SikakuKakuninn.frm** | `SyainName()` | Private | Function | SHAINMST を SELECT し社員名を返す |

### 6.1 **ThisWorkbook.cls**

| プロシージャ | 概要 |
| --- | --- |
| `Workbook_Open()` | `初期設定()` を Call |
| `Workbook_BeforeClose()` | CommandBars Standard/Formatting を表示、DisplayFormulaBar=True、`Saved=True` |
| `Workbook_BeforeSave()` | Debug フラグによる保存抑止はコメントアウト（TEST） |

### 6.2 **共通変数と開始処理.bas**

**Public 変数**

| 変数 | 用途 |
| --- | --- |
| `mpAprD(30, 2)` | メニューNo ごとの Path(0)/ファイル名(1)/表示文字(2) |
| `mpSyaName` | 資格確認後の社員名 |

| プロシージャ | 処理フロー概要 |
| --- | --- |
| `初期設定()` | 1. 画面最大化・AprList 保護 2. `MenuGetPath` でバージョン確認 3. `メニュー取込` 4. 各アプリの Path/存在/更新日チェックと FileCopy 5. メニュー UI 非表示化 6. Menu1〜30 設定 7. `メッセージ` |
| `MenuGetPath()` | `ExAprReadPath.txt` から元/先パス読込 → ExMenu.xlsm の FileDateTime 比較 |

### 6.3 **メニュー選択実行.bas**

| プロシージャ | 処理フロー概要 |
| --- | --- |
| `選択1()`〜`選択30()` | 各 Menu ボタンの OnAction スタブ |
| `二重起動チェックと起動()` | 1. PGList からファイル名再読込 2. 開済 Workbook 名と照合 3. 未起動なら `Workbooks.Open`（Err 1004 は無視） |

### 6.4 **SQL_Execution.bas**

接続文字列: `DSN=ricdb;UID=ric;PWD=t6101`

| プロシージャ | 概要 |
| --- | --- |
| `Open_oraconDB()` | ADODB.Connection を Open |
| `SQL_Exe()` | `oraconn.Execute(mySQL)` |
| `SQL_INSERT_UPDATE()` | COUNT 後 INSERT or UPDATE（トランザクション） |
| `SQL_Delete()` | DELETE 実行（本ブック未使用） |
| `Disp_Sheet()` | Recordset をシートに CopyFromRecordset |
| `Set_Array()` | Recordset を 2 次元配列に格納 |

### 6.5 **ファイル登録_取込.bas**

| プロシージャ | 概要 |
| --- | --- |
| `メニュー取込()` | ExSeihinj を SELECT → AprList I 列起点に Disp_Sheet |
| `メニュー登録Main()` | 確認 → 資格確認 → O 列=False の行を登録/消去 → 再起動確認 |
| `メニュー登録()` | 文字数チェック後 `SQL_INSERT_UPDATE` で ExSEIHINJ 更新。I〜N 列も同期 |
| `BackMenu()` | 「メニュー」シートを Select |
| `AddSoft()` | 【ファイルの追加・消去】から呼出し。「AprList」シートを Select し C4 セルにフォーカス |

### 6.6 **お知らせメッセージ.bas**

| プロシージャ | 概要 |
| --- | --- |
| `メッセージ()` | ExAprReadPath.txt 読込 → `お知らせ()` 呼出。失敗時 MsgBox で起動停止 |
| `お知らせ()` | 元ファイルが 10 日超古い場合スキップ。内容差分があれば MsgBox 表示。No 選択時にローカルコピー更新 |

### 6.7 **Debug用.bas**

| プロシージャ | 概要 |
| --- | --- |
| `AllCls()` | Debug フラグに応じ UI 要素を表示に戻し、PGList クリア、`DellMenu` |
| `DellMenu()` | Menu1〜30 の Text/OnAction をクリア |

### 6.8 **SikakuKakuninn.frm**

資格条件: `hshika='1'`（※推論: 有効社員フラグと推定。VBA・セルに定義なし） かつ `cshika='2' OR '3'`（顧客/製品登録資格者 ─ VBAメッセージ「資格要件：顧客、製品登録資格者」より）

---

## 7. ユーザーフォーム仕様

> ✓ = ユーザー入力を受け付ける、またはイベントで業務処理を起動するコントロール

### 7.1 **SikakuKakuninn.frm**（資格確認フォーム）

**目的**: メニュー登録操作（`メニュー登録Main`）前に、社員番号/パスワードで SHAINMST を照合し `mpSyaName` を設定する。
**モードレス表示**: Show（Modal 指定なし）

#### コントロール一覧

| ✓ | No | コントロール | 種別 | 用途 |
| --- | --- | --- | --- | --- |
|  | 1 | Syainn | Label | 社員番号ラベル |
|  | 2 | Label1 | Label | 補助ラベル |
| ✓ | 3 | SyainnNo | TextBox | 社員番号入力 |
| ✓ | 4 | PassWord | TextBox | パスワード入力 |
| ✓ | 5 | Kakuninn | CommandButton | 資格照合実行 |
| ✓ | 6 | CommandButton1 | CommandButton | フォームを閉じる |

#### イベント一覧

| ✓ | No | イベント | 動作概要 |
| --- | --- | --- | --- |
| ✓ | 1 | `Kakuninn_Click()` | 🗄️ `SyainName()` で SHAINMST SELECT。成功時 `mpSyaName` 設定して Unload |
| ✓ | 2 | `CommandButton1_Click()` | フォームを Unload（登録中止） |
|  | 3 | `UserForm_Activate()` | SyainnNo/PassWord をクリア |
|  | 4 | `UserForm_QueryClose()` | Unload |
| ✓ | 5 | `SyainName()` | 🗄️ `SELECT TRIM(shaname) FROM SHAINMST WHERE shano=... AND shask=... AND hshika='1' AND (cshika='2' OR cshika='3')` |

---

## 8. DB 接続・外部連携

### 8.1 ODBC 接続設定

| DSN 名 | UID | PWD | 用途 |
| --- | --- | --- | --- |
| `ricdb` | `ric` | `t6101` | 照射管理システムDB — ExSeihinj / ExSEIHINJ / SHAINMST へのアクセス |

> **DB サーバー IP**: 163.59.144.156（VBA コメント上の旧 UNC パス `\\163.59.144.156\ExRicSys\` 参照。現在の Path 判定は `C:\ラジエ工業\ExRicSys\`〈EXメニュー配置フォルダ〉）

### 8.2 テーブル一覧（参照/更新区分付き）

> ✓ = INSERT / UPDATE / DELETE の対象テーブル（参照のみのテーブルは ✓ なし）

| ✓ | No | テーブル名 | 区分 | 主な用途 | キー列 | 参照/更新列 |
| --- | --- | --- | --- | --- | --- | --- |
|  | 1 | `ExSeihinj` | 参照 | メニュー取込時のアプリ一覧 SELECT | `KAISYACD`（9001〜9030） | `FOLDER`, `FILENAME`, `KAIBIKOU`, `HIKITORI`, `TOUDATE`, `TOUNAME` |
| ✓ | 2 | `ExSEIHINJ` | **参照＋更新** | メニュー登録/消去 | `KAISYACD` | 更新: `FOLDER`, `FILENAME`, `KAIBIKOU`, `HIKITORI`, `TOUDATE`, `TOUNAME`（`メニュー登録()` / **SikakuKakuninn.frm** 経由） |
|  | 3 | `SHAINMST` | 参照 | 資格確認 | `shano` | `shaname`, `shask`, `hshika`, `cshika` |

> **「キー列」の定義**: JOIN 条件または UPDATE/DELETE の WHERE 句で使用される列を示す。

### 8.3 SQL 一覧

#### 8.3.1 アプリ一覧読み込み（`メニュー取込()` / **ファイル登録_取込.bas**）

```sql
SELECT TO_NUMBER(SUBSTR(KAISYACD,2)), FOLDER, FILENAME, KAIBIKOU,
       HIKITORI, TOUDATE, TOUNAME,
       FOLDER, FILENAME, KAIBIKOU, HIKITORI, TOUDATE, TOUNAME
FROM ExSeihinj
WHERE KAISYACD>'9000' ORDER BY KAISYACD
```

#### 8.3.2 アプリ登録（`メニュー登録()` / **ファイル登録_取込.bas**）

```sql
SELECT COUNT(*) FROM ExSEIHINJ WHERE KAISYACD='<9001〜9030>'
INSERT INTO ExSEIHINJ (KAISYACD, FOLDER, FILENAME, KAIBIKOU, HIKITORI, TOUDATE, TOUNAME)
  VALUES('<cd>', '<path>', '<file>', '<label>', '<type>', '<date>', '<name>')
UPDATE ExSEIHINJ SET FOLDER='<path>', FILENAME='<file>', ... WHERE KAISYACD='<cd>'
```

#### 8.3.3 資格確認（`SyainName()` / **SikakuKakuninn.frm**）

```sql
SELECT TRIM(shaname) FROM SHAINMST
WHERE shano='<社員番号>' AND shask='<パスワード>'
  AND hshika='1' AND (cshika='2' OR cshika='3')
```

### 8.4 外部ファイル連携

| ファイル | パス | ファイル名 | 処理 | 備考 |
| --- | --- | --- | --- | --- |
| パス設定 | `ThisWorkbook.Path\` | `ExAprReadPath.txt` | 読込 | サーバー元パス・コピー先パス（CSV 2 フィールド） |
| お知らせ | 元/先パス（ExAprReadPath.txt から取得） | `ExMenuからのお知らせ.txt` | 読込/書込 | 起動時差分表示。10 日超古い元ファイルは無視 |
| 登録アプリ | `PGList` C 列 Path + D 列ファイル名 | 各 .xlsm 等 | 読込/コピー | 起動時にサーバー→ローカルへ FileCopy |
| メニュー本体 | 元/先パス | `ExMenu.xlsm` | 参照 | MenuGetPath で更新日時比較 |

#### ExAprReadPath.txt

| 項目 | 内容 |
| --- | --- |
| 出力元 | `MenuGetPath()` / **共通変数と開始処理.bas**、`メッセージ()` / **お知らせメッセージ.bas** |
| パス生成 | `ThisWorkbook.Path & "\" & "ExAprReadPath.txt"` |
| 読込形式 | `Input #myFno, mpAprMotoPath, mpAprSakiPath`（2 フィールド CSV） |
| 記録内容 | フィールド1=サーバー元フォルダ、フィールド2=ローカルコピー先フォルダ |
| 用途 | ExMenu.xlsm およびお知らせ txt の元/先パス解決 |

#### ExMenuからのお知らせ.txt

| 項目 | 内容 |
| --- | --- |
| 出力元 | `お知らせ()` / **お知らせメッセージ.bas** |
| パス生成 | `myAprMotoPath & "ExMenuからのお知らせ.txt"` / `myAprSakiPath & "ExMenuからのお知らせ.txt"` |
| 読込内容 | 行単位テキスト（Input で配列化） |
| 更新ロジック | 元と先の行数/内容が異なる場合 MsgBox 表示。「次回もお知らせ?」で No → 先ファイルを元内容で上書き（Write） |
| 削除・作成 | ファイル未存在（Err 53）時は空ファイルを Output で自動作成 |

#### 登録アプリファイル（動的）

| 項目 | 内容 |
| --- | --- |
| 出力元 | `初期設定()` / **共通変数と開始処理.bas** |
| パス生成 | ソース: `mpAprD(i,0) & mpAprD(i,1)`、コピー先: `ThisWorkbook.Path & "\" & mpAprD(i,1)` |
| 処理 | 更新日時不一致時 `FileCopy` でサーバー→ローカル同期。Err 53 時は再試行（最大 50 回） |
| 起動 | `二重起動チェックと起動()` が `Workbooks.Open ThisWorkbook.Path & "\" & ファイル名` |

---

## 9. データフロー

各フローは「起点 → 処理 → 結果」の粒度で記述する。

### 9.1 起動フロー

| No | 起点 | 処理 | 結果 |
| --- | --- | --- | --- |
| 1 | ユーザーが ExMenu.xlsm を開く | 📄 `Workbook_Open()` / **ThisWorkbook.cls** | `初期設定()` を呼出し |
| 2 | 上記 | 📄 `MenuGetPath()` | 📄 `ExAprReadPath.txt` 読込 → ExMenu.xlsm 更新日時比較 |
| 3 | バージョン不一致 | 🖥️ MsgBox 確認 | No → ブック終了 / Yes → 継続 |
| 4 | 上記 | 🗄️ `メニュー取込()` | ExSeihinj SELECT → 📊 AprList I〜N 列書込み |
| 5 | 上記 | 📄 `初期設定()` ループ（1〜30） | Path/存在/PC 名チェック → 📄 FileCopy（必要時）→ `mpAprD` 構築 |
| 6 | 上記 | 📊 メニュー UI 設定 | ActiveWindow 非表示化 → Menu1〜30 に表示名/OnAction 設定 |
| 7 | 上記 | 📄 `メッセージ()` → `お知らせ()` | 📄 お知らせ txt 差分確認 → 🖥️ 必要時 MsgBox |

#### ツリー図（補助）

```
(ExMenu.xlsm を開く)
└─ 📄 Workbook_Open                 [ThisWorkbook.cls]
   └─ 📄 初期設定                    [共通変数と開始処理.bas]
      ├─ 📄 MenuGetPath             [ExAprReadPath.txt 読込]
      ├─ 🗄️ メニュー取込             [ExSeihinj → AprList]
      ├─ 📄 FileCopy ループ          [各アプリ同期]
      ├─ 📊 Menu1〜30 設定           [表示名 + OnAction]
      └─ 📄 メッセージ → お知らせ     [ExMenuからのお知らせ.txt]
```

### 9.2 アプリ起動フロー

| No | 起点 | 処理 | 結果 |
| --- | --- | --- | --- |
| 1 | 🖥️（MenuN Click） | 📄 `選択N()` | `二重起動チェックと起動(N)` 呼出し |
| 2 | 上記 | 📊 PGList からファイル名再読込 | `mpAprD(N,1)` 取得 |
| 3 | 上記 | 📄 全 Workbooks 走査 | 同名ブックあり → 🖥️ MsgBox + Activate |
| 4 | 未起動 | 📄 `Workbooks.Open` | ローカル Path からアプリ起動 |

#### ツリー図（補助）

```
(MenuN Click)
└─ 📄 選択N()                       [メニュー選択実行.bas]
   └─ 📄 二重起動チェックと起動(N)
      ├─ 📊 PGList 再読込
      ├─ [既に開いている] → 🖥️ MsgBox + Activate
      └─ [未起動] → 📄 Workbooks.Open
```

### 9.3 ファイル管理画面遷移フロー

| No | 起点 | 処理 | 結果 |
| --- | --- | --- | --- |
| 1 | 🖥️（【ファイルの追加・消去】Click） | 📄 `AddSoft()` / **ファイル登録_取込.bas** | 📊 「AprList」シートを Select し C4 セルにフォーカス |
| 2 | 📊 管理者が AprList C〜H 列を編集 | 📊 O 列の数式が自動評価 | O 列=False の行が変更あり状態になる |
| 3 | 🖥️（メニュー画面 Click） | 📄 `BackMenu()` / **ファイル登録_取込.bas** | 📊 「メニュー」シートを Select して戻る |

#### ツリー図（補助）

```
(【ファイルの追加・消去】Click)
└─ 📄 AddSoft()                     [ファイル登録_取込.bas]
   └─ 📊 AprList シート表示 + C4 フォーカス
      ├─ (管理者が C〜H 列を編集)
      ├─ (ﾌｧｲﾙ登録 Click) → 9.4 メニュー登録フローへ
      └─ (メニュー画面 Click) → 📄 BackMenu() → 📊 「メニュー」へ戻る
```

### 9.4 メニュー登録フロー

| No | 起点 | 処理 | 結果 |
| --- | --- | --- | --- |
| 1 | 📊 管理者が AprList C〜H 列を編集 | 📊 O 列が False（変更あり）になる | 登録待ち状態 |
| 2 | 🖥️（ファイル登録 Click） | 📄 `メニュー登録Main()` | 確認 MsgBox |
| 3 | 社員名未設定 | 🖥️ **SikakuKakuninn.frm** 表示 | 🗄️ SHAINMST 照合 → `mpSyaName` 設定 |
| 4 | 上記 | 📄 変更行ループ | Path/ファイル空 → 消去確認 → 🗄️ `メニュー登録()` |
| 5 | 上記 | 🗄️ `SQL_INSERT_UPDATE` | ExSEIHINJ INSERT/UPDATE + 📊 I〜N 列同期 |
| 6 | 更新あり | 🖥️ 再起動確認 Yes | 📄 `初期設定()` 再実行 |

#### ツリー図（補助）

```
(AprList 編集 → ﾌｧｲﾙ登録 Click)
└─ 📄 メニュー登録Main()            [ファイル登録_取込.bas]
   ├─ 🖥️ SikakuKakuninn.Show       [SHAINMST 照合]
   ├─ 🗄️ メニュー登録 (各行)       [ExSEIHINJ UPDATE]
   └─ 📄 初期設定 (再起動時)         [メニュー再構築]
```

---

## 10. セキュリティ注意事項


| No | カテゴリ | 内容 | リスク |
| --- | --- | --- | --- |
| 1 | 認証情報ハードコード | DSN=`ricdb`, UID=`ric`, PWD=`t6101` が VBA ソースに平文記載 | 中：VBAエディタで閲覧可能 |
| 2 | 社員パスワード | `SHAINMST.shask` を SQL 直結照合。ハッシュ化等の対策は不明 | 中：認証情報がソースとDB照会で露出 |
| 3 | ファイル自動コピー | 起動時にサーバーから最大 30 アプリをローカルへ自動コピー | 中：ローカルファイルの作成・上書きが可能 |
| 4 | ファイル操作 | Open / Write / Kill / MkDir を使用 | 中：ローカルファイルの作成・削除が可能 |
| 5 | 社内アドレス露出 | 旧サーバー IP がコメントアウト済み UNC パス内に残る。現在は `C:\ラジエ工業\ExRicSys\`〈EXメニュー配置フォルダ〉 | 低：社内アドレスの露出 |

## スコープ外（本仕様書に含まないもの）

- セル書式（色・罫線・フォント）
- 条件付き書式、グラフ・画像、印刷設定

必要な場合は Excel 画面のスクリーンショットで補完してください。
