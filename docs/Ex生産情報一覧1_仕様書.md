# Ex生産情報一覧1 仕様書

> **ファイル種別**: .xlsm（マクロ付き）
> **用途**: 照射装置（1号機・2号機・3号機・EB）の未照射在庫品の生産計画・出荷予定・照射スケジュールを管理し、DBからデータを取得して一覧表示・計画更新・出庫記録を行う総合管理ツール
> **VBA プロジェクト**: モジュール 56 本（.bas 42 / .cls 7 / .frm 7）
> **外部連携**: DSN=ricdb（Oracle DB）、DB接続先 IP: 163.59.144.156
> **解析日**: 2026-06-08（excel-to-md スキルによる自動解析・フィードバック反映改訂版）

---

## 凡例（本仕様書の表記ルール）

本仕様書では、保守時の判別を容易にするため、以下の表記ルールを使用します。

| 種別 | 表記 | 例 |
|---|---|---|
| モジュール（.bas / .cls） | **太字** | **スタート処理.bas** |
| ユーザーフォーム（.frm） | **太字** | **SoutiSenntaku.frm** |
| プロシージャ / イベント | `コード体()` | `生産情報開始処理()` ※Markdown 制約上、斜体ではなくコード体で統一 |
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
6. [VBA モジュール仕様](#6-vba-モジュール仕様)
7. [ユーザーフォーム仕様](#7-ユーザーフォーム仕様)
8. [DB 接続・外部連携](#8-db-接続外部連携)
9. [データフロー](#9-データフロー)
10. [セキュリティ注意事項](#10-セキュリティ注意事項)

---

## 1. ファイル構成

### 1.1 シート一覧（全 6 件）

> ✓ = ユーザーが直接操作する、または VBA が動的に表示/非表示を切り替えるシート

| ✓ | シート名 | 最大行 | 最大列 | 保存時 Visible | VBA による動的切替 |
|---|---|---|---|---|---|
| ✓ | 未処理品一覧 | 1017 | 73 (BU) | visible | — |
| ✓ | 受付番号別完了 | 1000 | 69 (BQ) | visible | 遷移時に表示/非表示 |
| ✓ | Ric3TR | 511 | 64 (BL) | visible | 3号機モード時のみ表示 |
|   | 設定 | 1015 | 44 (AR) | visible | DebugFlg=TRUE 時のみ手動表示 |
|   | TrakD | 29 | 9 (I) | visible | VBA から参照のみ |
|   | ２号機重量係数 | 52 | 5 (E) | visible | VBA から参照のみ |

### 1.2 ユーザーフォーム一覧（全 7 件）

> ✓ = ユーザー入力を受け付ける、または業務フローの起点となるフォーム

| ✓ | フォーム名 | コントロール数 | 主用途 |
|---|---|---|---|
| ✓ | SoutiSenntaku.frm | 30+ | 装置選択・表示方法・ソート設定（起動直後に表示） |
| ✓ | SyoriSettei.frm | 14+ | 処理条件設定（2/3号機パラメータ・業務資格登録） |
| ✓ | SikakuKakuninn.frm | 6 | 更新前の資格確認（社員ID/PW入力） |
| ✓ | Ric3Stg.frm | 100+ | 3号機ステージ表示（照射室内状況・出庫操作） |
|   | 製品情報.frm | 13+ | 製品特殊条件の詳細表示 |
|   | TokuSyuDsp.frm | 1 | 特殊条件ダイアログ（OKボタンのみ） |
|   | UserForm処理中.frm | 1 | 処理中メッセージ表示（操作なし） |

### 1.3 VBA モジュール一覧（全 56 件）

> ✓ = ユーザー操作の起点 / DB I/O を含む / 他モジュールから呼び出される / コード行数上位 25%

| ✓ | モジュール | 種別 | プロシージャ数 | 主な役割 |
|---|---|---|---|---|
| ✓ | **更新在庫.bas** | .bas | 5 | 在庫データの DB 更新（`ExKeikakuX`） |
| ✓ | **ThisWorkbook.cls** | .cls | 3 | 起動 / 終了イベント |
| ✓ | **R3出庫記録処理.bas** | .bas | 1 | 3号機出庫順の DB 記録 |
| ✓ | **R3計画作成.bas** | .bas | 3 | 3号機未照射品計画・照射中表示 |
|   | **Sheet2.cls** | .cls | 0 | 「受付番号別完了」シートモジュール |
| ✓ | **R2印刷範囲.bas** | .bas | 1 | 2号機印刷範囲設定 |
| ✓ | **装置稼働状況.bas** | .bas | 4 | 装置稼働状況の DB 取得・表示 |
| ✓ | **コマンド実行.bas** | .bas | 2 | 初期化（装置選択後の主制御）・ソート実行 |
| ✓ | **SyoriSettei.frm** | .frm | 14 | 処理条件設定フォーム |
| ✓ | **コマンドサブルーチン.bas** | .bas | 4 | CommandBar 登録・データ更新・行削除 |
|   | **R3出庫印刷.bas** | .bas | 2 | 3号機出庫印刷範囲設定 |
|   | **ユーティリティーイベント有効.bas** | .bas | 1 | R3計画クリア |
|   | **共通変数.bas** | .bas | 0 | グローバル変数宣言 |
|   | **TokuSyuDsp.frm** | .frm | 1 | 特殊条件ダイアログ |
| ✓ | **画面クリア_遷移1.bas** | .bas | 5 | シート遷移・画面初期化 |
| ✓ | **サブルーチン.bas** | .bas | 12 | 共通サブルーチン（休日読込・パラメータ読込等） |
| ✓ | **画面色と表示列設定.bas** | .bas | 2 | 行背景色・列表示/非表示設定 |
|   | **照射中パス集計.bas** | .bas | 1 | Ric3TR 照射中残パス集計 |
|   | **生産情報終了.bas** | .bas | 1 | 終了処理 |
| ✓ | **スタート処理.bas** | .bas | 6 | 起動シーケンス（DB読込・装置選択） |
| ✓ | **データ取得.bas** | .bas | 3 | DB からのデータ読込・表示 |
| ✓ | **在庫詳細.bas** | .bas | 4 | 在庫詳細表示・照射中未投入処理量 |
| ✓ | **製品情報.frm** | .frm | 2 | 製品情報表示フォーム |
|   | **ErrCheck.bas** | .bas | 1 | エラーチェックメッセージ |
| ✓ | **DB不要データ削除.bas** | .bas | 2 | DB 孤児レコード削除 |
| ✓ | **SikakuKakuninn.frm** | .frm | 3 | 資格確認フォーム |
| ✓ | **資格確認.bas** | .bas | 1 | 社員マスタ照合 |
| ✓ | **更新予約.bas** | .bas | 3 | 予約データの DB 更新（`ExYoyakuX`） |
| ✓ | **R2未照射急ぎ順.bas** | .bas | 1 | 2号機急ぎ順計算 |
|   | **R3最終完了日.bas** | .bas | 2 | 3号機最終完了日時計算 |
|   | **オートフィルタ.bas** | .bas | 1 | オートフィルター設定 |
| ✓ | **R2指定日チェック.bas** | .bas | 2 | 計画係数・日付変換 |
|   | **Sheet4.cls** | .cls | 0 | 「設定」シートモジュール |
| ✓ | **R3タイマーチェック.bas** | .bas | 1 | 3号機タイマーチェック |
|   | **Sheet5.cls** | .cls | 0 | 「TrakD」シートモジュール |
| ✓ | **R2計画作成.bas** | .bas | 4 | 2号機計画作成・計画済表示 |
|   | **R3フォーム_セル高さ調整.bas** | .bas | 2 | Ric3Stg フォーム初期化・行高さ調整 |
| ✓ | **R2予想計算メイン.bas** | .bas | 16 | 2号機予想計算（トラッキング・完了表示） |
|   | **シェープ表示非表示.bas** | .bas | 10 | シート上のシェープ（矢印・ボタン）表示制御 |
| ✓ | **R2計画指定.bas** | .bas | 1 | 2号機計画順指定 |
| ✓ | **R3Stg1挿入取消.bas** | .bas | 2 | 3号機出庫設定・指定取消 |
| ✓ | **Ric3Stg.frm** | .frm | 2 | 3号機ステージフォーム |
| ✓ | **R3予想計算.bas** | .bas | 8 | 3号機予想計算 |
| ✓ | **SQL_Execution.bas** | .bas | 7 | DB 接続・SQL 実行共通モジュール |
| ✓ | **R3出庫設定.bas** | .bas | 1 | 未指定全数出庫 |
| ✓ | **R3照射室.bas** | .bas | 6 | 3号機照射室ステージ表示 |
| ✓ | **R2速度チェック.bas** | .bas | 2 | 2号機速度・重量係数チェック |
| ✓ | **R3停止期間記録読出.bas** | .bas | 3 | 3号機停止期間の DB 記録/読込 |
| ✓ | **R3Data読込.bas** | .bas | 4 | 3号機予想データ読込・出荷日式コピー |
| ✓ | **ファンクション.bas** | .bas | 8 | 汎用関数（出荷日変換・照射日判定等） |
|   | **UserForm処理中.frm** | .frm | 0 | 処理中表示 |
| ✓ | **Sheet1.cls** | .cls | 2 | 「未処理品一覧」シートイベント |
| ✓ | **Sheet3.cls** | .cls | 5 | 「Ric3TR」シートイベント |
| ✓ | **SoutiSenntaku.frm** | .frm | 13 | 装置選択フォーム |
|   | **Sheet6.cls** | .cls | 0 | 「２号機重量係数」シートモジュール |
| ✓ | **出荷方法更新.bas** | .bas | 2 | 製品情報ファイル更新 |

---

## 2. シート詳細

### 2.0 シート可視性一覧

保存時はすべて visible だが、VBA が動的に表示/非表示を制御する。

| シート | VBA による非表示化 | 表示するタイミング | 非表示にするタイミング | 制御プロシージャ |
|---|---|---|---|---|
| 未処理品一覧 | — | 常時表示 | — | — |
| 受付番号別完了 | あり | `受付番号別完了遷移()` 実行時 | `未処理品一覧に戻る()` 実行時 | **画面クリア_遷移1.bas** |
| Ric3TR | あり | 装置選択で「3号機完了予想」選択時 | `Ric3TRに戻る()` / 別シート遷移時 | **コマンド実行.bas** `初期化()` / **画面クリア_遷移1.bas** |
| 設定 | 通常非表示 | DebugFlg=TRUE 時のみ手動表示 | — | — |
| TrakD | 通常非表示 | VBA から参照のみ（表示不要） | — | — |
| ２号機重量係数 | 通常非表示 | VBA から参照のみ（表示不要） | — | — |

> 以下の各シートのレイアウト構造表における ✓ = VBA から `Range()` で代入または参照され、業務ロジックに直結するセル

### 2.1 未処理品一覧

**目的**: メイン操作画面。DBから取得した未照射在庫品の生産計画データを一覧表示し、計画確認・出荷方法・備考の入力・更新を行う。

#### 非表示行・列

| 種類 | 対象 | 備考 |
|---|---|---|
| 非表示行 | 5〜14（10行） | ヘッダー・操作領域（VBA が制御） |
| 非表示列 | C, D, X, Y, Z, AA, AB, AC, AD, AJ, AK, AL, AY, AZ, BK, BL（16列） | 内部計算・比較用データ列 |

#### レイアウト構造

行 1〜4: 操作パネル（日付・処理メッセージ・表示方法等）
行 5〜14: 非表示ヘッダー領域
行 15: 列見出し行
行 16〜1015: データ本体（VBA が `Keikaku` 名前付き範囲に書込み）

| ✓ | セル | 名前付き範囲 | 種別 | 実態（値/数式/VBA代入） | 業務的意味 |
|---|---|---|---|---|---|
| ✓ | `$B$16:$BU$1015` | `Keikaku` | VBA代入 | `ReadDataFromDB()` で配列を一括書込み | 生産計画データ全体 |
| ✓ | `$D$16:$BU$1015` | `KeikakuS` | VBA代入 | 上記の D列以降部分 | 数式参照用（VLOOKUP の検索範囲） |
| ✓ | `$D$16:$W$1015` | `MotoData` | VBA代入 | DB から取得した元データ | 更新前後比較に使用 |
| ✓ | `$H$2` | `NowToday` | VBA代入 | `Range("NowToday") = Int(Now())` | 当日日付（基準日） |
| ✓ | `$AN$3` | `HyoujiHouhou` | VBA代入 | "在庫全て表示" / "全て表示" / "未照射品のみ" | 現在の表示方法 |
| ✓ | `$R$4` | `SyoriMSG` | VBA代入 | "製品情報" / "２号機完了予想" / "３号機完了予想" | 現在の処理モード表示 |
| ✓ | `$AD$2` | `RowNo` | VBA代入 | 整数（例: 158） | 最終データ表示行番号 |
| ✓ | `$AF$2` | `ZaikoKensu` | VBA代入 | 整数 | 在庫件数 |
| ✓ | `$AH$2` | `KeiMaxNo` | 数式 | `=MAX(X16:X1515)` | 計画最大 No |
|   | `$W$3` | `Amari` | VBA代入 | `Range("Amari") = .Cells(myRow,13) Mod .Cells(myRow,11)` | 未計画数 ÷ BOX入数 の余り（端数） |
|   | `$A$1` | `DebugFlg` | 手動設定 | 空欄 or "1" | デバッグモード切替フラグ |
|   | `$T$2` | `Hanuke` | VBA代入 | 文字列 | 歯抜け判定欄 |
|   | `$X$3` | `Ric2HpSuu` | VBA代入 | 整数 | 2号機 HP × 台数合計 |
|   | `$X$4` | `Ric2HSuu` | VBA代入 | 整数 | 2号機 台数合計 |
|   | `$X$16:$Y$1015` | `Ric2Keikaku` | VBA代入 | 整数配列 | 2号機計画データ |
|   | `$T$15` | `SeihinToku` | VBA代入 | "出荷時特殊条件\n製品管理情報" 等 | 特殊条件表示列の見出し |
|   | `$AO$3:$AX$3` | `SyukkaHou` | 設定値 | 出荷方法の選択肢リスト | 出荷方法一覧 |
|   | `$G$2` | — | 数式 | `=NowToday-1` | 前日日付 |
|   | `$I$2`〜`$O$2` | — | 数式 | `=NowToday+1`〜`=NowToday+7` | 翌日〜7日後の日付 |

### 2.2 受付番号別完了

**目的**: 受付番号単位の完了状況（出荷済数・残数・予想完了日など）を表示する。

#### 非表示行・列

| 種類 | 対象 | 備考 |
|---|---|---|
| 非表示行 | 3, 4（2行） | 予想計算式・ヘッダー |
| 非表示列 | B, D, E, F, G, H, K, L, M, O, V（11列） | 内部データ列 |

#### レイアウト構造

行 1〜4: 操作パネル
行 5: 列見出し（受付番号, 会社名, 納期, 出荷日, 未照射, 未投入数, 照射中, 状態, No）
行 6〜1000: データ本体（VBA が `ZaData` 名前付き範囲に書込み）

| ✓ | セル | 名前付き範囲 | 種別 | 実態 | 業務的意味 |
|---|---|---|---|---|---|
| ✓ | `$D$6:$AE$1000` | `ZaData` | VBA代入 | 受付番号別完了予想データ | 受付番号別データ全体 |
|   | `$U$2` | `Hanuke2` | 数式 | `=IF(AND(RRic2=TRUE,Hanuke<>""),Hanuke,"")` | 歯抜け判定表示 |
|   | `$AE$2` | `PrintEND` / `PrinｔE` | VBA代入 | 整数（例: 50） | 印刷終了行 |
|   | `$P$2` | `YosouMSG` | VBA代入 | "全ての受付番号別完了予想" | 予想メッセージ |
|   | `$Q$3:$T$3` | `YosouSiki` | 数式 | VLOOKUP 参照式 | 予想計算式 |

### 2.3 Ric3TR

**目的**: 3号機トラッキング計画シート。各受付番号がどのステージに配置されるかを管理し、出庫順を設定する。

#### 非表示行・列

| 種類 | 対象 | 備考 |
|---|---|---|
| 非表示行 | 5〜9, 11〜25（計 23 行） | 計算パラメータ行・予備行 |
| 非表示列 | B, D, F, I, K, N, S, Y, AA, AC, AL, AR, AX, BH（14列） | 内部計算列 |

#### レイアウト構造

行 1: 操作説明（"出庫設定の仕方：出庫順の列のクリックで指定／取消ができます。"）
行 3〜4: 停止/再開日時パラメータ
行 8: 計算パラメータ行（Ric3KeiSiki1〜3 等）
行 10: 列見出し
行 11〜510: データ本体（VBA が `Ric3Work` / `Ric3No` に書込み）

| ✓ | セル | 名前付き範囲 | 種別 | 実態 | 業務的意味 |
|---|---|---|---|---|---|
| ✓ | `$F$11:$BK$510` | `Ric3Work` | VBA代入 | 3号機作業データ全体 | ステージ配置・出庫順 |
| ✓ | `$E$11:$E$510` | `Ric3No` | VBA代入 | 連番 | 3号機 No 一覧 |
| ✓ | `$E$8` | `Ric3NoMax` | 数式 | `=COUNTIF(Ric3No,">0")` | 3号機データ件数 |
|   | `$I$8` | `Ric3SyukkoMax` | 数式 | `=IF(MAX(I11:I510)<0,0,MAX(I11:I510))` | 出庫最大値 |
|   | `$BK$4` | `Hanuke3S` | VBA代入 | 空欄 or 値 | 3号機歯抜け判定 |
|   | `$AW$3` | `Ric3Teisiji` | VBA代入 | 日時 | 3号機停止日時 |
|   | `$AW$4` | `Ric3Saikai` | VBA代入 | 日時 | 3号機再開日時 |
|   | `$BG$3` | `RIC3StopTime` | VBA代入 | 日時 | 3号機停止時間 |
|   | `$BG$4` | `Ric3StarTime` | VBA代入 | 日時 | 3号機開始時間 |
|   | `$AW$1` | `WarnMSG` | VBA代入 | 文字列 | 3号機警告メッセージ |

### 2.4 設定

**目的**: DB接続設定・パラメータ・マスタデータ（出荷日テーブル、計画HP、引取業者コード等）を格納する。

#### 非表示行・列

| 種類 | 対象 | 備考 |
|---|---|---|
| 非表示行 | なし | — |
| 非表示列 | D, H, J, K（4列） | 内部データ列 |

#### レイアウト構造

行 1〜3: 見出し（休祭日・出荷日カレンダー / 処理設定表 / 2号機/3号機パラメータ等）
行 4〜123: 出荷日テーブル（`SyukabiTB`）
行 4〜133: 発売日一覧（`Hatabi`）
行 4〜505: 引取業者コード（`HikitoriCD`）/ 計画HP（`KeikakuHP`）

| ✓ | セル | 名前付き範囲 | 種別 | 実態 | 業務的意味 |
|---|---|---|---|---|---|
| ✓ | `$A$4:$F$123` | `SyukabiTB` | 設定値/数式 | 出荷日カレンダーテーブル | 出荷日判定に使用 |
| ✓ | `$AD$4:$AE$505` | `HikitoriCD` | DB読込 | 引取業者コードテーブル | 出荷方法判定に使用 |
| ✓ | `$Y$6:$AB$505` | `KeikakuHP` | DB読込 | 計画HPテーブル | 2号機計画HP参照 |
| ✓ | `$O$33` | `SyainID` | VBA代入 | 社員番号 | 更新者ID（1,2号機用） |
| ✓ | `$O$34` | `SyainPW` | VBA代入 | パスワード | 更新者PW（1,2号機用） |
| ✓ | `$O$35` | `AccessTime` | VBA代入 | 日時 | 最終アクセス時刻 |
|   | `$O$4:$P$6` | `ParaM` | DB読込 | 2/3号機パラメータ | T/V/HP/PP/M 設定値 |
|   | `$O$31` | `SyoriMorde` | VBA代入 | TRUE/FALSE | 処理モード |
|   | `$O$13` | `Sort1` | VBA代入 | ソート項目名 | 第1ソート設定 |
|   | `$O$43` | `ReqNo` | VBA代入 | 整数 | 装置選択のリクエストNo |
|   | `$I$4:$I$133` | `Hatabi` | 設定値 | 日付リスト | 発売日一覧 |

### 2.5 TrakD

**目的**: 3号機トラック構成情報（各ステージ数・時間設定・半減期計算パラメータ）を格納する。

#### 非表示行・列

なし。

#### レイアウト構造

行 1: タイトル（"計算に使用する係数"）
行 2〜5: 基本パラメータ（年間稼動時間・測定日・1パス時間・ハンガー台数）
行 8〜15: ステージ別時間設定
行 16: 線源強度計算（半減期式 `=ROUND(G16*0.5^((TODAY()-F16)/1921),0)`）
行 20〜28: 合成時間（搬入R・サフR・未処理・新RRT・Stg6-4・Stg7-4・抽出）

| ✓ | セル | 名前付き範囲 | 種別 | 実態 | 業務的意味 |
|---|---|---|---|---|---|
| ✓ | `$D$5` | `NowRound` / `Hang4_5N` | 設定値 | 整数（例: 49） | 4Stg+5Stg のハンガー台数 |
| ✓ | `$D$16` | — | 数式 | `=ROUND(G16*0.5^((TODAY()-F16)/1921),0)` | 現在の線源強度（半減期計算） |
|   | `$E$20` | `KanRtime` | 数式 | `=D12+D14` | 搬入Rタイム |
|   | `$E$22` | `MisyTime` | 数式 | `=D8+D9` | 未処理タイム |
|   | `$E$23` | `NewRRT` | 数式 | `=E4/D16*D4` | 新RRT |

### 2.6 ２号機重量係数

**目的**: 2号機の出荷計画計算に使用する重量係数テーブル。

#### 非表示行・列

なし。

#### レイアウト構造

| 行 | 内容 |
|---|---|
| 1 | タイトル（"２号機照射室内総重量と速度係数（データのみ）"） |
| 2 | ヘッダー行: A=No, B=総重量, C=係数, E=VBA Case文生成式 |
| 3〜52 | 明細行（50行）: 重量閾値と速度係数の対応表 |

明細行の E 列は数式 `="Case Is <=" &B3&": R2speedk ="&C3` で VBA の Select Case 文を自動生成しており、`R2SpeedK()` 関数内で参照される。

---

## 3. 名前付き範囲一覧

全 **95** 件。

> ✓ = VBA から `Range()` で代入または参照され、業務ロジックに直結する名前付き範囲

| ✓ | 名前 | 参照先 | 業務的意味 |
|---|---|---|---|
| ✓ | `AccessTime` | 設定!$O$35 | 最終アクセス時刻（1,2号機） |
| ✓ | `AccessTimeR3` | 設定!$P$35 | 最終アクセス時刻（3号機） |
|   | `Amari` | 未処理品一覧!$W$3 | 未計画数÷BOX入数の余り |
|   | `ComList` | 設定!$N$23:$N$30 | コマンドリスト |
| ✓ | `DebugFlg` | 未処理品一覧!$A$1 | デバッグフラグ |
|   | `Hang4_5N` | TrakD!$D$5 | 3号機 4+5Stg ハンガー台数 |
|   | `Hanuke` | 未処理品一覧!$T$2 | 歯抜け判定欄 |
|   | `Hanuke2` | 受付番号別完了!$U$2 | 受付番号別完了判定欄 |
|   | `Hanuke3S` | Ric3TR!$BK$4 | 3号機歯抜け判定 |
|   | `Hatabi` | 設定!$I$4:$I$133 | 発売日一覧 |
| ✓ | `HikitoriCD` | 設定!$AD$4:$AE$505 | 引取業者コードテーブル |
| ✓ | `HyoujiHouhou` | 未処理品一覧!$AN$3 | 表示方法 |
|   | `KanRtime` | TrakD!$E$20 | 搬入Rタイム |
|   | `KeiMaxNo` | 未処理品一覧!$AH$2 | 計画最大No |
| ✓ | `KeikaTime` | 設定!$O$37 | 更新経過時間閾値（分） |
| ✓ | `KeikaTimeR3` | 設定!$P$37 | 更新経過時間閾値（3号機） |
| ✓ | `Keikaku` | 未処理品一覧!$B$16:$BU$1015 | 生産計画データ全体 |
| ✓ | `KeikakuHP` | 設定!$Y$6:$AB$505 | 計画HPテーブル |
| ✓ | `KeikakuS` | 未処理品一覧!$D$16:$BU$1015 | 生産計画データ（VLOOKUP用） |
|   | `MisyTime` | TrakD!$E$22 | 未処理タイム |
| ✓ | `MotoData` | 未処理品一覧!$D$16:$W$1015 | 元データ格納領域 |
|   | `NewRRT` | TrakD!$E$23 | 新RRT |
|   | `NowRound` | TrakD!$D$5 | 3号機現在ラウンド |
| ✓ | `NowToday` | 未処理品一覧!$H$2 | 現在日付 |
|   | `ParaM` | 設定!$O$4:$P$6 | 2/3号機パラメータ |
|   | `Point` | Ric3TR!$K$8 | ポイント値 |
|   | `PrintEND` | 受付番号別完了!$AE$2 | 印刷終了行 |
|   | `PrinｔE` | 受付番号別完了!$AE$2 | 印刷終了行（別名） |
|   | `RIC3StopTime` | Ric3TR!$BG$3 | 3号機停止時間 |
|   | `RIc2T` | 設定!$O$4 | 2号機T設定 |
| ✓ | `RRic1` | 設定!$O$18 | 1号機表示フラグ |
| ✓ | `RRic2` | 設定!$O$19 | 2号機表示フラグ |
| ✓ | `RRic3` | 設定!$O$20 | 3号機表示フラグ |
| ✓ | `ReqNo` | 設定!$O$43 | 装置選択リクエストNo |
|   | `Ric2HP` | 設定!$O$6 | 2号機HP設定値 |
|   | `Ric2HSuu` | 未処理品一覧!$X$4 | 2号機台数合計 |
|   | `Ric2HpSuu` | 未処理品一覧!$X$3 | 2号機HP×台数合計 |
|   | `Ric2Keikaku` | 未処理品一覧!$X$16:$Y$1015 | 2号機計画データ |
|   | `Ric2NowHP` | 設定!$O$8 | 2号機現在HP |
|   | `Ric2NowSokudo` | 設定!$O$7 | 2号機現在速度 |
|   | `Ric2V` | 設定!$O$5 | 2号機V設定 |
|   | `Ric3KeiSiki1` | Ric3TR!$F$8 | 3号機計算式1 |
|   | `Ric3KeiSiki2` | Ric3TR!$R$8:$BG$8 | 3号機計算式2 |
|   | `Ric3KeiSiki3` | Ric3TR!$L$8 | 3号機計算式3 |
|   | `Ric3M` | 設定!$P$5 | 3号機M設定 |
|   | `Ric3MisyoriN` | Ric3TR!$BJ$7 | 3号機未処理No |
| ✓ | `Ric3No` | Ric3TR!$E$11:$E$510 | 3号機No一覧 |
|   | `Ric3NoMax` | Ric3TR!$E$8 | 3号機No最大値 |
|   | `Ric3NowPP` | 設定!$P$8 | 3号機現在PP |
|   | `Ric3NowTime` | 設定!$P$7 | 3号機現在時間 |
|   | `Ric3PP` | 設定!$P$6 | 3号機PP設定 |
|   | `Ric3RowS` | Ric3TR!$I$7 | 3号機開始行 |
|   | `Ric3Saikai` | Ric3TR!$AW$4 | 3号機再開フラグ |
|   | `Ric3SetteiMax` | Ric3TR!$G$8 | 3号機設定最大値 |
|   | `Ric3StarTime` | Ric3TR!$BG$4 | 3号機開始時間 |
|   | `Ric3SyukkoMax` | Ric3TR!$I$8 | 3号機出庫最大値 |
|   | `Ric3T` | 設定!$P$4 | 3号機T設定 |
|   | `Ric3Teisiji` | Ric3TR!$AW$3 | 3号機停止指示 |
|   | `Ric3Track` | 設定!$AH$9:$AJ$12 | 3号機トラック設定 |
|   | `Ric3TyuuPP` | 設定!$X$3 | 3号機注PP |
| ✓ | `Ric3Work` | Ric3TR!$F$11:$BK$510 | 3号機作業領域全体 |
|   | `RicEB` | 設定!$O$21 | EB表示フラグ |
| ✓ | `RowNo` | 未処理品一覧!$AD$2 | 最終データ表示行番号 |
|   | `SeihinToku` | 未処理品一覧!$T$15 | 製品特殊条件見出し |
|   | `SenkNo` | 設定!$O$45 | 線源No |
|   | `SetteiV` | TrakD!$F$16:$H$16 | 設定V |
|   | `Sort1` | 設定!$O$13 | ソート1設定 |
|   | `Sort1M` | 設定!$P$13 | ソート1モード |
|   | `Sort1Modo` | 設定!$P$40 | ソート1モード |
|   | `Sort1Value` | 設定!$O$40 | ソート1値 |
|   | `Sort2` | 設定!$O$14 | ソート2設定 |
|   | `Sort2M` | 設定!$P$14 | ソート2モード |
|   | `Sort2Modo` | 設定!$P$41 | ソート2モード |
|   | `Sort2Value` | 設定!$O$41 | ソート2値 |
|   | `Stg2_3Time` | TrakD!$D$9 | ステージ2-3時間 |
|   | `Stg6_2Time` | TrakD!$D$14 | ステージ6-2時間 |
|   | `Stg6_4Time` | TrakD!$E$24 | ステージ6-4時間 |
|   | `Stg7_4Time` | TrakD!$E$25 | ステージ7-4時間 |
|   | `SufRtime` | TrakD!$E$21 | サフRタイム |
| ✓ | `SyainID` | 設定!$O$33 | 社員ID（1,2号機用） |
| ✓ | `SyainIDR3` | 設定!$P$33 | 社員ID（3号機用） |
| ✓ | `SyainName` | 設定!$P$36 | 社員名 |
| ✓ | `SyainPW` | 設定!$O$34 | 社員PW（1,2号機用） |
| ✓ | `SyainPWR3` | 設定!$P$34 | 社員PW（3号機用） |
| ✓ | `SyoriMSG` | 未処理品一覧!$R$4 | 処理メッセージ表示欄 |
| ✓ | `SyoriMorde` | 設定!$O$31 | 処理モード（TRUE/FALSE） |
| ✓ | `SyukabiTB` | 設定!$A$4:$F$123 | 出荷日テーブル |
|   | `SyukkaHou` | 未処理品一覧!$AO$3:$AX$3 | 出荷方法一覧 |
|   | `Tyuusyutu` | TrakD!$E$28 | 抽出設定 |
| ✓ | `Update` | 設定!$O$33:$O$35 | 更新設定（SyainID/PW/AccessTime） |
|   | `WarnMSG` | Ric3TR!$AW$1 | 3号機警告メッセージ |
|   | `YosouMSG` | 受付番号別完了!$P$2 | 予想メッセージ |
|   | `YosouSiki` | 受付番号別完了!$Q$3:$T$3 | 予想計算式 |
| ✓ | `ZaData` | 受付番号別完了!$D$6:$AE$1000 | 受付番号別完了データ |
| ✓ | `ZaikoKensu` | 未処理品一覧!$AF$2 | 在庫件数 |

---

## 4. 数式一覧

| シート | 数式件数 | 備考 |
|---|---|---|
| 未処理品一覧 | 31 | 日付計算・条件分岐表示 |
| 受付番号別完了 | 6 | VLOOKUP・条件分岐 |
| Ric3TR | 47 | VLOOKUP・集計・条件分岐 |
| 設定 | 5583 | カレンダー生成・HP参照 |
| TrakD | 23 | 時間計算・半減期 |
| ２号機重量係数 | 50 | VBA Case 文生成 |

### 4.1 未処理品一覧

| セル | 数式 | 説明 |
|---|---|---|
| `G2` | `=NowToday-1` | 前日日付 |
| `I2`〜`O2` | `=NowToday+1`〜`=NowToday+7` | 翌日〜7日後 |
| `AH2` | `=MAX(X16:X1515)` | 計画最大No |
| `G3`〜`O3` | `=IF(ISERROR(設定!AH4),"",IF(OR(RRic2=FALSE,SyoriMorde<>TRUE),"--",設定!AH4))` | 2号機HP/台数表示（条件付き） |
| `G4`〜`O4` | `=IF(ISERROR(設定!AH5),"",IF(OR(RRic3=FALSE,SyoriMorde<>TRUE),"--",設定!AH5))` | 3号機HP/台数表示（条件付き） |
| `T3` | `=IF(U3="--","","Ric2⇒")` | 2号機ラベル表示 |
| `U3` | `=IF(RRic2=TRUE,設定!O9,"--")` | 2号機現在HP表示 |
| `T4` | `=IF(U4="--","","Ric3⇒")` | 3号機ラベル表示 |
| `U4` | `=IF(RRic3=TRUE,設定!P9,"--")` | 3号機現在PP表示 |

### 4.2 受付番号別完了

| セル | 数式 | 説明 |
|---|---|---|
| `U2` | `=IF(AND(RRic2=TRUE,Hanuke<>""),Hanuke,"")` | 歯抜け判定表示 |
| `Z2` | `=NOW()` | 現在日時 |
| `Q3` | `=IF(P3="","",VLOOKUP(P3,ZaData,2,FALSE))` | 会社名取得 |
| `R3` | `=IF(P3="","",IFERROR(VLOOKUP(P3,ZaData,3,FALSE),"*"))` | 納期取得 |
| `S3` | `=IF(P3="","",IFERROR(VLOOKUP(P3,ZaData,4,FALSE),"*"))` | 出荷日取得 |
| `T3` | `=IF(Q3="","",IFERROR(VLOOKUP(P3,ZaData,8,FALSE),"*"))` | 状態取得 |

### 4.3 TrakD

| セル | 数式 | 説明 |
|---|---|---|
| `E2` | `=ROUND(D2/12,2)&" h/m"` | 年間稼動時間→月換算 |
| `E8`〜`E15` | `=ROUNDUP(D8*86400,0)` 等 | ステージ時間→秒換算 |
| `D16` | `=ROUND(G16*0.5^((TODAY()-F16)/1921),0)` | 線源強度（Co-60半減期計算） |
| `F16` | `=RIc2T` | 2号機T設定値参照 |
| `E20`〜`E25` | `=D12+D14` 等 | 合成ステージ時間 |
| `E23` | `=E4/D16*D4` | 新RRT計算 |

### 4.4 ２号機重量係数

| セル | 数式 | 説明 |
|---|---|---|
| `E3`〜`E52` | `="Case Is <=" &B3&": R2speedk ="&C3` | VBA の Select Case 文を自動生成 |

---

## 5. ボタン・マクロ対応

> ✓ = DB 更新・画面遷移・計算実行など副作用のある操作を起動するボタン

### 5.1 シート上のボタン（Form Control）

| ✓ | シート | ボタンラベル | 割り当てマクロ | 動作概要 |
|---|---|---|---|---|
| ✓ | 未処理品一覧 | 計画済 非表示 | `計画済行()` | 計画済み行の表示/非表示切替 |
| ✓ | 未処理品一覧 | 初期化 | `Ric2初期化()` | 2号機計画データの初期化 |
|   | 未処理品一覧 | ﾌｨﾙﾀ ON | `オートフィルター()` | オートフィルター ON/OFF |
| ✓ | 未処理品一覧 | 予想 | `RIC2予想計算メイン()` | 2号機照射完了予想の計算 |
| ✓ | Ric3TR | 停止記録 | `Ric3停止期間記録()` | 3号機停止期間のDB記録 |
|   | Ric3TR | 全数指定 | `未指定全数出庫()` | 未指定品の全数出庫設定 |
| ✓ | Ric3TR | 初期化 | `Ric3初期化()` | 3号機計画データの初期化 |
|   | Ric3TR | 室内表示 | `照射中表示非表示()` | 照射中品の表示切替 |
| ✓ | Ric3TR | 予想計算 | `RIC3予想計算メイン()` | 3号機照射完了予想の計算 |
| ✓ | Ric3TR | 予約記録 | `Ric3出庫記録()` | 3号機出庫順番のDB記録 |
|   | Ric3TR | 印刷 | `印刷範囲()` | 印刷範囲設定・印刷 |
|   | Ric3TR | 照射室表示 | `照射室内表示()` | 照射室内状況の表示 |

### 5.2 ユーザーフォーム上のボタン（サマリ）

| ✓ | フォーム | コントロール | イベント | 呼び出すプロシージャ | 動作概要 |
|---|---|---|---|---|---|
| ✓ | **SoutiSenntaku.frm** | comJikkou | Click | `comJikkou_Click()` → `初期化()` | 装置選択・DB読込開始 |
|   | **SoutiSenntaku.frm** | ComJyoukenn | Click | `ComJyoukenn_Click()` | SyoriSettei.frm を表示 |
|   | **SoutiSenntaku.frm** | ComSOrtRun | Click | `ComSOrtRun_Click()` → `ソート実行()` | ソート実行 |
|   | **SoutiSenntaku.frm** | Owari | Click | `Owari_Click()` | フォームを閉じる |
|   | **SoutiSenntaku.frm** | CommandButton1〜4 | Click | 各装置(1/2/3/EB)を単独選択 | 装置プリセット |
|   | **SoutiSenntaku.frm** | SubeteHyouji | Click | 全装置チェック ON | 全表示 |
|   | **SoutiSenntaku.frm** | Ric2Keikaku | Click | 2号機＋計画モード ON | 2号機計画モード |
|   | **SoutiSenntaku.frm** | Ric3Keikaku | Click | 3号機＋計画モード ON | 3号機出庫モード |
| ✓ | **SikakuKakuninn.frm** | Kakuninn | Click | `Kakuninn_Click()` → `SikakuCheck()` | 社員ID/PW で資格確認 |
|   | **SikakuKakuninn.frm** | CommandButton1 | Click | `CommandButton1_Click()` | キャンセル（Unload） |
| ✓ | **Ric3Stg.frm** | CommandButton2 | Click | `CommandButton2_Click()` | ステージ操作 |
|   | **Ric3Stg.frm** | 空P挿入 | Click | `空P挿入_Click()` | 空パレット挿入 |
|   | **TokuSyuDsp.frm** | OKButton | Click | `OKButton_Click()` | ダイアログ閉じる |
|   | **製品情報.frm** | Modori | Click | `Modori_Click()` | 戻る |

### 5.3 CommandBar に動的追加されるボタン（アドイン風）

**追加元**: **コマンドサブルーチン.bas** / `AddCmdBarButton()`
**追加タイミング**: `生産情報開始処理()` 内で呼び出し
**追加先 CommandBar**: "未計画一覧メニュー"（カスタム作成・Temporary=True）
**削除タイミング**: Workbook_BeforeClose（CommandBars はTemporary のため自動削除）

| ✓ | Caption | OnAction | FaceId | 動作概要 |
|---|---|---|---|---|
| ✓ | 最新データ | `D_UpdatetSort()` | 270 | 製品情報ファイル更新 → 装置選択画面を再表示 |
| ✓ | データ更新 | `S_Kousinn()` | 271 | 在庫・予約データの DB 更新（→ `更新処理()`） |
| ✓ | 予定行の削除 | `GyouDel()` | 293 | 選択行（予定行のみ）を DB から削除 |
|   | 在庫詳細 | `詳細表示()` | 444 | 選択行の在庫詳細を表示 |
|   | 終了 | `生産情報終了処理()` | 1088 | アプリケーション終了 |

---

## 6. VBA モジュール仕様

### 6.0 全プロシージャ一覧

全 **164** 件。

> ✓ = ユーザー操作の起点（Click イベント等） / DB I/O を実行 / 他モジュールから呼び出される Public

| ✓ | モジュール | プロシージャ | 種別 | 概要 |
|---|---|---|---|---|
| ✓ | **ThisWorkbook.cls** | `Workbook_Open()` | Event | 起動時に `生産情報開始処理()` を呼出 |
| ✓ | **ThisWorkbook.cls** | `Workbook_BeforeClose()` | Event | 終了確認・画面復帰・保存ダイアログ抑止 |
|   | **ThisWorkbook.cls** | `Workbook_BeforeSave()` | Event | デバッグ時の保存制御 |
| ✓ | **スタート処理.bas** | `生産情報開始処理()` | Sub | 休日チェック→コマンドバー設定→DB不要データ削除→装置選択 |
| ✓ | **スタート処理.bas** | `RICDB読込()` | Sub | パラメータ読込→DB読込→表示→集計 |
| ✓ | **スタート処理.bas** | `計画済未投入()` | Sub | 計画済未投入パス数のDB取得 |
| ✓ | **スタート処理.bas** | `照射日条件特殊条件Read()` | Sub | 特殊条件・照射指定日のメッセージ生成 |
|   | **スタート処理.bas** | `メッセージ()` | Sub | お知らせファイル読込（無効機能） |
|   | **スタート処理.bas** | `お知らせ()` | Function | お知らせファイル比較（無効機能） |
| ✓ | **データ取得.bas** | `ReadDataFromDB()` | Sub | zaiko+ExKeikakuX等をJOINしシートに展開 |
| ✓ | **データ取得.bas** | `ZaikoDataHyouji()` | Sub | 取得データの加工・シート書込み・会社名変換 |
| ✓ | **データ取得.bas** | `会社名()` | Sub | 会社コード→会社名変換 |
| ✓ | **コマンド実行.bas** | `初期化()` | Sub | 装置選択後の主制御（表示方法決定→RICDB読込→モード分岐） |
| ✓ | **コマンド実行.bas** | `ソート実行()` | Sub | 第1/第2ソートを実行 |
| ✓ | **コマンドサブルーチン.bas** | `AddCmdBarButton()` | Sub | CommandBar "未計画一覧メニュー" を作成・ボタン登録 |
| ✓ | **コマンドサブルーチン.bas** | `D_UpdatetSort()` | Sub | CommandBar「最新データ」→製品情報更新→装置選択再表示 |
| ✓ | **コマンドサブルーチン.bas** | `S_Kousinn()` | Sub | CommandBar「データ更新」→`更新処理()`呼出 |
| ✓ | **コマンドサブルーチン.bas** | `GyouDel()` | Sub | CommandBar「予定行の削除」→DB削除 |
| ✓ | **更新在庫.bas** | `更新処理()` | Sub | 資格確認→`ZaikoKousinn`→`YoyakuKousinn`→再読込 |
| ✓ | **更新在庫.bas** | `ZaikoKousinn()` | Sub | `ExKeikakuX` の UPDATE（kakunin/syukkabi/syuhouhou/bikou1） |
| ✓ | **更新在庫.bas** | `RunSort()` | Sub | 指定列でソート実行 |
| ✓ | **更新在庫.bas** | `ロギング記録()` | Sub | 操作ログ記録 |
|   | **更新在庫.bas** | `古いロギング削除()` | Sub | 古いログの削除 |
| ✓ | **更新予約.bas** | `YoyakuKousinn()` | Sub | `ExYoyakuX` の INSERT/UPDATE（全22カラム） |
| ✓ | **更新予約.bas** | `予約データ記録()` | Sub | 予約1件の DB 書込み |
| ✓ | **更新予約.bas** | `CheckYoyakuNo()` | Function | 予約番号重複チェック |
| ✓ | **R3出庫記録処理.bas** | `Ric3出庫記録()` | Sub | `ExR3SYukko` を DELETE→INSERT |
| ✓ | **R3計画作成.bas** | `Ric3未照射品計画()` | Sub | 3号機未照射品の計画配置 |
|   | **R3計画作成.bas** | `照射中表示非表示()` | Sub | 照射中行の表示/非表示 |
|   | **R3計画作成.bas** | `照射中表示非表示設定()` | Sub | 表示設定の適用 |
| ✓ | **SQL_Execution.bas** | `Open_oraconDB()` | Sub | Oracle DB 接続 |
| ✓ | **SQL_Execution.bas** | `SQL_Exe()` | Sub | SQL 文実行 |
| ✓ | **SQL_Execution.bas** | `SQL_INSERT_UPDATE()` | Sub | INSERT or UPDATE 実行 |
| ✓ | **SQL_Execution.bas** | `SQL_Delete()` | Sub | DELETE 実行 |
| ✓ | **SQL_Execution.bas** | `Disp_Sheet()` | Sub | SELECT結果をシートに展開 |
| ✓ | **SQL_Execution.bas** | `Set_Array()` | Sub | SELECT結果を配列に格納 |
|   | **SQL_Execution.bas** | `SQL_Close()` | Sub | DB 接続クローズ |
| ✓ | **DB不要データ削除.bas** | `計画削除()` | Sub | ExKeikakuX の孤児レコード削除 |
| ✓ | **DB不要データ削除.bas** | `予約削除()` | Sub | ExYoyakuX の無効・古い予約削除 |
| ✓ | **資格確認.bas** | `SikakuCheck()` | Sub | SHAINMST照合・認証 |
| ✓ | **R2計画作成.bas** | `計画作成()` | Sub | 2号機計画作成 |
|   | **R2計画作成.bas** | `計画済行()` | Sub | 計画済表示コマンド |
|   | **R2計画作成.bas** | `R2計画済表示非表示()` | Sub | 計画済行の表示/非表示 |
|   | **R2計画作成.bas** | `同線量表示()` | Sub | 同線量品の表示 |
| ✓ | **R2予想計算メイン.bas** | `Ric2初期化()` | Sub | 2号機予想の初期化 |
| ✓ | **R2予想計算メイン.bas** | `RIC2予想計算メイン()` | Sub | 2号機予想計算（メイン制御） |
|   | **R2予想計算メイン.bas** | `予想初期化()` | Sub | 予想データ初期化 |
|   | **R2予想計算メイン.bas** | `計算Main()` | Sub | 計算ループ本体 |
|   | **R2予想計算メイン.bas** | `SlineRun()` | Sub | 線量計走行シミュレーション |
|   | **R2予想計算メイン.bas** | `Ric2Trking()` | Sub | 2号機トラッキング |
|   | **R2予想計算メイン.bas** | `Ric3Picup()` | Function | 3号機ピックアップ |
|   | **R2予想計算メイン.bas** | `TrackData()` | Sub | トラックデータ処理 |
|   | **R2予想計算メイン.bas** | `bhjiu()` | Sub | （内部処理） |
|   | **R2予想計算メイン.bas** | `StgSyokiSet()` | Sub | ステージ初期設定 |
| ✓ | **R2予想計算メイン.bas** | `在庫読込と範囲名と納期日()` | Sub | 在庫データ読込 |
| ✓ | **R2予想計算メイン.bas** | `受付番号別完了表示()` | Sub | 受付番号別完了シートへ書込み |
|   | **R2予想計算メイン.bas** | `時刻別完了()` | Sub | 時刻別完了計算 |
|   | **R2予想計算メイン.bas** | `線量計別完了表示()` | Sub | 線量計別完了表示 |
|   | **R2予想計算メイン.bas** | `StgDisp()` | Sub | ステージ表示 |
|   | **R2予想計算メイン.bas** | `予想実行済のデータ初期化()` | Function | 予想実行済データのリセット |
| ✓ | **R3予想計算.bas** | `Ric3初期化()` | Sub | 3号機予想の初期化 |
| ✓ | **R3予想計算.bas** | `RIC3予想計算メイン()` | Sub | 3号機予想計算（メイン制御） |
|   | **R3予想計算.bas** | `SorsPass1()` | Sub | パスソート1 |
|   | **R3予想計算.bas** | `R3トラキング表示()` | Sub | 3号機トラッキング表示 |
| ✓ | **R3予想計算.bas** | `TrSyukkoDataRead()` | Sub | 出庫データ読込 |
|   | **R3予想計算.bas** | `空パレット挿入()` | Sub | 空パレット挿入 |
|   | **R3予想計算.bas** | `KanNo()` | Function | 搬入No取得 |
|   | **R3予想計算.bas** | `Ric3完了表示()` | Sub | 3号機完了表示 |
| ✓ | **R3Data読込.bas** | `Ric3号機予想データRead()` | Sub | 3号機予想データのDB読込 |
| ✓ | **R3Data読込.bas** | `出荷日等式をコピー()` | Sub | 出荷日計算式のコピー |
|   | **R3Data読込.bas** | `Ric3TRPass()` | Sub | Ric3TR パスデータ処理 |
| ✓ | **R3Data読込.bas** | `製品情報表示()` | Sub | 製品情報フォーム表示 |
| ✓ | **R3停止期間記録読出.bas** | `Ric3停止期間記録()` | Sub | 3号機停止期間のDB記録 |
| ✓ | **R3停止期間記録読出.bas** | `Ric3停止期間読込()` | Sub | 3号機停止期間のDB読込 |
|   | **R3停止期間記録読出.bas** | `R3停止ファクター()` | Sub | 停止ファクター計算 |
| ✓ | **R3Stg1挿入取消.bas** | `出庫設定()` | Sub | 出庫順設定（セルクリック） |
|   | **R3Stg1挿入取消.bas** | `指定取消()` | Sub | 出庫順指定取消 |
|   | **R3出庫設定.bas** | `未指定全数出庫()` | Sub | 未指定品を全数出庫設定 |
| ✓ | **R3照射室.bas** | `R3Pass2Stg()` | Sub | パス→ステージ変換 |
|   | **R3照射室.bas** | `STG1消去()` | Sub | ステージ1消去 |
|   | **R3照射室.bas** | `St1Sv表示()` | Sub | Stg1サーバ表示 |
|   | **R3照射室.bas** | `Stg1Form()` | Sub | ステージ1フォーム |
|   | **R3照射室.bas** | `R3Pass1Stg()` | Sub | パス1→ステージ変換 |
|   | **R3照射室.bas** | `照射室内表示()` | Sub | 照射室内状況表示 |
| ✓ | **在庫詳細.bas** | `詳細表示()` | Sub | 在庫詳細の表示（CommandBar「在庫詳細」） |
|   | **在庫詳細.bas** | `anann()` | Sub | （内部処理） |
| ✓ | **在庫詳細.bas** | `在庫詳細表示()` | Sub | 在庫詳細の DB 取得・表示 |
| ✓ | **在庫詳細.bas** | `照射中と未投入処理量()` | Sub | 照射中/未投入の処理量集計 |
| ✓ | **装置稼働状況.bas** | `稼働状況()` | Sub | 装置稼働状況のDB取得 |
|   | **装置稼働状況.bas** | `ahjk()` | Sub | （内部処理） |
|   | **装置稼働状況.bas** | `ExHenvan()` | Function | 変換関数 |
|   | **装置稼働状況.bas** | `ExDate()` | Function | 日付変換関数 |
| ✓ | **サブルーチン.bas** | `GamenSettei()` | Sub | 画面設定 |
| ✓ | **サブルーチン.bas** | `GamennFukki()` | Sub | 画面復帰 |
|   | **サブルーチン.bas** | `Sikicopy()` | Sub | 式コピー |
| ✓ | **サブルーチン.bas** | `KyuujituRead()` | Sub | 休日データDB読込 |
| ✓ | **サブルーチン.bas** | `KyuujituSuu()` | Function | 未登録休日数チェック |
| ✓ | **サブルーチン.bas** | `HikitoriRead()` | Sub | 引取業者DB読込 |
| ✓ | **サブルーチン.bas** | `Ric3ParaM_Read()` | Sub | 3号機パラメータDB読込 |
|   | **サブルーチン.bas** | `オートフィルター()` | Sub | オートフィルター設定 |
| ✓ | **サブルーチン.bas** | `Ric23ParaM_Read()` | Sub | 2,3号機パラメータDB読込 |
|   | **サブルーチン.bas** | `BubbleSort2()` | Sub | バブルソート |
|   | **サブルーチン.bas** | `コメント削除()` | Sub | コメント削除 |
|   | **サブルーチン.bas** | `コメント追加()` | Sub | コメント追加 |
| ✓ | **画面色と表示列設定.bas** | `GyouHaikei()` | Sub | 行背景色設定 |
| ✓ | **画面色と表示列設定.bas** | `列表示設定()` | Sub | 列の表示/非表示設定 |
|   | **照射中パス集計.bas** | `Ric3TR()` | Sub | Ric3照射中残パス集計 |
|   | **生産情報終了.bas** | `生産情報終了処理()` | Sub | アプリケーション終了処理 |
| ✓ | **画面クリア_遷移1.bas** | `Ric3TRに戻る()` | Sub | Ric3TR シートへ遷移 |
| ✓ | **画面クリア_遷移1.bas** | `未処理品一覧に戻る()` | Sub | 未処理品一覧シートへ遷移 |
| ✓ | **画面クリア_遷移1.bas** | `受付番号別完了遷移()` | Sub | 受付番号別完了シートへ遷移 |
|   | **画面クリア_遷移1.bas** | `KeikakuHClear()` | Sub | 計画HPクリア |
|   | **画面クリア_遷移1.bas** | `画面CLS()` | Sub | 画面クリア |
|   | **R2印刷範囲.bas** | `RIC2印刷範囲()` | Sub | 2号機印刷範囲設定 |
|   | **R3出庫印刷.bas** | `印刷範囲()` | Sub | 3号機印刷範囲設定 |
|   | **R3出庫印刷.bas** | `印刷範囲設定()` | Sub | 印刷範囲の適用 |
|   | **ユーティリティーイベント有効.bas** | `R3計画クリア()` | Sub | R3計画クリア |
|   | **R2未照射急ぎ順.bas** | `急ぎ順計算()` | Sub | 2号機急ぎ順計算 |
|   | **R3最終完了日.bas** | `最終完了日時計算()` | Sub | 3号機最終完了日時計算 |
|   | **R3最終完了日.bas** | `納期を年月日変換()` | Sub | 納期の日付変換 |
|   | **オートフィルタ.bas** | `オートフィルター設定()` | Sub | オートフィルター設定 |
|   | **R2指定日チェック.bas** | `計画係数()` | Function | 計画係数計算 |
|   | **R2指定日チェック.bas** | `FN_ConvertDay()` | Function | 日付変換 |
| ✓ | **R3タイマーチェック.bas** | `R3タイマーチェック処理()` | Sub | 3号機タイマーチェック |
|   | **R3フォーム_セル高さ調整.bas** | `Ric3StgForm_Initialize()` | Sub | Ric3Stg フォーム初期化 |
|   | **R3フォーム_セル高さ調整.bas** | `Ric3TR高さ調整()` | Sub | Ric3TR 行高さ調整 |
|   | **シェープ表示非表示.bas** | `右矢印表示()` | Sub | 「未処理品一覧」右矢印＋テキストボックス3 を表示 |
|   | **シェープ表示非表示.bas** | `右矢印非表示()` | Sub | 「未処理品一覧」右矢印＋テキストボックス3 を非表示 |
|   | **シェープ表示非表示.bas** | `予想ボタン表示()` | Sub | 「未処理品一覧」ボタン7（予想）を表示 |
|   | **シェープ表示非表示.bas** | `予想ボタン非表示()` | Sub | 「未処理品一覧」ボタン7（予想）を非表示 |
|   | **シェープ表示非表示.bas** | `初期化ボタン表示()` | Sub | 「未処理品一覧」Button5（初期化）を表示 |
|   | **シェープ表示非表示.bas** | `初期化ボタン非表示()` | Sub | 「未処理品一覧」Button5（初期化）を非表示 |
|   | **シェープ表示非表示.bas** | `計画済ボタン表示()` | Sub | 「未処理品一覧」Button8（計画済）を表示 |
|   | **シェープ表示非表示.bas** | `計画済ボタン非表示()` | Sub | 「未処理品一覧」Button8（計画済）を非表示 |
|   | **シェープ表示非表示.bas** | `警告メッセージ非表示()` | Sub | 「未処理品一覧」テキストボックス5 を非表示 |
|   | **シェープ表示非表示.bas** | `警告メッセージ表示()` | Sub | 「未処理品一覧」テキストボックス5 を表示 |
| ✓ | **R2計画指定.bas** | `R2計画順指定()` | Sub | 2号機計画順指定 |
| ✓ | **R2速度チェック.bas** | `R2速度チェック処理()` | Sub | 2号機速度チェック |
|   | **R2速度チェック.bas** | `R2SpeedK()` | Function | 重量→速度係数変換 |
| ✓ | **ファンクション.bas** | `ChengeSyukkaDay()` | Function | 出荷日変換 |
| ✓ | **ファンクション.bas** | `SyukkaHouhou()` | Function | 出荷方法取得 |
|   | **ファンクション.bas** | `ExchengeDay()` | Function | 日付→文字列変換 |
|   | **ファンクション.bas** | `ExchengeDATE()` | Function | 文字列→日付変換 |
|   | **ファンクション.bas** | `日付()` | Function | 日付変換 |
|   | **ファンクション.bas** | `照射日()` | Function | 照射日判定 |
|   | **ファンクション.bas** | `NGCount()` | Function | NG件数カウント |
|   | **ファンクション.bas** | `MisuuN()` | Function | 未数N取得 |
| ✓ | **Sheet1.cls** | `Worksheet_BeforeDoubleClick()` | Event | 「未処理品一覧」ダブルクリック |
|   | **Sheet1.cls** | `Worksheet_BeforeRightClick()` | Event | 「未処理品一覧」右クリック |
| ✓ | **Sheet3.cls** | `Worksheet_Activate()` | Event | 「Ric3TR」アクティブ時 |
| ✓ | **Sheet3.cls** | `Worksheet_Change()` | Event | 「Ric3TR」セル変更時 |
|   | **Sheet3.cls** | `Worksheet_Deactivate()` | Event | 「Ric3TR」非アクティブ時 |
|   | **Sheet3.cls** | `Worksheet_BeforeRightClick()` | Event | 「Ric3TR」右クリック |
|   | **Sheet3.cls** | `Worksheet_SelectionChange()` | Event | 「Ric3TR」選択変更時 |
| ✓ | **出荷方法更新.bas** | `製品情報ファイル更新()` | Sub | 製品情報テーブルの更新 |
|   | **出荷方法更新.bas** | `古い受付番号削除()` | Sub | 古い受付番号の削除 |
|   | **ErrCheck.bas** | `ErrCheckMsg()` | Sub | エラーメッセージ表示 |

---

## 7. ユーザーフォーム仕様

> ✓ = ユーザー入力を受け付ける、またはイベントで業務処理を起動するコントロール

### 7.1 SoutiSenntaku.frm（装置選択フォーム）

**目的**: 起動直後に表示。対象装置（1/2/3号機/EB）・表示方法・ソート条件を選択し「処理方法実行」で初期読込を開始する。

#### コントロール一覧

| ✓ | コントロール | 種別 | 用途 |
|---|---|---|---|
| ✓ | comJikkou | CommandButton | 処理方法実行（メイン起動ボタン） |
|   | ComJyoukenn | CommandButton | 条件設定（SyoriSettei.frm 表示） |
|   | ComSOrtRun | CommandButton | ソート実行 |
|   | Owari | CommandButton | 終了（フォーム閉じ） |
|   | CommandButton1〜4 | CommandButton | 装置プリセット（1号機/2号機/3号機/EB） |
|   | CheckBox1〜4 | CheckBox | 装置選択（1号機/2号機/3号機/EB） |
|   | SubeteHyouji | OptionButton | 全て表示 |
|   | Misyousya | OptionButton | 未照射品のみ |
|   | AllZaiko | OptionButton | 在庫全て表示 |
| ✓ | Ric2Keikaku | OptionButton | 2号機計画モード（→2号機完了予想） |
| ✓ | Ric3Keikaku | OptionButton | 3号機出庫モード（→3号機完了予想） |
|   | OptUke〜OptSoku | OptionButton | 特殊条件表示（受付/計画/投入/出荷/測定） |
|   | ComboBox1/2 | ComboBox | 第1/第2ソート項目選択 |
|   | OptionButton1〜5 | OptionButton | ソート昇順/降順 |
|   | TextBoxR1〜R3 | TextBox | 装置稼働状況表示 |
|   | LaR1Time〜LaR3Time | Label | 装置稼働時間表示 |
|   | Frame1〜8 | Frame | グループ枠 |

### 7.2 SikakuKakuninn.frm（資格確認フォーム）

**目的**: DB 更新操作前に表示。社員ID/パスワードを入力し、SHAINMST で資格を確認する。

#### コントロール一覧

| ✓ | コントロール | 種別 | 用途 |
|---|---|---|---|
| ✓ | SyainnNo | TextBox | 社員番号入力 |
| ✓ | PassWord | TextBox | パスワード入力 |
| ✓ | Kakuninn | CommandButton | 確認実行（→`SikakuCheck()`） |
|   | CommandButton1 | CommandButton | キャンセル |
|   | Syainn | Label | ラベル |
|   | Label1 | Label | ラベル |

### 7.3 SyoriSettei.frm（処理条件設定フォーム）

**目的**: 2/3号機のパラメータ設定・業務資格の登録/削除を行う。
**フォームキャプション**: "ファクター登録・処理"
**モードレス表示**: ShowModal=False

#### コントロール一覧

| ✓ | コントロール | 種別 | 親フレーム | 用途 |
|---|---|---|---|---|
| ✓ | ComKakunin | CommandButton | FramTouroku | 登録者資格確認（→`SikakuCheck`相当のDB照合） |
| ✓ | ComTouroku | CommandButton | FramHiTouroku | 営業業務資格登録（SHAINMST の `kshika='1'` に UPDATE） |
| ✓ | ComSakujyo | CommandButton | FramHiTouroku | 営業業務資格削除（SHAINMST の `kshika='0'` に UPDATE） |
| ✓ | Ric2Set | CommandButton | — | 2号機パラメータ設定（`ExKanriTB` SIKIBETU=4 を UPDATE） |
| ✓ | Ric3Set | CommandButton | — | 3号機パラメータ設定（`ExKanriTB` SIKIBETU=5 を UPDATE） |
|   | Ric2Hyouji | CommandButton | — | 2号機設定値の DB 読込・表示 |
|   | Ric3Hyouji | CommandButton | — | 3号機設定値の DB 読込・表示 |
|   | Ric2NowHyouji | CommandButton | — | 2号機現在値を「設定」シートから表示 |
|   | Ric3NowHyouji | CommandButton | — | 3号機現在値を「設定」シートから表示 |
|   | CommandButton7 | CommandButton | — | 閉じる（Unload → SoutiSenntaku.Show） |
|   | ComSyoriNouryoku | CommandButton | — | 処理能力設定（フォーム再表示） |
|   | TourokuCD | TextBox | FramTouroku | 登録者社員番号入力 |
|   | TourokuPW | TextBox | FramTouroku | 登録者パスワード入力 |
|   | HiTouCD | TextBox | FramHiTouroku | 被登録者社員番号入力 |
|   | Ric2Sokudo | TextBox | — | 2号機速度入力（Change で HP 自動計算） |
|   | Ric2Nenn / Ric2Tuki / Ric2Hi | TextBox×3 | — | 2号機線源増量 年/月/日 |
|   | Ric3Sokudo | TextBox | — | 3号機タイマー入力（Change で PP 自動計算） |
|   | Ric3Nenn / Ric3Tuki / Ric3Hi | TextBox×3 | — | 3号機設定 年/月/日 |
|   | LaSyName | Label | FramHiTouroku | 被登録者社員名表示 |
|   | Ric2HPset | Label | — | 2号機 HP 設定値（速度入力で自動算出: `速度×1.4885`） |
|   | Ric3PPSet | Label | — | 3号機 PP 設定値（タイマー入力で自動算出: `82800/タイマー`） |
|   | LRic2Sokudo / LRic2Nenn / LRic2Tuki / LRic2Hi / LRic2HP | Label×5 | — | 2号機現在値表示 |
|   | LRic3Sokudo / LRic3Nenn / LRic3Tuki / LRic3Hi / LRic3PP | Label×5 | — | 3号機現在値表示 |
|   | Label27〜Label30 | Label×4 | — | 固定ラベル（見出し） |
|   | FramTouroku | Frame | — | 登録者資格確認グループ |
|   | FramHiTouroku | Frame | — | 被登録者操作グループ |
|   | Frame1〜Frame4 | Frame×4 | — | レイアウト用グループ枠 |

#### イベント一覧

| ✓ | イベント | 動作概要 |
|---|---|---|
| ✓ | `ComKakunin_Click()` | 営業業務資格登録の登録者資格確認（SHAINMST で hshika='1' AND kshika='1' を照合） |
| ✓ | `ComTouroku_Click()` | 営業業務資格登録（SHAINMST の kshika を '1' に UPDATE） |
| ✓ | `ComSakujyo_Click()` | 営業業務資格削除（SHAINMST の kshika を '0' に UPDATE） |
|   | `HiTouCD_AfterUpdate()` | 被登録者の社員名を SHAINMST から取得し LaSyName に表示 |
| ✓ | `Ric2Hyouji_Click()` | ExKanriTB（SIKIBETU=4）から 2号機設定値を DB 読込・表示 |
| ✓ | `Ric2Set_Click()` | 2号機パラメータを「設定」シートに書込み → ExKanriTB を UPDATE |
|   | `Ric2NowHyouji_Click()` | 2号機現在速度/HP を「設定」シートから Label に表示 |
|   | `Ric2Sokudo_Change()` | 速度入力に連動して Ric2HPset を自動計算（`速度×1.4885`） |
| ✓ | `Ric3Hyouji_Click()` | ExKanriTB（SIKIBETU=5）から 3号機設定値を DB 読込・表示 |
| ✓ | `Ric3Set_Click()` | 3号機パラメータを「設定」シートに書込み → ExKanriTB を UPDATE |
|   | `Ric3NowHyouji_Click()` | 3号機現在タイマー/PP を「設定」シートから Label に表示 |
|   | `Ric3Sokudo_Change()` | タイマー入力に連動して Ric3PPSet を自動計算（`82800/タイマー`） |
|   | `CommandButton7_Click()` | フォーム閉じて SoutiSenntaku.Show |
|   | `ComSyoriNouryoku_Click()` | フォーム再表示 |

### 7.4 Ric3Stg.frm（3号機ステージフォーム）

**目的**: 3号機の照射室内状況をリアルタイム表示。各ステージの線量計番号・パス数・出庫順を表示する。

#### コントロール種別サマリ

| 種別 | 件数 | 内訳 |
|---|---|---|
| Label | 94+ | ステージ2（Sv01〜Sv16, SPas01〜SPas16, JPas01〜JPas16, SvP01〜SvP16 = 64件）＋ステージ1（St1Sv01〜St1Sv10, St1Pas01〜St1Pas10, St1SvP01〜St1SvP10 = 30件）＋固定ラベル |
| TextBox | 2 | TxPasu, TxSuu |
| CommandButton | 2 | CommandButton2, 空P挿入 |
| Frame | 1 | FrmKaraP |

#### コントロール一覧

| ✓ | コントロール | 種別 | 用途 |
|---|---|---|---|
|   | Sv01〜Sv16 | Label×16 | ステージ2 の各スロット（線量計番号表示・背景色で状態を示す） |
|   | St1Sv01〜St1Sv10 | Label×10 | ステージ1 の各スロット（出庫待ち線量計番号） |
|   | SPas01〜SPas16 | Label×16 | ステージ2 各スロットのパス数表示 |
|   | JPas01〜JPas16 | Label×16 | ステージ2 各スロットの受番表示 |
|   | SvP01〜SvP16 | Label×16 | ステージ2 各スロットの SvP 表示 |
|   | St1Pas01〜St1Pas10 | Label×10 | ステージ1 各スロットのパス数表示 |
|   | St1SvP01〜St1SvP10 | Label×10 | ステージ1 各スロットの SvP 表示 |
|   | Label102, Label103 | Label×2 | ステージ見出し |
|   | Label138, Label139, Label142〜Label144 | Label×5 | 固定ラベル（説明文） |
| ✓ | TxPasu | TextBox | 空パレット挿入時のパス数入力 |
| ✓ | TxSuu | TextBox | 空パレット台数入力 |
| ✓ | CommandButton2 | CommandButton | 画面非表示（`Ric3Stg.Hide`） |
| ✓ | 空P挿入 | CommandButton | 空パレット挿入（→ `空パレット挿入()` 呼出し） |
|   | FrmKaraP | Frame | 空パレット操作グループ枠 |

### 7.5 製品情報.frm（製品情報表示フォーム）

**目的**: 選択した在庫品の製品特殊条件・適否情報を表示する。

#### コントロール一覧

| コントロール | 種別 | 用途 |
|---|---|---|
| TextTokuSyu | TextBox | 特殊条件表示 |
| TekiMi / HutekiMi | TextBox | 適/不適 未照射 |
| TekiSu / HutekiSu | TextBox | 適/不適 照射済 |
| HenSuMi / HenSuSu | TextBox | 変数 未/済 |
| Modori | CommandButton | 戻るボタン |

### 7.6 TokuSyuDsp.frm / 7.7 UserForm処理中.frm

- **TokuSyuDsp.frm**: 特殊条件ダイアログ（OKButton のみ）
- **UserForm処理中.frm**: 処理中メッセージ表示（Label1 のみ。ユーザー操作なし）

---

## 8. DB 接続・外部連携

### 8.1 ODBC 接続設定

| DSN 名 | UID | PWD | 用途 |
|---|---|---|---|
| `ricdb` | `ric` | `t6101` | 全テーブルへのメインアクセス |

> **DB サーバー IP**: 163.59.144.156

### 8.2 テーブル一覧（参照/更新区分付き）

> ✓ = INSERT / UPDATE / DELETE の対象テーブル（参照のみのテーブルは ✓ なし）

| ✓ | テーブル名 | 区分 | 主な用途 | キー列 | 参照/更新列 |
|---|---|---|---|---|---|
| ✓ | `zaiko` | 参照 | γ線照射在庫マスタ | `uno` | `kaisyacd`, `misyousu`, `syouso`, `nyukabi`, `siteisn`, `pass` |
| ✓ | `ExKeikakuX` | **参照＋更新** | 生産計画 | `uno` | 更新: `kakunin`, `syukkabi`, `syuhouhou`, `bikou1` |
| ✓ | `ExYoyakuX` | **参照＋更新** | 予約データ | `yoyakuno` | 更新: 全22カラム（`nyukabi`, `kaisyacd`, `siteisn`, `pass`, `nouki`, `kakunin`, `syukkabi`, `syuhouhou`, `bikou`, `updateid`, `updateday`, `yuukou` 等） |
| ✓ | `ExR3SYukko` | **更新（DELETE+INSERT）** | 3号機出庫順 | `senkno` + `siteibi` | `syukkojyun`, `senkno`, `siteibi` |
|   | `tsyjmst` | 参照 | 特殊条件マスタ | `(複合)` | `tokjyo1`〜`5`, `dspno` |
|   | `ExSeihinJ` / `ExSeihinZ` | 参照 | 製品情報 | `kaisyacd` + `sehncd` | `tumikae` 等 |
|   | `sejofile` | 参照 | 生産条件 | — | `nyukajoken`, `shukajoken` |
|   | `syouk2` / `syouj2` | 参照 | 処理・照射中 | — | `syoriflg`, `pass` |
|   | `torak3` / `TORAK` | 参照 | トラッキング | — | `STNO`, `TRKDATA` |
| ✓ | `ExKanriTB` | 参照 | 装置管理パラメータ | `SIKIBETU` | `RICVM`, `HPPP` |
|   | `ExYasumiX` | 参照 | 休日テーブル | — | `kyuujitu1` |
| ✓ | `SHAINMST` | 参照 | 社員マスタ（資格確認） | `shano` | `shaname`, `shask`, `hshika`, `kshika`, `U3shika` |
|   | `zaikoeb` / `sehmst` | 参照 | EB在庫（2022/09改修） | — | — |

### 8.3 主要 SQL（γ在庫読込・抜粋）

```sql
SELECT z.Uno, z.nyukabi, TO_NUMBER(z.kaisyacd), TRIM(z.kainame),
       z.siteisn*1, z.pass, z.misyousu*1, k.kakunin, k.syukkabi, ...
FROM zaiko z, ExKeikakuX k, tsyjmst t, ExSeihinJ j, sejofile s, ExSeihinZ zs
WHERE z.uno = k.uno(+) AND ...
ORDER BY z.kaisyacd, z.uno
```

### 8.4 外部ファイル連携

- ログファイル: VBA により `Open ... For Output` で書込み・`Kill` で削除（パスは動的生成）
- お知らせファイル: `\\163.59.144.156\ExRicSys\生産情報一覧お知らせ.txt`（現在は無効機能）

---

## 9. データフロー

各フローは「起点 → 処理 → 結果」の粒度で記述する。

### 9.1 起動フロー

| # | 起点 | 処理 | 結果 |
|---|---|---|---|
| 1 | ユーザーがファイルを開く | 📄 `Workbook_Open()` / **ThisWorkbook.cls** | `生産情報開始処理()` を呼出し |
| 2 | 上記 | 📄 `生産情報開始処理()` / **スタート処理.bas** | 🗄️ 休日チェック → 🖥️ Ric3Stg 初期化 → 📄 `AddCmdBarButton()` → 🗄️ `計画削除()` → 🗄️ `予約削除()` → 📊 `画面CLS()` |
| 3 | 上記 | 🖥️ **SoutiSenntaku.frm** 表示 | ユーザーの装置選択待ち |

#### ツリー図（補助）

```
(ファイルを開く)
└─ 📄 Workbook_Open              [ThisWorkbook.cls]
   └─ 📄 生産情報開始処理          [スタート処理.bas]
      ├─ 🗄️ KyuujituSuu（休日チェック）
      ├─ 🖥️ Ric3StgForm_Initialize
      ├─ 📄 AddCmdBarButton        [コマンドサブルーチン.bas]
      ├─ 🗄️ 計画削除 / 予約削除    [DB不要データ削除.bas]
      ├─ 📊 画面CLS               [画面クリア_遷移1.bas]
      └─ 🖥️ SoutiSenntaku.Show    [SoutiSenntaku.frm]
```

### 9.2 データ表示フロー

| # | 起点 | 処理 | 結果 |
|---|---|---|---|
| 1 | 🖥️（SoutiSenntaku → 処理方法実行 Click） | 🖥️ `comJikkou_Click()` / **SoutiSenntaku.frm** | フォーム非表示 → `初期化()` 呼出し |
| 2 | 上記 | 📄 `初期化()` / **コマンド実行.bas** | 📊 装置選択条件を Range に書込み → 📄 `ReqNo` 算出 |
| 3 | 上記 | 📄 `RICDB読込()` / **スタート処理.bas** | 🗄️ パラメータ読込 → 🗄️ DB 読込 → 📊 表示 → 📊 集計 |
| 4 | 上記 | 🗄️ `ReadDataFromDB()` / **データ取得.bas** | `zaiko` + `ExKeikakuX` 等を JOIN し配列取得 |
| 5 | 上記 | 📊 `ZaikoDataHyouji()` / **データ取得.bas** | 「未処理品一覧」の `Keikaku` に一括書込み |
| 6 | 上記 | 📊 `Ric3TR()` / **照射中パス集計.bas** | 照射中残パス集計 |
| 7 | [2号機計画モード] | 📊 `R2速度チェック処理()` → 📊 `急ぎ順計算()` → 📊 `GyouHaikei()` | 2号機計画画面表示完了 |
| 8 | [3号機出庫モード] | 📊 「Ric3TR」シート表示 → 🗄️ `Ric3号機予想データRead()` → 📊 `Ric3未照射品計画()` → 📊 `照射中表示非表示()` | 3号機出庫画面表示完了 |
| 9 | [通常モード] | 📊 `ソート実行()` → 📊 `GyouHaikei()` | 未処理品一覧の表示完了 |

#### ツリー図（補助）

```
（SoutiSenntaku → 処理方法実行 Click）
└─ 🖥️ comJikkou_Click            [SoutiSenntaku.frm]
   └─ 📄 初期化                    [コマンド実行.bas]
      ├─ 📊 装置選択条件を Range 書込み
      ├─ 📄 ReqNo 算出
      └─ 📄 RICDB読込              [スタート処理.bas]
         ├─ 🗄️ Ric23ParaM_Read    [サブルーチン.bas]（パラメータDB読込）
         ├─ 🗄️ ReadDataFromDB     [データ取得.bas]（zaiko+ExKeikakuX JOIN→配列）
         ├─ 📊 ZaikoDataHyouji    [データ取得.bas]（Keikaku に一括書込み）
         ├─ 📊 Ric3TR             [照射中パス集計.bas]
         └─ [モード分岐]
            ├─ [2号機計画] 📊 R2速度チェック処理 → 📊 急ぎ順計算 → 📊 GyouHaikei
            ├─ [3号機出庫] 📊 Ric3号機予想データRead → 📊 Ric3未照射品計画 → 📊 照射中表示非表示
            └─ [通常]     📊 ソート実行 → 📊 GyouHaikei
```

### 9.3 更新フロー

| # | 起点 | 処理 | 結果 |
|---|---|---|---|
| 1 | 📊（「未処理品一覧」でユーザーがセル入力） | 📊 ユーザー入力 | `Keikaku` 領域のセルに変更が入る |
| 2 | 📄（CommandBar「データ更新」Click） | 📄 `S_Kousinn()` / **コマンドサブルーチン.bas** | 📄 計画モードチェック → `更新処理()` 呼出し |
| 3 | 上記 | 📄 `更新処理()` / **更新在庫.bas** | 📄 経過時間チェック → 資格確認判定 |
| 4 | [経過時間 > 閾値] | 🖥️ **SikakuKakuninn.frm** 表示 → 🗄️ `SikakuCheck()` / **資格確認.bas** | 社員ID/PW を `SHAINMST` で照合 |
| 5 | 認証成功 | 🗄️ `ZaikoKousinn()` / **更新在庫.bas** | `ExKeikakuX` の `kakunin`, `syukkabi`, `syuhouhou`, `bikou1` を UPDATE |
| 6 | 上記 | 🗄️ `YoyakuKousinn()` / **更新予約.bas** | `ExYoyakuX` の全22カラムを INSERT or UPDATE |
| 7 | 上記 | 🗄️ `ReadDataFromDB()` → 📊 `ZaikoDataHyouji()` | DB 再読込・表示更新 |
| 8 | 上記 | 📊 `RunSort()` → 📊 `GyouHaikei()` | ソート・行背景色適用 |

#### ツリー図（補助）

```
（CommandBar「データ更新」Click）
└─ 📄 S_Kousinn                   [コマンドサブルーチン.bas]
   ├─ 📄 計画モードチェック（mpKeiMode=2/3 なら中断）
   └─ 📄 更新処理                  [更新在庫.bas]
      ├─ 📄 経過時間チェック（AccessTime vs KeikaTime）
      ├─ [閾値超過] 🖥️ SikakuKakuninn.frm 表示
      │   └─ 🗄️ SikakuCheck       [資格確認.bas]（SHAINMST 照合）
      ├─ 🗄️ ZaikoKousinn          [更新在庫.bas]
      │   └─ ExKeikakuX UPDATE（kakunin, syukkabi, syuhouhou, bikou1）
      ├─ 🗄️ YoyakuKousinn         [更新予約.bas]
      │   └─ ExYoyakuX INSERT/UPDATE（全22カラム）
      ├─ 🗄️ ReadDataFromDB        [データ取得.bas]（DB 再読込）
      ├─ 📊 ZaikoDataHyouji       [データ取得.bas]（表示更新）
      ├─ 📊 RunSort               [更新在庫.bas]（ソート）
      └─ 📊 GyouHaikei            [画面色と表示列設定.bas]（行背景色適用）
```

### 9.4 3号機出庫フロー

| # | 起点 | 処理 | 結果 |
|---|---|---|---|
| 1 | 🖥️（SoutiSenntaku で Ric3Keikaku=True → 処理方法実行 Click） | 📄 `初期化()` 内の mpKeiMode=3 分岐 | 📊 「Ric3TR」シート表示 → 🗄️ 3号機データ読込 |
| 2 | 📊（「Ric3TR」上で出庫順列をセルクリック） | 📊 `Worksheet_BeforeDoubleClick()` / **Sheet3.cls** → 📊 `出庫設定()` / **R3Stg1挿入取消.bas** | 📊 `Ric3Work` に出庫順を記録（セル上のみ） |
| 3 | 📊（「予約記録」ボタン Click） | 🗄️ `Ric3出庫記録()` / **R3出庫記録処理.bas** | `ExR3SYukko` を DELETE → INSERT（`syukkojyun`, `senkno`, `siteibi`） |

#### ツリー図（補助）

```
（SoutiSenntaku → Ric3Keikaku=True → 処理方法実行 Click）
└─ 📄 初期化（mpKeiMode=3）       [コマンド実行.bas]
   └─ 📊 「Ric3TR」シート表示
      └─ 🗄️ Ric3号機予想データRead [R3Data読込.bas]
         ├─ 📊 Ric3未照射品計画    [R3計画作成.bas]
         └─ 📊 照射中表示非表示    [R3計画作成.bas]

（「Ric3TR」上で出庫順列をセルクリック）
└─ 📊 Worksheet_BeforeDoubleClick [Sheet3.cls]
   └─ 📊 出庫設定                  [R3Stg1挿入取消.bas]
      └─ 📊 Ric3Work に出庫順を記録（セル上のみ・DB未反映）

（「予約記録」ボタン Click）
└─ 🗄️ Ric3出庫記録               [R3出庫記録処理.bas]
   ├─ ExR3SYukko DELETE（senkno + siteibi）
   └─ ExR3SYukko INSERT（syukkojyun, senkno, siteibi）
```

### 9.5 予想計算フロー

| # | 起点 | 処理 | 結果 |
|---|---|---|---|
| 1 | 📊（「未処理品一覧」→「予想」ボタン Click） | 📄 `RIC2予想計算メイン()` / **R2予想計算メイン.bas** | 📄 2号機照射完了予想計算 → 📊 「受付番号別完了」更新 |
| 2 | 📊（「Ric3TR」→「予想計算」ボタン Click） | 📄 `RIC3予想計算メイン()` / **R3予想計算.bas** | 📄 3号機照射完了予想計算 → 📊 「受付番号別完了」更新 |

### 9.6 その他操作フロー

| 起点 | 処理 | 結果 |
|---|---|---|
| 📄（CommandBar「最新データ」Click） | 📄 `D_UpdatetSort()` → 🗄️ `製品情報ファイル更新()` → 🖥️ `SoutiSenntaku.Show` | 製品情報更新後、装置選択画面を再表示 |
| 📊（CommandBar「予定行の削除」Click） | 🖥️ 資格確認 → 🗄️ `GyouDel()` | 選択行の `ExYoyakuX` を無効化（`yuukou='0'`） → 📊 行削除 |
| 📊（CommandBar「在庫詳細」Click） | 🗄️ `詳細表示()` → 🗄️ `在庫詳細表示()` | DB から詳細取得 → 🖥️ 製品情報.frm 表示 |
| 📄（CommandBar「終了」Click） | 📊 `生産情報終了処理()` | 📊 画面復帰 → ブック閉じ |

---

## 10. セキュリティ注意事項

olevba 解析結果:

| 種別 | キーワード | 内容 |
|---|---|---|
| AutoExec | `Workbook_Open` | ファイルを開くと自動実行 |
| AutoExec | `Workbook_BeforeClose` | ファイルを閉じると自動実行 |
| Suspicious | `Environ` | システム環境変数を読取る可能性 |
| Suspicious | `Open` / `Write` / `Output` | ファイルへの書込み |
| Suspicious | `Kill` | ファイル削除の可能性 |
| Suspicious | `MkDir` | ディレクトリ作成の可能性 |
| Suspicious | `AppActivate` | 他アプリケーション制御の可能性 |
| Suspicious | `Chr` | 文字列難読化の可能性 |
| Suspicious | `Hex Strings` / `Base64 Strings` | エンコード文字列の検出 |
| IOC | `163.59.144.156` | IPv4 アドレス（DB サーバー） |

> **注意**: DB 接続文字列とパスワード（`DSN=ricdb;UID=ric;PWD=t6101`）が VBA に平文で埋め込まれています。外部共有時はマスキング・権限分離を推奨します。

---

## スコープ外（本仕様書に含まないもの）

- セル書式（色・罫線・フォント）
- 条件付き書式、グラフ・画像、印刷設定

必要な場合は Excel 画面のスクリーンショットで補完してください。
