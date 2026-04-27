# Ex生産情報一覧 仕様書

> **ファイル種別**: .xlsm（マクロ付き）
> **用途**: 照射装置（1号機・2号機・3号機・EB）の未照射在庫品の生産計画・出荷予定・照射スケジュールを管理し、DBからデータを取得して一覧表示・計画更新・出庫記録を行う総合管理ツール
> **VBA プロジェクトサイズ**: 約841KB（モジュール36本以上・フォーム6本以上）
> **外部連携**: DSN=ricdb（Oracle DB）、DB接続先 IP: 163.59.144.156

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
Ex生産情報一覧.xlsm
├── シート
│   ├── 未処理品一覧    （未照射在庫品の生産計画一覧・メイン操作画面、最大1000行×73列）
│   ├── 受付番号別完了  （受付番号単位の完了状況一覧、最大1000行×69列）
│   ├── Ric3TR         （3号機トラッキング計画・出庫設定、最大511行×64列）
│   ├── 設定           （DB接続設定・パラメータ・マスタデータ格納、最大1015行×44列）
│   ├── TrakD          （3号機トラック構成・各ステージ時間設定、29行×9列）
│   └── ２号機重量係数  （2号機の重量係数テーブル、52行×5列）
│
├── VBA モジュール（主要なもの）
│   ├── ThisWorkbook.cls          （Workbook_Open/BeforeClose/BeforeSave）
│   ├── SQL_Execution.bas         （ODBC接続・SQL実行共通処理）
│   ├── スタート処理.bas          （起動時初期化・DB読込・画面初期化）
│   ├── データ取得.bas            （ReadDataFromDB：在庫・計画・特殊条件の取得）
│   ├── 更新在庫.bas              （ZaikoKousinn：在庫データのDB更新）
│   ├── 更新予約.bas              （YoyakuKousinn：予約データのDB更新）
│   ├── R2計画作成.bas            （2号機照射計画の計算・生成）
│   ├── R2予想計算メイン.bas       （2号機の照射完了予想計算）
│   ├── R3計画作成.bas            （3号機照射計画の計算・生成）
│   ├── R3出庫記録処理.bas        （3号機出庫順番のDB記録・削除）
│   ├── R3予想計算.bas            （3号機の照射完了予想計算）
│   ├── R3Data読込.bas            （3号機トラッキングデータの読込・表示）
│   ├── R3照射室.bas              （照射室内表示）
│   ├── R3停止期間記録読出.bas     （3号機停止期間の記録・読出）
│   ├── 画面クリア_遷移1.bas      （画面クリア・シート遷移）
│   ├── 画面色と表示列設定.bas     （表示色・列の設定）
│   ├── サブルーチン.bas          （汎用サブルーチン群）
│   ├── ファンクション.bas        （共通関数群）
│   ├── 共通変数.bas              （モジュールレベル変数定義）
│   ├── 生産情報終了.bas          （ブック終了処理）
│   ├── ErrCheck.bas              （エラーチェック・ログ記録）
│   ├── DB不要データ削除.bas       （DBの不要レコード削除）
│   └── フォーム
│       ├── SikakuKakuninn.frm    （資格確認ダイアログ）
│       ├── SyoriSettei.frm       （処理設定ダイアログ）
│       ├── 製品情報.frm          （製品情報表示ダイアログ）
│       ├── Ric3Stg.frm           （3号機ステージ設定ダイアログ）
│       ├── TokuSyuDsp.frm        （特殊条件表示ダイアログ）
│       └── SoutiSenntaku.frm     （装置選択ダイアログ）
│
└── 外部連携
    └── Oracle DB（DSN=ricdb、IP: 163.59.144.156）
```

---

## 2. シート詳細

### 2.1 未処理品一覧

**目的**: メインの操作画面。DBから取得した未照射在庫品の生産計画データを一覧表示し、計画確認・出荷方法・備考の入力・更新を行う。

#### レイアウト構造

| セル範囲 | 名前付き範囲 | 内容 |
|---|---|---|
| A1 | `DebugFlg` | デバッグフラグ（空白=通常動作） |
| E2:F2 | — | ラベル「出荷日」（結合） |
| G2 | — | 出荷日カレンダー用 |
| H2 | `NowToday` | 現在日付 |
| R4:S4 | `SyoriMSG` | 処理メッセージ表示欄（結合） |
| T2 | `Hanuke` | 判定欄 |
| W3 | `Amari` | 残数 |
| X3 | `Ric2HpSuu` | 2号機HPスー |
| X4 | `Ric2HSuu` | 2号機H数 |
| AD2 | `RowNo` | 行番号 |
| AF2 | `ZaikoKensu` | 在庫件数（DBから取得） |
| AH2 | `KeiMaxNo` | 計画最大No |
| AN3 | `HyoujiHouhou` | 表示方法 |
| AO3:AX3 | `SyukkaHou` | 出荷方法一覧 |
| T15 | `SeihinToku` | 製品特殊条件 |
| B16:BU1015 | `Keikaku` | 生産計画データ表示領域（VBAが書き込み） |
| D16:BU1015 | `KeikakuS` | 生産計画データ（D列から） |
| D16:W1015 | `MotoData` | 元データ格納領域 |
| X16:Y1015 | `Ric2Keikaku` | 2号機計画データ |

（行番号 5〜15 は入力操作領域・ヘッダー行）

#### DB読み込み列（C列から、主要な列）

| 列オフセット | DBフィールド | 内容 |
|---|---|---|
| +0 | `uno` | 受付番号 |
| +1 | `nyukabi` | 入荷日 |
| +2 | `kaisyacd` | 会社コード |
| +3 | `kainame` | 会社名 |
| +4 | `siteisn` | 指定線量 |
| +5 | `pass` | パス |
| +6 | `jyougsn` | 上限線量 |
| +7 | `incnt` | 入数 |
| +8 | `nyukasu` | 入荷数 |
| +9 | `misyousu` | 未照射数 |
| +10 | （受付番号下4桁） | 受付番号略称 |
| +11 | `nouki` | 納期 |
| +12 | `kakunin` | 確認フラグ（計画済） |
| +13 | `syukkabi` | 出荷日 |
| +14 | `syuhouhou` | 出荷方法 |
| +15 | `bikou1` | 備考1 |

（列番号 D〜BU（列73まで）が使用される）

---

### 2.2 受付番号別完了

**目的**: 受付番号単位の完了状況（出荷済数・残数・予想完了日など）を表示するシート。

#### レイアウト構造

| セル範囲 | 名前付き範囲 | 内容 |
|---|---|---|
| P2 | `YosouMSG` | 予想メッセージ |
| Q3:T3 | `YosouSiki` | 予想計算式 |
| S2:T2 | — | 結合セル |
| U2 | `Hanuke2` | 判定欄 |
| X5:Y5 | — | 結合セル |
| Z2:AA2 | — | 結合セル |
| AC2:AD2 | — | 結合セル |
| AE2 | `PrintEND` / `PrinｔE` | 印刷終了行 |
| D6:AE1000 | `ZaData` | 完了データ表示領域 |

（行番号 1〜5 はヘッダー・設定行）

---

### 2.3 Ric3TR

**目的**: 3号機トラッキング計画シート。各受付番号がどのステージに配置されるかを管理する。

#### レイアウト構造

| セル範囲 | 名前付き範囲 | 内容 |
|---|---|---|
| AW1 | `WarnMSG` | 警告メッセージ |
| AW3 | `Ric3Teisiji` | 停止指示 |
| AW4 | `Ric3Saikai` | 再開 |
| AW1:BK2 | — | 結合セル（ヘッダー） |
| BG3 | `RIC3StopTime` | 3号機停止時間 |
| BG4 | `Ric3StarTime` | 3号機開始時間 |
| BJ7 | `Ric3MisyoriN` | 未処理No |
| BK4 | `Hanuke3S` | 判定欄S |
| E8 | `Ric3NoMax` | 3号機No最大値 |
| F8 | `Ric3KeiSiki1` | 計算式1 |
| G8 | `Ric3SetteiMax` | 設定最大値 |
| I7 | `Ric3RowS` | 開始行 |
| I8 | `Ric3SyukkoMax` | 出庫最大値 |
| K8 | `Point` | ポイント |
| L8 | `Ric3KeiSiki3` | 計算式3 |
| R8:BG8 | `Ric3KeiSiki2` | 計算式2（範囲） |
| E11:E510 | `Ric3No` | 3号機No一覧 |
| F11:BK510 | `Ric3Work` | 3号機作業領域 |

---

### 2.4 設定

**目的**: DB接続設定・パラメータ・マスタデータ（出荷日テーブル、計画HPテーブル、引取業者コード等）を格納する設定シート。

#### レイアウト構造（主要セル）

| セル範囲 | 名前付き範囲 | 内容 |
|---|---|---|
| A4:F123 | `SyukabiTB` | 出荷日テーブル |
| I4:I133 | `Hatabi` | 発売日一覧 |
| O4 | `RIc2T` | 2号機T設定 |
| O4:P6 | `ParaM` | 2/3号機パラメータ |
| O5 | `Ric2V` | 2号機V設定 |
| O6 | `Ric2HP` | 2号機HP設定 |
| O7 | `Ric2NowSokudo` | 2号機現在速度 |
| O8 | `Ric2NowHP` | 2号機現在HP |
| O13 | `Sort1` | ソート1設定 |
| O14 | `Sort2` | ソート2設定 |
| O18 | `RRic1` | 1号機表示フラグ |
| O19 | `RRic2` | 2号機表示フラグ |
| O20 | `RRic3` | 3号機表示フラグ |
| O21 | `RicEB` | EB表示フラグ |
| N23:N30 | `ComList` | コマンドリスト |
| O31 | `SyoriMorde` | 処理モード |
| O33 | `SyainID` | 社員ID（1,2号機用） |
| O33:O35 | `Update` | 更新設定 |
| O34 | `SyainPW` | 社員PW（1,2号機用） |
| O35 | `AccessTime` | 最終アクセス時刻 |
| O37 | `KeikaTime` | 経過時間閾値（分） |
| O40 | `Sort1Value` | ソート1値 |
| O41 | `Sort2Value` | ソート2値 |
| O43 | `ReqNo` | 要求No（装置選択）|
| O45 | `SenkNo` | 線源No |
| P4 | `Ric3T` | 3号機T設定 |
| P5 | `Ric3M` | 3号機M設定 |
| P6 | `Ric3PP` | 3号機PP設定 |
| P7 | `Ric3NowTime` | 3号機現在時間 |
| P8 | `Ric3NowPP` | 3号機現在PP |
| P13 | `Sort1M` | ソート1モード |
| P14 | `Sort2M` | ソート2モード |
| P33 | `SyainIDR3` | 社員ID（3号機用） |
| P34 | `SyainPWR3` | 社員PW（3号機用） |
| P35 | `AccessTimeR3` | 最終アクセス時刻（3号機用） |
| P36 | `SyainName` | 社員名 |
| P37 | `KeikaTimeR3` | 経過時間閾値（3号機用） |
| P40 | `Sort1Modo` | ソート1モード |
| P41 | `Sort2Modo` | ソート2モード |
| X3 | `Ric3TyuuPP` | 3号機注PP |
| Y6:AB505 | `KeikakuHP` | 計画HPテーブル |
| AD4:AE505 | `HikitoriCD` | 引取業者コードテーブル |
| AH9:AJ12 | `Ric3Track` | 3号機トラック設定 |

---

### 2.5 TrakD

**目的**: 3号機トラック構成情報（各ステージ数・時間設定）を格納する設定シート。

#### レイアウト構造

| セル範囲 | 名前付き範囲 | 内容 |
|---|---|---|
| D5 | `Hang4_5N` / `NowRound` | 現在ラウンド（4〜5段） |
| D9 | `Stg2_3Time` | ステージ2-3時間 |
| D14 | `Stg6_2Time` | ステージ6-2時間 |
| E20 | `KanRtime` | 搬入Rタイム |
| E21 | `SufRtime` | サフRタイム |
| E22 | `MisyTime` | 未処理タイム |
| E23 | `NewRRT` | 新RRT |
| E24 | `Stg6_4Time` | ステージ6-4時間 |
| E25 | `Stg7_4Time` | ステージ7-4時間 |
| E28 | `Tyuusyutu` | 抽出 |
| F16:H16 | `SetteiV` | 設定V |

（行番号 2〜28 は各ステージの結合セルで構成されるトラック構成図）

---

### 2.6 ２号機重量係数

**目的**: 2号機の出荷計画計算に使用する重量係数テーブル（52行×5列）。

---

## 3. 名前付き範囲一覧

| 名前 | 参照先 | 説明 |
|---|---|---|
| `AccessTime` | 設定!$O$35 | 最終アクセス時刻（1,2号機） |
| `AccessTimeR3` | 設定!$P$35 | 最終アクセス時刻（3号機） |
| `Amari` | 未処理品一覧!$W$3 | 残数表示 |
| `ComList` | 設定!$N$23:$N$30 | コマンドリスト |
| `DebugFlg` | 未処理品一覧!$A$1 | デバッグフラグ |
| `Hang4_5N` | TrakD!$D$5 | 3号機現在ラウンド数 |
| `Hanuke` | 未処理品一覧!$T$2 | 判定欄 |
| `Hanuke2` | 受付番号別完了!$U$2 | 受付番号別完了判定欄 |
| `Hanuke3S` | Ric3TR!$BK$4 | 3号機判定欄S |
| `Hatabi` | 設定!$I$4:$I$133 | 発売日一覧 |
| `HikitoriCD` | 設定!$AD$4:$AE$505 | 引取業者コードテーブル |
| `HyoujiHouhou` | 未処理品一覧!$AN$3 | 表示方法 |
| `KanRtime` | TrakD!$E$20 | 搬入Rタイム |
| `Keikaku` | 未処理品一覧!$B$16:$BU$1015 | 生産計画データ表示領域 |
| `KeikakuHP` | 設定!$Y$6:$AB$505 | 計画HPテーブル |
| `KeikakuS` | 未処理品一覧!$D$16:$BU$1015 | 生産計画データ（D列から） |
| `KeikaTime` | 設定!$O$37 | 経過時間閾値（分） |
| `KeikaTimeR3` | 設定!$P$37 | 経過時間閾値（3号機用） |
| `KeiMaxNo` | 未処理品一覧!$AH$2 | 計画最大No |
| `MisyTime` | TrakD!$E$22 | 未処理タイム |
| `MotoData` | 未処理品一覧!$D$16:$W$1015 | 元データ格納領域 |
| `NewRRT` | TrakD!$E$23 | 新RRT |
| `NowRound` | TrakD!$D$5 | 3号機現在ラウンド |
| `NowToday` | 未処理品一覧!$H$2 | 現在日付 |
| `ParaM` | 設定!$O$4:$P$6 | 2/3号機パラメータ |
| `Point` | Ric3TR!$K$8 | ポイント値 |
| `PrintEND` | 受付番号別完了!$AE$2 | 印刷終了行 |
| `ReqNo` | 設定!$O$43 | 要求No（装置選択） |
| `Ric2HP` | 設定!$O$6 | 2号機HP設定値 |
| `Ric2HpSuu` | 未処理品一覧!$X$3 | 2号機HPスー |
| `Ric2HSuu` | 未処理品一覧!$X$4 | 2号機H数 |
| `Ric2Keikaku` | 未処理品一覧!$X$16:$Y$1015 | 2号機計画データ |
| `Ric2NowHP` | 設定!$O$8 | 2号機現在HP |
| `Ric2NowSokudo` | 設定!$O$7 | 2号機現在速度 |
| `RIc2T` | 設定!$O$4 | 2号機T設定 |
| `Ric2V` | 設定!$O$5 | 2号機V設定 |
| `Ric3KeiSiki1` | Ric3TR!$F$8 | 3号機計算式1 |
| `Ric3KeiSiki2` | Ric3TR!$R$8:$BG$8 | 3号機計算式2 |
| `Ric3KeiSiki3` | Ric3TR!$L$8 | 3号機計算式3 |
| `Ric3M` | 設定!$P$5 | 3号機M設定 |
| `Ric3MisyoriN` | Ric3TR!$BJ$7 | 3号機未処理No |
| `Ric3No` | Ric3TR!$E$11:$E$510 | 3号機No一覧 |
| `Ric3NoMax` | Ric3TR!$E$8 | 3号機No最大値 |
| `Ric3NowPP` | 設定!$P$8 | 3号機現在PP |
| `Ric3NowTime` | 設定!$P$7 | 3号機現在時間 |
| `Ric3PP` | 設定!$P$6 | 3号機PP設定 |
| `Ric3RowS` | Ric3TR!$I$7 | 3号機開始行 |
| `Ric3Saikai` | Ric3TR!$AW$4 | 3号機再開フラグ |
| `Ric3SetteiMax` | Ric3TR!$G$8 | 3号機設定最大値 |
| `Ric3StarTime` | Ric3TR!$BG$4 | 3号機開始時間 |
| `RIC3StopTime` | Ric3TR!$BG$3 | 3号機停止時間 |
| `Ric3SyukkoMax` | Ric3TR!$I$8 | 3号機出庫最大値 |
| `Ric3T` | 設定!$P$4 | 3号機T設定 |
| `Ric3Teisiji` | Ric3TR!$AW$3 | 3号機停止指示 |
| `Ric3Track` | 設定!$AH$9:$AJ$12 | 3号機トラック設定 |
| `Ric3TyuuPP` | 設定!$X$3 | 3号機注PP |
| `Ric3Work` | Ric3TR!$F$11:$BK$510 | 3号機作業領域 |
| `RicEB` | 設定!$O$21 | EB表示フラグ |
| `RowNo` | 未処理品一覧!$AD$2 | 行番号 |
| `RRic1` | 設定!$O$18 | 1号機表示フラグ |
| `RRic2` | 設定!$O$19 | 2号機表示フラグ |
| `RRic3` | 設定!$O$20 | 3号機表示フラグ |
| `SeihinToku` | 未処理品一覧!$T$15 | 製品特殊条件 |
| `SenkNo` | 設定!$O$45 | 線源No |
| `SetteiV` | TrakD!$F$16:$H$16 | 設定V |
| `Sort1〜Sort2Modo` | 設定!$O/P$13〜$P$41 | ソート設定（キー・モード・値） |
| `Stg2_3Time` | TrakD!$D$9 | ステージ2-3時間 |
| `Stg6_2Time` | TrakD!$D$14 | ステージ6-2時間 |
| `Stg6_4Time` | TrakD!$E$24 | ステージ6-4時間 |
| `Stg7_4Time` | TrakD!$E$25 | ステージ7-4時間 |
| `SufRtime` | TrakD!$E$21 | サフRタイム |
| `SyainID` | 設定!$O$33 | 社員ID（1,2号機用） |
| `SyainIDR3` | 設定!$P$33 | 社員ID（3号機用） |
| `SyainName` | 設定!$P$36 | 社員名 |
| `SyainPW` | 設定!$O$34 | 社員PW（1,2号機用） |
| `SyainPWR3` | 設定!$P$34 | 社員PW（3号機用） |
| `SyoriMorde` | 設定!$O$31 | 処理モード |
| `SyoriMSG` | 未処理品一覧!$R$4 | 処理メッセージ表示欄 |
| `SyukabiTB` | 設定!$A$4:$F$123 | 出荷日テーブル |
| `SyukkaHou` | 未処理品一覧!$AO$3:$AX$3 | 出荷方法一覧 |
| `Tyuusyutu` | TrakD!$E$28 | 抽出設定 |
| `Update` | 設定!$O$33:$O$35 | 更新設定 |
| `WarnMSG` | Ric3TR!$AW$1 | 3号機警告メッセージ |
| `YosouMSG` | 受付番号別完了!$P$2 | 予想メッセージ |
| `YosouSiki` | 受付番号別完了!$Q$3:$T$3 | 予想計算式 |
| `ZaData` | 受付番号別完了!$D$6:$AE$1000 | 受付番号別完了データ |
| `ZaikoKensu` | 未処理品一覧!$AF$2 | 在庫件数 |

---

## 4. 数式一覧

### ２号機重量係数シート

| セル | 数式 | 説明 |
|---|---|---|
| E3〜E52 | `="Case Is <=" &B3&": R2speedk ="&C3` | VBA Selectケース文生成式（2号機速度係数コード生成） |

### 設定シート（一部）

複数の計算式が含まれるが、主に VBA から書き込まれる参照セルのため省略。

---

## 5. ボタン・マクロ対応

| シート | ボタンラベル | 割り当てマクロ | 動作概要 |
|---|---|---|---|
| 未処理品一覧 | 計画済表示 | `計画済行` | 計画済み行の表示/非表示切替 |
| 未処理品一覧 | 初期化 | `Ric2初期化` | 2号機の計画データを初期化 |
| 未処理品一覧 | フィルタON | `オートフィルター` | オートフィルターの ON/OFF 切替 |
| 未処理品一覧 | 予想 | `RIC2予想計算メイン` | 2号機の照射完了予想を計算 |
| Ric3TR | 停止記録 | `Ric3停止期間記録` | 3号機の停止期間をDBに記録 |
| Ric3TR | 全数指定 | `未指定全数出庫` | 未指定品を全数出庫設定 |
| Ric3TR | 初期化 | `Ric3初期化` | 3号機の計画データを初期化 |
| Ric3TR | 室内表示 | `照射中表示非表示` | 照射中品の表示切替 |
| Ric3TR | 予想計算 | `RIC3予想計算メイン` | 3号機の照射完了予想を計算 |
| Ric3TR | 予約記録 | `Ric3出庫記録` | 3号機出庫順番をDBに記録 |
| Ric3TR | 印刷 | `印刷範囲` | 印刷範囲を設定して印刷 |
| Ric3TR | 照射室表示 | `照射室内表示` | 照射室内状況を表示 |

---

## 6. VBA モジュール仕様

### 6.1 ThisWorkbook

#### `Workbook_Open()`

**処理概要**: ブック起動時に `mpStertF = True` を設定し、`生産情報開始処理` を呼び出す。

#### `Workbook_BeforeClose()`

**処理概要**: デバッグフラグが設定されている場合は終了確認ダイアログを表示する。画面復帰処理を実行後、保存ダイアログを抑止して閉じる。

#### `Workbook_BeforeSave()`

**処理概要**: デバッグフラグが空の場合に `画面CLS` を実行し、コメントを削除してから保存する。

---

### 6.2 SQL_Execution（詳細は Ex装置運転状況 と共通）

`Open_oraconDB`、`SQL_Exe`、`SQL_INSERT_UPDATE`、`SQL_Delete`、`Disp_Sheet`、`Set_Array` の各共通サブルーチンを提供。接続文字列は `DSN=ricdb;UID=ric;PWD=t6101`。

---

### 6.3 スタート処理

#### `生産情報開始処理()`

**処理概要**: 起動時に休日未登録チェック・処理中フォーム表示・コマンドボタン設定・DB不要データ削除・画面初期化を実行し、最後に `SoutiSenntaku`（装置選択ダイアログ）を表示する。

**処理フロー**:
1. 未登録休日チェック（5件未満で警告）
2. `UserForm処理中` を vbModeless で表示
3. `Ric3Stg` フォームをロード・初期化
4. コマンドバーボタンを設定
5. DB不要計画データ削除（`計画削除`）
6. DB不要予約データ削除（`予約削除`）
7. 画面CLS（初期化）
8. ウィンドウ最大化
9. `SoutiSenntaku`（装置選択ダイアログ）を表示

#### `RICDB読込()`

**処理概要**: DB から各種マスタ・在庫・計画・トラッキングデータを読み込み、表示を更新する。

---

### 6.4 データ取得

#### `ReadDataFromDB()`

**処理概要**: `ReqNo`（装置選択）に応じて在庫テーブル（`zaiko`）と計画テーブル（`ExKeikakuX`）を JOIN して取得し、「未処理品一覧」シートの `Keikaku` 領域に書き込む。γ在庫とEB在庫の両方を処理する。

**主要 SQL（γ在庫）**:
```sql
SELECT z.Uno, z.nyukabi, TO_NUMBER(z.kaisyacd), trim(z.kainame),
       z.siteisn*1, z.pass, z.jyougsn*1, z.incnt*1, z.nyukasu*1,
       z.misyousu*1, SUBSTR(z.uno,7,4), trim(z.nouki),
       k.kakunin, k.syukkabi, k.syuhouhou, k.bikou1, '',
       z.nyukasu*1-z.misyousu*1-z.syousu*1-z.syouzusu*1, -- 計画済数
       z.syuhnsu*1, TO_NUMBER(z.syouso), ...
FROM zaiko z, ExKeikakuX k, tsyjmst t, ExSeihinJ j, sejofile s, ExSeihinZ zs
WHERE z.uno=k.uno(+) AND z.kaisyacd=t.kaisyacd(+) AND z.sehncd=t.sehncd(+)
  AND z.kaisyacd=j.kaisyacd(+) AND z.uno=s.uno(+)
  AND z.kaisyacd=zs.kaisyacd(+) AND z.sehncd=zs.sehncd(+)
  AND z.misyousu*1 > 0 AND z.syouso = '2'  -- 例：2号機未照射品
ORDER BY z.kaisyacd, z.uno
```

---

### 6.5 更新在庫 / 更新予約

#### `更新処理()` / `ZaikoKousinn()` / `YoyakuKousinn()`

**処理概要**:
- `更新処理`: 資格確認・予約番号重複チェック後、在庫更新→予約更新→DB読込→表示更新の順で実行する
- `ZaikoKousinn`: 「未処理品一覧」シートの確認・出荷日・出荷方法・備考を `ExKeikakuX` テーブルに INSERT/UPDATE する
- `YoyakuKousinn`: 予約データを `ExYoyakuX` テーブルに INSERT/UPDATE する

**主要 SQL**:
```sql
-- 計画データ更新（INSERT or UPDATE）
SELECT COUNT(*) FROM ExKeikakuX WHERE uno='<受付番号>'
-- 0件: INSERT INTO ExKeikakuX (UNO, KAKUNIN, SYUKKABI, SYUHOUHOU, BIKOU1, ...)
-- 1件: UPDATE ExKeikakuX SET kakunin=..., syukkabi=..., ... WHERE uno='<受付番号>'
```

---

### 6.6 R3出庫記録処理

#### `Ric3出庫記録()`

**処理概要**: Ric3TR シートに設定された出庫順番を `ExR3SYukko` テーブルに記録する（既存データを DELETE 後に INSERT）。

```sql
DELETE ExR3SYukko WHERE syukkojyun > 0
-- その後、各行に対して:
INSERT INTO ExR3SYukko (syukkojyun, senkno, syainid, siteibi)
VALUES ('<出庫順>', '<線量計番号>', '<社員ID>', <設定日>)
```

---

### 6.7 DB不要データ削除

#### `計画削除()` / `予約削除()`

**処理概要**: `ExKeikakuX` テーブルに10000件以上のデータがある場合に `zaiko` に存在しない不要レコードを削除する。`ExYoyakuX` テーブルに200件以上かつ100日以上経過した無効データを削除する。

---

## 7. DB 接続・外部連携

### ODBC 接続設定

| DSN 名 | UID | PWD | 用途 |
|---|---|---|---|
| `ricdb` | `ric` | `t6101` | 全テーブルへのメインアクセス |

> **DB サーバーIP**: 163.59.144.156（olevba により IOC として検出）

### 参照テーブル一覧

| テーブル名 | 主な用途 | 主要カラム |
|---|---|---|
| `zaiko` | γ線照射在庫 | `uno`（受付番号）, `kaisyacd`, `kainame`, `siteisn`, `pass`, `nyukabi`, `nyukasu`, `misyousu`, `syouso`（1/2/3=装置） |
| `zaikor` | 在庫履歴 | `uno`, `kaisyacd`, `syouso` |
| `ExKeikakuX` | 生産計画（可変テーブル名） | `uno`, `kakunin`, `syukkabi`, `syuhouhou`, `bikou1`, `updateid`, `updateday` |
| `ExYoyakuX` | 予約データ | `uno`, `kaisyacd`, `yoyakubi`, `yoyakuno` |
| `ExR3SYukko` | 3号機出庫順番 | `syukkojyun`, `senkno`, `syainid`, `siteibi` |
| `tsyjmst` | 特殊条件マスタ | `kaisyacd`, `sehncd`, `tokjyo1〜5`, `ukedspno`, `syodspno`, `sendspno` |
| `ExSeihinJ` | 製品情報 | `kaisyacd`, `uno`, `housyube` |
| `sejofile` | 生産条件 | `uno`, `nyukajoken`, `shoshagojoken`, `shukajoken`, `tekibin` |
| `ExSeihinZ` | 製品Z情報 | `kaisyacd`, `sehncd`, `tumikae` |
| `syouk2` | 処理情報 | `syoriflg`, `pass` |
| `ExR3Stg` | 3号機ステージ設定 | ステージ管理情報 |

### 外部ファイル連携

ログファイルへの書き込み（`Open/Write/Kill` による）が実装されているが、パスはVBAコード内で動的に生成される。

---

## 8. データフロー

```
【起動フロー】
  Workbook_Open()
        ↓
  生産情報開始処理():
    - 休日チェック
    - DB不要データ削除
    - 画面CLS（クリア）
    - SoutiSenntaku（装置選択ダイアログ）
        ↓ 装置選択後
  RICDB読込():
    DB(zaiko + ExKeikakuX + tsyjmst + ...): データ取得
        ↓
  未処理品一覧シート Keikaku 領域に表示
  Ric3TR シートに3号機トラッキング表示

【データ更新フロー（1/2号機）】
  未処理品一覧: 確認・出荷日・出荷方法・備考 入力
        ↓ メニュー「更新」実行
  資格確認（SikakuKakuninn フォーム）→ ID/PW チェック
        ↓
  ZaikoKousinn(): DB(ExKeikakuX) に INSERT/UPDATE
  YoyakuKousinn(): DB(ExYoyakuX) に INSERT/UPDATE
        ↓
  ReadDataFromDB(): DB から最新データ再取得
  ZaikoDataHyouji(): 表示更新

【3号機出庫記録フロー】
  Ric3TRシート: 出庫順番を入力
        ↓ [予約記録]ボタン → Ric3出庫記録()
  DB(ExR3SYukko): DELETE + INSERT で出庫順を更新

【2号機予想計算フロー】
  未処理品一覧: [予想]ボタン → RIC2予想計算メイン()
  設定シート(Ric2T/V/HP等) をパラメータとして使用
  在庫データ + 速度係数(２号機重量係数シート) → 完了予想日計算
  受付番号別完了シート ZaData に結果表示
```

---

## 9. セキュリティ注意事項

| 種別 | キーワード | 内容 |
|---|---|---|
| AutoExec | `Workbook_Open` | 起動時に `生産情報開始処理` が自動実行される |
| AutoExec | `Workbook_BeforeClose` | 保存ダイアログ抑止・画面復帰処理が自動実行される |
| AutoExec | `ComKakunin_Click` | フォーム上のボタンがアクティブオブジェクトとして自動起動する |
| AutoExec | `Ric2Sokudo_Change` | 速度変更イベントが自動起動する |
| Suspicious | `Environ` | システム環境変数（ユーザー名等）を読み取る可能性がある |
| Suspicious | `Open/Write/Output` | ログファイルや外部ファイルへの書き込みが含まれる |
| Suspicious | `Kill` | ファイルの削除操作が含まれる |
| Suspicious | `MkDir` | ディレクトリを自動作成する処理が含まれる |
| Suspicious | `AppActivate` | 他のアプリケーションをアクティブにする処理が含まれる |
| Suspicious | `Chr` | メッセージ内の改行文字生成（難読化ではない） |
| Suspicious | Hex/Base64 Strings | xlsm 構造上の正常なバイナリデータ |
| **IOC** | `163.59.144.156` | **VBAコード内にDB接続先IPアドレスがハードコーディング** |

> **重大な注意**: DB接続文字列（`DSN=ricdb;UID=ric;PWD=t6101`）がVBAコード内にハードコーディングされており、パスワードが平文で埋め込まれています。また、`163.59.144.156` というIPアドレスが olevba によって IOC（侵害指標）として検出されました。これはDB接続先の内部IPアドレスと考えられますが、ファイルの外部共有には注意が必要です。
