# SDC 照射管理システム Web再構築 - 実装計画書

> **プロジェクト名**: SDC（照射管理システム Web版）
> **元システム**: Excel VBA (.xlsm) × 30本 + Oracle DB + Access DB
> **再構築先**: Next.js + Express + MySQL + Prisma
> **作成日**: 2026-04-27

---

## 目次

1. [プロジェクト概要](#1-プロジェクト概要)
2. [アーキテクチャ](#2-アーキテクチャ)
3. [技術スタック](#3-技術スタック)
4. [DBスキーマ設計](#4-dbスキーマ設計)
5. [画面一覧](#5-画面一覧)
6. [フェーズ別実装計画](#6-フェーズ別実装計画)
7. [実装ルール](#7-実装ルール)
8. [subagent プロンプトテンプレート](#8-subagent-プロンプトテンプレート)
9. [検証・完了基準](#9-検証完了基準)

---

## 1. プロジェクト概要

### 1.1 ビジネスドメイン

ラジエ工業の放射線照射管理システム。ガンマ線照射装置（1号機・2号機・3号機）およびEB（電子線）装置を用いた製品照射処理の全業務を管理する。

### 1.2 業務領域

| 領域 | 概要 |
|---|---|
| 照射管理 | 1〜3号機の照射計画・照射情報モニタリング・作業指図書 |
| 入出荷管理 | 入荷・出荷・在庫の集計・報告 |
| 線量管理 | 線量測定・温度補正・線量率計算・オーブン管理 |
| 報告書 | 線量不足報告書・出荷方法報告書・未発行検索 |
| 設備モニタリング | 1〜3号機の稼働時間・イベント・トレンドグラフ |
| 管理・マスタ | 休日カレンダー・勤務シフト表・単価・会社コード変換 |
| 汎用ツール | DBブラウザ・検索ツール |

### 1.3 移行方針

- 全30アプリを30画面として完全移行
- Oracle DB → MySQL への全面移行（並行運用なし・クリーンカットオーバー）
- Access DB → MySQL へ統合
- Excel VBA のビジネスロジックは `docs/` 配下の仕様書に基づいて新規実装

### 1.4 ユーザー規模

- 同時接続: 2人
- 利用部門: 照射課、入出荷、品証、営業

---

## 2. アーキテクチャ

```
┌──────────────────────────────────────────────┐
│  フロントエンド (Next.js App Router)            │
│  ポート: 3000                                   │
│  - shadcn/ui + Tailwind CSS                     │
│  - TanStack Query (サーバー状態管理)             │
│  - TanStack Table (テーブル)                     │
│  - Recharts (グラフ)                             │
│  - dhtmlx-gantt (ガントチャート)                  │
│  - @react-pdf/renderer (PDF生成)                 │
│  - SheetJS/xlsx (Excelエクスポート)               │
├──────────────────────────────────────────────┤
│  バックエンド (Express + TypeScript)              │
│  ポート: 4000                                    │
│  - Prisma ORM (MySQL)                            │
│  - passport + express-session (認証)              │
│  - Zod (バリデーション)                           │
│  - SSE (Server-Sent Events)                      │
├──────────────────────────────────────────────┤
│  データベース (MySQL 8.0)                         │
│  ポート: 3306                                    │
│  - 照射管理テーブル群 (Oracle移行分)               │
│  - 勤務管理テーブル群 (Access移行分)               │
│  - Web固有テーブル (認証・セッション・係数)         │
├──────────────────────────────────────────────┤
│  インフラ                                         │
│  - Docker Compose (開発・本番共通)                 │
│  - nginx (リバースプロキシ)                        │
│  - ローカルファイルストレージ                       │
└──────────────────────────────────────────────┘
```

### 2.1 通信フロー

```
ブラウザ → nginx(:80) → Next.js(:3000)  [画面配信]
                       → Express(:4000)  [API /api/*]
                       → Express(:4000)  [SSE /sse/*]
```

### 2.2 認証フロー

```
POST /api/auth/login { employeeId, password }
  → shainmst テーブルで照合
  → セッション発行 (express-session + MySQL store → `web_sessions` テーブル)
  → ロール付与 (users テーブルの role カラム)
  → JWT不使用（セッションベース）
```

---

## 3. 技術スタック

| 層 | 技術 | バージョン |
|---|---|---|
| フロントエンド | Next.js (App Router) | 15.x |
| UIライブラリ | shadcn/ui + Tailwind CSS | latest |
| 状態管理 | TanStack Query | v5 |
| テーブル | TanStack Table | v8 |
| グラフ | Recharts | latest |
| ガントチャート | dhtmlx-gantt | latest |
| PDF | @react-pdf/renderer | latest |
| Excel出力 | SheetJS (xlsx) | latest |
| バックエンド | Express + TypeScript | 4.x |
| ORM | Prisma | 6.x |
| DB | MySQL | 8.0 |
| 認証 | passport + express-session | latest |
| バリデーション | Zod | latest |
| ビルド | Turborepo | latest |
| コンテナ | Docker Compose | v2 |
| テスト | Vitest + Playwright | latest |
| 言語 | TypeScript | 5.x |

---

## 4. DBスキーマ設計

### 4.1 テーブル分類

#### マスタ系（9テーブル）

| テーブル名 | 用途 | 主要カラム |
|---|---|---|
| `tokumst` | 得意先マスタ | kaisyacd(PK), coname, kairname |
| `sehmst` | 製品マスタ | kaisyacd+sehncd(PK), seiname, syouho |
| `shainmst` | 社員マスタ | shano(PK), shaname, shask, hshika, cshika |
| `senkind` | 線量計種類 | 種類コード(PK), 名称 |
| `keicode` | 計算式コード | コード(PK), 係数群 |
| `kcdcnvmst` | GM/EB会社コード変換 | ガンマCD+EBCD(PK) |
| `tsyjmst` | 注文マスタ | 得意先関連 |
| `users` | **新規** Web認証 | id(PK), employee_id(FK→shainmst), role, password_hash |
| `correction_coefficients` | **新規** 補正係数 | id(PK), name, value, effective_from |

#### トランザクション系（15テーブル）

| テーブル名 | 用途 | 備考 |
|---|---|---|
| `zaiko` | 在庫（現在） | — |
| `zaikor` | 在庫（履歴） | `r` = 履歴(rireki) |
| `syouk1` | 1号機照射データ | — |
| `syouk2` | 2号機照射データ | — |
| `syoukj3` | 3号機照射データ | `j` は Oracle 元テーブル名に由来。`syouk3` ではない点に注意 |
| `syouj1` | 1号機照射詳細（現在） | — |
| `syoujr1` | 1号機照射詳細（履歴） | — |
| `syouj2` | 2/3号機照射詳細（現在） | **2号機・3号機共用**。テーブル名の「2」は2号専用ではない |
| `syoujr2` | 2/3号機照射詳細（履歴） | 同上（履歴版） |
| `syouj3ov` | 3号機オーブン照射実績 | `ov` = オーブン |
| `syukar` | 出荷記録 | — |
| `kansoku` | 線量測定データ | — |
| `ratetbl` | 線量率テーブル | Oracle 元名を維持。`tbl` = table（冗長表記） |
| `sejofile` | 製条ファイル | — |
| `web_sessions` | **新規** セッション | express-session 用。`sessions` だと store の自動スキーマと衝突しやすいため明示名 |

#### 装置センサー系（3テーブル）

| テーブル名 | 用途 | 状態判定方法 | 備考 |
|---|---|---|---|
| `sengnr1` | 1号機イベントログ | event コード値 (0/1/2/3/E/S) | 1/3 号は `sengnr` 系 |
| `kyouj2` | 2号機稼働ログ | rectime 5分間隔判定 | 2号機のみプレフィックスが異なる（Oracle 元名に由来） |
| `sengnr3` | 3号機イベントログ | event 16ビットデコード | — |

#### 照射管理拡張系（10テーブル）

> **命名補足**: `ex_` = 拡張(Extension)。サフィックス例: `_x` = 拡張系データ、`_tb` = パラメータ/マスタ形式の区別、末尾の数字（`ex_r1_keikaku1` / `2`）= ヘッダー/明細。`j` / `z`（`seihinj` / `seihinz`）は従来システムの区別。`stg` = staging。

| テーブル名 | 用途 | 備考 |
|---|---|---|
| `ex_seihinj` | 製品出荷方法・報告書・価格表 | 情報系。`j` = 従来区分（seihin**j**） |
| `ex_seihinz` | 製品単価 | 単価系。`z` = 従来区分（seihin**z**） |
| `ex_keikaku_x` | 出荷計画 | `_x` = 拡張系 |
| `ex_yoyaku_x` | 予約 | `_x` = 拡張系 |
| `ex_kanri_tb` | 線源管理パラメータ | `_tb` = 従来のテーブル区分。冗長表記の可能性あり |
| `ex_oven_tb` | オーブン管理 | 同上（`_tb`） |
| `ex_yasumi_x` | 休祭日 | `_x` = 拡張系 |
| `ex_r1_keikaku1` | 1号機計画ヘッダー | `1` = ヘッダー相当 |
| `ex_r1_keikaku2` | 1号機計画詳細 | `2` = 明細相当 |
| `ex_r3_stg` | 3号機ステージング | `stg` = staging |

#### 勤務管理系（旧Access → MySQL、4テーブル）

| テーブル名 | 用途 |
|---|---|
| `shift_schedules` | 勤務予定表 |
| `shift_employees` | ガンマ入出荷勤務社員 |
| `work_times` | 勤務時間マスタ |
| `circulation_list` | 回覧先 |

### 4.2 Oracle → MySQL SQL変換ルール

| Oracle | MySQL | 備考 |
|---|---|---|
| `TO_NUMBER(x)` | `CAST(x AS SIGNED)` | Prisma経由では不要な場合が多い |
| `TO_DATE(x, 'yyyy/mm/dd hh24:mi:ss')` | `STR_TO_DATE(x, '%Y/%m/%d %H:%i:%s')` | |
| `NVL(a, b)` | `IFNULL(a, b)` | |
| `SUBSTR(x, n, m)` | `SUBSTRING(x, n, m)` | |
| `t1.col = t2.col(+)` | `LEFT JOIN t2 ON t1.col = t2.col` | 全て ANSI JOIN に変換 |
| `ROWNUM <= N` | `LIMIT N` | |
| `\|\|` (連結) | `CONCAT(a, b)` | |
| `SYSDATE` | `NOW()` | |
| `TRIM(x)` | `TRIM(x)` | 互換 |
| `ROW_NUMBER() OVER(...)` | `ROW_NUMBER() OVER(...)` | MySQL 8.0 互換 |

---

## 5. 画面一覧

### 5.1 全画面マップ（30画面）

| # | 画面パス | 画面名 | 元アプリ | 種別 | リアルタイム |
|---|---|---|---|---|---|
| **ダッシュボード** | | | | | |
| 1 | `/dashboard` | ホームメニュー | ExMenu | ナビ | — |
| **照射管理** | | | | | |
| 2 | `/irradiation/machine1/monitor` | 1号機照射情報 | Ex1号機照射情報 | 参照 | SSE 30秒 |
| 3 | `/irradiation/machine1/plan` | 1号機照射計画 | Ex1号機照射計画 | CRUD | ポーリング 5分 |
| 4 | `/irradiation/machine1/work-order` | 1号機作業指図書 | Ex１号機作業指図書 | CRUD+印刷 | — |
| 5 | `/irradiation/machine-status` | 装置運転状況 | Ex装置運転状況 | 参照+グラフ | SSE 1分 |
| 6 | `/irradiation/production` | 生産情報一覧 | Ex生産情報一覧 | CRUD | — |
| **入出荷** | | | | | |
| 7 | `/inventory/shipment-summary` | 入出荷集計 | Ex入出荷集計 | 検索+集計 | — |
| 8 | `/inventory/company-arrival` | 会社別入荷集計 | Ex会社別入荷集計 | 検索+集計 | — |
| 9 | `/inventory/month-end` | 月末在庫集計 | Ex月末在庫集計 | 集計+出力 | — |
| 10 | `/inventory/customer-stock` | 顧客在庫報告 | Ex顧客在庫報告 | 集計+出力 | — |
| 11 | `/inventory/ric3-repack` | Ric3詰替作業 | ExRic3詰替作業 | CRUD | — |
| **線量管理** | | | | | |
| 12 | `/dosimetry/oven-management` | R3オーブン管理 | ExR3ｵｰﾌﾞﾝ管理温度補正有 | CRUD+計算 | — |
| 13 | `/dosimetry/dose-search` | 線量検索 | Ex線量検索 | 検索+出力 | — |
| 14 | `/dosimetry/irradiation-results` | 照射実績表示 | Ex照射実績表示2 | 検索 | — |
| 15 | `/dosimetry/jmm60` | JMM60φ記入用紙 | ExJMM60φ記入用紙 | 入力+印刷 | — |
| 16 | `/dosimetry/jmm90` | JMM90φ記入用紙 | ExJMM90φ記入用紙 | 入力+印刷 | — |
| 17 | `/dosimetry/dose-rate-calc` | 線量率計算 | 線量率計算v1.1 | 計算ツール | — |
| **報告書** | | | | | |
| 18 | `/reports/ric2-dose-shortage` | RIC2線量不足報告書 | ExRIC2線量不足報告書 | 生成+印刷 | — |
| 19 | `/reports/ric3-dose-shortage` | RIC3線量不足報告書 | ExRIC3線量不足報告書 | 生成+印刷 | — |
| 20 | `/reports/ric3-dose-shortage-sumiden` | RIC3線量不足報告書(住電) | ExRIC3線量不足報告書住電用 | 生成+印刷 | — |
| 21 | `/reports/shipment-method` | 出荷方法報告書発行登録 | Ex出荷方法報告書発行登録 | CRUD | — |
| 22 | `/reports/unissued-search` | 報告書未発行検索 | Ex報告書未発行検索 | 検索+CRUD | — |
| 23 | `/reports/gamma-results` | ガンマ照射課実績集計 | Exガンマ照射課実績集計 | 集計 | — |
| **管理・マスタ** | | | | | |
| 24 | `/admin/calendar` | カレンダー | Exカレンダー | CRUD | — |
| 25 | `/admin/shift-schedule` | 勤務表 | Exガンマ勤務表 | CRUD+印刷 | — |
| 26 | `/admin/price-search` | 単価検索 | Ex単価検索 | 検索+CRUD | — |
| 27 | `/admin/db-browser` | DBブラウザ | ExDBファイル表示 | 汎用検索 | — |
| 28 | `/admin/company-code-convert` | GM/EB会社コード変換 | ExGM_EB会社ｺｰﾄﾞ変換TB | CRUD | — |
| 29 | `/admin/shipment-method-verify` | 出荷方法報告書(検証) | ★検証_Ex出荷方法報告書 | CRUD | — |
| 30 | `/admin/shipment-method-ss` | 出荷方法報告書(SS検証) | ★SS検証_Ex出荷方法報告書 | 参照 | — |

### 5.2 共通コンポーネント

| コンポーネント | 用途 | 使用画面 |
|---|---|---|
| `<AppShell>` | サイドバー + ヘッダー + コンテンツ領域 | 全画面 |
| `<DataTable>` | ソート・フィルタ・ページネーション付きテーブル | ほぼ全画面 |
| `<DateRangePicker>` | 期間指定 | 集計系全般 |
| `<PrintButton>` | PDF生成 + ブラウザ印刷 | 報告書・帳票系 |
| `<ExcelExportButton>` | Excelファイルダウンロード | 集計系全般 |
| `<MachineStatusBadge>` | 装置状態表示（照射中/停止等） | #2, #5 |
| `<GanttChart>` | ガントチャート | #3 |
| `<TrendChart>` | トレンドグラフ（Recharts） | #5, #7 |

---

## 6. フェーズ別実装計画

### Phase 1: 基盤構築（Week 1-2）

**目的**: プロジェクト骨格・認証・共通基盤の構築

| タスクID | タスク名 | 変更対象 | 依存 | 並列グループ |
|---|---|---|---|---|
| 1-1 | monorepo 初期化 (Turborepo + Docker Compose) | `package.json`, `turbo.json`, `docker-compose.yml` | — | — (順列) |
| 1-2 | Prisma スキーマ定義 + MySQL マイグレーション | `packages/backend/prisma/schema.prisma` | 1-1 | — (順列) |
| 1-3 | Express サーバー骨格 + 共通ミドルウェア | `packages/backend/src/**` | 1-1 | A |
| 1-4 | Next.js 初期化 + shadcn/ui セットアップ + AppShell | `packages/frontend/**` | 1-1 | A |
| 1-5 | 認証 API (passport + session) | `packages/backend/src/routes/auth.ts`, `middleware/auth.ts` | 1-3 | B |
| 1-6 | 認証 UI (ログイン画面) + RBAC フロント | `packages/frontend/app/(auth)/**` | 1-4, 1-5 | B |
| 1-7 | SSE 基盤 | `packages/backend/src/sse/**` | 1-3 | B |
| 1-8 | 共通計算サービス (放射線減衰, 温度補正) | `packages/backend/src/services/radiation-decay.ts`, `temp-correction.ts` | 1-2 | A |

**実行順序**:

```
1-1 → 1-2 → [A: 1-3, 1-4, 1-8 並列] → [B: 1-5, 1-6, 1-7 並列]
```

**完了条件**:

- [ ] `docker compose up` でフロント・バック・DBが起動する
- [ ] ログイン→セッション発行→保護APIアクセスが動作する
- [ ] Prisma マイグレーションで全テーブルが作成される
- [ ] SSE エンドポイントからイベントが配信される
- [ ] 放射線減衰計算・温度補正計算のユニットテストが通る

---

### Phase 2: コア画面（Week 3-4）

**目的**: 日常業務で最も使用頻度の高い画面を実装

| タスクID | タスク名 | 画面# | 参照仕様書 | 並列グループ |
|---|---|---|---|---|
| 2-1 | ホームメニュー | #1 | `ExMenu_仕様書.md` | A |
| 2-2 | 1号機照射情報 (SSE) | #2 | `Ex1号機照射情報_仕様書.md` | A |
| 2-3 | 入出荷集計 (3シート統合) | #7 | `Ex入出荷集計_仕様書.md` | A |
| 2-4 | カレンダー (休日管理) | #24 | `Exカレンダー_仕様書.md` | A |
| 2-5 | 1号機照射計画 (ガントチャート) | #3 | `Ex1号機照射計画_仕様書.md` | B |
| 2-6 | 装置運転状況 (SSE + グラフ) | #5 | `Ex装置運転状況_仕様書.md` | B |
| 2-7 | 生産情報一覧 | #6 | `Ex生産情報一覧_仕様書.md` | B |
| 2-8 | 報告書未発行検索 | #22 | `Ex報告書未発行検索_仕様書.md` | B |

**実行順序**:

```
[A: 2-1, 2-2, 2-3, 2-4 並列] → [B: 2-5, 2-6, 2-7, 2-8 並列]
```

**完了条件**:

- [ ] 全8画面でCRUD or 表示が正常動作する
- [ ] SSE による1号機・装置運転状況のリアルタイム更新が動作する
- [ ] ガントチャートでの照射計画操作（配置・中断・空・削除）が動作する
- [ ] TypeScript 型エラー・ESLint エラーがない

---

### Phase 3: 入出荷・線量系（Week 5-6）

**目的**: 入出荷業務と線量管理の画面群を実装

| タスクID | タスク名 | 画面# | 参照仕様書 | 並列グループ |
|---|---|---|---|---|
| 3-1 | 会社別入荷集計 | #8 | `Ex会社別入荷集計_仕様書.md` | C |
| 3-2 | 月末在庫集計 | #9 | `Ex月末在庫集計_仕様書.md` | C |
| 3-3 | 顧客在庫報告 | #10 | `Ex顧客在庫報告_仕様書.md` | C |
| 3-4 | Ric3詰替作業 | #11 | `ExRic3詰替作業_仕様書.md` | C |
| 3-5 | R3オーブン管理 (温度補正) | #12 | `ExR3ｵｰﾌﾞﾝ管理温度補正有_仕様書.md` | D |
| 3-6 | 線量検索 | #13 | `Ex線量検索_仕様書.md` | D |
| 3-7 | 照射実績表示 | #14 | `Ex照射実績表示2_仕様書.md` | D |
| 3-8 | 線量率計算 | #17 | `線量率計算v1.1_仕様書.md` | D |

**実行順序**:

```
[C: 3-1, 3-2, 3-3, 3-4 並列] → [D: 3-5, 3-6, 3-7, 3-8 並列]
```

**完了条件**:

- [ ] 温度補正計算 `(-3.2478 + 23.386 * ABS/厚さ) * (-0.006 * 温度 + 1.0558)` が正確に動作
- [ ] 放射線減衰計算 `0.5^(経過日数/半減期日数)` が `ex_kanri_tb` の値で動作
- [ ] 月末在庫集計・顧客在庫報告のExcel出力が動作する

---

### Phase 4: 報告書・帳票系（Week 7-8）

**目的**: 帳票・報告書の生成・印刷機能を実装

| タスクID | タスク名 | 画面# | 参照仕様書 | 並列グループ |
|---|---|---|---|---|
| 4-1 | RIC2線量不足報告書 | #18 | `ExRIC2線量不足報告書_仕様書.md` | E |
| 4-2 | RIC3線量不足報告書 | #19 | `ExRIC3線量不足報告書_仕様書.md` | E |
| 4-3 | RIC3線量不足報告書 (住電用) | #20 | `ExRIC3線量不足報告書住電用_仕様書.md` | E |
| 4-4 | 出荷方法報告書発行登録 | #21 | `Ex出荷方法報告書発行登録_仕様書.md` | F |
| 4-5 | JMM60φ/90φ記入用紙 | #15, #16 | `ExJMM60φ記入用紙_仕様書.md`, `ExJMM90φ記入用紙_仕様書.md` | F |
| 4-6 | 1号機作業指図書 | #4 | `Ex１号機作業指図書_仕様書.md` | F |
| 4-7 | ガンマ照射課実績集計 | #23 | `Exガンマ照射課実績集計_仕様書.md` | F |

**実行順序**:

```
[E: 4-1, 4-2, 4-3 並列] → [F: 4-4, 4-5, 4-6, 4-7 並列]
```

**完了条件**:

- [ ] 全報告書がPDF生成 + ブラウザ印刷で出力できる
- [ ] 住電用フォーマットが通常版と正しく分岐する

---

### Phase 5: 管理・マスタ系（Week 9-10）

**目的**: 管理機能・マスタメンテナンス画面の実装

| タスクID | タスク名 | 画面# | 参照仕様書 | 並列グループ |
|---|---|---|---|---|
| 5-1 | 勤務表 | #25 | `Exガンマ勤務表_仕様書.md` | G |
| 5-2 | 単価検索 | #26 | `Ex単価検索_仕様書.md` | G |
| 5-3 | DBブラウザ | #27 | `ExDBファイル表示_仕様書.md` | G |
| 5-4 | GM/EB会社コード変換 | #28 | `ExGM_EB会社ｺｰﾄﾞ変換TB_仕様書.md` | H |
| 5-5 | 検証版画面 (2画面) | #29, #30 | `★検証_*_仕様書.md`, `★SS検証_*_仕様書.md` | H |

**実行順序**:

```
[G: 5-1, 5-2, 5-3 並列] → [H: 5-4, 5-5 並列]
```

**完了条件**:

- [ ] 勤務表の11日始まりサイクルが正しく動作する
- [ ] DBブラウザで任意テーブルの検索・条件絞り込み・Excel出力が動作する
- [ ] 勤務記号のコピー&ペースト操作がWebで再現されている

---

### Phase 6: 統合テスト・リリース（Week 11-12）

| タスクID | タスク名 | 依存 |
|---|---|---|
| 6-1 | E2Eテスト (Playwright) | Phase 1-5 全完了 |
| 6-2 | 全画面の機能網羅性チェック | 6-1 |
| 6-3 | パフォーマンステスト | 6-1 |
| 6-4 | 本番デプロイ・環境構築 | 6-2 |

---

## 7. 実装ルール

### 7.1 ディレクトリ構成

```
sdc-web/
├── packages/
│   ├── frontend/                # Next.js App Router
│   │   ├── app/
│   │   │   ├── (auth)/          # ログイン
│   │   │   │   └── login/page.tsx
│   │   │   ├── layout.tsx       # AppShell
│   │   │   ├── dashboard/       # ホーム
│   │   │   ├── irradiation/     # 照射管理
│   │   │   ├── inventory/       # 入出荷
│   │   │   ├── dosimetry/       # 線量管理
│   │   │   ├── reports/         # 報告書
│   │   │   └── admin/           # 管理
│   │   ├── components/
│   │   │   ├── ui/              # shadcn/ui
│   │   │   └── shared/          # 業務共通
│   │   └── lib/
│   │       ├── api-client.ts    # fetch ラッパー
│   │       └── utils.ts
│   │
│   ├── backend/                 # Express
│   │   ├── src/
│   │   │   ├── index.ts         # エントリポイント
│   │   │   ├── routes/          # ルート定義
│   │   │   │   ├── auth.ts
│   │   │   │   ├── irradiation.ts
│   │   │   │   ├── inventory.ts
│   │   │   │   ├── dosimetry.ts
│   │   │   │   ├── reports.ts
│   │   │   │   └── admin.ts
│   │   │   ├── services/        # ビジネスロジック
│   │   │   │   ├── radiation-decay.ts
│   │   │   │   ├── temp-correction.ts
│   │   │   │   └── machine-status/
│   │   │   │       ├── strategy.ts
│   │   │   │       ├── machine1.ts
│   │   │   │       ├── machine2.ts
│   │   │   │       └── machine3.ts
│   │   │   ├── middleware/
│   │   │   │   ├── auth.ts
│   │   │   │   ├── rbac.ts
│   │   │   │   └── error-handler.ts
│   │   │   └── sse/
│   │   │       └── machine-monitor.ts
│   │   ├── prisma/
│   │   │   └── schema.prisma
│   │   └── .env
│   │
│   └── shared/                  # 共通型定義
│       └── types/
│           ├── auth.ts
│           ├── inventory.ts
│           ├── irradiation.ts
│           └── dosimetry.ts
│
├── docs/                        # 仕様書群（既存）
├── docker-compose.yml
├── package.json
└── turbo.json
```

### 7.2 コーディング規約

- **言語**: TypeScript strict mode
- **命名**: camelCase (変数/関数), PascalCase (型/コンポーネント), snake_case (DBカラム)
- **Prisma モデルと DB 物理名**:
  - Oracle/Access 移行テーブル（短縮名・従来名）: 物理テーブル名は維持し、Prisma モデル名は英語の PascalCase とする。`@@map("物理テーブル名")` で対応付ける。

    ```prisma
    model IrradiationData1 {
      // ...
      @@map("syouk1")
    }
    // 2号機・3号機の照射詳細は同一テーブル
    model IrradiationDetail2 {
      // ...
      @@map("syouj2")
    }
    model Machine2OperationLog {
      // ...
      @@map("kyouj2")
    }
    ```

  - 新規テーブル（`users`, `web_sessions`, `correction_coefficients`, `shift_schedules`, `shift_employees`, `work_times`, `circulation_list`）: 物理名をそのままモデル名の snake 対応形にしてよく、`@@map` は原則不要（モデル名 `WorkTime` → `work_times` 等）。
- **API レスポンス形式**: `{ data: T, error?: string, meta?: { total: number } }`
- **エラーハンドリング**: Express の集中エラーハンドラー + Zod バリデーションエラー
- **コメント**: ビジネスロジックの意図のみ記載。自明なコードにはコメント不要
- **セッションストア**: `express-mysql-session` 等の設定で、セッションテーブル名を `web_sessions` に**明示的に**指定し、自動作成に任せる場合もスキーマをドキュメントと揃える

### 7.3 docs/ 仕様書の参照ルール

- 各画面の実装時は **必ず対応する仕様書を読んでから** 実装すること
- 仕様書の「VBA モジュール仕様」セクションのビジネスロジックを新規実装の基準とする
- 仕様書の「DB 接続・外部連携」セクションの SQL を MySQL 構文に変換して使用する
- 仕様書の「データフロー」セクションを画面遷移・API設計の根拠とする
- VBAコードをそのままコピーしない。ロジックの意図を理解して TypeScript で再実装する

### 7.4 装置状態判定の実装ルール

3台の装置はそれぞれ異なるロジックで状態判定する。Strategy パターンで実装する。

| 装置 | テーブル | 判定方法 | 参照仕様書セクション |
|---|---|---|---|
| 1号機 | `sengnr1` | event コード値判定 | Ex装置運転状況 §6.3 |
| 2号機 | `kyouj2` | 記録間隔 5分判定 | Ex装置運転状況 §6.3 |
| 3号機 | `sengnr3` | 16ビットデコード | Ex装置運転状況 §6.3, ExR3ｵｰﾌﾞﾝ管理 §6.4 |

### 7.5 RBAC（権限ロール）

| ロール | 権限 |
|---|---|
| `admin` | 全機能＋マスタ管理＋ユーザー管理 |
| `manager` | 照射計画の作成・承認、勤務表の作成 |
| `operator` | 日常業務（照射情報参照・入出荷操作） |
| `viewer` | 参照のみ |

---

## 8. subagent プロンプトテンプレート

各 implementor subagent に渡すプロンプトの標準テンプレート:

```markdown
以下のタスクを実装してください。

## タスク
[タスクID]: [タスク名]
[詳細説明]

## 参照ドキュメント（必ず読んでから実装すること）
- docs/[該当仕様書].md → 全セクション
- docs/implementation-plan.md → セクション「7. 実装ルール」

## 変更対象ファイル（参考）
- packages/backend/src/routes/[module].ts
- packages/backend/src/services/[service].ts
- packages/frontend/app/[path]/page.tsx
- packages/frontend/app/[path]/components/

## 技術スタック
- フロントエンド: Next.js 15 (App Router) + shadcn/ui + Tailwind CSS
- バックエンド: Express + TypeScript
- ORM: Prisma (MySQL)
- 状態管理: TanStack Query v5
- テーブル: TanStack Table v8
- グラフ: Recharts
- バリデーション: Zod

## Oracle→MySQL SQL変換ルール
- TO_NUMBER(x) → CAST(x AS SIGNED)
- SUBSTR(x,n,m) → SUBSTRING(x,n,m)
- t1.col=t2.col(+) → LEFT JOIN t2 ON t1.col=t2.col
- ROWNUM <= N → LIMIT N
- || → CONCAT()
- NVL(a,b) → IFNULL(a,b)
- TO_DATE(x,'fmt') → STR_TO_DATE(x,'%Y/%m/%d %H:%i:%s')

## 完了条件
1. 実装が完了すること
2. TypeScript の型エラーがないこと
3. ESLint エラーがないこと
4. [画面固有の条件]

実装が完了したら以下を報告してください:
- 変更したファイルの一覧
- スコープ外への影響がないことの確認
- 実装のポイント・判断した点
- 未解決の課題（あれば）
```

### 8.1 並列実行ルール（rebuild-system-execute 準拠）

- 並列グループ内は最大 4 タスクまで同時委任
- 次グループは前グループの全タスク完了後に開始
- 同一ファイル（`package.json`, ルーティング, Prisma schema）を触るタスクは順列

---

## 9. 検証・完了基準

### 9.1 フェーズ完了検証

各フェーズの全タスク完了後に実施:

```bash
# 型チェック
cd packages/backend && npx tsc --noEmit
cd packages/frontend && npx tsc --noEmit

# リント
npm run lint

# テスト
npm run test
```

### 9.2 フェーズ完了報告テンプレート

```markdown
## Phase [N] 完了報告

### 実装結果
| タスク | 状態 | 変更ファイル数 |
|--------|------|--------------|

### 統合テスト結果
- 型チェック: ✅ / ❌
- リント: ✅ / ❌
- テスト: ✅ / ❌ / 未実装

### ドキュメント更新
- [更新した仕様書があれば記載]

### 未解決課題
- [あれば記載]

### 次フェーズの準備状況
- [前提条件が満たされているか]
```

### 9.3 全体完了基準

```markdown
## 再構築完了報告

### サマリ
- 全フェーズ数: 6
- 総タスク数: 38
- 総画面数: 30

### 画面網羅率
- 実装済み: 30 / 30 (100%)

### API動作確認
- 全エンドポイントの正常系レスポンス確認

### 機能網羅性
| 画面 | CRUD | 検索 | 集計 | 印刷/出力 | SSE | 判定 |
|------|------|------|------|----------|-----|------|
```

### 9.4 次フェーズ開始判定

**自動開始の条件**（すべて満たす場合）:

1. 型チェック・リント通過
2. 未解決課題がない
3. ドキュメント更新完了

**ユーザー確認が必要な条件**:

- テストに失敗がある
- スコープ外の変更が必要と判明した
- ビジネスロジックの解釈に曖昧さがある

---

## 付録: rebuild-system-execute との対応

本計画書は `/Users/kohei/Projects/Meyasubako/.cursor/commands/rebuild-system-execute/SKILL.md` の Step 7〜10 に対応する。

| スキル | 本ドキュメントの該当 |
|---|---|
| Step 7 タスク委任 | §6 フェーズ別実装計画、§8 プロンプトテンプレート |
| Step 8 検証 | §9.1, §9.2 |
| Step 9 次フェーズ | §9.4 |
| Step 10 最終検証 | §9.3 |

実装開始時は `docs/` の仕様書一式と本書を前提とし、プロジェクト基盤（`apps/web`, `apps/api`, `packages/database` 等）の初期化後に Phase 1 から着手する。
