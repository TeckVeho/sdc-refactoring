import type { NextFunction, Request, Response } from "express";

import { isSampleDataMode } from "../lib/sample-data-mode.js";
import { DB_BROWSER_TABLES } from "../services/db-browser-allowlist.js";
import { computeDoseRate, getCoefficientsPayload } from "../services/dose-rate-calc.js";
import { shiftCycleFromYearMonth } from "../services/shift-cycle.js";
import { tempCorrection } from "../services/temp-correction.js";

type SampleMatch = { status: number; body: unknown };

function todayYmd(): string {
  return new Date().toISOString().slice(0, 10);
}

function pathOnly(req: Request): string {
  const raw = (req.originalUrl ?? req.url ?? "").split("?")[0] ?? "";
  if (raw.length > 1 && raw.endsWith("/")) {
    return raw.slice(0, -1);
  }
  return raw;
}

function qs(req: Request): URLSearchParams {
  const q = (req.originalUrl ?? req.url ?? "").split("?")[1];
  return new URLSearchParams(q ?? "");
}

const samplePlanId = "00000000-0000-4000-8000-00000000d1c0";

/** RIC3詰替: GET/POST の本文を共通化 */
function sampleRic3RepackDataInner() {
  return {
    version: 1,
    updatedAt: new Date().toISOString(),
    items: [
      {
        id: "item-1",
        lastRepackDate: "2026/05/01",
        irradiationStatus: "済",
        irradiationCode: "2",
        doseMeterNo: "DM01",
        receiptNo: "202601010001",
        companyName: "サンプル会社",
        dueDate: "05/31",
        shipDate: "20260610",
        passCount: "2",
        note: "",
        checked: false,
      },
    ],
    products: [
      {
        kaisyacd: "0999",
        sehncd: "P001",
        composite: "",
        companyName: "サンプル会社",
        productName: "サンプル樹脂",
        tumikae: "不要",
        originalTumikae: "不要",
      },
    ],
  };
}

function sampleGanttJson() {
  const d0 = todayYmd();
  return {
    data: [
      {
        id: 1,
        text: "製品A（サンプル）",
        start_date: d0,
        duration: 3,
        open: true,
      },
      {
        id: 2,
        text: "製品B（サンプル）",
        start_date: d0,
        duration: 5,
        open: true,
      },
    ],
    links: [{ id: 1, source: 1, target: 2, type: "0" }],
  };
}

const sampleMachineStatus = {
  updatedAt: new Date().toISOString(),
  machines: [
    { id: 1, timer: "123456", statusKey: "1", statusLabel: "照射", detail: "サンプル" },
    { id: 2, timer: "—", statusKey: "0", statusLabel: "停止", detail: "サンプル" },
    { id: 3, timer: "—", statusKey: "0x10", statusLabel: "イベント例", detail: "サンプル" },
  ],
  chart: {
    points: [
      { t: new Date(Date.now() - 3600_000).toISOString(), value: 0, machine: 1 },
      { t: new Date(Date.now() - 1800_000).toISOString(), value: 1, machine: 1 },
      { t: new Date().toISOString(), value: 0.5, machine: 1 },
    ],
  },
};

const sampleMonitor = {
  updatedAt: new Date().toISOString(),
  sourceState: {
    label: "照射中",
    sdate: "20260501",
    stime: "090000",
    event: "1",
    timer: "1000",
  },
  rows: [
    {
      no: 1,
      receiptNo: "202601010001",
      company: "サンプル工業㈱",
      dose: "25",
      qty: 10,
      position: "南ｺﾝ",
      remaining: "120",
      eta: "2026/05/05 17:30",
      shipDate: "2026/05/10",
      note: "デモデータ",
    },
  ],
};

const sampleRic3List = {
  kind: "list" as const,
  shortageCount: 128,
  candidates: [
    { uno: "202601010001", syono: "RIC3-9001", kainame: "サンプル会社" },
    { uno: "202601010002", syono: "RIC3-9002", kainame: "テスト㈱" },
  ],
  truncated: false,
};

function sampleRic3Report(uno: string) {
  return {
    kind: "report" as const,
    meta: {
      formNo: "G3-08",
      title: "３号機線量不足報告書（工程確認・追加照射指示書）",
      variant: "general" as const,
    },
    uno,
    senkNo: `RIC3-${uno.slice(-4)}`,
    doseMeterLast4: "9001",
    syoukj3: {
      syono: "RIC3-9001",
      kainame: "サンプル会社",
      syokind: "1",
      sdate: "20260502",
      edate: "",
      bikou: "サンプル",
    },
    measurement: {
      tuika: null,
      sokDate: "20260502",
      soktCd: null,
      saiSoku: null,
      sokutSn: 22.5,
      tani: "3",
      pass: 2,
      sikiKodo: "1",
      sosi: null,
      atusa: null,
      syokind: "1",
      bikou: "",
      j2Meisai: null,
      j2Status: null,
    },
    plans: [1, 2, 3, 4, 5].map((slot) => ({
      slot,
      ukeNo: slot === 1 ? uno : null,
      ricNoS: slot === 1 ? "RIC3-9001" : null,
      ricNoE: slot === 1 ? "01" : null,
      siteiSn: slot === 1 ? 25 : null,
      kagen: slot === 1 ? 24 : null,
      joug: slot === 1 ? 26 : null,
      jituP: slot === 1 ? 2 : null,
      kaiCd: slot === 1 ? "0001" : null,
      kaiName: slot === 1 ? "サンプル会社" : null,
    })),
    additional: [1, 2, 3, 4, 5].map((slot) => ({
      slot,
      ukeNo: slot === 1 ? uno : null,
      ricNoS: slot === 1 ? "RIC3-9001" : null,
      ricNoE: slot === 1 ? "01" : null,
      shortage: slot === 1 ? 1.5 : null,
      shortageDup: slot === 1 ? 1.5 : null,
      marginFromUpper: slot === 1 ? 3.5 : null,
      extraPasses: slot === 1 ? 1 : null,
      kaiName: slot === 1 ? "サンプル会社" : null,
    })),
    doseMeter: { senKind: "1", atusaT: null, keisask: "1" },
    warnings: { tooManyShortage: false, negativeExtraPasses: false, extraPassesMismatch: false },
    employees: [
      { shano: "dev", shaname: "開発ユーザー" },
      { shano: "s001", shaname: "山田一郎" },
    ],
  };
}

const sampleSumidenPayload = {
  layout: {
    variant: "sumiden" as const,
    reportTitle: "３号機線量不足報告書　住友電工専用（工程確認・追加照射指示書）",
    formCode: "G03-08",
    employeeTableRowCap: 39,
    additionalPassRounding: "sumiden_margin_lt2" as const,
  },
  senkNoOptions: ["SUM-9001", "SUM-9002"],
  sebkNoCount: 0,
  employees: [{ shano: "dev", shaname: "開発ユーザー" }],
  senkindTypes: [{ code: "1", label: "EB" }],
  mokutekiTable: [{ code: "1", description: "一般品" }],
  detail: {
    measurement: {
      tuikaN: "",
      sokDate: "20260501",
      sokutCd: "S1",
      saiSoku: "",
      sokutSn: 20,
      tani: "Gy",
      pass: 1,
      sikiKodo: "1",
      sosi: "",
      atusa: 2.5,
    },
    planRows: [],
  },
};

const sampleShipmentBoardRow = {
  kaisyacd: "0999",
  kairname: "サンプル",
  coname: "サンプル物流㈱",
  hikitoriCurrent: "引取",
  hikitoriDb: "",
  housyubeCurrent: "チャータ便",
  housyubeDb: "",
  reportFlagMax: 1,
  needsReport: true,
};

function sampleShipmentSummary(req: Request): SampleMatch | null {
  const params = qs(req);
  const view = params.get("view");
  const date = params.get("date") ?? "20260501";
  const from = params.get("from") ?? date;
  const to = params.get("to") ?? date;
  const asOf = new Date().toISOString();

  if (view === "arrival") {
    return {
      status: 200,
      body: {
        data: {
          view: "arrival" as const,
          asOf,
          date,
          series: {
            machines: [
              {
                machine: "1" as const,
                label: "1号機",
                rows: [{ kainame: "会社Aサンプル", nyukasu: 120 }],
                totalNyukasu: 120,
              },
              { machine: "2" as const, label: "2号機", rows: [{ kainame: "会社B", nyukasu: 35 }], totalNyukasu: 35 },
              { machine: "3" as const, label: "3号機", rows: [{ kainame: "会社C", nyukasu: 58 }], totalNyukasu: 58 },
            ],
          },
        },
      },
    };
  }

  if (view === "shipment") {
    return {
      status: 200,
      body: {
        data: {
          view: "shipment" as const,
          asOf,
          date,
          series: {
            rows: [
              { kaisyacd: "0999", companyName: "サンプル物流㈱", syukasu: 200 },
              { kaisyacd: "0888", companyName: "テスト商事", syukasu: 150 },
            ],
            totalSyukasu: 350,
          },
        },
      },
    };
  }

  if (view === "uptime") {
    const tp = Date.parse(`${to.slice(0, 4)}-${to.slice(4, 6)}-${to.slice(6, 8)}`);
    const fp = Date.parse(`${from.slice(0, 4)}-${from.slice(4, 6)}-${from.slice(6, 8)}`);
    const days =
      Number.isFinite(tp) && Number.isFinite(fp) && tp >= fp
        ? Math.max(1, Math.floor((tp - fp) / 86400000) + 1)
        : 7;

    return {
      status: 200,
      body: {
        data: {
          view: "uptime" as const,
          asOf,
          from,
          to,
          series: {
            periodDays: days,
            sengnr1: { count: 42, timerSum: 86400, last: { sdate: from, stime: "080000" } },
            kyouj2: { count: 12, lastRectime: `${to} 235959` },
            sengnr3: { count: 38, last: { sdate: to, sekitime: "180000" } },
            simpleUptimeRate: 0.42,
          },
        },
      },
    };
  }

  return null;
}

function sampleCalendarRowsForRange(fromYyyymm: string, toYyyymm: string): { ymd: string; reason: string | null }[] {
  const y1 = Number(fromYyyymm.slice(0, 4));
  const m1 = Number(fromYyyymm.slice(4, 6));
  const y2 = Number(toYyyymm.slice(0, 4));
  const m2 = Number(toYyyymm.slice(4, 6));
  const out: { ymd: string; reason: string | null }[] = [];
  let y = y1;
  let m = m1;
  while (y < y2 || (y === y2 && m <= m2)) {
    const ymd = `${String(y)}${String(m).padStart(2, "0")}03`;
    out.push({ ymd, reason: "サンプル休日" });
    m += 1;
    if (m > 12) {
      m = 1;
      y += 1;
    }
  }
  return out;
}

/** フロントが期待するクエリに対して常に妥当なサンプルを返す */
export function resolveSampleApiResponse(req: Request): SampleMatch | null {
  if (!isSampleDataMode()) return null;

  const method = req.method.toUpperCase();
  const p = pathOnly(req);

  if (p.startsWith("/api/auth/")) return null;

  switch (method) {
    case "GET":
      if (p === "/api/dashboard/notice") {
        return {
          status: 200,
          body: {
            data: {
              text: "【サンプル表示モード】\nSDC_SAMPLE_DATA が有効なため、画面上の一覧・集計はデモ用の固定データです。",
            },
          },
        };
      }
      if (p === "/api/irradiation/machine-status") {
        return { status: 200, body: { data: sampleMachineStatus } };
      }
      if (p === "/api/irradiation/machine1/monitor") {
        return { status: 200, body: { data: sampleMonitor } };
      }
      if (p === "/api/irradiation/machine1/plan") {
        return {
          status: 200,
          body: {
            data: {
              plans: [
                {
                  planId: samplePlanId,
                  json: sampleGanttJson(),
                  details: [],
                },
              ],
            },
          },
        };
      }
      if (p === "/api/irradiation/production") {
        return {
          status: 200,
          body: {
            data: {
              rows: [
                {
                  uno: "202601010001",
                  kaisyacd: "0999",
                  sehncd: "P001",
                  syouso: "1",
                  kainame: "サンプル会社",
                  nyukabi: "20260101000000",
                  nouki: "20260630120000",
                  pass: "2",
                  nyukasu: 100,
                  incnt: 100,
                  siteisn: "25",
                  misyousu: 10,
                  syosuu: 4,
                  senritu: "6.25",
                  syokind: "1",
                  syostat: "1",
                  syukkabi: "20260701",
                  bikou1: "サンプル備考",
                  kakunin: "",
                  syuhouhou: "チャーター",
                  yoyakubi: "",
                  yoyakuno: "",
                  yoyakuBikou: "",
                },
              ],
              meta: { total: 1 },
            },
          },
        };
      }
      if (p === "/api/irradiation/machine1/work-order") {
        const uno = qs(req).get("uno") ?? "202601010001";
        return {
          status: 200,
          body: {
            data: {
              meta: {
                formTitle: "１号機作業指図書（製品仕様書なし）自動印字用【サンプル】",
                specRef: "docs/Ex１号機作業指図書_仕様書.md",
              },
              uno,
              zaiko: {
                uno,
                kaisyacd: "0999",
                sehncd: "P001",
                kainame: "サンプル会社",
                nyukabi: "20260101000000",
              },
              syouk1: {
                uno,
                syono: "S1-001",
                kainame: "サンプル会社",
                siteisn: "25",
                syosuu: 4,
                senritu: 6.25,
                syokind: "1",
                bikou: "",
                sdate: "20260501",
                edate: "",
              },
              derived: {
                kainame: "サンプル会社",
                nyukaBi: "20260101000000",
                siteiSn: 25,
                syosuu: 4,
                senritu: 6.25,
                syono: "S1-001",
                syokind: "1",
                bikou: "",
                sdate: "20260501",
                edate: "",
              },
              employees: [{ shano: "dev", shaname: "開発ユーザー" }],
              ratetbl: [{ ratekey: "R001", rateval: "1.0" }],
            },
          },
        };
      }

      if (p === "/api/dosimetry/dose-rate-calc") {
        return {
          status: 200,
          body: {
            ...getCoefficientsPayload(),
            kanriSources: [
              { sikibetu: "DEMO-KANRI", kousinn: "2026-01-01", ricvm: "1.23", hppp: "0.97" },
            ],
          },
        };
      }
      if (p === "/api/dosimetry/dose-search") {
        return {
          status: 200,
          body: {
            data: {
              truncated: false,
              kansokuMasters: [{ kid: "K001", sokutei: "21.50", bikou: "サンプル" }],
              rows: [
                {
                  uno: "202601010001",
                  kaisyacd: "0999",
                  sehncd: "P001",
                  syouso: "1",
                  kainame: "サンプル会社",
                  nyukabi: "20260101000000",
                  nouki: "",
                  pass: "",
                  source: "zaiko" as const,
                  siteisn: "25",
                  senritu: "6.25",
                  syokind: "1",
                  syostat: "3",
                  syosuu: 4,
                  zhansuu: 0,
                  jno: 1,
                  j2meisai: "",
                  j2status: "",
                },
              ],
            },
          },
        };
      }
      if (p === "/api/dosimetry/irradiation-results") {
        return {
          status: 200,
          body: {
            data: {
              rows: [
                {
                  source: "syouk1" as const,
                  uno: "202601010001",
                  syono: "S001",
                  kainame: "サンプル会社",
                  siteisn: "25",
                  syosuu: 4,
                  syoichi: "",
                  syotime: "",
                  hansuu: 1,
                  stimer: "100",
                  ktimer: "500",
                  senritu: "6.25",
                  syostat: "3",
                  ctimer: "400",
                  zhansuu: 0,
                  htimer: "500",
                  slotno: "01",
                  sdate: "20260502",
                  edate: "20260502",
                  updfflg: "",
                  syokind: "2",
                  bikou: "サンプル",
                },
              ],
              meta: { total: 1, page: 1, pageSize: 50, totalPages: 1 },
            },
          },
        };
      }

      if (p === "/api/dosimetry/oven-management") {
        return {
          status: 200,
          body: {
            data: {
              exOvenTb: [
                { ovenkey: "OVEN-DEMO-1", json: { memo: "サンプルオーブン", ondoSet: "50" } },
              ],
              syouj3ov: [{ ovno: 1, uno: "202601010001", status: "1" }],
            },
          },
        };
      }

      if (p === "/api/dosimetry/oven-management/correction-preview") {
        const q = qs(req);
        const sokutei = Number(q.get("sokutei"));
        const atusa = Number(q.get("atusa"));
        const ondo = Number(q.get("ondo"));
        if (!Number.isFinite(sokutei) || !Number.isFinite(atusa) || !Number.isFinite(ondo)) {
          return { status: 400, body: { error: "sokutei, atusa, ondo は数値で指定してください" } };
        }
        if (atusa === 0) {
          return { status: 400, body: { error: "厚さ(atusa)が0です" } };
        }
        const ratio = sokutei / atusa;
        return {
          status: 200,
          body: {
            data: {
              absPerThickness: ratio,
              correctedKanriPoint: tempCorrection(ratio, ondo),
            },
          },
        };
      }

      if (p === "/api/dosimetry/jmm60" || p === "/api/dosimetry/jmm90") {
        const variant = p.endsWith("jmm60") ? "60" : "90";
        const uno = qs(req).get("uno") ?? "202601010001";
        const title =
          variant === "60"
            ? "1 号機線量測定記録（60φPE 丸棒用）【サンプル】"
            : "1 号機線量測定記録（90φPE 丸棒）【サンプル】";
        return {
          status: 200,
          body: {
            data: {
              variant,
              maxRows: variant === "60" ? 540 : 300,
              title,
              header: {
                uno,
                honnsuu: 10,
                dose: 25,
                sokuKiki: "#01",
                senKind: "EB",
                keicode: "K01",
                kagenn: 24,
                jyogen: 26,
                kainame: "サンプル会社",
                kaisyacd: "0999",
                sehncd: "P001",
                syono: "S001",
                syokind: "1",
                bikou: "デモデータ",
              },
              reference: {
                shipmentMethods: ["引取", "チャータ便"],
                reportKinds: ["照射後報告書発行"],
              },
              senkinds: [{ kindcd: "1", name: "一般" }],
              keicodes: [{ keicode: "K01", keisiki: "式Ⅰ" }],
            },
          },
        };
      }

      if (p === "/api/inventory/ric3-repack") {
        return {
          status: 200,
          body: {
            data: sampleRic3RepackDataInner(),
          },
        };
      }

      if (p === "/api/inventory/shipment-summary") {
        return sampleShipmentSummary(req);
      }

      if (p === "/api/inventory/company-arrival") {
        const kaisyacd = qs(req).get("kaisyacd");
        const date = qs(req).get("date") ?? "20260501";
        if (!kaisyacd) {
          return {
            status: 200,
            body: {
              data: {
                mode: "byCompany",
                date,
                rows: [
                  { kaisyacd: "0999", companyName: "サンプル物流㈱", nyukasu: 500 },
                  { kaisyacd: "0888", companyName: "テスト商事", nyukasu: 120 },
                ],
                totals: { nyukasu: 620, companyCount: 2 },
              },
            },
          };
        }
        return {
          status: 200,
          body: {
            data: {
              mode: "byProduct",
              date,
              kaisyacd,
              companyName: "ダミー社名",
              rows: [{ kaisyacd, sehncd: "P001", productName: "サンプル製品", nyukasu: 80 }],
              totals: { nyukasu: 80, productCount: 1 },
            },
          },
        };
      }

      if (p === "/api/inventory/month-end") {
        const asOf = qs(req).get("asOf") ?? qs(req).get("closingDate") ?? "20260531";
        return {
          status: 200,
          body: {
            data: {
              asOf,
              closingDate: asOf,
              yearMonth: `${asOf.slice(0, 4)}-${asOf.slice(4, 6)}`,
              generatedAt: new Date().toISOString(),
              mvpNote: "サンプル集計データです（SDC_SAMPLE_DATA）。",
              lines: [
                {
                  uno: "202601010001",
                  kaisyacd: "0999",
                  sehncd: "P001",
                  syouso: "1",
                  kainame: "サンプル会社",
                  nyukabi: "20260102000000",
                  nyukasu: 100,
                  pass: "",
                  incnt: 100,
                  nouki: "",
                  source: "zaiko",
                  bucket: "1",
                  bucketLabel: "1号機",
                },
              ],
              summary: [
                { bucket: "1", label: "1号機", rowCount: 1, stockProxyNyukasu: 100, wipAmount: 0 },
                { bucket: "2", label: "2号機", rowCount: 0, stockProxyNyukasu: 0, wipAmount: 0 },
                { bucket: "3", label: "3号機", rowCount: 0, stockProxyNyukasu: 0, wipAmount: 0 },
                { bucket: "EB", label: "EB", rowCount: 0, stockProxyNyukasu: 0, wipAmount: 0 },
                { bucket: "LOCA", label: "LOCA", rowCount: 0, stockProxyNyukasu: 0, wipAmount: 0 },
                { bucket: "other", label: "その他", rowCount: 0, stockProxyNyukasu: 0, wipAmount: 0 },
              ],
              totals: { rowCount: 1, totalNyukasu: 100 },
            },
          },
        };
      }

      if (p === "/api/inventory/customer-stock") {
        return {
          status: 200,
          body: {
            data: {
              asOf: new Date().toISOString(),
              basisDate: qs(req).get("asOf") ?? "20260504",
              kaisyacd: "0999",
              companyName: "サンプル物流㈱",
              note: "サンプル在庫（SDC_SAMPLE_DATA）。",
              rows: [
                {
                  no: 1,
                  uno: "202601010001",
                  nyukabi: "20260201093000",
                  sehncd: "P001",
                  productName: "製品アルファ",
                  kainame: "サンプル会社",
                  pass: "",
                  syouso: "1",
                  nyukasu: 50,
                  computedStock: 50,
                },
              ],
              totals: { rowCount: 1, nyukasu: 50, computedStock: 50 },
              shipmentAfterAsOfSum: 0,
            },
          },
        };
      }

      if (p === "/api/reports/ric2-dose-shortage/data") {
        return {
          status: 200,
          body: {
            data: {
              summary: { syouk2Count: 10, zaikoCount: 200, kansokuCount: 15, syouj2Count: 5 },
              recentPlansWithZaiko: [
                {
                  uno: "202601010001",
                  syono: "J2-01",
                  kainame: "サンプル",
                  syokind: "1",
                  sdate: "20260502",
                  edate: "",
                  bikou: "",
                  zaiko: { kaisyacd: "0999", sehncd: "P001", kainame: "サンプル会社" },
                },
              ],
              kansokuRows: [{ kid: "001", sokutei: "21.35", bikou: "サンプル" }],
              schemaNote: "サンプルデータです。",
            },
          },
        };
      }

      if (p === "/api/reports/ric3-dose-shortage") {
        const uno = qs(req).get("uno")?.trim();
        const senkNo = qs(req).get("senkNo")?.trim();
        if (!uno && !senkNo) {
          return { status: 200, body: { data: sampleRic3List } };
        }
        return { status: 200, body: { data: sampleRic3Report(uno ?? "202601010001") } };
      }

      if (p === "/api/reports/ric3-dose-shortage-sumiden") {
        return { status: 200, body: { data: sampleSumidenPayload } };
      }

      if (p === "/api/reports/shipment-method") {
        return {
          status: 200,
          body: {
            data: {
              rows: [sampleShipmentBoardRow],
              truncated: false,
            },
          },
        };
      }

      if (p === "/api/reports/unissued-search") {
        return {
          status: 200,
          body: {
            data: {
              truncated: false,
              rows: [
                {
                  rowkey: "sample-1",
                  uno: "202601010001",
                  kaisyacd: "0999",
                  kainame: "サンプル会社",
                  sehncd: "P001",
                  syinji: "0",
                  houfax: "",
                  hyouji: "",
                  rotno2: "L1",
                  syukkabi: "20260630",
                  json: {
                    uno: "202601010001",
                    kaisyacd: "0999",
                    syinji: "0",
                    houfax: "",
                    syukkabi: "20260630",
                  },
                },
              ],
            },
          },
        };
      }

      if (p === "/api/admin/db-browser/meta") {
        return { status: 200, body: { data: { tables: DB_BROWSER_TABLES, maxRows: 500 } } };
      }

      if (p === "/api/admin/db-browser") {
        return {
          status: 200,
          body: {
            data: {
              table: qs(req).get("table") ?? "(sample)",
              take: 100,
              skip: 0,
              rowCount: 1,
              rows: [{ uno: "202601010001", kaisyacd: "0999", _sample: true }],
            },
          },
        };
      }

      if (p === "/api/admin/shipment-method-ss") {
        return {
          status: 200,
          body: {
            data: {
              rows: [sampleShipmentBoardRow],
              truncated: false,
              view: "ss" as const,
              ref: "docs/★SS検証_Ex出荷方法報告書発行登録_仕様書.md",
            },
          },
        };
      }

      if (p === "/api/admin/shipment-method-verify") {
        return {
          status: 200,
          body: {
            data: {
              rows: [sampleShipmentBoardRow],
              truncated: false,
              view: "verify" as const,
              ref: "docs/★検証_Ex出荷方法報告書発行登録_仕様書.md",
            },
          },
        };
      }

      if (p === "/api/admin/company-code-convert") {
        return {
          status: 200,
          body: {
            data: {
              rows: [
                {
                  gammacd: "0999",
                  ebcd: "E999",
                  memo: "サンプル対応関係",
                },
              ],
            },
          },
        };
      }

      if (p === "/api/admin/price-search") {
        return {
          status: 200,
          body: {
            data: {
              uno: "202601010001",
              kaisyacd: "0999",
              sehncd: "P001",
              companyName: "サンプル物流㈱",
              companyShort: "サンプル",
              productName: "製品アルファ",
              zaiko: { uno: "202601010001", kaisyacd: "0999", sehncd: "P001" },
              priceRowkey: "price:0999:P001",
              price: {
                tanka: "120000",
                tanni: "式",
                tourokubi: "20260201",
                dose: "25",
                souti: "1号機",
                pass: "",
                folder: "",
                file: "",
              },
              note: "サンプル単価（SDC_SAMPLE_DATA）。",
            },
          },
        };
      }

      if (p === "/api/admin/shift-schedule") {
        const year = Number(qs(req).get("year") ?? "2026");
        const month = Number(qs(req).get("month") ?? "5");
        const cycle = Number.isFinite(year) && Number.isFinite(month) ? shiftCycleFromYearMonth(year, month) : shiftCycleFromYearMonth(2026, 5);
        return {
          status: 200,
          body: {
            data: {
              cycle,
              scheduleId: null,
              label: null,
              symbols: [{ symbol: "1", memo: "1号機勤務", hours: null as string | null }],
              circulationList: [] as { id: string; name: string; sortOrder: number | null }[],
              employees: [{ shano: "dev", shaname: "開発ユーザー", inRoster: true as const }],
              grid: {},
              rosterSize: 1,
            },
          },
        };
      }

      if (p === "/api/admin/calendar") {
        const qp = qs(req);
        const yearQ = qp.get("year");
        if (yearQ && /^\d{4}$/.test(yearQ)) {
          const fromYm = `${yearQ}01`;
          const toYm = `${yearQ}12`;
          return {
            status: 200,
            body: { data: sampleCalendarRowsForRange(fromYm, toYm) },
          };
        }
        const fromParam = qp.get("from");
        const toParam = qp.get("to");
        if (fromParam && toParam && /^\d{6}$/.test(fromParam) && /^\d{6}$/.test(toParam)) {
          return {
            status: 200,
            body: { data: sampleCalendarRowsForRange(fromParam, toParam) },
          };
        }
        const y = new Date().getFullYear();
        const cur = `${y}${String(new Date().getMonth() + 1).padStart(2, "0")}`;
        return {
          status: 200,
          body: { data: sampleCalendarRowsForRange(cur, cur) },
        };
      }

      break;

    case "POST": {
      if (p === "/api/reports/gamma-results/aggregate") {
        return {
          status: 200,
          body: {
            data: {
              params: {
                year: 2026,
                month: 5,
                unoMin: "20260101000001",
                unoMax: "20260131999999",
                kkmd1: "05/01",
                kkmd2: "05/31",
              },
              goukei: {
                ric1: { inboundQty: 950, simpleProcessingApprox: 123.456 },
                ric2: { inboundQty: 480, simpleProcessingApprox: 120 },
                ric3: { inboundQty: 610, simpleProcessingApprox: 88 },
              },
              notes: [
                "【サンプル】SDC_SAMPLE_DATA が有効です。",
                "実データの集計は無効時にご利用ください。",
              ],
            },
          },
        };
      }

      if (p === "/api/dosimetry/dose-rate-calc") {
        const fixed = computeDoseRate({
          mode: "RATE",
          potentiometerId: 1,
          ionChamberId: 1,
          readValue: 12.345,
          temperatureC: 22,
          pressureHpa: 1013.25,
          tpEnabled: true,
          uncertaintyOrg: "JQA",
          cableOn: true,
          targetDoseGy: 25,
          decay: null,
        });
        return { status: 200, body: { ok: true, result: fixed } };
      }

      if (p === "/api/admin/db-browser/query") {
        return {
          status: 200,
          body: {
            data: {
              table: (req.body as { table?: string })?.table ?? "zaiko",
              take: 100,
              skip: 0,
              rowCount: 1,
              rows: [{ message: "sample row", uno: "202601010001" }],
            },
          },
        };
      }

      if (p === "/api/irradiation/machine1/plan") {
        return {
          status: 201,
          body: {
            data: {
              planId: samplePlanId,
              json: sampleGanttJson(),
            },
          },
        };
      }

      if (p === "/api/admin/calendar") {
        const b = req.body as { ymd?: string };
        const ymd = typeof b?.ymd === "string" && /^\d{8}$/.test(b.ymd) ? b.ymd : "20261231";
        return {
          status: 201,
          body: {
            data: { ymd, reason: "サンプル登録（SDC_SAMPLE_DATA は DB に保存されませんでしたが成功を返しています）" },
          },
        };
      }

      if (p === "/api/admin/company-code-convert") {
        return { status: 200, body: { ok: true } };
      }

      if (p.startsWith("/api/inventory/ric3-repack/actions/")) {
        return {
          status: 200,
          body: {
            data: sampleRic3RepackDataInner(),
          },
        };
      }

      if (p === "/api/inventory/ric3-repack/products" || p === "/api/inventory/ric3-repack/items") {
        return {
          status: 200,
          body: {
            data: sampleRic3RepackDataInner(),
          },
        };
      }

      if (p === "/api/dosimetry/oven-management") {
        const b = req.body as { ovenkey?: string; json?: Record<string, unknown> };
        return {
          status: 201,
          body: {
            data: {
              ovenkey: b?.ovenkey ?? "OVEN-DEMO",
              json: b?.json ?? { note: "サンプル" },
            },
          },
        };
      }

      break;
    }

    case "PATCH": {
      if (p.startsWith("/api/reports/unissued-search/rows/")) {
        return { status: 200, body: { ok: true } };
      }

      const calPatch = /^\/api\/admin\/calendar\/(\d{8})$/.exec(p);
      if (calPatch) {
        return {
          status: 200,
          body: { data: { ymd: calPatch[1]!, reason: "サンプル更新" } },
        };
      }

      if (/^\/api\/dosimetry\/oven-management\/.+$/.test(p) && !p.includes("correction-preview")) {
        const ovenkey = p.split("/").pop()!;
        const b = req.body as { json?: Record<string, unknown> };
        return {
          status: 200,
          body: {
            data: {
              ovenkey,
              json: { ...(b?.json ?? {}), patched: true, sampleMode: true },
            },
          },
        };
      }

      if (p === "/api/inventory/ric3-repack/products" || p === "/api/inventory/ric3-repack/items") {
        return { status: 200, body: { data: sampleRic3RepackDataInner() } };
      }

      break;
    }

    case "PUT": {
      if (p === "/api/reports/shipment-method") {
        return { status: 200, body: { ok: true, updated: 0 } };
      }

      const planPut = p.match(/^\/api\/irradiation\/machine1\/plan\/([^/]+)$/);
      if (planPut) {
        const bodyJson = (req.body as { json?: unknown })?.json ?? sampleGanttJson();
        return {
          status: 200,
          body: { data: { planId: planPut[1]!, json: bodyJson } },
        };
      }

      if (p === "/api/admin/price-search") {
        const b = req.body as { kaisyacd?: string; sehncd?: string };
        const kaisyacd = b?.kaisyacd ?? "0999";
        const sehncd = b?.sehncd ?? "P001";
        return { status: 200, body: { ok: true, rowkey: `price:${kaisyacd}:${sehncd}` } };
      }

      break;
    }

    case "DELETE": {
      if (p === "/api/admin/company-code-convert") {
        return { status: 200, body: { ok: true } };
      }

      const calDel = /^\/api\/admin\/calendar\/(\d{8})$/.exec(p);
      if (calDel) {
        return { status: 200, body: { data: { ok: true as const, ymd: calDel[1]! } } };
      }

      break;
    }

    default:
      break;
  }

  if (method === "POST" && p === "/api/admin/shift-schedule") {
    return { status: 200, body: { ok: true, id: "sample-shift-id", updated: true } };
  }

  return null;
}

export function sampleDataApiMiddleware(req: Request, res: Response, next: NextFunction): void {
  if (!isSampleDataMode()) {
    next();
    return;
  }
  try {
    const hit = resolveSampleApiResponse(req);
    if (!hit) {
      next();
      return;
    }
    res.status(hit.status).json(hit.body);
  } catch {
    next();
  }
}
