const assert = require("node:assert/strict");
const { test } = require("node:test");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const source = fs.readFileSync(path.join(__dirname, "..", "app.js"), "utf8");

function functionSource(name) {
    const start = source.indexOf("    function " + name + "(");
    assert.notEqual(start, -1);
    const end = source.indexOf("\n    function ", start + 1);
    return source.slice(start, end);
}

function makeDashboard(label, latestMonthKey, rows) {
    const monthKeys = ["2026-01", "2026-02"];
    const skuMonthly = Object.fromEntries(Object.entries(rows).map(([sku, values]) => [
        sku,
        {
            "2026-01": { qty: values[0] || 0 },
            "2026-02": { qty: values[1] || 0 }
        }
    ]));
    const total = Object.values(rows).reduce((sum, values) => sum + values.reduce((itemSum, value) => itemSum + value, 0), 0);
    return {
        channel: { label },
        latestMonthKey,
        yearKeys: { 2026: monthKeys },
        yearlyTotals: { 2026: { qty: total } },
        skuMonthly
    };
}

function harness() {
    const context = vm.createContext({
        SKU_ANNUAL_CHANNEL_ORDER: ["nfm", "abt", "bsm", "rcw"],
        sumMetrics: (rows, months) => ({ qty: months.reduce((sum, month) => sum + (rows[month]?.qty || 0), 0) }),
        isBaseProductSku: sku => /^[A-Z]\d{3}$/i.test(String(sku || "").trim()),
        monthIndex: monthKey => {
            const parts = String(monthKey).split("-").map(Number);
            return parts[0] * 12 + parts[1] - 1;
        }
    });
    ["buildAnnualSkuShareModel", "getAnnualSkuShare"].forEach(name => vm.runInContext(functionSource(name), context));
    const dashboards = {
        nfm: makeDashboard("NFM", "2026-07", { S820: [60, 20], T920: [20, 0], S700: [0, 0], POP: [5, 0], "未归属调整": [-5, 0] }),
        abt: makeDashboard("Abt", "2026-08", { S820: [10, 10], T920: [30, 10] }),
        bsm: makeDashboard("BSM", "2026-08", { S820: [20, 20], E310: [40, 20] }),
        rcw: makeDashboard("RCW", "2026-06", { T920: [10, 10], POPKIT: [5, 5] })
    };
    return { context, dashboards };
}

test("builds four channel rows plus a direct NATM total row", () => {
    const { context, dashboards } = harness();
    const model = context.buildAnnualSkuShareModel(dashboards);
    assert.deepEqual(Array.from(model.rows, row => row.label), ["NFM", "Abt", "BSM", "RCW", "NATM 总量"]);
    assert.equal(model.totalRow.total, 290);
    assert.equal(model.totalRow.latestMonthKey, "2026-08");
    assert.equal(model.totalRow.quantities.S820, 140);
    assert.equal(model.totalRow.quantities.T920, 80);
    assert.equal(model.skus.includes("S700"), false);
});

test("groups POP and adjustment rows into Other without changing the denominator", () => {
    const { context, dashboards } = harness();
    const model = context.buildAnnualSkuShareModel(dashboards);
    const nfm = model.rows[0];
    const rcw = model.rows[3];
    assert.equal(nfm.quantities.Other, 0);
    assert.equal(rcw.quantities.Other, 10);
    assert.equal(nfm.total, 100);
    assert.equal(context.getAnnualSkuShare(nfm, "S820"), 80);
    assert.ok(Math.abs(context.getAnnualSkuShare(rcw, "Other") - 100 / 3) < 1e-9);
});

test("shares retain negative returns and reconcile to 100% for every positive row", () => {
    const { context, dashboards } = harness();
    dashboards.nfm = makeDashboard("NFM", "2026-08", { S820: [110, 0], T920: [0, 0], "未归属调整": [-10, 0] });
    const model = context.buildAnnualSkuShareModel(dashboards);
    model.rows.forEach(row => {
        const sum = model.skus.reduce((total, sku) => total + context.getAnnualSkuShare(row, sku), 0);
        assert.ok(Math.abs(sum - 100) < 1e-9, row.label + " should reconcile to 100%");
    });
    assert.equal(context.getAnnualSkuShare(model.rows[0], "Other"), -10);
});

test("does not calculate a share for a zero or negative channel total", () => {
    const { context } = harness();
    assert.equal(context.getAnnualSkuShare({ total: 0, quantities: { S820: 5 } }, "S820"), null);
    assert.equal(context.getAnnualSkuShare({ total: -10, quantities: { S820: -10 } }, "S820"), null);
});
