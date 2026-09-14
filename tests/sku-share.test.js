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

function harness() {
    const node = () => ({
        style: { setProperty() {} },
        classList: { toggle() {} },
        setAttribute() {},
        querySelectorAll() { return []; }
    });
    const chart = { setOption(option) { this.option = option; }, getOption() { return this.option; }, on(event, handler) { this[event] = handler; } };
    const context = vm.createContext({
        Intl,
        CONFIG: { adjustmentLabel: "Adjustment", colors: { ink: "#172033", muted: "#667085" } },
        SKU_MULTI_TREND_DEFAULT_LIMIT: 6,
        state: { activeChannelKey: "nfm", skuTrendView: "share", selectedSkuTrendByChannel: {}, charts: {} },
        els: Object.fromEntries(["Panel", "Controls", "Chart", "Title", "Description", "Note"].map(key => ["skuMultiTrend" + key, node()])),
        echarts: { init() { return chart; } },
        sumMetrics: (rows, months) => ({ qty: months.reduce((sum, month) => sum + (rows[month]?.qty || 0), 0) })
    });
    ["formatNumber", "escapeHtml", "getTrendableSkus", "getSkuQtyTotal", "getDefaultSkuTrendSelection",
        "getSelectedSkuTrendList", "getSkuMonthlyShare", "renderSkuMultiTrend"].forEach(name => vm.runInContext(functionSource(name), context));
    const dashboard = {
        channel: { key: "nfm", label: "NFM", accent: "#0f766e" },
        baseSkus: ["A", "B", "Adjustment"],
        monthKeys: ["2026-01", "2026-02", "2026-03", "2026-04", "2026-05"],
        skuMonthly: {
            A: { "2026-01": { qty: 60 }, "2026-02": { qty: 0 }, "2026-03": { qty: -2 }, "2026-04": { qty: 120 } },
            B: { "2026-01": { qty: 35 }, "2026-02": { qty: 0 }, "2026-03": { qty: 0 }, "2026-04": { qty: -20 } },
            Adjustment: { "2026-01": { qty: 5 }, "2026-04": { qty: 0 } }
        },
        monthlyTotals: {
            "2026-01": { qty: 100 }, "2026-02": { qty: 0 }, "2026-03": { qty: -2 },
            "2026-04": { qty: 100 }, "2026-05": { qty: 10 }
        }
    };
    context.getActiveDashboard = () => dashboard;
    return { context, dashboard, chart };
}

test("share uses the full monthly total, including adjustments, without rounding inputs", () => {
    const { context: c, dashboard: d } = harness();
    assert.equal(c.getSkuMonthlyShare(d, "A", "2026-01"), 60);
    assert.equal(d.baseSkus.reduce((sum, sku) => sum + c.getSkuMonthlyShare(d, sku, "2026-01"), 0), 100);
    d.monthlyTotals["2026-01"].qty = 300;
    assert.equal(c.getSkuMonthlyShare(d, "B", "2026-01"), 35 / 300 * 100);
});

test("zero or negative monthly totals are undefined; missing SKU rows in a positive month are zero", () => {
    const { context: c, dashboard: d } = harness();
    assert.equal(c.getSkuMonthlyShare(d, "A", "2026-02"), null);
    assert.equal(c.getSkuMonthlyShare(d, "A", "2026-03"), null);
    assert.equal(c.getSkuMonthlyShare(d, "A", "2026-05"), 0);
});

test("returns retain their sign and the share axis does not clip them", () => {
    const { context: c, dashboard: d, chart } = harness();
    assert.equal(c.getSkuMonthlyShare(d, "A", "2026-04"), 120);
    assert.equal(c.getSkuMonthlyShare(d, "B", "2026-04"), -20);
    c.renderSkuMultiTrend(d);
    assert.equal(chart.option.yAxis.min, -20);
    assert.equal(chart.option.yAxis.max, 120);
});

test("deselecting SKUs never changes another SKU's share or color", () => {
    const { context: c, dashboard: d, chart } = harness();
    c.renderSkuMultiTrend(d);
    const before = chart.option.series.find(item => item.name === "B");
    c.state.selectedSkuTrendByChannel.nfm = ["B"];
    c.renderSkuMultiTrend(d);
    assert.deepEqual(chart.option.series[0].data, before.data);
    assert.equal(chart.option.series[0].itemStyle.color, before.itemStyle.color);
    assert.equal(chart.option.series[0].type, "bar");
    assert.equal(chart.option.series[0].stack, "monthly-share");
});

test("switching to quantity keeps selections, raw values and negative returns", () => {
    const { context: c, dashboard: d, chart } = harness();
    c.state.selectedSkuTrendByChannel.nfm = ["B"];
    c.state.skuTrendView = "qty";
    c.renderSkuMultiTrend(d);
    assert.equal(chart.option.series.length, 1);
    assert.equal(chart.option.series[0].type, "line");
    assert.equal(chart.option.series[0].data[0], 35);
    assert.equal(chart.option.series[0].data[3], -20);
    assert.equal(chart.option.yAxis.max, undefined);
});

test("channel selections stay separate, including a deliberately empty selection", () => {
    const { context: c, dashboard: d, chart } = harness();
    c.state.selectedSkuTrendByChannel.nfm = [];
    c.renderSkuMultiTrend(d);
    assert.equal(chart.option.series.length, 0);
    assert.equal(chart.option.title.show, undefined);
    const other = { ...d, channel: { ...d.channel, key: "bsm", label: "BSM" } };
    c.renderSkuMultiTrend(other);
    assert.equal(chart.option.series.length, 2);
    c.renderSkuMultiTrend(d);
    assert.equal(chart.option.series.length, 0);
});

test("legend deselection updates the matching channel selection and chart", () => {
    const { context: c, dashboard: d, chart } = harness();
    c.renderSkuMultiTrend(d);
    chart.legendselectchanged({ selected: { A: false, B: true } });
    assert.equal(c.state.selectedSkuTrendByChannel.nfm.join(","), "B");
    assert.equal(chart.option.series.length, 1);
});

test("tooltip includes the channel, full denominator, raw quantity and 1-decimal share", () => {
    const { context: c, dashboard: d, chart } = harness();
    c.renderSkuMultiTrend(d);
    const tooltip = chart.option.tooltip.formatter([{ axisValue: "2026-01", seriesName: "B", value: 35, marker: "" }]);
    assert.match(tooltip, /NFM/);
    assert.match(tooltip, /100\.0/);
    assert.match(tooltip, /35\.0.*35\.0%/);
    assert.equal(chart.option.series[0].label.formatter({ value: 12.345 }), "12.3%");
    assert.equal(chart.option.series[0].label.formatter({ value: 1.2 }), "");
});

test("changing selections retains a user-adjusted time range", () => {
    const { context: c, dashboard: d, chart } = harness();
    c.renderSkuMultiTrend(d);
    chart.option.dataZoom[0].start = 25;
    chart.option.dataZoom[0].end = 75;
    c.state.selectedSkuTrendByChannel.nfm = ["B"];
    c.renderSkuMultiTrend(d);
    assert.equal(chart.option.dataZoom[0].start, 25);
    assert.equal(chart.option.dataZoom[0].end, 75);
});
