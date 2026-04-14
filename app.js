(function() {
    const WORKBOOK_ID = "1-3zBJZvum5M7yLED_ZUbpMugmkWZeTCFZVAvs1F4QqE";

    const CONFIG = {
        workbookId: WORKBOOK_ID,
        workbookUrl: "https://docs.google.com/spreadsheets/d/" + WORKBOOK_ID + "/edit?gid=1543835087#gid=1543835087",
        focusYears: [2024, 2025, 2026],
        startYear: 2024,
        structureStartMonth: "2025-09",
        adjustmentLabel: "未归属调整",
        channels: [
            { key: "nfm", label: "NFM", sheet: "NFM  Sales Data", accent: "#0f766e" },
            { key: "rcw", label: "RCW", sheet: "RCW Sales Data", accent: "#b45309" },
            { key: "abt", label: "Abt", sheet: "Abt Sales Data", accent: "#1d4ed8" },
            { key: "bsm", label: "BSM", sheet: "BSM Sales Data", accent: "#b42318" }
        ],
        channelProfiles: {
            nfm: {
                displayName: "NFM",
                businessTags: ["SO订单", "RMA退货", "EDI: TBD", "其他支持"],
                buyer: "anna.murphy@nfm.com / joe.feregrino@nfm.com (senior)",
                account: "jennifer.priebe@nfm.com (SO下单) / Detriee.Powe@bm1.brandsmart.com (invoice)",
                shipPrice2025: "DFI后为 0.6*0.98*0.95",
                shipPrice2026: "更正后 0.6*(1-0.05-0.02)",
                setup: "沿用当前流程，无额外新品 setup 备注",
                popSize: "120cm",
                popSkus: ["E310", "T920", "S803", "S710", "S820", "S821"]
            },
            rcw: {
                displayName: "RCW",
                businessTags: ["其他支持"],
                buyer: "david.mcloney@rcwilley.com",
                account: "/",
                shipPrice2025: "0.6",
                shipPrice2026: "0.6",
                setup: "联系 rep 及 buyer",
                popSize: "4*30cm",
                popSkus: ["S820", "S821", "T920", "E310", "S710"]
            },
            abt: {
                displayName: "Abt",
                businessTags: ["SO订单", "RMA退货", "其他支持"],
                buyer: "brozycki@abt.com",
                account: "soraya.matus@abt.com (RMA退货)",
                shipPrice2025: "0.6",
                shipPrice2026: "0.6",
                setup: "给 buyer 发送 Roadmap 即可",
                popSize: "暂无",
                popSkus: []
            },
            bsm: {
                displayName: "BSM",
                businessTags: ["SO订单", "RMA退货", "其他支持"],
                buyer: "gerard.sarvis@bm1.brandsmart.com",
                account: "Susan.Bailey@bm1.brandsmart.com (RMA退货)",
                shipPrice2025: "0.6",
                shipPrice2026: "0.6",
                setup: "给 buyer 发送 Roadmap 即可",
                popSize: "120cm",
                popSkus: ["E310", "T920", "S803", "S710", "S820", "S821"]
            }
        },
        colors: {
            qty: "#0f766e",
            sales: "#1d4ed8",
            yoy: "#b45309",
            negative: "#b42318",
            muted: "#667085",
            ink: "#172033"
        }
    };

    const state = {
        dashboards: {},
        activeChannelKey: "nfm",
        selectedMetric: "qty",
        compareMetric: "sales",
        selectedSkuByChannel: {},
        charts: {}
    };

    const els = {
        eyebrow: document.querySelector(".eyebrow"),
        heroTitle: document.querySelector(".hero h1"),
        heroCopy: document.querySelector(".hero-copy"),
        footer: document.querySelector(".footer"),
        businessOverviewGrid: document.getElementById("businessOverviewGrid"),
        channelOverviewGrid: document.getElementById("channelOverviewGrid"),
        channelSummaryBody: document.querySelector("#channelSummaryTable tbody"),
        channelTabs: document.getElementById("channelTabs"),
        kpiGrid: document.getElementById("kpiGrid"),
        insightPanel: document.getElementById("insightPanel"),
        yearSummaryGrid: document.getElementById("yearSummaryGrid"),
        monthlySummaryBody: document.querySelector("#monthlySummaryTable tbody"),
        skuMatrixTable: document.getElementById("skuMatrixTable"),
        skuSelect: document.getElementById("skuSelect"),
        skuMetricSelect: document.getElementById("skuMetricSelect"),
        skuDetailBody: document.querySelector("#skuDetailTable tbody"),
        skuSummaryBody: document.querySelector("#skuSummaryTable tbody"),
        refreshBtn: document.getElementById("refreshBtn"),
        statusBadge: document.getElementById("statusBadge"),
        lastUpdated: document.getElementById("lastUpdated"),
        dataRangeLabel: document.getElementById("dataRangeLabel")
    };

    function csvUrlForSheet(sheetName, cacheKey) {
        const key = cacheKey || Date.now();
        return "https://docs.google.com/spreadsheets/d/" + CONFIG.workbookId + "/gviz/tq?tqx=out:csv&sheet=" + encodeURIComponent(sheetName) + "&_=" + encodeURIComponent(String(key));
    }

    function parseNumber(value) {
        if (value === null || value === undefined) {
            return 0;
        }
        const cleaned = String(value).trim().replace(/,/g, "").replace(/\$/g, "");
        if (!cleaned) {
            return 0;
        }
        const parsed = Number(cleaned);
        return Number.isFinite(parsed) ? parsed : 0;
    }

    function formatNumber(value, digits) {
        return new Intl.NumberFormat("en-US", {
            maximumFractionDigits: digits || 0,
            minimumFractionDigits: digits || 0
        }).format(value);
    }

    function formatCurrency(value, digits) {
        return new Intl.NumberFormat("en-US", {
            style: "currency",
            currency: "USD",
            maximumFractionDigits: digits || 0,
            minimumFractionDigits: digits || 0
        }).format(value);
    }

    function formatPercent(value, absolute) {
        if (value === null || value === undefined || !Number.isFinite(value)) {
            return "—";
        }
        const pct = (absolute ? Math.abs(value) : value) * 100;
        const prefix = absolute ? "" : (pct > 0 ? "+" : "");
        return prefix + pct.toFixed(1) + "%";
    }

    function growthClass(value) {
        if (value === null || value === undefined || !Number.isFinite(value)) {
            return "neutral";
        }
        if (value > 0) {
            return "positive";
        }
        if (value < 0) {
            return "negative";
        }
        return "neutral";
    }

    function renderDelta(value) {
        return "<span class=\"delta " + growthClass(value) + "\">" + formatPercent(value, false) + "</span>";
    }

    function calculateGrowth(current, previous) {
        if (!Number.isFinite(previous) || previous === 0) {
            return null;
        }
        return (current - previous) / Math.abs(previous);
    }

    function describeGrowth(value, positiveWord, negativeWord) {
        if (value === null || value === undefined || !Number.isFinite(value)) {
            return "暂无可比同期";
        }
        if (Math.abs(value) < 0.0005) {
            return "与同期基本持平";
        }
        if (value > 0) {
            return positiveWord + " " + formatPercent(value, true);
        }
        return negativeWord + " " + formatPercent(value, true);
    }

    function hexToRgba(hex, alpha) {
        const normalized = String(hex || "").replace("#", "");
        if (normalized.length !== 6) {
            return "rgba(15, 118, 110, " + alpha + ")";
        }
        const red = Number.parseInt(normalized.slice(0, 2), 16);
        const green = Number.parseInt(normalized.slice(2, 4), 16);
        const blue = Number.parseInt(normalized.slice(4, 6), 16);
        return "rgba(" + red + ", " + green + ", " + blue + ", " + alpha + ")";
    }

    function toMonthKey(year, month) {
        return String(year) + "-" + String(month).padStart(2, "0");
    }

    function monthIndex(monthKey) {
        const parts = String(monthKey).split("-").map(Number);
        return parts[0] * 12 + parts[1] - 1;
    }

    function previousMonthKey(monthKey) {
        const index = monthIndex(monthKey) - 1;
        const year = Math.floor(index / 12);
        const month = index % 12 + 1;
        return toMonthKey(year, month);
    }

    function sameMonthLastYearKey(monthKey) {
        const parts = String(monthKey).split("-").map(Number);
        return toMonthKey(parts[0] - 1, parts[1]);
    }

    function baseSku(rawSku) {
        const value = String(rawSku || "").trim();
        if (!value) {
            return CONFIG.adjustmentLabel;
        }
        return value.split("-")[0] || CONFIG.adjustmentLabel;
    }

    function buildMonthRange(records) {
        if (!records.length) {
            return [];
        }
        const latestIndex = Math.max.apply(null, records.map(record => monthIndex(record.monthKey)));
        const startIndex = CONFIG.startYear * 12;
        const range = [];
        for (let index = startIndex; index <= latestIndex; index += 1) {
            const year = Math.floor(index / 12);
            const month = index % 12 + 1;
            if (CONFIG.focusYears.includes(year)) {
                range.push(toMonthKey(year, month));
            }
        }
        return range;
    }

    function blankMetricMap(monthKeys) {
        const map = {};
        monthKeys.forEach(monthKey => {
            map[monthKey] = { qty: 0, sales: 0 };
        });
        return map;
    }

    function sumMetrics(source, monthKeys) {
        return (monthKeys || []).reduce((acc, monthKey) => {
            const row = (source && source[monthKey]) || { qty: 0, sales: 0 };
            acc.qty += row.qty;
            acc.sales += row.sales;
            return acc;
        }, { qty: 0, sales: 0 });
    }

    function escapeHtml(value) {
        return String(value || "").replace(/[&<>"']/g, function(char) {
            return {
                "&": "&amp;",
                "<": "&lt;",
                ">": "&gt;",
                "\"": "&quot;",
                "'": "&#39;"
            }[char];
        });
    }

    function monthKeysSince(monthKeys, startMonthKey) {
        const startIndex = monthIndex(startMonthKey);
        return (monthKeys || []).filter(monthKey => monthIndex(monthKey) >= startIndex);
    }

    function buildProductMix(skuMonthly, monthKeys, startMonthKey) {
        const relevantKeys = monthKeysSince(monthKeys, startMonthKey);
        const items = Object.keys(skuMonthly).map(sku => {
            const totals = sumMetrics(skuMonthly[sku], relevantKeys);
            return {
                sku: sku,
                qty: totals.qty,
                sales: totals.sales
            };
        }).filter(item => {
            return item.sku !== CONFIG.adjustmentLabel &&
                /^[A-Z]\d{3}$/i.test(item.sku) &&
                (item.qty > 0 || item.sales > 0);
        }).sort((left, right) => {
            if (right.sales !== left.sales) {
                return right.sales - left.sales;
            }
            if (right.qty !== left.qty) {
                return right.qty - left.qty;
            }
            return left.sku.localeCompare(right.sku);
        });

        const totalSales = items.reduce((sum, item) => sum + Math.max(item.sales, 0), 0);
        const totalQty = items.reduce((sum, item) => sum + Math.max(item.qty, 0), 0);

        items.forEach(item => {
            item.salesShare = totalSales > 0 ? item.sales / totalSales : null;
            item.qtyShare = totalQty > 0 ? item.qty / totalQty : null;
        });

        return {
            startMonthKey: startMonthKey,
            monthKeys: relevantKeys,
            items: items,
            totalSales: totalSales,
            totalQty: totalQty
        };
    }

    function normalizeRecord(row) {
        const year = Number.parseInt(String(row.Year || "").trim(), 10);
        const month = Number.parseInt(String(row.Month || "").trim(), 10);
        if (!CONFIG.focusYears.includes(year) || !Number.isInteger(month) || month < 1 || month > 12) {
            return null;
        }

        const rawSku = String(row.SKU || "").trim();
        const qty = parseNumber(row["Total Quantity Sold"]);
        const sales = parseNumber(row["Total Sales"]);
        if (!rawSku && qty === 0 && sales === 0) {
            return null;
        }

        return {
            year: year,
            month: month,
            monthKey: toMonthKey(year, month),
            rawSku: rawSku,
            sku: baseSku(rawSku),
            qty: qty,
            sales: sales
        };
    }

    function sortSkus(left, right, skuMonthly, monthKeys) {
        if (left === CONFIG.adjustmentLabel) {
            return 1;
        }
        if (right === CONFIG.adjustmentLabel) {
            return -1;
        }

        const recentKeys = monthKeys.filter(monthKey => monthKey.startsWith("2026-"));
        const priorKeys = monthKeys.filter(monthKey => monthKey.startsWith("2025-"));
        const leftRecentSales = sumMetrics(skuMonthly[left], recentKeys).sales;
        const rightRecentSales = sumMetrics(skuMonthly[right], recentKeys).sales;
        if (rightRecentSales !== leftRecentSales) {
            return rightRecentSales - leftRecentSales;
        }

        const leftPriorSales = sumMetrics(skuMonthly[left], priorKeys).sales;
        const rightPriorSales = sumMetrics(skuMonthly[right], priorKeys).sales;
        if (rightPriorSales !== leftPriorSales) {
            return rightPriorSales - leftPriorSales;
        }

        return left.localeCompare(right);
    }

    function findTopSkuForMonths(skuMonthly, monthKeys) {
        let winner = null;
        Object.keys(skuMonthly).forEach(sku => {
            const totals = sumMetrics(skuMonthly[sku], monthKeys);
            if (totals.qty === 0 && totals.sales === 0) {
                return;
            }
            if (sku === CONFIG.adjustmentLabel && winner) {
                return;
            }
            if (!winner || totals.sales > winner.sales || (totals.sales === winner.sales && totals.qty > winner.qty)) {
                winner = {
                    sku: sku,
                    qty: totals.qty,
                    sales: totals.sales
                };
            }
        });
        return winner || {
            sku: CONFIG.adjustmentLabel,
            qty: 0,
            sales: 0
        };
    }

    function findBestGrowthSku(skuMonthly, currentKeys, previousKeys) {
        let winner = null;
        Object.keys(skuMonthly).forEach(sku => {
            if (sku === CONFIG.adjustmentLabel) {
                return;
            }
            const current = sumMetrics(skuMonthly[sku], currentKeys);
            const previous = sumMetrics(skuMonthly[sku], previousKeys);
            const growth = calculateGrowth(current.sales, previous.sales);
            if (!Number.isFinite(growth) || previous.sales <= 0) {
                return;
            }
            if (!winner || growth > winner.growth) {
                winner = {
                    sku: sku,
                    growth: growth,
                    current: current,
                    previous: previous
                };
            }
        });
        return winner;
    }

    function buildMonthlyRows(monthKeys, monthlyTotals, skuMonthly) {
        return monthKeys.map(monthKey => {
            const current = monthlyTotals[monthKey] || { qty: 0, sales: 0 };
            const previous = monthlyTotals[previousMonthKey(monthKey)] || { qty: 0, sales: 0 };
            const lastYear = monthlyTotals[sameMonthLastYearKey(monthKey)] || { qty: 0, sales: 0 };
            return {
                monthKey: monthKey,
                qty: current.qty,
                sales: current.sales,
                qtyMoM: calculateGrowth(current.qty, previous.qty),
                salesMoM: calculateGrowth(current.sales, previous.sales),
                qtyYoY: calculateGrowth(current.qty, lastYear.qty),
                salesYoY: calculateGrowth(current.sales, lastYear.sales),
                topSku: findTopSkuForMonths(skuMonthly, [monthKey])
            };
        });
    }

    function buildDashboard(records, channel) {
        const monthKeys = buildMonthRange(records);
        const monthlyTotals = blankMetricMap(monthKeys);
        const yearlyTotals = {};
        const skuMonthly = {};

        CONFIG.focusYears.forEach(year => {
            yearlyTotals[year] = { qty: 0, sales: 0 };
        });

        records.forEach(record => {
            if (!monthlyTotals[record.monthKey]) {
                monthlyTotals[record.monthKey] = { qty: 0, sales: 0 };
            }
            monthlyTotals[record.monthKey].qty += record.qty;
            monthlyTotals[record.monthKey].sales += record.sales;
            yearlyTotals[record.year].qty += record.qty;
            yearlyTotals[record.year].sales += record.sales;

            if (!skuMonthly[record.sku]) {
                skuMonthly[record.sku] = blankMetricMap(monthKeys);
            }
            skuMonthly[record.sku][record.monthKey].qty += record.qty;
            skuMonthly[record.sku][record.monthKey].sales += record.sales;
        });

        const baseSkus = Object.keys(skuMonthly).sort((left, right) => sortSkus(left, right, skuMonthly, monthKeys));
        const latestMonthKey = monthKeys[monthKeys.length - 1] || toMonthKey(CONFIG.startYear, 1);
        const latestMonthNumber = Number(latestMonthKey.slice(5, 7));
        const yearKeys = {};
        const samePeriodKeys = {};

        CONFIG.focusYears.forEach(year => {
            yearKeys[year] = monthKeys.filter(monthKey => monthKey.startsWith(String(year) + "-"));
            samePeriodKeys[year] = yearKeys[year].filter(monthKey => Number(monthKey.slice(5, 7)) <= latestMonthNumber);
        });

        const monthlyRows = buildMonthlyRows(monthKeys, monthlyTotals, skuMonthly);
        const latestMonth = monthlyTotals[latestMonthKey] || { qty: 0, sales: 0 };
        const previousMonth = monthlyTotals[previousMonthKey(latestMonthKey)] || { qty: 0, sales: 0 };
        const latestMonthLastYear = monthlyTotals[sameMonthLastYearKey(latestMonthKey)] || { qty: 0, sales: 0 };
        const samePeriodByYear = {};

        CONFIG.focusYears.forEach(year => {
            samePeriodByYear[year] = sumMetrics(monthlyTotals, samePeriodKeys[year]);
        });

        const productMixSinceStart = buildProductMix(skuMonthly, monthKeys, CONFIG.structureStartMonth);

        return {
            channel: channel,
            records: records,
            monthKeys: monthKeys,
            yearKeys: yearKeys,
            samePeriodKeys: samePeriodKeys,
            monthlyTotals: monthlyTotals,
            yearlyTotals: yearlyTotals,
            monthlyRows: monthlyRows,
            skuMonthly: skuMonthly,
            baseSkus: baseSkus,
            latestMonthKey: latestMonthKey,
            latestMonthNumber: latestMonthNumber,
            latestMonth: latestMonth,
            previousMonth: previousMonth,
            latestMonthLastYear: latestMonthLastYear,
            samePeriodByYear: samePeriodByYear,
            latestTopSku: findTopSkuForMonths(skuMonthly, [latestMonthKey]),
            ytdTopSku: findTopSkuForMonths(skuMonthly, samePeriodKeys[2026]),
            annualTopSku2025: findTopSkuForMonths(skuMonthly, yearKeys[2025]),
            bestGrowthSku2026: findBestGrowthSku(skuMonthly, samePeriodKeys[2026], samePeriodKeys[2025]),
            productMixSinceStart: productMixSinceStart
        };
    }

    function getActiveDashboard() {
        return state.dashboards[state.activeChannelKey] || null;
    }

    function channelStyle(channel) {
        return "--channel-accent:" + channel.accent + ";--channel-soft:" + hexToRgba(channel.accent, 0.14) + ";";
    }

    function hydrateStaticCopy() {
        document.title = "NATM | 2024-2026 Multi-Channel Dashboard";

        if (els.eyebrow) {
            els.eyebrow.textContent = "NATM Unified Dashboard";
        }
        if (els.heroTitle) {
            els.heroTitle.textContent = "NATM 2024-2026 多渠道销售分析";
        }
        if (els.heroCopy) {
            els.heroCopy.innerHTML = [
                "页面直接读取 Google Sheet 中 <code>NFM Sales Data</code>、<code>RCW Sales Data</code>、",
                "<code>Abt Sales Data</code> 与 <code>BSM Sales Data</code> 的 <code>A:F</code> 明细数据，",
                "按基础 SKU 自动聚合，并统一输出年度汇总、月度环比 / 同比、SKU 趋势和明细表。",
                "后续如继续在同一个工作簿增加 NATM 渠道，也可以沿用这套结构继续扩展。"
            ].join("");
        }
        if (els.refreshBtn) {
            els.refreshBtn.textContent = "刷新在线数据";
        }
        if (els.footer) {
            els.footer.innerHTML = [
                "数据源：",
                "<a href=\"" + CONFIG.workbookUrl + "\" target=\"_blank\" rel=\"noreferrer\">Google Sheet - NATM Sales Workbook</a>",
                "。页面实时读取四个 NATM 渠道工作表的 <code>A:F</code> 明细，",
                "并按 2024 全年、2025 全年和 2026 YTD 做统一口径分析。"
            ].join("");
        }

        document.querySelectorAll("#metricToggle .metric-btn").forEach(button => {
            button.textContent = button.dataset.metric === "sales" ? "销售额" : "销量";
        });

        if (els.skuMetricSelect) {
            Array.from(els.skuMetricSelect.options).forEach(option => {
                option.text = option.value === "sales" ? "销售额" : "销量";
            });
        }
    }

    function renderProfileTags(values, className) {
        if (!values || !values.length) {
            return "<span class=\"profile-tag muted\">暂无</span>";
        }
        return values.map(value => {
            return "<span class=\"profile-tag" + (className ? " " + className : "") + "\">" + escapeHtml(value) + "</span>";
        }).join("");
    }

    function renderProductMixTags(items, popSkuSet) {
        if (!items.length) {
            return "<span class=\"profile-tag muted\">2025-09 至今暂无卖出记录</span>";
        }
        return items.map(item => {
            const isPopSku = popSkuSet.has(item.sku);
            const shareText = item.salesShare !== null ? formatPercent(item.salesShare, true) : "—";
            return "<span class=\"profile-tag" + (isPopSku ? "" : " muted") + "\">" + escapeHtml(item.sku) + " " + formatCurrency(item.sales) + " · " + shareText + "</span>";
        }).join("");
    }

    function renderBusinessOverview() {
        els.businessOverviewGrid.innerHTML = CONFIG.channels.map(channel => {
            const profile = CONFIG.channelProfiles[channel.key];
            const dashboard = state.dashboards[channel.key];
            const productMix = dashboard.productMixSinceStart;
            const popSkuSet = new Set(profile.popSkus);
            const soldSkuSet = new Set(productMix.items.map(item => item.sku));
            const popInSales = profile.popSkus.filter(sku => soldSkuSet.has(sku));
            const nonPopSold = productMix.items.filter(item => !popSkuSet.has(item.sku)).map(item => item.sku);
            const popCoverageText = profile.popSkus.length ? ("POP 覆盖 " + popInSales.length + "/" + profile.popSkus.length) : "当前无 POP 组合";

            return [
                "<article class=\"profile-card\" style=\"" + channelStyle(channel) + "\">",
                "<div class=\"profile-card-head\">",
                "<div>",
                "<h3 class=\"profile-card-title\">" + escapeHtml(profile.displayName || channel.label) + "</h3>",
                "<p class=\"profile-card-sub\">商务对接 + 出货价 + POP产品情况 + 2025-09 至今产品结构</p>",
                "</div>",
                "<span class=\"channel-chip\">NATM</span>",
                "</div>",
                "<div class=\"profile-block\">",
                "<p class=\"profile-label\">商务对接</p>",
                "<div class=\"profile-tag-list\">" + renderProfileTags(profile.businessTags) + "</div>",
                "<p class=\"profile-copy\"><strong>Buyer：</strong>" + escapeHtml(profile.buyer) + "</p>",
                "<p class=\"profile-copy\"><strong>Account：</strong>" + escapeHtml(profile.account) + "</p>",
                "<p class=\"profile-copy\"><strong>Setup：</strong>" + escapeHtml(profile.setup) + "</p>",
                "</div>",
                "<div class=\"profile-block\">",
                "<p class=\"profile-label\">出货价</p>",
                "<div class=\"profile-price-grid\">",
                "<article class=\"profile-price\"><p class=\"profile-price-label\">2025 出货价</p><p class=\"profile-price-value\">" + escapeHtml(profile.shipPrice2025) + "</p></article>",
                "<article class=\"profile-price\"><p class=\"profile-price-label\">2026 出货价</p><p class=\"profile-price-value\">" + escapeHtml(profile.shipPrice2026) + "</p></article>",
                "</div>",
                "</div>",
                "<div class=\"profile-block\">",
                "<p class=\"profile-label\">POP产品情况</p>",
                "<p class=\"profile-copy\"><strong>POP尺寸：</strong>" + escapeHtml(profile.popSize) + "</p>",
                "<div class=\"profile-tag-list\">" + renderProfileTags(profile.popSkus) + "</div>",
                "</div>",
                "<div class=\"profile-block\">",
                "<p class=\"profile-label\">产品结构</p>",
                "<p class=\"profile-mix-note\">2025-09 至 " + escapeHtml(dashboard.latestMonthKey) + "，累计卖出 " + productMix.items.length + " 个 SKU，销量 " + formatNumber(productMix.totalQty) + "，销售额 " + formatCurrency(productMix.totalSales) + "。" + popCoverageText + (nonPopSold.length ? "，非POP卖出：" + escapeHtml(nonPopSold.join(" / ")) : "") + "</p>",
                "<div class=\"profile-tag-list\">" + renderProductMixTags(productMix.items, popSkuSet) + "</div>",
                "</div>",
                "</article>"
            ].join("");
        }).join("");
    }

    function renderChannelOverview() {
        els.channelOverviewGrid.innerHTML = CONFIG.channels.map(channel => {
            const dashboard = state.dashboards[channel.key];
            if (!dashboard) {
                return "";
            }
            const ytdSalesYoY = calculateGrowth(dashboard.samePeriodByYear[2026].sales, dashboard.samePeriodByYear[2025].sales);
            const activeClass = state.activeChannelKey === channel.key ? " active" : "";
            return [
                "<article class=\"channel-card" + activeClass + "\" data-channel=\"" + channel.key + "\" style=\"" + channelStyle(channel) + "\">",
                "<div class=\"channel-card-head\">",
                "<div>",
                "<p class=\"channel-card-title\">" + channel.label + "</p>",
                "<p class=\"channel-card-caption\">2026 YTD Sales</p>",
                "</div>",
                "<span class=\"channel-chip\">NATM</span>",
                "</div>",
                "<div class=\"channel-card-kpi\">" + formatCurrency(dashboard.samePeriodByYear[2026].sales) + "</div>",
                "<div class=\"channel-card-meta\">",
                "<span>2026 YTD 销量 " + formatNumber(dashboard.samePeriodByYear[2026].qty) + "</span>",
                renderDelta(ytdSalesYoY),
                "</div>",
                "<div class=\"channel-card-foot\">",
                "<span>最新月份 " + dashboard.latestMonthKey + "</span>",
                "<strong>" + formatCurrency(dashboard.latestMonth.sales) + "</strong>",
                "</div>",
                "</article>"
            ].join("");
        }).join("");
    }

    function renderChannelComparison() {
        const metric = state.compareMetric;
        const categories = ["2024", "2025", "2026 YTD"];

        if (!state.charts.channelCompare) {
            state.charts.channelCompare = echarts.init(document.getElementById("channelCompareChart"));
        }

        state.charts.channelCompare.setOption({
            animationDuration: 500,
            tooltip: {
                trigger: "axis"
            },
            legend: {
                top: 0,
                textStyle: { color: CONFIG.colors.muted }
            },
            grid: { left: 56, right: 24, top: 44, bottom: 42 },
            xAxis: {
                type: "category",
                data: categories,
                axisLine: { lineStyle: { color: "rgba(23,32,51,0.15)" } },
                axisLabel: { color: CONFIG.colors.muted }
            },
            yAxis: {
                type: "value",
                axisLabel: {
                    color: CONFIG.colors.muted,
                    formatter: function(value) {
                        return metric === "sales" ? formatCurrency(value, 0) : formatNumber(value);
                    }
                },
                splitLine: { lineStyle: { color: "rgba(23,32,51,0.08)" } }
            },
            series: CONFIG.channels.map(channel => {
                const dashboard = state.dashboards[channel.key];
                return {
                    name: channel.label,
                    type: "bar",
                    barMaxWidth: 24,
                    itemStyle: {
                        color: channel.accent,
                        borderRadius: [10, 10, 0, 0]
                    },
                    data: [
                        dashboard.yearlyTotals[2024][metric],
                        dashboard.yearlyTotals[2025][metric],
                        dashboard.samePeriodByYear[2026][metric]
                    ]
                };
            })
        });

        els.channelSummaryBody.innerHTML = CONFIG.channels.map(channel => {
            const dashboard = state.dashboards[channel.key];
            const annualQtyYoY2025 = calculateGrowth(dashboard.yearlyTotals[2025].qty, dashboard.yearlyTotals[2024].qty);
            const annualSalesYoY2025 = calculateGrowth(dashboard.yearlyTotals[2025].sales, dashboard.yearlyTotals[2024].sales);
            const ytdQtyYoY2026 = calculateGrowth(dashboard.samePeriodByYear[2026].qty, dashboard.samePeriodByYear[2025].qty);
            const ytdSalesYoY2026 = calculateGrowth(dashboard.samePeriodByYear[2026].sales, dashboard.samePeriodByYear[2025].sales);
            return [
                "<tr>",
                "<td><strong>" + channel.label + "</strong></td>",
                "<td>" + formatNumber(dashboard.yearlyTotals[2024].qty) + "</td>",
                "<td>" + formatCurrency(dashboard.yearlyTotals[2024].sales) + "</td>",
                "<td>" + formatNumber(dashboard.yearlyTotals[2025].qty) + "</td>",
                "<td>" + formatCurrency(dashboard.yearlyTotals[2025].sales) + "</td>",
                "<td>" + renderDelta(annualQtyYoY2025) + "</td>",
                "<td>" + renderDelta(annualSalesYoY2025) + "</td>",
                "<td>" + formatNumber(dashboard.samePeriodByYear[2026].qty) + "</td>",
                "<td>" + formatCurrency(dashboard.samePeriodByYear[2026].sales) + "</td>",
                "<td>" + renderDelta(ytdQtyYoY2026) + "</td>",
                "<td>" + renderDelta(ytdSalesYoY2026) + "</td>",
                "<td>" + dashboard.latestMonthKey + "<br>" + formatCurrency(dashboard.latestMonth.sales) + "</td>",
                "</tr>"
            ].join("");
        }).join("");
    }

    function renderChannelTabs() {
        els.channelTabs.innerHTML = CONFIG.channels.map(channel => {
            const dashboard = state.dashboards[channel.key];
            const activeClass = state.activeChannelKey === channel.key ? " active" : "";
            return [
                "<button class=\"channel-tab" + activeClass + "\" data-channel=\"" + channel.key + "\" style=\"" + channelStyle(channel) + "\">",
                "<span class=\"channel-tab-label\">" + channel.label + "</span>",
                "<span class=\"channel-tab-meta\">2026 YTD " + formatCurrency(dashboard.samePeriodByYear[2026].sales) + "</span>",
                "<span class=\"channel-tab-meta\">最新月份 " + dashboard.latestMonthKey + " · " + formatCurrency(dashboard.latestMonth.sales) + "</span>",
                "</button>"
            ].join("");
        }).join("");
    }

    function renderKpis(dashboard) {
        const latestSalesMoM = calculateGrowth(dashboard.latestMonth.sales, dashboard.previousMonth.sales);
        const latestSalesYoY = calculateGrowth(dashboard.latestMonth.sales, dashboard.latestMonthLastYear.sales);
        const latestQtyMoM = calculateGrowth(dashboard.latestMonth.qty, dashboard.previousMonth.qty);
        const latestQtyYoY = calculateGrowth(dashboard.latestMonth.qty, dashboard.latestMonthLastYear.qty);
        const ytdSalesYoY = calculateGrowth(dashboard.samePeriodByYear[2026].sales, dashboard.samePeriodByYear[2025].sales);
        const ytdQtyYoY = calculateGrowth(dashboard.samePeriodByYear[2026].qty, dashboard.samePeriodByYear[2025].qty);

        els.kpiGrid.innerHTML = [
            {
                title: dashboard.channel.label + " 最新月销售额",
                value: formatCurrency(dashboard.latestMonth.sales),
                sub: [
                    "<span>" + dashboard.latestMonthKey + "</span>",
                    "<span>环比</span>",
                    renderDelta(latestSalesMoM),
                    "<span>同比</span>",
                    renderDelta(latestSalesYoY)
                ].join("")
            },
            {
                title: dashboard.channel.label + " 最新月销量",
                value: formatNumber(dashboard.latestMonth.qty),
                sub: [
                    "<span>" + dashboard.latestMonthKey + "</span>",
                    "<span>环比</span>",
                    renderDelta(latestQtyMoM),
                    "<span>同比</span>",
                    renderDelta(latestQtyYoY)
                ].join("")
            },
            {
                title: dashboard.channel.label + " 2026 YTD 销售额",
                value: formatCurrency(dashboard.samePeriodByYear[2026].sales),
                sub: [
                    "<span>同期对比 2025</span>",
                    renderDelta(ytdSalesYoY),
                    "<span>2025 全年 " + formatCurrency(dashboard.yearlyTotals[2025].sales) + "</span>"
                ].join("")
            },
            {
                title: dashboard.channel.label + " 2026 YTD 销量",
                value: formatNumber(dashboard.samePeriodByYear[2026].qty),
                sub: [
                    "<span>同期对比 2025</span>",
                    renderDelta(ytdQtyYoY),
                    "<span>2025 全年 " + formatNumber(dashboard.yearlyTotals[2025].qty) + "</span>"
                ].join("")
            }
        ].map(card => {
            return [
                "<article class=\"kpi-card\">",
                "<p class=\"kpi-label\">" + card.title + "</p>",
                "<p class=\"kpi-value\">" + card.value + "</p>",
                "<div class=\"kpi-sub\">" + card.sub + "</div>",
                "</article>"
            ].join("");
        }).join("");
    }

    function renderInsights(dashboard) {
        const annualQtyYoY2025 = calculateGrowth(dashboard.yearlyTotals[2025].qty, dashboard.yearlyTotals[2024].qty);
        const annualSalesYoY2025 = calculateGrowth(dashboard.yearlyTotals[2025].sales, dashboard.yearlyTotals[2024].sales);
        const ytdQtyYoY2026 = calculateGrowth(dashboard.samePeriodByYear[2026].qty, dashboard.samePeriodByYear[2025].qty);
        const ytdSalesYoY2026 = calculateGrowth(dashboard.samePeriodByYear[2026].sales, dashboard.samePeriodByYear[2025].sales);
        const bestGrowthLine = dashboard.bestGrowthSku2026
            ? dashboard.bestGrowthSku2026.sku + " 是 2026 同期销售额增速最强的 SKU，较 2025 同期" + describeGrowth(dashboard.bestGrowthSku2026.growth, "增长", "下滑") + "。"
            : "当前缺少足够的可比基础，暂不输出 SKU 同期增速冠军。";

        els.insightPanel.innerHTML = [
            [
                "<section class=\"insight-callout\">",
                "<h3>" + dashboard.channel.label + " 年度节奏</h3>",
                "<p>",
                "2025 全年销量 " + formatNumber(dashboard.yearlyTotals[2025].qty) + "，销售额 " + formatCurrency(dashboard.yearlyTotals[2025].sales) + "。",
                "相较 2024 年，销量" + describeGrowth(annualQtyYoY2025, "增长", "下滑") + "，销售额" + describeGrowth(annualSalesYoY2025, "增长", "下滑") + "。",
                "</p>",
                "</section>"
            ].join(""),
            [
                "<section class=\"insight-callout\">",
                "<h3>2026 YTD 进度</h3>",
                "<p>",
                "截至 " + dashboard.latestMonthKey + "，销量累计 " + formatNumber(dashboard.samePeriodByYear[2026].qty) + "，销售额累计 " + formatCurrency(dashboard.samePeriodByYear[2026].sales) + "。",
                "对比 2025 同期，销量" + describeGrowth(ytdQtyYoY2026, "增长", "下滑") + "，销售额" + describeGrowth(ytdSalesYoY2026, "增长", "下滑") + "。",
                "</p>",
                "</section>"
            ].join(""),
            [
                "<section class=\"insight-callout\">",
                "<h3>SKU 聚焦</h3>",
                "<p>",
                "最新月份 " + dashboard.latestMonthKey + " 的销售额 Top SKU 为 " + dashboard.latestTopSku.sku + "，贡献 " + formatCurrency(dashboard.latestTopSku.sales) + "。",
                "若看 2026 YTD，则当前领先 SKU 为 " + dashboard.ytdTopSku.sku + "，累计销售额 " + formatCurrency(dashboard.ytdTopSku.sales) + "。",
                bestGrowthLine,
                "</p>",
                "</section>"
            ].join("")
        ].join("");
    }

    function renderYearSummary(dashboard) {
        const annualQtyYoY2025 = calculateGrowth(dashboard.yearlyTotals[2025].qty, dashboard.yearlyTotals[2024].qty);
        const annualSalesYoY2025 = calculateGrowth(dashboard.yearlyTotals[2025].sales, dashboard.yearlyTotals[2024].sales);
        const ytdQtyYoY2026 = calculateGrowth(dashboard.samePeriodByYear[2026].qty, dashboard.samePeriodByYear[2025].qty);
        const ytdSalesYoY2026 = calculateGrowth(dashboard.samePeriodByYear[2026].sales, dashboard.samePeriodByYear[2025].sales);

        els.yearSummaryGrid.innerHTML = [
            { label: "2024 全年销量", value: formatNumber(dashboard.yearlyTotals[2024].qty) },
            { label: "2024 全年销售额", value: formatCurrency(dashboard.yearlyTotals[2024].sales) },
            { label: "2025 全年销量", value: formatNumber(dashboard.yearlyTotals[2025].qty) },
            { label: "2025 全年销售额", value: formatCurrency(dashboard.yearlyTotals[2025].sales) },
            { label: "2025 全年同比", value: formatPercent(annualSalesYoY2025, false) + " / " + formatPercent(annualQtyYoY2025, false) },
            { label: "2026 YTD 同期同比", value: formatPercent(ytdSalesYoY2026, false) + " / " + formatPercent(ytdQtyYoY2026, false) }
        ].map(item => {
            return [
                "<article class=\"mini-card\">",
                "<p class=\"mini-card-label\">" + item.label + "</p>",
                "<p class=\"mini-card-value\">" + item.value + "</p>",
                "</article>"
            ].join("");
        }).join("");

        if (!state.charts.yearCompare) {
            state.charts.yearCompare = echarts.init(document.getElementById("yearCompareChart"));
        }

        state.charts.yearCompare.setOption({
            animationDuration: 500,
            color: [dashboard.channel.accent, CONFIG.colors.sales],
            tooltip: {
                trigger: "axis"
            },
            legend: {
                top: 0,
                textStyle: { color: CONFIG.colors.muted }
            },
            grid: { left: 54, right: 54, top: 42, bottom: 40 },
            xAxis: {
                type: "category",
                data: ["2024", "2025", "2026 YTD"],
                axisLine: { lineStyle: { color: "rgba(23,32,51,0.15)" } },
                axisLabel: { color: CONFIG.colors.muted }
            },
            yAxis: [
                {
                    type: "value",
                    name: "销量",
                    axisLabel: {
                        color: CONFIG.colors.muted,
                        formatter: function(value) {
                            return formatNumber(value);
                        }
                    },
                    splitLine: { lineStyle: { color: "rgba(23,32,51,0.08)" } }
                },
                {
                    type: "value",
                    name: "销售额",
                    axisLabel: {
                        color: CONFIG.colors.muted,
                        formatter: function(value) {
                            return formatCurrency(value, 0);
                        }
                    },
                    splitLine: { show: false }
                }
            ],
            series: [
                {
                    name: "销量",
                    type: "bar",
                    barMaxWidth: 34,
                    itemStyle: {
                        color: dashboard.channel.accent,
                        borderRadius: [10, 10, 0, 0]
                    },
                    data: [
                        dashboard.yearlyTotals[2024].qty,
                        dashboard.yearlyTotals[2025].qty,
                        dashboard.samePeriodByYear[2026].qty
                    ]
                },
                {
                    name: "销售额",
                    type: "line",
                    yAxisIndex: 1,
                    smooth: true,
                    symbolSize: 8,
                    lineStyle: { width: 3, color: CONFIG.colors.sales },
                    itemStyle: { color: CONFIG.colors.sales },
                    data: [
                        dashboard.yearlyTotals[2024].sales,
                        dashboard.yearlyTotals[2025].sales,
                        dashboard.samePeriodByYear[2026].sales
                    ]
                }
            ]
        });
    }

    function populateSkuSelector(dashboard) {
        const previousSku = state.selectedSkuByChannel[dashboard.channel.key];
        const nextSku = dashboard.baseSkus.includes(previousSku) ? previousSku : (dashboard.baseSkus[0] || "");
        state.selectedSkuByChannel[dashboard.channel.key] = nextSku;

        els.skuSelect.innerHTML = dashboard.baseSkus.map(sku => {
            return "<option value=\"" + sku + "\">" + sku + "</option>";
        }).join("");

        els.skuSelect.value = nextSku;
        els.skuMetricSelect.value = state.selectedMetric;
    }

    function renderMonthlyOverview(dashboard) {
        if (!state.charts.monthlyOverview) {
            state.charts.monthlyOverview = echarts.init(document.getElementById("monthlyOverviewChart"));
        }

        state.charts.monthlyOverview.setOption({
            animationDuration: 500,
            tooltip: {
                trigger: "axis"
            },
            legend: {
                top: 0,
                textStyle: { color: CONFIG.colors.muted }
            },
            grid: { left: 54, right: 54, top: 42, bottom: 50 },
            xAxis: {
                type: "category",
                data: dashboard.monthKeys,
                axisLine: { lineStyle: { color: "rgba(23,32,51,0.15)" } },
                axisLabel: { color: CONFIG.colors.muted, rotate: dashboard.monthKeys.length > 20 ? 38 : 0 }
            },
            yAxis: [
                {
                    type: "value",
                    name: "销售额",
                    axisLabel: {
                        color: CONFIG.colors.muted,
                        formatter: function(value) {
                            return formatCurrency(value, 0);
                        }
                    },
                    splitLine: { lineStyle: { color: "rgba(23,32,51,0.08)" } }
                },
                {
                    type: "value",
                    name: "销量",
                    axisLabel: {
                        color: CONFIG.colors.muted,
                        formatter: function(value) {
                            return formatNumber(value);
                        }
                    },
                    splitLine: { show: false }
                }
            ],
            series: [
                {
                    name: "销售额",
                    type: "bar",
                    barMaxWidth: 22,
                    itemStyle: {
                        color: dashboard.channel.accent,
                        borderRadius: [8, 8, 0, 0]
                    },
                    data: dashboard.monthKeys.map(monthKey => dashboard.monthlyTotals[monthKey].sales)
                },
                {
                    name: "销量",
                    type: "line",
                    yAxisIndex: 1,
                    smooth: true,
                    symbolSize: 7,
                    lineStyle: { width: 3, color: CONFIG.colors.sales },
                    itemStyle: { color: CONFIG.colors.sales },
                    data: dashboard.monthKeys.map(monthKey => dashboard.monthlyTotals[monthKey].qty)
                }
            ]
        });

        els.monthlySummaryBody.innerHTML = dashboard.monthlyRows.map(row => {
            return [
                "<tr>",
                "<td>" + row.monthKey + "</td>",
                "<td>" + formatNumber(row.qty) + "</td>",
                "<td>" + formatCurrency(row.sales) + "</td>",
                "<td>" + renderDelta(row.qtyMoM) + "</td>",
                "<td>" + renderDelta(row.salesMoM) + "</td>",
                "<td>" + renderDelta(row.qtyYoY) + "</td>",
                "<td>" + renderDelta(row.salesYoY) + "</td>",
                "<td>" + row.topSku.sku + " / " + formatCurrency(row.topSku.sales) + "</td>",
                "</tr>"
            ].join("");
        }).join("");
    }

    function renderSkuMatrix(dashboard) {
        const metric = state.selectedMetric;
        const headerCells = dashboard.monthKeys.map(monthKey => "<th>" + monthKey + "</th>").join("");
        const bodyRows = dashboard.baseSkus.map(sku => {
            const cells = dashboard.monthKeys.map(monthKey => {
                const value = dashboard.skuMonthly[sku][monthKey][metric];
                return "<td>" + (metric === "sales" ? formatCurrency(value) : formatNumber(value)) + "</td>";
            }).join("");
            return "<tr><td><strong>" + sku + "</strong></td>" + cells + "</tr>";
        }).join("");

        els.skuMatrixTable.innerHTML = [
            "<thead><tr><th>SKU</th>" + headerCells + "</tr></thead>",
            "<tbody>" + bodyRows + "</tbody>"
        ].join("");
    }

    function renderSkuTrend(dashboard) {
        const selectedSku = state.selectedSkuByChannel[dashboard.channel.key];
        if (!selectedSku) {
            return;
        }

        if (!state.charts.skuTrend) {
            state.charts.skuTrend = echarts.init(document.getElementById("skuTrendChart"));
        }

        const metric = state.selectedMetric;
        const series = dashboard.monthKeys.map(monthKey => dashboard.skuMonthly[selectedSku][monthKey][metric]);
        const yoySeries = dashboard.monthKeys.map(monthKey => {
            const current = dashboard.skuMonthly[selectedSku][monthKey][metric];
            const previous = (dashboard.skuMonthly[selectedSku][sameMonthLastYearKey(monthKey)] || { qty: 0, sales: 0 })[metric];
            const growth = calculateGrowth(current, previous);
            return Number.isFinite(growth) ? Number((growth * 100).toFixed(1)) : null;
        });

        state.charts.skuTrend.setOption({
            animationDuration: 500,
            color: [dashboard.channel.accent, CONFIG.colors.yoy],
            tooltip: {
                trigger: "axis"
            },
            legend: {
                top: 0,
                textStyle: { color: CONFIG.colors.muted }
            },
            grid: { left: 54, right: 72, top: 42, bottom: 46 },
            xAxis: {
                type: "category",
                data: dashboard.monthKeys,
                axisLine: { lineStyle: { color: "rgba(23,32,51,0.15)" } },
                axisLabel: { color: CONFIG.colors.muted, rotate: dashboard.monthKeys.length > 20 ? 38 : 0 }
            },
            yAxis: [
                {
                    type: "value",
                    name: metric === "sales" ? "销售额" : "销量",
                    axisLabel: {
                        color: CONFIG.colors.muted,
                        formatter: function(value) {
                            return metric === "sales" ? formatCurrency(value, 0) : formatNumber(value);
                        }
                    },
                    splitLine: { lineStyle: { color: "rgba(23,32,51,0.08)" } }
                },
                {
                    type: "value",
                    name: "同比%",
                    axisLabel: {
                        color: CONFIG.colors.muted,
                        formatter: function(value) {
                            return value + "%";
                        }
                    },
                    splitLine: { show: false }
                }
            ],
            series: [
                {
                    name: selectedSku + " " + (metric === "sales" ? "销售额" : "销量"),
                    type: "bar",
                    barMaxWidth: 26,
                    itemStyle: { borderRadius: [10, 10, 0, 0] },
                    data: series
                },
                {
                    name: "同比%",
                    type: "line",
                    yAxisIndex: 1,
                    smooth: true,
                    symbolSize: 8,
                    lineStyle: { width: 3 },
                    data: yoySeries
                }
            ]
        });
    }

    function renderSkuDetailTable(dashboard) {
        const selectedSku = state.selectedSkuByChannel[dashboard.channel.key];
        if (!selectedSku) {
            els.skuDetailBody.innerHTML = "";
            return;
        }

        els.skuDetailBody.innerHTML = dashboard.monthKeys.map(monthKey => {
            const current = dashboard.skuMonthly[selectedSku][monthKey];
            const previous = dashboard.skuMonthly[selectedSku][previousMonthKey(monthKey)] || { qty: 0, sales: 0 };
            const lastYear = dashboard.skuMonthly[selectedSku][sameMonthLastYearKey(monthKey)] || { qty: 0, sales: 0 };
            return [
                "<tr>",
                "<td>" + monthKey + "</td>",
                "<td>" + formatNumber(current.qty) + "</td>",
                "<td>" + formatCurrency(current.sales) + "</td>",
                "<td>" + renderDelta(calculateGrowth(current.qty, previous.qty)) + "</td>",
                "<td>" + renderDelta(calculateGrowth(current.sales, previous.sales)) + "</td>",
                "<td>" + renderDelta(calculateGrowth(current.qty, lastYear.qty)) + "</td>",
                "<td>" + renderDelta(calculateGrowth(current.sales, lastYear.sales)) + "</td>",
                "</tr>"
            ].join("");
        }).join("");
    }

    function renderSkuSummaryTable(dashboard) {
        els.skuSummaryBody.innerHTML = dashboard.baseSkus.map(sku => {
            const total2024 = sumMetrics(dashboard.skuMonthly[sku], dashboard.yearKeys[2024]);
            const total2025 = sumMetrics(dashboard.skuMonthly[sku], dashboard.yearKeys[2025]);
            const samePeriod2025 = sumMetrics(dashboard.skuMonthly[sku], dashboard.samePeriodKeys[2025]);
            const total2026 = sumMetrics(dashboard.skuMonthly[sku], dashboard.samePeriodKeys[2026]);
            const latestMonthValue = dashboard.skuMonthly[sku][dashboard.latestMonthKey] || { qty: 0, sales: 0 };
            return [
                "<tr>",
                "<td>" + sku + "</td>",
                "<td>" + formatNumber(total2024.qty) + "</td>",
                "<td>" + formatCurrency(total2024.sales) + "</td>",
                "<td>" + formatNumber(total2025.qty) + "</td>",
                "<td>" + formatCurrency(total2025.sales) + "</td>",
                "<td>" + renderDelta(calculateGrowth(total2025.qty, total2024.qty)) + "</td>",
                "<td>" + renderDelta(calculateGrowth(total2025.sales, total2024.sales)) + "</td>",
                "<td>" + formatNumber(total2026.qty) + "</td>",
                "<td>" + formatCurrency(total2026.sales) + "</td>",
                "<td>" + renderDelta(calculateGrowth(total2026.qty, samePeriod2025.qty)) + "</td>",
                "<td>" + renderDelta(calculateGrowth(total2026.sales, samePeriod2025.sales)) + "</td>",
                "<td>" + formatNumber(latestMonthValue.qty) + "</td>",
                "<td>" + formatCurrency(latestMonthValue.sales) + "</td>",
                "</tr>"
            ].join("");
        }).join("");
    }

    function renderActiveChannel() {
        const dashboard = getActiveDashboard();
        if (!dashboard) {
            return;
        }

        renderKpis(dashboard);
        renderInsights(dashboard);
        renderYearSummary(dashboard);
        populateSkuSelector(dashboard);
        renderMonthlyOverview(dashboard);
        renderSkuMatrix(dashboard);
        renderSkuTrend(dashboard);
        renderSkuDetailTable(dashboard);
        renderSkuSummaryTable(dashboard);

        const globalLatestMonth = CONFIG.channels.reduce((latest, channel) => {
            const item = state.dashboards[channel.key];
            if (!item) {
                return latest;
            }
            return !latest || monthIndex(item.latestMonthKey) > monthIndex(latest) ? item.latestMonthKey : latest;
        }, "");

        els.dataRangeLabel.textContent = "数据范围：2024-01 至 " + globalLatestMonth + " · 当前渠道 " + dashboard.channel.label;
    }

    function renderAll() {
        renderBusinessOverview();
        renderChannelOverview();
        renderChannelComparison();
        renderChannelTabs();
        renderActiveChannel();
    }

    function setStatus(text, tone) {
        els.statusBadge.textContent = text;
        els.statusBadge.className = "status-badge " + tone;
    }

    function showError(message) {
        const markup = "<div class=\"error-state\">数据加载失败：" + message + "</div>";
        els.businessOverviewGrid.innerHTML = markup;
        els.channelOverviewGrid.innerHTML = markup;
        els.channelSummaryBody.innerHTML = "";
        els.channelTabs.innerHTML = "";
        els.kpiGrid.innerHTML = markup;
        els.insightPanel.innerHTML = "";
        els.yearSummaryGrid.innerHTML = "";
        els.monthlySummaryBody.innerHTML = "";
        els.skuMatrixTable.innerHTML = "";
        els.skuDetailBody.innerHTML = "";
        els.skuSummaryBody.innerHTML = "";
    }

    async function fetchChannelDashboard(channel, cacheKey) {
        const response = await fetch(csvUrlForSheet(channel.sheet, cacheKey), { cache: "no-store" });
        if (!response.ok) {
            throw new Error(channel.label + " HTTP " + response.status);
        }

        const text = await response.text();
        const parsed = Papa.parse(text, {
            header: true,
            skipEmptyLines: true
        });
        const records = parsed.data.map(normalizeRecord).filter(Boolean);
        if (!records.length) {
            throw new Error(channel.label + " 未读取到有效数据");
        }

        return buildDashboard(records, channel);
    }

    async function loadData() {
        setStatus("正在连接 Google Sheet", "loading");
        els.refreshBtn.disabled = true;

        try {
            const cacheKey = Date.now();
            const dashboards = {};
            const results = await Promise.all(CONFIG.channels.map(channel => fetchChannelDashboard(channel, cacheKey)));
            results.forEach(dashboard => {
                dashboards[dashboard.channel.key] = dashboard;
            });

            state.dashboards = dashboards;
            if (!state.dashboards[state.activeChannelKey]) {
                state.activeChannelKey = CONFIG.channels[0].key;
            }

            renderAll();
            setStatus("四个渠道已同步", "success");
            els.lastUpdated.textContent = "上次刷新：" + new Date().toLocaleString("zh-CN");
        } catch (error) {
            console.error(error);
            showError(error.message || "未知错误");
            setStatus("加载失败", "error");
            els.lastUpdated.textContent = "上次刷新：失败";
        } finally {
            els.refreshBtn.disabled = false;
        }
    }

    document.getElementById("compareMetricToggle").addEventListener("click", event => {
        const button = event.target.closest(".metric-btn");
        if (!button) {
            return;
        }
        state.compareMetric = button.dataset.metric;
        document.querySelectorAll("#compareMetricToggle .metric-btn").forEach(item => {
            item.classList.toggle("active", item === button);
        });
        if (Object.keys(state.dashboards).length) {
            renderChannelComparison();
        }
    });

    document.getElementById("metricToggle").addEventListener("click", event => {
        const button = event.target.closest(".metric-btn");
        if (!button) {
            return;
        }
        state.selectedMetric = button.dataset.metric;
        els.skuMetricSelect.value = state.selectedMetric;
        document.querySelectorAll("#metricToggle .metric-btn").forEach(item => {
            item.classList.toggle("active", item === button);
        });
        const dashboard = getActiveDashboard();
        if (dashboard) {
            renderSkuMatrix(dashboard);
            renderSkuTrend(dashboard);
        }
    });

    els.channelOverviewGrid.addEventListener("click", event => {
        const card = event.target.closest("[data-channel]");
        if (!card) {
            return;
        }
        state.activeChannelKey = card.dataset.channel;
        renderAll();
    });

    els.channelTabs.addEventListener("click", event => {
        const tab = event.target.closest("[data-channel]");
        if (!tab) {
            return;
        }
        state.activeChannelKey = tab.dataset.channel;
        renderAll();
    });

    els.skuSelect.addEventListener("change", event => {
        const dashboard = getActiveDashboard();
        if (!dashboard) {
            return;
        }
        state.selectedSkuByChannel[dashboard.channel.key] = event.target.value;
        renderSkuTrend(dashboard);
        renderSkuDetailTable(dashboard);
    });

    els.skuMetricSelect.addEventListener("change", event => {
        state.selectedMetric = event.target.value;
        const targetButton = document.querySelector("#metricToggle .metric-btn[data-metric=\"" + state.selectedMetric + "\"]");
        document.querySelectorAll("#metricToggle .metric-btn").forEach(item => {
            item.classList.toggle("active", item === targetButton);
        });
        const dashboard = getActiveDashboard();
        if (dashboard) {
            renderSkuMatrix(dashboard);
            renderSkuTrend(dashboard);
        }
    });

    els.refreshBtn.addEventListener("click", loadData);

    window.addEventListener("resize", () => {
        Object.values(state.charts).forEach(chart => {
            if (chart && typeof chart.resize === "function") {
                chart.resize();
            }
        });
    });

    hydrateStaticCopy();
    loadData();
})();
