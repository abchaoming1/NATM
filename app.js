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
        popNextStepPlan: {
            effectiveDate: "6/4",
            replacements: [
                { channelKey: "nfm", currentSku: "E310", nextSku: "E320（OpenDots 2）" },
                { channelKey: "rcw", currentSku: "E310", nextSku: "E320（OpenDots 2）" },
                { channelKey: "bsm", currentSku: "E310", nextSku: "E320（OpenDots 2）" },
                { channelKey: "nfm", currentSku: "T920", nextSku: "T921" },
                { channelKey: "rcw", currentSku: "T920", nextSku: "T921" },
                { channelKey: "bsm", currentSku: "T920", nextSku: "T921" }
            ],
            targetCombos: [
                { channelKey: "nfm", skus: ["E320", "T921", "S803", "S710", "S820", "S821"] },
                { channelKey: "rcw", skus: ["S820", "S821", "T921", "E320", "S710"] },
                { channelKey: "bsm", skus: ["E320", "T921", "S803", "S710", "S820", "S821"] }
            ],
            augustReplacements: [
                { channelKey: "nfm", currentSku: "S803", nextSku: "NCE（OpenRun 2）" },
                { channelKey: "bsm", currentSku: "S803", nextSku: "NCE（OpenRun 2）" }
            ]
        },
        businessPlaceholders: {
            launchPlan: {
                kicker: "产品结构",
                title: "产品结构上市更新计划",
                description: "整理耳夹线与骨导线的上市节奏、价格方案和 GTM 争取时间点，方便和渠道 POP / 送样节奏一起看。",
                badge: "Roadmap",
                rows: [
                    {
                        title: "耳夹线",
                        text: "6月发布 OpenDots 2（代号 being），定价预计 $199，OpenDots 同步退市。",
                        meta: "8月上旬计划发布 OpenDots Air（代号 looking），定价预计 $129；GTM 正在争取 6 月和 OpenDots 2 一起上市。"
                    },
                    {
                        title: "骨导线",
                        text: "8月下旬计划发布 OpenRun 2（代号 NCE），GTM 正在争取提前到 6 月发布。",
                        meta: "当前优先口径采用方案 B：OpenRun 降价到 $99，OpenRun 2 定价为 $129。"
                    }
                ]
            },
            costMargin: {
                kicker: "产品结构",
                title: "产品结构及成本及毛利率",
                description: "后续补充各渠道主推型号的成本、出货价、毛利率和结构变化。",
                body: "暂时空着"
            },
            sampling: {
                kicker: "渠道推进",
                title: "样品提需及渠道送样情况",
                description: "同步记录四个渠道新品的提需、是否已和 buyer 对接，以及当前仍缺失的送样动作。",
                badge: "Samples",
                rows: [
                    {
                        title: "当前送样情况",
                        text: "目前四个渠道均未有 OpenDots 2、OpenFit Pro、OpenDots Air 和 OpenRun 2 的送样情况。"
                    },
                    {
                        title: "当前样品提需",
                        text: "四个渠道刚好都有 OpenDots Air 的黑色、紫色各一个，总计 4+4=8 个；提需已完成，但暂未和 buyer 产生联系，也尚未推进送样。"
                    },
                    {
                        title: "待补项目",
                        text: "当前还缺 OpenDots 2 的提需和 OpenRun 2 的提需，后续需要一并补齐并推进到 buyer 沟通。"
                    }
                ]
            }
        },
        channelNotes: [
            {
                kicker: "渠道洞察",
                title: "渠道特点及用户画像",
                description: "沉淀四个渠道的价格带、门店特征、典型消费者和购买场景。",
                body: "暂时空着"
            },
            {
                kicker: "增长方向",
                title: "核心机会",
                description: "补充 2026 年渠道扩张、POP 更新、重点机型切换和增量机会。",
                body: "暂时空着"
            },
            {
                kicker: "年度管理",
                title: "今年目标",
                description: "后续放年度销售目标、重点项目节点和阶段性里程碑。",
                body: "暂时空着"
            }
        ],
        growthGoalText: "目标同比增长 30%~40% 及以上，底线 30%，标准 40%，挑战 50%+。",
        channelTargets: {
            nfm: { qty: 5000, sales: 500000 },
            rcw: { qty: 3500, sales: 175000 },
            abt: { qty: 1600, sales: 185000 },
            bsm: { qty: 3800, sales: 420000 }
        },
        channelResearch: {
            nfm: {
                profileTitle: "区域目的地大店，家庭整单购物心智强",
                profileBody: "NFM 官方资料显示其以家具、家电、电子一站式大卖场运营，四个核心卖场拥有超大 showroom 与完整 electronics assortment；结合 PPT 里“大店流量高、停留时间长、耳机专区成熟”的结论，更适合做体验驱动成交。",
                persona: "以家庭决策型客群为主，包含搬家、装修、整屋升级以及会在同一次进店中顺带比较音频产品的中高客单消费者。",
                opportunityTitle: "把试听体验和家庭升级场景绑定，放大高客单转化",
                opportunityBody: "继续用 POP + 主销英雄 SKU 稳住露出，同时把 OpenDots 2、T921 这类新品放在“家居升级 / TV / 健身 / 户外”联动场景里做转化，争取从单纯陈列升级到体验成交。七八月还有一个 POP 快闪店机会，目前由 Kristen 策划、Raymond 辅助推进。"
            },
            rcw: {
                profileTitle: "西部区域家居电器零售商，偏家庭计划型购买",
                profileBody: "RC Willey 官网定位是服务西部市场的 home furnishings、appliances 与 electronics 零售商；结合 PPT 中“门店空间大、顾客停留久、独立店销售表现不错”的信息，这类店更像 destination store，而不是纯快进快出的电子卖场。",
                persona: "更偏郊区家庭、换新型用户和成套购买用户，重视 financing、门店服务和可视化陈列，愿意在店内做横向比较后再下单。",
                opportunityTitle: "围绕西部家庭客群做稳定扩店和结构升级",
                opportunityBody: "RCW 的增长空间在于把现有成功门店打法复制到更多门店与更多核心 SKU 上，用更清晰的 POP、培训和价格带分层，把 Shokz 从补充型音频 SKU 变成被主动推荐的功能型耳机。"
            },
            abt: {
                profileTitle: "单店高服务模型，专业咨询和高 ASP 更关键",
                profileBody: "Abt 官方公开信息强调其 Glenview 单店、90+ 年历史、公开 tour、培训体系与 award-winning customer service；它不是广撒网渠道，更像高服务密度的精品大店。",
                persona: "以研究型、信任导购型、愿意为服务和专业建议买单的 Chicagoland 用户为主，对产品故事、功能差异和售后体验更敏感。",
                opportunityTitle: "走精品化和顾问式成交，优先做高质量产出",
                opportunityBody: "Abt 更适合做少而精的结构优化，集中火力推高 ASP 主力款和具备明确卖点的新型号，通过导购教育、场景话术和本地化活动，拉高单点产出而不是盲目铺量。"
            },
            bsm: {
                profileTitle: "东南部价格驱动大卖场，促销和现货感知更强",
                profileBody: "BrandsMart USA 官网强调 huge selection、Audio & Headphones、110% Price Match、会员与 financing；结合 PPT 中“人流量高、竞品多、靠近 Costco / Sam / Best Buy”的判断，这一渠道更偏高流量、强促销、强对比购物环境。",
                persona: "价格敏感型、促销响应型和家庭换新型用户占比较高，重视即时可得、价格优势和清晰的卖点提示，容易在大促节点集中释放销量。",
                opportunityTitle: "用促销节点和主力价位带放大爆发力",
                opportunityBody: "BSM 需要把 POP 更新、主推 SKU、价格带和大促节奏绑在一起，重点做高流量月份的主力款曝光与补货，争取把自然陈列流量转成更高的销量和销售额贡献。"
            }
        },
        promoCalendar: [
            { month: "6月", note: "POP 更新切换，跟踪 E320 / T921 替换后的首轮 sell-through" },
            { month: "7月", note: "渠道对齐与跟进，复盘 POP 执行、补货和门店反馈" },
            { month: "8月", note: "关注 NFM POP 快闪店机会，以及 OpenDots Air / OpenRun 2 的上市与 POP 更新窗口" },
            { month: "9月+", note: "结合节庆和下半年促销节奏，继续补充重点活动节点" }
        ],
        executionModules: [
            {
                kicker: "执行节奏",
                title: "促销日历",
                description: "后续整理 NATM 渠道重点促销档期、节点安排和资源节奏。",
                body: "暂时空着"
            },
            {
                kicker: "营销推进",
                title: "营销事件时间线",
                description: "后续沉淀营销事件、门店动作、活动上线和复盘时间线。",
                body: "暂时空着"
            },
            {
                kicker: "项目管理",
                title: "当前进度",
                description: "后续更新各项重点事项的状态、负责人和下一步动作。",
                body: "暂时空着"
            },
            {
                kicker: "协同记录",
                title: "会议历史",
                description: "后续补充渠道会议纪要、关键决议和待跟进事项。",
                body: "暂时空着"
            }
        ],
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
        fullProductMixPanel: document.getElementById("fullProductMixPanel"),
        popPlanPanel: document.getElementById("popPlanPanel"),
        launchPlanPanel: document.getElementById("launchPlanPanel"),
        businessSupportGrid: document.getElementById("businessSupportGrid"),
        channelOverviewGrid: document.getElementById("channelOverviewGrid"),
        channelStrategyGrid: document.getElementById("channelStrategyGrid"),
        natmSummaryGrid: document.getElementById("natmSummaryGrid"),
        channelSummaryBody: document.querySelector("#channelSummaryTable tbody"),
        channelTabs: document.getElementById("channelTabs"),
        kpiGrid: document.getElementById("kpiGrid"),
        insightPanel: document.getElementById("insightPanel"),
        yearSummaryGrid: document.getElementById("yearSummaryGrid"),
        executionGrid: document.getElementById("executionGrid"),
        promoCalendarPanel: document.getElementById("promoCalendarPanel"),
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
                remainingInventory: null,
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

    function getChannelConfig(channelKey) {
        return CONFIG.channels.find(channel => channel.key === channelKey) || {
            key: channelKey,
            label: String(channelKey || "").toUpperCase(),
            accent: CONFIG.colors.qty
        };
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

    function renderMixBadge(isPopSku) {
        return "<span class=\"mix-badge" + (isPopSku ? "" : " out") + "\">" + (isPopSku ? "POP" : "非POP") + "</span>";
    }

    function renderProductMixTable(items, popSkuSet) {
        if (!items.length) {
            return "<div class=\"table-note\">2025-09 至今暂无卖出记录</div>";
        }
        return [
            "<div class=\"table-wrap profile-table-wrap\">",
            "<table class=\"profile-table\">",
            "<thead>",
            "<tr>",
            "<th>SKU</th>",
            "<th>属性</th>",
            "<th>销量</th>",
            "<th>销售额</th>",
            "<th>销售额占比</th>",
            "</tr>",
            "</thead>",
            "<tbody>",
            items.map(item => {
                const isPopSku = popSkuSet.has(item.sku);
                const shareText = item.salesShare !== null ? formatPercent(item.salesShare, true) : "—";
                return [
                    "<tr>",
                    "<td><strong>" + escapeHtml(item.sku) + "</strong></td>",
                    "<td>" + renderMixBadge(isPopSku) + "</td>",
                    "<td>" + formatNumber(item.qty) + "</td>",
                    "<td>" + formatCurrency(item.sales) + "</td>",
                    "<td>" + shareText + "</td>",
                    "</tr>"
                ].join("");
            }).join(""),
            "</tbody>",
            "</table>",
            "</div>"
        ].join("");
    }

    function renderBusinessMetricCard(label, value) {
        return [
            "<article class=\"business-metric-card\">",
            "<p class=\"business-metric-label\">" + escapeHtml(label) + "</p>",
            "<p class=\"business-metric-value\">" + escapeHtml(value) + "</p>",
            "</article>"
        ].join("");
    }

    function renderRankIndex(index) {
        return "<span class=\"rank-index\">" + String(index) + "</span>";
    }

    function renderBusinessRankTable(items, popSkuSet) {
        const topItems = (items || []).slice(0, 5);
        if (!topItems.length) {
            return "<div class=\"table-note\">暂无卖出SKU结构记录</div>";
        }
        return [
            "<div class=\"mini-rank-wrap\">",
            "<table class=\"mini-rank-table\">",
            "<thead>",
            "<tr>",
            "<th>排名</th>",
            "<th>SKU</th>",
            "<th>属性</th>",
            "<th>销量占比</th>",
            "<th>销售额占比</th>",
            "</tr>",
            "</thead>",
            "<tbody>",
            topItems.map((item, index) => {
                const qtyShare = item.qtyShare !== null ? formatPercent(item.qtyShare, true) : "—";
                const salesShare = item.salesShare !== null ? formatPercent(item.salesShare, true) : "—";
                return [
                    "<tr>",
                    "<td>" + renderRankIndex(index + 1) + "</td>",
                    "<td><strong>" + escapeHtml(item.sku) + "</strong></td>",
                    "<td>" + renderMixBadge(popSkuSet.has(item.sku)) + "</td>",
                    "<td>" + qtyShare + "</td>",
                    "<td>" + salesShare + "</td>",
                    "</tr>"
                ].join("");
            }).join(""),
            "</tbody>",
            "</table>",
            "</div>"
        ].join("");
    }

    function renderFullProductMixPanel() {
        if (!els.fullProductMixPanel) {
            return;
        }

        els.fullProductMixPanel.innerHTML = [
            "<article class=\"info-card\">",
            "<div class=\"info-card-head\">",
            "<div>",
            "<p class=\"info-kicker\">完整结构</p>",
            "<h3 class=\"info-card-title\">完整产品结构</h3>",
            "<p class=\"info-card-copy\">按四个 NATM 渠道展示 2025-09 至今所有实际卖出的 SKU 结构，保留完整销量、销售额与占比，方便后续继续分析结构变化。</p>",
            "</div>",
            "<span class=\"info-badge\">Full Mix</span>",
            "</div>",
            "<div class=\"table-wrap\">",
            "<table class=\"full-mix-table\">",
            "<thead>",
            "<tr>",
            "<th>渠道</th>",
            "<th>排名</th>",
            "<th>SKU</th>",
            "<th>属性</th>",
            "<th>销量</th>",
            "<th>销量占比</th>",
            "<th>销售额</th>",
            "<th>销售额占比</th>",
            "</tr>",
            "</thead>",
            "<tbody>",
            CONFIG.channels.map(channel => {
                const dashboard = state.dashboards[channel.key];
                const profile = CONFIG.channelProfiles[channel.key];
                const popSkuSet = new Set(profile.popSkus);
                const items = dashboard ? dashboard.productMixSinceStart.items : [];
                const channelLabel = profile.displayName || channel.label;

                if (!items.length) {
                    return [
                        "<tr class=\"mix-group-row\"><td colspan=\"8\">" + escapeHtml(channelLabel) + "</td></tr>",
                        "<tr><td>" + escapeHtml(channelLabel) + "</td><td colspan=\"7\">暂无卖出SKU结构记录</td></tr>"
                    ].join("");
                }

                return [
                    "<tr class=\"mix-group-row\"><td colspan=\"8\">" + escapeHtml(channelLabel) + "</td></tr>",
                    items.map((item, index) => {
                        const qtyShare = item.qtyShare !== null ? formatPercent(item.qtyShare, true) : "—";
                        const salesShare = item.salesShare !== null ? formatPercent(item.salesShare, true) : "—";
                        return [
                            "<tr>",
                            "<td>" + escapeHtml(channelLabel) + "</td>",
                            "<td>" + renderRankIndex(index + 1) + "</td>",
                            "<td><strong>" + escapeHtml(item.sku) + "</strong></td>",
                            "<td>" + renderMixBadge(popSkuSet.has(item.sku)) + "</td>",
                            "<td>" + formatNumber(item.qty) + "</td>",
                            "<td>" + qtyShare + "</td>",
                            "<td>" + formatCurrency(item.sales) + "</td>",
                            "<td>" + salesShare + "</td>",
                            "</tr>"
                        ].join("");
                    }).join("")
                ].join("");
            }).join(""),
            "</tbody>",
            "</table>",
            "</div>",
            "</article>"
        ].join("");
    }

    function renderInfoRow(row) {
        const channel = row.channelKey ? getChannelConfig(row.channelKey) : null;
        return [
            "<div class=\"note-row\"" + (channel ? " style=\"" + channelStyle(channel) + "\"" : "") + ">",
            "<div class=\"note-row-head\">",
            channel ? "<span class=\"profile-tag\">" + escapeHtml(channel.label) + "</span>" : "",
            row.title ? "<strong>" + escapeHtml(row.title) + "</strong>" : "",
            "</div>",
            row.metrics && row.metrics.length ? [
                "<div class=\"note-stat-grid\">",
                row.metrics.map(metric => {
                    return [
                        "<div class=\"note-stat\">",
                        "<span>" + escapeHtml(metric.label) + "</span>",
                        "<strong>" + escapeHtml(metric.value) + "</strong>",
                        "</div>"
                    ].join("");
                }).join(""),
                "</div>"
            ].join("") : "",
            row.text ? "<p class=\"note-row-copy\">" + escapeHtml(row.text) + "</p>" : "",
            row.meta ? "<p class=\"note-row-meta\">" + escapeHtml(row.meta) + "</p>" : "",
            "</div>"
        ].join("");
    }

    function renderInfoCard(module) {
        const badge = module.badge || (module.rows && module.rows.length ? "已更新" : "待补充");
        return [
            "<article class=\"info-card\">",
            "<div class=\"info-card-head\">",
            "<div>",
            module.kicker ? "<p class=\"info-kicker\">" + escapeHtml(module.kicker) + "</p>" : "",
            "<h3 class=\"info-card-title\">" + escapeHtml(module.title) + "</h3>",
            module.description ? "<p class=\"info-card-copy\">" + escapeHtml(module.description) + "</p>" : "",
            "</div>",
            "<span class=\"info-badge\">" + escapeHtml(badge) + "</span>",
            "</div>",
            module.rows && module.rows.length ? [
                "<div class=\"info-row-list\">",
                module.rows.map(renderInfoRow).join(""),
                "</div>"
            ].join("") : "<p class=\"placeholder-copy\">" + escapeHtml(module.body || "暂时空着") + "</p>",
            module.footnote ? "<p class=\"info-footnote\">" + escapeHtml(module.footnote) + "</p>" : "",
            "</article>"
        ].join("");
    }

    function renderPopPlanPanel() {
        if (!els.popPlanPanel) {
            return;
        }
        const plan = CONFIG.popNextStepPlan;
        els.popPlanPanel.innerHTML = [
            "<article class=\"info-card\">",
            "<div class=\"info-card-head\">",
            "<div>",
            "<p class=\"info-kicker\">POP 更新</p>",
            "<h3 class=\"info-card-title\">POP下一步更新计划</h3>",
            "<p class=\"info-card-copy\">先记录 6/4 已确认的 POP 替换动作和目标 POP 组合，八月计划后续继续补充。</p>",
            "</div>",
            "<span class=\"info-badge\">" + escapeHtml(plan.effectiveDate) + "</span>",
            "</div>",
            "<div class=\"info-section\">",
            "<p class=\"info-label\">6/4</p>",
            "<div class=\"table-wrap\">",
            "<table class=\"compact-table\">",
            "<thead><tr><th>渠道</th><th>当前SKU</th><th>更新后SKU</th></tr></thead>",
            "<tbody>",
            plan.replacements.map(item => {
                const channel = getChannelConfig(item.channelKey);
                return [
                    "<tr>",
                    "<td><strong>" + escapeHtml(channel.label) + "</strong></td>",
                    "<td>" + escapeHtml(item.currentSku) + "</td>",
                    "<td>" + escapeHtml(item.nextSku) + "</td>",
                    "</tr>"
                ].join("");
            }).join(""),
            "</tbody>",
            "</table>",
            "</div>",
            "</div>",
            "<div class=\"info-section\">",
            "<p class=\"info-label\">POP产品组合</p>",
            "<div class=\"plan-mix-grid\">",
            plan.targetCombos.map(group => {
                const channel = getChannelConfig(group.channelKey);
                return [
                    "<article class=\"plan-mini-card\" style=\"" + channelStyle(channel) + "\">",
                    "<h4 class=\"plan-mini-title\">" + escapeHtml(channel.label) + "</h4>",
                    "<div class=\"profile-tag-list\">" + renderProfileTags(group.skus) + "</div>",
                    "</article>"
                ].join("");
            }).join(""),
            "</div>",
            "</div>",
            "<div class=\"info-section\">",
            "<p class=\"info-label\">八月底计划</p>",
            plan.augustReplacements && plan.augustReplacements.length ? [
                "<div class=\"table-wrap\">",
                "<table class=\"compact-table\">",
                "<thead><tr><th>渠道</th><th>当前SKU</th><th>更新后SKU</th></tr></thead>",
                "<tbody>",
                plan.augustReplacements.map(item => {
                    const channel = getChannelConfig(item.channelKey);
                    return [
                        "<tr>",
                        "<td><strong>" + escapeHtml(channel.label) + "</strong></td>",
                        "<td>" + escapeHtml(item.currentSku) + "</td>",
                        "<td>" + escapeHtml(item.nextSku) + "</td>",
                        "</tr>"
                    ].join("");
                }).join(""),
                "</tbody>",
                "</table>",
                "</div>"
            ].join("") : "<p class=\"placeholder-copy\">暂时空着</p>",
            "</div>",
            "</article>"
        ].join("");
    }

    function renderLaunchPlanPanel() {
        if (els.launchPlanPanel) {
            els.launchPlanPanel.innerHTML = renderInfoCard(CONFIG.businessPlaceholders.launchPlan);
        }
    }

    function renderBusinessSupportGrid() {
        if (!els.businessSupportGrid) {
            return;
        }
        els.businessSupportGrid.innerHTML = [
            CONFIG.businessPlaceholders.costMargin,
            CONFIG.businessPlaceholders.sampling
        ].map(renderInfoCard).join("");
    }

    function buildChannelStrategyModules() {
        return [
            {
                kicker: "渠道洞察",
                title: "渠道特点及用户画像",
                description: "结合 NATM 介绍 PPT 与四家官网公开信息整理，用户画像为基于店型、品类与服务模型的业务推断。",
                badge: "Research",
                rows: CONFIG.channels.map(channel => {
                    const research = CONFIG.channelResearch[channel.key];
                    return {
                        channelKey: channel.key,
                        title: research.profileTitle,
                        text: research.profileBody,
                        meta: "用户画像：" + research.persona
                    };
                }),
                footnote: "公开信息主要来自 NFM、RC Willey、Abt、BrandsMart USA 官网；用户画像为结合 PPT 与店型结构的推断。"
            },
            {
                kicker: "增长方向",
                title: "核心机会",
                description: "按渠道定位去拆 2026 年最值得优先推进的增长动作，帮助后续把 POP、新品与营销资源放到正确的地方。",
                badge: "Action",
                rows: CONFIG.channels.map(channel => {
                    const research = CONFIG.channelResearch[channel.key];
                    return {
                        channelKey: channel.key,
                        title: research.opportunityTitle,
                        text: research.opportunityBody
                    };
                })
            },
            {
                kicker: "年度管理",
                title: "今年目标",
                description: "四个渠道统一按“至少 +30%，标准 +40%，挑战 +50%+”管理，并锁定 2026 的销量与销售额目标值。",
                badge: "2026",
                rows: CONFIG.channels.map(channel => {
                    const target = CONFIG.channelTargets[channel.key];
                    const dashboard = state.dashboards[channel.key];
                    const qtyGrowth = dashboard ? calculateGrowth(target.qty, dashboard.yearlyTotals[2025].qty) : null;
                    const salesGrowth = dashboard ? calculateGrowth(target.sales, dashboard.yearlyTotals[2025].sales) : null;
                    return {
                        channelKey: channel.key,
                        title: CONFIG.growthGoalText,
                        metrics: [
                            { label: "目标销量", value: formatNumber(target.qty) },
                            { label: "目标销售额", value: formatCurrency(target.sales) }
                        ],
                        text: "目标口径统一按至少 +30%、标准 +40%、挑战 +50%+推进。",
                        meta: dashboard
                            ? "相对 2025 实际：销量 " + formatPercent(qtyGrowth, false) + "，销售额 " + formatPercent(salesGrowth, false) + "。"
                            : "数据加载后自动回填相对 2025 的目标同比。"
                    };
                })
            }
        ];
    }

    function renderChannelStrategyGrid() {
        if (!els.channelStrategyGrid) {
            return;
        }
        els.channelStrategyGrid.innerHTML = buildChannelStrategyModules().map(renderInfoCard).join("");
    }

    function renderPromoCalendarPanel() {
        if (!els.promoCalendarPanel) {
            return;
        }
        els.promoCalendarPanel.innerHTML = [
            "<div class=\"calendar-strip\">",
            CONFIG.promoCalendar.map(item => {
                return "<div class=\"calendar-chip\"><strong>" + escapeHtml(item.month) + "</strong><span>" + escapeHtml(item.note) + "</span></div>";
            }).join(""),
            "</div>",
            "<p class=\"promo-compare-copy\">建议把这一块与下方月度总览一起看，重点观察促销节点前后 1-2 个月的销量、销售额与 Top SKU 变化。</p>"
        ].join("");
    }

    function sumChannelMetrics(resolver) {
        return CONFIG.channels.reduce((totals, channel) => {
            const metrics = resolver(channel) || {};
            totals.qty += Number(metrics.qty) || 0;
            totals.sales += Number(metrics.sales) || 0;
            return totals;
        }, { qty: 0, sales: 0 });
    }

    function buildNatmSummary() {
        const actual2024 = sumChannelMetrics(channel => {
            const dashboard = state.dashboards[channel.key];
            return dashboard ? dashboard.yearlyTotals[2024] : null;
        });
        const actual2025 = sumChannelMetrics(channel => {
            const dashboard = state.dashboards[channel.key];
            return dashboard ? dashboard.yearlyTotals[2025] : null;
        });
        const target2026 = sumChannelMetrics(channel => CONFIG.channelTargets[channel.key]);

        return {
            actual2024: actual2024,
            actual2025: actual2025,
            target2026: target2026,
            qtyYoY2025: calculateGrowth(actual2025.qty, actual2024.qty),
            salesYoY2025: calculateGrowth(actual2025.sales, actual2024.sales),
            qtyTargetGrowth: calculateGrowth(target2026.qty, actual2025.qty),
            salesTargetGrowth: calculateGrowth(target2026.sales, actual2025.sales)
        };
    }

    function renderNatmSummaryPanel() {
        if (!els.natmSummaryGrid) {
            return;
        }
        if (!Object.keys(state.dashboards).length) {
            els.natmSummaryGrid.innerHTML = "";
            return;
        }

        const summary = buildNatmSummary();
        const cards = [
            {
                eyebrow: "2024 Actual",
                title: "2024 NATM 总盘子",
                pairs: [
                    { label: "总销量", value: formatNumber(summary.actual2024.qty) },
                    { label: "总销售额", value: formatCurrency(summary.actual2024.sales) }
                ],
                meta: ["四个渠道全年实际表现"]
            },
            {
                eyebrow: "2025 Actual",
                title: "2025 NATM 总盘子",
                pairs: [
                    { label: "总销量", value: formatNumber(summary.actual2025.qty) },
                    { label: "总销售额", value: formatCurrency(summary.actual2025.sales) }
                ],
                meta: [
                    "销量同比 " + formatPercent(summary.qtyYoY2025, false),
                    "销售额同比 " + formatPercent(summary.salesYoY2025, false)
                ]
            },
            {
                eyebrow: "2026 Target",
                title: "2026 NATM 目标",
                pairs: [
                    { label: "目标销量", value: formatNumber(summary.target2026.qty) },
                    { label: "目标销售额", value: formatCurrency(summary.target2026.sales) }
                ],
                meta: [
                    "目标销量同比 " + formatPercent(summary.qtyTargetGrowth, false) + " vs 2025",
                    "目标销售额同比 " + formatPercent(summary.salesTargetGrowth, false) + " vs 2025"
                ]
            },
            {
                eyebrow: "Goal Range",
                title: "2026 目标口径",
                pairs: [
                    { label: "统一原则", value: "至少 +30%" },
                    { label: "标准 / 挑战", value: "+40% / +50%+" }
                ],
                meta: [CONFIG.growthGoalText]
            }
        ];

        els.natmSummaryGrid.innerHTML = cards.map(card => {
            return [
                "<article class=\"summary-total-card\">",
                "<p class=\"summary-total-eyebrow\">" + escapeHtml(card.eyebrow) + "</p>",
                "<h4 class=\"summary-total-title\">" + escapeHtml(card.title) + "</h4>",
                "<div class=\"summary-total-pairs\">",
                card.pairs.map(pair => {
                    return [
                        "<div class=\"summary-total-pair\">",
                        "<span>" + escapeHtml(pair.label) + "</span>",
                        "<strong>" + escapeHtml(pair.value) + "</strong>",
                        "</div>"
                    ].join("");
                }).join(""),
                "</div>",
                "<div class=\"summary-total-meta\">",
                card.meta.map(item => "<span>" + escapeHtml(item) + "</span>").join(""),
                "</div>",
                "</article>"
            ].join("");
        }).join("");
    }

    function renderExecutionGrid() {
        if (!els.executionGrid) {
            return;
        }
        const marketingTimeline = [
            { date: "Now", title: "POP 更新计划已建", text: "已把 6/4 替换动作与目标 POP 组合整理进页面。" },
            { date: "Next", title: "营销事件待补充", text: "后续补充每个渠道的营销动作、节奏与资源需求。" },
            { date: "Q3", title: "八月计划待确认", text: "预留给新品节奏、活动节点和后续 POP 调整。" }
        ];
        const progressTimeline = [
            { date: "01", title: "商务信息结构已完成", text: "商务对接、出货价、POP 情况和产品结构已形成统一框架。" },
            { date: "02", title: "渠道分析框架已预留", text: "渠道特点、核心机会和今年目标模块可继续逐项补充。" },
            { date: "03", title: "后续补内容即可", text: "当前以结构为主，后续直接逐块填入业务内容即可。" }
        ];
        const meetingTimeline = [
            { date: "TBD", title: "会议历史待补充", text: "后续可按时间沉淀会议纪要、决议和责任事项。" }
        ];

        els.executionGrid.innerHTML = [
            "<article class=\"info-card\">",
            "<div class=\"info-card-head\"><div><p class=\"info-kicker\">营销推进</p><h3 class=\"info-card-title\">营销事件时间线</h3><p class=\"info-card-copy\">改成时间线视图，后续加活动时会比普通卡片更直观。</p></div><span class=\"info-badge\">Timeline</span></div>",
            "<div class=\"timeline-list\">",
            marketingTimeline.map(item => {
                return "<div class=\"timeline-item\"><span class=\"timeline-date\">" + escapeHtml(item.date) + "</span><div class=\"timeline-body\"><strong>" + escapeHtml(item.title) + "</strong><span>" + escapeHtml(item.text) + "</span></div></div>";
            }).join(""),
            "</div>",
            "</article>",
            "<article class=\"info-card\">",
            "<div class=\"info-card-head\"><div><p class=\"info-kicker\">推进状态</p><h3 class=\"info-card-title\">当前进度</h3><p class=\"info-card-copy\">用步骤式时间线展示现在已经完成的结构和下一步待推进内容。</p></div><span class=\"info-badge\">Progress</span></div>",
            "<div class=\"timeline-list\">",
            progressTimeline.map(item => {
                return "<div class=\"timeline-item\"><span class=\"timeline-date\">" + escapeHtml(item.date) + "</span><div class=\"timeline-body\"><strong>" + escapeHtml(item.title) + "</strong><span>" + escapeHtml(item.text) + "</span></div></div>";
            }).join(""),
            "</div>",
            "</article>",
            "<article class=\"info-card\">",
            "<div class=\"info-card-head\"><div><p class=\"info-kicker\">会议沉淀</p><h3 class=\"info-card-title\">会议历史</h3><p class=\"info-card-copy\">先预留会议记录入口，后续可以直接追加会议纪要与 follow-up。</p></div><span class=\"info-badge\">Notes</span></div>",
            "<div class=\"timeline-list\">",
            meetingTimeline.map(item => {
                return "<div class=\"timeline-item\"><span class=\"timeline-date\">" + escapeHtml(item.date) + "</span><div class=\"timeline-body\"><strong>" + escapeHtml(item.title) + "</strong><span>" + escapeHtml(item.text) + "</span></div></div>";
            }).join(""),
            "</div>",
            "</article>"
        ].join("");
    }

    function renderStaticSections() {
        renderFullProductMixPanel();
        renderPopPlanPanel();
        renderLaunchPlanPanel();
        renderBusinessSupportGrid();
        renderChannelStrategyGrid();
        renderExecutionGrid();
        renderPromoCalendarPanel();
    }

    function renderBusinessOverview() {
        els.businessOverviewGrid.innerHTML = [
            "<table class=\"business-matrix\">",
            "<colgroup>",
            "<col class=\"col-channel\">",
            "<col class=\"col-business\">",
            "<col class=\"col-price\">",
            "<col class=\"col-pop\">",
            "<col class=\"col-overview\">",
            "</colgroup>",
            "<thead>",
            "<tr>",
            "<th>渠道</th>",
            "<th>商务对接</th>",
            "<th>出货价</th>",
            "<th>当前POP产品情况</th>",
            "<th>产品结构概览</th>",
            "</tr>",
            "</thead>",
            "<tbody>",
            CONFIG.channels.map(channel => {
            const profile = CONFIG.channelProfiles[channel.key];
            const dashboard = state.dashboards[channel.key];
            const productMix = dashboard.productMixSinceStart;
            const popSkuSet = new Set(profile.popSkus);
            const soldSkuSet = new Set(productMix.items.map(item => item.sku));
            const popInSales = profile.popSkus.filter(sku => soldSkuSet.has(sku));
            const popCoverageText = profile.popSkus.length ? ("POP 覆盖 " + popInSales.length + "/" + profile.popSkus.length) : "当前无 POP 组合";

            return [
                "<tr style=\"" + channelStyle(channel) + "\">",
                "<td>",
                "<div class=\"business-channel-cell\">",
                "<div><p class=\"business-channel-name\">" + escapeHtml(profile.displayName || channel.label) + "</p></div>",
                "<span class=\"channel-chip\">NATM</span>",
                "<p class=\"business-channel-meta\">最新月份 " + escapeHtml(dashboard.latestMonthKey) + "<br>2026 YTD " + formatCurrency(dashboard.samePeriodByYear[2026].sales) + "</p>",
                "</div>",
                "</td>",
                "<td>",
                "<div class=\"business-cell\">",
                "<div class=\"profile-tag-list\">" + renderProfileTags(profile.businessTags) + "</div>",
                "<p class=\"profile-copy\"><strong>Buyer：</strong>" + escapeHtml(profile.buyer) + "</p>",
                "<p class=\"profile-copy\"><strong>Account：</strong>" + escapeHtml(profile.account) + "</p>",
                "<p class=\"profile-copy\"><strong>Setup：</strong>" + escapeHtml(profile.setup) + "</p>",
                "</div>",
                "</td>",
                "<td>",
                "<div class=\"business-cell compact\">",
                renderBusinessMetricCard("2025 出货价", profile.shipPrice2025),
                renderBusinessMetricCard("2026 出货价", profile.shipPrice2026),
                "</div>",
                "</td>",
                "<td>",
                "<div class=\"business-cell\">",
                "<p class=\"profile-copy\"><strong>POP尺寸：</strong>" + escapeHtml(profile.popSize) + "</p>",
                "<div class=\"profile-tag-list\">" + renderProfileTags(profile.popSkus) + "</div>",
                "</div>",
                "</td>",
                "<td>",
                "<div class=\"business-cell\">",
                "<div class=\"business-metric-grid\">",
                renderBusinessMetricCard("卖出SKU", String(productMix.items.length)),
                renderBusinessMetricCard("POP覆盖", profile.popSkus.length ? (String(popInSales.length) + "/" + String(profile.popSkus.length)) : "无POP"),
                renderBusinessMetricCard("累计销量", formatNumber(productMix.totalQty)),
                renderBusinessMetricCard("销售额", formatCurrency(productMix.totalSales)),
                "</div>",
                "<p class=\"profile-mix-note\">统计口径：2025-09 至 " + escapeHtml(dashboard.latestMonthKey) + "。 " + popCoverageText + "</p>",
                renderBusinessRankTable(productMix.items, popSkuSet),
                "</div>",
                "</td>",
                "</tr>"
            ].join("");
        }).join(""),
            "</tbody>",
            "</table>"
        ].join("");
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

        renderNatmSummaryPanel();

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
                "<td><span class=\"delta neutral\">待补充</span></td>",
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
        renderStaticSections();
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
        if (els.natmSummaryGrid) {
            els.natmSummaryGrid.innerHTML = "";
        }
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

    try {
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
    renderStaticSections();
    loadData();
    } catch (error) {
        console.error("NATM init failed", error);
        if (els.statusBadge) {
            els.statusBadge.textContent = "初始化失败";
            els.statusBadge.className = "status-badge error";
        }
        if (els.lastUpdated) {
            els.lastUpdated.textContent = "上次刷新：初始化失败";
        }
        if (els.businessOverviewGrid) {
            els.businessOverviewGrid.innerHTML = "<div class=\"error-state\">页面初始化失败：" + escapeHtml(error.message || "未知错误") + "</div>";
        }
    }
})();
