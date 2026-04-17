(function() {
    const WORKBOOK_ID = "1-3zBJZvum5M7yLED_ZUbpMugmkWZeTCFZVAvs1F4QqE";
    const CHANNEL_PRODUCT_PRICING = window.CHANNEL_PRODUCT_PRICING || {};
    const PROMO_CALENDARS = {
        nfm: {
            source: "20260414 Shokz Q2 Promotion Calendar-NFM.xlsx",
            rows: [
                { month: "April", date: "4.29-5.10", product: "OpenRun Pro2", discount: 0.2, priceAsIs: 179.95, priceToBe: 139.95, saving: 40, rebate: 22.32, models: ["S820-ST-BK", "S820-ST-OR", "S821-MN-BK", "S821-MN-OR"] },
                { month: "April", date: "4.29-5.10", product: "OpenRun", discount: 0.3, priceAsIs: 129.95, priceToBe: 89.95, saving: 40, rebate: 22.32, models: ["S803-MN-BK", "S803-MN-BL", "S803-ST-BK", "S803-ST-BL", "S803-ST-RD"] },
                { month: "April", date: "4.29-5.10", product: "OpenFit 2+", discount: 0.15, priceAsIs: 199.95, priceToBe: 169.95, saving: 30, rebate: 16.74, models: ["T921-ST-BK", "T921-ST-GY"] },
                { month: "April", date: "4.29-5.10", product: "OpenDots ONE", discount: 0.2, priceAsIs: 199.95, priceToBe: 159.95, saving: 40, rebate: 22.32, models: ["E310-ST-BK", "E310-ST-BG"] },
                { month: "May", date: "5.18-5.25", product: "OpenFit Air", discount: 0.3, priceAsIs: 119.95, priceToBe: 79.95, saving: 40, rebate: 22.32, models: ["T511-ST-BK", "T511-ST-WT", "T511-ST-PK"] },
                { month: "May", date: "5.18-5.25", product: "OpenMove", discount: 0.3, priceAsIs: 79.95, priceToBe: 54.95, saving: 25, rebate: 13.95, models: ["S661-ST-BL", "S661-ST-GY", "S661-ST-PK", "S661-ST-WT"] }
            ]
        },
        rcw: {
            source: "20260414 Shokz Q2 Promotion Calendar-RCW.xlsx",
            rows: [
                { month: "April", date: "4.29-5.10", product: "OpenRun Pro2", discount: 0.2, priceAsIs: 179.95, priceToBe: 139.95, saving: 40, rebate: 24, models: ["S820-ST-BK", "S820-ST-OR", "S821-MN-BK", "S821-MN-OR"] },
                { month: "April", date: "4.29-5.10", product: "OpenRun", discount: 0.3, priceAsIs: 129.95, priceToBe: 89.95, saving: 40, rebate: 24, models: ["S803-MN-BK", "S803-MN-BL", "S803-ST-BK", "S803-ST-BL", "S803-ST-RD"] },
                { month: "April", date: "4.29-5.10", product: "OpenFit 2+", discount: 0.15, priceAsIs: 199.95, priceToBe: 169.95, saving: 30, rebate: 18, models: ["T921-ST-BK", "T921-ST-GY"] },
                { month: "April", date: "4.29-5.10", product: "OpenDots ONE", discount: 0.2, priceAsIs: 199.95, priceToBe: 159.95, saving: 40, rebate: 24, models: ["E310-ST-BK", "E310-ST-BG"] },
                { month: "May", date: "5.18-5.25", product: "OpenFit Air", discount: 0.3, priceAsIs: 119.95, priceToBe: 79.95, saving: 40, rebate: 24, models: ["T511-ST-BK", "T511-ST-WT", "T511-ST-PK"] },
                { month: "May", date: "5.18-5.25", product: "OpenMove", discount: 0.3, priceAsIs: 79.95, priceToBe: 54.95, saving: 25, rebate: 15, models: ["S661-ST-BL", "S661-ST-GY", "S661-ST-PK", "S661-ST-WT"] }
            ]
        },
        abt: {
            source: "20260414 Shokz Q2 Promotion Calendar-Abt.xlsx",
            rows: [
                { month: "April", date: "4.29-5.10", product: "OpenRun Pro2", discount: 0.2, priceAsIs: 179.95, priceToBe: 139.95, saving: 40, rebate: 24, models: ["S820-ST-BK", "S820-ST-OR", "S821-MN-BK", "S821-MN-OR"] },
                { month: "April", date: "4.29-5.10", product: "OpenRun", discount: 0.3, priceAsIs: 129.95, priceToBe: 89.95, saving: 40, rebate: 24, models: ["S803-MN-BK", "S803-MN-BL", "S803-ST-BK", "S803-ST-BL", "S803-ST-RD"] },
                { month: "April", date: "4.29-5.10", product: "OpenFit 2+", discount: 0.15, priceAsIs: 199.95, priceToBe: 169.95, saving: 30, rebate: 18, models: ["T921-ST-BK", "T921-ST-GY"] },
                { month: "April", date: "4.29-5.10", product: "OpenDots ONE", discount: 0.2, priceAsIs: 199.95, priceToBe: 159.95, saving: 40, rebate: 24, models: ["E310-ST-BK", "E310-ST-BG"] },
                { month: "May", date: "5.18-5.25", product: "OpenFit Air", discount: 0.3, priceAsIs: 119.95, priceToBe: 79.95, saving: 40, rebate: 24, models: ["T511-ST-BK", "T511-ST-WT", "T511-ST-PK"] },
                { month: "May", date: "5.18-5.25", product: "OpenMove", discount: 0.3, priceAsIs: 79.95, priceToBe: 54.95, saving: 25, rebate: 15, models: ["S661-ST-BL", "S661-ST-GY", "S661-ST-PK", "S661-ST-WT"] }
            ]
        },
        bsm: {
            source: "20260411 Shokz Q2 Promotion Calendar-BSM.xlsx",
            rows: [
                { month: "Jan", date: "1.1-1.14", product: "OpenRun Pro", discount: 0.3, priceAsIs: 159.95, priceToBe: 109.95, saving: 50, rebate: 30, models: ["S810STBK", "S810STBL", "S810STPK", "S811MNBG", "S811MNBK"] },
                { month: "Jan", date: "1.1-1.14", product: "OpenSwim Pro", discount: 0.15, priceAsIs: 179.95, priceToBe: 149.95, saving: 30, rebate: 18, models: ["S710STRD", "S710STGY"] },
                { month: "Jan", date: "1.7-1.20", product: "OpenFit", discount: 0.15, priceAsIs: 159.95, priceToBe: 134.95, saving: 25, rebate: 15, models: ["T910STBG", "T910STBK"] },
                { month: "Jan", date: "1.7-1.20", product: "OpenFit2", discount: 0.15, priceAsIs: 179.95, priceToBe: 149.95, saving: 30, rebate: 18, models: ["T920STBG", "T920STBK"] },
                { month: "Feb", date: "1.26-2.3", product: "OpenMove", discount: 0, priceAsIs: 79.95, priceToBe: 49.95, saving: 30, rebate: 18, models: ["S661STBL", "S661STGY", "S661STPK", "S661STWT"] },
                { month: "Feb", date: "2.4-2.14", product: "OpenRun Pro2", discount: 0.15, priceAsIs: 179.95, priceToBe: 149.95, saving: 30, rebate: 18, models: ["S820STBK", "S820STOR", "S821MNBK", "S821MNOR"] },
                { month: "Feb", date: "2.4-2.14", product: "OpenFit Air", discount: 0.3, priceAsIs: 119.95, priceToBe: 79.95, saving: 40, rebate: 24, models: ["T511STBK", "T511STWT", "T511STPK"] },
                { month: "Feb", date: "2.4-2.14", product: "OpenSwim Pro", discount: 0.2, priceAsIs: 179.95, priceToBe: 139.95, saving: 40, rebate: 24, models: ["S710STRD", "S710STGY"] },
                { month: "Feb", date: "2.4-2.14", product: "OpenRun", discount: 0.3, priceAsIs: 129.95, priceToBe: 89.95, saving: 40, rebate: 24, models: ["S803MNBK", "S803MNBL", "S803STBK", "S803STBL", "S803STRD"] },
                { month: "Feb", date: "2.4-2.14", product: "OpenFit2", discount: 0.15, priceAsIs: 179.95, priceToBe: 149.95, saving: 30, rebate: 18, models: ["T920STBG", "T920STBK"] },
                { month: "Feb", date: "2.4-2.14", product: "OpenDots ONE", discount: 0.15, priceAsIs: 199.95, priceToBe: 169.95, saving: 30, rebate: 18, models: ["E310STBK", "E310STBG"] },
                { month: "Feb", date: "2.4-2.14", product: "OpenMove", discount: 0.3, priceAsIs: 79.95, priceToBe: 54.95, saving: 25, rebate: 15, models: ["S661STBL", "S661STGY", "S661STPK", "S661STWT"] },
                { month: "Mar", date: "3.4-3.17", product: "OpenRun Pro", discount: 0.2, priceAsIs: 159.95, priceToBe: 124.95, saving: 35, rebate: 21, models: ["S810STBK", "S810STBL", "S810STPK", "S811MNBG", "S811MNBK"] },
                { month: "Mar", date: "3.25-4.7", product: "OpenRun Pro2", discount: 0.2, priceAsIs: 179.95, priceToBe: 139.95, saving: 40, rebate: 24, models: ["S820STBK", "S820STOR", "S821MNBK", "S821MNOR"] },
                { month: "Mar", date: "3.25-4.7", product: "OpenFit Air", discount: 0.3, priceAsIs: 119.95, priceToBe: 79.95, saving: 40, rebate: 24, models: ["T511STBK", "T511STWT", "T511STPK"] },
                { month: "Mar", date: "3.25-4.7", product: "OpenSwim Pro", discount: 0.2, priceAsIs: 179.95, priceToBe: 139.95, saving: 40, rebate: 24, models: ["S710STRD", "S710STGY"] },
                { month: "Mar", date: "3.25-4.7", product: "OpenRun", discount: 0.3, priceAsIs: 129.95, priceToBe: 89.95, saving: 40, rebate: 24, models: ["S803MNBK", "S803MNBL", "S803STBK", "S803STBL", "S803STRD"] },
                { month: "Mar", date: "3.25-4.7", product: "OpenDots ONE", discount: 0.2, priceAsIs: 199.95, priceToBe: 159.95, saving: 40, rebate: 24, models: ["E310STBK", "E310STBG"] },
                { month: "Mar", date: "3.25-4.7", product: "OpenMove", discount: 0.3, priceAsIs: 79.95, priceToBe: 54.95, saving: 25, rebate: 15, models: ["S661STBL", "S661STGY", "S661STPK", "S661STWT"] },
                { month: "April", date: "4.29-5.10", product: "OpenRun Pro2", discount: 0.2, priceAsIs: 179.95, priceToBe: 139.95, saving: 40, rebate: 24, models: ["S820STBK", "S820STOR", "S821MNBK", "S821MNOR"] },
                { month: "April", date: "4.29-5.10", product: "OpenRun", discount: 0.3, priceAsIs: 129.95, priceToBe: 89.95, saving: 40, rebate: 24, models: ["S803MNBK", "S803MNBL", "S803STBK", "S803STBL", "S803STRD"] },
                { month: "April", date: "4.29-5.10", product: "OpenFit 2+", discount: 0.15, priceAsIs: 199.95, priceToBe: 169.95, saving: 30, rebate: 18, models: ["T921STBK", "T921STGY"] },
                { month: "April", date: "4.29-5.10", product: "OpenDots ONE", discount: 0.2, priceAsIs: 199.95, priceToBe: 159.95, saving: 40, rebate: 24, models: ["E310STBK", "E310STBG"] },
                { month: "May", date: "5.18-5.25", product: "OpenFit Air", discount: 0.3, priceAsIs: 119.95, priceToBe: 79.95, saving: 40, rebate: 24, models: ["T511STBK", "T511STWT", "T511STPK"] },
                { month: "May", date: "5.18-5.25", product: "OpenMove", discount: 0.3, priceAsIs: 79.95, priceToBe: 54.95, saving: 25, rebate: 15, models: ["S661STBL", "S661STGY", "S661STPK", "S661STWT"] }
            ]
        }
    };
    const PROMO_MONTH_ORDER = {
        Jan: 1,
        Feb: 2,
        Mar: 3,
        April: 4,
        May: 5,
        June: 6,
        July: 7,
        August: 8,
        Sep: 9,
        Sept: 9,
        Oct: 10,
        Nov: 11,
        Dec: 12
    };

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
                popFormatLabel: "120cm",
                popSkus: ["E310", "T920", "S803", "S710", "S820", "S821"],
                popVisualSrc: "assets/natm-120cm-pop-example.jpg",
                popVisualAlt: "NFM 120cm POP reference",
                popVisualNote: "NFM 当前 POP 以这类 120cm 一体式陈列为主",
                storeSummary: "4家主门店",
                storeRegions: [
                    "NE: Omaha",
                    "KS: Kansas City",
                    "IA: Clive (Des Moines)",
                    "TX: The Colony (Dallas-Fort Worth)"
                ],
                storeNote: "另有 Omaha 的 Mrs. B's Clearance & Outlet"
            },
            rcw: {
                displayName: "RCW",
                businessTags: ["其他支持"],
                buyer: "david.mcloney@rcwilley.com",
                account: "/",
                shipPrice2025: "0.6",
                shipPrice2026: "0.6",
                setup: "联系 rep 及 buyer",
                popSize: "单独 standard POP，5*30cm",
                popFormatLabel: "单独 standard POP，5*30cm",
                popSkus: ["S820", "S821", "T920", "E310", "S710"],
                storeSummary: "10家零售门店",
                storeRegions: [
                    "UT: Draper / Layton / Orem / Salt Lake City",
                    "NV: Henderson / Reno / Summerlin",
                    "ID: Boise / Meridian",
                    "CA: Rocklin / Sacramento"
                ],
                storeNote: "按官方 Store Locations 统计，不含 warehouse",
                popStatusNote: "RCW 已下单 10 个 T010 单独 POP，预计 4 月底到店。",
                popVisualItems: [
                    { src: "assets/rcw-pop-store-1.png", alt: "RCW POP in store reference 1" },
                    { src: "assets/rcw-pop-store-2.png", alt: "RCW POP in store reference 2" }
                ],
                popVisualNote: "RCW 当前 POP 以该类桌面陈列形式为主"
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
                popFormatLabel: "暂无",
                popSkus: [],
                storeSummary: "1家单店 Showroom",
                storeRegions: [
                    "IL: Glenview (Chicago 北郊)"
                ]
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
                popFormatLabel: "120cm",
                popSkus: ["E310", "T920", "S803", "S710", "S820", "S821", "OpenComm2"],
                popVisualSrc: "assets/natm-120cm-pop-example.jpg",
                popVisualAlt: "BSM 120cm POP reference",
                popVisualNote: "BSM 当前 POP 以这类 120cm 一体式陈列为主",
                storeSummary: "11个到店点位",
                storeRegions: [
                    "FL: Miami / Dadeland / South Dade / Sawgrass Mills / West Palm Beach / Deerfield Beach / Dania Pointe / Outlet Center",
                    "GA: North Atlanta / South Atlanta / Kennesaw"
                ],
                storeNote: "含 Lauderhill 的 Outlet Center"
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
                { channelKey: "bsm", skus: ["E320", "T921", "S803", "S710", "S820", "S821", "OpenComm2"] }
            ],
            augustReplacements: [
                { channelKey: "nfm", currentSku: "S803", nextSku: "NCE（OpenRun 2）" },
                { channelKey: "bsm", currentSku: "S803", nextSku: "NCE（OpenRun 2）" }
            ]
        },
        businessPlaceholders: {
            topActions: [
                {
                    kicker: "节奏管理",
                    title: "固定动作",
                    description: "把 NATM 渠道需要固定推进的动作放在最上面，方便每次打开页面先检查节奏。",
                    badge: "Routine",
                    rows: [
                        {
                            title: "旧品 EOL 通知",
                            note: "检查旧品退市节奏、渠道在售尾货和 buyer 是否需要补充说明，避免旧品继续占 POP 或促销资源。",
                            meta: "节奏：每半个月一次 | 输出：EOL 清单 / buyer 通知"
                        },
                        {
                            title: "新品前 1 个月 SETUP",
                            note: "新品上市前倒推准备 setup、料号、页面信息、送样与 POP 切换，尽量把渠道准备动作前置。",
                            meta: "节奏：上市前 1 个月 | 输出：setup 追踪"
                        },
                        {
                            title: "更新促销日历",
                            note: "促销前至少半个月同步新 promo calendar，并对照下方月度总览观察促销前后销量与销售额变化。",
                            meta: "节奏：至少提前半个月 | 输出：promo calendar"
                        }
                    ]
                },
                {
                    kicker: "项目管理",
                    title: "待办事项",
                    description: "把当前最需要推进、但还没闭环的动作直接挂在首页，减少遗漏。",
                    badge: "To-Do",
                    rows: [
                        {
                            title: "NFM 七八月 POP 快闪店方案落地",
                            note: "继续跟 Kristen 对齐方案，由 Raymond 辅助推进，尽量把新品切换与快闪店机会一起打包。",
                            meta: "渠道：NFM | 状态：进行中 | 优先级：高"
                        },
                        {
                            title: "RCW 与 buyer 沟通季度下单改月度下单",
                            note: "减少季度性下单对库存和节奏的影响，争取建立更稳定的月度补货节奏。",
                            meta: "渠道：RCW | 状态：待会议 | 优先级：高"
                        },
                        {
                            title: "RCW 单独补签 Blair 年度合作合同",
                            note: "当前独立店与 Rep 的年度合作合同未包含 RCW，需要单独补签，避免后续执行脱节。",
                            meta: "渠道：RCW | 状态：待处理 | 优先级：高"
                        },
                        {
                            title: "RCW 争取把 4 台分散 POP 整合为 1 台 90cm POP",
                            note: "结合当前 standard POP 与 T010 单独 POP 的进展，继续争取更完整的一体化展示。",
                            meta: "渠道：RCW | 状态：推进中 | 优先级：中高"
                        },
                        {
                            title: "BSM 建立月末未下单提醒机制",
                            note: "每个月末检查 buyer 是否下单，若未下单则邮件提醒，持续提高下单密度。",
                            meta: "渠道：BSM | 状态：执行中 | 优先级：中高"
                        },
                        {
                            title: "补齐 OpenRun 2 的样品提需与送样动作",
                            note: "OpenDots 2 已有样品，OpenRun 2 仍缺提需，需要尽快补齐并和 buyer 对上送样节奏。",
                            meta: "渠道：All | 状态：待补齐 | 优先级：高"
                        }
                    ]
                }
            ],
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
                        text: "四个渠道刚好都有 OpenDots Air 的黑色、紫色各一个，总计 4+4=8 个；提需已完成，但暂未和 buyer 产生联系，也尚未推进送样。OpenDots 2 也已有样品，当前由 Kevin 保管，等他回来后再给他。"
                    },
                    {
                        title: "待补项目",
                        text: "当前还缺 OpenRun 2 的提需；后续需要继续推进 buyer 沟通，并明确 OpenDots 2 的下一步送样动作。"
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
            nfm: { qty: 4000, sales: 400000 },
            rcw: { qty: 3500, sales: 175000 },
            abt: { qty: 1800, sales: 185000 },
            bsm: { qty: 3800, sales: 400000 }
        },
        channelResearch: {
            nfm: {
                profileTitle: "区域目的地大店，家庭整单购物心智强",
                profileBody: "NFM 官方资料显示其以家具、家电、电子一站式大卖场运营，四个核心卖场拥有超大 showroom 与完整 electronics assortment；结合 PPT 里“大店流量高、停留时间长、耳机专区成熟”的结论，更适合做体验驱动成交。",
                persona: "以家庭决策型客群为主，包含搬家、装修、整屋升级以及会在同一次进店中顺带比较音频产品的中高客单消费者。",
                opportunityTitle: "把试听体验和家庭升级场景绑定，放大高客单转化",
                opportunityBody: "继续用 POP + 主销英雄 SKU 稳住露出，同时把 OpenDots 2、T921 这类新品放在“家居升级 / TV / 健身 / 户外”联动场景里做转化，争取从单纯陈列升级到体验成交。七八月还有一个 POP 快闪店机会，目前由 Kristen 策划、Raymond 辅助推进。",
                nextStepTitle: "优先推进 POP 快闪店和新品切换",
                nextStepBody: "继续跟进七八月 POP 快闪店机会，同时把 OpenDots 2、T921 与后续 NCE 的 POP / 上样 / 促销节奏提早对齐。"
            },
            rcw: {
                profileTitle: "西部区域家居电器零售商，偏家庭计划型购买",
                profileBody: "RC Willey 官网定位是服务西部市场的 home furnishings、appliances 与 electronics 零售商；结合 PPT 中“门店空间大、顾客停留久、独立店销售表现不错”的信息，这类店更像 destination store，而不是纯快进快出的电子卖场。",
                persona: "更偏郊区家庭、换新型用户和成套购买用户，重视 financing、门店服务和可视化陈列，愿意在店内做横向比较后再下单。",
                opportunityTitle: "围绕西部家庭客群做稳定扩店和结构升级",
                opportunityBody: "RCW 的增长空间在于把现有成功门店打法复制到更多门店与更多核心 SKU 上，用更清晰的 POP、培训和价格带分层，把 Shokz 从补充型音频 SKU 变成被主动推荐的功能型耳机。",
                nextStepTitle: "推动下单节奏与 POP 形态升级",
                nextStepBody: "和 buyer 讨论把季度性下单改成月度下单，补签 Blair 合同，并争取把 4 台分散 POP 整合成 1 台 90cm POP。"
            },
            abt: {
                profileTitle: "单店高服务模型，专业咨询和高 ASP 更关键",
                profileBody: "Abt 官方公开信息强调其 Glenview 单店、90+ 年历史、公开 tour、培训体系与 award-winning customer service；它不是广撒网渠道，更像高服务密度的精品大店。",
                persona: "以研究型、信任导购型、愿意为服务和专业建议买单的 Chicagoland 用户为主，对产品故事、功能差异和售后体验更敏感。",
                opportunityTitle: "走精品化和顾问式成交，优先做高质量产出",
                opportunityBody: "Abt 更适合做少而精的结构优化，集中火力推高 ASP 主力款和具备明确卖点的新型号，通过导购教育、场景话术和本地化活动，拉高单点产出而不是盲目铺量。",
                nextStepTitle: "承接 exclusive 促销结果推进全旗舰入店",
                nextStepBody: "复盘 2-3 月 exclusive store only 促销效果，继续争取全旗舰、全色在店内完整呈现，并提前准备后续新品 setup。"
            },
            bsm: {
                profileTitle: "东南部价格驱动大卖场，促销和现货感知更强",
                profileBody: "BrandsMart USA 官网强调 huge selection、Audio & Headphones、110% Price Match、会员与 financing；结合 PPT 中“人流量高、竞品多、靠近 Costco / Sam / Best Buy”的判断，这一渠道更偏高流量、强促销、强对比购物环境。",
                persona: "价格敏感型、促销响应型和家庭换新型用户占比较高，重视即时可得、价格优势和清晰的卖点提示，容易在大促节点集中释放销量。",
                opportunityTitle: "用促销节点和主力价位带放大爆发力",
                opportunityBody: "BSM 需要把 POP 更新、主推 SKU、价格带和大促节奏绑在一起，重点做高流量月份的主力款曝光与补货，争取把自然陈列流量转成更高的销量和销售额贡献。",
                nextStepTitle: "用月末提醒机制提升下单密度",
                nextStepBody: "既然已和 buyer 对齐提高下单密度，后续每个月末都检查是否下单，必要时邮件提醒，并同步跟进 E320 / T921 / NCE 切换。"
            }
        },
        promoCalendar: [
            { month: "6月", note: "POP 更新切换，跟踪 E320 / T921 替换后的首轮 sell-through" },
            { month: "7月", note: "渠道对齐与跟进，复盘 POP 执行、补货和门店反馈" },
            { month: "8月", note: "关注 NFM POP 快闪店机会，以及 OpenDots Air / OpenRun 2 的上市与 POP 更新窗口" },
            { month: "9月+", note: "结合节庆和下半年促销节奏，继续补充重点活动节点" }
        ],
        promoCalendars: PROMO_CALENDARS,
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
        selectedPromoMonth: "all",
        selectedPromoChannel: "all",
        selectedSkuByChannel: {},
        expandedMixBases: {},
        charts: {}
    };

    const els = {
        eyebrow: document.querySelector(".eyebrow"),
        heroTitle: document.querySelector(".hero h1"),
        heroCopy: document.querySelector(".hero-copy"),
        footer: document.querySelector(".footer"),
        sidebarChannelSwitch: document.getElementById("sidebarChannelSwitch"),
        businessOverviewGrid: document.getElementById("businessOverviewGrid"),
        fullProductMixPanel: document.getElementById("fullProductMixPanel"),
        popPlanPanel: document.getElementById("popPlanPanel"),
        launchPlanPanel: document.getElementById("launchPlanPanel"),
        businessSupportGrid: document.getElementById("businessSupportGrid"),
        topActionGrid: document.getElementById("topActionGrid"),
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
        const fractionDigits = digits === undefined ? 0 : Math.min(Number(digits) || 0, 1);
        return new Intl.NumberFormat("en-US", {
            maximumFractionDigits: fractionDigits,
            minimumFractionDigits: fractionDigits
        }).format(value);
    }

    function formatCurrency(value, digits) {
        const fractionDigits = digits === undefined ? 0 : Math.min(Number(digits) || 0, 1);
        return new Intl.NumberFormat("en-US", {
            style: "currency",
            currency: "USD",
            maximumFractionDigits: fractionDigits,
            minimumFractionDigits: fractionDigits
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

    function promoMonthValue(monthLabel) {
        return PROMO_MONTH_ORDER[String(monthLabel || "").trim()] || 99;
    }

    function promoDateValue(dateLabel) {
        const match = String(dateLabel || "").trim().match(/^(\d+)(?:\.(\d+))?/);
        if (!match) {
            return 9999;
        }
        return Number(match[1]) * 100 + Number(match[2] || 0);
    }

    function sortPromoRows(rows) {
        return (rows || []).slice().sort((left, right) => {
            return promoMonthValue(left.month) - promoMonthValue(right.month)
                || promoDateValue(left.date) - promoDateValue(right.date);
        });
    }

    function renderPromoModels(models) {
        if (!models || !models.length) {
            return "—";
        }
        return [
            "<div class=\"profile-tag-list\">",
            models.map(model => "<span class=\"profile-tag muted\">" + escapeHtml(model) + "</span>").join(""),
            "</div>"
        ].join("");
    }

    function getPromoMonthOptions() {
        const seen = new Set();
        return CONFIG.channels.reduce((months, channel) => {
            (CONFIG.promoCalendars[channel.key] && CONFIG.promoCalendars[channel.key].rows || []).forEach(row => {
                if (row.month && !seen.has(row.month)) {
                    seen.add(row.month);
                    months.push(row.month);
                }
            });
            return months;
        }, []).sort((left, right) => promoMonthValue(left) - promoMonthValue(right));
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

    function normalizeProductKey(rawSku) {
        return String(rawSku || "").trim().toUpperCase();
    }

    function priceTupleToObject(tuple) {
        if (!Array.isArray(tuple) || !tuple.length) {
            return { msrp: null, po: null };
        }
        const msrp = Number(tuple[0]);
        const po = Number(tuple[1]);
        return {
            msrp: Number.isFinite(msrp) ? msrp : null,
            po: Number.isFinite(po) ? po : null
        };
    }

    function buildExactPriceCandidates(rawSku) {
        const key = normalizeProductKey(rawSku);
        if (!key) {
            return [];
        }
        const candidates = [key];
        if (/\-US$/i.test(key)) {
            candidates.push(key.replace(/\-US$/i, ""));
        } else if (/\-/.test(key)) {
            candidates.push(key + "-US");
        }
        return candidates;
    }

    function lookupProductPrice(channelKey, row) {
        const pricing = CHANNEL_PRODUCT_PRICING[channelKey] || {};
        const exactPricing = pricing.exact || {};
        const basePricing = pricing.base || {};
        const exactCandidates = row.level === "specific" ? buildExactPriceCandidates(row.sku) : [];

        for (let index = 0; index < exactCandidates.length; index += 1) {
            if (exactPricing[exactCandidates[index]]) {
                return priceTupleToObject(exactPricing[exactCandidates[index]]);
            }
        }

        const baseKey = normalizeProductKey(row.baseSku || baseSku(row.sku));
        if (basePricing[baseKey]) {
            return priceTupleToObject(basePricing[baseKey]);
        }

        return { msrp: null, po: null };
    }

    function formatProductPrice(value) {
        if (!Number.isFinite(value)) {
            return "—";
        }
        const truncated = Math.trunc(value * 100) / 100;
        return new Intl.NumberFormat("en-US", {
            style: "currency",
            currency: "USD",
            maximumFractionDigits: 2,
            minimumFractionDigits: 2
        }).format(truncated);
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

    function isBaseProductSku(sku) {
        return /^[A-Z]\d{3}$/i.test(String(sku || "").trim());
    }

    function compareProductMixItems(left, right) {
        if (right.sales !== left.sales) {
            return right.sales - left.sales;
        }
        if (right.qty !== left.qty) {
            return right.qty - left.qty;
        }
        return String(left.sku || "").localeCompare(String(right.sku || ""));
    }

    function buildProductMix(records, skuMonthly, monthKeys, startMonthKey) {
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
                isBaseProductSku(item.sku) &&
                (item.qty > 0 || item.sales > 0);
        }).sort(compareProductMixItems);

        const totalSales = items.reduce((sum, item) => sum + Math.max(item.sales, 0), 0);
        const totalQty = items.reduce((sum, item) => sum + Math.max(item.qty, 0), 0);

        items.forEach(item => {
            item.salesShare = totalSales > 0 ? item.sales / totalSales : null;
            item.qtyShare = totalQty > 0 ? item.qty / totalQty : null;
        });

        const relevantKeySet = new Set(relevantKeys);
        const specificByBase = {};

        (records || []).forEach(record => {
            if (!relevantKeySet.has(record.monthKey) || !isBaseProductSku(record.sku)) {
                return;
            }
            if (record.sku === CONFIG.adjustmentLabel || (record.qty === 0 && record.sales === 0)) {
                return;
            }

            const specificSku = String(record.specificSku || "").trim();
            if (!specificSku || specificSku === record.sku || baseSku(specificSku) !== record.sku) {
                return;
            }

            if (!specificByBase[record.sku]) {
                specificByBase[record.sku] = {};
            }
            if (!specificByBase[record.sku][specificSku]) {
                specificByBase[record.sku][specificSku] = {
                    sku: specificSku,
                    qty: 0,
                    sales: 0
                };
            }

            specificByBase[record.sku][specificSku].qty += record.qty;
            specificByBase[record.sku][specificSku].sales += record.sales;
        });

        const detailedRows = [];

        items.forEach((item, index) => {
            const specificItems = Object.values(specificByBase[item.sku] || {}).sort(compareProductMixItems).map(specificItem => {
                return {
                    level: "specific",
                    baseSku: item.sku,
                    sku: specificItem.sku,
                    qty: specificItem.qty,
                    sales: specificItem.sales,
                    qtyShare: totalQty > 0 ? specificItem.qty / totalQty : null,
                    salesShare: totalSales > 0 ? specificItem.sales / totalSales : null
                };
            });

            item.specificItems = specificItems;
            detailedRows.push({
                level: "base",
                rank: index + 1,
                baseSku: item.sku,
                sku: item.sku,
                qty: item.qty,
                sales: item.sales,
                qtyShare: item.qtyShare,
                salesShare: item.salesShare,
                specificItems: specificItems
            });
            specificItems.forEach(specificItem => {
                detailedRows.push(specificItem);
            });
        });

        return {
            startMonthKey: startMonthKey,
            monthKeys: relevantKeys,
            items: items,
            detailedRows: detailedRows,
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
        const specificSku = String(row["Specific SKU"] || "").trim();
        const qty = parseNumber(row["Total Quantity Sold"]);
        const sales = parseNumber(row["Total Sales"]);
        if (!rawSku && !specificSku && qty === 0 && sales === 0) {
            return null;
        }

        const primarySku = rawSku || specificSku;

        return {
            year: year,
            month: month,
            monthKey: toMonthKey(year, month),
            rawSku: primarySku,
            sku: baseSku(primarySku),
            specificSku: specificSku || primarySku,
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

        const productMixSinceStart = buildProductMix(records, skuMonthly, monthKeys, CONFIG.structureStartMonth);

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

    function setActiveChannel(channelKey, options) {
        const nextChannel = getChannelConfig(channelKey);
        if (!nextChannel || !nextChannel.key) {
            return;
        }
        state.activeChannelKey = nextChannel.key;
        renderAll();

        if (options && options.scrollToBusinessDetail) {
            requestAnimationFrame(() => {
                const detail = document.getElementById("business-detail-" + nextChannel.key);
                if (detail && typeof detail.scrollIntoView === "function") {
                    detail.scrollIntoView({ behavior: "smooth", block: "nearest" });
                }
            });
        }
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

    function joinBusinessTags(values) {
        return (values || []).join(" / ") || "暂无";
    }

    function firstContactLine(value) {
        return String(value || "").split("/")[0].trim() || "暂无";
    }

    function mixExpansionKey(channelKey, baseSkuValue) {
        return String(channelKey || "") + ":" + String(baseSkuValue || "");
    }

    function isMixBaseExpanded(channelKey, baseSkuValue) {
        return Boolean(state.expandedMixBases[mixExpansionKey(channelKey, baseSkuValue)]);
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

    function renderMixWatchCard(label, value, note) {
        return [
            "<article class=\"business-metric-card mix-watch-card\">",
            "<p class=\"business-metric-label\">" + escapeHtml(label) + "</p>",
            "<p class=\"business-metric-value\">" + escapeHtml(value) + "</p>",
            note ? "<p class=\"mix-watch-note\">" + escapeHtml(note) + "</p>" : "",
            "</article>"
        ].join("");
    }

    function renderProgressMeta(item) {
        const parts = [item.status, item.owner, item.priority].filter(Boolean);
        if (!parts.length) {
            return "";
        }
        return "<div class=\"timeline-meta\">" + parts.map(part => {
            return "<span class=\"timeline-tag\">" + escapeHtml(part) + "</span>";
        }).join("") + "</div>";
    }

    function renderChecklistCard(module) {
        const items = module.checklistItems || module.rows || [];
        const emptyText = module.body || "鏆傛椂绌虹潃";
        return [
            "<article class=\"info-card checklist-card\">",
            "<div class=\"info-card-head\">",
            "<div>",
            module.kicker ? "<p class=\"info-kicker\">" + escapeHtml(module.kicker) + "</p>" : "",
            "<h3 class=\"info-card-title\">" + escapeHtml(module.title) + "</h3>",
            module.description ? "<p class=\"info-card-copy\">" + escapeHtml(module.description) + "</p>" : "",
            "</div>",
            "<span class=\"info-badge\">" + escapeHtml(module.badge || "Checklist") + "</span>",
            "</div>",
            items.length ? [
                "<div class=\"checklist-list\">",
                items.map(item => {
                    return [
                        "<div class=\"checklist-item" + (item.done ? " done" : "") + "\">",
                        "<span class=\"checklist-box\" aria-hidden=\"true\"></span>",
                        "<div class=\"checklist-copy\">",
                        "<strong>" + escapeHtml(item.title) + "</strong>",
                        item.note ? "<span class=\"checklist-note\">" + escapeHtml(item.note) + "</span>" : "",
                        item.meta ? "<span>" + escapeHtml(item.meta) + "</span>" : "",
                        "</div>",
                        "</div>"
                    ].join("");
                }).join(""),
                "</div>"
            ].join("") : "<div class=\"checklist-empty\">" + escapeHtml(emptyText) + "</div>",
            "</article>"
        ].join("");
    }

    function renderStoreFootprint(profile) {
        const regions = Array.isArray(profile.storeRegions) ? profile.storeRegions : [];
        if (!profile.storeSummary && !regions.length && !profile.storeNote) {
            return "";
        }

        return [
            "<div class=\"business-store-card\">",
            "<p class=\"business-store-title\">门店网络</p>",
            profile.storeSummary ? "<p class=\"business-store-summary\">" + escapeHtml(profile.storeSummary) + "</p>" : "",
            regions.length ? "<div class=\"business-store-list\">" + regions.map(region => {
                return "<p class=\"business-store-line\">" + escapeHtml(region) + "</p>";
            }).join("") + "</div>" : "",
            profile.storeNote ? "<p class=\"business-store-note\">" + escapeHtml(profile.storeNote) + "</p>" : "",
            "</div>"
        ].join("");
    }

    function renderPopVisual(profile) {
        const visualItems = Array.isArray(profile.popVisualItems) && profile.popVisualItems.length
            ? profile.popVisualItems
            : (profile.popVisualSrc ? [{ src: profile.popVisualSrc, alt: profile.popVisualAlt || "POP display reference" }] : []);
        if (!visualItems.length) {
            return "";
        }

        return [
            "<figure class=\"business-pop-visual\">",
            "<div class=\"business-pop-gallery\">",
            visualItems.map(item => {
                return "<img src=\"" + encodeURI(item.src) + "\" alt=\"" + escapeHtml(item.alt || "POP display reference") + "\" loading=\"lazy\">";
            }).join(""),
            "</div>",
            profile.popVisualNote ? "<figcaption>" + escapeHtml(profile.popVisualNote) + "</figcaption>" : "",
            "</figure>"
        ].join("");
    }

    function renderPopStatusNote(profile) {
        if (!profile.popStatusNote) {
            return "";
        }

        return "<p class=\"business-pop-note\"><strong>POP更新：</strong>" + escapeHtml(profile.popStatusNote) + "</p>";
    }

    function renderRankIndex(index) {
        return "<span class=\"rank-index\">" + String(index) + "</span>";
    }

    function renderMixLevelCell(level, isPopSku) {
        const levelLabel = level === "specific" ? "细分SKU" : "大类SKU";
        return [
            "<div class=\"mix-badge-group\">",
            "<span class=\"mix-badge level\">" + escapeHtml(levelLabel) + "</span>",
            renderMixBadge(isPopSku),
            "</div>"
        ].join("");
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

        const activeChannel = getChannelConfig(state.activeChannelKey);
        const activeDashboard = state.dashboards[state.activeChannelKey];
        const activeProfile = CONFIG.channelProfiles[state.activeChannelKey] || {};
        const activeMix = activeDashboard ? activeDashboard.productMixSinceStart : null;
        const activeItems = activeMix ? activeMix.items : [];
        const activePopSet = new Set(activeProfile.popSkus || []);
        const topMixItems = activeItems.slice(0, 3);
        const topNonPopItem = activeItems.find(item => !activePopSet.has(item.sku));
        const popSoldCount = (activeProfile.popSkus || []).filter(sku => activeItems.some(item => item.sku === sku)).length;

        els.fullProductMixPanel.innerHTML = [
            "<article class=\"info-card\">",
            "<div class=\"info-card-head\">",
            "<div>",
            "<p class=\"info-kicker\">完整结构</p>",
            "<h3 class=\"info-card-title\">完整产品结构</h3>",
            "<p class=\"info-card-copy\">按渠道折叠展示 2025-09 至今所有实际卖出的产品结构，同时保留大类 SKU 与 Specific SKU 两层记录；点击某个渠道即可展开查看完整明细。</p>",
            "</div>",
            "<span class=\"info-badge\">Full Mix</span>",
            "</div>",
            activeMix ? [
                "<div class=\"business-metric-grid mix-watch-grid\">",
                renderMixWatchCard("当前渠道", activeChannel.label, "下面默认展开当前选中的渠道结构"),
                renderMixWatchCard("Top SKU", topMixItems.length ? topMixItems.map(item => item.sku).join(" / ") : "暂无", topMixItems.length ? ("Top1 销售额 " + formatCurrency(topMixItems[0].sales)) : ""),
                renderMixWatchCard("非 POP 重点", topNonPopItem ? topNonPopItem.sku : "暂无", topNonPopItem ? ("销售额 " + formatCurrency(topNonPopItem.sales)) : "当前卖出结构已基本被 POP 覆盖"),
                renderMixWatchCard("POP 覆盖", activeProfile.popSkus && activeProfile.popSkus.length ? (String(popSoldCount) + "/" + String(activeProfile.popSkus.length)) : "暂无 POP", activeMix ? ("统计口径：2025-09 至 " + activeDashboard.latestMonthKey) : ""),
                "</div>"
            ].join("") : "",
            "<div class=\"mix-accordion\">",
            CONFIG.channels.map(channel => {
                const dashboard = state.dashboards[channel.key];
                const profile = CONFIG.channelProfiles[channel.key];
                const popSkuSet = new Set(profile.popSkus);
                const rows = dashboard ? dashboard.productMixSinceStart.detailedRows : [];
                const channelLabel = profile.displayName || channel.label;
                const productMix = dashboard ? dashboard.productMixSinceStart : { totalQty: 0, totalSales: 0 };
                const openAttr = state.activeChannelKey === channel.key ? " open" : "";

                if (!rows.length) {
                    return [
                        "<details class=\"mix-accordion-item\"" + openAttr + " style=\"" + channelStyle(channel) + "\">",
                        "<summary class=\"mix-accordion-summary\">",
                        "<div class=\"mix-accordion-title\"><strong>" + escapeHtml(channelLabel) + "</strong><span class=\"channel-chip\">NATM</span></div>",
                        "<div class=\"mix-accordion-meta\"><span>暂无卖出 SKU 结构记录</span></div>",
                        "</summary>",
                        "</details>"
                    ].join("");
                }

                return [
                    "<details class=\"mix-accordion-item\"" + openAttr + " style=\"" + channelStyle(channel) + "\">",
                    "<summary class=\"mix-accordion-summary\">",
                    "<div class=\"mix-accordion-title\"><strong>" + escapeHtml(channelLabel) + "</strong><span class=\"channel-chip\">NATM</span></div>",
                    "<div class=\"mix-accordion-meta\"><span>总销量 " + formatNumber(productMix.totalQty) + "</span><span>总销售额 " + formatCurrency(productMix.totalSales) + "</span><span>大类 SKU " + formatNumber(dashboard.productMixSinceStart.items.length) + "</span></div>",
                    "</summary>",
                    "<div class=\"mix-accordion-body\">",
                    "<div class=\"table-wrap\">",
                    "<table class=\"full-mix-table\">",
                    "<thead>",
                    "<tr>",
                    "<th>排名</th>",
                    "<th>SKU</th>",
                    "<th>记录层级</th>",
                    "<th>MSRP</th>",
                    "<th>PO Price</th>",
                    "<th>销量</th>",
                    "<th>销量占比</th>",
                    "<th>销售额</th>",
                    "<th>销售额占比</th>",
                    "</tr>",
                    "</thead>",
                    "<tbody>",
                    rows.map(row => {
                        const qtyShare = row.qtyShare !== null ? formatPercent(row.qtyShare, true) : "—";
                        const salesShare = row.salesShare !== null ? formatPercent(row.salesShare, true) : "—";
                        const isPopSku = popSkuSet.has(row.baseSku || row.sku);
                        const hasSpecificRows = row.level === "base" && row.specificItems && row.specificItems.length;
                        const isExpanded = row.level === "base" ? isMixBaseExpanded(channel.key, row.baseSku) : false;
                        const isHidden = row.level === "specific" && !isMixBaseExpanded(channel.key, row.baseSku);
                        const rowClass = row.level === "specific" ? "mix-detail-row" : "mix-base-row";
                        const pricing = lookupProductPrice(channel.key, row);
                        const skuCell = row.level === "specific"
                            ? "<span class=\"mix-sku-sub\">" + escapeHtml(row.sku) + "</span>"
                            : [
                                "<div class=\"mix-base-sku-wrap\">",
                                hasSpecificRows
                                    ? "<button class=\"mix-toggle-btn\" type=\"button\" data-mix-channel=\"" + escapeHtml(channel.key) + "\" data-mix-base=\"" + escapeHtml(row.baseSku) + "\">" + (isExpanded ? "收起" : "展开") + "</button>"
                                    : "",
                                "<strong>" + escapeHtml(row.sku) + "</strong>",
                                "</div>"
                            ].join("");
                        return [
                            "<tr class=\"" + rowClass + (isHidden ? " is-hidden" : "") + "\">",
                            "<td>" + (row.level === "base" ? renderRankIndex(row.rank) : "") + "</td>",
                            "<td>" + skuCell + "</td>",
                            "<td>" + renderMixLevelCell(row.level, isPopSku) + "</td>",
                            "<td>" + formatProductPrice(pricing.msrp) + "</td>",
                            "<td>" + formatProductPrice(pricing.po) + "</td>",
                            "<td>" + formatNumber(row.qty) + "</td>",
                            "<td>" + qtyShare + "</td>",
                            "<td>" + formatCurrency(row.sales) + "</td>",
                            "<td>" + salesShare + "</td>",
                            "</tr>"
                        ].join("");
                    }).join(""),
                    "</tbody>",
                    "</table>",
                    "</div>",
                    "</div>",
                    "</details>"
                ].join("");
            }).join(""),
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
                const profile = CONFIG.channelProfiles[group.channelKey];
                return [
                    "<article class=\"plan-mini-card\" style=\"" + channelStyle(channel) + "\">",
                    "<h4 class=\"plan-mini-title\">" + escapeHtml(channel.label) + "</h4>",
                    profile && profile.popFormatLabel ? "<p class=\"plan-mini-meta\">" + escapeHtml(profile.popFormatLabel) + "</p>" : "",
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

    function renderTopActionGrid() {
        if (!els.topActionGrid) {
            return;
        }
        els.topActionGrid.innerHTML = (CONFIG.businessPlaceholders.topActions || []).map(renderChecklistCard).join("");
    }

    function buildChannelStrategyModules() {
        const activeChannel = getChannelConfig(state.activeChannelKey);
        const research = CONFIG.channelResearch[activeChannel.key] || {};
        const target = CONFIG.channelTargets[activeChannel.key] || { qty: 0, sales: 0 };
        const dashboard = state.dashboards[activeChannel.key];
        const qtyGrowth = dashboard ? calculateGrowth(target.qty, dashboard.yearlyTotals[2025].qty) : null;
        const salesGrowth = dashboard ? calculateGrowth(target.sales, dashboard.yearlyTotals[2025].sales) : null;

        return [
            {
                kicker: "渠道洞察",
                title: "渠道特点及用户画像",
                description: "当前展示 " + activeChannel.label + " 的渠道特点与用户画像；点击上方渠道卡片可切换。",
                badge: activeChannel.label,
                rows: [
                    {
                        channelKey: activeChannel.key,
                        title: research.profileTitle,
                        text: research.profileBody,
                        meta: "用户画像：" + research.persona
                    }
                ],
                footnote: "公开信息主要来自 NFM、RC Willey、Abt、BrandsMart USA 官网；用户画像为结合 PPT 与店型结构的推断。"
            },
            {
                kicker: "增长方向",
                title: "核心机会",
                description: "当前展示 " + activeChannel.label + " 的优先增长动作，后续可继续按渠道逐条细化。",
                badge: activeChannel.label,
                rows: [
                    {
                        channelKey: activeChannel.key,
                        title: research.opportunityTitle,
                        text: research.opportunityBody
                    }
                ]
            },
            {
                kicker: "下一步动作",
                title: "下一步动作",
                description: "当前展示 " + activeChannel.label + " 接下来最值得优先推进的动作。",
                badge: activeChannel.label,
                rows: [
                    {
                        channelKey: activeChannel.key,
                        title: research.nextStepTitle || "待补充",
                        text: research.nextStepBody || "后续补充该渠道的下一步动作。"
                    }
                ]
            },
            {
                kicker: "年度管理",
                title: "今年目标",
                description: "当前展示 " + activeChannel.label + " 的 2026 年目标值和相对 2025 实际的目标同比。",
                badge: "2026",
                rows: [
                    {
                        channelKey: activeChannel.key,
                        title: CONFIG.growthGoalText,
                        metrics: [
                            { label: "目标销量", value: formatNumber(target.qty) },
                            { label: "目标销售额", value: formatCurrency(target.sales) }
                        ],
                        text: "目标口径统一按至少 +30%、标准 +40%、挑战 +50%+推进。",
                        meta: dashboard
                            ? "相对 2025 实际：销量 " + formatPercent(qtyGrowth, false) + "，销售额 " + formatPercent(salesGrowth, false) + "。"
                            : "数据加载后自动回填相对 2025 的目标同比。"
                    }
                ]
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
        const monthOptions = getPromoMonthOptions();
        const selectedMonth = state.selectedPromoMonth || "all";
        const selectedChannel = state.selectedPromoChannel || "all";
        const visibleChannels = selectedChannel === "all"
            ? CONFIG.channels
            : CONFIG.channels.filter(channel => channel.key === selectedChannel);
        els.promoCalendarPanel.innerHTML = [
            "<p class=\"promo-panel-intro\">已根据四个渠道的 promo calendar Excel 汇总成表格，支持按月份和渠道筛选，便于直接和下方月度总览对照查看。</p>",
            "<div class=\"promo-filter-bar\">",
            ["all"].concat(monthOptions).map(month => {
                const isActive = month === selectedMonth ? " active" : "";
                const label = month === "all" ? "All" : month;
                return "<button class=\"metric-btn" + isActive + "\" type=\"button\" data-promo-month=\"" + escapeHtml(month) + "\">" + escapeHtml(label) + "</button>";
            }).join(""),
            "</div>",
            "<div class=\"promo-filter-bar compact\">",
            ["all"].concat(CONFIG.channels.map(channel => channel.key)).map(channelKey => {
                const isActive = channelKey === selectedChannel ? " active" : "";
                const label = channelKey === "all" ? "All Channels" : getChannelConfig(channelKey).label;
                return "<button class=\"metric-btn" + isActive + "\" type=\"button\" data-promo-channel=\"" + escapeHtml(channelKey) + "\">" + escapeHtml(label) + "</button>";
            }).join(""),
            "</div>",
            "<div class=\"mix-accordion\">",
            visibleChannels.map(channel => {
                const calendar = CONFIG.promoCalendars[channel.key] || { source: "", rows: [] };
                const allRows = sortPromoRows(calendar.rows);
                const rows = selectedMonth === "all"
                    ? allRows
                    : allRows.filter(row => row.month === selectedMonth);
                const months = Array.from(new Set(allRows.map(row => row.month)));
                const coverage = selectedMonth === "all"
                    ? (months.length ? months[0] + (months.length > 1 ? " - " + months[months.length - 1] : "") : "—")
                    : selectedMonth;
                const modelCount = rows.reduce((total, row) => total + row.models.length, 0);
                const openAttr = selectedChannel === "all"
                    ? (state.activeChannelKey === channel.key ? " open" : "")
                    : " open";
                const coverageLabel = selectedMonth === "all" ? "覆盖 " : "筛选 ";

                return [
                    "<details class=\"mix-accordion-item\"" + openAttr + " style=\"" + channelStyle(channel) + "\">",
                    "<summary class=\"mix-accordion-summary\">",
                    "<div class=\"mix-accordion-title\"><strong>" + escapeHtml(channel.label) + "</strong><span class=\"channel-chip\">Promo</span></div>",
                    "<div class=\"mix-accordion-meta\"><span>" + coverageLabel + escapeHtml(coverage) + "</span><span>" + formatNumber(rows.length) + " 条促销</span><span>" + formatNumber(modelCount) + " 个型号</span></div>",
                    "</summary>",
                    "<div class=\"mix-accordion-body\">",
                    rows.length ? [
                        "<div class=\"table-wrap\">",
                        "<table class=\"promo-calendar-table\">",
                        "<thead><tr><th>月份</th><th>档期</th><th>产品</th><th>折扣</th><th>原价</th><th>活动价</th><th>Promo Saving</th><th>Rebate</th><th>涉及型号</th></tr></thead>",
                        "<tbody>",
                        rows.map(row => {
                            return [
                                "<tr>",
                                "<td><strong>" + escapeHtml(row.month) + "</strong></td>",
                                "<td>" + escapeHtml(row.date) + "</td>",
                                "<td>" + escapeHtml(row.product) + "</td>",
                                "<td>" + formatPercent(row.discount, true) + "</td>",
                                "<td>" + formatCurrency(row.priceAsIs, 2) + "</td>",
                                "<td>" + formatCurrency(row.priceToBe, 2) + "</td>",
                                "<td>" + formatCurrency(row.saving, 2) + "</td>",
                                "<td>" + formatCurrency(row.rebate, 2) + "</td>",
                                "<td class=\"promo-model-cell\">" + renderPromoModels(row.models) + "</td>",
                                "</tr>"
                            ].join("");
                        }).join(""),
                        "</tbody>",
                        "</table>",
                        "</div>"
                    ].join("") : "<p class=\"placeholder-copy\">" + escapeHtml(channel.label) + " 在 " + escapeHtml(selectedMonth) + " 暂无促销</p>",
                    calendar.source ? "<p class=\"table-note\">来源：" + escapeHtml(calendar.source) + "</p>" : "",
                    "</div>",
                    "</details>"
                ].join("");
            }).join(""),
            "</div>",
            "<p class=\"promo-compare-copy\">建议把这块与下方月度总览一起看，重点对照促销档期前后 1-2 个月的销量、销售额和 Top SKU 变化。</p>"
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

    function renderBusinessDetailPanel(channel, profile, dashboard) {
        const productMix = dashboard.productMixSinceStart;
        const popSkuSet = new Set(profile.popSkus);
        const soldSkuSet = new Set(productMix.items.map(item => item.sku));
        const popInSales = profile.popSkus.filter(sku => soldSkuSet.has(sku));
        const popCoverageText = profile.popSkus.length ? ("POP 覆盖 " + popInSales.length + "/" + profile.popSkus.length) : "当前无 POP 组合";

        return [
            "<details class=\"business-detail-item\" id=\"business-detail-" + escapeHtml(channel.key) + "\"" + (state.activeChannelKey === channel.key ? " open" : "") + " style=\"" + channelStyle(channel) + "\">",
            "<summary class=\"business-detail-summary\">",
            "<div class=\"business-detail-title\"><strong>" + escapeHtml(profile.displayName || channel.label) + "</strong><span class=\"channel-chip\">Detail</span></div>",
            "<div class=\"business-detail-meta\"><span>门店 " + escapeHtml(profile.storeSummary || "—") + "</span><span>卖出SKU " + formatNumber(productMix.items.length) + "</span><span>销售额 " + formatCurrency(productMix.totalSales) + "</span></div>",
            "</summary>",
            "<div class=\"business-detail-body\">",
            "<div class=\"business-detail-grid\">",
            "<article class=\"profile-card\">",
            "<div class=\"profile-card-head\"><div><h4 class=\"profile-card-title\">" + escapeHtml(profile.displayName || channel.label) + "</h4><p class=\"profile-card-sub\">最新月份 " + escapeHtml(dashboard.latestMonthKey) + " · 2026 YTD " + formatCurrency(dashboard.samePeriodByYear[2026].sales) + "</p></div><span class=\"channel-chip\">NATM</span></div>",
            "<div class=\"profile-block\">",
            "<p class=\"profile-label\">门店网络</p>",
            renderStoreFootprint(profile),
            "</div>",
            "</article>",
            "<article class=\"profile-card\">",
            "<div class=\"profile-block\">",
            "<p class=\"profile-label\">商务对接</p>",
            "<div class=\"profile-tag-list\">" + renderProfileTags(profile.businessTags) + "</div>",
            "<p class=\"profile-copy\"><strong>Buyer：</strong>" + escapeHtml(profile.buyer) + "</p>",
            "<p class=\"profile-copy\"><strong>Account：</strong>" + escapeHtml(profile.account) + "</p>",
            "<p class=\"profile-copy\"><strong>Setup：</strong>" + escapeHtml(profile.setup) + "</p>",
            "</div>",
            "</article>",
            "<article class=\"profile-card\">",
            "<div class=\"profile-block\">",
            "<p class=\"profile-label\">出货价</p>",
            "<div class=\"profile-price-grid\">",
            "<div class=\"profile-price\"><p class=\"profile-price-label\">2025 出货价</p><p class=\"profile-price-value\">" + escapeHtml(profile.shipPrice2025) + "</p></div>",
            "<div class=\"profile-price\"><p class=\"profile-price-label\">2026 出货价</p><p class=\"profile-price-value\">" + escapeHtml(profile.shipPrice2026) + "</p></div>",
            "</div>",
            "</div>",
            "</article>",
            "<article class=\"profile-card\">",
            "<div class=\"profile-block\">",
            "<p class=\"profile-label\">当前POP产品情况</p>",
            "<p class=\"profile-copy\"><strong>POP形式：</strong>" + escapeHtml(profile.popFormatLabel || profile.popSize) + "</p>",
            renderPopStatusNote(profile),
            "<div class=\"profile-tag-list\">" + renderProfileTags(profile.popSkus) + "</div>",
            renderPopVisual(profile),
            "</div>",
            "</article>",
            "<article class=\"profile-card business-detail-wide\">",
            "<div class=\"profile-block\">",
            "<p class=\"profile-label\">产品结构概览</p>",
            "<div class=\"business-metric-grid\">",
            renderBusinessMetricCard("卖出SKU", String(productMix.items.length)),
            renderBusinessMetricCard("POP覆盖", profile.popSkus.length ? (String(popInSales.length) + "/" + String(profile.popSkus.length)) : "无POP"),
            renderBusinessMetricCard("累计销量", formatNumber(productMix.totalQty)),
            renderBusinessMetricCard("销售额", formatCurrency(productMix.totalSales)),
            "</div>",
            "<p class=\"profile-mix-note\">统计口径：2025-09 至 " + escapeHtml(dashboard.latestMonthKey) + "。 " + popCoverageText + "</p>",
            renderBusinessRankTable(productMix.items, popSkuSet),
            "</div>",
            "</article>",
            "</div>",
            "</div>",
            "</details>"
        ].join("");
    }

    function renderExecutionGrid() {
        if (!els.executionGrid) {
            return;
        }
        const activeChannel = getChannelConfig(state.activeChannelKey);
        const marketingTimeline = [
            { date: "Now", title: "POP 更新计划已建", text: "已把 6/4 替换动作与目标 POP 组合整理进页面。" },
            { date: "Next", title: "营销事件待补充", text: "后续补充每个渠道的营销动作、节奏与资源需求。" },
            { date: "Q3", title: "八月计划待确认", text: "预留给新品节奏、活动节点和后续 POP 调整。" }
        ];
        const progressTimeline = [
            { date: "NFM", title: "推进七八月 POP 快闪店机会", text: "目前交由 Kristen 策划、Raymond 辅助推进，建议把快闪店展示与新品切换节奏一起对齐。", status: "进行中", owner: "Kristen / Raymond", priority: "高优先级" },
            { date: "NFM", title: "4月已上 T921", text: "NFM 当前 POP 结构已在 4 月完成 T920 -> T921 的切换，后续重点看 sell-through 与顾客反馈。", status: "已完成", owner: "POP Update", priority: "已落地" },
            { date: "ABT", title: "已通过 2-3 月 exclusive store only 促销", text: "具体促销信息已放进 Abt 文件夹；渠道支持度较高，正在积极尝试把全旗舰、全色放进店内。", status: "已推进", owner: "ABT Channel", priority: "持续跟进" },
            { date: "BSM", title: "提高下单密度", text: "已与 buyer 达成一致，后续提高下单密度；若每个月末仍未下单，可以邮件提醒。", status: "执行中", owner: "Buyer Follow-up", priority: "中高" },
            { date: "BSM", title: "4月已上 T921", text: "BSM 当前 POP 结构已在 4 月完成 T920 -> T921 的切换，后续要继续观察新品切换后的卖出表现。", status: "已完成", owner: "POP Update", priority: "已落地" },
            { date: "BSM", title: "4/17 新增 OpenComm2 需求结构", text: "BSM 在 4/17 新增了 OpenComm2 的产品结构需求，当前已纳入渠道需求结构，后续继续跟进 buyer 对应的承接与上样节奏。", status: "新增需求", owner: "Product Structure", priority: "持续跟进" },
            { date: "RCW", title: "争取从季度下单改成月度下单", text: "当前渠道以季度性下单为主，会影响库存状况；后续可与 buyer 开会建议改为按月下单。", status: "待会议", owner: "Buyer Meeting", priority: "高优先级" },
            { date: "RCW", title: "补签独立店与 Rep（Blair）合作合同", text: "现有年度合作合同未包含 RCW，需要单独签署对应合作合同。", status: "待处理", owner: "Contract", priority: "高优先级" },
            { date: "RCW", title: "尝试整合 POP 形式", text: "目前是 4 台分散的 POP，后续可争取整合为 1 台 90cm POP。", status: "推进中", owner: "POP Upgrade", priority: "中高" },
            { date: "ALL", title: "补齐 OpenRun 2 样品提需", text: "OpenDots 2 已有样品且暂由 Kevin 保管，OpenRun 2 的提需和下一步送样动作仍需补齐。", status: "待补齐", owner: "Samples", priority: "高优先级" }
        ];
        const meetingTimeline = [
            { date: "TBD", title: "会议历史待补充", text: "后续可按时间沉淀会议纪要、决议和责任事项。" }
        ];
        const activeProgressItems = progressTimeline.filter(item => {
            const itemDate = String(item.date).toLowerCase();
            return itemDate === String(activeChannel.label).toLowerCase() || itemDate === "all";
        });
        const activeHighPriorityCount = activeProgressItems.filter(item => String(item.priority || "").includes("高")).length;
        const activePendingCount = activeProgressItems.filter(item => /待/.test(String(item.status || ""))).length;
        const activeOwnerSummary = Array.from(new Set(activeProgressItems.map(item => item.owner).filter(Boolean))).join(" / ") || "待补充";

        els.executionGrid.innerHTML = [
            "<article class=\"info-card\">",
            "<div class=\"info-card-head\"><div><p class=\"info-kicker\">营销推进</p><h3 class=\"info-card-title\">营销事件时间线</h3><p class=\"info-card-copy\">改成时间线视图，后续加活动时会比普通卡片更直观。</p></div><span class=\"info-badge\">Timeline</span></div>",
            "<div class=\"timeline-list\">",
            marketingTimeline.map(item => {
                return "<div class=\"timeline-item\"><span class=\"timeline-date\">" + escapeHtml(item.date) + "</span><div class=\"timeline-body\"><strong>" + escapeHtml(item.title) + "</strong>" + renderProgressMeta(item) + "<span>" + escapeHtml(item.text) + "</span></div></div>";
            }).join(""),
            "</div>",
            "</article>",
            "<article class=\"info-card execution-card-wide\">",
            "<div class=\"info-card-head\"><div><p class=\"info-kicker\">推进状态</p><h3 class=\"info-card-title\">当前进度 · " + escapeHtml(activeChannel.label) + "</h3><p class=\"info-card-copy\">这里会跟随当前分析渠道联动，切换渠道卡后只看该渠道的业务推进。</p></div><span class=\"info-badge\">Progress</span></div>",
            "<div class=\"timeline-list\">",
            activeProgressItems.length ? activeProgressItems.map(item => {
                return "<div class=\"timeline-item\"><span class=\"timeline-date\">" + escapeHtml(item.date) + "</span><div class=\"timeline-body\"><strong>" + escapeHtml(item.title) + "</strong>" + renderProgressMeta(item) + "<span>" + escapeHtml(item.text) + "</span></div></div>";
            }).join("") : "<p class=\"placeholder-copy\">" + escapeHtml(activeChannel.label) + " 当前暂未补充业务推进内容。</p>",
            "</div>",
            "</article>",
            "<article class=\"info-card execution-card-wide\">",
            "<div class=\"info-card-head\"><div><p class=\"info-kicker\">当前分析渠道</p><h3 class=\"info-card-title\">推进摘要 · " + escapeHtml(activeChannel.label) + "</h3><p class=\"info-card-copy\">把当前分析渠道的推进重点、待办密度和 owner 集中显示，方便你在一个屏幕内判断优先级。</p></div><span class=\"info-badge\">Linked</span></div>",
            "<div class=\"business-metric-grid mix-watch-grid\">",
            renderMixWatchCard("推进事项", formatNumber(activeProgressItems.length), "当前分析渠道正在跟进的业务动作数量"),
            renderMixWatchCard("高优先级", formatNumber(activeHighPriorityCount), "优先级标记里包含“高”的事项"),
            renderMixWatchCard("待处理", formatNumber(activePendingCount), "当前状态里带“待”的事项"),
            renderMixWatchCard("关键 Owner", activeOwnerSummary, "切换渠道后这里会同步更新"),
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
        renderTopActionGrid();
        renderFullProductMixPanel();
        renderPopPlanPanel();
        renderLaunchPlanPanel();
        renderBusinessSupportGrid();
        renderChannelStrategyGrid();
        renderExecutionGrid();
        renderPromoCalendarPanel();
    }

    function renderBusinessOverview() {
        if (!els.businessOverviewGrid) {
            return;
        }
        els.businessOverviewGrid.innerHTML = [
            "<article class=\"info-card business-summary-card\">",
            "<div class=\"info-card-head\"><div><p class=\"info-kicker\">Overview</p><h3 class=\"info-card-title\">渠道总表</h3><p class=\"info-card-copy\">先用总表快速横向比，再展开单渠道详情看具体商务、POP 和产品结构。</p></div><span class=\"info-badge\">Summary</span></div>",
            "<div class=\"table-wrap\">",
            "<table class=\"business-summary-table\">",
            "<thead><tr><th>渠道</th><th>门店</th><th>商务摘要</th><th>POP形式</th><th>卖出SKU</th><th>累计销量</th><th>销售额</th><th>详情</th></tr></thead>",
            "<tbody>",
            CONFIG.channels.map(channel => {
                const profile = CONFIG.channelProfiles[channel.key];
                const dashboard = state.dashboards[channel.key];
                const productMix = dashboard.productMixSinceStart;
                return [
                    "<tr style=\"" + channelStyle(channel) + "\">",
                    "<td><strong>" + escapeHtml(profile.displayName || channel.label) + "</strong><div class=\"table-subcopy\">最新月份 " + escapeHtml(dashboard.latestMonthKey) + "</div></td>",
                    "<td>" + escapeHtml(profile.storeSummary || "—") + "</td>",
                    "<td>" + escapeHtml(joinBusinessTags(profile.businessTags)) + "<div class=\"table-subcopy\">" + escapeHtml(firstContactLine(profile.buyer)) + "</div></td>",
                    "<td>" + escapeHtml(profile.popFormatLabel || profile.popSize) + "</td>",
                    "<td>" + formatNumber(productMix.items.length) + "</td>",
                    "<td>" + formatNumber(productMix.totalQty) + "</td>",
                    "<td>" + formatCurrency(productMix.totalSales) + "</td>",
                    "<td><button class=\"business-detail-btn\" type=\"button\" data-business-channel=\"" + escapeHtml(channel.key) + "\">查看</button></td>",
                    "</tr>"
                ].join("");
            }).join(""),
            "</tbody></table></div></article>",
            "<div class=\"business-detail-accordion\">",
            CONFIG.channels.map(channel => {
                const profile = CONFIG.channelProfiles[channel.key];
                const dashboard = state.dashboards[channel.key];
                return renderBusinessDetailPanel(channel, profile, dashboard);
            }).join(""),
            "</div>"
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
            const target2026 = CONFIG.channelTargets[channel.key] || { qty: 0, sales: 0 };
            const annualQtyYoY2025 = calculateGrowth(dashboard.yearlyTotals[2025].qty, dashboard.yearlyTotals[2024].qty);
            const annualSalesYoY2025 = calculateGrowth(dashboard.yearlyTotals[2025].sales, dashboard.yearlyTotals[2024].sales);
            const targetQtyYoY2026 = calculateGrowth(target2026.qty, dashboard.yearlyTotals[2025].qty);
            const targetSalesYoY2026 = calculateGrowth(target2026.sales, dashboard.yearlyTotals[2025].sales);
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
                "<td>" + formatNumber(target2026.qty) + "</td>",
                "<td>" + formatCurrency(target2026.sales) + "</td>",
                "<td>" + renderDelta(targetQtyYoY2026) + "</td>",
                "<td>" + renderDelta(targetSalesYoY2026) + "</td>",
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

    function renderSidebarChannelSwitch() {
        if (!els.sidebarChannelSwitch) {
            return;
        }
        els.sidebarChannelSwitch.innerHTML = CONFIG.channels.map(channel => {
            const dashboard = state.dashboards[channel.key];
            const activeClass = state.activeChannelKey === channel.key ? " active" : "";
            const metaTop = dashboard
                ? "2026 YTD " + formatCurrency(dashboard.samePeriodByYear[2026].sales)
                : "Loading";
            const metaBottom = dashboard
                ? ("最新 " + dashboard.latestMonthKey + " · " + formatCurrency(dashboard.latestMonth.sales))
                : "同步中";
            return [
                "<button class=\"sidebar-channel-btn" + activeClass + "\" type=\"button\" data-sidebar-channel=\"" + channel.key + "\" style=\"" + channelStyle(channel) + "\">",
                "<span class=\"sidebar-channel-label\">" + escapeHtml(channel.label) + "</span>",
                "<span class=\"sidebar-channel-meta\">" + escapeHtml(metaTop) + "</span>",
                "<span class=\"sidebar-channel-meta\">" + escapeHtml(metaBottom) + "</span>",
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
        renderSidebarChannelSwitch();
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
        if (els.sidebarChannelSwitch) {
            els.sidebarChannelSwitch.innerHTML = "";
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
        setActiveChannel(card.dataset.channel);
    });

    els.channelTabs.addEventListener("click", event => {
        const tab = event.target.closest("[data-channel]");
        if (!tab) {
            return;
        }
        setActiveChannel(tab.dataset.channel);
    });

    if (els.sidebarChannelSwitch) {
        els.sidebarChannelSwitch.addEventListener("click", event => {
            const button = event.target.closest("[data-sidebar-channel]");
            if (!button) {
                return;
            }
            setActiveChannel(button.dataset.sidebarChannel);
        });
    }

    els.promoCalendarPanel.addEventListener("click", event => {
        const monthButton = event.target.closest("[data-promo-month]");
        if (monthButton) {
            const nextMonth = monthButton.dataset.promoMonth || "all";
            if (state.selectedPromoMonth !== nextMonth) {
                state.selectedPromoMonth = nextMonth;
                renderPromoCalendarPanel();
            }
            return;
        }

        const channelButton = event.target.closest("[data-promo-channel]");
        if (!channelButton) {
            return;
        }
        const nextChannel = channelButton.dataset.promoChannel || "all";
        if (state.selectedPromoChannel === nextChannel) {
            return;
        }
        state.selectedPromoChannel = nextChannel;
        renderPromoCalendarPanel();
    });

    els.fullProductMixPanel.addEventListener("click", event => {
        const button = event.target.closest("[data-mix-channel][data-mix-base]");
        if (!button) {
            return;
        }
        const key = mixExpansionKey(button.dataset.mixChannel, button.dataset.mixBase);
        state.expandedMixBases[key] = !state.expandedMixBases[key];
        renderFullProductMixPanel();
    });

    els.businessOverviewGrid.addEventListener("click", event => {
        const button = event.target.closest("[data-business-channel]");
        if (!button) {
            return;
        }
        const nextChannel = button.dataset.businessChannel;
        if (!nextChannel) {
            return;
        }
        setActiveChannel(nextChannel, { scrollToBusinessDetail: true });
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
