const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
        Header, Footer, AlignmentType, LevelFormat, HeadingLevel,
        BorderStyle, WidthType, ShadingType, PageNumber, TabStopType, TabStopPosition } = require('docx');
const fs = require('fs');

// 边框样式
const border = { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" };
const borders = { top: border, bottom: border, left: border, right: border };
const noBorder = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const noBorders = { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder };

// 表头样式
const headerShading = { fill: "1E4D78", type: ShadingType.CLEAR };
const altRowShading = { fill: "F5F8FA", type: ShadingType.CLEAR };

// 创建表格单元格
function createCell(text, isHeader = false, width = 2000, isBold = false) {
    return new TableCell({
        borders,
        width: { size: width, type: WidthType.DXA },
        shading: isHeader ? headerShading : { fill: "FFFFFF", type: ShadingType.CLEAR },
        margins: { top: 80, bottom: 80, left: 120, right: 120 },
        children: [new Paragraph({
            children: [new TextRun({
                text: text,
                bold: isBold || isHeader,
                color: isHeader ? "FFFFFF" : "333333",
                size: 22
            })]
        })]
    });
}

// 创建带底边框的段落（分隔线效果）
function createSeparatorParagraph(text = "", isBold = false) {
    return new Paragraph({
        border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: "1E4D78", space: 1 } },
        spacing: { before: 200, after: 200 },
        children: [new TextRun({ text: text, bold: isBold, size: 28, color: "1E4D78" })]
    });
}

// 创建普通段落
function createParagraph(text, isBold = false, indent = 0) {
    return new Paragraph({
        spacing: { before: 100, after: 100 },
        indent: indent ? { left: indent } : undefined,
        children: [new TextRun({ text: text, bold: isBold, size: 22 })]
    });
}

// 创建带编号的段落
function createNumberedParagraph(num, text, indent = 0) {
    return new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        spacing: { before: 60, after: 60 },
        indent: indent ? { left: indent } : { left: 720, hanging: 360 },
        children: [new TextRun({ text: text, size: 22 })]
    });
}

// 创建带项目符号的段落
function createBulletParagraph(text, bold = false, indent = 0) {
    return new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        spacing: { before: 60, after: 60 },
        indent: indent ? { left: indent } : { left: 720, hanging: 360 },
        children: [new TextRun({ text: text, bold: bold, size: 22 })]
    });
}

// 创建引用段落
function createQuoteParagraph(text) {
    return new Paragraph({
        spacing: { before: 120, after: 120 },
        indent: { left: 720, right: 720 },
        border: { left: { style: BorderStyle.SINGLE, size: 12, color: "1E4D78", space: 10 } },
        children: [new TextRun({ text: text, size: 22, italics: true })]
    });
}

// 创建勾选框段落
function createCheckBoxParagraph(text, week) {
    return new Paragraph({
        spacing: { before: 80, after: 80 },
        children: [
            new TextRun({ text: "☐ ", size: 22 }),
            new TextRun({ text: text, size: 22 })
        ]
    });
}

// 创建标题1
function createHeading1(text) {
    return new Paragraph({
        heading: HeadingLevel.HEADING_1,
        spacing: { before: 360, after: 180 },
        children: [new TextRun({ text: text, bold: true, size: 36, color: "1E4D78" })]
    });
}

// 创建标题2
function createHeading2(text) {
    return new Paragraph({
        heading: HeadingLevel.HEADING_2,
        spacing: { before: 280, after: 140 },
        children: [new TextRun({ text: text, bold: true, size: 28, color: "333333" })]
    });
}

// 创建标题3
function createHeading3(text) {
    return new Paragraph({
        heading: HeadingLevel.HEADING_3,
        spacing: { before: 200, after: 100 },
        children: [new TextRun({ text: text, bold: true, size: 24, color: "555555" })]
    });
}

const doc = new Document({
    styles: {
        default: {
            document: {
                run: { font: "Microsoft YaHei", size: 22 }
            }
        },
        paragraphStyles: [
            {
                id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
                run: { size: 36, bold: true, font: "Microsoft YaHei", color: "1E4D78" },
                paragraph: { spacing: { before: 360, after: 180 }, outlineLevel: 0 }
            },
            {
                id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
                run: { size: 28, bold: true, font: "Microsoft YaHei", color: "333333" },
                paragraph: { spacing: { before: 280, after: 140 }, outlineLevel: 1 }
            },
            {
                id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
                run: { size: 24, bold: true, font: "Microsoft YaHei", color: "555555" },
                paragraph: { spacing: { before: 200, after: 100 }, outlineLevel: 2 }
            }
        ]
    },
    numbering: {
        config: [
            {
                reference: "bullets",
                levels: [{
                    level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
                    style: { paragraph: { indent: { left: 720, hanging: 360 } } }
                }]
            },
            {
                reference: "numbers",
                levels: [{
                    level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT,
                    style: { paragraph: { indent: { left: 720, hanging: 360 } } }
                }]
            }
        ]
    },
    sections: [{
        properties: {
            page: {
                size: { width: 11906, height: 16838 }, // A4
                margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 }
            }
        },
        headers: {
            default: new Header({
                children: [new Paragraph({
                    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: "1E4D78", space: 4 } },
                    children: [new TextRun({ text: "独居老人安全监护：低成本落地方案", size: 20, color: "666666" })]
                })]
            })
        },
        footers: {
            default: new Footer({
                children: [new Paragraph({
                    border: { top: { style: BorderStyle.SINGLE, size: 6, color: "CCCCCC", space: 4 } },
                    alignment: AlignmentType.CENTER,
                    tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }],
                    children: [
                        new TextRun({ text: "养老行业赛道分析", size: 18, color: "999999" }),
                        new TextRun({ text: "\t第 ", size: 18, color: "999999" }),
                        new TextRun({ children: [PageNumber.CURRENT], size: 18, color: "999999" }),
                        new TextRun({ text: " 页", size: 18, color: "999999" })
                    ]
                })]
            })
        },
        children: [
            // 封面标题
            new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { before: 2000, after: 400 },
                children: [new TextRun({ text: "独居老人安全监护", bold: true, size: 56, color: "1E4D78" })]
            }),
            new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { before: 0, after: 800 },
                children: [new TextRun({ text: "低成本落地方案", bold: true, size: 44, color: "333333" })]
            }),

            // 分隔线
            new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { before: 400, after: 2000 },
                border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: "1E4D78", space: 1 } },
                children: []
            }),

            // 第一部分
            createHeading1("一、项目定位"),
            createParagraph("聚焦独居老人居家安全监护，通过轻量化设备+SaaS服务，让子女实时掌握老人状态，老人发生跌倒等意外时秒级报警通知子女和社区。"),

            // 分隔线
            new Paragraph({ spacing: { before: 400, after: 200 }, border: {}, children: [] }),

            // 第二部分
            createHeading1("二、产品方案设计"),
            createHeading2("2.1 硬件配置（三种套餐）"),

            // 硬件套餐表格
            new Table({
                width: { size: 9026, type: WidthType.DXA },
                columnWidths: [1805, 3611, 1805, 1805],
                rows: [
                    new TableRow({
                        children: [
                            createCell("套餐", true, 1805),
                            createCell("设备组成", true, 3611),
                            createCell("硬件成本", true, 1805),
                            createCell("适合场景", true, 1805)
                        ]
                    }),
                    new TableRow({
                        children: [
                            createCell("基础版", false, 1805),
                            createCell("智能摄像头（客厅）+ SOS手环", false, 3611),
                            createCell("400-600元", false, 1805),
                            createCell("预算有限、子女陪伴较多", false, 1805)
                        ]
                    }),
                    new TableRow({
                        children: [
                            createCell("标准版", false, 1805, true),
                            createCell("基础版 + 毫米波雷达（卫生间/浴室）+ 门磁传感器", false, 3611, true),
                            createCell("1000-1500元", false, 1805, true),
                            createCell("推荐主推，覆盖核心场景", false, 1805, true)
                        ]
                    }),
                    new TableRow({
                        children: [
                            createCell("尊享版", false, 1805),
                            createCell("标准版 + 智能手环（健康监测）+ 烟雾/燃气报警器", false, 3611),
                            createCell("2000-3000元", false, 1805),
                            createCell("高净值家庭/失能半失能老人", false, 1805)
                        ]
                    })
                ]
            }),

            new Paragraph({ spacing: { before: 300, after: 100 }, children: [] }),
            createHeading2("2.2 核心功能"),

            // 核心功能表格
            new Table({
                width: { size: 9026, type: WidthType.DXA },
                columnWidths: [2257, 3385, 3384],
                rows: [
                    new TableRow({
                        children: [
                            createCell("功能", true, 2257),
                            createCell("技术实现", true, 3385),
                            createCell("价值", true, 3384)
                        ]
                    }),
                    new TableRow({
                        children: [
                            createCell("跌倒检测", false, 2257),
                            createCell("毫米波雷达（无摄像头侵犯隐私）", false, 3385),
                            createCell("老人跌倒后自动报警，救命关键", false, 3384)
                        ]
                    }),
                    new TableRow({
                        children: [
                            createCell("活动监测", false, 2257),
                            createCell("摄像头AI分析+红外传感器", false, 3385),
                            createCell("长时间无活动自动预警", false, 3384)
                        ]
                    }),
                    new TableRow({
                        children: [
                            createCell("一键SOS", false, 2257),
                            createCell("手环/摄像头物理按键", false, 3385),
                            createCell("老人主动求助，秒达子女", false, 3384)
                        ]
                    }),
                    new TableRow({
                        children: [
                            createCell("实时查看", false, 2257),
                            createCell("摄像头+手机APP", false, 3385),
                            createCell("子女随时查看老人状态", false, 3384)
                        ]
                    }),
                    new TableRow({
                        children: [
                            createCell("健康数据", false, 2257),
                            createCell("智能手环", false, 3385),
                            createCell("心率、血压、睡眠数据", false, 3384)
                        ]
                    })
                ]
            }),

            new Paragraph({ spacing: { before: 300, after: 100 }, children: [] }),
            createHeading2("2.3 软件平台"),
            createBulletParagraph("SaaS管理后台：子女端APP（报警推送+实时查看+健康数据）"),
            createBulletParagraph("老人端设备：配对绑定、报警触发"),
            createBulletParagraph("运营端：社区/物业管理视图、数据统计"),

            // 分隔线
            new Paragraph({ spacing: { before: 400, after: 200 }, border: {}, children: [] }),

            // 第三部分
            createHeading1("三、盈利模式"),
            createHeading2("3.1 收入结构"),

            // 收入结构表格
            new Table({
                width: { size: 9026, type: WidthType.DXA },
                columnWidths: [3009, 3008, 3009],
                rows: [
                    new TableRow({
                        children: [
                            createCell("收入来源", true, 3009),
                            createCell("定价策略", true, 3008),
                            createCell("月收入/用户", true, 3009)
                        ]
                    }),
                    new TableRow({
                        children: [
                            createCell("设备销售", false, 3009),
                            createCell("硬件成本价+20-30%毛利", false, 3008),
                            createCell("一次性300-800元", false, 3009)
                        ]
                    }),
                    new TableRow({
                        children: [
                            createCell("订阅服务", false, 3009, true),
                            createCell("月/年订阅（必选）", false, 3008, true),
                            createCell("59-99元/月，499-899元/年", false, 3009, true)
                        ]
                    }),
                    new TableRow({
                        children: [
                            createCell("增值服务", false, 3009),
                            createCell("上门服务对接/健康管理", false, 3008),
                            createCell("额外20-30%抽佣", false, 3009)
                        ]
                    })
                ]
            }),

            new Paragraph({ spacing: { before: 200, after: 100 }, children: [] }),
            createParagraph("核心收入来源是订阅服务，硬件平进平出或微利均可。", true),

            new Paragraph({ spacing: { before: 300, after: 100 }, children: [] }),
            createHeading2("3.2 测算模型"),

            // 测算模型表格
            new Table({
                width: { size: 9026, type: WidthType.DXA },
                columnWidths: [4513, 4513],
                rows: [
                    new TableRow({
                        children: [
                            createCell("指标", true, 4513),
                            createCell("数据", true, 4513)
                        ]
                    }),
                    new TableRow({
                        children: [
                            createCell("单用户ARPU", false, 4513),
                            createCell("约800元/年（含硬件分摊）", false, 4513)
                        ]
                    }),
                    new TableRow({
                        children: [
                            createCell("硬件成本", false, 4513),
                            createCell("1000-1500元/套", false, 4513)
                        ]
                    }),
                    new TableRow({
                        children: [
                            createCell("毛利空间", false, 4513),
                            createCell("订阅服务贡献主要利润", false, 4513)
                        ]
                    }),
                    new TableRow({
                        children: [
                            createCell("盈亏平衡", false, 4513),
                            createCell("50-80个年付费用户", false, 4513)
                        ]
                    })
                ]
            }),

            // 分隔线
            new Paragraph({ spacing: { before: 400, after: 200 }, border: {}, children: [] }),

            // 第四部分
            createHeading1("四、低成本启动策略"),
            createHeading2("4.1 阶段一：最小MVP验证（预算2-5万）"),
            createParagraph("目标：跑通服务闭环，验证市场需求"),

            // MVP阶段表格
            new Table({
                width: { size: 9026, type: WidthType.DXA },
                columnWidths: [3009, 2006, 4011],
                rows: [
                    new TableRow({
                        children: [
                            createCell("任务", true, 3009),
                            createCell("预算", true, 2006),
                            createCell("说明", true, 4011)
                        ]
                    }),
                    new TableRow({
                        children: [
                            createCell("采购硬件样品", false, 3009),
                            createCell("3000-5000元", false, 2006),
                            createCell("采购2-3套设备测试", false, 4011)
                        ]
                    }),
                    new TableRow({
                        children: [
                            createCell("开发/购买SaaS平台", false, 3009),
                            createCell("5000-20000元", false, 2006),
                            createCell("低代码SaaS模板或现成解决方案", false, 4011)
                        ]
                    }),
                    new TableRow({
                        children: [
                            createCell("地推获客", false, 3009),
                            createCell("5000-10000元", false, 2006),
                            createCell("2-3个小区试运营", false, 4011)
                        ]
                    }),
                    new TableRow({
                        children: [
                            createCell("运营备用金", false, 3009),
                            createCell("5000元", false, 2006),
                            createCell("处理突发情况", false, 4011)
                        ]
                    })
                ]
            }),

            new Paragraph({ spacing: { before: 200, after: 100 }, children: [] }),
            createParagraph("执行要点：", true),
            createBulletParagraph("先用现成设备（小米/萤石摄像头+第三方SOS设备）快速验证"),
            createBulletParagraph("软件先用SaaS模板或低代码平台搭建"),
            createBulletParagraph("选1个小区、10-20户老人开始试运营"),

            new Paragraph({ spacing: { before: 300, after: 100 }, children: [] }),
            createHeading2("4.2 阶段二：模式复制（预算10-20万）"),
            createParagraph("目标：形成可复制的社区运营模式"),
            createBulletParagraph("开发/完善自有SaaS平台"),
            createBulletParagraph("拓展至3-5个社区"),
            createBulletParagraph("建立服务响应机制（与社区/物业合作）"),
            createBulletParagraph("申请政府养老服务采购资质"),

            new Paragraph({ spacing: { before: 300, after: 100 }, children: [] }),
            createHeading2("4.3 阶段三：规模化（预算30-50万）"),
            createParagraph("目标：跨区域复制，对接政府资源"),
            createBulletParagraph("申请智慧养老示范项目"),
            createBulletParagraph("接入民政局养老服务平台"),
            createBulletParagraph("发展城市合伙人/代理商"),
            createBulletParagraph("接入医保/商业保险"),

            // 分隔线
            new Paragraph({ spacing: { before: 400, after: 200 }, border: {}, children: [] }),

            // 第五部分
            createHeading1("五、获客渠道"),
            createHeading2("5.1 低成本获客路径"),

            // 获客渠道表格
            new Table({
                width: { size: 9026, type: WidthType.DXA },
                columnWidths: [3009, 1504, 1504, 3009],
                rows: [
                    new TableRow({
                        children: [
                            createCell("渠道", true, 3009),
                            createCell("成本", true, 1504),
                            createCell("转化率", true, 1504),
                            createCell("优先级", true, 3009)
                        ]
                    }),
                    new TableRow({
                        children: [
                            createCell("社区居委会/物业合作", false, 3009),
                            createCell("低", false, 1504),
                            createCell("高", false, 1504),
                            createCell("⭐⭐⭐", false, 3009)
                        ]
                    }),
                    new TableRow({
                        children: [
                            createCell("社区微信群推广", false, 3009),
                            createCell("极低", false, 1504),
                            createCell("中", false, 1504),
                            createCell("⭐⭐⭐", false, 3009)
                        ]
                    }),
                    new TableRow({
                        children: [
                            createCell("抖音/小红书（子女内容）", false, 3009),
                            createCell("低", false, 1504),
                            createCell("中高", false, 1504),
                            createCell("⭐⭐", false, 3009)
                        ]
                    }),
                    new TableRow({
                        children: [
                            createCell("医院/社区卫生中心导流", false, 3009),
                            createCell("中", false, 1504),
                            createCell("高", false, 1504),
                            createCell("⭐⭐", false, 3009)
                        ]
                    }),
                    new TableRow({
                        children: [
                            createCell("政府采购/民政采购", false, 3009),
                            createCell("中", false, 1504),
                            createCell("高", false, 1504),
                            createCell("⭐⭐⭐", false, 3009)
                        ]
                    }),
                    new TableRow({
                        children: [
                            createCell("老龄委/老年协会合作", false, 3009),
                            createCell("低", false, 1504),
                            createCell("中", false, 1504),
                            createCell("⭐⭐", false, 3009)
                        ]
                    })
                ]
            }),

            new Paragraph({ spacing: { before: 300, after: 100 }, children: [] }),
            createHeading2("5.2 核心获客话术"),
            createParagraph("针对子女：", true),
            createQuoteParagraph("只需一部手机的钱，就能24小时守护父母安全。老人跌倒、异常活动实时推送，比请保姆更划算。"),
            new Paragraph({ spacing: { before: 150, after: 100 }, children: [] }),
            createParagraph("针对社区/物业：", true),
            createQuoteParagraph("帮助社区建立独居老人安全档案，减少意外事故纠纷，降低管理风险。"),

            // 分隔线
            new Paragraph({ spacing: { before: 400, after: 200 }, border: {}, children: [] }),

            // 第六部分
            createHeading1("六、政策利用"),

            // 政策表格
            new Table({
                width: { size: 9026, type: WidthType.DXA },
                columnWidths: [3009, 3008, 3009],
                rows: [
                    new TableRow({
                        children: [
                            createCell("政策", true, 3009),
                            createCell("内容", true, 3008),
                            createCell("如何利用", true, 3009)
                        ]
                    }),
                    new TableRow({
                        children: [
                            createCell("家庭养老床位建设", false, 3009),
                            createCell("政府补贴3000-5000元/床", false, 3008),
                            createCell("对接民政局，进入采购名单", false, 3009)
                        ]
                    }),
                    new TableRow({
                        children: [
                            createCell("高龄独居老人补贴", false, 3009),
                            createCell("各地200-500元/月不等", false, 3008),
                            createCell("为老人申请补贴，降低付费门槛", false, 3009)
                        ]
                    }),
                    new TableRow({
                        children: [
                            createCell("智慧养老示范项目", false, 3009),
                            createCell("一次性补贴5-20万", false, 3008),
                            createCell("申报成为示范点", false, 3009)
                        ]
                    }),
                    new TableRow({
                        children: [
                            createCell("适老化改造补贴", false, 3009),
                            createCell("各地1000-5000元/户", false, 3008),
                            createCell("将设备纳入适老化改造包", false, 3009)
                        ]
                    })
                ]
            }),

            new Paragraph({ spacing: { before: 200, after: 100 }, children: [] }),
            createParagraph("建议：主动联系当地民政局养老服务科，了解本地补贴政策。", true),

            // 分隔线
            new Paragraph({ spacing: { before: 400, after: 200 }, border: {}, children: [] }),

            // 第七部分
            createHeading1("七、关键成功因素"),
            createNumberedParagraph(1, "报警准确率>95%：误报太多会被子女卸载"),
            createNumberedParagraph(2, "7×24小时响应：报警后5分钟内必须响应"),
            createNumberedParagraph(3, "本地化服务网络：与社区/物业深度合作"),
            createNumberedParagraph(4, "子女体验优先：APP要简洁、推送要及时"),

            // 分隔线
            new Paragraph({ spacing: { before: 400, after: 200 }, border: {}, children: [] }),

            // 第八部分
            createHeading1("八、风险与应对"),

            // 风险表格
            new Table({
                width: { size: 9026, type: WidthType.DXA },
                columnWidths: [4513, 4513],
                rows: [
                    new TableRow({
                        children: [
                            createCell("风险", true, 4513),
                            createCell("应对策略", true, 4513)
                        ]
                    }),
                    new TableRow({
                        children: [
                            createCell("误报率高导致用户流失", false, 4513),
                            createCell("选用高质量毫米波雷达，多次算法调优", false, 4513)
                        ]
                    }),
                    new TableRow({
                        children: [
                            createCell("老人不愿使用设备", false, 4513),
                            createCell("强调\"不侵犯隐私\"，上门帮老人适应", false, 4513)
                        ]
                    }),
                    new TableRow({
                        children: [
                            createCell("竞争加剧", false, 4513),
                            createCell("深耕本地化服务，建立社区关系壁垒", false, 4513)
                        ]
                    }),
                    new TableRow({
                        children: [
                            createCell("政策变化", false, 4513),
                            createCell("多元化收入，不依赖单一补贴来源", false, 4513)
                        ]
                    })
                ]
            }),

            // 分隔线
            new Paragraph({ spacing: { before: 400, after: 200 }, border: {}, children: [] }),

            // 第九部分
            createHeading1("九、下一步行动清单"),
            createCheckBoxParagraph("第一周：调研本地市场，走访2-3个社区了解需求"),
            createCheckBoxParagraph("第二周：采购样品设备，实测3-5款产品"),
            createCheckBoxParagraph("第三周：搭建最小MVP（微信小程序+现成设备）"),
            createCheckBoxParagraph("第四周：在1个小区招募10户免费试用"),
            createCheckBoxParagraph("第一个月后：收集反馈，优化服务，准备正式推广"),

            // 结束语
            new Paragraph({ spacing: { before: 500, after: 200 }, border: {}, children: [] }),
            new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { before: 400, after: 400 },
                border: { top: { style: BorderStyle.SINGLE, size: 6, color: "CCCCCC", space: 8 } },
                children: [new TextRun({ text: "如需更详细的某一部分（如技术方案、软件平台选型、财务测算模型），可以继续深挖。", size: 20, color: "666666", italics: true })]
            })
        ]
    }]
});

// 生成文档
Packer.toBuffer(doc).then(buffer => {
    fs.writeFileSync('/Users/likaiying/WorkBuddy/20260401111529/独居老人安全监护_落地方案.docx', buffer);
    console.log('Word文档已生成：独居老人安全监护_落地方案.docx');
});
