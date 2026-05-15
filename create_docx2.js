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
const headerShading = { fill: "2E7D32", type: ShadingType.CLEAR };

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
                size: 20
            })]
        })]
    });
}

// 创建标题1
function createHeading1(text) {
    return new Paragraph({
        heading: HeadingLevel.HEADING_1,
        spacing: { before: 360, after: 180 },
        children: [new TextRun({ text: text, bold: true, size: 36, color: "2E7D32" })]
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

// 创建普通段落
function createParagraph(text, isBold = false, indent = 0) {
    return new Paragraph({
        spacing: { before: 100, after: 100 },
        indent: indent ? { left: indent } : undefined,
        children: [new TextRun({ text: text, bold: isBold, size: 22 })]
    });
}

// 创建带项目符号的段落
function createBulletParagraph(text, bold = false, indent = 720) {
    return new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        spacing: { before: 60, after: 60 },
        indent: { left: indent, hanging: 360 },
        children: [new TextRun({ text: text, bold: bold, size: 22 })]
    });
}

// 创建带编号的段落
function createNumberedParagraph(num, text, bold = false) {
    return new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        spacing: { before: 60, after: 60 },
        indent: { left: 720, hanging: 360 },
        children: [new TextRun({ text: text, bold: bold, size: 22 })]
    });
}

// 创建引用段落
function createQuoteParagraph(text) {
    return new Paragraph({
        spacing: { before: 120, after: 120 },
        indent: { left: 720, right: 720 },
        border: { left: { style: BorderStyle.SINGLE, size: 12, color: "2E7D32", space: 10 } },
        children: [new TextRun({ text: text, size: 22, italics: true })]
    });
}

// 创建勾选框段落
function createCheckBoxParagraph(text) {
    return new Paragraph({
        spacing: { before: 80, after: 80 },
        children: [
            new TextRun({ text: "☐ ", size: 22 }),
            new TextRun({ text: text, size: 22 })
        ]
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
                run: { size: 36, bold: true, font: "Microsoft YaHei", color: "2E7D32" },
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
                size: { width: 11906, height: 16838 },
                margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 }
            }
        },
        headers: {
            default: new Header({
                children: [new Paragraph({
                    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: "2E7D32", space: 4 } },
                    children: [new TextRun({ text: "居家养老服务：低成本落地方案", size: 20, color: "666666" })]
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
                        new TextRun({ text: "居家养老服务创业指南", size: 18, color: "999999" }),
                        new TextRun({ text: "\t第 ", size: 18, color: "999999" }),
                        new TextRun({ children: [PageNumber.CURRENT], size: 18, color: "999999" }),
                        new TextRun({ text: " 页", size: 18, color: "999999" })
                    ]
                })]
            })
        },
        children: [
            // 封面
            new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { before: 2000, after: 400 },
                children: [new TextRun({ text: "居家养老服务", bold: true, size: 56, color: "2E7D32" })]
            }),
            new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { before: 0, after: 800 },
                children: [new TextRun({ text: "低成本落地方案", bold: true, size: 44, color: "333333" })]
            }),
            new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { before: 400, after: 2000 },
                border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: "2E7D32", space: 1 } },
                children: []
            }),

            // 第一部分
            createHeading1("一、项目定位"),
            createParagraph("聚焦社区居家养老服务这一万亿级蓝海市场，通过搭建需求对接+服务交付+质量管控的轻资产平台，为社区老人提供专业化、标准化、人性化的上门养老服务，让老人在家安心养老、子女安心工作。"),

            // 第二部分
            createHeading1("二、市场机遇"),
            createHeading2("2.1 市场空间"),
            new Table({
                width: { size: 9026, type: WidthType.DXA },
                columnWidths: [4513, 4513],
                rows: [
                    new TableRow({ children: [createCell("指标", true, 4513), createCell("数据", true, 4513)] }),
                    new TableRow({ children: [createCell("60岁以上人口", false, 4513), createCell("3.1亿（2024年底）", false, 4513)] }),
                    new TableRow({ children: [createCell("选择居家养老比例", false, 4513), createCell("90%", false, 4513)] }),
                    new TableRow({ children: [createCell("银发经济市场规模", false, 4513), createCell("超6万亿（2025年预测）", false, 4513)] }),
                    new TableRow({ children: [createCell("2026年政策支持", false, 4513), createCell("国家密集出台\"9073\"养老体系政策", false, 4513)] })
                ]
            }),

            new Paragraph({ spacing: { before: 300, after: 100 }, children: [] }),
            createHeading2("2.2 核心痛点"),
            new Table({
                width: { size: 9026, type: WidthType.DXA },
                columnWidths: [3009, 6017],
                rows: [
                    new TableRow({ children: [createCell("痛点", true, 3009), createCell("描述", true, 6017)] }),
                    new TableRow({ children: [createCell("供给不足", false, 3009), createCell("专业护理人员缺口超500万", false, 6017)] }),
                    new TableRow({ children: [createCell("信任缺失", false, 3009), createCell("老人/子女不放心陌生人上门", false, 6017)] }),
                    new TableRow({ children: [createCell("价格混乱", false, 3009), createCell("市场定价不透明，标准缺失", false, 6017)] }),
                    new TableRow({ children: [createCell("质量参差", false, 3009), createCell("服务质量参差不齐，难评价", false, 6017)] }),
                    new TableRow({ children: [createCell("响应慢", false, 3009), createCell("紧急需求无法快速匹配", false, 6017)] })
                ]
            }),

            new Paragraph({ spacing: { before: 300, after: 100 }, children: [] }),
            createHeading2("2.3 政策红利"),
            createBulletParagraph("政府购买服务：各地政府大量采购居家养老服务"),
            createBulletParagraph("税收优惠：企业所得税\"三免三减半\""),
            createBulletParagraph("补贴叠加：高龄/失能老人服务补贴200-600元/月"),
            createBulletParagraph("准入简化：取消养老服务部分资质门槛"),

            // 第三部分
            createHeading1("三、服务产品体系"),
            createHeading2("A. 生活照料类（高频刚需）"),
            new Table({
                width: { size: 9026, type: WidthType.DXA },
                columnWidths: [3009, 2006, 4011],
                rows: [
                    new TableRow({ children: [createCell("服务项目", true, 3009), createCell("单价", true, 2006), createCell("说明", true, 4011)] }),
                    new TableRow({ children: [createCell("助洁服务", false, 3009), createCell("35-50元/小时", false, 2006), createCell("打扫卫生、收纳整理", false, 4011)] }),
                    new TableRow({ children: [createCell("助餐服务", false, 3009), createCell("按次收费", false, 2006), createCell("送餐上门或助餐", false, 4011)] }),
                    new TableRow({ children: [createCell("助浴服务", false, 3009), createCell("80-150元/次", false, 2006), createCell("上门助浴（含安全防护）", false, 4011)] }),
                    new TableRow({ children: [createCell("助行服务", false, 3009), createCell("50-80元/次", false, 2006), createCell("陪同外出、散步", false, 4011)] }),
                    new TableRow({ children: [createCell("代购代办", false, 3009), createCell("30-50元/次", false, 2006), createCell("买菜、买药、缴费", false, 4011)] })
                ]
            }),

            new Paragraph({ spacing: { before: 300, after: 100 }, children: [] }),
            createHeading2("B. 医疗护理类（高价值）"),
            new Table({
                width: { size: 9026, type: WidthType.DXA },
                columnWidths: [3009, 2006, 4011],
                rows: [
                    new TableRow({ children: [createCell("服务项目", true, 3009), createCell("单价", true, 2006), createCell("说明", true, 4011)] }),
                    new TableRow({ children: [createCell("陪诊就医", false, 3009), createCell("120-200元/次", false, 2006), createCell("全程陪同看病", false, 4011)] }),
                    new TableRow({ children: [createCell("康复护理", false, 3009), createCell("80-150元/小时", false, 2006), createCell("术后康复指导", false, 4011)] }),
                    new TableRow({ children: [createCell("压疮护理", false, 3009), createCell("100-180元/次", false, 2006), createCell("专业护理", false, 4011)] }),
                    new TableRow({ children: [createCell("管道护理", false, 3009), createCell("150-250元/次", false, 2006), createCell("鼻饲管、导尿管等", false, 4011)] }),
                    new TableRow({ children: [createCell("健康监测", false, 3009), createCell("30-50元/次", false, 2006), createCell("血压血糖测量", false, 4011)] })
                ]
            }),

            new Paragraph({ spacing: { before: 300, after: 100 }, children: [] }),
            createHeading2("C. 陪伴关怀类（情感价值）"),
            new Table({
                width: { size: 9026, type: WidthType.DXA },
                columnWidths: [3009, 2006, 4011],
                rows: [
                    new TableRow({ children: [createCell("服务项目", true, 3009), createCell("单价", true, 2006), createCell("说明", true, 4011)] }),
                    new TableRow({ children: [createCell("情感陪聊", false, 3009), createCell("30-50元/小时", false, 2006), createCell("心理疏导、陪伴聊天", false, 4011)] }),
                    new TableRow({ children: [createCell("文娱陪伴", false, 3009), createCell("40-60元/小时", false, 2006), createCell("下棋、读书、散步", false, 4011)] }),
                    new TableRow({ children: [createCell("节日关怀", false, 3009), createCell("按次收费", false, 2006), createCell("生日、节日特别服务", false, 4011)] })
                ]
            }),

            new Paragraph({ spacing: { before: 300, after: 100 }, children: [] }),
            createHeading2("D. 紧急响应类（增值服务）"),
            new Table({
                width: { size: 9026, type: WidthType.DXA },
                columnWidths: [3009, 2006, 4011],
                rows: [
                    new TableRow({ children: [createCell("服务项目", true, 3009), createCell("单价", true, 2006), createCell("说明", true, 4011)] }),
                    new TableRow({ children: [createCell("夜间陪护", false, 3009), createCell("200-400元/夜", false, 2006), createCell("24小时应急响应", false, 4011)] }),
                    new TableRow({ children: [createCell("紧急救援协助", false, 3009), createCell("含在套餐内", false, 2006), createCell("联动120、社区", false, 4011)] }),
                    new TableRow({ children: [createCell("定期巡查", false, 3009), createCell("100-200元/月", false, 2006), createCell("独居老人安全巡查", false, 4011)] })
                ]
            }),

            new Paragraph({ spacing: { before: 300, after: 100 }, children: [] }),
            createHeading2("3.2 服务套餐设计"),
            new Table({
                width: { size: 9026, type: WidthType.DXA },
                columnWidths: [2257, 2257, 3009, 1503],
                rows: [
                    new TableRow({ children: [createCell("套餐类型", true, 2257), createCell("月定价", true, 2257), createCell("服务内容", true, 3009), createCell("适合人群", true, 1503)] }),
                    new TableRow({ children: [createCell("基础套餐", false, 2257), createCell("399元/月", false, 2257), createCell("4次助洁+2次助餐+1次健康监测", false, 3009), createCell("自理老人", false, 1503)] }),
                    new TableRow({ children: [createCell("标准套餐", false, 2257), createCell("799元/月", false, 2257), createCell("8次服务（任选）+健康监测+陪聊", false, 3009), createCell("半自理老人", false, 1503)] }),
                    new TableRow({ children: [createCell("尊享套餐", false, 2257), createCell("1599元/月", false, 2257), createCell("16次服务+陪诊+康复+夜间响应", false, 3009), createCell("失能/失智老人", false, 1503)] }),
                    new TableRow({ children: [createCell("按次付费", false, 2257), createCell("—", false, 2257), createCell("随时预约，灵活选择", false, 3009), createCell("所有用户", false, 1503)] })
                ]
            }),

            // 第四部分
            createHeading1("四、商业模式设计"),
            createHeading2("4.1 收入结构"),
            new Table({
                width: { size: 9026, type: WidthType.DXA },
                columnWidths: [3009, 2006, 4011],
                rows: [
                    new TableRow({ children: [createCell("收入来源", true, 3009), createCell("占比目标", true, 2006), createCell("说明", true, 4011)] }),
                    new TableRow({ children: [createCell("服务佣金", false, 3009), createCell("60-70%", false, 2006), createCell("每单抽取15-25%佣金", false, 4011)] }),
                    new TableRow({ children: [createCell("套餐预售", false, 3009), createCell("20-25%", false, 2006), createCell("月度/季度/年度套餐", false, 4011)] }),
                    new TableRow({ children: [createCell("增值服务", false, 3009), createCell("5-10%", false, 2006), createCell("紧急响应、健康报告等", false, 4011)] }),
                    new TableRow({ children: [createCell("政府补贴", false, 3009), createCell("5-10%", false, 2006), createCell("接入政府采购后", false, 4011)] })
                ]
            }),

            new Paragraph({ spacing: { before: 300, after: 100 }, children: [] }),
            createHeading2("4.2 盈利测算"),
            new Table({
                width: { size: 9026, type: WidthType.DXA },
                columnWidths: [4513, 4513],
                rows: [
                    new TableRow({ children: [createCell("指标", true, 4513), createCell("数据", true, 4513)] }),
                    new TableRow({ children: [createCell("单用户月均消费", false, 4513), createCell("500-800元", false, 4513)] }),
                    new TableRow({ children: [createCell("单护理员月产出", false, 4513), createCell("6000-10000元（服务额）", false, 4513)] }),
                    new TableRow({ children: [createCell("平台佣金率", false, 4513), createCell("15-25%", false, 4513)] }),
                    new TableRow({ children: [createCell("单护理员月贡献利润", false, 4513), createCell("900-2500元", false, 4513)] }),
                    new TableRow({ children: [createCell("盈亏平衡", false, 4513), createCell("20-30名活跃护理员 + 100个稳定客户", false, 4513)] })
                ]
            }),

            // 第五部分
            createHeading1("五、低成本启动策略"),
            createHeading2("5.1 阶段一：最小MVP验证（预算1-3万）"),
            createParagraph("目标：跑通服务闭环，验证核心假设"),
            new Table({
                width: { size: 9026, type: WidthType.DXA },
                columnWidths: [3009, 2006, 4011],
                rows: [
                    new TableRow({ children: [createCell("任务", true, 3009), createCell("预算", true, 2006), createCell("说明", true, 4011)] }),
                    new TableRow({ children: [createCell("平台搭建", false, 3009), createCell("3000-8000元", false, 2006), createCell("微信小程序/h5页面（模板）", false, 4011)] }),
                    new TableRow({ children: [createCell("首批护理员招募", false, 3009), createCell("0-2000元", false, 2006), createCell("3-5人，社群里招募", false, 4011)] }),
                    new TableRow({ children: [createCell("种子用户获取", false, 3009), createCell("3000-8000元", false, 2006), createCell("1-2个小区地推", false, 4011)] }),
                    new TableRow({ children: [createCell("物料与工具", false, 3009), createCell("2000元", false, 2006), createCell("工牌、服装、服务包", false, 4011)] }),
                    new TableRow({ children: [createCell("备用金", false, 3009), createCell("5000元", false, 2006), createCell("处理纠纷、应急", false, 4011)] })
                ]
            }),
            new Paragraph({ spacing: { before: 200, after: 100 }, children: [] }),
            createParagraph("执行要点：", true),
            createBulletParagraph("先做\"服务中介\"：接单后派给护理员，不自养团队"),
            createBulletParagraph("聚焦1个小区、20-30户老人"),
            createBulletParagraph("选高频刚需的助洁、陪诊切入"),

            new Paragraph({ spacing: { before: 300, after: 100 }, children: [] }),
            createHeading2("5.2 阶段二：模式打磨（预算5-10万）"),
            createParagraph("目标：建立服务标准，形成口碑"),
            createBulletParagraph("开发/优化接单平台（小程序）"),
            createBulletParagraph("组建核心护理员团队（8-15人）"),
            createBulletParagraph("建立培训体系和服务标准SOP"),
            createBulletParagraph("拓展至3-5个社区"),
            createBulletParagraph("申请政府购买服务资质"),

            new Paragraph({ spacing: { before: 300, after: 100 }, children: [] }),
            createHeading2("5.3 阶段三：规模扩张（预算15-30万）"),
            createParagraph("目标：跨区域复制，品牌化运营"),
            createBulletParagraph("招募城市合伙人/加盟商"),
            createBulletParagraph("接入政府养老服务平台"),
            createBulletParagraph("开发APP和智能派单系统"),
            createBulletParagraph("建立护理员培训认证体系"),
            createBulletParagraph("接入商业保险、医疗资源"),

            // 第六部分
            createHeading1("六、获客渠道"),
            createHeading2("6.1 B端获客（快速起量）"),
            new Table({
                width: { size: 9026, type: WidthType.DXA },
                columnWidths: [3009, 4011, 2006],
                rows: [
                    new TableRow({ children: [createCell("渠道", true, 3009), createCell("策略", true, 4011), createCell("优先级", true, 2006)] }),
                    new TableRow({ children: [createCell("社区居委会", false, 3009), createCell("合作开展活动，获得官方背书", false, 4011), createCell("⭐⭐⭐", false, 2006)] }),
                    new TableRow({ children: [createCell("物业公司", false, 3009), createCell("嵌入物业服务体系", false, 4011), createCell("⭐⭐⭐", false, 2006)] }),
                    new TableRow({ children: [createCell("医院/社区卫生中心", false, 3009), createCell("导流术后康复、慢病老人", false, 4011), createCell("⭐⭐", false, 2006)] }),
                    new TableRow({ children: [createCell("养老机构", false, 3009), createCell("合作出院后衔接服务", false, 4011), createCell("⭐⭐", false, 2006)] })
                ]
            }),

            new Paragraph({ spacing: { before: 300, after: 100 }, children: [] }),
            createHeading2("6.2 C端获客（口碑裂变）"),
            new Table({
                width: { size: 9026, type: WidthType.DXA },
                columnWidths: [3009, 4011, 2006],
                rows: [
                    new TableRow({ children: [createCell("渠道", true, 3009), createCell("策略", true, 4011), createCell("优先级", true, 2006)] }),
                    new TableRow({ children: [createCell("口碑转介绍", false, 3009), createCell("老客户推荐奖励", false, 4011), createCell("⭐⭐⭐", false, 2006)] }),
                    new TableRow({ children: [createCell("子女微信群", false, 3009), createCell("精准触达决策者", false, 4011), createCell("⭐⭐⭐", false, 2006)] }),
                    new TableRow({ children: [createCell("社区地推", false, 3009), createCell("摆摊、体验活动", false, 4011), createCell("⭐⭐", false, 2006)] }),
                    new TableRow({ children: [createCell("抖音/视频号", false, 3009), createCell("护理员故事、暖心服务", false, 4011), createCell("⭐⭐", false, 2006)] }),
                    new TableRow({ children: [createCell("老年人活动中心", false, 3009), createCell("线下渗透", false, 4011), createCell("⭐⭐", false, 2006)] })
                ]
            }),

            new Paragraph({ spacing: { before: 300, after: 100 }, children: [] }),
            createHeading2("6.3 核心话术"),
            createParagraph("针对子女：", true),
            createQuoteParagraph("\"专业护理员上门服务，子女上班更安心。首次体验半价，点击预约。\""),
            new Paragraph({ spacing: { before: 100, after: 100 }, children: [] }),
            createParagraph("针对老人：", true),
            createQuoteParagraph("\"政府认证服务，价格透明。不满意随时换人。\""),
            new Paragraph({ spacing: { before: 100, after: 100 }, children: [] }),
            createParagraph("针对社区/物业：", true),
            createQuoteParagraph("\"帮助社区建立养老服务档案，提升居民满意度，配合政府考核。\""),

            // 第七部分
            createHeading1("七、团队组建"),
            createHeading2("7.1 初期团队（3-5人）"),
            new Table({
                width: { size: 9026, type: WidthType.DXA },
                columnWidths: [3009, 4011, 2006],
                rows: [
                    new TableRow({ children: [createCell("角色", true, 3009), createCell("职责", true, 4011), createCell("来源", true, 2006)] }),
                    new TableRow({ children: [createCell("项目负责人", false, 3009), createCell("整体运营、渠道对接", false, 4011), createCell("创始人兼任", false, 2006)] }),
                    new TableRow({ children: [createCell("护理员", false, 3009), createCell("核心服务交付", false, 4011), createCell("退休护士、4050人员", false, 2006)] }),
                    new TableRow({ children: [createCell("兼职调度", false, 3009), createCell("订单匹配、排班", false, 4011), createCell("兼职或自动化", false, 2006)] }),
                    new TableRow({ children: [createCell("推广人员", false, 3009), createCell("获客、活动执行", false, 4011), createCell("兼职或合伙人", false, 2006)] })
                ]
            }),

            new Paragraph({ spacing: { before: 300, after: 100 }, children: [] }),
            createHeading2("7.2 护理员招募标准"),
            createBulletParagraph("年龄：45-60岁优先（与老人有共鸣）"),
            createBulletParagraph("背景：无犯罪记录，有健康证"),
            createBulletParagraph("技能：持有护理证/育婴师证优先"),
            createBulletParagraph("性格：耐心、细心、有爱心"),
            createBulletParagraph("地域：社区周边居民优先"),

            new Paragraph({ spacing: { before: 300, after: 100 }, children: [] }),
            createHeading2("7.3 护理员激励"),
            new Table({
                width: { size: 9026, type: WidthType.DXA },
                columnWidths: [3009, 6017],
                rows: [
                    new TableRow({ children: [createCell("激励类型", true, 3009), createCell("内容", true, 6017)] }),
                    new TableRow({ children: [createCell("基础报酬", false, 3009), createCell("25-40元/小时（高于市场价）", false, 6017)] }),
                    new TableRow({ children: [createCell("订单奖励", false, 3009), createCell("好评奖励、续单奖励", false, 6017)] }),
                    new TableRow({ children: [createCell("晋升通道", false, 3009), createCell("星级护理员→督导→合伙人", false, 6017)] }),
                    new TableRow({ children: [createCell("归属感", false, 3009), createCell("团建、培训、节日福利", false, 6017)] })
                ]
            }),

            // 第八部分
            createHeading1("八、风险与应对"),
            new Table({
                width: { size: 9026, type: WidthType.DXA },
                columnWidths: [4513, 4513],
                rows: [
                    new TableRow({ children: [createCell("风险", true, 4513), createCell("应对策略", true, 4513)] }),
                    new TableRow({ children: [createCell("护理员流失", false, 4513), createCell("建立社群情感维系，多储备备份人员", false, 4513)] }),
                    new TableRow({ children: [createCell("安全事故", false, 4513), createCell("购买雇主责任险，服务前风险告知", false, 4513)] }),
                    new TableRow({ children: [createCell("老人投诉", false, 4513), createCell("24小时响应机制，快速换人", false, 4513)] }),
                    new TableRow({ children: [createCell("竞争模仿", false, 4513), createCell("深耕本地服务，建立口碑壁垒", false, 4513)] }),
                    new TableRow({ children: [createCell("政策变化", false, 4513), createCell("多元化获客，不依赖单一政府订单", false, 4513)] })
                ]
            }),

            // 第九部分
            createHeading1("九、下一步行动清单"),
            createCheckBoxParagraph("第一周：调研本地养老需求，走访2-3个社区"),
            createCheckBoxParagraph("第二周：完成小程序/接单平台搭建（模板）"),
            createCheckBoxParagraph("第三周：招募首批3-5名护理员，完成培训"),
            createCheckBoxParagraph("第四周：在1个小区开展免费体验活动"),
            createCheckBoxParagraph("第一个月：获取20个种子用户，收集反馈"),
            createCheckBoxParagraph("第二个月：优化服务流程，准备复制"),

            // 第十部分
            createHeading1("十、成功案例参考"),
            createHeading2("小柏家护（对标学习）"),
            createBulletParagraph("引进美国高端养老护理经验"),
            createBulletParagraph("国内率先实现护理员培训体系"),
            createBulletParagraph("从单一服务到多元化生态"),

            new Paragraph({ spacing: { before: 300, after: 100 }, children: [] }),
            createHeading2("社区嵌入式养老（上海模式）"),
            createBulletParagraph("15分钟养老服务圈"),
            createBulletParagraph("社区综合为老服务中心"),
            createBulletParagraph("可复制、可推广"),

            // 结束语
            new Paragraph({ spacing: { before: 500, after: 200 }, children: [] }),
            new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { before: 400, after: 400 },
                border: { top: { style: BorderStyle.SINGLE, size: 6, color: "CCCCCC", space: 8 } },
                children: [new TextRun({ text: "如需更详细的某一部分（如财务测算模型、服务SOP手册、护理员培训资料），可以继续深挖。", size: 20, color: "666666", italics: true })]
            })
        ]
    }]
});

Packer.toBuffer(doc).then(buffer => {
    fs.writeFileSync('/Users/likaiying/WorkBuddy/20260401111529/居家养老服务_落地方案.docx', buffer);
    console.log('Word文档已生成：居家养老服务_落地方案.docx');
});
