"use strict";
const { Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType,
        Table, TableRow, TableCell, WidthType, ShadingType, VerticalAlign,
        PageBreak, convertInchesToTwip, Header, Footer, PageNumber,
        NumberFormat, LevelFormat } = require("docx");
const fs = require("fs");

const PRIMARY   = "1F4E79";
const SECONDARY = "2E75B6";
const ACCENT    = "E67E22";
const LIGHT_BG  = "EBF5FB";
const WHITE     = "FFFFFF";
const DARK      = "2C3E50";
const GRAY      = "7F8C8D";

// ---- helpers ----
const sp  = (pts) => new Paragraph({ spacing: { before: pts * 10, after: 0 }, children: [] });
const p   = (text, o={}) => new Paragraph({
  alignment: o.align || AlignmentType.LEFT,
  spacing: { before:(o.sb||120)*10, after:(o.sa||120)*10 },
  children: [new TextRun({ text, font: o.font||"Arial",
    size: o.size||24, bold: !!o.bold, color: o.color||DARK,
    italics: !!o.italic })],
});
const h2  = (text) => new Paragraph({
  heading: HeadingLevel.HEADING_2,
  spacing: { before:400, after:200 },
  children: [new TextRun({ text, font:"Arial", size:28, bold:true, color:PRIMARY })],
});
const h3  = (text) => new Paragraph({
  heading: HeadingLevel.HEADING_3,
  spacing: { before:280, after:120 },
  children: [new TextRun({ text, font:"Arial", size:24, bold:true, color:SECONDARY })],
});
const pb  = () => new Paragraph({ children: [new PageBreak()] });

const borders = { top:{style:1,size:4,color:"BDC3C7"}, bottom:{style:1,size:4,color:"BDC3C7"},
                  left:{style:1,size:4,color:"BDC3C7"}, right:{style:1,size:4,color:"BDC3C7"} };

const hc = (t, w, bg) => new TableCell({
  borders, width:{size:w,type:WidthType.DXA},
  shading:{fill:bg||PRIMARY,type:ShadingType.CLEAR},
  margins:{top:100,bottom:100,left:120,right:120},verticalAlign:VerticalAlign.CENTER,
  children:[new Paragraph({alignment:AlignmentType.CENTER,
    children:[new TextRun({text:t,font:"Arial",size:20,bold:true,color:WHITE})]})]
});
const dc = (t, w, o={}) => new TableCell({
  borders, width:{size:w,type:WidthType.DXA},
  shading:{fill:o.bg||WHITE,type:ShadingType.CLEAR},
  margins:{top:80,bottom:80,left:120,right:120},verticalAlign:VerticalAlign.CENTER,
  children:[new Paragraph({alignment:o.al||AlignmentType.LEFT,
    children:[new TextRun({text:t,font:"Arial",size:20,color:o.co||DARK,bold:!!o.bd})]})]
});
const tr = (...cells) => new TableRow({children:cells});
const tbl = (rows, w) => new Table({width:{size:w||100,type:WidthType.PCT},rows});

const bullets = [
  "你不需要自己造硬件：选1-2款成熟的智能设备（摄像头/毫米波雷达/智能手环），你专注做软件平台+24小时值守服务",
  "从一开始就把服务做闭环：监测到异常 → 24小时值守响应 → 协调120/子女/邻居 → 事后跟进，不留断点",
  "用B端撬动C端：先去社区、物业、养老机构谈合作，用B端渠道铺开，再慢慢建立口碑拉C端用户",
  "硬件代销+服务订阅两条腿走路：硬件赚差价，服务订阅赚长期现金流",
  "从小城市切入：大厂都在一二线卷，三四线和农村几乎是空白，这就是你的窗口期",
  "把\"本地服务\"当成核心壁垒：AI可以抄，但你在每个城市有线下服务团队、有合作医护，这抄不走",
];

const risks = [
  {r:"用户付费意愿低", m:"先用免费体验打开市场，第1-3个月不收费，让用户感受到价值后再谈付费"},
  {r:"紧急响应能力不足", m:"第1年联合社区卫生服务中心、120急救中心，建立绿色通道"},
  {r:"硬件故障导致漏报", m:"选3家以上供应商，主备设备冗余，APP内每天自检提示"},
  {r:"隐私争议", m:"严格合规，设备安装需老人+子女双方知情同意，数据加密存储"},
  {r:"规模扩张慢", m:"前3个月专注1个小区/社区做透，拿到真实数据后再复制，不贪快"},
];

const ph1 = [
  ["项　目", 2200, PRIMARY, "内　容", 7600, LIGHT_BG],
  ["项目名称", 2200, "", "独居老人安全守护", 7600, ""],
  ["核心定位", 2200, "", "24小时值守响应服务闭环 + 智能硬件监测", 7600, ""],
  ["切入角度", 2200, "", "填补市场空白：目前所有竞品只\"看\"不\"管\"，我们做\"监测+响应\"闭环", 7600, ""],
  ["目标市场", 2200, "", "三四线城市 + 农村独居老人家庭", 7600, ""],
  ["差异化核心", 2200, "", "不卖硬件卖服务；不是工具，是保障", 7600, ""],
  ["最小验证", 2200, "", "1个社区 + 50-100户，运行3个月看数据", 7600, ""],
];

const ph2 = [
  ["产品层面", ""],
  ["硬件选型策略", "不自主研发硬件，选择成熟供应商：\n① 萤石/海康摄像头（200-800元/台）：实时视频+跌倒AI检测\n② 毫米波雷达（300-600元/台）：无感监测，适合卧室/卫生间（隐私敏感区）\n③ 智能手环（200-500元/个）：心率/SOS/定位，适合行动尚可的老人\n\n选品原则：WiFi直连+APP联动+支持API对接，拒绝蓝牙断联的坑"],
  ["软件平台", "自主开发核心APP（外包或SaaS化）：\n① 家属端APP：实时数据看板 + 异常推送 + 值班员联系方式\n② 老人端（大字版）：SOS一键呼叫 + 每日健康打卡 + 用药提醒\n③ 值班员工作台：异常事件列表 + 快速响应按钮 + 协同通知"],
  ["24小时值守中心", "核心壁垒，闭环关键：\n① 招聘2-3名全职值班员（三班倒），或外包给社区卫生服务中心\n② 异常事件触发 → 值班员5秒内接单 → 10秒内电话联系老人确认情况\n③ 确认紧急 → 立即拨打120 + 通知子女 + 必要时联系物业/邻居\n④ 非紧急 → 定时跟进，确保老人平安\n⑤ 夜间/节假日：启用AI语音助手辅助初步筛查，减少误报打扰"],
  ["增值服务包", "基础版（49元/月）：硬件监测 + 异常推送\n标准版（79元/月）：基础版 + 24小时人工值守响应\n尊享版（149元/月）：标准版 + 每月1次电话问安 + 每年2次上门探访"],
];

const ph3 = [
  ["启动成本估算（第一个社区）", ""],
  ["一次性投入", "① 硬件采购（50户×3件）：约4.5万元\n② APP开发/订阅（SaaS）：约3万元\n③ 值班员2人×3个月工资：约3.6万元\n④ 其他（SIM卡、辅料、培训）：约1万元\n合计：约12万元"],
  ["月度运营成本", "① 值班员工资（2人）：约1.2万元/月\n② 云服务/短信：约2000元/月\n③ 硬件售后/更换备用：约2000元/月\n④ 推广运营：约3000元/月\n合计：约1.9万元/月"],
  ["盈亏平衡测算", "50户 × 79元/月（标准版均价）= 3,950元/月\n距盈亏平衡（月成本1.9万）需要约5倍用户量（250户）\n前3个月：目标50-100户，用B端补贴C端，尽快跨过临界点"],
  ["规模化路径", "阶段一（第1-3个月）：1个社区，50-100户，验证服务流程\n阶段二（第4-6个月）：3-5个社区，200-500户，自建值班中心\n阶段三（第7-12个月）：10+个社区，1000+户，考虑开放加盟"],
];

const ph4 = [
  ["渠道合作策略", ""],
  ["B端合作（优先）", "① 社区居委会/物业：免费赠送1-2套设备体验，换取社区推广背书\n② 社区卫生服务中心：合作建立绿色通道，值班员确认为紧急后直连120\n③ 养老机构/日间照料中心：设备租赁+服务订阅，批量获客\n④ 房产中介/链家/贝壳找房：针对子女为父母购房/租房场景植入服务"],
  ["C端推广", "① 老龄委/民政局补贴合作：申请智慧养老补贴，降低用户付费门槛\n② 口碑传播：前100个用户免前3个月服务费，换取真实评价和转介绍\n③ 社区地推：与社区合作开展\"智慧养老体验日\"，现场演示+免费安装"],
  ["线上渠道", "① 抖音/视频号：拍真实案例短视频（已授权），子女买给父母的礼物\n② 微信社群：社区群、小区业主群，精准触达子女\n③ 适老化产品测评博主合作"],
];

const ph5 = [
  ["关键里程碑", ""],
  ["第1个月", "完成APP开发/SaaS接入，选定硬件供应商，安装10-20套设备试运行，验证技术稳定性"],
  ["第2-3个月", "扩展到50-100户，建立值班流程，跑通\"监测→响应→闭环\"全流程，拿到第一批真实用户数据"],
  ["第4-6个月", "扩展到200-500户，建立本地服务团队，开始谈B端渠道合作，验证B2B2C模式"],
  ["第7-12个月", "覆盖1000+户，考虑城市合伙人/加盟模式，开始有稳定的月度订阅收入"],
  ["12个月后", "有了数据背书，可以谈融资，或被大厂/养老机构收购，实现退出"],
];

const ph6 = [
  ["核心问题解答", ""],
  ["为什么大厂做不成？", "萤石/华为把养老当副业，没有服务闭环，只是卖硬件。真正的监护需要人在环里，他们做不到。"],
  ["为什么现在是好时机？", "① 老龄化加速，政策支持（智慧养老纳入十四五）\n② 子女在外打拼，独居老人多，监护需求刚性\n③ 硬件成本下降，AI算法成熟，基础设施已具备\n④ 竞品都在\"占位\"，没人真正做透"],
  ["最小化风险？", "先做一个小社区的验证，不要一开始就砸大钱建平台。用最小成本跑通流程，数据说话后再决定是否加码。"],
  ["未来想象空间？", "成为平台后，可以接入：慢病管理、在线问诊、药品配送、护理服务……围绕老年人健康管理做一个超级APP。"],
];

const ph7 = [
  ["竞品对比", ""],
  ["萤石/海康威视（上市公司）", "弱", "只做硬件+云存储，没有服务闭环，不碰线下"],
  ["华为/小米生态", "弱", "大厂副业，没有服务，做C端没有线下渠道"],
  ["颐养通/爱牵挂（SaaS平台）", "中", "专注B端G端，不做C端家庭用户"],
  ["亲鹿（创业公司）", "中", "最直接的竞争对手，产品体验好，但没有线下服务落地"],
  ["亲鹿（相比我们的优势）", "强", "我们有24小时人工值守响应，亲鹿只有APP推送"],
  ["传统养老院/保姆", "强", "费用高（3000+/月），老年人不愿离开家"],
];

const ph8 = [
  ["风险与对策", ""],
];
for (const {r, m} of risks) {
  ph8.push([r, m]);
}

const ph9 = [
  ["立即行动清单", ""],
  ["第1周", "① 选定1-2个目标社区/物业，上门谈合作\n② 注册公司+申请营业执照\n③ 确定硬件供应商，索要样品测试"],
  ["第2周", "① 选定APP开发方案（自建还是SaaS）\n② 设计LOGO和基本品牌\n③ 招聘或确定首批值班员人选"],
  ["第3-4周", "① 开发/配置APP，完成基本功能\n② 安装首批10-20套设备试运行\n③ 跑通全流程：监测→推送→值守→响应"],
  ["第1-3个月", "① 扩展到50-100户\n② 收集真实用户反馈，优化服务流程\n③ 打磨话术和服务标准SOP\n④ 申请老龄委/民政局智慧养老试点资格"],
];

// ---- BUILD ----
const content = [];

// 封面
content.push(
  sp(2400),
  p("独居老人安全监护", {align:AlignmentType.CENTER,sb:0,sa:200,size:72,bold:true,color:PRIMARY}),
  p("低成本落地方案", {align:AlignmentType.CENTER,sb:0,sa:200,size:72,bold:true,color:PRIMARY}),
  sp(400),
  p("（v2.0 基于竞品分析全面升级）", {align:AlignmentType.CENTER,sb:0,sa:100,size:32,color:GRAY}),
  sp(800),
  p("核心差异化：24小时值守响应服务闭环", {align:AlignmentType.CENTER,sb:0,sa:100,size:26,color:ACCENT}),
  sp(800),
  tbl([
    tr(hc("方案定位",2000,PRIMARY),hc("不卖硬件卖服务，做独居老人安全的\"守门人\"",7800,ACCENT)),
    tr(hc("切入角度",2000,PRIMARY),hc("填补竞品空白：目前所有竞品只\"看\"不\"管\"，我们做\"监测+响应\"闭环",7800,ACCENT)),
    tr(hc("目标市场",2000,PRIMARY),hc("三四线城市 + 农村独居老人家庭（竞品空白带）",7800,ACCENT)),
    tr(hc("启动资金",2000,PRIMARY),hc("约12万元（硬件4.5万 + 软件3万 + 人工3.6万 + 其他1万）",7800,ACCENT)),
    tr(hc("最小验证",2000,PRIMARY),hc("1个社区 + 50户，运行3个月，用真实数据验证商业模式",7800,ACCENT)),
  ], 100),
);

// 目录提示
content.push(
  sp(800),
  p("━━━━  目  录  ━━━━", {align:AlignmentType.CENTER,size:32,bold:true,color:SECONDARY}),
  sp(200),
  p("一、项目概述与核心差异化", {size:24,color:DARK}),
  p("二、产品体系与服务设计", {size:24,color:DARK}),
  p("三、启动成本与盈利模式", {size:24,color:DARK}),
  p("四、渠道策略", {size:24,color:DARK}),
  p("五、关键里程碑", {size:24,color:DARK}),
  p("六、竞品对比分析", {size:24,color:DARK}),
  p("七、风险与对策", {size:24,color:DARK}),
  p("八、立即行动清单", {size:24,color:DARK}),
);

content.push(pb());

// 一
content.push(h2("一、项目概述与核心差异化"));
content.push(h3("1.1 核心差异化：为什么你能成，竞品不能？"));
content.push(p("所有竞品都只解决了\"看\"的问题，没人解决\"管\"的问题。", {bold:true}));
content.push(p("他们的逻辑是：监测 → 推送子女 → 完事（子女在外地，看完焦虑也没办法）。"));
content.push(p("我们的逻辑是：监测 → 推送子女 + 24小时值守响应 → 协调120/子女/邻居 → 事后跟进，闭环。"));
content.push(p("这不是技术差异，是服务闭环差异。谁能让人真的安心睡好觉，谁就赢。", {color:ACCENT,bold:true}));
content.push(sp(200));

content.push(h3("1.2 方案总览"));
content.push(tbl([
  tr(hc("项　目",2200,PRIMARY), hc("内　容",7600,PRIMARY)),
  tr(dc("项目名称",2200,{bg:LIGHT_BG,bd:true}), dc("独居老人安全守护",7600)),
  tr(dc("核心定位",2200,{bg:LIGHT_BG,bd:true}), dc("24小时值守响应服务闭环 + 智能硬件监测",7600)),
  tr(dc("切入角度",2200,{bg:LIGHT_BG,bd:true}), dc("填补市场空白：目前所有竞品只\"看\"不\"管\"，我们做\"监测+响应\"闭环",7600)),
  tr(dc("目标市场",2200,{bg:LIGHT_BG,bd:true}), dc("三四线城市 + 农村独居老人家庭",7600)),
  tr(dc("差异化核心",2200,{bg:LIGHT_BG,bd:true}), dc("不卖硬件卖服务；不是工具，是保障",7600)),
  tr(dc("最小验证",2200,{bg:LIGHT_BG,bd:true}), dc("1个社区 + 50-100户，运行3个月看数据",7600)),
], 100));

// 二
content.push(pb());
content.push(h2("二、产品体系与服务设计"));

content.push(h3("2.1 硬件选型策略：借力成熟产品，不自己造轮子"));
content.push(p("选品原则：WiFi直连 + APP联动 + 支持API对接，拒绝蓝牙断联的坑", {color:GRAY,italic:true}));
content.push(p("① 萤石/海康摄像头（200-800元/台）", {bold:true}));
content.push(p("实时视频 + 跌倒AI检测，安装在客厅，推荐萤石C6CN系列（性价比高，APP成熟）"));
content.push(p("② 毫米波雷达（300-600元/台）", {bold:true}));
content.push(p("无感监测，适合卧室/卫生间（隐私敏感区，不装摄像头），跌倒检测用雷达更合适"));
content.push(p("③ 智能手环（200-500元/个）", {bold:true}));
content.push(p("心率/SOS/定位，适合行动尚可的老人，推荐华为/小米生态链产品"));

content.push(h3("2.2 软件平台：自主开发核心APP"));
content.push(p("① 家属端APP：实时数据看板 + 异常推送 + 值班员联系方式", {bold:true}));
content.push(p("② 老人端（大字版）：SOS一键呼叫 + 每日健康打卡 + 用药提醒", {bold:true}));
content.push(p("③ 值班员工作台：异常事件列表 + 快速响应按钮 + 协同通知", {bold:true}));
content.push(p("开发方案：前期用SaaS平台（如织信/简道云）快速搭建MVP，验证后再自建或外包", {color:GRAY,italic:true}));

content.push(h3("2.3 24小时值守中心：核心壁垒，闭环关键"));
content.push(p("招聘2-3名全职值班员（三班倒），或外包给社区卫生服务中心", {bold:true}));
content.push(p("异常事件触发 → 值班员5秒内接单 → 10秒内电话联系老人确认情况"));
content.push(p("确认紧急 → 立即拨打120 + 通知子女 + 必要时联系物业/邻居"));
content.push(p("非紧急 → 定时跟进，确保老人平安"));
content.push(p("夜间/节假日：启用AI语音助手辅助初步筛查，减少误报打扰"));

content.push(h3("2.4 服务定价套餐"));
content.push(tbl([
  tr(hc("套餐版本",2400,PRIMARY), hc("价格",1600,PRIMARY), hc("包含内容",5600,PRIMARY)),
  tr(dc("基础版",2400,{bg:LIGHT_BG,bd:true}), dc("49元/月",1600,{al:AlignmentType.CENTER}), dc("硬件监测 + 异常推送",5600)),
  tr(dc("标准版（推荐）",2400,{bg:LIGHT_BG,bd:true}), dc("79元/月",1600,{al:AlignmentType.CENTER}), dc("基础版 + 24小时人工值守响应 + 紧急协调",5600)),
  tr(dc("尊享版",2400,{bg:LIGHT_BG,bd:true}), dc("149元/月",1600,{al:AlignmentType.CENTER}), dc("标准版 + 每月1次电话问安 + 每年2次上门探访",5600)),
], 100));

// 三
content.push(pb());
content.push(h2("三、启动成本与盈利模式"));

content.push(h3("3.1 启动成本估算（第一个社区，50户）"));
content.push(tbl([
  tr(hc("费用类别",2200,PRIMARY), hc("明细",7600,PRIMARY)),
  tr(dc("硬件采购",2200,{bg:LIGHT_BG,bd:true}), dc("50户 × 3件设备（摄像头+雷达+手环），均价900元 = 4.5万元",7600)),
  tr(dc("软件/APP",2200,{bg:LIGHT_BG,bd:true}), dc("SaaS平台订阅或外包开发，约3万元",7600)),
  tr(dc("人工（首批）",2200,{bg:LIGHT_BG,bd:true}), dc("2名值班员 × 3个月 × 6000元/月 = 3.6万元",7600)),
  tr(dc("其他杂费",2200,{bg:LIGHT_BG,bd:true}), dc("SIM卡、辅料、培训、推广物料等，约1万元",7600)),
  tr(dc("合计启动成本",2200,{bg:ACCENT,bd:true}), dc("约12万元",7600,{co:ACCENT,bd:true})),
], 100));

content.push(h3("3.2 月度运营成本"));
content.push(p("值班员工资（2人）：约1.2万元/月"));
content.push(p("云服务/短信推送：约2000元/月"));
content.push(p("硬件售后/备用：约2000元/月"));
content.push(p("推广运营：约3000元/月"));
content.push(p("合计：约1.9万元/月", {bold:true,color:ACCENT}));

content.push(h3("3.3 盈亏平衡测算"));
content.push(p("50户 × 79元/月（标准版均价）= 3,950元/月"));
content.push(p("距盈亏平衡（月成本1.9万）需要约5倍用户量（250户）"));
content.push(p("前3个月：目标50-100户，用B端补贴C端，尽快跨过临界点", {color:GRAY,italic:true}));
content.push(p("规模化后：500户 × 79元 = 3.95万/月，覆盖成本后有稳定现金流", {bold:true}));

content.push(h3("3.4 收入模式"));
content.push(tbl([
  tr(hc("收入来源",3400,PRIMARY), hc("说明",6400,PRIMARY)),
  tr(dc("硬件差价",3400,{bg:LIGHT_BG,bd:true}), dc("代销萤石/雷达/手环，批量采购可拿到更低价差，约15-25%毛利",6400)),
  tr(dc("服务订阅（月费）",3400,{bg:LIGHT_BG,bd:true}), dc("主要现金流来源，79元/月，500户 = 3.95万/月",6400)),
  tr(dc("B端合作费",3400,{bg:LIGHT_BG,bd:true}), dc("养老机构/日间照料中心设备租赁+服务费",6400)),
  tr(dc("政府补贴",3400,{bg:LIGHT_BG,bd:true}), dc("申请智慧养老试点补贴，降低获客成本",6400)),
], 100));

// 四
content.push(pb());
content.push(h2("四、渠道策略"));

content.push(h3("4.1 B端合作（优先级最高）"));
content.push(p("① 社区居委会/物业：免费赠送1-2套设备体验，换取社区推广背书和业主群推广资源", {bold:true}));
content.push(p("② 社区卫生服务中心：合作建立绿色通道，值班员确认为紧急后直连120，减少响应时间", {bold:true}));
content.push(p("③ 养老机构/日间照料中心：设备租赁+服务订阅，批量获客，降低边际成本", {bold:true}));
content.push(p("④ 房产中介：针对子女为父母购房/租房场景植入服务，精准触达有支付能力的家庭", {bold:true}));

content.push(h3("4.2 C端推广"));
content.push(p("① 老龄委/民政局补贴合作：申请智慧养老补贴，降低用户付费门槛", {bold:true}));
content.push(p("② 口碑传播：前100个用户免前3个月服务费，换取真实评价和转介绍", {bold:true}));
content.push(p("③ 社区地推：与社区合作开展\"智慧养老体验日\"，现场演示+免费安装", {bold:true}));

content.push(h3("4.3 线上渠道"));
content.push(p("① 抖音/视频号：拍真实案例短视频（已授权），子女买给父母的礼物，情感营销"));
content.push(p("② 微信社群：社区群、小区业主群，精准触达子女"));
content.push(p("③ 适老化产品测评博主合作"));

// 五
content.push(pb());
content.push(h2("五、关键里程碑"));
content.push(tbl([
  tr(hc("阶段",1600,PRIMARY), hc("时间",1200,PRIMARY), hc("目标与行动",7000,PRIMARY)),
  tr(dc("验证期",1600,{bg:LIGHT_BG,bd:true}), dc("第1-3月",1200,{al:AlignmentType.CENTER}), dc("1个社区，50-100户，跑通\"监测→响应→闭环\"全流程，拿到第一批真实数据",7000)),
  tr(dc("扩张期",1600,{bg:LIGHT_BG,bd:true}), dc("第4-6月",1200,{al:AlignmentType.CENTER}), dc("3-5个社区，200-500户，建立本地服务团队，谈成2-3个B端渠道合作",7000)),
  tr(dc("成长期",1600,{bg:LIGHT_BG,bd:true}), dc("第7-12月",1200,{al:AlignmentType.CENTER}), dc("10+个社区，1000+户，考虑城市合伙人/加盟模式，月收入突破10万",7000)),
  tr(dc("规模化",1600,{bg:LIGHT_BG,bd:true}), dc("12月后",1200,{al:AlignmentType.CENTER}), dc("有了数据背书，可以谈融资，或被大厂/养老机构收购，实现退出",7000)),
], 100));

// 六
content.push(pb());
content.push(h2("六、竞品对比分析"));
content.push(p("基于对市面上10+家竞品的调研，以下是我们相比竞品的核心优势：", {color:GRAY,italic:true}));
content.push(tbl([
  tr(hc("竞品类型",2400,PRIMARY), hc("威胁等级",1400,PRIMARY), hc("我们的优势",6000,PRIMARY)),
  tr(dc("萤石/海康威视（上市公司）",2400,{bg:LIGHT_BG,bd:true}), dc("弱",1400,{al:AlignmentType.CENTER}), dc("只做硬件+云存储，没有服务闭环，不碰线下",6000)),
  tr(dc("华为/小米生态",2400,{bg:LIGHT_BG,bd:true}), dc("弱",1400,{al:AlignmentType.CENTER}), dc("大厂副业，没有服务，做C端没有线下渠道",6000)),
  tr(dc("颐养通/爱牵挂（SaaS平台）",2400,{bg:LIGHT_BG,bd:true}), dc("中",1400,{al:AlignmentType.CENTER}), dc("专注B端G端，不做C端家庭用户，我们做他们不做的事",6000)),
  tr(dc("亲鹿（创业公司）",2400,{bg:LIGHT_BG,bd:true}), dc("中",1400,{al:AlignmentType.CENTER}), dc("最直接的竞争对手，产品体验好，但没有线下服务落地，我们有24小时人工值守响应",6000)),
  tr(dc("传统养老院/保姆",2400,{bg:LIGHT_BG,bd:true}), dc("强（替代优势）",1400,{al:AlignmentType.CENTER}), dc("费用高（3000+/月），老年人不愿离开家；我们在家里就能享受专业监护服务",6000)),
], 100));
content.push(sp(200));
content.push(p("核心结论：没有一家做到\"监护+响应\"服务闭环，这就是我们的机会窗口。", {bold:true,color:ACCENT}));

// 七
content.push(pb());
content.push(h2("七、风险与对策"));
content.push(tbl([
  tr(hc("核心风险",2800,PRIMARY), hc("应对策略",7000,PRIMARY)),
  tr(dc("用户付费意愿低",2800,{bg:LIGHT_BG,bd:true}), dc("先用免费体验打开市场，第1-3个月不收费，让用户感受到价值后再谈付费",7000)),
  tr(dc("紧急响应能力不足",2800,{bg:LIGHT_BG,bd:true}), dc("第1年联合社区卫生服务中心、120急救中心，建立绿色通道",7000)),
  tr(dc("硬件故障导致漏报",2800,{bg:LIGHT_BG,bd:true}), dc("选3家以上供应商，主备设备冗余，APP内每天自检提示",7000)),
  tr(dc("隐私争议",2800,{bg:LIGHT_BG,bd:true}), dc("严格合规，设备安装需老人+子女双方知情同意，数据加密存储",7000)),
  tr(dc("规模扩张慢",2800,{bg:LIGHT_BG,bd:true}), dc("前3个月专注1个小区/社区做透，拿到真实数据后再复制，不贪快",7000)),
], 100));

// 八
content.push(pb());
content.push(h2("八、立即行动清单（30天）"));
content.push(tbl([
  tr(hc("时间",1600,PRIMARY), hc("行动项",8200,PRIMARY)),
  tr(dc("第1周",1600,{bg:LIGHT_BG,bd:true}), dc("① 选定1-2个目标社区/物业，上门谈合作\n② 注册公司+申请营业执照\n③ 确定硬件供应商，索要样品测试",8200)),
  tr(dc("第2周",1600,{bg:LIGHT_BG,bd:true}), dc("① 选定APP开发方案（自建还是SaaS）\n② 设计LOGO和基本品牌\n③ 招聘或确定首批值班员人选",8200)),
  tr(dc("第3-4周",1600,{bg:LIGHT_BG,bd:true}), dc("① 开发/配置APP，完成基本功能\n② 安装首批10-20套设备试运行\n③ 跑通全流程：监测→推送→值守→响应",8200)),
  tr(dc("第1-3个月",1600,{bg:LIGHT_BG,bd:true}), dc("① 扩展到50-100户\n② 收集真实用户反馈，优化服务流程\n③ 打磨话术和服务标准SOP\n④ 申请老龄委/民政局智慧养老试点资格",8200)),
], 100));

// 结尾
content.push(pb());
content.push(sp(800));
content.push(p("━━━━ 方案总结 ━━━━", {align:AlignmentType.CENTER,size:32,bold:true,color:PRIMARY}));
content.push(sp(400));
content.push(p("竞品都在\"占位\"，没有人真正做到服务闭环。", {align:AlignmentType.CENTER,size:28,bold:true,color:ACCENT}));
content.push(p("这就是你最小成本切入的窗口期。", {align:AlignmentType.CENTER,size:28,bold:true,color:ACCENT}));
content.push(sp(400));
content.push(p("硬件可以被复制，但24小时值守 + 本地化服务团队才是真正的护城河。", {align:AlignmentType.CENTER,size:24,color:GRAY}));
content.push(p("先做一个小社区验证，数据说话，再决定是否加码。", {align:AlignmentType.CENTER,size:24,color:GRAY}));
content.push(sp(600));
content.push(p("下一步：选定目标社区，注册公司，开始干。", {align:AlignmentType.CENTER,size:28,bold:true,color:PRIMARY}));

// ---- DOC ----
const doc = new Document({
  sections: [{ children: content }],
  styles: {
    default: { document: { run: { font: "Arial", size: 24 } } },
  },
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync("/Users/likaiying/WorkBuddy/20260401111529/独居老人安全监护_落地方案_v2.docx", buf);
  console.log("OK: 独居老人安全监护_落地方案_v2.docx");
}).catch(e => { console.error(e); process.exit(1); });
