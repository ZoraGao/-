import fs from "node:fs/promises";
import vm from "node:vm";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const source = await fs.readFile("index.html", "utf8");

function extractConst(name) {
  const marker = `const ${name} = `;
  const start = source.indexOf(marker);
  if (start < 0) throw new Error(`Missing ${name}`);
  const arrayOpen = source.indexOf("[", start + marker.length);
  const objectOpen = source.indexOf("{", start + marker.length);
  const open = arrayOpen >= 0 && (objectOpen < 0 || arrayOpen < objectOpen) ? arrayOpen : objectOpen;
  const closeChar = source[open] === "[" ? "]" : "}";
  let depth = 0;
  let quote = "";
  let escaped = false;
  for (let i = open; i < source.length; i += 1) {
    const char = source[i];
    if (quote) {
      if (escaped) escaped = false;
      else if (char === "\\") escaped = true;
      else if (char === quote) quote = "";
      continue;
    }
    if (char === "\"" || char === "'" || char === "`") {
      quote = char;
      continue;
    }
    if (char === source[open]) depth += 1;
    if (char === closeChar) depth -= 1;
    if (depth === 0) return source.slice(open, i + 1);
  }
  throw new Error(`Unclosed ${name}`);
}

function parseConst(name) {
  return vm.runInNewContext(`(${extractConst(name)})`);
}

const eventGroups = [
  ["生活事件", parseConst("lifeEvents")],
  ["周末事件", parseConst("weekendEvents")],
  ["工作事件", parseConst("dayWorkEvents")],
  ["茶水间事件", parseConst("teaEvents")],
  ["夜晚事件", parseConst("nightEvents")]
];
const monthlyFinanceEvents = parseConst("monthlyFinanceEvents");
const monthlyBills = parseConst("monthlyBills");
const outputDir = "outputs/mouse-sim-events";
const outputPath = `${outputDir}/鼠鼠模拟器事件数值表.xlsx`;

const penalties = [
  ["系统惩罚", 0, "上级评价过低：被辞退", "上级评价 < 18", "月末绩效沟通未通过，岗位不再继续保留。", "月末自动触发终局", 0, 0, 0, 0, 0, 0, "辞退结局"],
  ["系统惩罚", 1, "整洁过低：被房东赶出门", "整洁度 < 18 且月末概率触发", "房东临时看房发现房间太乱，不再续租，只能加钱换房。", "月末概率触发", -1, -2, -8, -2600, 0, 35, "低整洁度重大事件"],
  ["系统惩罚", 2, "发量过低：压力就医", "发量 < 28", "连续熬夜和通勤把头皮状态拖到报警，请半天假去医院。", "月末自动触发", -2, 12, -6, -650, 0, -3, "低发量惩罚"],
  ["系统惩罚", 3, "同事印象过低：被背锅", "同事印象 < 28", "协作印象差时，活动数据异常更容易被推锅。", "月末自动触发", -7, -2, -6, 0, -10, -2, "低同事印象惩罚"],
  ["系统惩罚", 4, "心态过低：效率崩盘", "心态 < 24", "打开文档都像加载失败，本月开局效率下降。", "月末自动触发", -4, -2, 4, 0, -2, -2, "低心态惩罚"],
  ["系统惩罚", 5, "现金不足：余额焦虑", "现金 < 下月房租+水电", "每天开始时现金低于 3320，会因为下月房租水电焦虑。", "每日自动触发", 0, 0, -1, 0, 0, 0, "每日心态扣减"],
  ["系统惩罚", 6, "现金为负：破产结局", "现金 < 0", "现金不能小于 0，否则直接进入破产结局：灰溜溜滚回老家。", "即时终局", 0, 0, 0, 0, 0, 0, "破产结局"],
  ["系统惩罚", 7, "整洁过低：出租屋反噬", "整洁度 < 20 且未触发赶人", "房间乱到找不到工牌，外卖袋和衣服堆在一起。", "月末自动触发", -2, -3, -4, 0, 0, -5, "低整洁度惩罚"],
  ["系统规则", 8, "周末不触发夜晚事件", "周末事件结束后", "周末结算选项后直接进入下一天；如果是每月第 28 天，则直接进入月末结算。", "自动流程", 0, 0, 0, 0, 0, 0, "流程规则"],
  ["系统规则", 9, "每日饭钱", "每天开始时", "每天白天开始时扣 25 元饭钱；读档不会重复扣同一天。", "每日自动触发", 0, 0, 0, -25, 0, 0, "每日现金扣减"]
];

const endings = [
  ["现实结果", 1, "现实：无法转正", "上级评价 < 50", "试用期答辩后被判定交付质量和节奏不达标，没能转正。", "通关判定", 0, 0, 0, 0, 0, 0, "现实结果三选一"],
  ["现实结果", 2, "现实：普通转正", "50 <= 上级评价 < 85", "撑完试用期并普通转正。若表现低于 70，会同时获得组解散裁员风险。", "通关判定", 0, 0, 0, 0, 0, 0, "现实结果三选一"],
  ["现实结果", 3, "现实：优秀转正", "上级评价 >= 85", "顺利转正且表现优秀。", "通关判定", 0, 0, 0, 0, 0, 0, "现实结果三选一"],
  ["未来成就", 1, "未来：心态长期告急", "心态 <= 15", "两年后对消息提示音越来越敏感，被迫离开岗位长期休养。", "可叠加成就", 0, 0, 0, 0, 0, 0, "可与现实转正同时出现"],
  ["未来成就", 2, "未来：头顶发亮", "发量 < 25", "两年后照镜子发现头顶发亮。", "可叠加成就", 0, 0, 0, 0, 0, 0, "可与现实转正同时出现"],
  ["未来成就", 3, "未来：背锅体质", "同事印象 < 30", "两年后活动数据异常被追责，因为同事评价太低而背锅。", "可叠加成就", 0, 0, 0, 0, 0, 0, "可叠加成就"],
  ["未来成就", 4, "未来：首付希望", "现金 >= 28000", "两年后靠存款和奖金攒出首付款。", "可叠加成就", 0, 0, 0, 0, 0, 0, "可叠加成就"],
  ["未来成就", 5, "未来：组解散裁员风险", "50 <= 上级评价 < 70", "表现不好不坏，业务线变化时容易被当成可替换的人。", "可叠加成就", 0, 0, 0, 0, 0, 0, "普通转正的风险分支"],
  ["未来成就", 6, "未来：余额焦虑", "现金 < 3320", "现金离下个月房租水电太近。", "可叠加成就", 0, 0, 0, 0, 0, 0, "可叠加成就"],
  ["未来成就", 7, "未来：出租屋反噬", "整洁度 < 25", "出租屋状态反过来影响生活效率。", "可叠加成就", 0, 0, 0, 0, 0, 0, "可叠加成就"],
  ["未来成就", 8, "未来：精神续航", "心态 >= 70", "保留喘气空间。", "可叠加成就", 0, 0, 0, 0, 0, 0, "正向成就"],
  ["未来成就", 9, "未来：同事缓冲垫", "同事印象 >= 80", "协作关系良好，出事时有人愿意听解释。", "可叠加成就", 0, 0, 0, 0, 0, 0, "正向成就"],
  ["未来成就", 10, "未来：发量守门成功", "发量 >= 80", "没有把每一次熬夜都写在头顶上。", "可叠加成就", 0, 0, 0, 0, 0, 0, "正向成就"]
];

const rows = [[
  "事件类型",
  "序号",
  "标签/标题",
  "触发条件",
  "事件正文",
  "选项",
  "上级评价变化",
  "发量变化",
  "心态变化",
  "现金变化",
  "同事印象变化",
  "整洁度变化",
  "备注"
]];
const mergeSpecs = [];

function nightChoiceType(choiceIndex) {
  if (choiceIndex === 0) return "housework";
  if (choiceIndex === 1) return "overtime";
  if (choiceIndex === 2) return "outing";
  return "other";
}

function choiceCleanDelta(choiceIndex) {
  const type = nightChoiceType(choiceIndex);
  if (type === "housework") return 8;
  if (type === "overtime") return -8;
  if (type === "outing") return -8;
  return 0;
}

for (const [type, events] of eventGroups) {
  events.forEach((event, eventIndex) => {
    const label = event.tag || event.title || "";
    let condition = event.commute ? `通勤=${event.commute}` : "";
    if (type === "茶水间事件") condition = "工作日约 18% 概率触发，事件池用完后不再重置";
    const body = event.body || "";
    const startRow = rows.length + 1;
    event.choices.forEach((choice, choiceIndex) => {
      const [text, workDelta = 0, moodDelta = 0, cashDelta = 0, peerDelta = 0, healthDelta = 0, explicitCleanDelta] = choice;
      const nightType = type === "夜晚事件" ? nightChoiceType(choiceIndex) : "";
      const cleanDelta = typeof explicitCleanDelta === "number" ? explicitCleanDelta : type === "夜晚事件" ? choiceCleanDelta(choiceIndex) : 0;
      let note = "";
      if (type === "夜晚事件" && nightType === "housework") note = "家务类选项；回南天或整洁度过低时只保留此类选项；整洁度过低时有几率触发抓老鼠";
      if (type === "夜晚事件" && nightType === "overtime") note = "加班类选项";
      if (type === "夜晚事件" && nightType === "outing") note = "外出类选项；台风天、发量过低或整洁度过低时会隐藏";
      if (type === "工作事件" && eventIndex === 0) note = "只在 9月1日触发";
      rows.push([
        type,
        eventIndex + 1,
        label,
        condition,
        body,
        text,
        workDelta,
        healthDelta,
        moodDelta,
        cashDelta,
        peerDelta,
        cleanDelta,
        note || `选项${choiceIndex + 1}`
      ]);
    });
    const endRow = rows.length;
    if (endRow > startRow) mergeSpecs.push({ startRow, endRow });
  });
}

const salaryRules = [
  ["上级评价 >= 85", 12500],
  ["70 <= 上级评价 < 85", 10500],
  ["50 <= 上级评价 < 70", 8500],
  ["30 <= 上级评价 < 50", 6500],
  ["18 <= 上级评价 < 30", 4200],
  ["上级评价 < 18", 0]
];

salaryRules.forEach(([condition, value], index) => {
  rows.push(["月末固定结算", index + 1, "月末工资", condition, "月末工资和上级评价挂钩。上级评价太低时会被辞退。", "月末工资", 0, 0, 0, value, 0, 0, "工资分档规则"]);
});

Object.entries(monthlyBills).filter(([key, value]) => !["salary", "dailyMeal"].includes(key) && value !== 0).forEach(([key, value], index) => {
  const labels = {
    rent: "房租",
    utilities: "水电网",
    commuteMeal: "通勤餐饮"
  };
  rows.push(["月末固定结算", salaryRules.length + index + 1, labels[key] || key, "每月第 28 天夜晚结束后", "月末发工资并结算固定生活成本。", labels[key] || key, 0, 0, 0, value, 0, 0, "固定月度现金流"]);
});

monthlyFinanceEvents.forEach((event, index) => {
  const [workDelta = 0, moodDelta = 0, cashDelta = 0, peerDelta = 0, healthDelta = 0, cleanDelta = 0] = event.delta;
  rows.push([
    "月末随机收支",
    index + 1,
    event.tag,
    "每月月末随机抽取，尽量不重复",
    event.body,
    event.tag,
    workDelta,
    healthDelta,
    moodDelta,
    cashDelta,
    peerDelta,
    cleanDelta,
    "随机收入/支出事件"
  ]);
});

penalties.forEach((row) => rows.push(row));
endings.forEach((row) => rows.push(row));

const workbook = Workbook.create();
const sheet = workbook.worksheets.add("事件数值表");
sheet.getRangeByIndexes(0, 0, rows.length, rows[0].length).values = rows;
sheet.freezePanes.freezeRows(1);
sheet.showGridLines = false;
sheet.getRange("A1:M1").format = {
  fill: "#1F2937",
  font: { bold: true, color: "#FFFFFF", name: "Arial", size: 10 },
  verticalAlignment: "center"
};
sheet.getRangeByIndexes(1, 0, rows.length - 1, rows[0].length).format = {
  font: { name: "Arial", size: 10 },
  verticalAlignment: "top"
};
sheet.getRangeByIndexes(0, 0, rows.length, rows[0].length).format.borders = {
  insideHorizontal: { style: "thin", color: "#E5E7EB" },
  top: { style: "thin", color: "#D1D5DB" },
  bottom: { style: "thin", color: "#D1D5DB" }
};
sheet.getRange("A:A").format.columnWidth = 14;
sheet.getRange("B:B").format.columnWidth = 8;
sheet.getRange("C:C").format.columnWidth = 20;
sheet.getRange("D:D").format.columnWidth = 22;
sheet.getRange("E:E").format.columnWidth = 70;
sheet.getRange("F:F").format.columnWidth = 34;
sheet.getRange("G:L").format.columnWidth = 12;
sheet.getRange("M:M").format.columnWidth = 30;
for (let row = 2; row <= rows.length; row += 1) {
  if (row % 2 === 0) sheet.getRange(`A${row}:M${row}`).format.fill = "#EAF7FB";
}
for (const { startRow, endRow } of mergeSpecs) {
  ["A", "B", "C", "D", "E"].forEach((col) => {
    sheet.mergeCells(`${col}${startRow}:${col}${endRow}`);
  });
  sheet.getRange(`A${startRow}:E${endRow}`).format.verticalAlignment = "center";
}
sheet.getRange("E:E").format.wrapText = true;
sheet.getRange("F:F").format.wrapText = true;
sheet.getRange("M:M").format.wrapText = true;
sheet.getRangeByIndexes(0, 0, rows.length, rows[0].length).format.autofitRows();

await workbook.recalculate();
await fs.mkdir(outputDir, { recursive: true });
const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(outputPath);
console.log(outputPath);
