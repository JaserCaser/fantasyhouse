const pptxgen = require("pptxgenjs");

let pres = new pptxgen();
pres.layout = 'LAYOUT_16x9';
pres.title = '文件知识库产品介绍';
pres.author = 'Gemini CLI';

// Theme Colors
const COLORS = {
    primary: "028090",
    secondary: "00A896",
    accent: "02C39A",
    text: "1E293B",
    muted: "64748B",
    bg: "F8FAFC",
    white: "FFFFFF"
};

// --- Slide 1: Title Slide ---
let slide1 = pres.addSlide();
slide1.background = { color: COLORS.primary };

slide1.addText("文件知识库", {
    x: 1, y: 1.5, w: 8, h: 1,
    fontSize: 54, bold: true, color: COLORS.white,
    align: "center", fontFace: "Microsoft YaHei"
});

slide1.addText("智能、高效、安全的团队协作与知识管理平台", {
    x: 1, y: 2.6, w: 8, h: 0.5,
    fontSize: 24, color: COLORS.white,
    align: "center", fontFace: "Microsoft YaHei",
    opacity: 0.9
});

slide1.addShape(pres.shapes.RECTANGLE, {
    x: 4.5, y: 3.5, w: 1, h: 0.05,
    fill: { color: COLORS.accent }
});

slide1.addText("2026年5月", {
    x: 1, y: 4.5, w: 8, h: 0.4,
    fontSize: 14, color: COLORS.white,
    align: "center", fontFace: "Arial",
    opacity: 0.7
});

// --- Slide 2: Project Background ---
let slide2 = pres.addSlide();
slide2.addText("面临的挑战", { x: 0.5, y: 0.3, w: 9, h: 0.6, fontSize: 32, bold: true, color: COLORS.primary, fontFace: "Microsoft YaHei" });
slide2.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1, w: 1.5, h: 0.04, fill: { color: COLORS.accent } });

const challenges = [
    { title: "信息碎片化", desc: "文档分散在不同设备和目录中，难以统一管理" },
    { title: "搜索效率低", desc: "无法快速定位文件内容，全文本搜索缺失" },
    { title: "知识沉淀难", desc: "文件间缺乏关联，难以形成系统的知识体系" },
    { title: "协作成本高", desc: "版本冲突、权限混乱，影响团队产出" }
];

challenges.forEach((item, idx) => {
    let x = 0.5 + (idx % 2) * 4.5;
    let y = 1.5 + Math.floor(idx / 2) * 1.8;
    
    slide2.addShape(pres.shapes.RECTANGLE, {
        x: x, y: y, w: 4, h: 1.5,
        fill: { color: COLORS.white },
        line: { color: COLORS.bg, width: 1 },
        shadow: { type: "outer", color: "000000", blur: 10, offset: 2, opacity: 0.05 }
    });
    
    slide2.addText(item.title, {
        x: x + 0.3, y: y + 0.3, w: 3.4, h: 0.4,
        fontSize: 20, bold: true, color: COLORS.primary, fontFace: "Microsoft YaHei"
    });
    
    slide2.addText(item.desc, {
        x: x + 0.3, y: y + 0.8, w: 3.4, h: 0.5,
        fontSize: 14, color: COLORS.muted, fontFace: "Microsoft YaHei"
    });
});

// --- Slide 3: Core Technology ---
let slide3 = pres.addSlide();
slide3.addText("技术架构与优势", { x: 0.5, y: 0.3, w: 9, h: 0.6, fontSize: 32, bold: true, color: COLORS.primary, fontFace: "Microsoft YaHei" });
slide3.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1, w: 1.5, h: 0.04, fill: { color: COLORS.accent } });

const tech = [
    { label: "后端核心", val: "FastAPI + Python", icon: "⚙️" },
    { label: "前端引擎", val: "Vue.js + Apple Style UI", icon: "💻" },
    { label: "内容检索", val: "SQLite FTS5 + CJK Support", icon: "🔍" },
    { label: "协同办公", val: "OnlyOffice Integration", icon: "🤝" }
];

tech.forEach((item, idx) => {
    slide3.addText(`${item.icon} ${item.label}: ${item.val}`, {
        x: 1, y: 1.8 + idx * 0.8, w: 8, h: 0.5,
        fontSize: 18, color: COLORS.text, fontFace: "Microsoft YaHei",
        bullet: true
    });
});

slide3.addShape(pres.shapes.RECTANGLE, {
    x: 6, y: 1.8, w: 3, h: 3,
    fill: { color: COLORS.primary, transparency: 90 },
    line: { color: COLORS.primary, width: 1 }
});
slide3.addText("高性能存储与毫秒级搜索响应", {
    x: 6.2, y: 2.8, w: 2.6, h: 1,
    fontSize: 16, bold: true, color: COLORS.primary, align: "center", fontFace: "Microsoft YaHei"
});

// --- Slide 4: Feature: Document Lifecycle ---
let slide4 = pres.addSlide();
slide4.addText("全生命周期文件管理", { x: 0.5, y: 0.3, w: 9, h: 0.6, fontSize: 32, bold: true, color: COLORS.primary, fontFace: "Microsoft YaHei" });

slide4.addTable([
    ["阶段", "核心功能", "应用场景"],
    ["上传阶段", "多格式支持、自动分析、配额控制", "团队资料入库、个人笔记上传"],
    ["处理阶段", "全文本索引、Top关键词、字符统计", "快速理解长文档核心内容"],
    ["编辑阶段", "在线协同编辑、多版本管理、草稿箱", "多人协作、回溯历史版本"],
    ["消费阶段", "多维度搜索、分类目录、收藏夹", "日常查阅、知识库快速检索"]
], {
    x: 0.5, y: 1.5, w: 9, h: 3.5,
    fontSize: 14, fontFace: "Microsoft YaHei",
    border: { pt: 1, color: "E2E8F0" },
    fill: { color: COLORS.white },
    colW: [1.5, 4, 3.5],
    valign: "middle"
});

// --- Slide 5: Feature: Intelligence & AI ---
let slide5 = pres.addSlide();
slide5.addText("智能驱动：AI 与 深度检索", { x: 0.5, y: 0.3, w: 9, h: 0.6, fontSize: 32, bold: true, color: COLORS.primary, fontFace: "Microsoft YaHei" });

slide5.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 1.2, w: 4.2, h: 3.5,
    fill: { color: COLORS.bg }, line: { color: COLORS.accent, width: 1 }
});
slide5.addText("🤖 智能问答 (QA)", { x: 0.8, y: 1.4, w: 3.6, h: 0.4, fontSize: 20, bold: true, color: COLORS.primary, fontFace: "Microsoft YaHei" });
slide5.addText("基于大语言模型（LLM）的知识问答系统，能够直接从上传的文档中提取答案，实现真正的“对话式搜索”。", {
    x: 0.8, y: 2.0, w: 3.6, h: 2, fontSize: 14, color: COLORS.text, fontFace: "Microsoft YaHei"
});

slide5.addShape(pres.shapes.RECTANGLE, {
    x: 5.3, y: 1.2, w: 4.2, h: 3.5,
    fill: { color: COLORS.bg }, line: { color: COLORS.accent, width: 1 }
});
slide5.addText("⚡ 深度搜索 (FTS)", { x: 5.6, y: 1.4, w: 3.6, h: 0.4, fontSize: 20, bold: true, color: COLORS.primary, fontFace: "Microsoft YaHei" });
slide5.addText("秒级完成百万字级别的全文本检索，支持中文分词优化，高亮显示命中片段，让文档搜索不再局限于文件名。", {
    x: 5.6, y: 2.0, w: 3.6, h: 2, fontSize: 14, color: COLORS.text, fontFace: "Microsoft YaHei"
});

// --- Slide 6: Visual Knowledge ---
let slide6 = pres.addSlide();
slide6.addText("可视化与交互体验", { x: 0.5, y: 0.3, w: 9, h: 0.6, fontSize: 32, bold: true, color: COLORS.primary, fontFace: "Microsoft YaHei" });

slide6.addText("让知识流动起来", { x: 0.5, y: 0.9, w: 9, h: 0.4, fontSize: 16, color: COLORS.muted, fontFace: "Microsoft YaHei" });

const visuals = [
    { title: "知识图谱", desc: "可视化展示文件间的关联关系，发现潜在的知识连接。" },
    { title: "扇形文件轮播", desc: "极具视觉冲击力的文件展示方式，提升交互趣味性。" },
    { title: "Apple Style UI", desc: "极致简约的界面设计，支持深色/浅色模式随心切换。" }
];

visuals.forEach((item, idx) => {
    slide6.addShape(pres.shapes.RECTANGLE, {
        x: 0.5 + idx * 3.2, y: 2.0, w: 2.8, h: 2.5,
        fill: { color: COLORS.primary, transparency: 95 },
        line: { color: COLORS.primary, width: 1 }
    });
    slide6.addText(item.title, {
        x: 0.6 + idx * 3.2, y: 2.2, w: 2.6, h: 0.4,
        fontSize: 18, bold: true, color: COLORS.primary, align: "center", fontFace: "Microsoft YaHei"
    });
    slide6.addText(item.desc, {
        x: 0.6 + idx * 3.2, y: 2.8, w: 2.6, h: 1.5,
        fontSize: 13, color: COLORS.text, align: "center", fontFace: "Microsoft YaHei"
    });
});

// --- Slide 7: Management & Governance ---
let slide7 = pres.addSlide();
slide7.addText("系统治理与安全性", { x: 0.5, y: 0.3, w: 9, h: 0.6, fontSize: 32, bold: true, color: COLORS.primary, fontFace: "Microsoft YaHei" });

const governance = [
    "基于工作空间（Workspace）的资源隔离与配额限制",
    "详尽的管理员日志，全流程追踪文件操作历史",
    "灵活的权限控制系统：查看、编辑、删除、转移",
    "报表中心：SQL 驱动的自定义数据洞察与分析"
];

governance.forEach((text, idx) => {
    slide7.addText(text, {
        x: 1, y: 1.5 + idx * 0.8, w: 8, h: 0.5,
        fontSize: 18, color: COLORS.text, fontFace: "Microsoft YaHei",
        bullet: true
    });
});

// --- Slide 8: End Slide ---
let slide8 = pres.addSlide();
slide8.background = { color: COLORS.primary };

slide8.addText("谢谢观看", {
    x: 1, y: 2, w: 8, h: 1,
    fontSize: 48, bold: true, color: COLORS.white,
    align: "center", fontFace: "Microsoft YaHei"
});

slide8.addText("让每个文档都成为智慧的阶梯", {
    x: 1, y: 3, w: 8, h: 0.5,
    fontSize: 20, color: COLORS.white,
    align: "center", fontFace: "Microsoft YaHei",
    opacity: 0.8
});

pres.writeFile({ fileName: "File_Knowledge_Base_Product.pptx" }).then(fileName => {
    console.log(`PPT generated: ${fileName}`);
});
