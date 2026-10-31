/**
 * 全站唯一文案来源。电脑版与手机版都从这里渲染，禁止在 desktop/、mobile/ 里硬编码文案。
 * version 与 changelog 由发版流程在构建时注入（inject_site.py 依赖 id 为 changelog 的 screen），
 * 这里的值仅用于本地开发。
 */
window.SITE_CONTENT = {
  meta: {
    title: "Study Buddy · 帮你每门课都及格的AI学习搭子",
    description: "把课件整理成讲义，规划每周，按讲义带学，考前速通，作业和论文直接出成品。",
  },
  repo: "Liunian06/study-buddy",
  version: "1.0.0",
  nav: [
    { id: "home", label: "首页", hash: "#/" },
    { id: "guide", label: "使用指南", hash: "#/guide" },
    { id: "changelog", label: "更新日志", hash: "#/changelog" },
  ],
  install: {
    command: "npx skills add Liunian06/study-buddy --skill study-buddy -g",
    zipUrl: "https://github.com/Liunian06/study-buddy/releases/latest",
    copyLabel: "复制",
    copiedLabel: "已复制",
  },
  pain: {
    inboxTitle: "班级群与邮箱",
    unread: "99+条未读",
    mails: [
      { from: "微观经济学课程群", subject: "【课件更新】第1-9章课件已上传（共312页）", tag: "课件", kind: "warn", time: "09:12" },
      { from: "管理学原理", subject: "【通知】小组pre下周三课上展示，请各组确认分工", tag: "pre", kind: "warn", time: "08:47" },
      { from: "市场营销", subject: "【作业】星巴克case study，2000字，周五23:59截止", tag: "作业", kind: "warn", time: "昨天" },
      { from: "会计学", subject: "期末考试时间已公布：第16周周三 14:00", tag: "考试", kind: "bad", time: "昨天" },
      { from: "统计学", subject: "往年题已上传，供同学们参考", tag: "往年题", kind: "info", time: "周二" },
      { from: "国际贸易", subject: "课程论文选题请于本周日前提交，3000字", tag: "论文", kind: "warn", time: "周二" },
      { from: "教务处", subject: "本学期期末考试安排及注意事项", tag: "教务", kind: "info", time: "周一" },
      { from: "宏观经济学", subject: "【提醒】课堂小测下周一进行，范围第3-5章", tag: "小测", kind: "bad", time: "周一" },
    ],
    narration: "五门课的课件、作业和考试挤在一起，每一门都看不懂，还都得过。",
  },
  agent: {
    appName: "AI 助手",
    prompt: "老师发了一堆课件，帮我整理成讲义",
    thinking: "正在处理",
    newChat: "新对话",
    recents: ["微观经济学课件整理", "管理学pre讲稿", "后天考试速通计划", "星巴克case study"],
    model: "Claude",
    skill: "study-buddy",
    tools: [
      { verb: "复制", target: "原始资料/ → 课程资料/课件_微观经济学.pptx" },
      { verb: "转写", target: "课件文字与图片" },
      { verb: "写入", target: "讲义/第01章_供需与弹性.md" },
      { verb: "写入", target: "讲义目录.md" },
    ],
    reply: "已整理完成，目录结构如下：",
    placeholder: "回复…",
    tree: [
      { name: "study-buddy-data/0101_微观经济学/", dir: true },
      { name: "原始资料/ 和 课程资料/" },
      { name: "讲义目录.md" },
      { name: "错题本.md" },
      { name: "讲义/第01章_供需与弹性.md" },
    ],
    narration1: "一句话，学习搭子就能把课件整理成讲义。",
    narration2: "转写课件和图片，按知识点重排，标出考点……",
  },
  generate: {
    srcTitle: "老师的课件（第3章，共48页）",
    srcLines: [
      "第3章 消费者选择理论",
      "无差异曲线与预算约束",
      "**边际效用递减** 规律",
      "**消费者均衡** 条件：MRS = Px/Py",
    ],
    outTitle: "讲义/第03章_消费者选择.md",
    outLines: [
      "考点权重：高 · 往年每年必考",
      "**边际效用递减**：多吃一口，满足感越来越少",
      "**消费者均衡**：花光预算时，每块钱买到的满足感相同",
      "小测：下面哪种情形已达到均衡？",
    ],
    previews: [
      { file: "学期规划.md", body: ["## 及格路径", "- 微观：期末占60%，考到40分即可", "- 本周：整理第1-3章，做完小测"] },
      { file: "错题本.md", body: ["## 错题本", "- 混淆了替代效应和收入效应", "- 复习：第3天、第7天"] },
      { file: "2026-12-28_速通计划.md", body: ["## 速通计划（剩余2天）", "1. 先过往年题高频考点", "2. 做一套模拟卷", "3. 生成考前速记清单"] },
    ],
    narration: "讲义、计划、错题本一次到位，只做拿分必需的部分。",
    words: ["资料整理", "学期规划", "跟课学习", "期末速通"],
  },
  cram: {
    reply: "距离考试还有2天，我先按往年题把考点分了优先级。这两天的安排：",
    steps: ["今天上午：过一遍高频考点，每个只记结论和一个例子", "今天下午：做一套模拟卷，错题记进错题本", "明天晚上：生成考前速记清单，睡前过一遍"],
    toast: "2026-12-28_速通计划.md",
    narration: "pre、课程论文、case study也一样，直接出完整成品，转成docx或pdf提交。",
  },
  brand: {
    name: "Study Buddy",
    tagline: "帮你每门课都及格的AI学习搭子",
    narration: "把时间留给真正重要的事，及格就好。",
    cta: "查看使用指南",
  },
  guide: {
    title: "使用指南",
    intro: "同时修好几门课，只想每门都及格：整理课件、规划每周、带学、速通、做作业。",
    sections: [
      { h: "适合谁用", p: ["同时修多门晦涩课程、只想用最少时间及格的大学生", "老师发了大量课件，不知道从哪看起", "要做pre、课程论文、case study等作业"] },
      { h: "安装", p: ["需要 Windows 和已安装的 PowerPoint、Word，Node.js 22.20 及以上，Python 3.10 及以上，以及支持 Skills 的 AI 编程助手（如 Claude Code）。", "推荐直接把仓库 README 里的安装提示词发给你的助手，它会自动安装并检查环境。"], code: "npx skills add Liunian06/study-buddy --skill study-buddy -g" },
      { h: "安装 Python 依赖", p: ["在你的助手确认过的 Python 环境里运行："], code: "python -m pip install pymupdf==2.0 pywin32==312 requests python-docx==1.2.0 markdown-it-py pillow" },
      { h: "使用流程", steps: [
        ["整理资料", "把课件、教材、往年题放进该课的“原始资料”文件夹，确认范围后，助手复制改名到“课程资料”并转成讲义"],
        ["规划学期", "提供评分构成和截止时间，算出及格路径和每周清单"],
        ["跟课学习", "说一句“带我学下一节”，白话讲解、小测、错题本"],
        ["期末速通", "按剩余时间和往年题生成速通计划和模拟卷"],
        ["Presentation", "大纲、讲稿、Q&A，生成pptx；小组只做你负责的部分"],
        ["课程论文", "完整论文，引用按 GB/T 7714，转成 docx 或 pdf"],
        ["完成作业", "case study、task、习题，提交语言每次都先问你"],
      ] },
      { h: "配置 OCR 与 VLM", p: ["第一次整理资料时，会在 study-buddy-data/ 下生成 config.json。OCR 引擎默认用百度 Paddle，需要自己申请免费 key，不填则退用 SoMark 的 OXR；VLM 用来转写课件里的图片，不填也能用，只是图片不会被转写。", "key 由你自己填写，助手不会询问、读取或复述。"] },
      { h: "开始对话", p: ["安装完成后，直接在助手里说需求，例如：", "「老师发了一堆课件，帮我整理一下」", "「后天考微观，还没看过」"] },
      { h: "需要知道的限制", p: ["目标是及格，不追求高分或学透", "考试或限时测验进行中不代答，考前复习和模拟照常帮忙", "依赖 Windows 和 Office，转换 PPT、DOC 和生成 pptx 都要用到"] },
    ],
    feedback: { text: "遇到问题或想反馈建议，欢迎到 GitHub 留言。", label: "前往 Issues", url: "https://github.com/Liunian06/study-buddy/issues" },
  },
  screens: [
    {
      id: "changelog",
      title: "更新日志",
      subtitle: "每次更新都有记录。",
      changelog: [
        { version: "1.0.0", date: "2026-10-29", items: ["初始版本，包含工作流：资料整理、学期规划、跟课学习、期末速通、Presentation、课程论文、完成作业"] },
      ],
    },
  ],
};
