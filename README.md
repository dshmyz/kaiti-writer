# kaiti-writer

一个用于 AI 智能体的技能，辅助研究生撰写学位论文开题报告。支持 Claude Code、Trae、Codex 等多个平台。

## 功能特点

- **选题引导**：四问齐发（行业→细分→领域→材料）、领域映射表动态给选项、候选卡 AskUserQuestion 点选、题目过大先拆、换一批机制
- **文献检索**：三级通道（开放API → Playwright知网直搜 → WebSearch+百度学术兜底）→ GB/T 7714 整理 → 文献综述每3篇必有1句评价句
- **逐节起草**：质量闸控制、每节写完即去AI化、篇幅档位控制（精简/标准/充实）、扩写优先级（数据>案例>机制>同义替换）
- **问卷/访谈**：自动拟问卷量表与访谈提纲，插入附录
- **技术路线图**：route.json → Pillow PNG（第5步，docx生成之前）
- **docx 生成**：基于官方 Word 模板（改而不建，保留全部样式）、空节检测、标题长度检查、改稿diff输出
- **PPT 生成**：内容密度规范（12-15页/8分钟）、从content.json自动派生、选最大文本框防溢出、字号自适应、字体统一美化
- **导师反馈循环**：散乱意见 → 结构化清单 → 逐条执行 → 自查
- **一站式收尾**：PPT → 演讲稿（逐页配PPT）→ 答辩问答（最后写，信息最全），默认Word输出
- **断点续传**：progress.json 记录进度，新会话可从上次停下的地方继续
- **评委挑战模拟**：自查后扮演严厉评委追问4个方向，帮用户准备答辩

---

## 快速安装

### 方式 1：一键安装脚本（最简单）

不需要 git，只需要一条命令：

```bash
curl -fsSL https://raw.githubusercontent.com/dshmyz/kaiti-writer/main/install.sh | bash
```

或者下载后运行：

```bash
curl -fsSL -o install.sh https://raw.githubusercontent.com/dshmyz/kaiti-writer/main/install.sh
bash install.sh
```

### 方式 2：git clone

```bash
git clone https://github.com/dshmyz/kaiti-writer.git
cd kaiti-writer
bash install.sh
```

### 方式 3：下载 ZIP

1. 访问 https://github.com/dshmyz/kaiti-writer
2. 点击绿色 **"Code"** 按钮 → **"Download ZIP"**
3. 解压后运行 `bash install.sh`

### 方式 4：指定平台安装

```bash
bash install.sh --platform claude      # Claude Code
bash install.sh --platform trae        # Trae
bash install.sh --platform codebuddy   # 腾讯 CodeBuddy
bash install.sh --platform lingma      # 阿里通义灵码
bash install.sh --all                  # 所有平台
```

## 安装方式对比

| 方式 | 难度 | 需要 git | 适用人群 |
|------|------|---------|---------|
| **一键脚本** | 最简单 | 不需要 | 所有人 |
| **下载 ZIP** | 简单 | 不需要 | 不会命令行的用户 |
| **git clone** | 中等 | 需要 | 有基础的用户 |

---

## 支持的平台

### 国际平台

| 平台 | 支持程度 | 安装路径 | 说明 |
|------|---------|---------|------|
| **Claude Code** | 完全支持 | `~/.claude/skills/` | 推荐 |
| **Trae** | 完全支持 | `~/.trae/builtin_skills/` | 格式兼容 |
| **Codex** | 完全支持 | `~/.codex/skills/` | 自动生成 AGENTS.md |

### 国内平台

| 平台 | 公司 | 支持程度 | 规则格式 |
|------|------|---------|---------|
| **CodeBuddy** | 腾讯 | 支持 | Markdown 规则文件 |
| **通义灵码** | 阿里 | 支持 | Markdown 规则文件 |
| **CodeGeeX** | 智谱AI | 支持 | Markdown 规则文件 |
| **Fitten Code** | — | 支持 | Markdown 规则文件 |

---

## 全流程说明

技能的工作流程分 **9 步**，支持断点续传（中断后可从上次停下的地方继续）：

```
┌─────────────────────────────────────────────────────┐
│ 入口判断：起草新稿 vs 审已有稿（含导师反馈）         │
└────────────────────────┬────────────────────────────┘
                         ▼
┌─────────────────────────────────────────────────────┐
│ 第0步：引导式提问                                     │
│  Round 1：学术方向 + 题目状态                         │
│  Round 2：行业大类→细分→领域→材料（四问齐发）         │
│           领域按行业映射表动态给选项（表A/表B）        │
│  Round 3：定层级 + 写法偏好                           │
│  收束：研究层次/论文类型/问卷访谈/篇幅档位            │
│  一句话摘要检验                                       │
└────────────────────────┬────────────────────────────┘
                         ▼
┌─────────────────────────────────────────────────────┐
│ 第1步：定边界（论文类型决策树复核）                    │
│ 第2步：联网检索文献（三级通道+百度学术兜底）           │
│ 第3步：搭大纲（大纲自查四问 → 一句话摘要确认）        │
│ 第4步：逐节起草（每节写完即去AI化检查）               │
│ 第5步：生成技术路线图（route.json → PNG）             │
│ 第6步：生成docx（content.json → 官方模板）            │
│ 第7步：自查（7项必备+逐节质量闸+篇幅+格式）           │
│ 第7.5步：评委挑战模拟（可选，扮演评委追问4方向）      │
└────────────────────────┬────────────────────────────┘
                         ▼
┌─────────────────────────────────────────────────────┐
│ 第8步：一站式收尾（PPT → 演讲稿 → 答辩问答）         │
│  PPT: content.json自动派生 → 密度校验 → 字号自适应    │
│  演讲稿: 逐页配PPT，Word输出                         │
│  答辩问答: 基于完整论文+PPT，最后写                   │
│ 第9步：查重提示                                      │
└─────────────────────────────────────────────────────┘
```

**断点续传**：每完成一个大步骤自动写 `progress.json`，新会话启动时读取并问用户是否继续。支持撤销/回退到上一步、换题时清理旧文件。

### 项目结构

```
kaiti-writer/
├── SKILL.md                          # 技能入口（调度逻辑全在这里）
├── README.md                         # 使用文档
├── CHANGELOG.md                      # 版本变更记录
├── install.sh                        # 多平台安装脚本
├── .gitignore
├── assets/templates/
│   ├── 开题报告模板.docx             # 官方 Word 模板
│   └── ppt/
│       ├── 模板1-北航答辩通用.pptx   # PPT 模板（文件名含学校名，按需重命名）
│       ├── 模板2-北航答辩通用.pptx
│       └── 模板3-北航答辩通用.pptx
├── references/
│   ├── 选题引导.md                   # 四问脚本 + 领域映射表 + 候选卡
│   ├── 文献检索.md                   # 三级检索通道 + 百度学术兜底
│   ├── 问卷与访谈提纲.md             # 问卷量表/访谈提纲设计规范
│   ├── 篇幅档位.md                   # 精简/标准/充实三档 + 扩写优先级
│   ├── 论文类型详解.md               # 四类论文结构与写作要点
│   ├── 格式规范.md                   # 字体字号页面等排版要求
│   ├── 参考文献著录.md               # GB/T 7714 著录规范
│   ├── 去AI化写作守则.md             # 去 AI 化规则
│   ├── 自查清单.md                   # 输出前逐条核对清单
│   ├── 审阅已有稿.md                 # 审阅/修订已有报告流程
│   ├── 一站式收尾.md                 # PPT/演讲稿/答辩问答 + 内容密度规范
│   └── ars-imports/                  # ARS 学术方法论（CC BY-NC 4.0）
├── scripts/
│   ├── build_from_template.py        # docx 生成（lxml）+ diff + 空节检测
│   ├── build_ppt_from_template.py    # pptx 生成 + 密度校验 + 字号自适应 + 美化
│   ├── derive_ppt_content.py         # 从 content.json 自动派生 ppt_content.json
│   ├── make_route_figure.py          # 技术路线图（Pillow）
│   └── read_docx.py                  # 已有 docx 审阅（只读）
└── templates/
    ├── 开题报告大纲.md               # 内容草稿脚手架
    ├── 范例-问题研究型.md            # 参考范文
    └── 内容示例-*.json               # 四类各一 content.json 示例
```

### 核心组件

| 文件 | 用途 |
|------|------|
| `SKILL.md` | 技能定义，包含触发条件、9步完整工作流、断点续传、进度可视化 |
| `build_from_template.py` | docx 生成引擎，解析模板、填充内容、保留格式、改稿diff、空节检测 |
| `build_ppt_from_template.py` | pptx 生成引擎，密度校验、字号自适应、字体美化 |
| `derive_ppt_content.py` | 从 content.json 自动派生 ppt_content.json |
| `make_route_figure.py` | 技术路线图 PNG 生成 |
| `read_docx.py` | 已有 docx 只读分析，按节切分 |

### 使用方法

#### 方法 1：通过 AI 智能体使用（推荐）

1. 安装技能到你的 AI 智能体
2. 重启智能体
3. 输入：`帮我写一篇 MPA 开题报告，题目是「xxx」`
4. 按提示完成选题、大纲、起草全流程

#### 方法 2：直接使用脚本

```bash
SKILL=~/.claude/skills/kaiti-writer

# 生成开题报告 .docx
python "$SKILL/scripts/build_from_template.py" \
  --template "$SKILL/assets/templates/开题报告模板.docx" \
  --content content.json \
  --output "开题报告-<题目>.docx"

# 生成汇报 PPT
python "$SKILL/scripts/build_ppt_from_template.py" \
  --template "$SKILL/assets/templates/ppt/模板1-北航答辩通用.pptx" \
  --content ppt_content.json \
  --output "开题汇报-<题目>.pptx"

# 生成技术路线图
python "$SKILL/scripts/make_route_figure.py" \
  --content route.json \
  --output 路线图.png

# 审阅已有 docx
python "$SKILL/scripts/read_docx.py" \
  --input "用户给的路径.docx"
```

### content.json 格式

```json
{
  "title": "论文题目",
  "author": "姓名",
  "cover": {
    "title": "论文题目",
    "student": "姓名",
    "advisor": "导师姓名",
    "major": "公共管理",
    "date": "2026年08月"
  },
  "abstract": "摘要正文...",
  "keywords": "关键词1；关键词2；关键词3",
  "content_by_section": {
    "一、选题依据与意义": ["段落1...", "段落2..."],
    "二、国内外研究现状": ["段落1..."],
    "三、研究内容与方法": ["段落1..."],
    "四、研究计划与进度": ["段落1..."],
    "五、预期目标和成果": ["段落1..."],
    "参考文献": ["[1] 作者. 题名[J]. 期刊, 年, 卷(期): 页码."]
  }
}
```

### 论文类型决策树

| 类型 | 适用场景 | 核心方法 |
|------|---------|---------|
| **问题研究型** | 有明确问题、可公开数据 | 问题分析 + 对策建议 |
| **案例分析型** | 能进入具体社区/街道 | 案例描述 + 经验提炼 |
| **调研分析型** | 能发放 ≥200 份问卷 | 问卷调查 + 统计分析 |
| **政策分析型** | 有政策文本可获取 | 政策梳理 + 效果评估 |

### 去AI化规则

写入时即执行，不等自查。核心规则：

- 禁用词：进一步、切实、多措并举、技术赋能、筑牢、深度融合、新范式
- 禁止三句以上平行排比
- 禁止抽象堆叠（"具有重要的理论意义和现实意义"）
- 文献综述必须分类梳理，不能罗列
- 研究方法必须可操作（对象 + 工具 + 步骤 + 信效度）

### 常见问题

**Q: 提示 `ModuleNotFoundError: No module named 'lxml'`**
```bash
pip install lxml
```

**Q: 提示 `ModuleNotFoundError: No module named 'pptx'`**
```bash
pip install python-pptx
```

**Q: Word 打开生成的 docx 提示"需要修复"**
通常是用 stdlib `xml.etree` 替代 lxml 造成的命名空间丢失。确认脚本用的是 `from lxml import etree`。

**Q: 路线图报"找不到可用的中文字体"**
需要系统安装任一中文宋体/黑体。macOS 用系统自带字体；Linux `apt install fonts-arphic-uming`。

**Q: PPT 里还显示别人的名字**
脚本会自动清空模板 docProps。若仍有残留，确认 `python-pptx` 版本 ≥ 0.6.18。

**Q: 非 Claude 环境下怎么用？**
SKILL.md 顶部有工具名映射表：`AskUserQuestion` → 编号选项 + 停下等回答；`WebSearch`/`WebFetch` → `curl` 直连 OpenAlex/Crossref。详见 README 跨平台章节。

---

## 依赖

| 脚本 | 依赖 | 备注 |
|------|------|------|
| `build_from_template.py` | `lxml`, `python-docx`(diff用) | 必须用 lxml，不能用 stdlib `xml.etree` |
| `build_ppt_from_template.py` | `python-pptx` | 生成 PPT 时才需要 |
| `derive_ppt_content.py` | （无） | 纯 stdlib |
| `make_route_figure.py` | `Pillow` | 生成路线图时才需要 |
| `read_docx.py` | （无） | stdlib `xml.etree` 足够 |

---

## 贡献指南

欢迎贡献！请遵循以下步骤：

1. Fork 本仓库
2. 创建特性分支：`git checkout -b feature/amazing-feature`
3. 提交更改：`git commit -m 'Add amazing feature'`
4. 推送分支：`git push origin feature/amazing-feature`
5. 创建 Pull Request

---

## 更新日志

### v1.1.0 (2026-08-10)

**选题引导升级**
- 第2轮从三问升级为四问（行业大类→细分→领域→材料）
- 领域选项从固定4个改为按行业映射表动态给（表A+表B+衔接规则）
- 候选卡改用 AskUserQuestion 呈现，用户直接点选

**文献检索增强**
- 新增 Playwright 知网直搜通道（先检测环境，不可用则问用户是否安装）
- 新增百度学术兜底（不触发域名拦截）
- 文献综述每3篇必有1句评价句硬标准

**PPT 生成重写**
- 内容密度规范（12-15页/8分钟，3-5条bullets×25字）
- 从 content.json 自动派生 ppt_content.json（derive_ppt_content.py）
- 选最大文本框 + 字号自适应防溢出
- 统一字体（宋体+Calibri）和段间距美化

**流程完善**
- 新增第5步「生成技术路线图」
- 去AI化从模糊中间扫描改为每节写完硬性节点
- 大纲确认后加一句话摘要检验
- 自查后加评委挑战模拟（扮演严厉评委追问4方向）
- progress.json 断点续传 + 进度条可视化
- 撤销/回退指引 + 换题清理机制
- 收尾物生成顺序：PPT→演讲稿→答辩问答
- 标题长度检查（>25字提示精简）
- 改稿 diff 输出（--diff 参数）
- 空节检测（正文<50字警告）

**脚本质量**
- Python 编码护栏（open 用 encoding、json.dump 用 ensure_ascii）
- 模板版本跟踪（打印 MD5）
- 错误输出统一（❌ 前缀）

### v1.0.0 (2026-08-10)

- 选题引导：三轮固定提问、候选卡六栏、换一批机制
- 文献检索：联网真实文献 + GB/T 7714
- 逐节起草：质量闸 + 去AI化 + 篇幅档位
- docx/PPT/路线图生成
- 导师反馈循环：散乱意见 → 结构化清单 → 逐条执行
- 大纲大改分支：最多退回两次 + 收敛策略
- 多平台支持：Claude Code / Trae / Codex / 国内 IDE
- 一键安装脚本

---

## 许可证

- ARS 学术方法论（`references/ars-imports/`）：[CC BY-NC 4.0](https://creativecommons.org/licenses/by-nc/4.0/)
- 官方模板及答辩 PPT 模板：仅供个人学位论文使用，请勿重新分发
- 其余文件：MIT License

---

## 免责声明

### 学术诚信

- **本工具仅供参考和学习使用**，不建议直接用 AI 生成开题报告并提交
- 使用 AI 工具辅助理解选题方向、梳理文献、理清研究思路是合理的，但**直接提交 AI 生成的内容可能违反学术诚信要求**
- 各学校对 AI 工具的使用政策不同，请先了解并遵守你所在学校的规定
- 本工具生成的内容**不能替代个人的思考、研究和写作过程**

### 正确使用方式

- 使用本工具**学习开题报告的结构和格式规范**
- 参考生成的内容**理解如何组织论证**
- 用生成的框架作为**草稿**，再进行个人化修改和完善
- 借鉴格式和排版**提高文档制作效率**

### 责任声明

- 本工具开发者不对因使用本工具而产生的任何学术后果负责
- 用户应自行判断使用方式，并承担相应责任
- 本工具按"现状"提供，不保证生成内容的学术适用性
