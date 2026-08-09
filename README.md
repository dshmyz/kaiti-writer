# 学位论文开题报告 Claude Skill

撰写、审阅、修订**研究生学位论文开题报告**，面向 MPA/社科方向（理工可切编号）。覆盖从选题引导到答辩 PPT 的完整闭环。

## 功能概览

| 模块 | 做什么 | 入口文件 |
|------|--------|----------|
| 选题引导 | 三轮固定提问脚本、候选卡六栏格式、题目过大先拆、大纲自查四问 | `references/选题引导.md` |
| 文献检索 | 联网检索真实文献 → GB/T 7714 整理 → 支撑文献综述 | `references/文献检索.md` |
| 起草 | 逐节质量闸、去AI化写入、篇幅档位控制、问卷/访谈自动插入附录 | `SKILL.md`（第 0 步 → 第 5 步） |
| 生成 .docx | 基于官方 Word 模板（改而不建，保留全部样式） | `scripts/build_from_template.py` |
| 生成 PPT | 内置答辩模板可选，自动清水印/占位 | `scripts/build_ppt_from_template.py` |
| 技术路线图 | Pillow 直接出 PNG（零新增依赖） | `scripts/make_route_figure.py` |
| 审阅已有稿 | 只读分析、按节切分、缺失节/缺附录报警 | `scripts/read_docx.py` |
| 自查对照 | 7 项必备 + 逐节质量闸 + 篇幅 + 格式 + 评审标准 | `references/自查清单.md` |

## 目录结构

```
开题报告/
├── SKILL.md                          # 入口指令（技能调度逻辑全在这里）
├── assets/templates/
│   ├── 开题报告模板.docx             # 官方 Word 模板（生成时用，勿改）
│   └── ppt/
│       ├── 模板1-北航答辩通用.pptx   # 三套 PPT 模板（文件名含学校名，按需重命名）
│       ├── 模板2-北航答辩通用.pptx
│       └── 模板3-北航答辩通用.pptx
├── references/
│   ├── 选题引导.md                   # 三轮提问脚本 + 候选卡 + 收窄 + 大纲自查
│   ├── 文献检索.md                   # 联网检索完整流程
│   ├── 问卷与访谈提纲.md             # 问卷量表/访谈提纲设计规范
│   ├── 篇幅档位.md                   # 精简/标准/充实三档字数控制
│   ├── 论文类型详解.md               # 四类论文结构与写作要点
│   ├── 格式规范.md                   # 字体字号页面等排版要求
│   ├── 参考文献著录.md               # GB/T 7714 著录规范
│   ├── 去AI化写作守则.md             # 中强度去 AI 化规则（写入时执行）
│   ├── 自查清单.md                   # 输出前逐条核对清单
│   ├── 审阅已有稿.md                 # 审阅/修订已有报告流程
│   ├── 一站式收尾.md                 # PPT / 演讲稿 / 答问预测 / 多格式导出
│   └── ars-imports/                  # ARS 学术方法论（CC BY-NC 4.0）
│       ├── 内置说明.md
│       ├── source_quality_hierarchy.md
│       ├── literature_matrix_template.md
│       ├── literature_monitoring_strategies.md
│       ├── openalex_api_protocol.md
│       └── crossref_api_protocol.md
├── scripts/
│   ├── build_from_template.py        # docx 生成（lxml，非 stdlib ET）
│   ├── build_ppt_from_template.py    # pptx 生成（python-pptx）
│   ├── make_route_figure.py          # 技术路线图（Pillow）
│   └── read_docx.py                  # 已有 docx 审阅（只读）
├── templates/
│   ├── 开题报告大纲.md               # 内容草稿脚手架
│   ├── 范例-问题研究型.md             # 参考范文
│   ├── 内容示例-社区减负.json         # 四类各一 content.json 示例
│   ├── 内容示例-案例分析.json
│   ├── 内容示例-调研分析.json
│   └── 内容示例-政策分析.json
└── README.md                         # ← 本文件
```

## 快速开始

### Claude Code（默认）

直接在对话中说：

> 帮我写一篇 MPA 开题报告，题目是「T市基层公安民警职业倦怠与工作投入的关系研究」

技能会自动激活（触发词：开题报告 / MPA 论文 / 推荐开题方向 / 选题 / 没想好论文题目），走完三轮选题 → 大纲 → 起草 → 生成 .docx → 自查 → 收尾（PPT / 演讲稿等）全流程。

### 手动生成（命令行）

```bash
# 1. 定位本技能目录（所有命令通过 $SKILL 引用）
SKILL=~/.claude/skills/开题报告
ls "$SKILL/scripts"   # 应看到 4 个 .py

# 2. 生成开题报告 .docx（先按 templates/内容示例-社区减负.json 格式准备 content.json）
python "$SKILL/scripts/build_from_template.py" \
  --template "$SKILL/assets/templates/开题报告模板.docx" \
  --content content.json \
  --output "开题报告-<题目>.docx"

# 3. 生成汇报 PPT（按 content.json 的 chapters 格式准备 ppt_content.json）
python "$SKILL/scripts/build_ppt_from_template.py" \
  --template "$SKILL/assets/templates/ppt/模板1-北航答辩通用.pptx" \
  --content ppt_content.json \
  --output "开题汇报-<题目>.pptx"

# 4. 生成技术路线图
cat > route.json <<'EOF'
{"nodes":["问题提出",["理论梳理","政策梳理"],"实证分析","对策建议"]}
EOF
python "$SKILL/scripts/make_route_figure.py" --content route.json --output 路线图.png

# 5. 审阅已有 docx
python "$SKILL/scripts/read_docx.py" --input "用户给的路径.docx"
```

## 跨智能体使用

本技能最初为 Claude Code 编写，但**核心逻辑与所有 Python 脚本均不依赖任何 Claude 专有 API**。只需适配两件事：技能存放位置的路径，以及工具名差异。

### 支持的智能体

| 智能体 | 技能存放位置 | 规则文件 | 状态 |
|--------|-------------|---------|------|
| Claude Code | `~/.claude/skills/开题报告/` | `SKILL.md`（自动扫描） | 原生支持 |
| Codex (OpenAI) | 项目根目录任意子目录 | `AGENTS.md` 或 `agents.md` | ✅ 支持 |
| Trae (字节跳动) | 项目根目录任意子目录 | `AGENTS.md`（v3.5.19+）| ✅ 支持 |
| Cursor | 项目根目录 | `AGENTS.md` 或 `.cursorrules` | ✅ 支持 |
| Gemini CLI (Google) | 项目根目录 | `AGENTS.md` | ✅ 支持 |
| Windsurf | 项目根目录 | `AGENTS.md` 或 `.windsurfrules` | ✅ 支持 |
| GitHub Copilot | 项目根目录 | `.github/copilot-instructions.md` 或 `AGENTS.md` | ✅ 支持 |
| WorkBuddy (腾讯) | — | **不支持本地规则文件** | ⚠️ 见下文 |

> **AGENTS.md** 是一个开放规范（[agents.md](https://agents.md)），已被 60k+ 开源项目和所有上述本地编码智能体支持。内容格式与 `SKILL.md` 相同：Markdown，智能体启动时自动读取。

### 安装到其他智能体

**步骤 1 — 复制技能目录到项目内**

```bash
# 复制到项目根目录（以项目名命名，避免语义暴露）
cp -r ~/.claude/skills/开题报告 ./docs/thesis-proposal-skill
```

**步骤 2 — 创建/追加 `AGENTS.md`**

在项目根目录创建 `AGENTS.md`（或追加到已有文件），内容一行即可：

```markdown
# Thesis proposal skill
详见 `docs/thesis-proposal-skill/SKILL.md` — 遇到"开题报告""MPA 论文"等关键词时按其执行。
```

部分智能体（如 Cursor）也支持在 `.cursorrules` 里写同样内容。

**步骤 3 — 设置 `$SKILL` 环境变量**

在 `SKILL.md` 开头或对话里说明：

```bash
SKILL=./docs/thesis-proposal-skill
```

下文所有命令引用 `"$SKILL/..."`，到此已全部可运行。

### 工具名差异对照

SKILL.md 里有两个 Claude 专有工具名，其他智能体下按以下方式等价替换：

| SKILL.md 里的叫法 | 其他智能体怎么做 |
|---|---|
| **`AskUserQuestion`**（一次性列出 2–3 个带选项的问题，等用户回答） | 在回复里写编号选项（如 `A/B/C`，注明"也可自己输入"），然后**停下等用户回答**。三轮脚本内容与顺序不变，只是提问方式不同。 |
| **`WebSearch` / `WebFetch`** | 任意联网检索能力。没有就用 `curl` 直连 OpenAlex / Crossref 开放 API（见 `references/文献检索.md`）。**完全无网时不许编造文献**，按该文件第二级请用户提供。 |

这些替代方案已写在 `SKILL.md` 的顶部注释里，智能体读取时会看到。

### WorkBuddy（腾讯）

WorkBuddy 是腾讯的托管型办公 Agent，**不支持上传本地规则文件或技能包**。若要使用本技能的内容辅助：

1. 打开 WorkBuddy 对话，贴入：

> 我要写 MPA 开题报告。以下是核心约束：开题题目；选题依据（意义/必要性/前沿性）；文献综述；研究方案（目标/内容/关键问题/方法/技术路线）；预期目标和成果；实施计划；主要参考文献，七项缺一不可。研究方法须可操作、文献须真实、写入时即去 AI 化。现在请帮我选题，按以下节奏来：第一问同时问清"学术方向 + 题目状态"；题目状态分岔：已定题 → 直接走大纲；有范围或没想好 → 三轮固定提问脚本（方向/范围/收束）。

2. 需要生成 docx 时，让 WorkBuddy 输出 `content.json`，然后在本地用命令行生成：

```bash
SKILL=~/.claude/skills/开题报告   # 或实际安装路径
python "$SKILL/scripts/build_from_template.py" \
  --template "$SKILL/assets/templates/开题报告模板.docx" \
  --content content.json \
  --output "开题报告.docx"
```

> WorkBuddy 的优势是能直接生成 PPT、Excel 报表等，对于收尾阶段（答辩 PPT / 演讲稿）可以就地完成，不依赖本地脚本。但 Word 模板生成（保留官方样式）仍需本地运行 `build_from_template.py`。

## 依赖

| 脚本 | 依赖 | 备注 |
|------|------|------|
| `build_from_template.py` | `lxml` | `pip install lxml`；必须用 lxml 不能用 stdlib `xml.etree`（后者丢命名空间前缀，Word 会报修复） |
| `build_ppt_from_template.py` | `python-pptx` | `pip install python-pptx`；生成 PPT 时才需要 |
| `make_route_figure.py` | `Pillow` | `pip install Pillow`；macOS/Linux/Windows 系统中文字体需任一可用 |
| `read_docx.py` | （无） | stdlib `xml.etree` 足够，只读不写 |

## 安全与隐私

- **内置模板已清除全部第三方个人信息**（作者姓名、邮箱、单位、修改记录、打印时间），不会把你导师或同学的信息带进生成稿。
- **用户自带模板时照样清理**——`docProps/core.xml` 里的 `dc:creator`、`cp:lastModifiedBy`、`cp:lastPrinted`、`docProps/app.xml` 的 `Company`/`Manager` 均在写入前清空。
- **不许伪造文献**：占位符 `〈〉` 可以出现在示例/大纲里，但生成正式稿时必须清掉。脚本自检会分别报出"正文占位"、"参考文献占位"、"封面占位"。
- **本地文献目录读取**必须由用户提供路径，不许自己猜路径或遍历无关目录。

## 常见问题

**Q: 提示 `ModuleNotFoundError: No module named 'lxml'`**
```bash
pip install lxml
```
如果是 `python-pptx` 缺失也同理：`pip install python-pptx`。技能不会擅自安装，会先问用户。

**Q: Word 打开生成的 docx 提示"需要修复"**
通常是用 stdlib `xml.etree` 替代 lxml 造成的命名空间丢失。确认脚本用的是 `from lxml import etree`（本技能已锁定）。

**Q: 路线图生成报"找不到可用的中文字体"**
需要系统安装任一中文宋体/黑体。macOS 用系统自带字体；Linux `apt install fonts-arphic-uming`；Windows 通常已有 simsun.ttc。

**Q: 生成的 PPT 里还显示别人的名字**
脚本会自动清空模板 docProps，若仍有残留请确认 `python-pptx` 版本 ≥ 0.6.18。

## 许可

- ARS 学术方法论（`references/ars-imports/`）：[CC BY-NC 4.0](https://creativecommons.org/licenses/by-nc/4.0/)
- 官方模板及答辩 PPT 模板：仅供个人学位论文使用，请勿重新分发。
- 其余文件随技能分发，无特殊许可限制。
