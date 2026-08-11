# kaiti-writer

**帮你写开题报告的 AI 技能。** 装上之后，跟 AI 说"帮我写一篇开题报告"，它会一步一步引导你——从选题目到交稿，全程帮你搞定。

## 它能帮你做什么

一句话：**从"还没想好题目"到"交稿"的全流程。**

具体来说：

| 你要做的事 | 它怎么帮你 |
|---|---|
| **选题目** | 问你几个问题（什么行业、什么方向、能拿到什么数据），然后给你 3 个题目建议，告诉你哪个最好写、风险在哪 |
| **查文献** | 自动上网搜知网、百度学术、英文数据库，帮你找真实论文，整理成标准格式 |
| **搭大纲** | 按学校模板帮你列好提纲，确认后还会用一句话总结"你这篇论文到底要干什么" |
| **写正文** | 一节一节帮你写，写完就检查有没有 AI 味太重的句子，当场改掉 |
| **生成 Word** | 自动套学校官方模板，格式、字号、行距全都对，不用你调 |
| **生成 PPT** | 帮你做汇报用的答辩 PPT，12-15 页，自动从 Word 内容里提炼要点 |
| **准备答辩** | 帮你写演讲稿（逐字稿）、预测评委可能问的 10-15 个问题并教你怎么答 |
| **改稿** | 导师让改？把导师意见扔给它，它帮你逐条改，改完还能对比"改了哪里" |

**一句话：你负责想，它负责写和排版。**

## 装上就能用

只需要一条命令：

```bash
curl -fsSL https://raw.githubusercontent.com/dshmyz/kaiti-writer/main/install.sh | bash
```

装完重启你的 AI 智能体（Claude Code / Trae / Codex 等），然后直接说：

> 帮我写一篇 MPA 开题报告

它就会开始引导你了。

<details>
<summary>其他安装方式（点击展开）</summary>

### git clone

```bash
git clone https://github.com/dshmyz/kaiti-writer.git
cd kaiti-writer
bash install.sh
```

### 下载 ZIP

1. 访问 https://github.com/dshmyz/kaiti-writer
2. 点绿色 **"Code"** → **"Download ZIP"**
3. 解压后运行 `bash install.sh`

### 指定平台安装

```bash
bash install.sh --platform claude      # Claude Code
bash install.sh --platform trae        # Trae
bash install.sh --platform codebuddy   # 腾讯 CodeBuddy
bash install.sh --platform lingma      # 阿里通义灵码
bash install.sh --all                  # 所有平台
```

</details>

## 怎么用

装好之后，打开你的 AI 智能体，直接用大白话说你要干什么：

- `帮我写一篇 MPA 开题报告，我想写社区治理方向`
- `我还没想好题目，帮我选一个`
- `我有一份写了一半的开题报告，帮我看看`
- `导师说我的综述写得不好，帮我改`

**它不会替你做决定**——它会问你问题（什么行业、能拿到什么数据、想写多长），然后给你建议，你点头它才往下写。

## 它的工作流程

```
选题目（问你几个问题 → 给你3个建议）
    ↓
查文献（上网搜真实论文 → 整理格式）
    ↓
搭大纲（列提纲 → 一句话确认"到底要写什么"）
    ↓
写正文（一节一节写 → 每节去AI味）
    ↓
生成 Word（套学校模板 → 自动排版）
    ↓
自查（检查格式/内容/占位 → 给你清单）
    ↓
收尾（PPT + 演讲稿 + 答辩问答 → 全部搞定）
```

**中途断了？** 下次打开会问你"上次停在XX步，要继续吗？"——不会让你从头来。

## 它默认支持谁

默认是**北航 MPA（公共管理）**的格式和模板。但如果你是别的学校，可以自己换模板——技能会自动识别你的模板格式。

<details>
<summary>怎么换自己学校的模板（点击展开）</summary>

1. 把你学校的开题报告 Word 模板放到 `assets/templates/开题报告模板.docx`
2. 打开你的模板，看看里面有哪几节标题（比如"一、选题背景"、"二、研究意义"之类的）
3. 在生成 `content.json` 时，把节标题写成你模板里的标题就行

技能的脚本是按文本匹配节标题的，不限定具体是哪几节——你模板里有什么标题，它就能往什么标题后面填内容。

</details>

## 依赖

核心功能只需要 Python 3.8+。按需安装：

| 你想要什么 | 装什么 | 命令 |
|---|---|---|
| **生成 Word 开题报告** | lxml | `pip install lxml` |
| **生成 PPT 汇报** | python-pptx | `pip install python-pptx` |
| **生成技术路线图** | Pillow | `pip install Pillow` |
| **知网直搜文献** | playwright | `pip install playwright && playwright install chromium` |

不装也能用——技能会在需要时问你"要不要安装"。

## 常见问题

**Q: 它会直接帮我交作业吗？**
不会。它生成的是草稿，你需要自己检查、修改、确认。学术诚信是底线。

**Q: 我不是北航的能用吗？**
能用。默认是北航格式，但你可以换自己学校的模板（见上面）。

**Q: 文献是真的吗？**
是的。它从知网、百度学术、OpenAlex 等数据库搜索真实论文。搜不到的会标"待核验"，不会编造。

**Q: 生成的 PPT 好看吗？**
基于北航答辩模板，自动清水印、填内容、调字号。如果模板本身好看，生成的就不会差。

**Q: 中途断了怎么办？**
下次打开会自动问你"上次停在XX步，要继续吗？"。

**Q: 我的 AI 没有 AskUserQuestion 工具怎么办？**
技能会降级成在回复里写编号选项（A/B/C），你回复字母就行。

## 项目结构

```
kaiti-writer/
├── SKILL.md                    # 技能核心（调度逻辑全在这里）
├── README.md                   # 你现在看的这个文件
├── CHANGELOG.md                # 版本更新记录
├── install.sh                  # 一键安装脚本
├── assets/templates/
│   ├── 开题报告模板.docx       # Word 模板（可以换成你学校的）
│   └── ppt/
│       ├── 模板1-北航答辩通用.pptx
│       ├── 模板2-北航答辩通用.pptx
│       └── 模板3-北航答辩通用.pptx
├── references/                 # 技能的各种规范和指引
│   ├── 选题引导.md             # 怎么帮你选题目
│   ├── 文献检索.md             # 怎么帮你搜文献
│   ├── 篇幅档位.md             # 写多长（精简/标准/充实）
│   ├── 去AI化写作守则.md       # 怎么去掉AI味
│   └── ...
├── scripts/                    # 自动生成文件的脚本
│   ├── build_from_template.py  # 生成 Word
│   ├── build_ppt_from_template.py  # 生成 PPT
│   ├── derive_ppt_content.py   # 从 Word 内容自动提炼 PPT 要点
│   ├── make_route_figure.py    # 生成技术路线图
│   └── read_docx.py            # 读取并分析已有 Word
└── templates/                  # 内容示例
    └── 内容示例-*.json          # 各类论文的示例数据
```

## 许可证

- ARS 学术方法论（`references/ars-imports/`）：CC BY-NC 4.0
- 官方模板及答辩 PPT 模板：仅供个人学位论文使用，请勿重新分发
- 其余文件：MIT License

## 免责声明

这个工具帮你写草稿、排格式、整理文献，**但不替你做研究**。生成的内容需要你自己检查、修改、确认后再提交。各学校对 AI 工具的使用政策不同，请先了解并遵守你所在学校的规定。
