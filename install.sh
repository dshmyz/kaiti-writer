#!/bin/bash
# 多平台技能安装脚本
# 支持：Claude Code, Trae, Codex, 国内平台
# 用法: bash install.sh [--platform PLATFORM]

set -e

# 颜色定义
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
RED='\033[0;31m'
NC='\033[0m'

echo "🚀 kaiti-writer 安装程序"
echo ""

SKILL_NAME="kaiti-writer"

# 检测平台
detect_platform() {
    if [ -d "$HOME/.claude" ]; then
        echo "claude"
    elif [ -d "$HOME/.trae" ]; then
        echo "trae"
    elif [ -d "$HOME/.codex" ]; then
        echo "codex"
    else
        echo "unknown"
    fi
}

# 核心文件列表
copy_common_files() {
    local TARGET_DIR="$1"
    echo "📦 复制文件..."

    cp ./SKILL.md "$TARGET_DIR/"
    cp ./README.md "$TARGET_DIR/"
    cp ./.gitignore "$TARGET_DIR/"

    # scripts
    mkdir -p "$TARGET_DIR/scripts"
    cp ./scripts/*.py "$TARGET_DIR/scripts/"

    # references
    mkdir -p "$TARGET_DIR/references/ars-imports"
    cp ./references/*.md "$TARGET_DIR/references/"
    cp ./references/ars-imports/*.md "$TARGET_DIR/references/ars-imports/"

    # assets
    mkdir -p "$TARGET_DIR/assets/templates/ppt"
    cp ./assets/templates/*.docx "$TARGET_DIR/assets/templates/"
    cp ./assets/templates/ppt/*.pptx "$TARGET_DIR/assets/templates/ppt/"

    # templates
    mkdir -p "$TARGET_DIR/templates"
    cp ./templates/*.md "$TARGET_DIR/templates/"
    cp ./templates/*.json "$TARGET_DIR/templates/"
}

# 安装到 Claude Code
install_claude() {
    echo "📦 安装到 Claude Code..."
    TARGET_DIR="$HOME/.claude/skills/$SKILL_NAME"

    if [ -d "$TARGET_DIR" ]; then
        read -p "⚠️  目录已存在，是否覆盖? (y/N): " confirm
        if [[ ! $confirm =~ ^[Yy]$ ]]; then
            echo "❌ 安装取消"
            exit 0
        fi
        rm -rf "$TARGET_DIR"
    fi

    mkdir -p "$TARGET_DIR"
    copy_common_files "$TARGET_DIR"

    echo "✅ Claude Code 安装完成"
    echo "💡 重启 Claude Code 即可使用"
}

# 安装到 Trae
install_trae() {
    echo "📦 安装到 Trae..."
    TARGET_DIR="$HOME/.trae/builtin_skills/$SKILL_NAME"

    if [ -d "$TARGET_DIR" ]; then
        read -p "⚠️  目录已存在，是否覆盖? (y/N): " confirm
        if [[ ! $confirm =~ ^[Yy]$ ]]; then
            echo "❌ 安装取消"
            exit 0
        fi
        rm -rf "$TARGET_DIR"
    fi

    mkdir -p "$TARGET_DIR"
    copy_common_files "$TARGET_DIR"

    echo "✅ Trae 安装完成"
    echo "💡 重启 Trae 即可使用"
}

# 安装到 Codex（AGENTS.md 格式）
install_codex() {
    echo "📦 安装到 Codex..."
    TARGET_DIR="$HOME/.codex/skills/$SKILL_NAME"

    if [ -d "$TARGET_DIR" ]; then
        read -p "⚠️  目录已存在，是否覆盖? (y/N): " confirm
        if [[ ! $confirm =~ ^[Yy]$ ]]; then
            echo "❌ 安装取消"
            exit 0
        fi
        rm -rf "$TARGET_DIR"
    fi

    mkdir -p "$TARGET_DIR"
    copy_common_files "$TARGET_DIR"

    # 生成 AGENTS.md
    echo "📝 生成 AGENTS.md..."
    cat > "$TARGET_DIR/AGENTS.md" << 'AGENTSEOF'
# Thesis Proposal Writer

遇到"开题报告""MPA 论文""选题""推荐开题方向"等关键词时，按 SKILL.md 执行。
AGENTSEOF

    echo "✅ Codex 安装完成"
    echo "💡 重启 Codex 即可使用"
}

# 安装到国内平台（转换为规则文件）
install_rules_platform() {
    local PLATFORM_NAME="$1"
    local TARGET_DIR="$2"
    local FILE_NAME="$3"

    echo "📦 安装到 $PLATFORM_NAME..."
    mkdir -p "$TARGET_DIR"

    python3 -c "
import re
from pathlib import Path

skill_md = Path('./SKILL.md').read_text(encoding='utf-8')
skill_md = re.sub(r'^---.*?---\n', '', skill_md, flags=re.DOTALL)

rules_md = f'''# kaiti-writer

> 研究生学位论文开题报告写作技能

{skill_md}
'''

Path('$TARGET_DIR/$FILE_NAME').write_text(rules_md, encoding='utf-8')
print('✅ 已转换为 $PLATFORM_NAME 规则')
"

    echo "✅ $PLATFORM_NAME 安装完成"
    echo "💡 重启 IDE 即可使用"
}

# 显示帮助
show_help() {
    echo "用法: bash install.sh [选项]"
    echo ""
    echo "选项:"
    echo "  -p, --platform PLATFORM   指定安装平台"
    echo "  -a, --all                 安装到所有检测到的平台"
    echo "  -l, --list                列出所有支持的平台"
    echo "  -h, --help                显示此帮助信息"
    echo ""
    echo "支持的平台:"
    echo "  Claude Code (推荐)    ~/.claude/skills/"
    echo "  Trae                  ~/.trae/builtin_skills/"
    echo "  Codex (OpenAI)        ~/.codex/skills/"
    echo "  CodeBuddy (腾讯)      .codebuddy/rules/"
    echo "  通义灵码 (阿里)       .lingma/rules/"
    echo "  CodeGeeX (智谱AI)     .codegeex/"
    echo "  Fitten Code           ~/.fitten/rules/"
    echo ""
    echo "示例:"
    echo "  bash install.sh                        # 自动检测并安装"
    echo "  bash install.sh --platform claude      # 安装到 Claude Code"
    echo "  bash install.sh --all                  # 安装到所有平台"
}

# 列出支持的平台
list_platforms() {
    echo "支持的平台:"
    echo ""
    echo "国际平台:"
    echo "  claude    Claude Code (推荐)"
    echo "  trae      Trae (ByteDance)"
    echo "  codex     Codex (OpenAI)"
    echo ""
    echo "国内平台:"
    echo "  codebuddy 腾讯 CodeBuddy"
    echo "  lingma    阿里 通义灵码"
    echo "  codegeex  智谱AI CodeGeeX"
    echo "  fitten    Fitten Code"
    echo ""
    PLATFORM=$(detect_platform)
    if [ "$PLATFORM" != "unknown" ]; then
        echo "✅ 检测到平台: $PLATFORM"
    else
        echo "⚠️  未检测到已知平台，请用 --platform 指定"
    fi
}

# 解析参数
PLATFORM=""
INSTALL_ALL=false

while [ $# -gt 0 ]; do
    case $1 in
        -p|--platform)
            PLATFORM="$2"
            shift 2
            ;;
        -a|--all)
            INSTALL_ALL=true
            shift
            ;;
        -l|--list)
            list_platforms
            exit 0
            ;;
        -h|--help)
            show_help
            exit 0
            ;;
        *)
            echo "❌ 未知选项: $1"
            show_help
            exit 1
            ;;
    esac
done

# 自动检测
if [ -z "$PLATFORM" ] && [ "$INSTALL_ALL" = false ]; then
    PLATFORM=$(detect_platform)
    if [ "$PLATFORM" = "unknown" ]; then
        echo "❌ 未检测到已知平台"
        echo "💡 请使用 --platform 指定平台"
        list_platforms
        exit 1
    fi
    echo "🔍 自动检测到平台: $PLATFORM"
    echo ""
fi

# 执行安装
if [ "$INSTALL_ALL" = true ]; then
    echo "📦 安装到所有平台..."
    echo ""
    [ -d "$HOME/.claude" ] && install_claude && echo ""
    [ -d "$HOME/.trae" ] && install_trae && echo ""
    [ -d "$HOME/.codex" ] && install_codex && echo ""
    echo "✅ 所有平台安装完成"
else
    case $PLATFORM in
        claude)   install_claude ;;
        trae)     install_trae ;;
        codex)    install_codex ;;
        codebuddy) install_rules_platform "CodeBuddy" ".codebuddy/rules" "kaiti-writer.md" ;;
        lingma)   install_rules_platform "通义灵码" ".lingma/rules" "kaiti-writer.md" ;;
        codegeex) install_rules_platform "CodeGeeX" ".codegeex" "rules.md" ;;
        fitten)   install_rules_platform "Fitten Code" "$HOME/.fitten/rules" "kaiti-writer.md" ;;
        *)
            echo "❌ 未知平台: $PLATFORM"
            list_platforms
            exit 1
            ;;
    esac
fi

echo ""
echo "🎉 安装完成!"
echo "📖 使用文档: $(pwd)/README.md"
echo ""
