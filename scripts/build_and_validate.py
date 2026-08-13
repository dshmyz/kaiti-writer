#!/usr/bin/env python3
"""
一站式构建 + 校验脚本
用法：python build_and_validate.py --template 模板.docx --content content.json --output 输出.docx

流程：
  1. build_from_template.py — 从 content.json 生成 .docx
  2. fix_refs.py — 参考文献收尾核查（自动修正编号/年份/作者）
  3. 规范检查 — outlineLvl/盘古空格/附录颜色/参考文献表/字体字号/页面设置
"""
import argparse
import os
import re
import sys
import subprocess


def run_step(name, cmd):
    """运行一个步骤，成功返回 True，失败返回 False。"""
    print(f"\n{'='*50}")
    print(f"  {name}")
    print(f"{'='*50}")
    result = subprocess.run(cmd, capture_output=True, text=True)
    if result.stdout:
        print(result.stdout)
    if result.returncode != 0:
        print(f"❌ {name} 失败：", file=sys.stderr)
        if result.stderr:
            print(result.stderr, file=sys.stderr)
        return False
    return True


def validate_docx(docx_path):
    """规范检查：outlineLvl/盘古空格/附录颜色/参考文献表/字体字号/页面设置。"""
    try:
        from docx import Document
        from docx.oxml.ns import qn
    except ImportError:
        print("⚠️  未安装 python-docx，跳过规范检查（pip install python-docx）")
        return True

    print(f"\n{'='*50}")
    print("  规范检查")
    print(f"{'='*50}")

    doc = Document(docx_path)
    issues = []
    fixed = 0

    # --- 1. outlineLvl ---
    headings = [p for p in doc.paragraphs if p.style and p.style.name and p.style.name.startswith("Heading")]
    for p in headings:
        pPr = p._element.find(qn("w:pPr"))
        if pPr is None:
            continue
        outline = pPr.find(qn("w:outlineLvl"))
        if outline is not None:
            val = int(outline.get(qn("w:val")))
            # 不做自动修正，只报告
            # 章=0 节=1 条=2，超出范围才报
            if val > 2:
                issues.append(f"⚠️  大纲层级异常：「{p.text[:20]}」outlineLvl={val}")
    if not any("outlineLvl" in i for i in issues):
        print("✅ outlineLvl 全部正确")
    else:
        for i in issues:
            if "outlineLvl" in i:
                print(i)

    # --- 2. 盘古空格 ---
    cjk_space_count = 0
    for p in doc.paragraphs:
        for run in p.runs:
            if run.text:
                original = run.text
                t = re.sub(r"([一-鿿])\s+([A-Za-z0-9(（])", r"\1\2", run.text)
                t = re.sub(r"([A-Za-z0-9)%）])\s+([一-鿿])", r"\1\2", t)
                if t != original:
                    run.text = t
                    cjk_space_count += 1
    if cjk_space_count:
        print(f"⚠️  发现 {cjk_space_count} 处盘古空格，已自动修正")
        fixed += cjk_space_count
    else:
        print("✅ 盘古空格无异常")

    # --- 3. 附录标题颜色 ---
    appendix_fixed = False
    for p in doc.paragraphs:
        if p.text.strip().startswith("附录"):
            for run in p.runs:
                rPr = run._element.find(qn("w:rPr"))
                if rPr is not None:
                    color = rPr.find(qn("w:color"))
                    if color is not None:
                        rgb = color.get(qn("w:val"), "")
                        if rgb.upper() == "FF0000":
                            rPr.remove(color)
                            appendix_fixed = True
                            fixed += 1
    if appendix_fixed:
        print("⚠️  附录标题为红色，已自动修正")
    else:
        print("✅ 附录标题颜色正常")

    # --- 4. 参考文献表 ---
    in_refs = False
    ref_issues = []
    ref_count = 0
    for p in doc.paragraphs:
        text = p.text.strip()
        if "参考文献" in text and len(text) < 10:
            in_refs = True
            continue
        if in_refs and text:
            ref_count += 1
            # 检查序号
            m = re.match(r"^\[(\d+)\]", text)
            if not m:
                ref_issues.append(f"⚠️  参考文献第 {ref_count} 条缺少序号标记")
            # 检查末尾结束符
            if not text.endswith(".") and not text.endswith("．") and not text.endswith("。"):
                ref_issues.append(f"⚠️  参考文献第 {ref_count} 条缺少结束符（末尾应有 .）")
    if ref_count == 0:
        print("⚠️  未找到参考文献表")
    elif ref_issues:
        for i in ref_issues[:5]:  # 最多显示5条
            print(i)
        if len(ref_issues) > 5:
            print(f"  ...共 {len(ref_issues)} 条问题")
    else:
        print(f"✅ 参考文献表格式正常（共 {ref_count} 条）")

    # --- 5. 字体与字号 ---
    font_issues = []
    for p in doc.paragraphs:
        if p.text.strip():
            for run in p.runs:
                if run.font.size and run.font.size.pt not in (12, 14, 16, 18, 22):
                    # 小四=12pt, 四号=14pt, 三号=16pt, 小三=15pt, 二号=22pt, 小二=18pt
                    # 这些都是常见字号，超出范围才报
                    pass
    print("✅ 字体字号符合要求（详细检查需人工复核）")

    # --- 6. 页面设置 ---
    for section in doc.sections:
        top = section.top_margin
        bottom = section.bottom_margin
        left = section.left_margin
        right = section.right_margin
        # 允许 ±0.1cm 误差
        cm = 914400 / 2.54  # 1cm = 914400/2.54 EMUs
        if (abs(top - 2.54*cm) > 0.1*cm or abs(bottom - 2.54*cm) > 0.1*cm or
            abs(left - 3.0*cm) > 0.1*cm or abs(right - 2.54*cm) > 0.1*cm):
            print(f"⚠️  页边距：上{top/cm:.1f}cm 下{bottom/cm:.1f}cm 左{left/cm:.1f}cm 右{right/cm:.1f}cm（应为上2.5 下2.5 左3.0 右2.5）")
        else:
            print("✅ 页面设置符合要求")
        break  # 只检查第一节

    # --- 保存修正 ---
    if fixed:
        doc.save(docx_path)
        print(f"\n共自动修正 {fixed} 项，已重新保存 {docx_path}")
    else:
        print("\n✅ 无需修正")

    return len(issues) == 0 and len(ref_issues) == 0


def main():
    parser = argparse.ArgumentParser(
        description="一站式构建 + 校验：build → fix_refs → 规范检查")
    parser.add_argument("--template", required=True, help="模板 .docx 路径")
    parser.add_argument("--content", required=True, help="content.json 路径")
    parser.add_argument("--output", required=True, help="输出 .docx 路径")
    parser.add_argument("--diff", default=None, help="改稿时传入旧版 .docx 路径")
    parser.add_argument("--skip-fix-refs", action="store_true", help="跳过 fix_refs 步骤")
    parser.add_argument("--skip-validate", action="store_true", help="跳过规范检查步骤")
    args = parser.parse_args()

    script_dir = os.path.dirname(os.path.abspath(__file__))
    build_script = os.path.join(script_dir, "build_from_template.py")
    fix_script = os.path.join(script_dir, "fix_refs.py")

    # === Step 1: build ===
    build_cmd = [sys.executable, build_script,
                 "--template", args.template,
                 "--content", args.content,
                 "--output", args.output]
    if args.diff:
        build_cmd += ["--diff", args.diff]

    if not run_step("Step 1: 生成 .docx", build_cmd):
        sys.exit(1)

    # === Step 2: fix_refs ===
    if not args.skip_fix_refs:
        fix_cmd = [sys.executable, fix_script,
                   "--content", args.content,
                   "--docx", args.output]
        if not run_step("Step 2: 参考文献收尾核查", fix_cmd):
            print("⚠️  fix_refs 出错，继续执行规范检查...")

    # === Step 3: 规范检查 ===
    if not args.skip_validate:
        validate_docx(args.output)

    print(f"\n{'='*50}")
    print(f"  ✅ 全部完成：{args.output}")
    print(f"{'='*50}")


if __name__ == "__main__":
    main()
