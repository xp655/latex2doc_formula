#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Word LaTeX 公式转换器
=====================
将 .docx 文档中的 $...$（行内公式）和 $$...$$（独立公式）
转换为 Word 原生公式（OMML 格式），可在 Word 中直接编辑。

依赖安装：
    pip install python-docx lxml
    # 并安装 pandoc: https://pandoc.org/installing.html

用法示例：
    python word_latex_converter.py 论文.docx
    python word_latex_converter.py 论文.docx -o 输出.docx
    python word_latex_converter.py *.docx --outdir ./converted
    python word_latex_converter.py 论文.docx --in-place
"""

import re
import os
import sys
import io
import shutil
import zipfile
import tempfile
import subprocess
import argparse
import logging
from pathlib import Path
from copy import deepcopy
from concurrent.futures import ThreadPoolExecutor, as_completed

from lxml import etree

# ---------------------------------------------------------------------------
# 日志配置
# ---------------------------------------------------------------------------
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
)
logger = logging.getLogger("latex2word")

# ---------------------------------------------------------------------------
# 命名空间
# ---------------------------------------------------------------------------
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
XML_NS = "http://www.w3.org/XML/1998/namespace"

NSMAP = {"w": W_NS, "m": M_NS, "xml": XML_NS}


def wqn(tag):
    """构造 w: 命名空间标签"""
    return f"{{{W_NS}}}{tag}"


def mqn(tag):
    """构造 m: 命名空间标签"""
    return f"{{{M_NS}}}{tag}"


# ---------------------------------------------------------------------------
# 正则：匹配 $$...$$（独立公式）和 $...$（行内公式）
# 先匹配 $$ 再匹配 $；(?<!\\) 排除转义的 \$
# ---------------------------------------------------------------------------
LATEX_PATTERN = re.compile(
    r"(?<!\\)(\$\$(.+?)\$\$|\$(.+?)\$)",
    re.DOTALL,
)

# 公式缓存：(latex_code, display) -> OMML 元素
_omml_cache: dict[tuple[str, bool], etree._Element | None] = {}


# ===========================================================================
# 第一部分：LaTeX -> OMML 转换引擎（基于 pandoc）
# ===========================================================================

def check_pandoc() -> str | None:
    """检测 pandoc 是否可用，返回版本字符串或 None"""
    try:
        result = subprocess.run(
            ["pandoc", "--version"],
            capture_output=True, text=True, timeout=10,
        )
        if result.returncode == 0:
            return result.stdout.splitlines()[0]
    except (FileNotFoundError, subprocess.TimeoutExpired):
        pass
    return None


def _convert_single_pandoc(latex_code: str, display: bool) -> etree._Element | None:
    """
    用 pandoc 将单个 LaTeX 公式转为 OMML 元素。
    display=True  -> 返回 <m:oMathPara>（独立公式段落）
    display=False -> 返回 <m:oMath>（行内公式）
    """
    latex_code = latex_code.strip()
    if not latex_code:
        return None

    # 构造 markdown 输入
    if display:
        md_content = f"$$\n{latex_code}\n$$"
    else:
        md_content = f"${latex_code}$"

    with tempfile.TemporaryDirectory() as tmpdir:
        md_path = os.path.join(tmpdir, "formula.md")
        docx_path = os.path.join(tmpdir, "formula.docx")

        with open(md_path, "w", encoding="utf-8") as f:
            f.write(md_content)

        result = subprocess.run(
            ["pandoc", md_path, "-f", "markdown", "-t", "docx", "-o", docx_path],
            capture_output=True, text=True, timeout=30,
        )
        if result.returncode != 0:
            logger.error(f"pandoc 转换失败 [{latex_code[:40]}]: {result.stderr.strip()}")
            return None

        # 从生成的 docx 中提取 OMML
        try:
            with zipfile.ZipFile(docx_path, "r") as z:
                with z.open("word/document.xml") as f:
                    tree = etree.parse(f)
        except Exception as e:
            logger.error(f"读取 pandoc 输出失败: {e}")
            return None

        root = tree.getroot()

        if display:
            # 独立公式：找 m:oMathPara
            omath_paras = root.findall(f".//{mqn('oMathPara')}")
            if omath_paras:
                return deepcopy(omath_paras[0])
            # fallback：如果只有 oMath，包一层 oMathPara
            omaths = root.findall(f".//{mqn('oMath')}")
            if omaths:
                para = etree.Element(mqn("oMathPara"))
                para.append(deepcopy(omaths[0]))
                return para
        else:
            # 行内公式：找不在 oMathPara 内的 oMath
            omaths = root.xpath(
                f".//m:oMath[not(parent::m:oMathPara)]",
                namespaces={"m": M_NS},
            )
            if omaths:
                return deepcopy(omaths[0])
            # fallback
            all_omaths = root.findall(f".//{mqn('oMath')}")
            if all_omaths:
                return deepcopy(all_omaths[0])

        logger.error(f"未找到 OMML 元素 [{latex_code[:40]}]")
        return None


def get_omml(latex_code: str, display: bool) -> etree._Element | None:
    """获取公式的 OMML（带缓存）"""
    key = (latex_code.strip(), display)
    if key in _omml_cache:
        cached = _omml_cache[key]
        return deepcopy(cached) if cached is not None else None

    result = _convert_single_pandoc(latex_code, display)
    _omml_cache[key] = deepcopy(result) if result is not None else None
    return result


def batch_convert_formulas(formulas: list[tuple[str, bool]], max_workers: int = 4) -> None:
    """
    批量并行转换公式，结果写入全局缓存。
    formulas: [(latex_code, display), ...]
    """
    # 去重
    unique = list(set((lc.strip(), d) for lc, d in formulas if lc.strip()))
    if not unique:
        return

    # 过滤已缓存的
    todo = [f for f in unique if f not in _omml_cache]
    if not todo:
        return

    logger.info(f"正在转换 {len(todo)} 个唯一公式（并行 {max_workers} 线程）...")

    success = 0
    fail = 0
    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        future_map = {
            executor.submit(_convert_single_pandoc, lc, d): (lc, d)
            for lc, d in todo
        }
        for future in as_completed(future_map):
            lc, d = future_map[future]
            try:
                result = future.result()
                _omml_cache[(lc, d)] = deepcopy(result) if result is not None else None
                if result is not None:
                    success += 1
                else:
                    fail += 1
            except Exception as e:
                _omml_cache[(lc, d)] = None
                fail += 1
                logger.error(f"公式转换异常 [{lc[:40]}]: {e}")

    logger.info(f"公式转换完成：成功 {success}，失败 {fail}")


# ===========================================================================
# 第二部分：文档段落处理
# ===========================================================================

def get_paragraph_text(p_elem: etree._Element) -> str:
    """获取段落的完整文本（拼接所有 w:t）"""
    texts = []
    for t in p_elem.iter(wqn("t")):
        if t.text:
            texts.append(t.text)
    return "".join(texts)


def get_first_run_rpr(p_elem: etree._Element) -> etree._Element | None:
    """获取段落中第一个 run 的格式属性（用于保留字体/字号等）"""
    for r in p_elem.iter(wqn("r")):
        rPr = r.find(wqn("rPr"))
        if rPr is not None:
            return deepcopy(rPr)
    return None


def has_complex_elements(p_elem: etree._Element) -> bool:
    """
    检查段落是否包含超链接、域代码、批注等复杂元素。
    包含这些元素时，重建段落可能丢失格式，选择跳过并警告。
    """
    complex_tags = [
        wqn("hyperlink"),
        wqn("fldChar"),
        wqn("instrText"),
        wqn("commentReference"),
        wqn("footnoteReference"),
        wqn("endnoteReference"),
    ]
    for tag in complex_tags:
        if p_elem.find(f".//{tag}") is not None:
            return True
    return False


def add_text_run(
    parent: etree._Element,
    text: str,
    rPr: etree._Element | None = None,
) -> etree._Element:
    """在父元素中添加一个文本 run，保留空格"""
    r = etree.SubElement(parent, wqn("r"))
    if rPr is not None:
        r.append(deepcopy(rPr))
    t = etree.SubElement(r, wqn("t"))
    t.text = text
    t.set(f"{{{XML_NS}}}space", "preserve")
    return r


def set_paragraph_alignment(p_elem: etree._Element, align: str = "center") -> None:
    """设置段落对齐方式"""
    pPr = p_elem.find(wqn("pPr"))
    if pPr is None:
        pPr = etree.Element(wqn("pPr"))
        p_elem.insert(0, pPr)
    jc = pPr.find(wqn("jc"))
    if jc is None:
        jc = etree.SubElement(pPr, wqn("jc"))
    jc.set(wqn("val"), align)


def process_paragraph_element(p_elem: etree._Element) -> int:
    """
    处理单个段落 XML 元素中的 LaTeX 公式。
    返回成功转换的公式数量。
    """
    full_text = get_paragraph_text(p_elem)
    if not full_text or "$" not in full_text:
        return 0

    matches = list(LATEX_PATTERN.finditer(full_text))
    if not matches:
        return 0

    # 包含复杂元素时跳过，避免丢失超链接/域等格式
    if has_complex_elements(p_elem):
        logger.warning(
            f"段落含超链接/域等复杂元素，跳过以保留格式: "
            f"\"{full_text[:60]}{'...' if len(full_text) > 60 else ''}\""
        )
        return 0

    # 保存第一个 run 的格式（尽量保留字体字号）
    first_rPr = get_first_run_rpr(p_elem)

    # 清除段落中除 w:pPr 外的所有子元素
    for child in list(p_elem):
        if child.tag != wqn("pPr"):
            p_elem.remove(child)

    last_end = 0
    converted = 0
    last_display = False

    for match in matches:
        start, end = match.span()

        # 公式前的文本
        if start > last_end:
            text = full_text[last_end:start]
            if text:
                add_text_run(p_elem, text, first_rPr)

        # 判断是 $$...$$ 还是 $...$
        if match.group(2) is not None:
            latex_code = match.group(2)
            display = True
        else:
            latex_code = match.group(3)
            display = False
        last_display = display

        # 获取 OMML 并插入
        omml = get_omml(latex_code, display)
        if omml is not None:
            p_elem.append(deepcopy(omml))
            converted += 1
        else:
            # 转换失败，保留原文
            add_text_run(p_elem, match.group(0), first_rPr)
            logger.warning(f"公式转换失败，保留原文: ${latex_code[:50]}$")

        last_end = end

    # 公式后的文本
    if last_end < len(full_text):
        text = full_text[last_end:]
        if text:
            add_text_run(p_elem, text, first_rPr)

    # 如果独立公式独占一段，自动居中
    if (
        converted == 1
        and last_display
        and full_text.strip() == matches[0].group(0)
    ):
        set_paragraph_alignment(p_elem, "center")

    return converted


# ===========================================================================
# 第三部分：文档级处理
# ===========================================================================

def collect_formulas_from_paragraph(p_elem: etree._Element) -> list[tuple[str, bool]]:
    """从段落中收集所有需要转换的公式"""
    full_text = get_paragraph_text(p_elem)
    if not full_text or "$" not in full_text:
        return []

    formulas = []
    for match in LATEX_PATTERN.finditer(full_text):
        if match.group(2) is not None:
            formulas.append((match.group(2), True))
        else:
            formulas.append((match.group(3), False))
    return formulas


def iter_all_paragraph_elements(doc) -> list[etree._Element]:
    """
    获取文档中所有段落元素，包括：
    - 正文（含嵌套在表格、内容控件 SDT 中的段落）
    - 页眉 / 页脚
    """
    paragraphs = []

    # 正文所有 w:p（递归遍历，覆盖表格、SDT 等嵌套结构）
    for p in doc.element.body.iter(wqn("p")):
        paragraphs.append(p)

    # 页眉页脚
    for section in doc.sections:
        for attr_name in [
            "header", "footer",
            "first_page_header", "first_page_footer",
            "even_page_header", "even_page_footer",
        ]:
            try:
                part = getattr(section, attr_name)
                if part is not None and not part.is_linked_to_previous:
                    for p in part._element.iter(wqn("p")):
                        paragraphs.append(p)
            except (AttributeError, KeyError, TypeError):
                pass

    return paragraphs


def process_document(input_path: str, output_path: str, in_place: bool = False) -> bool:
    """
    处理单个文档。
    返回是否成功。
    """
    logger.info(f"正在处理: {input_path}")

    # 备份原文件
    if in_place:
        backup_path = input_path + ".bak"
        shutil.copy2(input_path, backup_path)
        logger.info(f"已备份原文件: {backup_path}")

    try:
        from docx import Document
    except ImportError:
        logger.error("未安装 python-docx，请运行: pip install python-docx")
        return False

    try:
        doc = Document(input_path)
    except Exception as e:
        logger.error(f"无法打开文档: {e}")
        return False

    # 第一步：收集所有公式
    all_paragraphs = iter_all_paragraph_elements(doc)
    all_formulas = []
    for p in all_paragraphs:
        all_formulas.extend(collect_formulas_from_paragraph(p))

    if not all_formulas:
        logger.info("未找到任何 LaTeX 公式，无需转换。")
        if not in_place and input_path != output_path:
            shutil.copy2(input_path, output_path)
        return True

    logger.info(f"共发现 {len(all_formulas)} 个公式（{len(set((f[0].strip(), f[1]) for f in all_formulas))} 个唯一）")

    # 第二步：批量转换
    batch_convert_formulas(all_formulas)

    # 第三步：替换文档中的公式
    total_converted = 0
    for p in all_paragraphs:
        total_converted += process_paragraph_element(p)

    logger.info(f"文档中成功转换 {total_converted} 个公式")

    # 保存
    try:
        doc.save(output_path)
        logger.info(f"已保存: {output_path}")
    except Exception as e:
        logger.error(f"保存文档失败: {e}")
        return False

    return True


# ===========================================================================
# 第四部分：命令行入口
# ===========================================================================

def parse_args():
    parser = argparse.ArgumentParser(
        description="将 Word 文档中的 $...$ / $$...$$ LaTeX 公式转换为 Word 原生公式",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例:
  %(prog)s thesis.docx                         # 生成 thesis_converted.docx
  %(prog)s thesis.docx -o out.docx             # 指定输出文件名
  %(prog)s *.docx --outdir ./converted         # 批量转换到指定目录
  %(prog)s thesis.docx --in-place              # 直接修改原文件（自动备份）
  %(prog)s thesis.docx --in-place --no-backup  # 直接修改且不备份
        """,
    )
    parser.add_argument("inputs", nargs="+", help="输入的 .docx 文件（支持多个）")
    parser.add_argument("-o", "--output", help="输出文件名（仅单文件时有效）")
    parser.add_argument("--outdir", help="输出目录（批量文件时使用）")
    parser.add_argument("--in-place", action="store_true", help="直接修改原文件")
    parser.add_argument("--no-backup", action="store_true", help="配合 --in-place 使用，不创建备份")
    parser.add_argument("--workers", type=int, default=4, help="并行转换线程数（默认 4）")
    return parser.parse_args()


def main():
    args = parse_args()

    # 检测 pandoc
    pandoc_ver = check_pandoc()
    if not pandoc_ver:
        logger.error("未检测到 pandoc，请先安装: https://pandoc.org/installing.html")
        sys.exit(1)
    logger.info(f"检测到 {pandoc_ver}")

    # 验证输入文件
    input_files = []
    for f in args.inputs:
        path = Path(f)
        if not path.exists():
            logger.error(f"文件不存在: {f}")
            continue
        if path.suffix.lower() != ".docx":
            logger.warning(f"非 .docx 文件，跳过: {f}")
            continue
        input_files.append(str(path))

    if not input_files:
        logger.error("没有可处理的 .docx 文件")
        sys.exit(1)

    # 确定输出路径
    outputs = []
    if len(input_files) == 1 and args.output:
        outputs.append(args.output)
    elif args.outdir:
        outdir = Path(args.outdir)
        outdir.mkdir(parents=True, exist_ok=True)
        for f in input_files:
            outputs.append(str(outdir / Path(f).name))
    else:
        for f in input_files:
            p = Path(f)
            outputs.append(str(p.with_name(p.stem + "_converted" + p.suffix)))

    if len(input_files) != len(outputs):
        logger.error("输入输出数量不匹配")
        sys.exit(1)

    # 处理每个文件
    success_count = 0
    for input_path, output_path in zip(input_files, outputs):
        in_place = args.in_place or (input_path == output_path)
        if in_place and args.no_backup:
            # --no-backup 时不备份：直接输出到原文件，不创建 .bak
            # 这里通过临时文件实现：先输出到临时文件，成功后替换
            tmp_output = input_path + ".tmp_convert"
            ok = process_document(input_path, tmp_output, in_place=False)
            if ok:
                os.replace(tmp_output, input_path)
                success_count += 1
                logger.info(f"已原地更新（无备份）: {input_path}")
            else:
                if os.path.exists(tmp_output):
                    os.remove(tmp_output)
        else:
            if process_document(input_path, output_path, in_place=in_place):
                success_count += 1

    logger.info(f"全部完成：成功 {success_count}/{len(input_files)} 个文件")


if __name__ == "__main__":
    main()
