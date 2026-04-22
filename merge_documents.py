# -*- coding: utf-8 -*-
"""
债券尽调说明文件自动化处理工具V1.1
功能：
1. 自动从盖章页合集PDF每页提取文字
   - 数字PDF：直接提取文字
   - 扫描件：自动OCR识别（使用 rapidocr，无需安装 Tesseract，首次运行自动安装）
2. 与Word文件名模糊匹配，自动生成 matches.json
3. 将Word文件转换为PDF（后台静默，不弹Word窗口）
4. 可选：删除待处理文件转换后PDF的最后一页（用于去除旧盖章页）
5. 按 matches.json 精确合并说明文件PDF + 对应盖章页
6. 可选：删除拼接完成后PDF的倒数第二页
7. 处理完毕后自动清理临时文件

使用方式：
  将待处理的Word文件放入「待处理文件」文件夹，将盖章页合集PDF放入「盖章页」文件夹，输出文件将保存在「输出文件夹」中。
  打开 merge_documents.py，点击运行即可。
"""

import json
import os
import re
import shutil
import subprocess
import sys
import time
from difflib import SequenceMatcher
from pathlib import Path

try:
    from pypdf import PdfReader, PdfWriter
except ImportError:
    from PyPDF2 import PdfReader, PdfWriter


# ─────────────────────────────────────────────
#  自动安装缺失的依赖
# ─────────────────────────────────────────────

def _pip_install(*packages):
    """静默安装pip包"""
    print(f"  正在安装依赖: {' '.join(packages)} ...")
    subprocess.run(
        [sys.executable, "-m", "pip", "install", "--quiet", *packages],
        check=True,
        capture_output=True,
    )
    print(f"  安装完成: {' '.join(packages)}")


def ensure_deps():
    """
    确保 OCR 所需依赖已安装：
      - pymupdf  (PDF → 图片渲染)
      - rapidocr-onnxruntime  (中文OCR，无需Tesseract二进制)
    """
    missing = []
    try:
        import fitz  # noqa: F401
    except ImportError:
        missing.append("pymupdf")

    try:
        from rapidocr_onnxruntime import RapidOCR  # noqa: F401
    except ImportError:
        missing.append("rapidocr-onnxruntime")

    if missing:
        print("\n  [首次运行] 检测到缺少OCR依赖，正在自动安装...")
        _pip_install(*missing)
        print("  OCR依赖安装完成，继续处理...\n")


# ─────────────────────────────────────────────
#  文本工具
# ─────────────────────────────────────────────

def normalize_text(text):
    """标准化文本：去除空白、标点差异，便于相似度比较"""
    text = re.sub(r"[\s\u3000\n\r\t]+", "", text)
    text = text.replace("（", "(").replace("）", ")")
    text = text.replace("，", ",").replace("、", "").replace("。", "")
    text = text.replace("：", ":").replace("【", "").replace("】", "")
    return text.lower()


def similarity(a, b):
    """计算两个字符串的相似度（0~1）"""
    a_n = normalize_text(a)
    b_n = normalize_text(b)
    if not a_n or not b_n:
        return 0.0
    return SequenceMatcher(None, a_n, b_n).ratio()


def get_keyword_from_filename(filename):
    """
    从文件名提取匹配用关键词：
    去掉扩展名 + 去掉开头的编号前缀（如 '1-1-1-4 '、'1-2-3-10-1 '）
    """
    name = os.path.splitext(filename)[0]
    name = re.sub(r"^[\d\-]+\s*", "", name)
    return name


# ─────────────────────────────────────────────
#  PDF 文字提取（数字PDF + 扫描件OCR）
# ─────────────────────────────────────────────

def extract_text_digital(page):
    """使用 pypdf 直接提取页面文字（仅适用于数字PDF）"""
    try:
        return (page.extract_text() or "").strip()
    except Exception:
        return ""


def ocr_pdf_page(pdf_path, page_index, dpi=150):
    """
    使用 rapidocr 对 PDF 单页进行 OCR
    page_index: 从0开始的页码
    返回识别到的文字（拼接所有行），失败返回空字符串
    """
    try:
        import fitz
        import numpy as np
        from rapidocr_onnxruntime import RapidOCR

        ocr_engine = RapidOCR()

        # 用 pymupdf 将 PDF 页渲染为图片（numpy array）
        doc = fitz.open(pdf_path)
        page = doc[page_index]
        mat = fitz.Matrix(dpi / 72, dpi / 72)
        pix = page.get_pixmap(matrix=mat, colorspace=fitz.csRGB)
        doc.close()

        # 转为 numpy array
        img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(
            pix.height, pix.width, pix.n
        )
        if pix.n == 4:  # RGBA → RGB
            img = img[:, :, :3]

        # OCR
        result, _ = ocr_engine(img)
        if result:
            lines = [item[1] for item in result if item and len(item) > 1]
            return "\n".join(lines)
        return ""

    except Exception:
        return ""


def extract_all_page_texts(stamp_pdf_path, temp_folder):
    """
    提取盖章页PDF所有页面的文字：
    - 先用 pypdf 直接提取
    - 若某页文字 < 10 字符（扫描件），自动用 OCR 识别
    返回 {页码(1-based): 文字} 字典
    """
    _ = temp_folder
    reader = PdfReader(stamp_pdf_path)
    total_pages = len(reader.pages)
    page_texts = {}

    digital_count = 0
    ocr_count = 0

    for i in range(total_pages):
        text = extract_text_digital(reader.pages[i])
        if len(text) >= 10:
            digital_count += 1
        else:
            # 文字过少 → 扫描件，使用OCR
            ocr_text = ocr_pdf_page(stamp_pdf_path, i)
            if ocr_text.strip():
                text = ocr_text
                ocr_count += 1
        page_texts[i + 1] = text

    status_parts = []
    if digital_count:
        status_parts.append(f"{digital_count} 页数字文字提取")
    if ocr_count:
        status_parts.append(f"{ocr_count} 页OCR识别")
    if status_parts:
        print(f"  文字提取完成：{' / '.join(status_parts)}")

    no_text = sum(1 for t in page_texts.values() if len(t) < 10)
    if no_text:
        print(f"  [注意] 仍有 {no_text} 页未能提取到足够文字，这些页将使用顺序兜底匹配")

    return page_texts


# ─────────────────────────────────────────────
#  自动生成 matches.json
# ─────────────────────────────────────────────

def build_matches(stamp_pdf_path, word_files, temp_folder, matches_save_path):
    """
    从盖章页PDF提取文字，与Word文件名做模糊匹配，
    生成并保存 matches.json

    匹配策略（贪心，最大化总相似度）：
    1. 提取每页文字（含OCR）
    2. 计算所有 (Word文件, 盖章页) 的相似度
    3. 从高到低贪心分配，已占用的不重复使用
    4. 相似度极低的未匹配文件用剩余页顺序兜底
    """
    print("\n【步骤2】自动匹配盖章页...")
    print("  正在提取盖章PDF各页文字（扫描件自动OCR）...")
    page_texts = extract_all_page_texts(stamp_pdf_path, temp_folder)
    total_pages = len(page_texts)

    pages_with_text = sum(1 for t in page_texts.values() if len(t) >= 10)
    use_text_match = pages_with_text >= max(1, total_pages * 0.2)

    matches = {}

    if use_text_match:
        print(f"  {pages_with_text}/{total_pages} 页有可用文字，使用文字相似度匹配")

        # 构建全量相似度矩阵
        all_pairs = []
        for wf in word_files:
            kw = get_keyword_from_filename(wf.name)
            for page_num, text in page_texts.items():
                score = similarity(kw, text)
                all_pairs.append((score, wf.name, page_num))
        all_pairs.sort(reverse=True)

        assigned_words = set()
        used_pages = set()
        for score, word_name, page_num in all_pairs:
            if word_name in assigned_words or page_num in used_pages:
                continue
            if score < 0.04:
                break
            matches[word_name] = page_num
            assigned_words.add(word_name)
            used_pages.add(page_num)

        # 未匹配的顺序兜底
        unmatched = [f for f in word_files if f.name not in matches]
        remaining = sorted(p for p in page_texts if p not in used_pages)
        for idx, wf in enumerate(unmatched):
            if idx < len(remaining):
                matches[wf.name] = remaining[idx]
                print(f"  [顺序兜底] 第 {remaining[idx]:>2} 页 ← {wf.name[:45]}...")

    else:
        print("  文字提取不足，全部使用顺序匹配（Word文件按字母排序对应盖章第1~N页）")
        for idx, wf in enumerate(word_files):
            if idx + 1 <= total_pages:
                matches[wf.name] = idx + 1

    # 保存 matches.json
    with open(matches_save_path, "w", encoding="utf-8") as f:
        json.dump(matches, f, ensure_ascii=False, indent=2)

    # 打印匹配报告
    print("\n  ── 匹配结果预览（已保存至 matches.json，如有误可手动修改后重跑）──")
    for wf in word_files:
        pg = matches.get(wf.name, "未匹配")
        label = (wf.name[:55] + "...") if len(wf.name) > 55 else wf.name
        print(f"  [第{str(pg):>3}页] {label}")

    return matches


# ─────────────────────────────────────────────
#  拆分盖章页PDF
# ─────────────────────────────────────────────

def split_all_stamp_pages(stamp_pdf_path, output_folder):
    """将盖章页合集PDF按页拆分，返回 {页码(1-based): 路径}"""
    print("\n【步骤3】拆分盖章页PDF...")
    reader = PdfReader(stamp_pdf_path)
    split_pages = {}
    for i, page in enumerate(reader.pages):
        out = os.path.join(output_folder, f"stamp_page_{i + 1:03d}.pdf")
        w = PdfWriter()
        w.add_page(page)
        with open(out, "wb") as f:
            w.write(f)
        split_pages[i + 1] = out
    print(f"  共拆分 {len(split_pages)} 页")
    return split_pages


# ─────────────────────────────────────────────
#  Word → PDF 转换
# ─────────────────────────────────────────────

def kill_word_processes():
    """关闭所有残留的Word进程"""
    try:
        subprocess.run(
            "taskkill /F /IM WINWORD.EXE", shell=True, capture_output=True, timeout=5
        )
        time.sleep(0.3)
    except Exception:
        pass


def convert_word_to_pdf(word_path, output_dir, max_retries=3):
    """
    Word转PDF：优先docx2pdf，失败则用win32com后台静默模式
    遇到COM错误自动重试，每次重试前清理残留进程
    """
    os.makedirs(output_dir, exist_ok=True)
    word_path = os.path.abspath(word_path)
    base_name = os.path.splitext(os.path.basename(word_path))[0]
    pdf_path = os.path.join(output_dir, base_name + ".pdf")

    # 每次转换前先清理残留Word进程
    kill_word_processes()

    # 方法1：docx2pdf
    for attempt in range(max_retries):
        try:
            from docx2pdf import convert

            convert(word_path, pdf_path)
            if os.path.exists(pdf_path) and os.path.getsize(pdf_path) > 1024:
                return pdf_path
        except ImportError:
            break
        except Exception as e:
            error_msg = str(e)
            if "RPC_E_CALL_REJECTED" in error_msg or "拒绝接收呼叫" in error_msg:
                print(f"    docx2pdf 重试 ({attempt + 1}/{max_retries})...")
                kill_word_processes()
                time.sleep(1)
                continue
            print(f"    docx2pdf 失败: {e}")
            break

    # 方法2：win32com（Visible=False，不弹窗）
    for attempt in range(max_retries):
        kill_word_processes()
        time.sleep(0.5)

        try:
            import pythoncom

            pythoncom.CoUninitialize()
            pythoncom.CoInitialize()
        except Exception:
            pass

        word_app = None
        try:
            import win32com.client as win32

            word_app = win32.DispatchEx("Word.Application")
            word_app.Visible = False
            word_app.DisplayAlerts = False
            doc = word_app.Documents.Open(
                word_path, ReadOnly=True, ConfirmConversions=False
            )
            doc.SaveAs(pdf_path, FileFormat=17)
            doc.Close(False)
            word_app.Quit()
            word_app = None

            if os.path.exists(pdf_path) and os.path.getsize(pdf_path) > 1024:
                return pdf_path

        except Exception as e:
            error_msg = str(e)
            if (
                "RPC_E_CALL_REJECTED" in error_msg
                or "拒绝接收呼叫" in error_msg
                or "call was rejected" in error_msg.lower()
            ):
                print(f"    win32com 重试 ({attempt + 1}/{max_retries})...")
                if word_app:
                    try:
                        word_app.Quit()
                    except Exception:
                        pass
                kill_word_processes()
                time.sleep(1.5)
                continue
            print(f"    win32com 失败: {e}")
            if word_app:
                try:
                    word_app.Quit()
                except Exception:
                    pass
            break

    return None


def remove_last_page_from_pdf(pdf_path):
    """删除 PDF 最后一页（若仅1页则不删除），返回 (是否成功删除, 总页数)。"""
    try:
        with open(pdf_path, "rb") as f:
            reader = PdfReader(f)
            total = len(reader.pages)

        if total <= 1:
            return False, total

        writer = PdfWriter()
        with open(pdf_path, "rb") as f:
            reader = PdfReader(f)
            for i in range(total - 1):
                writer.add_page(reader.pages[i])

        with open(pdf_path, "wb") as f:
            writer.write(f)
        return True, total
    except Exception:
        return False, 0


def remove_penultimate_page_from_pdf(pdf_path):
    """删除 PDF 倒数第二页（页数不足2页则不删除），返回 (是否成功删除, 总页数)。"""
    try:
        with open(pdf_path, "rb") as f:
            reader = PdfReader(f)
            total = len(reader.pages)

        if total <= 1:
            return False, total

        target_index = total - 2  # 倒数第二页（0-based）
        writer = PdfWriter()
        with open(pdf_path, "rb") as f:
            reader = PdfReader(f)
            for i, page in enumerate(reader.pages):
                if i != target_index:
                    writer.add_page(page)

        with open(pdf_path, "wb") as f:
            writer.write(f)
        return True, total
    except Exception:
        return False, 0


# ─────────────────────────────────────────────
#  合并PDF
# ─────────────────────────────────────────────

def merge_pdfs(pdf1_path, pdf2_path, output_path):
    writer = PdfWriter()
    for p in [pdf1_path, pdf2_path]:
        with open(p, "rb") as f:
            for page in PdfReader(f).pages:
                writer.add_page(page)
    with open(output_path, "wb") as f:
        writer.write(f)


# ─────────────────────────────────────────────
#  清理临时文件
# ─────────────────────────────────────────────

def cleanup_temp(base_dir):
    for td in ["temp", "temp_images", "extract", "temppdf", "tempimages"]:
        p = os.path.join(base_dir, td)
        if os.path.exists(p):
            try:
                shutil.rmtree(p)
                print(f"  已清理: {td}")
            except Exception as e:
                print(f"  清理 {td} 失败: {e}")


# ─────────────────────────────────────────────
#  获取Word文件列表
# ─────────────────────────────────────────────

def get_word_files(folder_path):
    word_files = []
    for ext in [".docx", ".doc"]:
        files = list(Path(folder_path).glob(f"*{ext}"))
        files = [f for f in files if not f.name.startswith("~$")]
        word_files.extend(files)
    return sorted(word_files, key=lambda f: f.name)


def ask_remove_last_page():
    """询问用户是否删除待处理文件转换后PDF的最后一页。"""
    prompt = (
        "\n是否删除每个待处理文件的最后一页？\n"
        "（适用于文件末页已包含旧盖章页；输入 y 或 n，默认 n）："
    )
    while True:
        choice = input(prompt).strip().lower()
        if choice in {"", "n", "no", "否"}:
            return False
        if choice in {"y", "yes", "是"}:
            return True
        print("请输入 y 或 n。")


def ask_remove_penultimate_after_merge():
    """询问用户是否删除拼接完成后PDF的倒数第二页。"""
    prompt = (
        "\n是否删除拼接完成后每个输出PDF的倒数第二页？\n"
        "（输入 y 或 n，默认 n）："
    )
    while True:
        choice = input(prompt).strip().lower()
        if choice in {"", "n", "no", "否"}:
            return False
        if choice in {"y", "yes", "是"}:
            return True
        print("请输入 y 或 n。")


# ─────────────────────────────────────────────
#  主流程
# ─────────────────────────────────────────────

def main():
    base_dir = os.path.dirname(os.path.abspath(__file__))
    if "BASE_DIR" in os.environ:
        base_dir = os.environ["BASE_DIR"]

    # ========== 文件夹配置（可自行修改）==========
    word_folder = os.path.join(base_dir, "待处理文件")  # 放入Word文件
    stamp_folder = os.path.join(base_dir, "盖章页")  # 放入盖章页合集PDF
    output_folder = os.path.join(base_dir, "输出文件夹")  # 输出合并后的PDF
    # ===========================================

    temp_folder = os.path.join(base_dir, "temp")
    matches_path = os.path.join(base_dir, "matches.json")

    print("=" * 60)
    print("债券尽调说明文件自动化处理工具")
    print("=" * 60)

    remove_last_page = ask_remove_last_page()
    remove_penultimate_after_merge = ask_remove_penultimate_after_merge()
    if remove_last_page:
        print("  已启用：将删除每个待处理文件转换后PDF的最后一页")
    else:
        print("  未启用：保留待处理文件全部页面")
    if remove_penultimate_after_merge:
        print("  已启用：将删除每个输出PDF（拼接后）的倒数第二页")
    else:
        print("  未启用：保留拼接完成后的全部页面")

    # 确保OCR依赖已安装（首次运行自动安装）
    ensure_deps()

    # 查找盖章页PDF
    stamp_pdf = None
    for f in os.listdir(stamp_folder):
        if f.lower().endswith(".pdf"):
            stamp_pdf = os.path.join(stamp_folder, f)
            break
    if not stamp_pdf:
        print(f"[错误] 未找到盖章页PDF！请将文件放入:\n  {stamp_folder}")
        input("\n按回车键退出...")
        sys.exit(1)

    print(f"盖章页PDF:  {os.path.basename(stamp_pdf)}")
    print(f"Word目录:   {word_folder}")
    print(f"输出目录:   {output_folder}")
    print("=" * 60)

    os.makedirs(output_folder, exist_ok=True)
    os.makedirs(temp_folder, exist_ok=True)

    # 步骤1：扫描Word文件
    print("\n【步骤1】扫描说明文件Word文件...")
    word_files = get_word_files(word_folder)
    if not word_files:
        print("[错误] 「说明文件」目录中未找到Word文件！")
        input("\n按回车键退出...")
        sys.exit(1)
    print(f"  找到 {len(word_files)} 个Word文件")

    # 步骤2：自动匹配（每次运行重新生成matches.json）
    matches = build_matches(stamp_pdf, word_files, temp_folder, matches_path)

    # 步骤3：拆分盖章页
    split_pages = split_all_stamp_pages(stamp_pdf, temp_folder)

    # 步骤4：逐文件转换并合并
    print("\n【步骤4】转换Word并合并盖章页...")
    success_count = 0
    failed_count = 0

    for i, word_file in enumerate(word_files):
        print(f"\n  [{i + 1}/{len(word_files)}] {word_file.name}")
        print("    转换Word → PDF...")

        pdf_path = convert_word_to_pdf(str(word_file), temp_folder)
        if not pdf_path:
            print("    ✗ 转换失败，跳过")
            failed_count += 1
            continue

        if remove_last_page:
            removed, total_pages = remove_last_page_from_pdf(pdf_path)
            if removed:
                print(f"    ✓ 已删除最后一页（原共 {total_pages} 页）")
            elif total_pages == 1:
                print("    [提示] 文件仅1页，未删除最后一页")
            else:
                print("    [提示] 删除最后一页失败，继续使用原PDF")

        page_num = matches.get(word_file.name)
        stamp_path = split_pages.get(int(page_num)) if page_num else None
        output_pdf = os.path.join(
            output_folder, os.path.splitext(word_file.name)[0] + ".pdf"
        )

        if stamp_path:
            try:
                merge_pdfs(pdf_path, stamp_path, output_pdf)
                print(f"    ✓ 合并完成（盖章页第 {page_num} 页）")
                if remove_penultimate_after_merge:
                    removed, total_pages = remove_penultimate_page_from_pdf(output_pdf)
                    if removed:
                        print(f"    ✓ 已删除倒数第二页（原共 {total_pages} 页）")
                    elif total_pages <= 1:
                        print("    [提示] 文件页数不足2页，未删除倒数第二页")
                    else:
                        print("    [提示] 删除倒数第二页失败，继续使用当前PDF")
                success_count += 1
            except Exception as e:
                print(f"    ✗ 合并失败: {e}")
                failed_count += 1
        else:
            shutil.copy(pdf_path, output_pdf)
            print("    ✓ 已输出（无对应盖章页，直接保存）")
            if remove_penultimate_after_merge:
                removed, total_pages = remove_penultimate_page_from_pdf(output_pdf)
                if removed:
                    print(f"    ✓ 已删除倒数第二页（原共 {total_pages} 页）")
                elif total_pages <= 1:
                    print("    [提示] 文件页数不足2页，未删除倒数第二页")
                else:
                    print("    [提示] 删除倒数第二页失败，继续使用当前PDF")
            success_count += 1

    # 步骤5：清理临时文件
    print("\n【步骤5】清理临时文件...")
    cleanup_temp(base_dir)

    # 汇总
    print("\n" + "=" * 60)
    print(f"处理完成！成功: {success_count} 个，失败: {failed_count} 个")
    print(f"输出目录: {output_folder}")
    print("=" * 60)
    if failed_count > 0:
        print("\n[提示] 转换失败通常是Word文件被其他程序占用。")
        print("       请关闭所有打开的Word文档后重试。")

    input("\n按回车键退出...")


if __name__ == "__main__":
    main()
