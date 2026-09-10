#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PDF 批量转 Word (docx) 工具

功能：
    把本程序所在文件夹（或拖入的文件夹 / 指定目录）中的所有 PDF
    自动转换为同名的 .docx 文件。

用法：
    1. 双击运行（exe 或 py），程序自动处理自己所在目录下的所有 PDF。
    2. 也可以把 PDF 文件 / 文件夹拖到 exe 上运行。
    3. 命令行： pdf2word.exe [目录或PDF文件...]

作者：WorkBuddy
"""

import contextlib
import io
import logging
import os
import sys
import time
import traceback
import warnings

os.environ.setdefault("PYTHONIOENCODING", "utf-8")

# ---------------------------------------------------------------------------
# 让底层 PDF 库闭嘴，保持进度输出干净
#
# pdf2docx / PyMuPDF 会打印一堆 [INFO] 日志和废弃提示，对使用者没有价值。
# 这里三重保险：
#   1. 提高 logging 级别，吞掉 [INFO]
#   2. warnings.filterwarnings，吞掉 Python 层警告
#   3. _mute_output_fd()，吞掉 C 层直写 stdout 的提示
# ---------------------------------------------------------------------------
logging.basicConfig(level=logging.WARNING)
logging.getLogger().setLevel(logging.WARNING)
for _name in list(logging.root.manager.loggerDict) + [
    "pdf2docx", "fitz", "pymupdf", "PyMuPDF",
]:
    try:
        _lg = logging.getLogger(_name)
        _lg.setLevel(logging.WARNING)
        _lg.propagate = False
    except Exception:
        pass

warnings.filterwarnings("ignore")


BANNER = r"""
==============================================================
                 PDF  ->  Word  批量转换工具
==============================================================
"""


def _quiet_pymupdf():
    """关掉 PyMuPDF 的提示/警告输出（必须在导入 pdf2docx 前调用）。"""
    try:
        import pymupdf
        pymupdf.set_messages(0)
        return
    except Exception:
        pass
    try:
        import fitz
        fitz.set_messages(0)
    except Exception:
        pass


@contextlib.contextmanager
def _mute_output_fd():
    """把操作系统层面的 stdout + stderr 临时整体静音。

    PyMuPDF 的 "fitz API is deprecated" 提示是在 import 时由 C 层直接
    写出来的一句字符串——实测它走的是 stdout 而不是 stderr，
    因此必须在文件描述符层面同时把 1 和 2 都顶住，Python 层的
    sys.stdout / sys.stderr 重定向对它无效。
    """
    saved = []
    devnull = None
    try:
        devnull = os.open(os.devnull, os.O_WRONLY)
        for fd in (1, 2):
            try:
                saved.append((fd, os.dup(fd)))
                os.dup2(devnull, fd)
            except OSError:
                pass
        yield
    finally:
        for fd, backup in saved:
            try:
                os.dup2(backup, fd)
                os.close(backup)
            except OSError:
                pass
        if devnull is not None:
            try:
                os.close(devnull)
            except OSError:
                pass


def get_target_dir():
    """决定要扫描哪个目录。

    优先级：
        1. 命令行参数中传入的文件夹 / PDF 文件
        2. 把文件拖到 exe 图标上时，系统传入的文件路径
        3. 打包成 exe 时：exe 自身所在目录
        4. 直接跑 py 脚本时：脚本所在目录
    """
    # 情况 1 & 2：命令行参数
    args = [a for a in sys.argv[1:] if a.strip()]

    # PyInstaller 打包后，如果用户把文件拖到 exe 上，参数就是被拖的文件
    # 如果没有任何参数，argv[0] 是程序自身路径
    if args:
        # 只取第一个存在的路径作为输入源
        for a in args:
            p = os.path.abspath(a)
            if os.path.isfile(p):
                # 拖进来的是单个 PDF 文件
                return os.path.dirname(p), [p]
            if os.path.isdir(p):
                return p, None

    # 情况 3 & 4：程序自身所在目录
    if getattr(sys, "frozen", False):
        # 打包后的 exe
        base = os.path.dirname(sys.executable)
    else:
        # 直接运行 py 脚本
        base = os.path.dirname(os.path.abspath(__file__))
    return base, None


def unique_path(path):
    """如果目标文件已存在，自动加序号 (1) (2) ... 避免覆盖。"""
    if not os.path.exists(path):
        return path

    root, ext = os.path.splitext(path)
    i = 1
    while True:
        candidate = f"{root}({i}){ext}"
        if not os.path.exists(candidate):
            return candidate
        i += 1


def find_pdfs(directory, explicit_files=None):
    """找出待转换的 PDF 文件列表。"""
    if explicit_files:
        return [f for f in explicit_files if f.lower().endswith(".pdf")]

    pdfs = []
    for name in sorted(os.listdir(directory)):
        full = os.path.join(directory, name)
        # 只处理普通文件，且扩展名为 .pdf（忽略大小写、忽略已经生成的 word）
        if os.path.isfile(full) and name.lower().endswith(".pdf"):
            # 跳过临时文件（如 ~$xxx.pdf、.xxx.pdf）
            if name.startswith("~$") or name.startswith("."):
                continue
            pdfs.append(full)
    return pdfs


def convert_one(pdf_path, out_dir):
    """转换单个 PDF，返回生成的 docx 路径。"""
    # 在静音环境中导入并转换，避免底层库的提示信息插进进度输出里
    with _mute_output_fd():
        _quiet_pymupdf()
        from pdf2docx import Converter

    base = os.path.splitext(os.path.basename(pdf_path))[0]
    target = unique_path(os.path.join(out_dir, base + ".docx"))

    cv = Converter(pdf_path)
    try:
        cv.convert(target)
    finally:
        cv.close()

    return target


def human_size(num):
    for unit in ("B", "KB", "MB", "GB"):
        if num < 1024:
            return f"{num:.1f}{unit}"
        num /= 1024
    return f"{num:.1f}TB"


def pause_if_needed(has_error=False):
    """让窗口停住，等用户按回车，避免一闪而过。"""
    # 只有在"双击运行"这种交互场景才暂停：
    # 如果 stdout 不是终端（被重定向），就不暂停
    try:
        is_tty = sys.stdin is not None and sys.stdin.isatty()
    except Exception:
        is_tty = False

    if not is_tty:
        return

    print()
    if has_error:
        print("提示：部分文件转换失败，请查看上方日志。")
    try:
        input("按回车键退出 ...")
    except (EOFError, KeyboardInterrupt):
        pass


def main():
    # 先用静音环境做一次预热导入：pdf2docx / PyMuPDF 在首次导入时会
    # 往控制台打提示信息，提前在这里消化掉，之后的输出就干净了。
    with _mute_output_fd():
        _quiet_pymupdf()
        try:
            import pdf2docx  # noqa: F401
        except Exception:
            pass

    print(BANNER)

    target_dir, explicit_files = get_target_dir()
    target_dir = os.path.abspath(target_dir)
    print(f"扫描目录：{target_dir}")
    print()

    if not os.path.isdir(target_dir):
        print("[错误] 目录不存在，请检查后重试。")
        pause_if_needed(True)
        return 1

    pdfs = find_pdfs(target_dir, explicit_files)

    if not pdfs:
        print("没有找到任何 PDF 文件。")
        print()
        print("说明：请把本程序放到「存放 PDF 的文件夹」里，")
        print("      或者把 PDF / 文件夹直接拖到本程序图标上再运行。")
        pause_if_needed()
        return 0

    print(f"共找到 {len(pdfs)} 个 PDF 文件，开始转换 ...")
    print("-" * 62)

    ok, failed = 0, 0
    failures = []
    t_all = time.time()

    for idx, pdf in enumerate(pdfs, 1):
        name = os.path.basename(pdf)
        size = human_size(os.path.getsize(pdf))
        print(f"[{idx}/{len(pdfs)}] {name}  ({size})")

        t0 = time.time()
        try:
            out = convert_one(pdf, target_dir)
            cost = time.time() - t0
            print(f"          -> 完成：{os.path.basename(out)}  ({cost:.1f}秒)")
            ok += 1
        except Exception as e:
            print(f"          -> 失败：{e}")
            failures.append((name, repr(e)))
            failed += 1
            # 打印简要堆栈，方便排查
            if os.environ.get("PDF2WORD_DEBUG"):
                traceback.print_exc()
        print()

    total = time.time() - t_all
    print("=" * 62)
    print(f"全部完成：成功 {ok} 个，失败 {failed} 个，耗时 {total:.1f} 秒")
    if failures:
        print()
        print("失败清单：")
        for n, err in failures:
            print(f"  - {n} : {err}")
    print("=" * 62)

    pause_if_needed(failed > 0)
    return 0 if failed == 0 else 2


if __name__ == "__main__":
    try:
        sys.exit(main())
    except KeyboardInterrupt:
        print("\n已取消。")
        sys.exit(130)
