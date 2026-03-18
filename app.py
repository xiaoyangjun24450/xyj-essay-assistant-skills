#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
论文改写助手主程序
从 docx 文件开始，执行预处理、改写和重建流程
"""

import sys
from pathlib import Path

# 添加 scripts 目录到路径
if getattr(sys, 'frozen', False):
    # 打包后的 exe 环境
    SCRIPT_DIR = Path(sys.executable).parent
else:
    # 开发环境
    SCRIPT_DIR = Path(__file__).parent

sys.path.insert(0, str(SCRIPT_DIR / 'scripts'))

from docx_preprocessor import DocxPreprocessor
from chunks_rewriter import run_rewrite
from docx_chunks_restorer import DocxChunksRestorer


def main():
    # 1. 获取用户输入
    template_docx = input("请输入 docx 文件路径: ").strip()

    # 验证文件是否存在
    docx_path = Path(template_docx)
    if not docx_path.exists():
        print(f"错误: 文件不存在: {template_docx}")
        sys.exit(1)

    # 2. 获取改写需求
    requirement = input("请输入改写需求: ").strip()

    # 3. 创建工作目录
    import datetime
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    work_dir = SCRIPT_DIR / "output" / timestamp
    work_dir.mkdir(parents=True, exist_ok=True)

    # 4. 执行预处理
    print("正在处理原始文档...")
    try:
        processor = DocxPreprocessor(template_docx)
        processor.process(str(work_dir))
        print(f"处理完成！输出目录: {work_dir}")
    except Exception as e:
        print(f"预处理失败！错误: {e}")
        sys.exit(1)

    # 5. 执行改写
    chunks_dir = work_dir / "chunks"
    chunks_new_dir = work_dir / "chunks_new"
    print("正在改写文档...")
    try:
        run_rewrite(str(chunks_dir), str(chunks_new_dir), requirement)
        print(f"改写完成！输出目录: {chunks_new_dir}")
    except Exception as e:
        print(f"改写失败！错误: {e}")
        sys.exit(1)

    # 6. 重建文档
    output_docx = work_dir / "output.docx"
    print("正在重建文档...")
    try:
        restorer = DocxChunksRestorer(str(work_dir / "unzipped"), str(chunks_new_dir))
        restorer.restore(str(output_docx))
        print(f"合并完成！输出文件: {output_docx}")
    except Exception as e:
        print(f"重建失败！错误: {e}")
        sys.exit(1)

    print("\n全部完成！")


if __name__ == '__main__':
    main()
