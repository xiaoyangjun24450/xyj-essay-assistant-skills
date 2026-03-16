#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
校验 chunks/ 和 origin_chunks/ 下的同名文件
检查每一行的 [PARA_ID] 是否匹配，是否有遗漏的行
"""

import re
import sys
from pathlib import Path
from typing import List, Tuple, Optional


class ChunksVerifier:
    """Chunks文件校验器"""

    def __init__(self, chunks_dir: str, origin_chunks_dir: str):
        self.chunks_dir = Path(chunks_dir)
        self.origin_chunks_dir = Path(origin_chunks_dir)
        self.errors: List[str] = []
        self.warnings: List[str] = []

    def verify(self) -> bool:
        """执行校验，返回是否通过"""
        print(f"校验目录:")
        print(f"  chunks: {self.chunks_dir}")
        print(f"  origin_chunks: {self.origin_chunks_dir}")
        print()

        if not self.chunks_dir.exists():
            self.errors.append(f"chunks目录不存在: {self.chunks_dir}")
            return False

        if not self.origin_chunks_dir.exists():
            self.errors.append(f"origin_chunks目录不存在: {self.origin_chunks_dir}")
            return False

        # 获取所有chunk文件
        chunk_files = sorted(self.chunks_dir.glob('chunk_*.md'))
        origin_files = sorted(self.origin_chunks_dir.glob('chunk_*.md'))

        chunk_names = {f.name for f in chunk_files}
        origin_names = {f.name for f in origin_files}

        # 检查文件数量是否一致
        if len(chunk_files) != len(origin_files):
            self.warnings.append(f"文件数量不一致: chunks={len(chunk_files)}, origin_chunks={len(origin_files)}")

        # 检查是否有缺失的文件
        only_in_chunks = chunk_names - origin_names
        only_in_origin = origin_names - chunk_names

        if only_in_chunks:
            self.errors.append(f"仅在chunks中存在: {', '.join(sorted(only_in_chunks))}")
        if only_in_origin:
            self.errors.append(f"仅在origin_chunks中存在: {', '.join(sorted(only_in_origin))}")

        # 校验同名文件
        common_files = chunk_names & origin_names
        if not common_files:
            self.errors.append("没有共同的文件可以校验")
            return False

        print(f"找到 {len(common_files)} 个共同文件")
        print()

        all_passed = True
        for filename in sorted(common_files):
            chunk_file = self.chunks_dir / filename
            origin_file = self.origin_chunks_dir / filename
            passed = self._verify_file_pair(chunk_file, origin_file)
            if not passed:
                all_passed = False

        return all_passed

    def _verify_file_pair(self, chunk_file: Path, origin_file: Path) -> bool:
        """校验一对文件，返回是否通过"""
        print(f"校验: {chunk_file.name}")

        chunk_lines = self._read_file_lines(chunk_file)
        origin_lines = self._read_file_lines(origin_file)

        # 检查行数
        if len(chunk_lines) != len(origin_lines):
            self.errors.append(
                f"{chunk_file.name}: 行数不一致 "
                f"(chunks={len(chunk_lines)}, origin={len(origin_lines)})"
            )

        # 逐行校验PARA_ID
        passed = True
        max_lines = max(len(chunk_lines), len(origin_lines))

        for i in range(max_lines):
            chunk_line = chunk_lines[i] if i < len(chunk_lines) else None
            origin_line = origin_lines[i] if i < len(origin_lines) else None

            if chunk_line is None:
                self.errors.append(f"{chunk_file.name}: 第{i+1}行 - chunks缺少该行")
                passed = False
                continue

            if origin_line is None:
                self.errors.append(f"{chunk_file.name}: 第{i+1}行 - origin缺少该行")
                passed = False
                continue

            chunk_para_id = self._extract_para_id(chunk_line)
            origin_para_id = self._extract_para_id(origin_line)

            if chunk_para_id is None:
                self.errors.append(f"{chunk_file.name}: 第{i+1}行 - chunks未找到PARA_ID")
                passed = False
                continue

            if origin_para_id is None:
                self.errors.append(f"{chunk_file.name}: 第{i+1}行 - origin未找到PARA_ID")
                passed = False
                continue

            if chunk_para_id != origin_para_id:
                self.errors.append(
                    f"{chunk_file.name}: 第{i+1}行 - PARA_ID不匹配 "
                    f"(chunks={chunk_para_id}, origin={origin_para_id})"
                )
                passed = False
            
            # 校验格式标签（只检查 chunks 文件）
            if not self._verify_format_tags(chunk_line, chunk_file.name, i + 1):
                passed = False
            
            # 比较 origin 和 chunks 的标签差异
            if not self._compare_format_tags(origin_line, chunk_line, chunk_file.name, i + 1):
                passed = False

        if passed:
            print(f"  ✓ 通过 ({len(chunk_lines)}行)")
        else:
            print(f"  ✗ 失败")

        return passed

    def _read_file_lines(self, file_path: Path) -> List[str]:
        """读取文件所有行"""
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                return [line.rstrip('\n') for line in f.readlines()]
        except Exception as e:
            self.errors.append(f"读取文件失败 {file_path}: {e}")
            return []

    def _extract_para_id(self, line: str) -> Optional[str]:
        """从行中提取PARA_ID，格式为 [XXXXXXXX]"""
        match = re.match(r'^\[([0-9A-Fa-f]{8})\]', line)
        if match:
            return match.group(1).upper()
        return None

    def _extract_tag_types(self, line: str) -> set:
        """
        从行中提取所有格式标签的种类（不重复）
        返回: {属性类型集合}
        例如: <i=true> 提取出 {'i=true'}
              <f=黑体,b=true> 提取出 {'f=黑体', 'b=true'}
        """
        tag_types = set()
        # 匹配完整的标签对: <attrs>content</attrs>
        tag_pair_pattern = r'<([fisbzchuv,=\w]+)>([^<]*)</\1>'
        
        for match in re.finditer(tag_pair_pattern, line):
            attrs = match.group(1)
            # 将属性拆分为单个类型
            for attr in attrs.split(','):
                if '=' in attr:
                    tag_types.add(attr.strip())
        
        return tag_types
    
    def _compare_format_tags(self, origin_line: str, chunk_line: str, filename: str, line_num: int) -> bool:
        """
        比较 origin 和 chunk 行的标签种类差异
        检测 chunk 中是否缺少了 origin 中的标签种类（不检查数量）
        返回: 是否通过校验
        """
        passed = True
        
        # 提取所有标签种类
        origin_tag_types = self._extract_tag_types(origin_line)
        chunk_tag_types = self._extract_tag_types(chunk_line)
        
        # 检查 origin 中的标签种类是否在 chunk 中存在
        missing_types = origin_tag_types - chunk_tag_types
        
        for missing in sorted(missing_types):
            self.errors.append(
                f"{filename}: 第{line_num}行 - 缺少格式类型: {missing}"
            )
            passed = False
        
        return passed

    def _verify_format_tags(self, line: str, filename: str, line_num: int) -> bool:
        """
        校验行中的格式标签是否匹配
        格式标记示例: <f=黑体,b=true>文本</f=黑体,b=true>
        返回: 是否通过校验
        """
        passed = True
        
        # 查找所有格式标记 (开标签和闭标签)
        # 开标签: <attrs>
        # 闭标签: </attrs>
        tag_pattern = r'<([fisbzchuv,=\w]*)>|</([fisbzchuv,=\w]*)>'
        tags = list(re.finditer(tag_pattern, line))
        
        if not tags:
            return True  # 没有标签，无需检查
        
        # 使用栈来检查标签匹配
        stack = []
        
        for match in tags:
            open_tag = match.group(1)
            close_tag = match.group(2)
            
            if open_tag is not None:
                # 开标签，压入栈
                stack.append((open_tag, match.start()))
            elif close_tag is not None:
                # 闭标签，检查是否匹配
                if not stack:
                    self.errors.append(
                        f"{filename}: 第{line_num}行 - 多余的闭标签 </{close_tag}> "
                        f"(位置: {match.start()})"
                    )
                    passed = False
                    continue
                
                expected_open, open_pos = stack.pop()
                if expected_open != close_tag:
                    self.errors.append(
                        f"{filename}: 第{line_num}行 - 标签不匹配: "
                        f"开标签 <{expected_open}> (位置: {open_pos}) 与 "
                        f"闭标签 </{close_tag}> (位置: {match.start()}) 不匹配"
                    )
                    passed = False
        
        # 检查是否有未闭合的标签
        for remaining_tag, pos in stack:
            self.errors.append(
                f"{filename}: 第{line_num}行 - 标签未闭合: <{remaining_tag}> "
                f"(位置: {pos})"
            )
            passed = False
        
        return passed

    def print_report(self):
        """打印校验报告"""
        print()
        print("=" * 50)
        print("校验报告")
        print("=" * 50)

        if self.warnings:
            print()
            print(f"警告 ({len(self.warnings)}项):")
            for warning in self.warnings:
                print(f"  ! {warning}")

        if self.errors:
            print()
            print(f"错误 ({len(self.errors)}项):")
            for error in self.errors:
                print(f"  ✗ {error}")
            print()
            print("结果: 校验失败")
            return False
        else:
            print()
            print("结果: 全部通过 ✓")
            return True


def main():
    import argparse

    parser = argparse.ArgumentParser(description='校验chunks和origin_chunks文件')
    parser.add_argument('chunks_dir', nargs='?', default='output/chunks',
                        help='chunks目录路径 (默认: output/chunks)')
    parser.add_argument('origin_chunks_dir', nargs='?', default='output/origin_chunks',
                        help='origin_chunks目录路径 (默认: output/origin_chunks)')

    args = parser.parse_args()

    verifier = ChunksVerifier(args.chunks_dir, args.origin_chunks_dir)
    verifier.verify()
    success = verifier.print_report()

    sys.exit(0 if success else 1)


if __name__ == '__main__':
    main()
