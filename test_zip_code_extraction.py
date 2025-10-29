#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
测试ZIP文件单号提取功能的脚本
"""

import sys
import os
import re

# 添加项目路径到sys.path
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from uploader.utils import extract_info_from_zip

def test_zip_code_extraction():
    """测试ZIP文件单号提取功能"""
    # 使用现有的测试ZIP文件
    zip_path = r"test_case\EOSC_4712508269337893_KC+官渡区兰兴招待所+官渡区关上街道办事处双桥村637号(1).zip"
    
    if not os.path.exists(zip_path):
        print(f"测试ZIP文件不存在: {zip_path}")
        return
    
    print(f"开始测试ZIP文件单号提取: {zip_path}")
    
    # 测试新的提取逻辑
    zip_file_name = os.path.basename(zip_path)
    print(f"ZIP文件名: {zip_file_name}")
    
    # 新的提取逻辑
    zip_code_match = re.match(r'([A-Za-z0-9_]+)', zip_file_name)
    zip_code_part = zip_code_match.group(1) if zip_code_match else None
    print(f"提取到的单号: {zip_code_part}")
    
    # 旧的提取逻辑（仅提取数字）
    old_zip_code_match = re.search(r'EOSC_(\d+)_', zip_file_name)
    old_zip_code_part = old_zip_code_match.group(1) if old_zip_code_match else None
    print(f"旧方法提取到的单号: {old_zip_code_part}")
    
    print("\n" + "="*50)
    print("完整处理结果:")
    print("="*50)
    
    # 完整处理测试
    results = extract_info_from_zip(zip_path)
    
    for result in results:
        if 'error' in result:
            print(f"错误: {result['error']}")
        else:
            print(f"文件名: {result['file_name']}")
            print(f"单号: {result['code_part']}")
            print(f"维护费（含税）合计: {result['doc_maintenance_total']}")
            print(f"总估算价格: {result['total_price']}")
            print(f"宽带维护费: {result['maintenance_fee']}")
            print(f"宽带服务费: {result['service_fee']}")
            print(f"终端费: {result['terminal_fee']}")
            print(f"费用总和: {result['total_fees']}")
            print(f"光缆长度: {result['fiber_length']}")
            print(f"验算通过: {result['verification_passed']}")
            print(f"提取状态: {result['extraction_status']}")
            print("-" * 30)

if __name__ == "__main__":
    test_zip_code_extraction()