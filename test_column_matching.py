#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import pandas as pd

def test_column_matching():
    """测试列名匹配逻辑"""
    df = pd.read_excel('yxc.xlsx')
    print('原始列名:', list(df.columns))
    print('列名匹配测试:')
    
    # 查找剩余天数列
    possible_remaining = ['剩余天数', '剩余时间', '到期天数', '过期天数', '天数', 'days', 'remaining_days', '剩余', '到期', '过期']
    for col in df.columns:
        col_str = str(col).lower()
        print(f'列: {col} -> {col_str}')
        for name in possible_remaining:
            if name in col_str:
                print(f'  匹配剩余: {name}')
    
    # 查找总天数列
    possible_total = ['总天数', '总时间', '总天', 'total_days', 'total', '天']
    for col in df.columns:
        col_str = str(col).lower()
        for name in possible_total:
            if name in col_str:
                print(f'  匹配总天: {name}')
    
    # 查找开始时间列
    possible_start = ['开始时间', '开始日期', 'start_date', 'start_time', '开始']
    for col in df.columns:
        col_str = str(col).lower()
        for name in possible_start:
            if name in col_str:
                print(f'  匹配开始: {name}')

if __name__ == "__main__":
    test_column_matching() 