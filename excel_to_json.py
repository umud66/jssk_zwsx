#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
将江苏省各地市职位表Excel文件转换为JSON格式

描述：
    读取指定目录下的所有Excel文件，提取职位信息并转换为JSON格式，
    保存为data.json文件供网页筛选程序使用。

Args：
    excel_dir: Excel文件所在目录路径
    
Returns：
    dict: 包含所有职位信息的字典，格式为 {'data': [...], 'total': int}
"""

import os
import json
import pandas as pd
from pathlib import Path
from typing import Dict, List, Any


def read_excel_file(file_path: str, city_name: str) -> List[Dict[str, Any]]:
    """
    读取Excel文件并转换为字典列表
    
    描述：
        读取单个Excel文件，解析职位信息并转换为标准化的字典格式
        
    Args：
        file_path: Excel文件路径
        city_name: 城市名称，从文件名提取
        
    Returns：
        List[Dict[str, Any]]: 职位信息列表
    """
    try:
        # 根据文件扩展名选择正确的引擎
        file_ext = Path(file_path).suffix.lower()
        if file_ext == '.xls':
            # .xls 文件使用 xlrd 引擎
            engine = 'xlrd'
        elif file_ext == '.xlsx':
            # .xlsx 文件使用 openpyxl 引擎
            engine = 'openpyxl'
        else:
            # 默认尝试 xlrd
            engine = 'xlrd'
        
        # 读取Excel文件，跳过第一行（标题行）
        df = pd.read_excel(file_path, header=1, engine=engine)
        
        # 清理列名，去除前后空格
        df.columns = df.columns.str.strip()
        
        # 如果第一列是标题行，跳过
        # 检查第一行是否包含列名信息
        if len(df) > 0:
            first_row = df.iloc[0]
            # 如果第一行看起来像是列名行，则从第二行开始读取
            if pd.isna(first_row.iloc[0]) or str(first_row.iloc[0]).strip() in ['隶属', '市', '省']:
                # 检查第二行是否是真正的列名
                if len(df) > 1:
                    second_row = df.iloc[1]
                    if '隶属' in str(second_row.iloc[0]) or '关系' in str(second_row.iloc[0]):
                        # 重新读取，使用第二行作为header
                        df = pd.read_excel(file_path, header=2, engine=engine)
                        df.columns = df.columns.str.strip()
        
        # 获取实际的列名（根据Excel文件的实际结构）
        # 通常列名在第一行或第二行
        if len(df.columns) > 13:
            # 尝试识别列名
            columns_map = {
                0: '隶属关系',
                1: '地区代码',
                2: '地区名称',
                3: '单位代码',
                4: '单位名称',
                5: '职位代码',
                6: '职位名称',
                7: '职位简介',
                8: '考试类别',
                9: '开考比例',
                10: '招考人数',
                11: '学历',
                12: '专业',
                13: '其它'
            }
        else:
            # 使用默认列名映射
            columns_map = {}
        
        jobs = []
        
        # 如果列名已经正确，直接使用
        # 否则尝试从数据中识别
        for idx, row in df.iterrows():
            # 跳过空行
            if pd.isna(row.iloc[0]) or str(row.iloc[0]).strip() == '':
                continue
            
            # 跳过标题行
            if '隶属' in str(row.iloc[0]) or '关系' in str(row.iloc[0]):
                continue
            
            try:
                # 使用短字段名以减小文件大小
                job = {
                    'c': city_name,  # city
                    'r': str(row.iloc[0]).strip() if not pd.isna(row.iloc[0]) else '',  # 隶属关系
                    'rc': str(row.iloc[1]).strip() if len(row) > 1 and not pd.isna(row.iloc[1]) else '',  # 地区代码
                    'rn': str(row.iloc[2]).strip() if len(row) > 2 and not pd.isna(row.iloc[2]) else '',  # 地区名称
                    'uc': str(row.iloc[3]).strip() if len(row) > 3 and not pd.isna(row.iloc[3]) else '',  # 单位代码
                    'un': str(row.iloc[4]).strip() if len(row) > 4 and not pd.isna(row.iloc[4]) else '',  # 单位名称
                    'pc': str(row.iloc[5]).strip() if len(row) > 5 and not pd.isna(row.iloc[5]) else '',  # 职位代码
                    'pn': str(row.iloc[6]).strip() if len(row) > 6 and not pd.isna(row.iloc[6]) else '',  # 职位名称
                    'pd': str(row.iloc[7]).strip() if len(row) > 7 and not pd.isna(row.iloc[7]) else '',  # 职位简介
                    'et': str(row.iloc[8]).strip() if len(row) > 8 and not pd.isna(row.iloc[8]) else '',  # 考试类别
                    'kr': str(row.iloc[9]).strip() if len(row) > 9 and not pd.isna(row.iloc[9]) else '',  # 开考比例
                    'rh': str(row.iloc[10]).strip() if len(row) > 10 and not pd.isna(row.iloc[10]) else '',  # 招考人数
                    'ed': str(row.iloc[11]).strip() if len(row) > 11 and not pd.isna(row.iloc[11]) else '',  # 学历
                    'mj': str(row.iloc[12]).strip() if len(row) > 12 and not pd.isna(row.iloc[12]) else '',  # 专业
                    'ot': str(row.iloc[13]).strip() if len(row) > 13 and not pd.isna(row.iloc[13]) else '',  # 其它
                }
                
                # 过滤掉完全空白的记录（使用短字段名）
                if job['pn'] or job['un']:
                    jobs.append(job)
            except Exception as e:
                print(f"处理第 {idx} 行时出错: {e}")
                continue
        
        return jobs
    
    except Exception as e:
        print(f"读取文件 {file_path} 时出错: {e}")
        return []


def process_all_excel_files(excel_dir: str) -> Dict[str, Any]:
    """
    处理目录下所有Excel文件
    
    描述：
        遍历指定目录下的所有Excel文件，读取并转换为JSON格式
        
    Args：
        excel_dir: Excel文件所在目录路径
        
    Returns：
        Dict[str, Any]: 包含所有职位信息和统计数据的字典
    """
    excel_dir_path = Path(excel_dir)
    all_jobs = []
    
    # 获取所有Excel文件，按文件名排序
    excel_files = sorted(excel_dir_path.glob('*.xls'))
    
    if not excel_files:
        excel_files = sorted(excel_dir_path.glob('*.xlsx'))
    
    print(f"找到 {len(excel_files)} 个Excel文件")
    
    for excel_file in excel_files:
        # 从文件名提取城市名称（去除序号和扩展名）
        city_name = excel_file.stem
        # 去除开头的序号（如 "01-", "14-" 等）
        if '-' in city_name:
            city_name = city_name.split('-', 1)[1]
        
        print(f"正在处理: {excel_file.name} ({city_name})")
        jobs = read_excel_file(str(excel_file), city_name)
        print(f"  提取了 {len(jobs)} 条职位信息")
        all_jobs.extend(jobs)
    
    result = {
        'data': all_jobs,
        'total': len(all_jobs),
        'cities': list(set(job['c'] for job in all_jobs))  # 使用短字段名 'c' 代替 'city'
    }
    
    print(f"\n总计: {result['total']} 条职位信息，涵盖 {len(result['cities'])} 个城市")
    
    return result


def main():
    """
    主函数
    
    描述：
        执行Excel到JSON的转换流程
        
    Args：
        无
        
    Returns：
        无
    """
    # 获取脚本所在目录
    script_dir = Path(__file__).parent
    excel_dir = script_dir / "江苏省2026年度考试录用公务员各地职位表"
    output_file = script_dir / "data.json"
    
    if not excel_dir.exists():
        print(f"错误: 目录不存在 {excel_dir}")
        return
    
    # 处理所有Excel文件
    result = process_all_excel_files(str(excel_dir))
    
    # 保存为JSON文件（压缩格式，无缩进，减小文件大小）
    with open(output_file, 'w', encoding='utf-8') as f:
        json.dump(result, f, ensure_ascii=False, separators=(',', ':'))
    
    print(f"\n数据已保存到: {output_file}")


if __name__ == "__main__":
    main()

