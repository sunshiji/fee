# -*- coding: utf-8 -*-
# 测试核心功能模块
import os
import sys
import io

# 设置标准输出编码为UTF-8
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

# 添加当前目录到路径
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

print("=" * 60)
print("测试核心功能模块")
print("=" * 60)

all_passed = True

try:
    print("\n1. 测试导入 config 模块...")
    from config import BASE_DIR, DATA_DIR, OUTPUT_DIR, COLLEGE_NAME
    print(f"   BASE_DIR: {BASE_DIR}")
    print(f"   DATA_DIR: {DATA_DIR}")
    print(f"   OUTPUT_DIR: {OUTPUT_DIR}")
    print(f"   COLLEGE_NAME: {COLLEGE_NAME}")
    print("   OK: 成功")
except Exception as e:
    print(f"   FAIL: {e}")
    all_passed = False

try:
    print("\n2. 测试导入 data_models 模块...")
    from data_models import PartyMember, PartyBranch, SalaryInfo, DeductionInfo
    
    salary = SalaryInfo(position_salary=5000, rank_salary=3000, fixed_salary=0, basic_performance=2000)
    deduction = DeductionInfo(housing_fund=800, medical_insurance=200, pension_insurance=600, 
                               occupational_annuity=200, large_medical=100, unemployment_insurance=50, 
                               personal_income_tax=0)
    member = PartyMember(name="测试党员", salary_info=salary, deduction_info=deduction)
    branch = PartyBranch(name="测试党支部", sequence=1)
    
    print(f"   创建党员对象: {member.name}")
    print(f"   创建支部对象: {branch.name}")
    print("   OK: 成功")
except Exception as e:
    print(f"   FAIL: {e}")
    all_passed = False

try:
    print("\n3. 测试导入 fee_calculator 模块...")
    from fee_calculator import FeeCalculator
    
    FeeCalculator.calculate_member_fee(member)
    print(f"   缴费基数: {member.payment_base}")
    print(f"   月党费: {member.monthly_fee}")
    print("   OK: 成功")
except Exception as e:
    print(f"   FAIL: {e}")
    all_passed = False

try:
    print("\n4. 测试导入 member_manager 模块...")
    from member_manager import MemberManager
    
    manager = MemberManager()
    manager.add_branch(branch)
    manager.add_member(member, branch.name)
    
    branches = manager.get_all_branches()
    print(f"   支部数量: {len(branches)}")
    print(f"   党员数量: {len(manager.get_all_members())}")
    print("   OK: 成功")
except Exception as e:
    print(f"   FAIL: {e}")
    all_passed = False

try:
    print("\n5. 测试导入 excel_generator 模块...")
    from excel_generator import ExcelGenerator
    
    generator = ExcelGenerator()
    print("   OK: 成功")
except Exception as e:
    print(f"   FAIL: {e}")
    all_passed = False

try:
    print("\n6. 测试导入 database_manager 模块...")
    from database_manager import DatabaseManager
    
    db_manager = DatabaseManager()
    print("   OK: 成功")
except Exception as e:
    print(f"   FAIL: {e}")
    all_passed = False

try:
    print("\n7. 测试导入 statistics 模块...")
    from statistics import FeeStatistics
    
    report = FeeStatistics.generate_statistics_report(branches, 2026, 4)
    print(f"   统计报告长度: {len(report)} 字符")
    print("   OK: 成功")
except Exception as e:
    print(f"   FAIL: {e}")
    all_passed = False

try:
    print("\n8. 测试导入 excel_importer 模块...")
    from excel_importer import ExcelImporter
    
    importer = ExcelImporter()
    print("   OK: 成功")
except Exception as e:
    print(f"   FAIL: {e}")
    all_passed = False

try:
    print("\n9. 测试导入 word_importer 模块...")
    from word_importer import WordImporter
    
    word_importer = WordImporter()
    print("   OK: 成功")
except Exception as e:
    print(f"   FAIL: {e}")
    all_passed = False

try:
    print("\n10. 测试导入 tkinter 模块...")
    import tkinter as tk
    from tkinter import ttk, messagebox, filedialog, scrolledtext
    
    print("   OK: 成功")
except Exception as e:
    print(f"   FAIL: {e}")
    all_passed = False

try:
    print("\n11. 测试导入 openpyxl 模块...")
    import openpyxl
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
    
    print(f"   openpyxl 版本: {openpyxl.__version__}")
    print("   OK: 成功")
except Exception as e:
    print(f"   FAIL: {e}")
    all_passed = False

try:
    print("\n12. 测试导入 pandas 模块...")
    import pandas as pd
    
    print(f"   pandas 版本: {pd.__version__}")
    print("   OK: 成功")
except Exception as e:
    print(f"   FAIL: {e}")
    all_passed = False

try:
    print("\n13. 测试导入 docx 模块...")
    import docx
    
    print(f"   python-docx 版本: {docx.__version__}")
    print("   OK: 成功")
except Exception as e:
    print(f"   FAIL: {e}")
    all_passed = False

try:
    print("\n14. 测试导入 sqlite3 模块...")
    import sqlite3
    
    print(f"   SQLite 版本: {sqlite3.sqlite_version}")
    print("   OK: 成功")
except Exception as e:
    print(f"   FAIL: {e}")
    all_passed = False

print("\n" + "=" * 60)
if all_passed:
    print("所有核心功能模块测试通过！")
else:
    print("部分模块测试失败，请检查错误信息。")
print("=" * 60)
