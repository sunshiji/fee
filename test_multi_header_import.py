# 测试多级表头和带空格的列名
import os
import sys
from openpyxl import Workbook
from openpyxl.utils import get_column_letter

# 添加当前目录到路径
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from config import OUTPUT_DIR


def create_test_multi_header_file():
    """创建模拟用户实际表格的测试文件（多级表头、带空格的列名）"""
    print("\n" + "=" * 60)
    print("创建模拟实际表格的测试文件")
    print("=" * 60)
    
    wb = Workbook()
    ws = wb.active
    ws.title = "2月"
    
    # 第1行：标题（合并单元格）
    ws['A1'] = "计算机科学与技术学院2026年2月党费收缴明细表"
    ws.merge_cells('A1:Q1')
    
    # 第2行：分组表头
    # A2: (空)
    # B2: (空)
    # C2: (空)
    # D2: (空)
    # E2:H2: 工资项目（元）
    # I2:P2: 扣款项目（元）
    # Q2: 缴费基数
    # R2: 每月应缴党费（元）
    # S2: 支部每月应缴党费（元）
    
    ws['E2'] = "工资项目（元）"
    ws.merge_cells('E2:H2')
    
    ws['I2'] = "扣款项目（元）"
    ws.merge_cells('I2:P2')
    
    ws['Q2'] = "缴费基数"
    ws['R2'] = "每月应缴党费（元）"
    ws['S2'] = "支部每月应缴党费（元）"
    
    # 第3行：实际列名（带空格）
    # A3: 支部序号
    # B3: 所属党支部
    # C3: 序号
    # D3: 姓 名
    # E3: 岗位工资
    # F3: 薪级工资
    # G3: 高定工资
    # H3: 基础性绩效
    # I3: 住房公积金
    # J3: 医疗保险
    # K3: 养老保险
    # L3: 职业年金
    # M3: 大额医疗
    # N3: 失业保险
    # O3: 个人所得税
    # P3: (空)
    # Q3: 缴费基数
    # R3: 每月应缴党费（元）
    # S3: 支部每月应缴党费（元）
    
    headers_row3 = [
        '支部序号', '所属党支部', '序号', '姓 名',
        '岗位工资', '薪级工资', '高定工资', '基础性绩效',
        '住房公积金', '医疗保险', '养老保险', '职业年金',
        '大额医疗', '失业保险', '个人所得税', '',
        '缴费基数', '每月应缴党费（元）', '支部每月应缴党费（元）'
    ]
    
    for col_idx, header in enumerate(headers_row3, 1):
        ws.cell(row=3, column=col_idx, value=header)
    
    # 第4行开始：数据
    # 教工第一党支部的党员
    test_data = [
        # 支部1：教工第一党支部（合并单元格）
        (1, '教工第一党支部', 1, '张三', 5000, 3000, 0, 2000, 800, 200, 400, 100, 50, 50, 0, '', 8600, 129, ''),
        (1, '教工第一党支部', 2, '李四', 6000, 4000, 0, 2500, 900, 250, 500, 120, 60, 60, 0, '', 10850, 217, ''),
        # 支部2：本科生第一党支部
        (2, '本科生第一党支部', 1, '王五', 3000, 2000, 0, 1000, 500, 100, 200, 50, 25, 25, 0, '', 5200, 78, ''),
        (2, '本科生第一党支部', 2, '赵六', 4000, 2500, 0, 1500, 600, 150, 300, 70, 35, 35, 0, '', 6950, 104.2, ''),
        (2, '本科生第一党支部', 3, '钱七', 3500, 2200, 0, 1200, 550, 120, 250, 60, 30, 30, 0, '', 5980, 89.7, ''),
    ]
    
    for row_idx, row_data in enumerate(test_data, 4):
        for col_idx, value in enumerate(row_data, 1):
            ws.cell(row=row_idx, column=col_idx, value=value)
    
    # 保存文件
    test_file_path = os.path.join(OUTPUT_DIR, '测试_模拟实际表格_多级表头.xlsx')
    wb.save(test_file_path)
    
    print(f"[OK] 已创建测试文件: {test_file_path}")
    print(f"     结构: 3级表头（标题行+分组行+列名行）")
    print(f"     列名带空格: 如'姓 名'")
    print(f"     包含: 2个支部, 5名党员")
    
    return test_file_path


def create_test_single_branch_file():
    """创建单个支部的测试文件（没有支部信息列）"""
    print("\n" + "=" * 60)
    print("创建单个支部的测试文件（无支部信息列）")
    print("=" * 60)
    
    wb = Workbook()
    ws = wb.active
    ws.title = "教工第一党支部"
    
    # 第1行：标题
    ws['A1'] = "教工第一党支部 - 2026年2月党费收缴明细表"
    ws.merge_cells('A1:O1')
    
    # 第2行：分组表头
    ws['C2'] = "工资项目（元）"
    ws.merge_cells('C2:F2')
    
    ws['G2'] = "扣款项目（元）"
    ws.merge_cells('G2:N2')
    
    # 第3行：列名
    headers_row3 = [
        '序号', '姓 名',
        '岗位工资', '薪级工资', '高定工资', '基础性绩效',
        '住房公积金', '医疗保险', '养老保险', '职业年金',
        '大额医疗', '失业保险', '个人所得税',
        '缴费基数', '每月应缴党费'
    ]
    
    for col_idx, header in enumerate(headers_row3, 1):
        ws.cell(row=3, column=col_idx, value=header)
    
    # 数据
    test_data = [
        (1, '张三', 5000, 3000, 0, 2000, 800, 200, 400, 100, 50, 50, 0, 8600, 129),
        (2, '李四', 6000, 4000, 0, 2500, 900, 250, 500, 120, 60, 60, 0, 10850, 217),
        (3, '王五', 4500, 2800, 0, 1800, 700, 180, 350, 90, 45, 45, 0, 7680, 115.2),
    ]
    
    for row_idx, row_data in enumerate(test_data, 4):
        for col_idx, value in enumerate(row_data, 1):
            ws.cell(row=row_idx, column=col_idx, value=value)
    
    # 保存文件
    test_file_path = os.path.join(OUTPUT_DIR, '测试_单个支部_无支部列.xlsx')
    wb.save(test_file_path)
    
    print(f"[OK] 已创建测试文件: {test_file_path}")
    print(f"     结构: 3级表头，无支部信息列")
    print(f"     支部名称: 从标题中提取")
    print(f"     包含: 1个支部, 3名党员")
    
    return test_file_path


def test_improved_importer():
    """测试改进后的导入器"""
    from excel_importer import ExcelImporter
    
    print("\n" + "#" * 60)
    print("# 测试改进后的Excel导入器")
    print("#" * 60)
    
    # 创建测试文件
    file1 = create_test_multi_header_file()
    file2 = create_test_single_branch_file()
    
    importer = ExcelImporter()
    
    # 测试1：多级表头、带空格列名、有支部信息列
    print("\n" + "=" * 60)
    print("测试1：多级表头、带空格列名、有支部信息列")
    print("=" * 60)
    
    result1 = importer.import_from_file(file1, sheet_type='detail')
    
    print(f"\n导入结果:")
    print(f"  成功: {result1.success}")
    print(f"  支部数量: {result1.total_branches}")
    print(f"  党员数量: {result1.total_members}")
    
    if result1.errors:
        print(f"\n错误:")
        for e in result1.errors:
            print(f"  - {e}")
    
    if result1.success:
        print(f"\n导入的支部详情:")
        print("-" * 60)
        for branch in result1.branches:
            print(f"\n支部: {branch.name} (序号: {branch.sequence})")
            print(f"  党员数量: {len(branch.members)}")
            for member in branch.members:
                print(f"    - {member.name}: 缴费基数={member.payment_base:.2f}, 党费={member.monthly_fee:.2f}")
    
    # 测试2：单个支部、无支部信息列
    print("\n" + "=" * 60)
    print("测试2：单个支部、无支部信息列")
    print("=" * 60)
    
    result2 = importer.import_from_file(file2, sheet_type='detail')
    
    print(f"\n导入结果:")
    print(f"  成功: {result2.success}")
    print(f"  支部数量: {result2.total_branches}")
    print(f"  党员数量: {result2.total_members}")
    
    if result2.errors:
        print(f"\n错误:")
        for e in result2.errors:
            print(f"  - {e}")
    
    if result2.success:
        print(f"\n导入的支部详情:")
        print("-" * 60)
        for branch in result2.branches:
            print(f"\n支部: {branch.name} (序号: {branch.sequence})")
            print(f"  党员数量: {len(branch.members)}")
            for member in branch.members:
                print(f"    - {member.name}: 缴费基数={member.payment_base:.2f}, 党费={member.monthly_fee:.2f}")
    
    # 总结
    print("\n" + "#" * 60)
    if result1.success and result2.success:
        print("# 所有测试通过！")
    else:
        print("# 部分测试失败")
    print("#" * 60)
    
    return result1.success and result2.success


if __name__ == "__main__":
    test_improved_importer()
