# 测试合并单元格的导入问题
import os
import sys
from openpyxl import Workbook
from openpyxl.utils import get_column_letter

# 添加当前目录到路径
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from config import OUTPUT_DIR
from excel_importer import ExcelImporter


def create_test_file_with_merged_cells():
    """创建包含合并单元格的测试文件（模拟实际表格）"""
    print("\n" + "=" * 60)
    print("创建包含合并单元格的测试文件")
    print("=" * 60)
    
    wb = Workbook()
    ws = wb.active
    ws.title = "2月"
    
    # 第1行：标题（合并单元格）
    ws['A1'] = "计算机科学与技术学院2026年2月党费收缴明细表"
    ws.merge_cells('A1:Q1')
    
    # 第2行：分组表头
    ws['E2'] = "工资项目（元）"
    ws.merge_cells('E2:H2')
    
    ws['I2'] = "扣款项目（元）"
    ws.merge_cells('I2:P2')
    
    ws['Q2'] = "缴费基数"
    ws['R2'] = "每月应缴党费（元）"
    
    # 第3行：实际列名
    headers_row3 = [
        '支部序号', '所属党支部', '序号', '姓 名',
        '岗位工资', '薪级工资', '高定工资', '基础性绩效',
        '住房公积金', '医疗保险', '养老保险', '职业年金',
        '大额医疗', '失业保险', '个人所得税', '',
        '缴费基数', '每月应缴党费（元）'
    ]
    
    for col_idx, header in enumerate(headers_row3, 1):
        ws.cell(row=3, column=col_idx, value=header)
    
    # 第4行开始：数据（使用合并单元格表示支部信息）
    # 支部1：教工第一党支部（合并单元格）
    # 第4-5行：支部序号=1，所属党支部=教工第一党支部（合并）
    
    # 设置合并单元格
    ws.merge_cells('A4:A5')  # 支部序号列合并
    ws.merge_cells('B4:B5')  # 所属党支部列合并
    
    # 设置合并单元格的值（只有左上角单元格有值）
    ws['A4'] = 1
    ws['B4'] = "教工第一党支部"
    
    # 第4行数据
    ws['C4'] = 1
    ws['D4'] = "张三"
    ws['E4'] = 5000
    ws['F4'] = 3000
    ws['G4'] = 0
    ws['H4'] = 2000
    ws['I4'] = 800
    ws['J4'] = 200
    ws['K4'] = 400
    ws['L4'] = 100
    ws['M4'] = 50
    ws['N4'] = 50
    ws['O4'] = 0
    ws['Q4'] = 8600
    ws['R4'] = 129
    
    # 第5行数据（注意：A5和B5是合并单元格的一部分，没有值）
    ws['C5'] = 2
    ws['D5'] = "李四"
    ws['E5'] = 6000
    ws['F5'] = 4000
    ws['G5'] = 0
    ws['H5'] = 2500
    ws['I5'] = 900
    ws['J5'] = 250
    ws['K5'] = 500
    ws['L5'] = 120
    ws['M5'] = 60
    ws['N5'] = 60
    ws['O5'] = 0
    ws['Q5'] = 10850
    ws['R5'] = 217
    
    # 支部2：本科生第一党支部（合并单元格）
    # 第6-8行：支部序号=2，所属党支部=本科生第一党支部（合并）
    
    # 设置合并单元格
    ws.merge_cells('A6:A8')  # 支部序号列合并
    ws.merge_cells('B6:B8')  # 所属党支部列合并
    
    # 设置合并单元格的值
    ws['A6'] = 2
    ws['B6'] = "本科生第一党支部"
    
    # 第6行数据
    ws['C6'] = 1
    ws['D6'] = "王五"
    ws['E6'] = 3000
    ws['F6'] = 2000
    ws['G6'] = 0
    ws['H6'] = 1000
    ws['I6'] = 500
    ws['J6'] = 100
    ws['K6'] = 200
    ws['L6'] = 50
    ws['M6'] = 25
    ws['N6'] = 25
    ws['O6'] = 0
    ws['Q6'] = 5200
    ws['R6'] = 78
    
    # 第7行数据
    ws['C7'] = 2
    ws['D7'] = "赵六"
    ws['E7'] = 4000
    ws['F7'] = 2500
    ws['G7'] = 0
    ws['H7'] = 1500
    ws['I7'] = 600
    ws['J7'] = 150
    ws['K7'] = 300
    ws['L7'] = 70
    ws['M7'] = 35
    ws['N7'] = 35
    ws['O7'] = 0
    ws['Q7'] = 6950
    ws['R7'] = 104.2
    
    # 第8行数据
    ws['C8'] = 3
    ws['D8'] = "钱七"
    ws['E8'] = 3500
    ws['F8'] = 2200
    ws['G8'] = 0
    ws['H8'] = 1200
    ws['I8'] = 550
    ws['J8'] = 120
    ws['K8'] = 250
    ws['L8'] = 60
    ws['M8'] = 30
    ws['N8'] = 30
    ws['O8'] = 0
    ws['Q8'] = 5980
    ws['R8'] = 89.7
    
    # 保存文件
    test_file_path = os.path.join(OUTPUT_DIR, '测试_合并单元格_实际格式.xlsx')
    wb.save(test_file_path)
    
    print(f"[OK] 已创建测试文件: {test_file_path}")
    print(f"     结构: 3级表头，支部信息使用合并单元格")
    print(f"     合并单元格:")
    print(f"       - 支部1: A4:A5, B4:B5 (2名党员)")
    print(f"       - 支部2: A6:A8, B6:B8 (3名党员)")
    print(f"     注意: 合并单元格只有左上角有值，其他单元格为空")
    
    return test_file_path


def test_import_with_merged_cells():
    """测试导入包含合并单元格的文件"""
    print("\n" + "=" * 60)
    print("测试导入包含合并单元格的文件")
    print("=" * 60)
    
    file_path = create_test_file_with_merged_cells()
    
    # 首先直接读取Excel文件，检查合并单元格的值
    print("\n--- 直接读取Excel文件，检查合并单元格 ---")
    from openpyxl import load_workbook
    
    wb = load_workbook(file_path, data_only=True)
    ws = wb.active
    
    print("\n检查支部信息列（A列和B列）:")
    print(f"{'行号':<6}{'A列(支部序号)':<15}{'B列(所属党支部)':<20}{'D列(姓名)':<10}")
    print("-" * 60)
    
    for row in range(4, 9):
        a_value = ws.cell(row=row, column=1).value
        b_value = ws.cell(row=row, column=2).value
        d_value = ws.cell(row=row, column=4).value
        print(f"{row:<6}{str(a_value):<15}{str(b_value):<20}{str(d_value):<10}")
    
    print("\n检查合并单元格范围:")
    for merged_range in ws.merged_cells.ranges:
        print(f"  {merged_range}")
        # 检查合并区域的左上角值
        min_row, min_col, max_row, max_col = merged_range.bounds
        top_left_value = ws.cell(row=min_row, column=min_col).value
        print(f"    左上角值: {top_left_value}")
    
    wb.close()
    
    # 测试导入器
    print("\n--- 测试导入器 ---")
    importer = ExcelImporter()
    
    # 测试 _get_merged_cell_value 方法
    print("\n测试 _get_merged_cell_value 方法:")
    wb2 = load_workbook(file_path, data_only=True)
    ws2 = wb2.active
    
    for row in range(4, 9):
        # 使用 _get_merged_cell_value 方法
        a_value = importer._get_merged_cell_value(ws2, row, 1)
        b_value = importer._get_merged_cell_value(ws2, row, 2)
        d_value = importer._get_merged_cell_value(ws2, row, 4)
        print(f"行{row}: 支部序号={a_value}, 所属党支部={b_value}, 姓名={d_value}")
    
    wb2.close()
    
    # 实际导入测试
    print("\n--- 实际导入测试 ---")
    result = importer.import_from_file(file_path, sheet_type='detail')
    
    print(f"\n导入结果:")
    print(f"  成功: {result.success}")
    print(f"  支部数量: {result.total_branches}")
    print(f"  党员数量: {result.total_members}")
    
    if result.errors:
        print(f"\n错误:")
        for e in result.errors:
            print(f"  - {e}")
    
    if result.warnings:
        print(f"\n警告:")
        for w in result.warnings:
            print(f"  - {w}")
    
    if result.success:
        print(f"\n导入的支部详情:")
        print("-" * 60)
        for branch in result.branches:
            print(f"\n支部: {branch.name} (序号: {branch.sequence})")
            print(f"  党员数量: {len(branch.members)}")
            for member in branch.members:
                print(f"    - {member.name}: 缴费基数={member.payment_base:.2f}, 党费={member.monthly_fee:.2f}")
    
    return result


if __name__ == "__main__":
    print("\n" + "#" * 60)
    print("# 测试合并单元格的导入问题")
    print("#" * 60)
    
    result = test_import_with_merged_cells()
    
    print("\n" + "#" * 60)
    if result.success and result.total_members == 5:
        print("# 测试通过！合并单元格处理正确")
    else:
        print("# 测试失败！")
        print(f"# 预期: 5名党员, 实际: {result.total_members}名党员")
    print("#" * 60)
