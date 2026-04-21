# 系统测试脚本
import os
import sys

# 添加项目目录到Python路径
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from data_models import PartyMember, PartyBranch, SalaryInfo, DeductionInfo
from member_manager import MemberManager
from fee_calculator import FeeCalculator
from excel_generator import ExcelGenerator
from statistics import FeeStatistics
from config import OUTPUT_DIR


def create_test_data():
    """创建测试数据"""
    print("创建测试数据...")
    
    # 创建党支部
    branch1 = PartyBranch(
        name="教工第一支部",
        sequence=1
    )
    
    branch2 = PartyBranch(
        name="本科生第一支部",
        sequence=2
    )
    
    # 创建教工党员（有工资信息）
    member1 = PartyMember(
        name="张三",
        salary_info=SalaryInfo(
            position_salary=2185.0,
            rank_salary=692.0,
            fixed_salary=0.0,
            basic_performance=1100.0
        ),
        deduction_info=DeductionInfo(
            housing_fund=1094.0,
            medical_insurance=152.62,
            pension_insurance=706.72,
            occupational_annuity=353.36,
            large_medical=39.66,
            unemployment_insurance=44.17,
            personal_income_tax=0.0
        )
    )
    
    member2 = PartyMember(
        name="李四",
        salary_info=SalaryInfo(
            position_salary=2375.0,
            rank_salary=845.0,
            fixed_salary=0.0,
            basic_performance=2030.0
        ),
        deduction_info=DeductionInfo(
            housing_fund=1200.0,
            medical_insurance=165.78,
            pension_insurance=785.62,
            occupational_annuity=392.81,
            large_medical=45.32,
            unemployment_insurance=52.43,
            personal_income_tax=67.57
        )
    )
    
    # 创建学生党员（工资信息为0）
    member3 = PartyMember(
        name="王五",
        salary_info=SalaryInfo(
            position_salary=0.0,
            rank_salary=0.0,
            fixed_salary=0.0,
            basic_performance=0.0
        ),
        deduction_info=DeductionInfo(
            housing_fund=0.0,
            medical_insurance=0.0,
            pension_insurance=0.0,
            occupational_annuity=0.0,
            large_medical=0.0,
            unemployment_insurance=0.0,
            personal_income_tax=0.0
        )
    )
    
    member4 = PartyMember(
        name="赵六",
        salary_info=SalaryInfo(
            position_salary=0.0,
            rank_salary=0.0,
            fixed_salary=0.0,
            basic_performance=0.0
        ),
        deduction_info=DeductionInfo(
            housing_fund=0.0,
            medical_insurance=0.0,
            pension_insurance=0.0,
            occupational_annuity=0.0,
            large_medical=0.0,
            unemployment_insurance=0.0,
            personal_income_tax=0.0
        )
    )
    
    # 将党员添加到支部
    branch1.members = [member1, member2]
    branch2.members = [member3, member4]
    
    # 为党员分配序号
    for i, member in enumerate(branch1.members, 1):
        member.sequence = i
        member.branch_name = branch1.name
        member.branch_sequence = branch1.sequence
    
    for i, member in enumerate(branch2.members, 1):
        member.sequence = i
        member.branch_name = branch2.name
        member.branch_sequence = branch2.sequence
    
    return [branch1, branch2]


def test_fee_calculation(branches):
    """测试党费计算"""
    print("\n测试党费计算...")
    print("-" * 60)
    
    # 计算所有党员的党费
    FeeCalculator.calculate_all_fees(branches)
    
    for branch in branches:
        print(f"\n【{branch.name}】")
        for member in branch.members:
            print(f"  {member.name}: 缴费基数={member.payment_base:.2f}元, 月党费={member.monthly_fee:.2f}元")
    
    # 计算统计信息
    total_members = sum(len(b.members) for b in branches)
    total_fee = sum(sum(m.monthly_fee for m in b.members) for b in branches)
    
    print(f"\n总计: {total_members}名党员, 总党费: {total_fee:.2f}元")
    print("党费计算测试通过!")


def test_member_manager(branches):
    """测试人员管理器"""
    print("\n测试人员管理器...")
    print("-" * 60)
    
    manager = MemberManager()
    
    # 添加支部
    for branch in branches:
        manager.add_branch(branch)
    
    # 测试获取所有支部
    all_branches = manager.get_all_branches()
    print(f"支部数量: {len(all_branches)}")
    
    # 测试获取单个支部
    branch = manager.get_branch("教工第一支部")
    if branch:
        print(f"找到支部: {branch.name}, 党员数: {len(branch.members)}")
    
    # 测试获取所有党员
    all_members = manager.get_all_members()
    print(f"党员总数: {len(all_members)}")
    
    # 测试获取单个党员
    member = manager.get_member("张三")
    if member:
        print(f"找到党员: {member.name}, 所属支部: {member.branch_name}")
    
    # 测试保存和加载
    print("\n测试数据保存和加载...")
    filepath = manager.save_to_file(2026, 2)
    print(f"数据已保存到: {filepath}")
    
    # 创建新的管理器并加载数据
    new_manager = MemberManager()
    new_manager.load_from_file(2026, 2)
    
    loaded_branches = new_manager.get_all_branches()
    print(f"加载后支部数量: {len(loaded_branches)}")
    
    for branch in loaded_branches:
        print(f"  {branch.name}: {len(branch.members)}名党员")
    
    print("人员管理器测试通过!")


def test_excel_generator(branches):
    """测试Excel生成器"""
    print("\n测试Excel生成器...")
    print("-" * 60)
    
    # 确保输出目录存在
    if not os.path.exists(OUTPUT_DIR):
        os.makedirs(OUTPUT_DIR)
    
    generator = ExcelGenerator()
    
    # 测试生成党费明细表
    print("\n生成党费明细表...")
    filepath = generator.generate_fee_detail_sheet(branches, 2026, 2)
    print(f"文件已生成: {filepath}")
    
    # 测试生成各支部党费收缴子表
    print("\n生成各支部党费收缴子表...")
    filepaths = generator.generate_all_branch_sheets(branches, 2026, 2)
    for fp in filepaths:
        print(f"文件已生成: {fp}")
    
    # 测试生成党费汇总表
    print("\n生成党费汇总表...")
    filepath = generator.generate_summary_sheet(branches, 2026, 2)
    print(f"文件已生成: {filepath}")
    
    print("Excel生成器测试通过!")


def test_statistics(branches):
    """测试统计功能"""
    print("\n测试统计功能...")
    print("-" * 60)
    
    # 计算党费
    FeeCalculator.calculate_all_fees(branches)
    
    # 测试支部统计
    print("\n各支部统计:")
    branch_stats = FeeStatistics.calculate_all_branch_statistics(branches)
    for stats in branch_stats:
        print(f"  {stats.branch_name}: {stats.member_count}人, 总党费{stats.total_fee:.2f}元, 人均{stats.avg_fee:.2f}元")
    
    # 测试总体统计
    print("\n总体统计:")
    overall_stats = FeeStatistics.calculate_overall_statistics(branches)
    print(f"  支部数量: {overall_stats.total_branches}")
    print(f"  党员总数: {overall_stats.total_members}")
    print(f"  总党费: {overall_stats.total_fee:.2f}元")
    print(f"  人均党费: {overall_stats.avg_fee_per_member:.2f}元")
    print(f"  党费最高支部: {overall_stats.max_fee_branch}")
    print(f"  党费最低支部: {overall_stats.min_fee_branch}")
    
    # 测试生成统计报告
    print("\n生成统计报告...")
    report = FeeStatistics.generate_statistics_report(branches, 2026, 2)
    print(report)
    
    # 测试分布统计
    print("\n党费分布统计:")
    fee_dist = FeeStatistics.get_fee_distribution(branches)
    for range_str, count in fee_dist.items():
        if count > 0:
            print(f"  {range_str}: {count}人")
    
    print("\n缴费基数分布统计:")
    base_dist = FeeStatistics.get_payment_base_distribution(branches)
    for range_str, count in base_dist.items():
        if count > 0:
            print(f"  {range_str}: {count}人")
    
    print("统计功能测试通过!")


def main():
    """主测试函数"""
    print("=" * 60)
    print("党费收取系统测试")
    print("=" * 60)
    
    # 创建测试数据
    branches = create_test_data()
    
    # 测试党费计算
    test_fee_calculation(branches)
    
    # 重新创建测试数据（因为上面的测试已经修改了数据）
    branches = create_test_data()
    
    # 测试人员管理器
    test_member_manager(branches)
    
    # 重新创建测试数据
    branches = create_test_data()
    FeeCalculator.calculate_all_fees(branches)
    
    # 测试Excel生成器
    test_excel_generator(branches)
    
    # 重新创建测试数据
    branches = create_test_data()
    FeeCalculator.calculate_all_fees(branches)
    
    # 测试统计功能
    test_statistics(branches)
    
    print("\n" + "=" * 60)
    print("所有测试通过！")
    print("=" * 60)


if __name__ == "__main__":
    main()
