# 测试改进后的系统功能
import os
import sys
import json

# 添加项目目录到Python路径
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from data_models import PartyMember, PartyBranch, SalaryInfo, DeductionInfo
from member_manager import MemberManager
from fee_calculator import FeeCalculator
from config import DATA_DIR


def test_data_persistence():
    """测试数据持久化功能"""
    print("=" * 60)
    print("测试数据持久化功能")
    print("=" * 60)
    
    # 创建测试数据
    print("\n1. 创建测试数据...")
    manager = MemberManager()
    
    # 添加支部
    branch1 = PartyBranch(name="测试支部1", sequence=1)
    branch2 = PartyBranch(name="测试支部2", sequence=2)
    
    manager.add_branch(branch1)
    manager.add_branch(branch2)
    
    # 添加党员
    member1 = PartyMember(
        name="测试党员1",
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
        name="测试党员2",
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
    
    # 计算党费
    FeeCalculator.calculate_member_fee(member1)
    FeeCalculator.calculate_member_fee(member2)
    
    # 添加到支部
    manager.add_member(member1, "测试支部1")
    manager.add_member(member2, "测试支部2")
    
    print(f"   - 创建了 {len(manager.get_all_branches())} 个党支部")
    print(f"   - 创建了 {len(manager.get_all_members())} 名党员")
    
    # 保存数据
    print("\n2. 保存数据到文件...")
    filepath = manager.save_to_file(2026, 4)
    print(f"   - 数据已保存到: {filepath}")
    
    # 验证文件是否存在
    if os.path.exists(filepath):
        print("   [OK] 文件保存成功！")
    else:
        print("   [FAIL] 文件保存失败！")
        return False
    
    # 读取文件内容验证
    print("\n3. 验证文件内容...")
    with open(filepath, 'r', encoding='utf-8') as f:
        data = json.load(f)
    
    print(f"   - 年份: {data.get('year')}")
    print(f"   - 月份: {data.get('month')}")
    print(f"   - 支部数量: {len(data.get('branches', []))}")
    
    # 加载数据到新的管理器
    print("\n4. 从文件加载数据...")
    new_manager = MemberManager()
    new_manager.load_from_file(2026, 4)
    
    branches = new_manager.get_all_branches()
    members = new_manager.get_all_members()
    
    print(f"   - 加载了 {len(branches)} 个党支部")
    print(f"   - 加载了 {len(members)} 名党员")
    
    # 验证数据完整性
    print("\n5. 验证数据完整性...")
    
    # 检查支部数量
    if len(branches) == 2:
        print("   [OK] 支部数量正确！")
    else:
        print("   [FAIL] 支部数量不正确！")
        return False
    
    # 检查党员数量
    if len(members) == 2:
        print("   [OK] 党员数量正确！")
    else:
        print("   [FAIL] 党员数量不正确！")
        return False
    
    # 检查党费计算
    for member in members:
        # 重新计算验证
        expected_base = FeeCalculator.calculate_payment_base(member)
        expected_fee = FeeCalculator.calculate_monthly_fee(expected_base)
        
        print(f"\n   党员: {member.name}")
        print(f"   - 存储的缴费基数: {member.payment_base:.2f}")
        print(f"   - 存储的月党费: {member.monthly_fee:.2f}")
        print(f"   - 重新计算的缴费基数: {expected_base:.2f}")
        print(f"   - 重新计算的月党费: {expected_fee:.2f}")
        
        if abs(member.payment_base - expected_base) < 0.01 and abs(member.monthly_fee - expected_fee) < 0.01:
            print("   [OK] 党费计算正确！")
        else:
            print("   [FAIL] 党费计算不正确！")
            return False
    
    print("\n" + "=" * 60)
    print("数据持久化测试通过！")
    print("=" * 60)
    return True


def test_fee_calculation_logic():
    """测试党费计算逻辑"""
    print("\n" + "=" * 60)
    print("测试党费计算逻辑")
    print("=" * 60)
    
    # 测试场景1：缴费基数 <= 3000元
    print("\n测试场景1：缴费基数 <= 3000元（0.5%费率）")
    
    # 创建党员对象
    member = PartyMember(
        name="测试党员",
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
    
    # 手动计算
    salary_total = 2185.0 + 692.0 + 0.0 + 1100.0  # 3977.0
    deduction_total = 1094.0 + 152.62 + 706.72 + 353.36 + 39.66 + 44.17 + 0.0  # 2390.53
    expected_base = salary_total - deduction_total  # 1586.47
    expected_fee = round(expected_base * 0.005, 1)  # 7.9
    
    print(f"   工资项目总和: {salary_total:.2f} 元")
    print(f"   扣款项目总和: {deduction_total:.2f} 元")
    print(f"   预期缴费基数: {expected_base:.2f} 元")
    print(f"   预期月党费: {expected_fee:.2f} 元")
    
    # 使用计算器计算
    FeeCalculator.calculate_member_fee(member)
    
    print(f"\n   实际缴费基数: {member.payment_base:.2f} 元")
    print(f"   实际月党费: {member.monthly_fee:.2f} 元")
    
    if abs(member.payment_base - expected_base) < 0.01 and abs(member.monthly_fee - expected_fee) < 0.01:
        print("   [OK] 计算结果正确！")
    else:
        print("   [FAIL] 计算结果不正确！")
        return False
    
    # 测试场景2：不同费率区间
    print("\n测试场景2：不同费率区间验证")
    
    test_cases = [
        (2000.0, 0.005, 10.0),   # 2000元 * 0.5% = 10元
        (3000.0, 0.005, 15.0),   # 3000元 * 0.5% = 15元
        (3500.0, 0.01, 35.0),    # 3500元 * 1% = 35元
        (5000.0, 0.01, 50.0),    # 5000元 * 1% = 50元
        (7500.0, 0.015, 112.5),  # 7500元 * 1.5% = 112.5元
        (10000.0, 0.015, 150.0), # 10000元 * 1.5% = 150元
        (15000.0, 0.02, 300.0),  # 15000元 * 2% = 300元
    ]
    
    for base, expected_rate, expected_fee in test_cases:
        # 创建党员对象（工资=缴费基数，扣款=0）
        member = PartyMember(
            name="测试",
            salary_info=SalaryInfo(
                position_salary=base,
                rank_salary=0.0,
                fixed_salary=0.0,
                basic_performance=0.0
            ),
            deduction_info=DeductionInfo()
        )
        
        FeeCalculator.calculate_member_fee(member)
        
        if abs(member.monthly_fee - expected_fee) < 0.01:
            print(f"   [OK] 缴费基数 {base:.0f} 元: 月党费 {member.monthly_fee:.1f} 元（预期: {expected_fee:.1f} 元）")
        else:
            print(f"   [FAIL] 缴费基数 {base:.0f} 元: 月党费 {member.monthly_fee:.1f} 元（预期: {expected_fee:.1f} 元）")
            return False
    
    print("\n" + "=" * 60)
    print("党费计算逻辑测试通过！")
    print("=" * 60)
    return True


def main():
    """主测试函数"""
    print("\n改进功能测试")
    print("=" * 60)
    
    # 测试1：数据持久化
    test1_passed = test_data_persistence()
    
    # 测试2：党费计算逻辑
    test2_passed = test_fee_calculation_logic()
    
    # 总结
    print("\n" + "=" * 60)
    print("测试总结")
    print("=" * 60)
    
    if test1_passed and test2_passed:
        print("\n[成功] 所有测试通过！")
        print("\n系统改进说明：")
        print("  1. 系统启动时自动检测并加载历史数据")
        print("  2. 添加党员时提供确认和修改功能")
        print("  3. 添加支部时提供确认和修改功能")
        print("  4. 党费计算逻辑正确实现")
    else:
        print("\n[失败] 部分测试未通过！")
    
    print("=" * 60)


if __name__ == "__main__":
    main()
