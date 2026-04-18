# 测试党费计算修复的脚本
import os
import sys

# 添加项目目录到Python路径
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from data_models import PartyMember, PartyBranch, SalaryInfo, DeductionInfo
from member_manager import MemberManager
from fee_calculator import FeeCalculator


def test_fee_calculation_logic():
    """测试党费计算逻辑"""
    print("=" * 60)
    print("测试党费计算逻辑")
    print("=" * 60)
    
    # 测试场景1：张三的工资和扣款信息（根据测试数据）
    # 岗位工资: 2185.0, 薪级工资: 692.0, 高定工资: 0.0, 基础性绩效: 1100.0
    # 住房公积金: 1094.0, 医疗保险: 152.62, 养老保险: 706.72, 职业年金: 353.36
    # 大额医疗: 39.66, 失业保险: 44.17, 个人所得税: 0.0
    
    print("\n【测试场景1：张三的党费计算】")
    salary_total = 2185.0 + 692.0 + 0.0 + 1100.0
    deduction_total = 1094.0 + 152.62 + 706.72 + 353.36 + 39.66 + 44.17 + 0.0
    
    print(f"  工资项目总和: {salary_total:.2f} 元")
    print(f"  扣款项目总和: {deduction_total:.2f} 元")
    
    payment_base = salary_total - deduction_total
    print(f"  缴费基数: {payment_base:.2f} 元")
    
    # 计算党费（3000元以下按0.5%计算）
    if payment_base <= 3000:
        monthly_fee = payment_base * 0.005
        print(f"  党费计算: {payment_base:.2f} × 0.5% = {monthly_fee:.2f} 元")
    elif payment_base <= 5000:
        monthly_fee = payment_base * 0.01
        print(f"  党费计算: {payment_base:.2f} × 1% = {monthly_fee:.2f} 元")
    elif payment_base <= 10000:
        monthly_fee = payment_base * 0.015
        print(f"  党费计算: {payment_base:.2f} × 1.5% = {monthly_fee:.2f} 元")
    else:
        monthly_fee = payment_base * 0.02
        print(f"  党费计算: {payment_base:.2f} × 2% = {monthly_fee:.2f} 元")
    
    print(f"  四舍五入后: {round(monthly_fee, 1):.2f} 元")
    
    # 验证FeeCalculator的计算结果
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
    
    # 计算前的值
    print(f"\n  计算前: payment_base={member1.payment_base:.2f}, monthly_fee={member1.monthly_fee:.2f}")
    
    # 计算党费
    FeeCalculator.calculate_member_fee(member1)
    
    # 计算后的值
    print(f"  计算后: payment_base={member1.payment_base:.2f}, monthly_fee={member1.monthly_fee:.2f}")
    
    # 验证
    expected_payment_base = salary_total - deduction_total
    expected_monthly_fee = round(expected_payment_base * 0.005, 1)
    
    print(f"\n  预期结果: payment_base={expected_payment_base:.2f}, monthly_fee={expected_monthly_fee:.2f}")
    
    if abs(member1.payment_base - expected_payment_base) < 0.01 and abs(member1.monthly_fee - expected_monthly_fee) < 0.01:
        print("  [OK] 计算结果正确！")
    else:
        print("  [FAIL] 计算结果不正确！")
    
    # 测试场景2：李四的工资和扣款信息
    print("\n【测试场景2：李四的党费计算】")
    salary_total2 = 2375.0 + 845.0 + 0.0 + 2030.0
    deduction_total2 = 1200.0 + 165.78 + 785.62 + 392.81 + 45.32 + 52.43 + 67.57
    
    print(f"  工资项目总和: {salary_total2:.2f} 元")
    print(f"  扣款项目总和: {deduction_total2:.2f} 元")
    
    payment_base2 = salary_total2 - deduction_total2
    print(f"  缴费基数: {payment_base2:.2f} 元")
    
    # 计算党费
    if payment_base2 <= 3000:
        monthly_fee2 = payment_base2 * 0.005
        print(f"  党费计算: {payment_base2:.2f} × 0.5% = {monthly_fee2:.2f} 元")
    elif payment_base2 <= 5000:
        monthly_fee2 = payment_base2 * 0.01
        print(f"  党费计算: {payment_base2:.2f} × 1% = {monthly_fee2:.2f} 元")
    elif payment_base2 <= 10000:
        monthly_fee2 = payment_base2 * 0.015
        print(f"  党费计算: {payment_base2:.2f} × 1.5% = {monthly_fee2:.2f} 元")
    else:
        monthly_fee2 = payment_base2 * 0.02
        print(f"  党费计算: {payment_base2:.2f} × 2% = {monthly_fee2:.2f} 元")
    
    print(f"  四舍五入后: {round(monthly_fee2, 1):.2f} 元")
    
    # 验证FeeCalculator的计算结果
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
    
    FeeCalculator.calculate_member_fee(member2)
    
    expected_payment_base2 = salary_total2 - deduction_total2
    # 根据缴费基数确定费率
    if expected_payment_base2 <= 3000:
        expected_monthly_fee2 = round(expected_payment_base2 * 0.005, 1)
    elif expected_payment_base2 <= 5000:
        expected_monthly_fee2 = round(expected_payment_base2 * 0.01, 1)
    elif expected_payment_base2 <= 10000:
        expected_monthly_fee2 = round(expected_payment_base2 * 0.015, 1)
    else:
        expected_monthly_fee2 = round(expected_payment_base2 * 0.02, 1)
    
    print(f"\n  计算结果: payment_base={member2.payment_base:.2f}, monthly_fee={member2.monthly_fee:.2f}")
    print(f"  预期结果: payment_base={expected_payment_base2:.2f}, monthly_fee={expected_monthly_fee2:.2f}")
    
    if abs(member2.payment_base - expected_payment_base2) < 0.01 and abs(member2.monthly_fee - expected_monthly_fee2) < 0.01:
        print("  [OK] 计算结果正确！")
    else:
        print("  [FAIL] 计算结果不正确！")
    
    # 测试场景3：学生党员（工资为0）
    print("\n【测试场景3：学生党员（工资为0）】")
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
    
    print(f"  计算前: payment_base={member3.payment_base:.2f}, monthly_fee={member3.monthly_fee:.2f}")
    
    FeeCalculator.calculate_member_fee(member3)
    
    print(f"  计算后: payment_base={member3.payment_base:.2f}, monthly_fee={member3.monthly_fee:.2f}")
    
    if member3.payment_base == 0.0 and member3.monthly_fee == 0.0:
        print("  [OK] 计算结果正确（工资为0，党费为0）！")
    else:
        print("  [FAIL] 计算结果不正确！")


def test_member_manager_with_auto_calculation():
    """测试人员管理器的自动计算功能"""
    print("\n" + "=" * 60)
    print("测试人员管理器的自动计算功能")
    print("=" * 60)
    
    # 创建管理器
    manager = MemberManager()
    
    # 创建支部
    branch = PartyBranch(
        name="测试支部",
        sequence=1
    )
    manager.add_branch(branch)
    
    # 创建党员
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
    
    print("\n【测试1：添加党员前】")
    print(f"  payment_base={member.payment_base:.2f}, monthly_fee={member.monthly_fee:.2f}")
    
    # 添加党员
    manager.add_member(member, "测试支部")
    
    # 注意：这里需要手动计算，因为member_manager.add_member方法不会自动计算
    # 这个功能是在main.py的add_member方法中实现的
    
    print("\n【测试2：添加党员后（未计算）】")
    print(f"  payment_base={member.payment_base:.2f}, monthly_fee={member.monthly_fee:.2f}")
    
    # 手动计算
    FeeCalculator.calculate_member_fee(member)
    
    print("\n【测试3：手动计算后】")
    print(f"  payment_base={member.payment_base:.2f}, monthly_fee={member.monthly_fee:.2f}")
    
    # 验证
    salary_total = 2185.0 + 692.0 + 0.0 + 1100.0
    deduction_total = 1094.0 + 152.62 + 706.72 + 353.36 + 39.66 + 44.17 + 0.0
    expected_payment_base = salary_total - deduction_total
    expected_monthly_fee = round(expected_payment_base * 0.005, 1)
    
    print(f"\n  预期结果: payment_base={expected_payment_base:.2f}, monthly_fee={expected_monthly_fee:.2f}")
    
    if abs(member.payment_base - expected_payment_base) < 0.01 and abs(member.monthly_fee - expected_monthly_fee) < 0.01:
        print("  [OK] 计算结果正确！")
    else:
        print("  [FAIL] 计算结果不正确！")


def main():
    """主测试函数"""
    print("\n党费计算修复测试")
    print("=" * 60)
    
    # 测试党费计算逻辑
    test_fee_calculation_logic()
    
    # 测试人员管理器
    test_member_manager_with_auto_calculation()
    
    print("\n" + "=" * 60)
    print("测试完成！")
    print("=" * 60)


if __name__ == "__main__":
    main()
