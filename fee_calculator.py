# 党费计算模块
from typing import List

from config import PARTY_FEE_RATES
from data_models import PartyMember, PartyBranch, SalaryInfo, DeductionInfo


class FeeCalculator:
    """党费计算器"""
    
    @staticmethod
    def calculate_payment_base(member: PartyMember) -> float:
        """
        计算缴费基数
        
        缴费基数 = 工资项目总和 - 扣款项目总和
        工资项目包括：岗位工资、薪级工资、高定工资、基础性绩效
        扣款项目包括：住房公积金、医疗保险、养老保险、职业年金、大额医疗、失业保险、个人所得税
        """
        # 计算工资项目总和
        salary_total = (
            member.salary_info.position_salary +
            member.salary_info.rank_salary +
            member.salary_info.fixed_salary +
            member.salary_info.basic_performance
        )
        
        # 计算扣款项目总和
        deduction_total = (
            member.deduction_info.housing_fund +
            member.deduction_info.medical_insurance +
            member.deduction_info.pension_insurance +
            member.deduction_info.occupational_annuity +
            member.deduction_info.large_medical +
            member.deduction_info.unemployment_insurance +
            member.deduction_info.personal_income_tax
        )
        
        # 计算缴费基数
        payment_base = salary_total - deduction_total
        
        # 确保缴费基数不为负数
        return max(0.0, payment_base)
    
    @staticmethod
    def calculate_monthly_fee(payment_base: float) -> float:
        """
        根据缴费基数计算每月应缴党费
        
        计算规则：
        - 3000元以下(含3000元)：0.5%
        - 3000元以上至5000元(含5000元)：1%
        - 5000元以上至10000元(含10000元)：1.5%
        - 10000元以上：2%
        """
        if payment_base <= 0:
            return 0.0
        
        # 根据缴费基数确定费率
        if payment_base <= 3000:
            rate = PARTY_FEE_RATES["below_3000"]
        elif payment_base <= 5000:
            rate = PARTY_FEE_RATES["3000_5000"]
        elif payment_base <= 10000:
            rate = PARTY_FEE_RATES["5000_10000"]
        else:
            rate = PARTY_FEE_RATES["above_10000"]
        
        # 计算党费
        monthly_fee = payment_base * rate
        
        # 四舍五入到小数点后一位
        return round(monthly_fee, 1)
    
    @staticmethod
    def calculate_member_fee(member: PartyMember, force: bool = False) -> PartyMember:
        """
        计算单个党员的党费
        
        Args:
            member: 党员对象
            force: 是否强制重新计算，即使已有值
        """
        # 只有当force为True，或者payment_base和monthly_fee都为0时，才进行计算
        if force or (member.payment_base == 0 and member.monthly_fee == 0):
            # 计算缴费基数
            member.payment_base = FeeCalculator.calculate_payment_base(member)
            
            # 计算每月应缴党费
            member.monthly_fee = FeeCalculator.calculate_monthly_fee(member.payment_base)
        
        return member
    
    @staticmethod
    def calculate_branch_fees(branch: PartyBranch, force: bool = False) -> PartyBranch:
        """
        计算整个支部所有党员的党费
        
        Args:
            branch: 支部对象
            force: 是否强制重新计算，即使已有值
        """
        for member in branch.members:
            FeeCalculator.calculate_member_fee(member, force)
        
        return branch
    
    @staticmethod
    def calculate_all_fees(branches: List[PartyBranch], force: bool = False) -> List[PartyBranch]:
        """
        计算所有支部所有党员的党费
        
        Args:
            branches: 支部列表
            force: 是否强制重新计算，即使已有值
        """
        for branch in branches:
            FeeCalculator.calculate_branch_fees(branch, force)
        
        return branches
    
    @staticmethod
    def get_branch_total_fee(branch: PartyBranch) -> float:
        """
        获取支部总党费
        """
        return sum(member.monthly_fee for member in branch.members)
    
    @staticmethod
    def get_branches_total_fee(branches: List[PartyBranch]) -> float:
        """
        获取所有支部总党费
        """
        total = 0.0
        for branch in branches:
            total += FeeCalculator.get_branch_total_fee(branch)
        return total
    
    @staticmethod
    def get_branch_member_count(branch: PartyBranch) -> int:
        """
        获取支部党员人数
        """
        return len(branch.members)
    
    @staticmethod
    def get_branches_total_member_count(branches: List[PartyBranch]) -> int:
        """
        获取所有支部总党员人数
        """
        total = 0
        for branch in branches:
            total += FeeCalculator.get_branch_member_count(branch)
        return total
