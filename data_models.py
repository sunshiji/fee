# 数据模型定义
from dataclasses import dataclass, field
from typing import List, Optional
from datetime import datetime


@dataclass
class SalaryInfo:
    """工资信息"""
    position_salary: float = 0.0  # 岗位工资
    rank_salary: float = 0.0       # 薪级工资
    fixed_salary: float = 0.0      # 高定工资
    basic_performance: float = 0.0  # 基础性绩效


@dataclass
class DeductionInfo:
    """扣款信息"""
    housing_fund: float = 0.0       # 住房公积金
    medical_insurance: float = 0.0  # 医疗保险
    pension_insurance: float = 0.0   # 养老保险
    occupational_annuity: float = 0.0  # 职业年金
    large_medical: float = 0.0       # 大额医疗
    unemployment_insurance: float = 0.0  # 失业保险
    personal_income_tax: float = 0.0    # 个人所得税


@dataclass
class PartyMember:
    """党员信息"""
    id: str = ""  # 唯一标识
    name: str = ""  # 姓名
    branch_name: str = ""  # 所属党支部
    branch_sequence: int = 0  # 支部序号
    sequence: int = 0  # 序号
    salary_info: SalaryInfo = field(default_factory=SalaryInfo)
    deduction_info: DeductionInfo = field(default_factory=DeductionInfo)
    payment_base: float = 0.0  # 缴费基数
    monthly_fee: float = 0.0   # 每月应缴党费


@dataclass
class PartyBranch:
    """党支部信息"""
    id: str = ""  # 唯一标识
    name: str = ""  # 支部名称
    sequence: int = 0  # 支部序号
    members: List[PartyMember] = field(default_factory=list)
    
    @property
    def member_count(self) -> int:
        """党员人数"""
        return len(self.members)
    
    @property
    def total_fee(self) -> float:
        """总党费"""
        return sum(member.monthly_fee for member in self.members)


@dataclass
class MonthlyFeeData:
    """月度党费数据"""
    year: int = 0
    month: int = 0
    branches: List[PartyBranch] = field(default_factory=list)
    created_at: datetime = field(default_factory=datetime.now)
    
    @property
    def total_members(self) -> int:
        """总党员人数"""
        return sum(branch.member_count for branch in self.branches)
    
    @property
    def total_fee(self) -> float:
        """总党费"""
        return sum(branch.total_fee for branch in self.branches)
