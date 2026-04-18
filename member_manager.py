# 人员信息管理模块
import json
import os
from typing import List, Optional
from datetime import datetime

from config import DATA_DIR
from data_models import PartyMember, PartyBranch, MonthlyFeeData, SalaryInfo, DeductionInfo


class MemberManager:
    """人员信息管理器"""
    
    def __init__(self):
        """初始化人员信息管理器"""
        self.branches: List[PartyBranch] = []
        self._ensure_data_dir()
    
    def _ensure_data_dir(self):
        """确保数据目录存在"""
        if not os.path.exists(DATA_DIR):
            os.makedirs(DATA_DIR)
    
    def add_branch(self, branch: PartyBranch) -> None:
        """添加党支部"""
        # 检查是否已存在同名支部
        for existing_branch in self.branches:
            if existing_branch.name == branch.name:
                raise ValueError(f"党支部 '{branch.name}' 已存在")
            if existing_branch.sequence == branch.sequence:
                raise ValueError(f"支部序号 '{branch.sequence}' 已存在")
        
        self.branches.append(branch)
        # 按支部序号排序
        self.branches.sort(key=lambda b: b.sequence)
    
    def remove_branch(self, branch_name: str) -> bool:
        """删除党支部"""
        for i, branch in enumerate(self.branches):
            if branch.name == branch_name:
                self.branches.pop(i)
                return True
        return False
    
    def get_branch(self, branch_name: str) -> Optional[PartyBranch]:
        """获取指定党支部"""
        for branch in self.branches:
            if branch.name == branch_name:
                return branch
        return None
    
    def get_all_branches(self) -> List[PartyBranch]:
        """获取所有党支部"""
        return self.branches
    
    def add_member(self, member: PartyMember, branch_name: str) -> bool:
        """添加党员到指定党支部"""
        branch = self.get_branch(branch_name)
        if not branch:
            raise ValueError(f"党支部 '{branch_name}' 不存在")
        
        # 检查党员是否已存在
        for existing_member in branch.members:
            if existing_member.name == member.name:
                raise ValueError(f"党员 '{member.name}' 已在支部 '{branch_name}' 中")
        
        # 为党员分配支部信息
        member.branch_name = branch_name
        member.branch_sequence = branch.sequence
        
        # 为党员分配序号
        if not branch.members:
            member.sequence = 1
        else:
            member.sequence = max(m.sequence for m in branch.members) + 1
        
        branch.members.append(member)
        # 按序号排序
        branch.members.sort(key=lambda m: m.sequence)
        return True
    
    def remove_member(self, member_name: str, branch_name: str) -> bool:
        """从指定党支部删除党员"""
        branch = self.get_branch(branch_name)
        if not branch:
            return False
        
        for i, member in enumerate(branch.members):
            if member.name == member_name:
                branch.members.pop(i)
                # 重新排序号
                for j, m in enumerate(branch.members):
                    m.sequence = j + 1
                return True
        return False
    
    def update_member(self, member_name: str, branch_name: str, 
                      salary_info: Optional[SalaryInfo] = None,
                      deduction_info: Optional[DeductionInfo] = None) -> bool:
        """更新党员信息"""
        branch = self.get_branch(branch_name)
        if not branch:
            return False
        
        for member in branch.members:
            if member.name == member_name:
                if salary_info:
                    member.salary_info = salary_info
                if deduction_info:
                    member.deduction_info = deduction_info
                return True
        return False
    
    def get_member(self, member_name: str, branch_name: Optional[str] = None) -> Optional[PartyMember]:
        """获取党员信息"""
        if branch_name:
            branch = self.get_branch(branch_name)
            if branch:
                for member in branch.members:
                    if member.name == member_name:
                        return member
        else:
            for branch in self.branches:
                for member in branch.members:
                    if member.name == member_name:
                        return member
        return None
    
    def get_all_members(self) -> List[PartyMember]:
        """获取所有党员"""
        all_members = []
        for branch in self.branches:
            all_members.extend(branch.members)
        return all_members
    
    def save_to_file(self, year: int, month: int) -> str:
        """保存数据到文件"""
        filename = f"party_fee_data_{year}_{month:02d}.json"
        filepath = os.path.join(DATA_DIR, filename)
        
        # 构建数据结构
        data = {
            "year": year,
            "month": month,
            "created_at": datetime.now().isoformat(),
            "branches": []
        }
        
        for branch in self.branches:
            branch_data = {
                "id": branch.id,
                "name": branch.name,
                "sequence": branch.sequence,
                "members": []
            }
            
            for member in branch.members:
                member_data = {
                    "id": member.id,
                    "name": member.name,
                    "branch_name": member.branch_name,
                    "branch_sequence": member.branch_sequence,
                    "sequence": member.sequence,
                    "salary_info": {
                        "position_salary": member.salary_info.position_salary,
                        "rank_salary": member.salary_info.rank_salary,
                        "fixed_salary": member.salary_info.fixed_salary,
                        "basic_performance": member.salary_info.basic_performance
                    },
                    "deduction_info": {
                        "housing_fund": member.deduction_info.housing_fund,
                        "medical_insurance": member.deduction_info.medical_insurance,
                        "pension_insurance": member.deduction_info.pension_insurance,
                        "occupational_annuity": member.deduction_info.occupational_annuity,
                        "large_medical": member.deduction_info.large_medical,
                        "unemployment_insurance": member.deduction_info.unemployment_insurance,
                        "personal_income_tax": member.deduction_info.personal_income_tax
                    },
                    "payment_base": member.payment_base,
                    "monthly_fee": member.monthly_fee
                }
                branch_data["members"].append(member_data)
            
            data["branches"].append(branch_data)
        
        # 保存到文件
        with open(filepath, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        
        return filepath
    
    def load_from_file(self, year: int, month: int) -> bool:
        """从文件加载数据"""
        filename = f"party_fee_data_{year}_{month:02d}.json"
        filepath = os.path.join(DATA_DIR, filename)
        
        if not os.path.exists(filepath):
            return False
        
        with open(filepath, 'r', encoding='utf-8') as f:
            data = json.load(f)
        
        # 清空现有数据
        self.branches = []
        
        # 加载支部数据
        for branch_data in data.get("branches", []):
            branch = PartyBranch(
                id=branch_data.get("id", ""),
                name=branch_data.get("name", ""),
                sequence=branch_data.get("sequence", 0)
            )
            
            # 加载党员数据
            for member_data in branch_data.get("members", []):
                salary_info = SalaryInfo(
                    position_salary=member_data.get("salary_info", {}).get("position_salary", 0.0),
                    rank_salary=member_data.get("salary_info", {}).get("rank_salary", 0.0),
                    fixed_salary=member_data.get("salary_info", {}).get("fixed_salary", 0.0),
                    basic_performance=member_data.get("salary_info", {}).get("basic_performance", 0.0)
                )
                
                deduction_info = DeductionInfo(
                    housing_fund=member_data.get("deduction_info", {}).get("housing_fund", 0.0),
                    medical_insurance=member_data.get("deduction_info", {}).get("medical_insurance", 0.0),
                    pension_insurance=member_data.get("deduction_info", {}).get("pension_insurance", 0.0),
                    occupational_annuity=member_data.get("deduction_info", {}).get("occupational_annuity", 0.0),
                    large_medical=member_data.get("deduction_info", {}).get("large_medical", 0.0),
                    unemployment_insurance=member_data.get("deduction_info", {}).get("unemployment_insurance", 0.0),
                    personal_income_tax=member_data.get("deduction_info", {}).get("personal_income_tax", 0.0)
                )
                
                member = PartyMember(
                    id=member_data.get("id", ""),
                    name=member_data.get("name", ""),
                    branch_name=member_data.get("branch_name", ""),
                    branch_sequence=member_data.get("branch_sequence", 0),
                    sequence=member_data.get("sequence", 0),
                    salary_info=salary_info,
                    deduction_info=deduction_info,
                    payment_base=member_data.get("payment_base", 0.0),
                    monthly_fee=member_data.get("monthly_fee", 0.0)
                )
                
                branch.members.append(member)
            
            self.branches.append(branch)
        
        return True
