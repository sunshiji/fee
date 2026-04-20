# 党费收缴统计模块
from typing import List, Dict, Optional
from dataclasses import dataclass

from data_models import PartyBranch, PartyMember
from fee_calculator import FeeCalculator


@dataclass
class BranchStatistics:
    """支部统计信息"""
    branch_name: str = ""
    branch_sequence: int = 0
    member_count: int = 0
    total_fee: float = 0.0
    avg_fee: float = 0.0
    max_fee: float = 0.0
    min_fee: float = 0.0
    payment_base_range: str = ""


@dataclass
class OverallStatistics:
    """总体统计信息"""
    total_branches: int = 0
    total_members: int = 0
    total_fee: float = 0.0
    avg_fee_per_member: float = 0.0
    avg_fee_per_branch: float = 0.0
    max_fee_branch: str = ""
    min_fee_branch: str = ""
    largest_branch: str = ""
    smallest_branch: str = ""


class FeeStatistics:
    """党费统计器"""
    
    @staticmethod
    def calculate_branch_statistics(branch: PartyBranch) -> BranchStatistics:
        """
        计算单个支部的统计信息
        """
        stats = BranchStatistics()
        stats.branch_name = branch.name
        stats.branch_sequence = branch.sequence
        stats.member_count = len(branch.members)
        
        if stats.member_count == 0:
            return stats
        
        # 计算总党费
        stats.total_fee = FeeCalculator.get_branch_total_fee(branch)
        
        # 计算平均党费
        stats.avg_fee = stats.total_fee / stats.member_count
        
        # 计算最高和最低党费
        fees = [member.monthly_fee for member in branch.members]
        stats.max_fee = max(fees)
        stats.min_fee = min(fees)
        
        # 计算缴费基数范围
        payment_bases = [member.payment_base for member in branch.members]
        if payment_bases:
            min_base = min(payment_bases)
            max_base = max(payment_bases)
            stats.payment_base_range = f"{min_base:.2f} - {max_base:.2f}"
        
        return stats
    
    @staticmethod
    def calculate_all_branch_statistics(branches: List[PartyBranch]) -> List[BranchStatistics]:
        """
        计算所有支部的统计信息
        """
        all_stats = []
        for branch in branches:
            stats = FeeStatistics.calculate_branch_statistics(branch)
            all_stats.append(stats)
        
        # 按支部序号排序
        all_stats.sort(key=lambda s: s.branch_sequence)
        return all_stats
    
    @staticmethod
    def calculate_overall_statistics(branches: List[PartyBranch]) -> OverallStatistics:
        """
        计算总体统计信息
        """
        stats = OverallStatistics()
        stats.total_branches = len(branches)
        
        if stats.total_branches == 0:
            return stats
        
        # 计算总党员人数
        stats.total_members = FeeCalculator.get_branches_total_member_count(branches)
        
        # 计算总党费
        stats.total_fee = FeeCalculator.get_branches_total_fee(branches)
        
        # 计算平均党费
        if stats.total_members > 0:
            stats.avg_fee_per_member = stats.total_fee / stats.total_members
        
        stats.avg_fee_per_branch = stats.total_fee / stats.total_branches
        
        # 找出党费最高和最低的支部
        branch_stats = FeeStatistics.calculate_all_branch_statistics(branches)
        
        if branch_stats:
            # 按总党费排序
            sorted_by_fee = sorted(branch_stats, key=lambda s: s.total_fee, reverse=True)
            stats.max_fee_branch = sorted_by_fee[0].branch_name
            stats.min_fee_branch = sorted_by_fee[-1].branch_name
            
            # 按党员人数排序
            sorted_by_members = sorted(branch_stats, key=lambda s: s.member_count, reverse=True)
            stats.largest_branch = sorted_by_members[0].branch_name
            stats.smallest_branch = sorted_by_members[-1].branch_name
        
        return stats
    
    @staticmethod
    def generate_statistics_report(branches: List[PartyBranch], 
                                    year: int, month: int) -> str:
        """
        生成统计报告文本
        """
        report_lines = []
        report_lines.append("=" * 80)
        report_lines.append(f"{year}年{month}月党费收缴统计报告")
        report_lines.append("=" * 80)
        report_lines.append("")
        
        # 总体统计
        overall_stats = FeeStatistics.calculate_overall_statistics(branches)
        report_lines.append("【总体统计】")
        report_lines.append(f"  党支部数量: {overall_stats.total_branches}")
        report_lines.append(f"  党员总人数: {overall_stats.total_members}")
        report_lines.append(f"  党费总金额: {overall_stats.total_fee:.2f} 元")
        report_lines.append(f"  人均党费: {overall_stats.avg_fee_per_member:.2f} 元")
        report_lines.append(f"  支部平均党费: {overall_stats.avg_fee_per_branch:.2f} 元")
        report_lines.append(f"  党费最高支部: {overall_stats.max_fee_branch}")
        report_lines.append(f"  党费最低支部: {overall_stats.min_fee_branch}")
        report_lines.append(f"  人数最多支部: {overall_stats.largest_branch}")
        report_lines.append(f"  人数最少支部: {overall_stats.smallest_branch}")
        report_lines.append("")
        
        # 各支部统计
        report_lines.append("【各支部统计】")
        report_lines.append("-" * 80)
        report_lines.append(f"{'序号':<6}{'支部名称':<30}{'党员人数':<10}{'总党费(元)':<15}{'人均党费(元)':<15}{'缴费基数范围':<20}")
        report_lines.append("-" * 80)
        
        branch_stats = FeeStatistics.calculate_all_branch_statistics(branches)
        for i, stats in enumerate(branch_stats, 1):
            report_lines.append(
                f"{i:<6}{stats.branch_name:<30}{stats.member_count:<10}"
                f"{stats.total_fee:<15.2f}{stats.avg_fee:<15.2f}{stats.payment_base_range:<20}"
            )
        
        report_lines.append("-" * 80)
        report_lines.append("")
        report_lines.append("=" * 80)
        
        return "\n".join(report_lines)
    
    @staticmethod
    def print_statistics_report(branches: List[PartyBranch], 
                                 year: int, month: int) -> None:
        """
        打印统计报告
        """
        report = FeeStatistics.generate_statistics_report(branches, year, month)
        print(report)
    
    @staticmethod
    def get_fee_distribution(branches: List[PartyBranch]) -> Dict[str, int]:
        """
        获取党费分布情况（按党费金额区间统计人数）
        """
        distribution = {
            "0元以下": 0,
            "0-0.5元": 0,
            "0.5-1元": 0,
            "1-5元": 0,
            "5-10元": 0,
            "10-20元": 0,
            "20-50元": 0,
            "50-100元": 0,
            "100元以上": 0
        }
        
        for branch in branches:
            for member in branch.members:
                fee = member.monthly_fee
                if fee <= 0:
                    distribution["0元以下"] += 1
                elif fee <= 0.5:
                    distribution["0-0.5元"] += 1
                elif fee <= 1:
                    distribution["0.5-1元"] += 1
                elif fee <= 5:
                    distribution["1-5元"] += 1
                elif fee <= 10:
                    distribution["5-10元"] += 1
                elif fee <= 20:
                    distribution["10-20元"] += 1
                elif fee <= 50:
                    distribution["20-50元"] += 1
                elif fee <= 100:
                    distribution["50-100元"] += 1
                else:
                    distribution["100元以上"] += 1
        
        return distribution
    
    @staticmethod
    def get_payment_base_distribution(branches: List[PartyBranch]) -> Dict[str, int]:
        """
        获取缴费基数分布情况
        """
        distribution = {
            "0元以下": 0,
            "0-3000元": 0,
            "3000-5000元": 0,
            "5000-10000元": 0,
            "10000元以上": 0
        }
        
        for branch in branches:
            for member in branch.members:
                base = member.payment_base
                if base <= 0:
                    distribution["0元以下"] += 1
                elif base <= 3000:
                    distribution["0-3000元"] += 1
                elif base <= 5000:
                    distribution["3000-5000元"] += 1
                elif base <= 10000:
                    distribution["5000-10000元"] += 1
                else:
                    distribution["10000元以上"] += 1
        
        return distribution
