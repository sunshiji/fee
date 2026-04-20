# SQLite数据库管理模块
import sqlite3
import os
import json
from typing import List, Optional, Dict, Any, Tuple
from dataclasses import dataclass, asdict
from datetime import datetime

from config import BASE_DIR
from data_models import PartyMember, PartyBranch, SalaryInfo, DeductionInfo


@dataclass
class DatabaseConfig:
    """数据库配置"""
    db_path: str = None
    
    def __post_init__(self):
        if self.db_path is None:
            db_dir = os.path.join(BASE_DIR, "data")
            if not os.path.exists(db_dir):
                os.makedirs(db_dir)
            self.db_path = os.path.join(db_dir, "party_fee.db")


class DatabaseManager:
    """数据库管理器"""
    
    def __init__(self, config: DatabaseConfig = None):
        """初始化数据库管理器"""
        self.config = config or DatabaseConfig()
        self._init_database()
    
    def _init_database(self):
        """初始化数据库表结构"""
        conn = sqlite3.connect(self.config.db_path)
        cursor = conn.cursor()
        
        # 创建月度数据表
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS monthly_data (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                year INTEGER NOT NULL,
                month INTEGER NOT NULL,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL,
                UNIQUE(year, month)
            )
        ''')
        
        # 创建党支部表
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS branches (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                monthly_data_id INTEGER NOT NULL,
                branch_name TEXT NOT NULL,
                branch_sequence INTEGER NOT NULL,
                FOREIGN KEY (monthly_data_id) REFERENCES monthly_data(id),
                UNIQUE(monthly_data_id, branch_name)
            )
        ''')
        
        # 创建党员表
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS members (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                branch_id INTEGER NOT NULL,
                member_name TEXT NOT NULL,
                member_sequence INTEGER NOT NULL,
                -- 工资信息
                position_salary REAL DEFAULT 0.0,
                rank_salary REAL DEFAULT 0.0,
                fixed_salary REAL DEFAULT 0.0,
                basic_performance REAL DEFAULT 0.0,
                -- 扣款信息
                housing_fund REAL DEFAULT 0.0,
                medical_insurance REAL DEFAULT 0.0,
                pension_insurance REAL DEFAULT 0.0,
                occupational_annuity REAL DEFAULT 0.0,
                large_medical REAL DEFAULT 0.0,
                unemployment_insurance REAL DEFAULT 0.0,
                personal_income_tax REAL DEFAULT 0.0,
                -- 计算结果
                payment_base REAL DEFAULT 0.0,
                monthly_fee REAL DEFAULT 0.0,
                FOREIGN KEY (branch_id) REFERENCES branches(id),
                UNIQUE(branch_id, member_name)
            )
        ''')
        
        # 创建索引
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_monthly_data_year_month ON monthly_data(year, month)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_branches_monthly_data_id ON branches(monthly_data_id)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_members_branch_id ON members(branch_id)')
        
        conn.commit()
        conn.close()
    
    def save_monthly_data(self, year: int, month: int, branches: List[PartyBranch]) -> bool:
        """
        保存月度数据到数据库
        
        Args:
            year: 年份
            month: 月份
            branches: 党支部列表
            
        Returns:
            是否保存成功
        """
        conn = sqlite3.connect(self.config.db_path)
        cursor = conn.cursor()
        
        try:
            now = datetime.now().isoformat()
            
            # 检查是否已存在该月份的数据
            cursor.execute('SELECT id FROM monthly_data WHERE year = ? AND month = ?', (year, month))
            existing = cursor.fetchone()
            
            if existing:
                # 更新现有数据
                monthly_data_id = existing[0]
                cursor.execute('UPDATE monthly_data SET updated_at = ? WHERE id = ?', (now, monthly_data_id))
                
                # 删除旧的支部和党员数据
                cursor.execute('DELETE FROM members WHERE branch_id IN (SELECT id FROM branches WHERE monthly_data_id = ?)', (monthly_data_id,))
                cursor.execute('DELETE FROM branches WHERE monthly_data_id = ?', (monthly_data_id,))
            else:
                # 插入新的月度数据
                cursor.execute('''
                    INSERT INTO monthly_data (year, month, created_at, updated_at)
                    VALUES (?, ?, ?, ?)
                ''', (year, month, now, now))
                monthly_data_id = cursor.lastrowid
            
            # 插入支部数据
            for branch in branches:
                cursor.execute('''
                    INSERT INTO branches (monthly_data_id, branch_name, branch_sequence)
                    VALUES (?, ?, ?)
                ''', (monthly_data_id, branch.name, branch.sequence))
                branch_id = cursor.lastrowid
                
                # 插入党员数据
                for member in branch.members:
                    cursor.execute('''
                        INSERT INTO members (
                            branch_id, member_name, member_sequence,
                            position_salary, rank_salary, fixed_salary, basic_performance,
                            housing_fund, medical_insurance, pension_insurance, occupational_annuity,
                            large_medical, unemployment_insurance, personal_income_tax,
                            payment_base, monthly_fee
                        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    ''', (
                        branch_id, member.name, member.sequence,
                        member.salary_info.position_salary,
                        member.salary_info.rank_salary,
                        member.salary_info.fixed_salary,
                        member.salary_info.basic_performance,
                        member.deduction_info.housing_fund,
                        member.deduction_info.medical_insurance,
                        member.deduction_info.pension_insurance,
                        member.deduction_info.occupational_annuity,
                        member.deduction_info.large_medical,
                        member.deduction_info.unemployment_insurance,
                        member.deduction_info.personal_income_tax,
                        member.payment_base,
                        member.monthly_fee
                    ))
            
            conn.commit()
            return True
            
        except Exception as e:
            conn.rollback()
            print(f"数据库保存失败: {e}")
            return False
        finally:
            conn.close()
    
    def load_monthly_data(self, year: int, month: int) -> Optional[List[PartyBranch]]:
        """
        从数据库加载月度数据
        
        Args:
            year: 年份
            month: 月份
            
        Returns:
            党支部列表，如果不存在返回None
        """
        conn = sqlite3.connect(self.config.db_path)
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()
        
        try:
            # 查询月度数据
            cursor.execute('SELECT id FROM monthly_data WHERE year = ? AND month = ?', (year, month))
            monthly_data_row = cursor.fetchone()
            
            if not monthly_data_row:
                conn.close()
                return None
            
            monthly_data_id = monthly_data_row['id']
            
            # 查询所有支部
            cursor.execute('''
                SELECT id, branch_name, branch_sequence
                FROM branches
                WHERE monthly_data_id = ?
                ORDER BY branch_sequence
            ''', (monthly_data_id,))
            
            branches = []
            for branch_row in cursor.fetchall():
                branch = PartyBranch(
                    name=branch_row['branch_name'],
                    sequence=branch_row['branch_sequence']
                )
                
                # 查询该支部的党员
                cursor.execute('''
                    SELECT 
                        member_name, member_sequence,
                        position_salary, rank_salary, fixed_salary, basic_performance,
                        housing_fund, medical_insurance, pension_insurance, occupational_annuity,
                        large_medical, unemployment_insurance, personal_income_tax,
                        payment_base, monthly_fee
                    FROM members
                    WHERE branch_id = ?
                    ORDER BY member_sequence
                ''', (branch_row['id'],))
                
                for member_row in cursor.fetchall():
                    salary_info = SalaryInfo(
                        position_salary=member_row['position_salary'],
                        rank_salary=member_row['rank_salary'],
                        fixed_salary=member_row['fixed_salary'],
                        basic_performance=member_row['basic_performance']
                    )
                    
                    deduction_info = DeductionInfo(
                        housing_fund=member_row['housing_fund'],
                        medical_insurance=member_row['medical_insurance'],
                        pension_insurance=member_row['pension_insurance'],
                        occupational_annuity=member_row['occupational_annuity'],
                        large_medical=member_row['large_medical'],
                        unemployment_insurance=member_row['unemployment_insurance'],
                        personal_income_tax=member_row['personal_income_tax']
                    )
                    
                    member = PartyMember(
                        name=member_row['member_name'],
                        sequence=member_row['member_sequence'],
                        branch_name=branch.name,
                        branch_sequence=branch.sequence,
                        salary_info=salary_info,
                        deduction_info=deduction_info,
                        payment_base=member_row['payment_base'],
                        monthly_fee=member_row['monthly_fee']
                    )
                    
                    branch.members.append(member)
                
                branches.append(branch)
            
            conn.close()
            return branches
            
        except Exception as e:
            conn.close()
            print(f"数据库加载失败: {e}")
            return None
    
    def get_available_months(self) -> List[Tuple[int, int]]:
        """
        获取所有可用的月份数据
        
        Returns:
            [(年份, 月份), ...] 列表，按时间倒序排列
        """
        conn = sqlite3.connect(self.config.db_path)
        cursor = conn.cursor()
        
        try:
            cursor.execute('''
                SELECT year, month
                FROM monthly_data
                ORDER BY year DESC, month DESC
            ''')
            
            months = [(row[0], row[1]) for row in cursor.fetchall()]
            conn.close()
            return months
            
        except Exception as e:
            conn.close()
            print(f"获取可用月份失败: {e}")
            return []
    
    def delete_monthly_data(self, year: int, month: int) -> bool:
        """
        删除指定月份的数据
        
        Args:
            year: 年份
            month: 月份
            
        Returns:
            是否删除成功
        """
        conn = sqlite3.connect(self.config.db_path)
        cursor = conn.cursor()
        
        try:
            # 查询月度数据ID
            cursor.execute('SELECT id FROM monthly_data WHERE year = ? AND month = ?', (year, month))
            monthly_data_row = cursor.fetchone()
            
            if not monthly_data_row:
                conn.close()
                return False
            
            monthly_data_id = monthly_data_row[0]
            
            # 删除党员数据
            cursor.execute('DELETE FROM members WHERE branch_id IN (SELECT id FROM branches WHERE monthly_data_id = ?)', (monthly_data_id,))
            
            # 删除支部数据
            cursor.execute('DELETE FROM branches WHERE monthly_data_id = ?', (monthly_data_id,))
            
            # 删除月度数据
            cursor.execute('DELETE FROM monthly_data WHERE id = ?', (monthly_data_id,))
            
            conn.commit()
            conn.close()
            return True
            
        except Exception as e:
            conn.rollback()
            conn.close()
            print(f"删除数据失败: {e}")
            return False


@dataclass
class ComparisonResult:
    """对比结果"""
    base_year: int
    base_month: int
    compare_year: int
    compare_month: int
    
    # 总体对比
    base_total_members: int = 0
    compare_total_members: int = 0
    member_change: int = 0
    member_change_percent: float = 0.0
    
    base_total_fee: float = 0.0
    compare_total_fee: float = 0.0
    fee_change: float = 0.0
    fee_change_percent: float = 0.0
    
    # 支部对比
    branch_comparisons: List[Dict] = None
    
    # 新增/减少的党员
    new_members: List[Dict] = None
    removed_members: List[Dict] = None
    
    # 党员变化
    member_changes: List[Dict] = None
    
    def __post_init__(self):
        if self.branch_comparisons is None:
            self.branch_comparisons = []
        if self.new_members is None:
            self.new_members = []
        if self.removed_members is None:
            self.removed_members = []
        if self.member_changes is None:
            self.member_changes = []


class DataComparator:
    """数据对比分析器"""
    
    def __init__(self, db_manager: DatabaseManager = None):
        """初始化对比分析器"""
        self.db_manager = db_manager or DatabaseManager()
    
    def compare_months(self, base_year: int, base_month: int, 
                       compare_year: int, compare_month: int) -> Optional[ComparisonResult]:
        """
        对比两个月份的数据
        
        Args:
            base_year: 基准年份
            base_month: 基准月份
            compare_year: 对比年份
            compare_month: 对比月份
            
        Returns:
            对比结果
        """
        # 加载两个月份的数据
        base_branches = self.db_manager.load_monthly_data(base_year, base_month)
        compare_branches = self.db_manager.load_monthly_data(compare_year, compare_month)
        
        if base_branches is None or compare_branches is None:
            return None
        
        result = ComparisonResult(
            base_year=base_year,
            base_month=base_month,
            compare_year=compare_year,
            compare_month=compare_month
        )
        
        # 计算总体统计
        result.base_total_members = sum(len(b.members) for b in base_branches)
        result.compare_total_members = sum(len(b.members) for b in compare_branches)
        result.member_change = result.compare_total_members - result.base_total_members
        if result.base_total_members > 0:
            result.member_change_percent = (result.member_change / result.base_total_members) * 100
        
        result.base_total_fee = sum(sum(m.monthly_fee for m in b.members) for b in base_branches)
        result.compare_total_fee = sum(sum(m.monthly_fee for m in b.members) for b in compare_branches)
        result.fee_change = result.compare_total_fee - result.base_total_fee
        if result.base_total_fee > 0:
            result.fee_change_percent = (result.fee_change / result.base_total_fee) * 100
        
        # 构建查找字典
        base_members_dict = {}
        for branch in base_branches:
            for member in branch.members:
                key = f"{member.name}_{branch.name}"
                base_members_dict[key] = {
                    'name': member.name,
                    'branch': branch.name,
                    'member': member
                }
        
        compare_members_dict = {}
        for branch in compare_branches:
            for member in branch.members:
                key = f"{member.name}_{branch.name}"
                compare_members_dict[key] = {
                    'name': member.name,
                    'branch': branch.name,
                    'member': member
                }
        
        # 查找新增党员
        for key, data in compare_members_dict.items():
            if key not in base_members_dict:
                result.new_members.append({
                    'name': data['name'],
                    'branch': data['branch'],
                    'payment_base': data['member'].payment_base,
                    'monthly_fee': data['member'].monthly_fee
                })
        
        # 查找减少的党员
        for key, data in base_members_dict.items():
            if key not in compare_members_dict:
                result.removed_members.append({
                    'name': data['name'],
                    'branch': data['branch'],
                    'payment_base': data['member'].payment_base,
                    'monthly_fee': data['member'].monthly_fee
                })
        
        # 查找有变化的党员
        for key, base_data in base_members_dict.items():
            if key in compare_members_dict:
                compare_data = compare_members_dict[key]
                base_member = base_data['member']
                compare_member = compare_data['member']
                
                # 检查是否有变化
                has_change = False
                changes = []
                
                if base_member.payment_base != compare_member.payment_base:
                    has_change = True
                    changes.append({
                        'field': '缴费基数',
                        'old': base_member.payment_base,
                        'new': compare_member.payment_base,
                        'change': compare_member.payment_base - base_member.payment_base
                    })
                
                if base_member.monthly_fee != compare_member.monthly_fee:
                    has_change = True
                    changes.append({
                        'field': '月党费',
                        'old': base_member.monthly_fee,
                        'new': compare_member.monthly_fee,
                        'change': compare_member.monthly_fee - base_member.monthly_fee
                    })
                
                if has_change:
                    result.member_changes.append({
                        'name': base_member.name,
                        'branch': base_member.branch_name,
                        'changes': changes
                    })
        
        # 支部对比
        base_branches_dict = {b.name: b for b in base_branches}
        compare_branches_dict = {b.name: b for b in compare_branches}
        
        all_branch_names = set(base_branches_dict.keys()) | set(compare_branches_dict.keys())
        
        for branch_name in sorted(all_branch_names):
            base_branch = base_branches_dict.get(branch_name)
            compare_branch = compare_branches_dict.get(branch_name)
            
            base_members_count = len(base_branch.members) if base_branch else 0
            compare_members_count = len(compare_branch.members) if compare_branch else 0
            
            base_fee = sum(m.monthly_fee for m in base_branch.members) if base_branch else 0.0
            compare_fee = sum(m.monthly_fee for m in compare_branch.members) if compare_branch else 0.0
            
            result.branch_comparisons.append({
                'branch_name': branch_name,
                'base_members': base_members_count,
                'compare_members': compare_members_count,
                'member_change': compare_members_count - base_members_count,
                'base_fee': base_fee,
                'compare_fee': compare_fee,
                'fee_change': compare_fee - base_fee,
                'fee_change_percent': ((compare_fee - base_fee) / base_fee * 100) if base_fee > 0 else 0.0
            })
        
        return result
    
    def generate_comparison_report(self, result: ComparisonResult) -> str:
        """
        生成对比报告文本
        
        Args:
            result: 对比结果
            
        Returns:
            报告文本
        """
        lines = []
        lines.append("=" * 80)
        lines.append(f"数据对比分析报告")
        lines.append(f"基准月份: {result.base_year}年{result.base_month}月")
        lines.append(f"对比月份: {result.compare_year}年{result.compare_month}月")
        lines.append("=" * 80)
        lines.append("")
        
        # 总体对比
        lines.append("【总体对比】")
        lines.append("-" * 80)
        lines.append(f"{'项目':<20}{'基准月份':<15}{'对比月份':<15}{'变化量':<15}{'变化率':<15}")
        lines.append("-" * 80)
        
        member_change_str = f"+{result.member_change}" if result.member_change >= 0 else f"{result.member_change}"
        member_percent_str = f"+{result.member_change_percent:.2f}%" if result.member_change_percent >= 0 else f"{result.member_change_percent:.2f}%"
        lines.append(f"{'党员人数':<20}{result.base_total_members:<15}{result.compare_total_members:<15}{member_change_str:<15}{member_percent_str:<15}")
        
        fee_change_str = f"+{result.fee_change:.2f}" if result.fee_change >= 0 else f"{result.fee_change:.2f}"
        fee_percent_str = f"+{result.fee_change_percent:.2f}%" if result.fee_change_percent >= 0 else f"{result.fee_change_percent:.2f}%"
        lines.append(f"{'总党费(元)':<20}{result.base_total_fee:.2f:<15}{result.compare_total_fee:.2f:<15}{fee_change_str:<15}{fee_percent_str:<15}")
        lines.append("")
        
        # 支部对比
        lines.append("【各支部对比】")
        lines.append("-" * 80)
        lines.append(f"{'支部名称':<30}{'党员人数变化':<15}{'党费变化(元)':<15}{'变化率':<15}")
        lines.append("-" * 80)
        
        for bc in result.branch_comparisons:
            member_change = f"+{bc['member_change']}" if bc['member_change'] >= 0 else f"{bc['member_change']}"
            fee_change = f"+{bc['fee_change']:.2f}" if bc['fee_change'] >= 0 else f"{bc['fee_change']:.2f}"
            percent = f"+{bc['fee_change_percent']:.2f}%" if bc['fee_change_percent'] >= 0 else f"{bc['fee_change_percent']:.2f}%"
            
            lines.append(f"{bc['branch_name']:<30}{member_change:<15}{fee_change:<15}{percent:<15}")
        lines.append("")
        
        # 新增党员
        if result.new_members:
            lines.append(f"【新增党员】共 {len(result.new_members)} 人")
            lines.append("-" * 80)
            lines.append(f"{'姓名':<15}{'所属支部':<30}{'缴费基数(元)':<15}{'月党费(元)':<15}")
            lines.append("-" * 80)
            for m in result.new_members:
                lines.append(f"{m['name']:<15}{m['branch']:<30}{m['payment_base']:.2f:<15}{m['monthly_fee']:.2f:<15}")
            lines.append("")
        
        # 减少党员
        if result.removed_members:
            lines.append(f"【减少党员】共 {len(result.removed_members)} 人")
            lines.append("-" * 80)
            lines.append(f"{'姓名':<15}{'所属支部':<30}{'缴费基数(元)':<15}{'月党费(元)':<15}")
            lines.append("-" * 80)
            for m in result.removed_members:
                lines.append(f"{m['name']:<15}{m['branch']:<30}{m['payment_base']:.2f:<15}{m['monthly_fee']:.2f:<15}")
            lines.append("")
        
        # 党员变化
        if result.member_changes:
            lines.append(f"【党员信息变化】共 {len(result.member_changes)} 人")
            lines.append("-" * 80)
            for mc in result.member_changes:
                lines.append(f"\n党员: {mc['name']} (支部: {mc['branch']})")
                for change in mc['changes']:
                    change_str = f"+{change['change']:.2f}" if change['change'] >= 0 else f"{change['change']:.2f}"
                    lines.append(f"  {change['field']}: {change['old']:.2f} -> {change['new']:.2f} ({change_str})")
            lines.append("")
        
        lines.append("=" * 80)
        
        return "\n".join(lines)
