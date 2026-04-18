# 党费收取系统主程序
import os
import sys
from typing import List, Optional

from config import DEFAULT_YEAR, DEFAULT_MONTH, OUTPUT_DIR
from data_models import PartyMember, PartyBranch, SalaryInfo, DeductionInfo
from member_manager import MemberManager
from fee_calculator import FeeCalculator
from excel_generator import ExcelGenerator
from statistics import FeeStatistics


class PartyFeeSystem:
    """党费收取系统主类"""
    
    def __init__(self):
        """初始化系统"""
        self.member_manager = MemberManager()
        self.excel_generator = ExcelGenerator()
        self.current_year = DEFAULT_YEAR
        self.current_month = DEFAULT_MONTH
    
    def display_menu(self):
        """显示主菜单"""
        print("\n" + "=" * 60)
        print("党费收取系统 v1.0")
        print("=" * 60)
        print(f"当前月份: {self.current_year}年{self.current_month}月")
        print("-" * 60)
        print("1. 人员维护管理")
        print("2. 支部维护管理")
        print("3. 计算党费")
        print("4. 生成党费明细表")
        print("5. 生成各支部党费收缴子表")
        print("6. 生成党费汇总表")
        print("7. 生成所有报表")
        print("8. 查看统计信息")
        print("9. 保存数据")
        print("10. 加载数据")
        print("11. 设置年月")
        print("0. 退出系统")
        print("=" * 60)
    
    def display_person_management_menu(self):
        """显示人员维护菜单"""
        print("\n" + "-" * 60)
        print("人员维护管理")
        print("-" * 60)
        print("1. 添加党员")
        print("2. 删除党员")
        print("3. 修改党员信息")
        print("4. 查看党员列表")
        print("5. 查看党员详情")
        print("0. 返回主菜单")
        print("-" * 60)
    
    def display_branch_management_menu(self):
        """显示支部维护菜单"""
        print("\n" + "-" * 60)
        print("支部维护管理")
        print("-" * 60)
        print("1. 添加党支部")
        print("2. 删除党支部")
        print("3. 查看党支部列表")
        print("4. 查看党支部详情")
        print("0. 返回主菜单")
        print("-" * 60)
    
    def input_float(self, prompt: str, default: float = 0.0) -> float:
        """输入浮点数"""
        while True:
            try:
                value = input(f"{prompt} (默认: {default}): ")
                if value.strip() == "":
                    return default
                return float(value)
            except ValueError:
                print("请输入有效的数字！")
    
    def input_int(self, prompt: str, default: int = 0) -> int:
        """输入整数"""
        while True:
            try:
                value = input(f"{prompt} (默认: {default}): ")
                if value.strip() == "":
                    return default
                return int(value)
            except ValueError:
                print("请输入有效的整数！")
    
    def add_member(self):
        """添加党员"""
        print("\n添加党员")
        print("-" * 40)
        
        # 显示支部列表供选择
        branches = self.member_manager.get_all_branches()
        if not branches:
            print("暂无党支部，请先添加党支部！")
            return
        
        print("\n请选择党支部：")
        for i, branch in enumerate(branches, 1):
            print(f"{i}. {branch.name}")
        
        while True:
            choice = self.input_int("请输入选项编号", 0)
            if 1 <= choice <= len(branches):
                selected_branch = branches[choice - 1]
                break
            else:
                print("无效的选项，请重新输入！")
        
        # 输入党员基本信息
        name = input("请输入党员姓名: ").strip()
        if not name:
            print("姓名不能为空！")
            return
        
        # 检查党员是否已存在
        existing_member = self.member_manager.get_member(name, selected_branch.name)
        if existing_member:
            print(f"党员 '{name}' 已在支部 '{selected_branch.name}' 中！")
            return
        
        # 输入工资信息
        print("\n请输入工资信息（单位：元）：")
        position_salary = self.input_float("岗位工资", 0.0)
        rank_salary = self.input_float("薪级工资", 0.0)
        fixed_salary = self.input_float("高定工资", 0.0)
        basic_performance = self.input_float("基础性绩效", 0.0)
        
        # 输入扣款信息
        print("\n请输入扣款信息（单位：元）：")
        housing_fund = self.input_float("住房公积金", 0.0)
        medical_insurance = self.input_float("医疗保险", 0.0)
        pension_insurance = self.input_float("养老保险", 0.0)
        occupational_annuity = self.input_float("职业年金", 0.0)
        large_medical = self.input_float("大额医疗", 0.0)
        unemployment_insurance = self.input_float("失业保险", 0.0)
        personal_income_tax = self.input_float("个人所得税", 0.0)
        
        # 创建党员对象
        salary_info = SalaryInfo(
            position_salary=position_salary,
            rank_salary=rank_salary,
            fixed_salary=fixed_salary,
            basic_performance=basic_performance
        )
        
        deduction_info = DeductionInfo(
            housing_fund=housing_fund,
            medical_insurance=medical_insurance,
            pension_insurance=pension_insurance,
            occupational_annuity=occupational_annuity,
            large_medical=large_medical,
            unemployment_insurance=unemployment_insurance,
            personal_income_tax=personal_income_tax
        )
        
        member = PartyMember(
            name=name,
            salary_info=salary_info,
            deduction_info=deduction_info
        )
        
        # 添加党员
        try:
            self.member_manager.add_member(member, selected_branch.name)
            print(f"\n党员 '{name}' 已成功添加到支部 '{selected_branch.name}'！")
        except Exception as e:
            print(f"添加失败：{e}")
    
    def remove_member(self):
        """删除党员"""
        print("\n删除党员")
        print("-" * 40)
        
        # 显示支部列表供选择
        branches = self.member_manager.get_all_branches()
        if not branches:
            print("暂无党支部！")
            return
        
        print("\n请选择党支部：")
        for i, branch in enumerate(branches, 1):
            print(f"{i}. {branch.name}")
        
        while True:
            choice = self.input_int("请输入选项编号", 0)
            if 1 <= choice <= len(branches):
                selected_branch = branches[choice - 1]
                break
            else:
                print("无效的选项，请重新输入！")
        
        if not selected_branch.members:
            print(f"支部 '{selected_branch.name}' 暂无党员！")
            return
        
        # 显示该支部的党员列表
        print(f"\n支部 '{selected_branch.name}' 的党员列表：")
        for i, member in enumerate(selected_branch.members, 1):
            print(f"{i}. {member.name}")
        
        while True:
            choice = self.input_int("请输入要删除的党员编号（0取消）", 0)
            if choice == 0:
                return
            if 1 <= choice <= len(selected_branch.members):
                selected_member = selected_branch.members[choice - 1]
                break
            else:
                print("无效的选项，请重新输入！")
        
        # 确认删除
        confirm = input(f"确定要删除党员 '{selected_member.name}' 吗？(y/n): ").strip().lower()
        if confirm != 'y':
            print("已取消删除。")
            return
        
        # 执行删除
        if self.member_manager.remove_member(selected_member.name, selected_branch.name):
            print(f"党员 '{selected_member.name}' 已成功删除！")
        else:
            print("删除失败！")
    
    def update_member(self):
        """修改党员信息"""
        print("\n修改党员信息")
        print("-" * 40)
        
        # 显示支部列表供选择
        branches = self.member_manager.get_all_branches()
        if not branches:
            print("暂无党支部！")
            return
        
        print("\n请选择党支部：")
        for i, branch in enumerate(branches, 1):
            print(f"{i}. {branch.name}")
        
        while True:
            choice = self.input_int("请输入选项编号", 0)
            if 1 <= choice <= len(branches):
                selected_branch = branches[choice - 1]
                break
            else:
                print("无效的选项，请重新输入！")
        
        if not selected_branch.members:
            print(f"支部 '{selected_branch.name}' 暂无党员！")
            return
        
        # 显示该支部的党员列表
        print(f"\n支部 '{selected_branch.name}' 的党员列表：")
        for i, member in enumerate(selected_branch.members, 1):
            print(f"{i}. {member.name}")
        
        while True:
            choice = self.input_int("请输入要修改的党员编号（0取消）", 0)
            if choice == 0:
                return
            if 1 <= choice <= len(selected_branch.members):
                selected_member = selected_branch.members[choice - 1]
                break
            else:
                print("无效的选项，请重新输入！")
        
        # 显示当前信息
        print(f"\n党员 '{selected_member.name}' 当前信息：")
        print(f"  岗位工资: {selected_member.salary_info.position_salary}")
        print(f"  薪级工资: {selected_member.salary_info.rank_salary}")
        print(f"  高定工资: {selected_member.salary_info.fixed_salary}")
        print(f"  基础性绩效: {selected_member.salary_info.basic_performance}")
        print(f"  住房公积金: {selected_member.deduction_info.housing_fund}")
        print(f"  医疗保险: {selected_member.deduction_info.medical_insurance}")
        print(f"  养老保险: {selected_member.deduction_info.pension_insurance}")
        print(f"  职业年金: {selected_member.deduction_info.occupational_annuity}")
        print(f"  大额医疗: {selected_member.deduction_info.large_medical}")
        print(f"  失业保险: {selected_member.deduction_info.unemployment_insurance}")
        print(f"  个人所得税: {selected_member.deduction_info.personal_income_tax}")
        
        # 输入新信息
        print("\n请输入新的信息（直接回车保持原值）：")
        position_salary = self.input_float("岗位工资", selected_member.salary_info.position_salary)
        rank_salary = self.input_float("薪级工资", selected_member.salary_info.rank_salary)
        fixed_salary = self.input_float("高定工资", selected_member.salary_info.fixed_salary)
        basic_performance = self.input_float("基础性绩效", selected_member.salary_info.basic_performance)
        
        housing_fund = self.input_float("住房公积金", selected_member.deduction_info.housing_fund)
        medical_insurance = self.input_float("医疗保险", selected_member.deduction_info.medical_insurance)
        pension_insurance = self.input_float("养老保险", selected_member.deduction_info.pension_insurance)
        occupational_annuity = self.input_float("职业年金", selected_member.deduction_info.occupational_annuity)
        large_medical = self.input_float("大额医疗", selected_member.deduction_info.large_medical)
        unemployment_insurance = self.input_float("失业保险", selected_member.deduction_info.unemployment_insurance)
        personal_income_tax = self.input_float("个人所得税", selected_member.deduction_info.personal_income_tax)
        
        # 创建新的工资和扣款信息对象
        salary_info = SalaryInfo(
            position_salary=position_salary,
            rank_salary=rank_salary,
            fixed_salary=fixed_salary,
            basic_performance=basic_performance
        )
        
        deduction_info = DeductionInfo(
            housing_fund=housing_fund,
            medical_insurance=medical_insurance,
            pension_insurance=pension_insurance,
            occupational_annuity=occupational_annuity,
            large_medical=large_medical,
            unemployment_insurance=unemployment_insurance,
            personal_income_tax=personal_income_tax
        )
        
        # 更新党员信息
        if self.member_manager.update_member(
            selected_member.name, 
            selected_branch.name, 
            salary_info, 
            deduction_info
        ):
            print(f"党员 '{selected_member.name}' 信息已成功更新！")
        else:
            print("更新失败！")
    
    def view_member_list(self):
        """查看党员列表"""
        print("\n党员列表")
        print("-" * 60)
        
        all_members = self.member_manager.get_all_members()
        if not all_members:
            print("暂无党员数据！")
            return
        
        # 按支部分组显示
        branches = self.member_manager.get_all_branches()
        for branch in branches:
            if branch.members:
                print(f"\n【{branch.name}】（共{len(branch.members)}人）")
                print(f"{'序号':<8}{'姓名':<15}{'缴费基数':<15}{'月党费(元)':<10}")
                print("-" * 50)
                for member in branch.members:
                    print(f"{member.sequence:<8}{member.name:<15}{member.payment_base:<15.2f}{member.monthly_fee:<10.2f}")
        
        print(f"\n总计：{len(all_members)}名党员")
    
    def view_member_detail(self):
        """查看党员详情"""
        print("\n查看党员详情")
        print("-" * 40)
        
        name = input("请输入党员姓名: ").strip()
        if not name:
            print("姓名不能为空！")
            return
        
        member = self.member_manager.get_member(name)
        if not member:
            print(f"未找到党员 '{name}'！")
            return
        
        print("\n" + "=" * 50)
        print(f"党员 '{member.name}' 详细信息")
        print("=" * 50)
        print(f"所属党支部: {member.branch_name}")
        print(f"支部序号: {member.branch_sequence}")
        print(f"党员序号: {member.sequence}")
        print("\n【工资信息】")
        print(f"  岗位工资: {member.salary_info.position_salary:.2f} 元")
        print(f"  薪级工资: {member.salary_info.rank_salary:.2f} 元")
        print(f"  高定工资: {member.salary_info.fixed_salary:.2f} 元")
        print(f"  基础性绩效: {member.salary_info.basic_performance:.2f} 元")
        print("\n【扣款信息】")
        print(f"  住房公积金: {member.deduction_info.housing_fund:.2f} 元")
        print(f"  医疗保险: {member.deduction_info.medical_insurance:.2f} 元")
        print(f"  养老保险: {member.deduction_info.pension_insurance:.2f} 元")
        print(f"  职业年金: {member.deduction_info.occupational_annuity:.2f} 元")
        print(f"  大额医疗: {member.deduction_info.large_medical:.2f} 元")
        print(f"  失业保险: {member.deduction_info.unemployment_insurance:.2f} 元")
        print(f"  个人所得税: {member.deduction_info.personal_income_tax:.2f} 元")
        print("\n【党费信息】")
        print(f"  缴费基数: {member.payment_base:.2f} 元")
        print(f"  每月应缴党费: {member.monthly_fee:.2f} 元")
        print("=" * 50)
    
    def add_branch(self):
        """添加党支部"""
        print("\n添加党支部")
        print("-" * 40)
        
        name = input("请输入党支部名称: ").strip()
        if not name:
            print("党支部名称不能为空！")
            return
        
        # 检查是否已存在
        existing_branch = self.member_manager.get_branch(name)
        if existing_branch:
            print(f"党支部 '{name}' 已存在！")
            return
        
        # 输入支部序号
        while True:
            sequence = self.input_int("请输入支部序号", 0)
            if sequence <= 0:
                print("支部序号必须大于0！")
                continue
            
            # 检查序号是否已存在
            branches = self.member_manager.get_all_branches()
            sequence_exists = any(b.sequence == sequence for b in branches)
            if sequence_exists:
                print(f"支部序号 '{sequence}' 已存在，请使用其他序号！")
                continue
            
            break
        
        # 创建党支部对象
        branch = PartyBranch(
            name=name,
            sequence=sequence
        )
        
        # 添加党支部
        try:
            self.member_manager.add_branch(branch)
            print(f"党支部 '{name}' 已成功添加！")
        except Exception as e:
            print(f"添加失败：{e}")
    
    def remove_branch(self):
        """删除党支部"""
        print("\n删除党支部")
        print("-" * 40)
        
        branches = self.member_manager.get_all_branches()
        if not branches:
            print("暂无党支部！")
            return
        
        print("\n党支部列表：")
        for i, branch in enumerate(branches, 1):
            print(f"{i}. {branch.name}（{len(branch.members)}名党员）")
        
        while True:
            choice = self.input_int("请输入要删除的支部编号（0取消）", 0)
            if choice == 0:
                return
            if 1 <= choice <= len(branches):
                selected_branch = branches[choice - 1]
                break
            else:
                print("无效的选项，请重新输入！")
        
        # 检查支部是否有党员
        if selected_branch.members:
            print(f"警告：支部 '{selected_branch.name}' 有 {len(selected_branch.members)} 名党员！")
            confirm = input("确定要删除该支部及其所有党员吗？(y/n): ").strip().lower()
            if confirm != 'y':
                print("已取消删除。")
                return
        else:
            confirm = input(f"确定要删除支部 '{selected_branch.name}' 吗？(y/n): ").strip().lower()
            if confirm != 'y':
                print("已取消删除。")
                return
        
        # 执行删除
        if self.member_manager.remove_branch(selected_branch.name):
            print(f"党支部 '{selected_branch.name}' 已成功删除！")
        else:
            print("删除失败！")
    
    def view_branch_list(self):
        """查看党支部列表"""
        print("\n党支部列表")
        print("-" * 80)
        
        branches = self.member_manager.get_all_branches()
        if not branches:
            print("暂无党支部数据！")
            return
        
        print(f"{'序号':<8}{'支部序号':<10}{'支部名称':<40}{'党员人数':<10}{'总党费(元)':<12}")
        print("-" * 80)
        
        for i, branch in enumerate(branches, 1):
            total_fee = sum(m.monthly_fee for m in branch.members)
            print(f"{i:<8}{branch.sequence:<10}{branch.name:<40}{len(branch.members):<10}{total_fee:<12.2f}")
        
        print("-" * 80)
        total_members = sum(len(b.members) for b in branches)
        total_fee = sum(sum(m.monthly_fee for m in b.members) for b in branches)
        print(f"总计：{len(branches)}个党支部，{total_members}名党员，总党费：{total_fee:.2f}元")
    
    def view_branch_detail(self):
        """查看党支部详情"""
        print("\n查看党支部详情")
        print("-" * 40)
        
        branches = self.member_manager.get_all_branches()
        if not branches:
            print("暂无党支部数据！")
            return
        
        print("\n请选择党支部：")
        for i, branch in enumerate(branches, 1):
            print(f"{i}. {branch.name}")
        
        while True:
            choice = self.input_int("请输入选项编号", 0)
            if 1 <= choice <= len(branches):
                selected_branch = branches[choice - 1]
                break
            else:
                print("无效的选项，请重新输入！")
        
        # 计算支部统计信息
        total_fee = sum(m.monthly_fee for m in selected_branch.members)
        avg_fee = total_fee / len(selected_branch.members) if selected_branch.members else 0
        
        print("\n" + "=" * 60)
        print(f"党支部 '{selected_branch.name}' 详细信息")
        print("=" * 60)
        print(f"支部序号: {selected_branch.sequence}")
        print(f"党员人数: {len(selected_branch.members)}")
        print(f"总党费: {total_fee:.2f} 元")
        print(f"人均党费: {avg_fee:.2f} 元")
        
        if selected_branch.members:
            print("\n【党员列表】")
            print(f"{'序号':<8}{'姓名':<15}{'缴费基数':<15}{'月党费(元)':<10}")
            print("-" * 50)
            for member in selected_branch.members:
                print(f"{member.sequence:<8}{member.name:<15}{member.payment_base:<15.2f}{member.monthly_fee:<10.2f}")
        
        print("=" * 60)
    
    def calculate_fees(self):
        """计算党费"""
        print("\n计算党费")
        print("-" * 40)
        
        branches = self.member_manager.get_all_branches()
        if not branches:
            print("暂无党支部数据！")
            return
        
        if not any(b.members for b in branches):
            print("暂无党员数据！")
            return
        
        # 计算所有党员的党费
        FeeCalculator.calculate_all_fees(branches)
        
        # 显示计算结果
        total_members = sum(len(b.members) for b in branches)
        total_fee = sum(sum(m.monthly_fee for m in b.members) for b in branches)
        
        print(f"\n党费计算完成！")
        print(f"计算党员人数: {total_members}")
        print(f"总党费金额: {total_fee:.2f} 元")
        print(f"\n各支部党费情况：")
        for branch in branches:
            if branch.members:
                branch_total = sum(m.monthly_fee for m in branch.members)
                print(f"  {branch.name}: {branch_total:.2f} 元（{len(branch.members)}人）")
    
    def generate_fee_detail_sheet(self):
        """生成党费明细表"""
        print("\n生成党费明细表")
        print("-" * 40)
        
        branches = self.member_manager.get_all_branches()
        if not branches:
            print("暂无党支部数据！")
            return
        
        if not any(b.members for b in branches):
            print("暂无党员数据！")
            return
        
        try:
            filepath = self.excel_generator.generate_fee_detail_sheet(
                branches, self.current_year, self.current_month
            )
            print(f"\n党费明细表已成功生成！")
            print(f"文件路径: {filepath}")
        except Exception as e:
            print(f"生成失败：{e}")
    
    def generate_branch_fee_sheets(self):
        """生成各支部党费收缴子表"""
        print("\n生成各支部党费收缴子表")
        print("-" * 40)
        
        branches = self.member_manager.get_all_branches()
        if not branches:
            print("暂无党支部数据！")
            return
        
        if not any(b.members for b in branches):
            print("暂无党员数据！")
            return
        
        try:
            filepaths = self.excel_generator.generate_all_branch_sheets(
                branches, self.current_year, self.current_month
            )
            print(f"\n成功生成 {len(filepaths)} 个支部党费收缴子表！")
            print("生成的文件：")
            for filepath in filepaths:
                print(f"  - {filepath}")
        except Exception as e:
            print(f"生成失败：{e}")
    
    def generate_summary_sheet(self):
        """生成党费汇总表"""
        print("\n生成党费汇总表")
        print("-" * 40)
        
        branches = self.member_manager.get_all_branches()
        if not branches:
            print("暂无党支部数据！")
            return
        
        try:
            filepath = self.excel_generator.generate_summary_sheet(
                branches, self.current_year, self.current_month
            )
            print(f"\n党费汇总表已成功生成！")
            print(f"文件路径: {filepath}")
        except Exception as e:
            print(f"生成失败：{e}")
    
    def generate_all_reports(self):
        """生成所有报表"""
        print("\n生成所有报表")
        print("-" * 40)
        
        branches = self.member_manager.get_all_branches()
        if not branches:
            print("暂无党支部数据！")
            return
        
        # 先计算党费
        print("\n1. 计算党费...")
        FeeCalculator.calculate_all_fees(branches)
        
        # 生成党费明细表
        print("2. 生成党费明细表...")
        try:
            filepath1 = self.excel_generator.generate_fee_detail_sheet(
                branches, self.current_year, self.current_month
            )
            print(f"   成功: {filepath1}")
        except Exception as e:
            print(f"   失败: {e}")
        
        # 生成各支部党费收缴子表
        print("3. 生成各支部党费收缴子表...")
        try:
            filepaths = self.excel_generator.generate_all_branch_sheets(
                branches, self.current_year, self.current_month
            )
            print(f"   成功生成 {len(filepaths)} 个文件")
        except Exception as e:
            print(f"   失败: {e}")
        
        # 生成党费汇总表
        print("4. 生成党费汇总表...")
        try:
            filepath3 = self.excel_generator.generate_summary_sheet(
                branches, self.current_year, self.current_month
            )
            print(f"   成功: {filepath3}")
        except Exception as e:
            print(f"   失败: {e}")
        
        print(f"\n所有报表已生成完成！文件保存在: {OUTPUT_DIR}")
    
    def view_statistics(self):
        """查看统计信息"""
        print("\n查看统计信息")
        print("-" * 40)
        
        branches = self.member_manager.get_all_branches()
        if not branches:
            print("暂无党支部数据！")
            return
        
        # 先计算党费
        FeeCalculator.calculate_all_fees(branches)
        
        # 打印统计报告
        FeeStatistics.print_statistics_report(branches, self.current_year, self.current_month)
    
    def save_data(self):
        """保存数据"""
        print("\n保存数据")
        print("-" * 40)
        
        try:
            filepath = self.member_manager.save_to_file(self.current_year, self.current_month)
            print(f"数据已成功保存到: {filepath}")
        except Exception as e:
            print(f"保存失败：{e}")
    
    def load_data(self):
        """加载数据"""
        print("\n加载数据")
        print("-" * 40)
        
        # 询问年份和月份
        year = self.input_int("请输入年份", self.current_year)
        month = self.input_int("请输入月份", self.current_month)
        
        # 更新当前年月
        self.current_year = year
        self.current_month = month
        
        try:
            if self.member_manager.load_from_file(year, month):
                print(f"成功加载 {year}年{month}月 的数据！")
            else:
                print(f"未找到 {year}年{month}月 的数据文件，已创建空数据。")
        except Exception as e:
            print(f"加载失败：{e}")
    
    def set_year_month(self):
        """设置年月"""
        print("\n设置年月")
        print("-" * 40)
        print(f"当前年月: {self.current_year}年{self.current_month}月")
        
        year = self.input_int("请输入年份", self.current_year)
        month = self.input_int("请输入月份", self.current_month)
        
        # 验证月份
        if not (1 <= month <= 12):
            print("月份必须在1-12之间！")
            return
        
        self.current_year = year
        self.current_month = month
        
        print(f"\n已设置为: {year}年{month}月")
    
    def run_person_management(self):
        """运行人员维护管理"""
        while True:
            self.display_person_management_menu()
            choice = self.input_int("请选择操作", 0)
            
            if choice == 0:
                break
            elif choice == 1:
                self.add_member()
            elif choice == 2:
                self.remove_member()
            elif choice == 3:
                self.update_member()
            elif choice == 4:
                self.view_member_list()
            elif choice == 5:
                self.view_member_detail()
            else:
                print("无效的选项，请重新输入！")
    
    def run_branch_management(self):
        """运行支部维护管理"""
        while True:
            self.display_branch_management_menu()
            choice = self.input_int("请选择操作", 0)
            
            if choice == 0:
                break
            elif choice == 1:
                self.add_branch()
            elif choice == 2:
                self.remove_branch()
            elif choice == 3:
                self.view_branch_list()
            elif choice == 4:
                self.view_branch_detail()
            else:
                print("无效的选项，请重新输入！")
    
    def run(self):
        """运行主程序"""
        print("\n欢迎使用党费收取系统！")
        
        while True:
            self.display_menu()
            choice = self.input_int("请选择操作", 0)
            
            if choice == 0:
                print("\n感谢使用党费收取系统，再见！")
                break
            elif choice == 1:
                self.run_person_management()
            elif choice == 2:
                self.run_branch_management()
            elif choice == 3:
                self.calculate_fees()
            elif choice == 4:
                self.generate_fee_detail_sheet()
            elif choice == 5:
                self.generate_branch_fee_sheets()
            elif choice == 6:
                self.generate_summary_sheet()
            elif choice == 7:
                self.generate_all_reports()
            elif choice == 8:
                self.view_statistics()
            elif choice == 9:
                self.save_data()
            elif choice == 10:
                self.load_data()
            elif choice == 11:
                self.set_year_month()
            else:
                print("无效的选项，请重新输入！")


def main():
    """主函数"""
    system = PartyFeeSystem()
    system.run()


if __name__ == "__main__":
    main()
