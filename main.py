# 党费收取系统主程序
import os
import sys
import glob
from typing import List, Optional, Tuple

from config import DEFAULT_YEAR, DEFAULT_MONTH, OUTPUT_DIR, DATA_DIR
from data_models import PartyMember, PartyBranch, SalaryInfo, DeductionInfo
from member_manager import MemberManager
from fee_calculator import FeeCalculator
from excel_generator import ExcelGenerator
from statistics import FeeStatistics


class ImportMode:
    """导入模式"""
    ADD = "add"
    REPLACE = "replace"
    MERGE = "merge"


class PartyFeeSystem:
    """党费收取系统主类"""
    
    def __init__(self):
        """初始化系统"""
        self.member_manager = MemberManager()
        self.excel_generator = ExcelGenerator()
        self.current_year = DEFAULT_YEAR
        self.current_month = DEFAULT_MONTH
        
        # 自动加载已有数据
        self._auto_load_data()
    
    def _auto_load_data(self):
        """自动加载已有数据"""
        print("\n" + "=" * 60)
        print("党费收取系统 v1.0")
        print("=" * 60)
        
        # 检查数据目录
        if not os.path.exists(DATA_DIR):
            print("\n数据目录不存在，这是第一次使用系统。")
            print("请先添加党支部和党员信息。")
            print("=" * 60)
            return
        
        # 查找所有数据文件
        data_files = glob.glob(os.path.join(DATA_DIR, "party_fee_data_*.json"))
        
        if not data_files:
            print("\n未找到历史数据文件，这是第一次使用系统。")
            print("请先添加党支部和党员信息。")
            print("=" * 60)
            return
        
        # 解析数据文件中的年月信息
        available_data = []
        for filepath in data_files:
            filename = os.path.basename(filepath)
            # 从文件名中提取年月，格式为 party_fee_data_YYYY_MM.json
            parts = filename.replace(".json", "").split("_")
            if len(parts) >= 5:
                try:
                    year = int(parts[3])
                    month = int(parts[4])
                    available_data.append((year, month, filepath))
                except (ValueError, IndexError):
                    continue
        
        if not available_data:
            print("\n未找到有效的历史数据文件。")
            print("=" * 60)
            return
        
        # 按年月排序（最新的在前）
        available_data.sort(key=lambda x: (x[0], x[1]), reverse=True)
        
        print(f"\n找到 {len(available_data)} 个历史数据文件：")
        print("-" * 60)
        for i, (year, month, filepath) in enumerate(available_data, 1):
            print(f"{i}. {year}年{month}月")
        
        print("-" * 60)
        print("0. 开始新的月份（不加载历史数据）")
        print("-" * 60)
        
        # 让用户选择
        while True:
            try:
                choice = input(f"\n请选择要加载的数据 [默认: 1 - {available_data[0][0]}年{available_data[0][1]}月]: ").strip()
                
                if choice == "" or choice == "1":
                    # 默认加载最新的
                    selected = available_data[0]
                    break
                elif choice == "0":
                    print("\n已选择开始新的月份。")
                    print("=" * 60)
                    return
                else:
                    choice_num = int(choice)
                    if 1 <= choice_num <= len(available_data):
                        selected = available_data[choice_num - 1]
                        break
                    else:
                        print(f"请输入 0 到 {len(available_data)} 之间的数字！")
            except ValueError:
                print("请输入有效的数字！")
        
        # 加载选中的数据
        year, month, filepath = selected
        self.current_year = year
        self.current_month = month
        
        if self.member_manager.load_from_file(year, month):
            # 计算所有党费（确保数据是最新的）
            branches = self.member_manager.get_all_branches()
            FeeCalculator.calculate_all_fees(branches)
            
            total_branches = len(branches)
            total_members = sum(len(b.members) for b in branches)
            
            print(f"\n成功加载 {year}年{month}月 的数据！")
            print(f"  - 党支部数量: {total_branches}")
            print(f"  - 党员总人数: {total_members}")
            print("=" * 60)
        else:
            print(f"\n加载 {year}年{month}月 的数据失败！")
            print("=" * 60)
    
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
        print("12. 导入外部表格数据")
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
        print("-" * 60)
        
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
        
        # 初始化党员信息收集
        while True:
            # 输入党员基本信息
            print("\n" + "=" * 60)
            print("请输入党员信息（输入 'q' 或 'quit' 取消）")
            print("-" * 60)
            
            # 输入姓名
            while True:
                name = input("\n请输入党员姓名: ").strip()
                if name.lower() in ['q', 'quit']:
                    print("已取消添加党员。")
                    return
                if name:
                    # 检查党员是否已存在
                    existing_member = self.member_manager.get_member(name, selected_branch.name)
                    if existing_member:
                        print(f"党员 '{name}' 已在支部 '{selected_branch.name}' 中！")
                        continue
                    break
                print("姓名不能为空！")
            
            # 输入工资信息
            print("\n请输入工资信息（单位：元，直接回车使用默认值 0）：")
            print("-" * 60)
            
            position_salary = self.input_float("1. 岗位工资", 0.0)
            rank_salary = self.input_float("2. 薪级工资", 0.0)
            fixed_salary = self.input_float("3. 高定工资", 0.0)
            basic_performance = self.input_float("4. 基础性绩效", 0.0)
            
            # 输入扣款信息
            print("\n请输入扣款信息（单位：元，直接回车使用默认值 0）：")
            print("-" * 60)
            
            housing_fund = self.input_float("5. 住房公积金", 0.0)
            medical_insurance = self.input_float("6. 医疗保险", 0.0)
            pension_insurance = self.input_float("7. 养老保险", 0.0)
            occupational_annuity = self.input_float("8. 职业年金", 0.0)
            large_medical = self.input_float("9. 大额医疗", 0.0)
            unemployment_insurance = self.input_float("10. 失业保险", 0.0)
            personal_income_tax = self.input_float("11. 个人所得税", 0.0)
            
            # 显示确认信息
            while True:
                print("\n" + "=" * 60)
                print("请确认以下信息：")
                print("-" * 60)
                print(f"所属党支部: {selected_branch.name}")
                print(f"党员姓名: {name}")
                print("\n【工资信息】")
                print(f"  1. 岗位工资: {position_salary:.2f} 元")
                print(f"  2. 薪级工资: {rank_salary:.2f} 元")
                print(f"  3. 高定工资: {fixed_salary:.2f} 元")
                print(f"  4. 基础性绩效: {basic_performance:.2f} 元")
                print("\n【扣款信息】")
                print(f"  5. 住房公积金: {housing_fund:.2f} 元")
                print(f"  6. 医疗保险: {medical_insurance:.2f} 元")
                print(f"  7. 养老保险: {pension_insurance:.2f} 元")
                print(f"  8. 职业年金: {occupational_annuity:.2f} 元")
                print(f"  9. 大额医疗: {large_medical:.2f} 元")
                print(f"  10. 失业保险: {unemployment_insurance:.2f} 元")
                print(f"  11. 个人所得税: {personal_income_tax:.2f} 元")
                print("-" * 60)
                print("\n请选择操作：")
                print("  1. 确认添加")
                print("  2. 修改信息")
                print("  3. 取消添加")
                print("  或输入字段编号（1-11）修改对应字段")
                print("-" * 60)
                
                confirm_choice = input("\n请选择 [默认: 1): ").strip()
                
                if confirm_choice == "" or confirm_choice == "1":
                    # 确认添加
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
                        
                        # 自动计算党费
                        FeeCalculator.calculate_member_fee(member)
                        
                        print(f"\n[成功] 党员 '{name}' 已添加到支部 '{selected_branch.name}'！")
                        print(f"  缴费基数: {member.payment_base:.2f} 元")
                        print(f"  每月应缴党费: {member.monthly_fee:.2f} 元")
                    except Exception as e:
                        print(f"添加失败：{e}")
                    return
                    
                elif confirm_choice == "2":
                    # 重新输入所有信息
                    print("\n>>> 重新输入所有信息...")
                    break  # 跳出确认循环，重新输入
                    
                elif confirm_choice == "3":
                    print("已取消添加党员。")
                    return
                    
                elif confirm_choice in ["1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11"]:
                    # 修改特定字段
                    field_num = int(confirm_choice)
                    if field_num == 1:
                        position_salary = self.input_float("请输入新的岗位工资", position_salary)
                    elif field_num == 2:
                        rank_salary = self.input_float("请输入新的薪级工资", rank_salary)
                    elif field_num == 3:
                        fixed_salary = self.input_float("请输入新的高定工资", fixed_salary)
                    elif field_num == 4:
                        basic_performance = self.input_float("请输入新的基础性绩效", basic_performance)
                    elif field_num == 5:
                        housing_fund = self.input_float("请输入新的住房公积金", housing_fund)
                    elif field_num == 6:
                        medical_insurance = self.input_float("请输入新的医疗保险", medical_insurance)
                    elif field_num == 7:
                        pension_insurance = self.input_float("请输入新的养老保险", pension_insurance)
                    elif field_num == 8:
                        occupational_annuity = self.input_float("请输入新的职业年金", occupational_annuity)
                    elif field_num == 9:
                        large_medical = self.input_float("请输入新的大额医疗", large_medical)
                    elif field_num == 10:
                        unemployment_insurance = self.input_float("请输入新的失业保险", unemployment_insurance)
                    elif field_num == 11:
                        personal_income_tax = self.input_float("请输入新的个人所得税", personal_income_tax)
                    print(f"\n>>> 已修改字段，返回确认菜单...")
                else:
                    print("无效的选项，请重新输入！")
    
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
            # 重新计算党费
            FeeCalculator.calculate_member_fee(selected_member)
            
            print(f"党员 '{selected_member.name}' 信息已成功更新！")
            print(f"  新的缴费基数: {selected_member.payment_base:.2f} 元")
            print(f"  新的每月应缴党费: {selected_member.monthly_fee:.2f} 元")
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
        
        # 确保所有党员的党费都已计算
        branches = self.member_manager.get_all_branches()
        FeeCalculator.calculate_all_fees(branches)
        
        # 按支部分组显示
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
        
        # 确保党费已计算
        FeeCalculator.calculate_member_fee(member)
        
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
        print("-" * 60)
        
        while True:
            # 输入党支部信息收集循环
            print("\n" + "=" * 60)
            print("请输入党支部信息（输入 'q' 或 'quit' 取消）")
            print("-" * 60)
            
            # 输入支部名称
            while True:
                name = input("\n请输入党支部名称: ").strip()
                if name.lower() in ['q', 'quit']:
                    print("已取消添加党支部。")
                    return
                if name:
                    # 检查是否已存在
                    existing_branch = self.member_manager.get_branch(name)
                    if existing_branch:
                        print(f"党支部 '{name}' 已存在！")
                        continue
                    break
                print("党支部名称不能为空！")
            
            # 输入支部序号
            print("\n请输入支部序号：")
            
            while True:
                sequence = self.input_int("支部序号", 0)
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
            
            # 显示确认信息
            while True:
                print("\n" + "=" * 60)
                print("请确认以下信息：")
                print("-" * 60)
                print(f"1. 党支部名称: {name}")
                print(f"2. 支部序号: {sequence}")
                print("-" * 60)
                print("\n请选择操作：")
                print("  1. 确认添加")
                print("  2. 修改信息")
                print("  3. 取消添加")
                print("  或输入字段编号（1-2）修改对应字段")
                print("-" * 60)
                
                confirm_choice = input("\n请选择 [默认: 1): ").strip()
                
                if confirm_choice == "" or confirm_choice == "1":
                    # 确认添加
                    # 创建党支部对象
                    branch = PartyBranch(
                        name=name,
                        sequence=sequence
                    )
                    
                    # 添加党支部
                    try:
                        self.member_manager.add_branch(branch)
                        print(f"\n[成功] 党支部 '{name}' 已添加！")
                    except Exception as e:
                        print(f"添加失败：{e}")
                    return
                    
                elif confirm_choice == "2":
                    # 重新输入所有信息
                    print("\n>>> 重新输入所有信息...")
                    break  # 跳出确认循环，重新输入
                    
                elif confirm_choice == "3":
                    print("已取消添加党支部。")
                    return
                    
                elif confirm_choice in ["1", "2"]:
                    # 修改特定字段
                    field_num = int(confirm_choice)
                    if field_num == 1:
                        while True:
                            new_name = input("请输入新的党支部名称: ").strip()
                            if new_name:
                                # 检查是否已存在
                                existing_branch = self.member_manager.get_branch(new_name)
                                if existing_branch and new_name != name:
                                    print(f"党支部 '{new_name}' 已存在！")
                                    continue
                                name = new_name
                                break
                            print(f"\n>>> 已修改支部名称，返回确认菜单...")
                            break
                    elif field_num == 2:
                        while True:
                            new_sequence = self.input_int("请输入新的支部序号", sequence)
                            if new_sequence <= 0:
                                print("支部序号必须大于0！")
                                continue
                            
                            # 检查序号是否已存在
                            branches = self.member_manager.get_all_branches()
                            sequence_exists = any(b.sequence == new_sequence for b in branches)
                            if sequence_exists and new_sequence != sequence:
                                print(f"支部序号 '{new_sequence}' 已存在，请使用其他序号！")
                                continue
                            
                            sequence = new_sequence
                            break
                        print(f"\n>>> 已修改支部序号，返回确认菜单...")
                else:
                    print("无效的选项，请重新输入！")
    
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
    
    def display_import_menu(self):
        """显示导入菜单"""
        print("\n" + "-" * 60)
        print("导入外部表格数据")
        print("-" * 60)
        print("1. 导入Excel明细表（总表）")
        print("2. 导入Excel子表（各支部表）")
        print("3. 导入Word汇总表（仅查看统计）")
        print("4. 自动检测并导入")
        print("0. 返回主菜单")
        print("-" * 60)
        print("\n说明：")
        print("  - Excel明细表：包含所有支部所有党员的完整信息")
        print("  - Excel子表：每个工作表对应一个支部的党员信息")
        print("  - Word汇总表：仅包含统计信息，不含详细党员数据")
        print("  - 自动检测：根据文件名和结构自动判断类型")
    
    def import_external_data(self):
        """导入外部数据主方法"""
        from excel_importer import ExcelImporter
        from word_importer import WordImporter, check_word_dependency
        
        importer = ExcelImporter()
        word_importer = WordImporter()
        
        while True:
            self.display_import_menu()
            choice = self.input_int("请选择操作", 0)
            
            if choice == 0:
                break
            
            # 询问文件路径
            print("\n" + "-" * 60)
            print("文件路径输入说明：")
            print("  - 可以输入绝对路径（如：D:\\data\\2026年2月党费收缴明细表.xlsx）")
            print("  - 可以输入相对路径（如：2026年2月党费收缴明细表.xlsx）")
            print("  - 输入 'q' 或 'quit' 取消操作")
            print("-" * 60)
            
            filepath = input("\n请输入文件路径: ").strip()
            
            if filepath.lower() in ['q', 'quit']:
                print("已取消导入")
                continue
            
            # 处理路径中的引号
            filepath = filepath.strip('"').strip("'")
            
            if not os.path.exists(filepath):
                print(f"\n错误：文件不存在 - {filepath}")
                # 尝试在output目录查找
                possible_path = os.path.join(OUTPUT_DIR, filepath)
                if os.path.exists(possible_path):
                    print(f"提示：在 output 目录找到了同名文件")
                    use_it = input("是否使用该文件？(y/n): ").strip().lower()
                    if use_it == 'y':
                        filepath = possible_path
                    else:
                        continue
                else:
                    continue
            
            # 根据选择导入
            if choice == 1:
                # 导入Excel明细表
                result = importer.import_from_file(filepath, sheet_type='detail')
            elif choice == 2:
                # 导入Excel子表
                result = importer.import_from_file(filepath, sheet_type='branch')
            elif choice == 3:
                # 导入Word汇总表
                if not check_word_dependency():
                    print("\n" + "!" * 60)
                    print("警告：未安装 python-docx 库，无法解析Word文件")
                    print("请运行: pip install python-docx")
                    print("!" * 60)
                    continue
                
                word_result = word_importer.import_from_file(filepath)
                
                if word_result.errors:
                    print("\n错误信息：")
                    for error in word_result.errors:
                        print(f"  - {error}")
                else:
                    print("\n" + "=" * 60)
                    print("Word汇总表信息")
                    print("=" * 60)
                    
                    if word_result.summary_data.get('title'):
                        print(f"文档标题: {word_result.summary_data['title']}")
                    
                    branches = word_result.summary_data.get('branches', [])
                    if branches:
                        print(f"\n共找到 {len(branches)} 个支部的汇总信息：")
                        print("-" * 60)
                        print(f"{'序号':<6}{'支部名称':<25}{'党员人数':<10}{'金额':<12}{'备注':<15}")
                        print("-" * 60)
                        
                        total_count = 0
                        total_amount = 0.0
                        
                        for item in branches:
                            seq = item.get('sequence', '-')
                            name = item.get('branch_name', '未知')
                            count = item.get('member_count', 0)
                            amount = item.get('amount', 0.0)
                            remark = item.get('remark', '')
                            
                            try:
                                total_count += int(count)
                            except:
                                pass
                            try:
                                total_amount += float(amount)
                            except:
                                pass
                            
                            print(f"{str(seq):<6}{name:<25}{str(count):<10}{str(amount):<12}{remark:<15}")
                        
                        print("-" * 60)
                        print(f"{'合计':<6}{'':<25}{total_count:<10}{total_amount:<12.2f}")
                        print("=" * 60)
                        print("\n注意：Word汇总表仅包含统计信息，不包含详细党员数据")
                        print("如果需要导入完整的党员信息，请使用Excel格式的明细表或子表")
                    else:
                        print("未找到可解析的汇总数据")
                
                continue
            
            elif choice == 4:
                # 自动检测并导入
                file_type = importer.detect_file_type(filepath)
                print(f"\n检测到文件类型: {file_type}")
                
                if file_type in ['detail', 'branch']:
                    result = importer.import_from_file(filepath, sheet_type=file_type)
                elif file_type == 'summary':
                    print("检测到汇总表，仅显示统计信息")
                    if filepath.endswith(('.xlsx', '.xls')):
                        result = importer.import_from_file(filepath, sheet_type='summary')
                    else:
                        # 可能是Word文件
                        if not check_word_dependency():
                            print("\n警告：未安装 python-docx 库")
                            print("请运行: pip install python-docx")
                            continue
                        word_result = word_importer.import_from_file(filepath)
                        if word_result.summary_data.get('branches'):
                            print(f"\n找到 {len(word_result.summary_data['branches'])} 个支部的汇总信息")
                        continue
                else:
                    # 默认为明细表
                    result = importer.import_from_file(filepath, sheet_type='detail')
            else:
                print("无效的选项，请重新输入！")
                continue
            
            # 处理导入结果
            if result.warnings:
                print("\n提示信息：")
                for warning in result.warnings:
                    print(f"  - {warning}")
            
            if result.errors:
                print("\n错误信息：")
                for error in result.errors:
                    print(f"  - {error}")
            
            if not result.success:
                print("\n导入失败！")
                continue
            
            # 显示导入预览
            print("\n" + "=" * 60)
            print("导入预览")
            print("=" * 60)
            print(f"导入的支部数量: {result.total_branches}")
            print(f"导入的党员数量: {result.total_members}")
            print("\n支部列表：")
            print("-" * 60)
            print(f"{'序号':<6}{'支部名称':<30}{'党员人数':<10}")
            print("-" * 60)
            
            for branch in result.branches:
                print(f"{branch.sequence:<6}{branch.name:<30}{len(branch.members):<10}")
            
            print("=" * 60)
            
            # 询问导入模式
            print("\n请选择导入模式：")
            print("  1. 追加（在现有数据基础上添加）")
            print("  2. 替换（删除现有数据，使用导入的数据）")
            print("  3. 合并（合并同名支部的党员）")
            print("  0. 取消导入")
            
            mode_choice = self.input_int("请选择", 0)
            
            if mode_choice == 0:
                print("已取消导入")
                continue
            
            # 执行导入
            if mode_choice == 1:
                # 追加模式
                self._import_append(result.branches)
                print("\n[OK] 已追加导入数据")
            elif mode_choice == 2:
                # 替换模式
                self._import_replace(result.branches)
                print("\n[OK] 已替换现有数据")
            elif mode_choice == 3:
                # 合并模式
                self._import_merge(result.branches)
                print("\n[OK] 已合并导入数据")
            else:
                print("无效的选项，已取消导入")
                continue
            
            # 提示保存
            print("\n提示：导入的数据还未保存，建议使用菜单选项 '9. 保存数据' 进行保存")
    
    def _import_append(self, new_branches: List[PartyBranch]):
        """
        追加模式导入
        
        将新的支部和党员添加到现有数据中
        """
        # 获取当前最大的支部序号和党员序号
        existing_branches = self.member_manager.get_all_branches()
        max_branch_seq = max((b.sequence for b in existing_branches), default=0)
        
        # 为新的支部分配序号
        for branch in new_branches:
            max_branch_seq += 1
            branch.sequence = max_branch_seq
            # 更新党员的支部序号
            for member in branch.members:
                member.branch_sequence = max_branch_seq
            # 添加到管理器
            self.member_manager.add_branch(branch)
    
    def _import_replace(self, new_branches: List[PartyBranch]):
        """
        替换模式导入
        
        删除现有数据，使用导入的数据
        """
        # 清除现有数据
        self.member_manager = MemberManager()
        
        # 添加导入的数据
        for branch in new_branches:
            self.member_manager.add_branch(branch)
    
    def _import_merge(self, new_branches: List[PartyBranch]):
        """
        合并模式导入
        
        合并同名支部的党员
        """
        existing_branches = self.member_manager.get_all_branches()
        existing_branch_dict = {b.name: b for b in existing_branches}
        
        for new_branch in new_branches:
            if new_branch.name in existing_branch_dict:
                # 同名支部存在，合并党员
                existing_branch = existing_branch_dict[new_branch.name]
                
                # 为新党员分配序号
                start_seq = len(existing_branch.members) + 1
                for i, member in enumerate(new_branch.members):
                    member.sequence = start_seq + i
                    member.branch_sequence = existing_branch.sequence
                    existing_branch.members.append(member)
            else:
                # 新支部，直接添加
                max_branch_seq = max((b.sequence for b in existing_branches), default=0)
                new_branch.sequence = max_branch_seq + 1
                for member in new_branch.members:
                    member.branch_sequence = new_branch.sequence
                self.member_manager.add_branch(new_branch)
    
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
            elif choice == 12:
                self.import_external_data()
            else:
                print("无效的选项，请重新输入！")


def main():
    """主函数"""
    system = PartyFeeSystem()
    system.run()


if __name__ == "__main__":
    main()
