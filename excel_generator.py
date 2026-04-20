# Excel文件生成模块
import os
import shutil
from typing import List, Optional
from datetime import datetime

from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter

from config import OUTPUT_DIR, COLLEGE_NAME
from data_models import PartyBranch, PartyMember


# 中文月份数字映射
CHINESE_MONTHS = {
    1: "一", 2: "二", 3: "三", 4: "四", 5: "五", 6: "六",
    7: "七", 8: "八", 9: "九", 10: "十", 11: "十一", 12: "十二"
}

# 需要移除的支部名称前缀和后缀
BRANCH_NAME_PREFIXES = ["中共", "中国共产党", "中共和"]
BRANCH_NAME_SUFFIXES = ["委员会", "党委会", "党总支", "分党委", "党委"]


def get_chinese_month(month: int) -> str:
    """
    将数字月份转换为中文月份
    
    Args:
        month: 数字月份 (1-12)
        
    Returns:
        中文月份字符串，如"一"、"十二"
    """
    return CHINESE_MONTHS.get(month, str(month))


def simplify_branch_name(name: str) -> str:
    """
    简化支部名称，去掉前缀和后缀
    
    Args:
        name: 原始支部名称，如"中共xxxx计算机科学与技术学院研究生第九支部委员会"
        
    Returns:
        简化后的名称，如"计算机科学与技术学院研究生第九支部"
    """
    simplified = name
    
    # 移除前缀
    for prefix in BRANCH_NAME_PREFIXES:
        if simplified.startswith(prefix):
            simplified = simplified[len(prefix):]
            break
    
    # 移除后缀（从后往前找最长的匹配）
    for suffix in sorted(BRANCH_NAME_SUFFIXES, key=len, reverse=True):
        if simplified.endswith(suffix):
            simplified = simplified[:-len(suffix)]
            break
    
    # 确保结果不是空字符串
    if not simplified.strip():
        simplified = name
    
    return simplified.strip()


def get_display_branch_name(name: str, custom_college_name: Optional[str] = None) -> str:
    """
    获取用于表头显示的支部名称
    
    Args:
        name: 原始支部名称
        custom_college_name: 自定义学院名称
        
    Returns:
        用于表头显示的支部名称
    """
    simplified = simplify_branch_name(name)
    
    # 如果有自定义学院名称，检查是否需要替换
    if custom_college_name and custom_college_name in simplified:
        # 已经包含学院名称，直接返回简化后的名称
        return simplified
    
    return simplified


class ExcelGenerator:
    """Excel文件生成器"""
    
    def __init__(self):
        """初始化Excel生成器"""
        self._ensure_output_dir()
        # 默认学院显示名称（可自定义）
        self.college_display_name = COLLEGE_NAME
        
        # 定义样式
        self.title_font = Font(bold=True, size=16)
        self.header_font = Font(bold=True, size=12)
        self.normal_font = Font(size=11)
        
        self.center_alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        self.left_alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
        
        self.thin_border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
        
        self.header_fill = PatternFill(start_color='E8E8E8', end_color='E8E8E8', fill_type='solid')
        self.light_fill = PatternFill(start_color='F5F5F5', end_color='F5F5F5', fill_type='solid')
    
    def _ensure_output_dir(self):
        """确保输出目录存在"""
        if not os.path.exists(OUTPUT_DIR):
            os.makedirs(OUTPUT_DIR)
    
    def _get_month_dir(self, year: int, month: int, create: bool = True) -> str:
        """
        获取月份目录路径
        
        Args:
            year: 年份
            month: 月份
            create: 是否创建目录
            
        Returns:
            月份目录的完整路径
        """
        month_dir = os.path.join(OUTPUT_DIR, f"{year}年{month}月")
        
        if create and not os.path.exists(month_dir):
            os.makedirs(month_dir)
        
        return month_dir
    
    def _get_sub_sheet_dir(self, year: int, month: int, create: bool = True) -> str:
        """
        获取子表目录路径（党费收缴子表）
        
        Args:
            year: 年份
            month: 月份
            create: 是否创建目录
            
        Returns:
            子表目录的完整路径
        """
        month_dir = self._get_month_dir(year, month, create)
        sub_sheet_dir = os.path.join(month_dir, "党费收缴子表")
        
        if create and not os.path.exists(sub_sheet_dir):
            os.makedirs(sub_sheet_dir)
        
        return sub_sheet_dir
    
    def _cleanup_old_files(self, year: int, month: int):
        """
        清理旧文件（在生成新报表前调用）
        
        Args:
            year: 年份
            month: 月份
        """
        month_dir = self._get_month_dir(year, month, create=False)
        
        if os.path.exists(month_dir):
            # 删除整个月份目录，然后重新创建
            shutil.rmtree(month_dir)
            # 重新创建目录结构
            self._get_month_dir(year, month, create=True)
            self._get_sub_sheet_dir(year, month, create=True)
    
    def generate_fee_detail_sheet(self, branches: List[PartyBranch], 
                                    year: int, month: int,
                                    custom_title: Optional[str] = None,
                                    custom_college_name: Optional[str] = None,
                                    use_chinese_month: bool = False) -> str:
        """
        生成党费收缴明细表
        
        包含所有支部的所有党员详细信息
        
        Args:
            branches: 党支部列表
            year: 年份
            month: 月份
            custom_title: 自定义完整标题（如果提供，将覆盖其他设置）
            custom_college_name: 自定义学院名称（如果不提供，使用默认的college_display_name）
            use_chinese_month: 是否使用中文月份（如"二月"而不是"2月"）
        """
        wb = Workbook()
        ws = wb.active
        
        # 确定学院显示名称
        college_name = custom_college_name if custom_college_name else self.college_display_name
        
        # 确定月份显示格式
        if use_chinese_month:
            month_str = f"{get_chinese_month(month)}月"
        else:
            month_str = f"{month}月"
        
        ws.title = f"{year}年{month_str}党费收缴明细表"
        
        # 生成表格标题
        if custom_title:
            title = custom_title
        else:
            title = f"{college_name}{year}年{month_str}党费收缴明细表"
        
        # 定义列宽
        column_widths = [8, 25, 8, 15, 12, 12, 12, 14, 12, 12, 12, 12, 12, 12, 12, 12, 15, 12]
        for i, width in enumerate(column_widths, 1):
            ws.column_dimensions[get_column_letter(i)].width = width
        
        # 第一行：标题
        ws.merge_cells('A1:R1')
        ws['A1'] = title
        ws['A1'].font = self.title_font
        ws['A1'].alignment = self.center_alignment
        
        # 第二行：二级标题（工资项目和扣款项目）
        ws.merge_cells('F2:I2')  # 工资项目
        ws['F2'] = "工资项目（元）"
        ws['F2'].font = self.header_font
        ws['F2'].alignment = self.center_alignment
        
        ws.merge_cells('J2:Q2')  # 扣款项目
        ws['J2'] = "扣款项目（元）"
        ws['J2'].font = self.header_font
        ws['J2'].alignment = self.center_alignment
        
        # 第三行：列标题
        headers = [
            "支部序号", "所属党支部", "序号", "姓名",
            "岗位工资", "薪级工资", "高定工资", "基础性绩效",  # 工资项目
            "住房公积金", "医疗保险", "养老保险", "职业年金", 
            "大额医疗", "失业保险", "个人所得税",  # 扣款项目
            "缴费基数", "每月应缴党费（元）", "支部每月应缴党费（元）"
        ]
        
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=3, column=col, value=header)
            cell.font = self.header_font
            cell.alignment = self.center_alignment
            cell.border = self.thin_border
            cell.fill = self.header_fill
        
        # 填充数据
        row = 4
        current_branch_sequence = None
        current_branch_name = None
        branch_start_row = None
        
        for branch in sorted(branches, key=lambda b: b.sequence):
            # 记录支部信息
            branch_sequence = branch.sequence
            branch_name = branch.name
            branch_total_fee = sum(m.monthly_fee for m in branch.members)
            
            for member in branch.members:
                # 支部序号（合并单元格）
                if branch_sequence != current_branch_sequence:
                    if current_branch_sequence is not None and branch_start_row is not None:
                        # 合并上一个支部的序号单元格
                        ws.merge_cells(f'A{branch_start_row}:A{row-1}')
                        # 合并上一个支部的名称单元格
                        ws.merge_cells(f'B{branch_start_row}:B{row-1}')
                        # 合并上一个支部的总党费单元格
                        ws.merge_cells(f'R{branch_start_row}:R{row-1}')
                    
                    current_branch_sequence = branch_sequence
                    current_branch_name = branch_name
                    branch_start_row = row
                
                # 填充数据
                ws.cell(row=row, column=1, value=branch_sequence).alignment = self.center_alignment
                ws.cell(row=row, column=2, value=branch_name).alignment = self.center_alignment
                ws.cell(row=row, column=3, value=member.sequence).alignment = self.center_alignment
                ws.cell(row=row, column=4, value=member.name).alignment = self.left_alignment
                
                # 工资项目
                ws.cell(row=row, column=5, value=member.salary_info.position_salary).alignment = self.center_alignment
                ws.cell(row=row, column=6, value=member.salary_info.rank_salary).alignment = self.center_alignment
                ws.cell(row=row, column=7, value=member.salary_info.fixed_salary).alignment = self.center_alignment
                ws.cell(row=row, column=8, value=member.salary_info.basic_performance).alignment = self.center_alignment
                
                # 扣款项目
                ws.cell(row=row, column=9, value=member.deduction_info.housing_fund).alignment = self.center_alignment
                ws.cell(row=row, column=10, value=member.deduction_info.medical_insurance).alignment = self.center_alignment
                ws.cell(row=row, column=11, value=member.deduction_info.pension_insurance).alignment = self.center_alignment
                ws.cell(row=row, column=12, value=member.deduction_info.occupational_annuity).alignment = self.center_alignment
                ws.cell(row=row, column=13, value=member.deduction_info.large_medical).alignment = self.center_alignment
                ws.cell(row=row, column=14, value=member.deduction_info.unemployment_insurance).alignment = self.center_alignment
                ws.cell(row=row, column=15, value=member.deduction_info.personal_income_tax).alignment = self.center_alignment
                
                # 缴费基数和党费
                ws.cell(row=row, column=16, value=member.payment_base).alignment = self.center_alignment
                ws.cell(row=row, column=17, value=member.monthly_fee).alignment = self.center_alignment
                ws.cell(row=row, column=18, value=branch_total_fee).alignment = self.center_alignment
                
                # 应用边框
                for col in range(1, 19):
                    ws.cell(row=row, column=col).border = self.thin_border
                    # 交替行背景色
                    if row % 2 == 0:
                        ws.cell(row=row, column=col).fill = self.light_fill
                
                row += 1
            
            # 处理最后一个支部的合并单元格
            if branch_start_row is not None and branch_start_row < row:
                ws.merge_cells(f'A{branch_start_row}:A{row-1}')
                ws.merge_cells(f'B{branch_start_row}:B{row-1}')
                ws.merge_cells(f'R{branch_start_row}:R{row-1}')
        
        # 设置行高
        ws.row_dimensions[1].height = 30
        ws.row_dimensions[2].height = 25
        ws.row_dimensions[3].height = 40
        
        # 保存文件到月份目录
        month_dir = self._get_month_dir(year, month)
        filename = f"{year}年{month}月党费收缴明细表.xlsx"
        filepath = os.path.join(month_dir, filename)
        wb.save(filepath)
        
        return filepath
    
    def generate_branch_fee_sheet(self, branch: PartyBranch, 
                                    year: int, month: int,
                                    custom_title: Optional[str] = None,
                                    custom_college_name: Optional[str] = None,
                                    use_chinese_month: bool = True) -> str:
        """
        生成单个支部的党费收缴明细
        
        每个支部一个单独的表格
        
        Args:
            branch: 党支部对象
            year: 年份
            month: 月份
            custom_title: 自定义完整标题（如果提供，将覆盖其他设置）
            custom_college_name: 自定义学院名称（如果不提供，使用默认的college_display_name）
            use_chinese_month: 是否使用中文月份（默认True，如"二月"而不是"2月"）
        """
        wb = Workbook()
        
        # 简化支部名称（去掉"中共"、"委员会"等前缀后缀）
        simplified_branch_name = simplify_branch_name(branch.name)
        
        ws = wb.active
        ws.title = f"{simplified_branch_name}党费收缴明细"
        
        # 确定学院显示名称
        college_name = custom_college_name if custom_college_name else self.college_display_name
        
        # 确定月份显示格式（子表默认使用中文月份）
        if use_chinese_month:
            month_str = f"{get_chinese_month(month)}月"
        else:
            month_str = f"{month}月"
        
        # 生成表格标题
        # 子表格式："2026年计算机科学与技术学院本科生第一党支部二月党费收缴明细"
        # 使用简化的支部名称
        if custom_title:
            title = custom_title
        else:
            title = f"{year}年{college_name}{simplified_branch_name}{month_str}党费收缴明细"
        
        # 定义列宽
        column_widths = [8, 15, 12, 12, 12, 14, 12, 12, 12, 12, 12, 12, 12, 15, 12]
        for i, width in enumerate(column_widths, 1):
            ws.column_dimensions[get_column_letter(i)].width = width
        
        # 第一行：标题
        ws.merge_cells('A1:O1')
        ws['A1'] = title
        ws['A1'].font = self.title_font
        ws['A1'].alignment = self.center_alignment
        
        # 第二行：二级标题（工资项目和扣款项目）
        ws.merge_cells('C2:F2')  # 工资项目
        ws['C2'] = "工资项目（元）"
        ws['C2'].font = self.header_font
        ws['C2'].alignment = self.center_alignment
        
        ws.merge_cells('G2:N2')  # 扣款项目
        ws['G2'] = "扣款项目（元）"
        ws['G2'].font = self.header_font
        ws['G2'].alignment = self.center_alignment
        
        # 第三行：列标题
        headers = [
            "序号", "姓名",
            "岗位工资", "薪级工资", "高定工资", "基础性绩效",  # 工资项目
            "住房公积金", "医疗保险", "养老保险", "职业年金", 
            "大额医疗", "失业保险", "个人所得税",  # 扣款项目
            "缴费基数", "每月应缴党费（元）"
        ]
        
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=3, column=col, value=header)
            cell.font = self.header_font
            cell.alignment = self.center_alignment
            cell.border = self.thin_border
            cell.fill = self.header_fill
        
        # 填充数据
        row = 4
        for member in branch.members:
            # 填充数据
            ws.cell(row=row, column=1, value=member.sequence).alignment = self.center_alignment
            ws.cell(row=row, column=2, value=member.name).alignment = self.left_alignment
            
            # 工资项目
            ws.cell(row=row, column=3, value=member.salary_info.position_salary).alignment = self.center_alignment
            ws.cell(row=row, column=4, value=member.salary_info.rank_salary).alignment = self.center_alignment
            ws.cell(row=row, column=5, value=member.salary_info.fixed_salary).alignment = self.center_alignment
            ws.cell(row=row, column=6, value=member.salary_info.basic_performance).alignment = self.center_alignment
            
            # 扣款项目
            ws.cell(row=row, column=7, value=member.deduction_info.housing_fund).alignment = self.center_alignment
            ws.cell(row=row, column=8, value=member.deduction_info.medical_insurance).alignment = self.center_alignment
            ws.cell(row=row, column=9, value=member.deduction_info.pension_insurance).alignment = self.center_alignment
            ws.cell(row=row, column=10, value=member.deduction_info.occupational_annuity).alignment = self.center_alignment
            ws.cell(row=row, column=11, value=member.deduction_info.large_medical).alignment = self.center_alignment
            ws.cell(row=row, column=12, value=member.deduction_info.unemployment_insurance).alignment = self.center_alignment
            ws.cell(row=row, column=13, value=member.deduction_info.personal_income_tax).alignment = self.center_alignment
            
            # 缴费基数和党费
            ws.cell(row=row, column=14, value=member.payment_base).alignment = self.center_alignment
            ws.cell(row=row, column=15, value=member.monthly_fee).alignment = self.center_alignment
            
            # 应用边框
            for col in range(1, 16):
                ws.cell(row=row, column=col).border = self.thin_border
                # 交替行背景色
                if row % 2 == 0:
                    ws.cell(row=row, column=col).fill = self.light_fill
            
            row += 1
        
        # 添加合计行
        total_row = row
        ws.merge_cells(f'A{total_row}:N{total_row}')
        ws.cell(row=total_row, column=1, value="合计").font = self.header_font
        ws.cell(row=total_row, column=1).alignment = self.center_alignment
        ws.cell(row=total_row, column=1).fill = self.header_fill
        
        # 计算总党费
        total_fee = sum(member.monthly_fee for member in branch.members)
        ws.cell(row=total_row, column=15, value=total_fee).font = self.header_font
        ws.cell(row=total_row, column=15).alignment = self.center_alignment
        ws.cell(row=total_row, column=15).fill = self.header_fill
        
        # 应用边框到合计行
        for col in range(1, 16):
            ws.cell(row=total_row, column=col).border = self.thin_border
        
        # 设置行高
        ws.row_dimensions[1].height = 30
        ws.row_dimensions[2].height = 25
        ws.row_dimensions[3].height = 40
        
        # 保存文件到子表目录（党费收缴子表）
        sub_sheet_dir = self._get_sub_sheet_dir(year, month)
        # 改进命名规则，与总表有明显区别
        # 总表命名：2026年2月党费收缴明细表.xlsx
        # 子表命名：2026年2月-计算机科学与技术学院研究生第八党支部-党费收缴子表.xlsx
        # 使用简化的支部名称，去掉"中共"、"委员会"等前缀后缀
        filename = f"{year}年{month}月-{simplified_branch_name}-党费收缴子表.xlsx"
        filepath = os.path.join(sub_sheet_dir, filename)
        wb.save(filepath)
        
        return filepath
    
    def generate_all_branch_sheets(self, branches: List[PartyBranch], 
                                     year: int, month: int,
                                     custom_college_name: Optional[str] = None,
                                     use_chinese_month: bool = True) -> List[str]:
        """
        生成所有支部的党费收缴明细
        
        每个支部一个单独的Excel文件
        
        Args:
            branches: 党支部列表
            year: 年份
            month: 月份
            custom_college_name: 自定义学院名称
            use_chinese_month: 是否使用中文月份
        """
        generated_files = []
        
        for branch in branches:
            filepath = self.generate_branch_fee_sheet(
                branch, year, month,
                custom_college_name=custom_college_name,
                use_chinese_month=use_chinese_month
            )
            generated_files.append(filepath)
        
        return generated_files
    
    def generate_summary_sheet(self, branches: List[PartyBranch], 
                                year: int, month: int) -> str:
        """
        生成党费汇总表
        
        包含各支部的党员人数和党费金额汇总
        """
        wb = Workbook()
        ws = wb.active
        ws.title = f"{year}年{month}月党费汇总表"
        
        # 生成表格标题
        title_line1 = f"{year}年{COLLEGE_NAME}"
        title_line2 = f"{month}月党费汇总表"
        
        # 定义列宽
        column_widths = [8, 30, 12, 12, 20]
        for i, width in enumerate(column_widths, 1):
            ws.column_dimensions[get_column_letter(i)].width = width
        
        # 第一行：标题第一部分
        ws.merge_cells('A1:E1')
        ws['A1'] = title_line1
        ws['A1'].font = self.title_font
        ws['A1'].alignment = self.center_alignment
        
        # 第二行：标题第二部分
        ws.merge_cells('A2:E2')
        ws['A2'] = title_line2
        ws['A2'].font = self.title_font
        ws['A2'].alignment = self.center_alignment
        
        # 第三行：列标题
        headers = [
            "序号", "支部名称", "党员人数", "金额", "备注"
        ]
        
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=3, column=col, value=header)
            cell.font = self.header_font
            cell.alignment = self.center_alignment
            cell.border = self.thin_border
            cell.fill = self.header_fill
        
        # 填充数据
        row = 4
        total_members = 0
        total_fee = 0.0
        
        for i, branch in enumerate(sorted(branches, key=lambda b: b.sequence), 1):
            branch_member_count = len(branch.members)
            branch_total_fee = sum(m.monthly_fee for m in branch.members)
            
            total_members += branch_member_count
            total_fee += branch_total_fee
            
            # 填充数据
            ws.cell(row=row, column=1, value=i).alignment = self.center_alignment
            ws.cell(row=row, column=2, value=branch.name).alignment = self.left_alignment
            ws.cell(row=row, column=3, value=branch_member_count).alignment = self.center_alignment
            ws.cell(row=row, column=4, value=branch_total_fee).alignment = self.center_alignment
            ws.cell(row=row, column=5, value="").alignment = self.left_alignment
            
            # 应用边框
            for col in range(1, 6):
                ws.cell(row=row, column=col).border = self.thin_border
                # 交替行背景色
                if row % 2 == 0:
                    ws.cell(row=row, column=col).fill = self.light_fill
            
            row += 1
        
        # 添加合计行
        total_row = row
        ws.merge_cells(f'A{total_row}:B{total_row}')
        ws.cell(row=total_row, column=1, value="合计").font = self.header_font
        ws.cell(row=total_row, column=1).alignment = self.center_alignment
        ws.cell(row=total_row, column=1).fill = self.header_fill
        
        ws.cell(row=total_row, column=3, value=total_members).font = self.header_font
        ws.cell(row=total_row, column=3).alignment = self.center_alignment
        ws.cell(row=total_row, column=3).fill = self.header_fill
        
        ws.cell(row=total_row, column=4, value=total_fee).font = self.header_font
        ws.cell(row=total_row, column=4).alignment = self.center_alignment
        ws.cell(row=total_row, column=4).fill = self.header_fill
        
        ws.cell(row=total_row, column=5, value="").fill = self.header_fill
        
        # 应用边框到合计行
        for col in range(1, 6):
            ws.cell(row=total_row, column=col).border = self.thin_border
        
        # 设置行高
        ws.row_dimensions[1].height = 30
        ws.row_dimensions[2].height = 30
        ws.row_dimensions[3].height = 25
        
        # 保存文件到月份目录
        month_dir = self._get_month_dir(year, month)
        filename = f"{year}年{month}月党费汇总表.xlsx"
        filepath = os.path.join(month_dir, filename)
        wb.save(filepath)
        
        return filepath
