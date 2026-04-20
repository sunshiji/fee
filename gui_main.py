# 党费收取系统图形界面
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, scrolledtext
from typing import List, Optional
import os
import glob

from config import DEFAULT_YEAR, DEFAULT_MONTH, OUTPUT_DIR, DATA_DIR
from data_models import PartyMember, PartyBranch, SalaryInfo, DeductionInfo
from member_manager import MemberManager
from fee_calculator import FeeCalculator
from excel_generator import ExcelGenerator
from statistics import FeeStatistics
from database_manager import DatabaseManager, DataComparator


class PartyFeeGUI:
    """党费收取系统图形界面主类"""
    
    def __init__(self, root: tk.Tk):
        """初始化GUI系统"""
        self.root = root
        self.root.title("党费收取系统 v1.0 - 可视化版")
        self.root.geometry("1200x800")
        self.root.minsize(1000, 700)
        
        # 初始化业务逻辑
        self.member_manager = MemberManager()
        self.excel_generator = ExcelGenerator()
        self.db_manager = DatabaseManager()
        self.data_comparator = DataComparator(self.db_manager)
        self.current_year = DEFAULT_YEAR
        self.current_month = DEFAULT_MONTH
        
        # 自动加载数据
        self._auto_load_data()
        
        # 创建界面
        self._create_menu()
        self._create_toolbar()
        self._create_main_content()
        self._create_status_bar()
        
        # 刷新显示
        self._refresh_display()
    
    def _auto_load_data(self):
        """自动加载最新的数据"""
        if not os.path.exists(DATA_DIR):
            return
        
        data_files = glob.glob(os.path.join(DATA_DIR, "party_fee_data_*.json"))
        if not data_files:
            return
        
        available_data = []
        for filepath in data_files:
            filename = os.path.basename(filepath)
            parts = filename.replace(".json", "").split("_")
            if len(parts) >= 5:
                try:
                    year = int(parts[3])
                    month = int(parts[4])
                    available_data.append((year, month, filepath))
                except (ValueError, IndexError):
                    continue
        
        if not available_data:
            return
        
        available_data.sort(key=lambda x: (x[0], x[1]), reverse=True)
        latest = available_data[0]
        
        self.current_year = latest[0]
        self.current_month = latest[1]
        
        if self.member_manager.load_from_file(latest[0], latest[1]):
            branches = self.member_manager.get_all_branches()
            FeeCalculator.calculate_all_fees(branches)
    
    def _create_menu(self):
        """创建菜单栏"""
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        
        # 文件菜单
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="文件(F)", menu=file_menu, underline=5)
        file_menu.add_command(label="新建月份数据", command=self._new_month, accelerator="Ctrl+N")
        file_menu.add_command(label="保存数据", command=self._save_data, accelerator="Ctrl+S")
        file_menu.add_command(label="加载数据", command=self._load_data, accelerator="Ctrl+L")
        file_menu.add_separator()
        file_menu.add_command(label="设置年月", command=self._set_year_month)
        file_menu.add_separator()
        file_menu.add_command(label="退出", command=self._on_close, accelerator="Ctrl+Q")
        
        # 数据管理菜单
        data_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="数据管理(D)", menu=data_menu, underline=5)
        data_menu.add_command(label="人员维护管理", command=self._show_member_management)
        data_menu.add_command(label="支部维护管理", command=self._show_branch_management)
        data_menu.add_separator()
        data_menu.add_command(label="计算党费", command=self._calculate_fees)
        data_menu.add_separator()
        data_menu.add_command(label="导入外部Excel数据", command=self._import_excel_data)
        
        # 报表菜单
        report_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="报表(R)", menu=report_menu, underline=3)
        report_menu.add_command(label="生成党费明细表", command=self._generate_detail_sheet)
        report_menu.add_command(label="生成各支部党费收缴子表", command=self._generate_branch_sheets)
        report_menu.add_command(label="生成党费汇总表", command=self._generate_summary_sheet)
        report_menu.add_separator()
        report_menu.add_command(label="生成所有报表", command=self._generate_all_reports)
        
        # 统计菜单
        stats_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="统计(S)", menu=stats_menu, underline=3)
        stats_menu.add_command(label="查看统计信息", command=self._show_statistics)
        stats_menu.add_separator()
        stats_menu.add_command(label="历史数据对比", command=self._show_data_comparison)
        
        # 数据库菜单
        db_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="数据库(D)", menu=db_menu, underline=3)
        db_menu.add_command(label="保存到数据库", command=self._save_to_database)
        db_menu.add_command(label="从数据库加载", command=self._load_from_database)
        db_menu.add_separator()
        db_menu.add_command(label="查看数据库中的月份", command=self._list_database_months)
        
        # 帮助菜单
        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="帮助(H)", menu=help_menu, underline=3)
        help_menu.add_command(label="关于", command=self._show_about)
        
        # 绑定快捷键
        self.root.bind("<Control-n>", lambda e: self._new_month())
        self.root.bind("<Control-s>", lambda e: self._save_data())
        self.root.bind("<Control-l>", lambda e: self._load_data())
        self.root.bind("<Control-q>", lambda e: self._on_close())
    
    def _create_toolbar(self):
        """创建工具栏"""
        toolbar = ttk.Frame(self.root)
        toolbar.pack(fill=tk.X, padx=5, pady=5)
        
        # 人员管理按钮
        ttk.Button(toolbar, text="人员管理", command=self._show_member_management).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="支部管理", command=self._show_branch_management).pack(side=tk.LEFT, padx=2)
        ttk.Separator(toolbar, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=5)
        
        # 业务功能按钮
        ttk.Button(toolbar, text="计算党费", command=self._calculate_fees).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="生成报表", command=self._show_generate_report_dialog).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="导入数据", command=self._import_excel_data).pack(side=tk.LEFT, padx=2)
        ttk.Separator(toolbar, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=5)
        
        # 数据操作按钮
        ttk.Button(toolbar, text="保存", command=self._save_data).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="加载", command=self._load_data).pack(side=tk.LEFT, padx=2)
        ttk.Separator(toolbar, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=5)
        
        # 年月显示和设置
        ttk.Label(toolbar, text="当前年月:").pack(side=tk.LEFT, padx=2)
        self.year_month_var = tk.StringVar(value=f"{self.current_year}年{self.current_month}月")
        ttk.Label(toolbar, textvariable=self.year_month_var, font=("Arial", 10, "bold")).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="设置", command=self._set_year_month).pack(side=tk.LEFT, padx=2)
    
    def _create_main_content(self):
        """创建主内容区域"""
        # 创建PanedWindow
        main_paned = ttk.PanedWindow(self.root, orient=tk.HORIZONTAL)
        main_paned.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 左侧：支部和党员列表
        left_frame = ttk.LabelFrame(main_paned, text="数据视图", padding=5)
        main_paned.add(left_frame, weight=1)
        
        # 支部树视图
        ttk.Label(left_frame, text="党支部列表:").pack(anchor=tk.W)
        self.branch_tree = ttk.Treeview(left_frame, columns=("name", "count", "total_fee"), show="tree headings", height=10)
        self.branch_tree.heading("name", text="支部名称")
        self.branch_tree.heading("count", text="党员人数")
        self.branch_tree.heading("total_fee", text="总党费(元)")
        self.branch_tree.column("name", width=200)
        self.branch_tree.column("count", width=80)
        self.branch_tree.column("total_fee", width=100)
        self.branch_tree.pack(fill=tk.BOTH, expand=True, pady=5)
        self.branch_tree.bind("<<TreeviewSelect>>", self._on_branch_select)
        
        # 党员列表
        ttk.Label(left_frame, text="党员列表:").pack(anchor=tk.W, pady=(10, 0))
        self.member_tree = ttk.Treeview(left_frame, columns=("name", "base", "fee"), show="headings", height=10)
        self.member_tree.heading("name", text="姓名")
        self.member_tree.heading("base", text="缴费基数(元)")
        self.member_tree.heading("fee", text="月党费(元)")
        self.member_tree.column("name", width=120)
        self.member_tree.column("base", width=120)
        self.member_tree.column("fee", width=100)
        self.member_tree.pack(fill=tk.BOTH, expand=True, pady=5)
        self.member_tree.bind("<<TreeviewSelect>>", self._on_member_select)
        
        # 右侧：详情和操作区域
        right_frame = ttk.LabelFrame(main_paned, text="详情面板", padding=5)
        main_paned.add(right_frame, weight=2)
        
        # 创建Notebook用于标签页
        self.notebook = ttk.Notebook(right_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True)
        
        # 党员详情标签页
        member_detail_frame = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(member_detail_frame, text="党员详情")
        self._create_member_detail_panel(member_detail_frame)
        
        # 操作日志标签页（先创建，因为其他面板可能需要调用_log()
        log_frame = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(log_frame, text="操作日志")
        self._create_log_panel(log_frame)
        
        # 统计信息标签页
        stats_frame = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(stats_frame, text="统计信息")
        self._create_stats_panel(stats_frame)
        
        # 数据对比标签页
        comparison_frame = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(comparison_frame, text="数据对比")
        self._create_comparison_panel(comparison_frame)
    
    def _create_member_detail_panel(self, parent):
        """创建党员详情面板"""
        # 基本信息
        basic_frame = ttk.LabelFrame(parent, text="基本信息", padding=10)
        basic_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(basic_frame, text="姓名:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=2)
        self.detail_name_var = tk.StringVar()
        ttk.Entry(basic_frame, textvariable=self.detail_name_var, state="readonly", width=30).grid(row=0, column=1, sticky=tk.W, padx=5, pady=2)
        
        ttk.Label(basic_frame, text="所属支部:").grid(row=0, column=2, sticky=tk.W, padx=5, pady=2)
        self.detail_branch_var = tk.StringVar()
        ttk.Entry(basic_frame, textvariable=self.detail_branch_var, state="readonly", width=30).grid(row=0, column=3, sticky=tk.W, padx=5, pady=2)
        
        ttk.Label(basic_frame, text="序号:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=2)
        self.detail_seq_var = tk.StringVar()
        ttk.Entry(basic_frame, textvariable=self.detail_seq_var, state="readonly", width=30).grid(row=1, column=1, sticky=tk.W, padx=5, pady=2)
        
        # 工资信息
        salary_frame = ttk.LabelFrame(parent, text="工资信息（元）", padding=10)
        salary_frame.pack(fill=tk.X, pady=5)
        
        self.salary_vars = {}
        salary_fields = [
            ("岗位工资", "position_salary"),
            ("薪级工资", "rank_salary"),
            ("高定工资", "fixed_salary"),
            ("基础性绩效", "basic_performance"),
        ]
        
        for i, (label, field) in enumerate(salary_fields):
            ttk.Label(salary_frame, text=f"{label}:").grid(row=i, column=0, sticky=tk.W, padx=5, pady=2)
            var = tk.StringVar()
            self.salary_vars[field] = var
            ttk.Entry(salary_frame, textvariable=var, width=20).grid(row=i, column=1, sticky=tk.W, padx=5, pady=2)
        
        # 扣款信息
        deduction_frame = ttk.LabelFrame(parent, text="扣款信息（元）", padding=10)
        deduction_frame.pack(fill=tk.X, pady=5)
        
        self.deduction_vars = {}
        deduction_fields = [
            ("住房公积金", "housing_fund"),
            ("医疗保险", "medical_insurance"),
            ("养老保险", "pension_insurance"),
            ("职业年金", "occupational_annuity"),
            ("大额医疗", "large_medical"),
            ("失业保险", "unemployment_insurance"),
            ("个人所得税", "personal_income_tax"),
        ]
        
        for i, (label, field) in enumerate(deduction_fields):
            row = i // 2
            col = (i % 2) * 2
            ttk.Label(deduction_frame, text=f"{label}:").grid(row=row, column=col, sticky=tk.W, padx=5, pady=2)
            var = tk.StringVar()
            self.deduction_vars[field] = var
            ttk.Entry(deduction_frame, textvariable=var, width=20).grid(row=row, column=col+1, sticky=tk.W, padx=5, pady=2)
        
        # 计算结果
        result_frame = ttk.LabelFrame(parent, text="计算结果", padding=10)
        result_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(result_frame, text="缴费基数:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=2)
        self.detail_base_var = tk.StringVar()
        ttk.Entry(result_frame, textvariable=self.detail_base_var, state="readonly", width=20, font=("Arial", 10, "bold")).grid(row=0, column=1, sticky=tk.W, padx=5, pady=2)
        
        ttk.Label(result_frame, text="月党费:").grid(row=0, column=2, sticky=tk.W, padx=5, pady=2)
        self.detail_fee_var = tk.StringVar()
        ttk.Entry(result_frame, textvariable=self.detail_fee_var, state="readonly", width=20, font=("Arial", 10, "bold"), foreground="red").grid(row=0, column=3, sticky=tk.W, padx=5, pady=2)
        
        # 操作按钮
        btn_frame = ttk.Frame(parent)
        btn_frame.pack(fill=tk.X, pady=10)
        
        ttk.Button(btn_frame, text="添加党员", command=self._add_member_dialog).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="修改党员", command=self._update_member).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="删除党员", command=self._delete_member).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="重新计算", command=self._recalculate_member).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="清空表单", command=self._clear_member_detail).pack(side=tk.LEFT, padx=5)
    
    def _create_stats_panel(self, parent):
        """创建统计信息面板"""
        self.stats_text = scrolledtext.ScrolledText(parent, wrap=tk.WORD, font=("Consolas", 10))
        self.stats_text.pack(fill=tk.BOTH, expand=True)
    
    def _create_log_panel(self, parent):
        """创建日志面板"""
        self.log_text = scrolledtext.ScrolledText(parent, wrap=tk.WORD, font=("Consolas", 9), height=15)
        self.log_text.pack(fill=tk.BOTH, expand=True)
        self._log("系统启动完成")
    
    def _create_status_bar(self):
        """创建状态栏"""
        self.status_var = tk.StringVar(value="就绪")
        status_bar = ttk.Label(self.root, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)
        
        # 统计信息状态
        self.stats_status_var = tk.StringVar(value="支部: 0 | 党员: 0 | 总党费: 0.00元")
        stats_status = ttk.Label(self.root, textvariable=self.stats_status_var, relief=tk.SUNKEN, anchor=tk.E)
        stats_status.pack(side=tk.BOTTOM, fill=tk.X)
    
    def _log(self, message: str):
        """记录日志"""
        # 安全检查：确保log_text存在
        if not hasattr(self, 'log_text') or self.log_text is None:
            return
        
        from datetime import datetime
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        try:
            self.log_text.insert(tk.END, f"[{timestamp}] {message}\n")
            self.log_text.see(tk.END)
        except Exception:
            # 忽略日志记录时的错误
            pass
    
    def _refresh_display(self, preserve_selection: bool = True):
        """
        刷新显示
        
        Args:
            preserve_selection: 是否保持当前选中的支部和党员
        """
        # 保存当前选中的支部
        selected_branch = None
        if preserve_selection:
            selection = self.branch_tree.selection()
            if selection:
                selected_branch = selection[0]
        
        # 更新年月显示
        self.year_month_var.set(f"{self.current_year}年{self.current_month}月")
        
        # 清空支部树
        for item in self.branch_tree.get_children():
            self.branch_tree.delete(item)
        
        # 清空党员树
        for item in self.member_tree.get_children():
            self.member_tree.delete(item)
        
        # 重新计算党费
        branches = self.member_manager.get_all_branches()
        FeeCalculator.calculate_all_fees(branches)
        
        # 填充支部树
        for branch in branches:
            branch_total = sum(m.monthly_fee for m in branch.members)
            self.branch_tree.insert("", tk.END, iid=branch.name, values=(branch.name, len(branch.members), f"{branch_total:.2f}"))
        
        # 恢复选中的支部并填充党员列表
        if preserve_selection and selected_branch:
            # 检查该支部是否还存在
            if selected_branch in self.branch_tree.get_children():
                self.branch_tree.selection_set(selected_branch)
                # 手动触发支部选择事件，填充党员列表
                self._on_branch_select(None)
        
        # 更新统计状态栏
        total_branches = len(branches)
        total_members = sum(len(b.members) for b in branches)
        total_fee = sum(sum(m.monthly_fee for m in b.members) for b in branches)
        self.stats_status_var.set(f"支部: {total_branches} | 党员: {total_members} | 总党费: {total_fee:.2f}元")
        
        # 更新统计面板
        self._update_stats_panel()
    
    def _update_stats_panel(self):
        """更新统计面板"""
        branches = self.member_manager.get_all_branches()
        if not branches:
            self.stats_text.delete(1.0, tk.END)
            self.stats_text.insert(tk.END, "暂无数据")
            return
        
        report = FeeStatistics.generate_statistics_report(branches, self.current_year, self.current_month)
        self.stats_text.delete(1.0, tk.END)
        self.stats_text.insert(tk.END, report)
    
    def _on_branch_select(self, event):
        """支部选择事件"""
        selection = self.branch_tree.selection()
        if not selection:
            return
        
        branch_name = selection[0]
        
        # 清空党员树
        for item in self.member_tree.get_children():
            self.member_tree.delete(item)
        
        # 填充党员树
        branch = self.member_manager.get_branch(branch_name)
        if branch:
            for member in branch.members:
                self.member_tree.insert("", tk.END, iid=f"{branch_name}_{member.name}", 
                                         values=(member.name, f"{member.payment_base:.2f}", f"{member.monthly_fee:.2f}"))
    
    def _on_member_select(self, event):
        """党员选择事件"""
        selection = self.member_tree.selection()
        if not selection:
            return
        
        item_id = selection[0]
        parts = item_id.split("_", 1)
        if len(parts) < 2:
            return
        
        branch_name = parts[0]
        member_name = parts[1]
        
        member = self.member_manager.get_member(member_name, branch_name)
        if member:
            self._fill_member_detail(member)
    
    def _fill_member_detail(self, member: PartyMember):
        """填充党员详情"""
        self.detail_name_var.set(member.name)
        self.detail_branch_var.set(member.branch_name)
        self.detail_seq_var.set(str(member.sequence))
        
        # 工资信息
        self.salary_vars["position_salary"].set(f"{member.salary_info.position_salary:.2f}")
        self.salary_vars["rank_salary"].set(f"{member.salary_info.rank_salary:.2f}")
        self.salary_vars["fixed_salary"].set(f"{member.salary_info.fixed_salary:.2f}")
        self.salary_vars["basic_performance"].set(f"{member.salary_info.basic_performance:.2f}")
        
        # 扣款信息
        self.deduction_vars["housing_fund"].set(f"{member.deduction_info.housing_fund:.2f}")
        self.deduction_vars["medical_insurance"].set(f"{member.deduction_info.medical_insurance:.2f}")
        self.deduction_vars["pension_insurance"].set(f"{member.deduction_info.pension_insurance:.2f}")
        self.deduction_vars["occupational_annuity"].set(f"{member.deduction_info.occupational_annuity:.2f}")
        self.deduction_vars["large_medical"].set(f"{member.deduction_info.large_medical:.2f}")
        self.deduction_vars["unemployment_insurance"].set(f"{member.deduction_info.unemployment_insurance:.2f}")
        self.deduction_vars["personal_income_tax"].set(f"{member.deduction_info.personal_income_tax:.2f}")
        
        # 计算结果
        self.detail_base_var.set(f"{member.payment_base:.2f}")
        self.detail_fee_var.set(f"{member.monthly_fee:.2f}")
    
    def _clear_member_detail(self):
        """清空党员详情表单"""
        self.detail_name_var.set("")
        self.detail_branch_var.set("")
        self.detail_seq_var.set("")
        
        for var in self.salary_vars.values():
            var.set("0.00")
        
        for var in self.deduction_vars.values():
            var.set("0.00")
        
        self.detail_base_var.set("")
        self.detail_fee_var.set("")
        
        self._log("已清空表单")
    
    def _show_member_management(self):
        """显示人员管理界面"""
        self.notebook.select(0)
        self._log("切换到人员管理")
    
    def _show_branch_management(self):
        """显示支部管理对话框"""
        dialog = tk.Toplevel(self.root)
        dialog.title("支部维护管理")
        dialog.geometry("600x500")
        dialog.transient(self.root)
        dialog.grab_set()
        
        # 支部列表
        list_frame = ttk.LabelFrame(dialog, text="党支部列表", padding=10)
        list_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        columns = ("name", "sequence", "count", "total_fee")
        tree = ttk.Treeview(list_frame, columns=columns, show="headings", height=15)
        tree.heading("name", text="支部名称")
        tree.heading("sequence", text="支部序号")
        tree.heading("count", text="党员人数")
        tree.heading("total_fee", text="总党费(元)")
        tree.column("name", width=250)
        tree.column("sequence", width=80)
        tree.column("count", width=80)
        tree.column("total_fee", width=100)
        tree.pack(fill=tk.BOTH, expand=True)
        
        # 填充数据
        branches = self.member_manager.get_all_branches()
        for branch in branches:
            branch_total = sum(m.monthly_fee for m in branch.members)
            tree.insert("", tk.END, values=(branch.name, branch.sequence, len(branch.members), f"{branch_total:.2f}"))
        
        # 操作按钮
        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(fill=tk.X, padx=10, pady=10)
        
        ttk.Button(btn_frame, text="添加支部", command=lambda: self._add_branch_dialog(tree)).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="删除支部", command=lambda: self._delete_branch_dialog(tree)).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="关闭", command=dialog.destroy).pack(side=tk.RIGHT, padx=5)
    
    def _add_branch_dialog(self, tree: ttk.Treeview):
        """添加支部对话框"""
        add_dialog = tk.Toplevel(self.root)
        add_dialog.title("添加党支部")
        add_dialog.geometry("400x150")
        add_dialog.transient(self.root)
        add_dialog.grab_set()
        
        ttk.Label(add_dialog, text="支部名称:").pack(pady=5)
        name_var = tk.StringVar()
        ttk.Entry(add_dialog, textvariable=name_var, width=40).pack(pady=5)
        
        ttk.Label(add_dialog, text="支部序号:").pack(pady=5)
        seq_var = tk.StringVar()
        ttk.Entry(add_dialog, textvariable=seq_var, width=20).pack(pady=5)
        
        def do_add():
            name = name_var.get().strip()
            seq = seq_var.get().strip()
            
            if not name:
                messagebox.showerror("错误", "支部名称不能为空！")
                return
            
            if not seq:
                messagebox.showerror("错误", "支部序号不能为空！")
                return
            
            try:
                seq_int = int(seq)
            except ValueError:
                messagebox.showerror("错误", "支部序号必须是整数！")
                return
            
            try:
                from data_models import PartyBranch
                branch = PartyBranch(name=name, sequence=seq_int)
                self.member_manager.add_branch(branch)
                self._refresh_display()
                
                # 更新列表
                for item in tree.get_children():
                    tree.delete(item)
                branches = self.member_manager.get_all_branches()
                for b in branches:
                    b_total = sum(m.monthly_fee for m in b.members)
                    tree.insert("", tk.END, values=(b.name, b.sequence, len(b.members), f"{b_total:.2f}"))
                
                self._log(f"添加党支部: {name}")
                messagebox.showinfo("成功", f"党支部 '{name}' 添加成功！")
                add_dialog.destroy()
            except ValueError as e:
                messagebox.showerror("错误", str(e))
        
        btn_frame = ttk.Frame(add_dialog)
        btn_frame.pack(pady=10)
        ttk.Button(btn_frame, text="确定", command=do_add).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="取消", command=add_dialog.destroy).pack(side=tk.LEFT, padx=5)
    
    def _delete_branch_dialog(self, tree: ttk.Treeview):
        """删除支部对话框"""
        selection = tree.selection()
        if not selection:
            messagebox.showwarning("警告", "请先选择要删除的支部！")
            return
        
        item = tree.item(selection[0])
        branch_name = item["values"][0]
        
        if not messagebox.askyesno("确认删除", f"确定要删除党支部 '{branch_name}' 吗？\n\n该操作将同时删除支部内的所有党员！"):
            return
        
        if self.member_manager.remove_branch(branch_name):
            self._refresh_display()
            
            # 更新列表
            for item in tree.get_children():
                tree.delete(item)
            branches = self.member_manager.get_all_branches()
            for b in branches:
                b_total = sum(m.monthly_fee for m in b.members)
                tree.insert("", tk.END, values=(b.name, b.sequence, len(b.members), f"{b_total:.2f}"))
            
            self._log(f"删除党支部: {branch_name}")
            messagebox.showinfo("成功", f"党支部 '{branch_name}' 删除成功！")
        else:
            messagebox.showerror("错误", "删除失败！")
    
    def _add_member_dialog(self):
        """添加党员对话框"""
        branches = self.member_manager.get_all_branches()
        if not branches:
            messagebox.showwarning("警告", "请先添加党支部！")
            return
        
        dialog = tk.Toplevel(self.root)
        dialog.title("添加党员")
        dialog.geometry("500x600")
        dialog.transient(self.root)
        dialog.grab_set()
        
        # 基本信息
        basic_frame = ttk.LabelFrame(dialog, text="基本信息", padding=10)
        basic_frame.pack(fill=tk.X, padx=10, pady=5)
        
        ttk.Label(basic_frame, text="姓名:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=2)
        name_var = tk.StringVar()
        ttk.Entry(basic_frame, textvariable=name_var, width=30).grid(row=0, column=1, sticky=tk.W, padx=5, pady=2)
        
        ttk.Label(basic_frame, text="所属支部:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=2)
        branch_var = tk.StringVar()
        branch_combo = ttk.Combobox(basic_frame, textvariable=branch_var, state="readonly", width=28)
        branch_combo["values"] = [b.name for b in branches]
        if branches:
            branch_combo.current(0)
        branch_combo.grid(row=1, column=1, sticky=tk.W, padx=5, pady=2)
        
        # 工资信息
        salary_frame = ttk.LabelFrame(dialog, text="工资信息（元）", padding=10)
        salary_frame.pack(fill=tk.X, padx=10, pady=5)
        
        salary_vars = {}
        salary_fields = [
            ("岗位工资", "position_salary"),
            ("薪级工资", "rank_salary"),
            ("高定工资", "fixed_salary"),
            ("基础性绩效", "basic_performance"),
        ]
        
        for i, (label, field) in enumerate(salary_fields):
            ttk.Label(salary_frame, text=f"{label}:").grid(row=i, column=0, sticky=tk.W, padx=5, pady=2)
            var = tk.StringVar(value="0.00")
            salary_vars[field] = var
            ttk.Entry(salary_frame, textvariable=var, width=20).grid(row=i, column=1, sticky=tk.W, padx=5, pady=2)
        
        # 扣款信息
        deduction_frame = ttk.LabelFrame(dialog, text="扣款信息（元）", padding=10)
        deduction_frame.pack(fill=tk.X, padx=10, pady=5)
        
        deduction_vars = {}
        deduction_fields = [
            ("住房公积金", "housing_fund"),
            ("医疗保险", "medical_insurance"),
            ("养老保险", "pension_insurance"),
            ("职业年金", "occupational_annuity"),
            ("大额医疗", "large_medical"),
            ("失业保险", "unemployment_insurance"),
            ("个人所得税", "personal_income_tax"),
        ]
        
        for i, (label, field) in enumerate(deduction_fields):
            row = i // 2
            col = (i % 2) * 2
            ttk.Label(deduction_frame, text=f"{label}:").grid(row=row, column=col, sticky=tk.W, padx=5, pady=2)
            var = tk.StringVar(value="0.00")
            deduction_vars[field] = var
            ttk.Entry(deduction_frame, textvariable=var, width=20).grid(row=row, column=col+1, sticky=tk.W, padx=5, pady=2)
        
        def do_add():
            name = name_var.get().strip()
            branch_name = branch_var.get()
            
            if not name:
                messagebox.showerror("错误", "姓名不能为空！")
                return
            
            if not branch_name:
                messagebox.showerror("错误", "请选择所属支部！")
                return
            
            try:
                # 解析工资信息
                salary_info = SalaryInfo(
                    position_salary=float(salary_vars["position_salary"].get() or 0),
                    rank_salary=float(salary_vars["rank_salary"].get() or 0),
                    fixed_salary=float(salary_vars["fixed_salary"].get() or 0),
                    basic_performance=float(salary_vars["basic_performance"].get() or 0)
                )
                
                # 解析扣款信息
                deduction_info = DeductionInfo(
                    housing_fund=float(deduction_vars["housing_fund"].get() or 0),
                    medical_insurance=float(deduction_vars["medical_insurance"].get() or 0),
                    pension_insurance=float(deduction_vars["pension_insurance"].get() or 0),
                    occupational_annuity=float(deduction_vars["occupational_annuity"].get() or 0),
                    large_medical=float(deduction_vars["large_medical"].get() or 0),
                    unemployment_insurance=float(deduction_vars["unemployment_insurance"].get() or 0),
                    personal_income_tax=float(deduction_vars["personal_income_tax"].get() or 0)
                )
                
                member = PartyMember(
                    name=name,
                    salary_info=salary_info,
                    deduction_info=deduction_info
                )
                
                self.member_manager.add_member(member, branch_name)
                FeeCalculator.calculate_member_fee(member)
                
                self._refresh_display()
                self._log(f"添加党员: {name} (支部: {branch_name})")
                messagebox.showinfo("成功", f"党员 '{name}' 添加成功！\n缴费基数: {member.payment_base:.2f} 元\n月党费: {member.monthly_fee:.2f} 元")
                dialog.destroy()
                
            except ValueError as e:
                messagebox.showerror("错误", str(e))
            except Exception as e:
                messagebox.showerror("错误", f"添加失败: {e}")
        
        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(pady=15)
        ttk.Button(btn_frame, text="确定添加", command=do_add).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="取消", command=dialog.destroy).pack(side=tk.LEFT, padx=5)
    
    def _update_member(self):
        """修改党员信息"""
        name = self.detail_name_var.get()
        branch_name = self.detail_branch_var.get()
        
        if not name or not branch_name:
            messagebox.showwarning("警告", "请先选择要修改的党员！")
            return
        
        member = self.member_manager.get_member(name, branch_name)
        if not member:
            messagebox.showerror("错误", "党员不存在！")
            return
        
        try:
            # 解析工资信息
            salary_info = SalaryInfo(
                position_salary=float(self.salary_vars["position_salary"].get() or 0),
                rank_salary=float(self.salary_vars["rank_salary"].get() or 0),
                fixed_salary=float(self.salary_vars["fixed_salary"].get() or 0),
                basic_performance=float(self.salary_vars["basic_performance"].get() or 0)
            )
            
            # 解析扣款信息
            deduction_info = DeductionInfo(
                housing_fund=float(self.deduction_vars["housing_fund"].get() or 0),
                medical_insurance=float(self.deduction_vars["medical_insurance"].get() or 0),
                pension_insurance=float(self.deduction_vars["pension_insurance"].get() or 0),
                occupational_annuity=float(self.deduction_vars["occupational_annuity"].get() or 0),
                large_medical=float(self.deduction_vars["large_medical"].get() or 0),
                unemployment_insurance=float(self.deduction_vars["unemployment_insurance"].get() or 0),
                personal_income_tax=float(self.deduction_vars["personal_income_tax"].get() or 0)
            )
            
            if self.member_manager.update_member(name, branch_name, salary_info, deduction_info):
                # 重新计算党费
                FeeCalculator.calculate_member_fee(member)
                
                # 更新显示
                self._fill_member_detail(member)
                self._refresh_display()
                
                self._log(f"修改党员信息: {name}")
                messagebox.showinfo("成功", f"党员 '{name}' 信息更新成功！\n新缴费基数: {member.payment_base:.2f} 元\n新月党费: {member.monthly_fee:.2f} 元")
            else:
                messagebox.showerror("错误", "更新失败！")
                
        except ValueError as e:
            messagebox.showerror("错误", f"请输入有效的数值: {e}")
    
    def _delete_member(self):
        """删除党员"""
        name = self.detail_name_var.get()
        branch_name = self.detail_branch_var.get()
        
        if not name or not branch_name:
            messagebox.showwarning("警告", "请先选择要删除的党员！")
            return
        
        if not messagebox.askyesno("确认删除", f"确定要删除党员 '{name}' 吗？"):
            return
        
        if self.member_manager.remove_member(name, branch_name):
            self._clear_member_detail()
            self._refresh_display()
            self._log(f"删除党员: {name}")
            messagebox.showinfo("成功", f"党员 '{name}' 删除成功！")
        else:
            messagebox.showerror("错误", "删除失败！")
    
    def _recalculate_member(self):
        """重新计算选中党员的党费"""
        name = self.detail_name_var.get()
        branch_name = self.detail_branch_var.get()
        
        if not name or not branch_name:
            messagebox.showwarning("警告", "请先选择党员！")
            return
        
        # 先更新表单中的值到对象
        try:
            salary_info = SalaryInfo(
                position_salary=float(self.salary_vars["position_salary"].get() or 0),
                rank_salary=float(self.salary_vars["rank_salary"].get() or 0),
                fixed_salary=float(self.salary_vars["fixed_salary"].get() or 0),
                basic_performance=float(self.salary_vars["basic_performance"].get() or 0)
            )
            
            deduction_info = DeductionInfo(
                housing_fund=float(self.deduction_vars["housing_fund"].get() or 0),
                medical_insurance=float(self.deduction_vars["medical_insurance"].get() or 0),
                pension_insurance=float(self.deduction_vars["pension_insurance"].get() or 0),
                occupational_annuity=float(self.deduction_vars["occupational_annuity"].get() or 0),
                large_medical=float(self.deduction_vars["large_medical"].get() or 0),
                unemployment_insurance=float(self.deduction_vars["unemployment_insurance"].get() or 0),
                personal_income_tax=float(self.deduction_vars["personal_income_tax"].get() or 0)
            )
            
            member = self.member_manager.get_member(name, branch_name)
            if member:
                member.salary_info = salary_info
                member.deduction_info = deduction_info
                FeeCalculator.calculate_member_fee(member)
                
                self._fill_member_detail(member)
                self._refresh_display()
                
                self._log(f"重新计算党费: {name}")
                messagebox.showinfo("计算完成", f"党员 '{name}' 党费重新计算完成！\n缴费基数: {member.payment_base:.2f} 元\n月党费: {member.monthly_fee:.2f} 元")
                
        except ValueError as e:
            messagebox.showerror("错误", f"请输入有效的数值: {e}")
    
    def _calculate_fees(self):
        """计算所有党费"""
        branches = self.member_manager.get_all_branches()
        if not branches:
            messagebox.showwarning("警告", "暂无党支部数据！")
            return
        
        if not any(b.members for b in branches):
            messagebox.showwarning("警告", "暂无党员数据！")
            return
        
        FeeCalculator.calculate_all_fees(branches)
        self._refresh_display()
        
        total_members = sum(len(b.members) for b in branches)
        total_fee = sum(sum(m.monthly_fee for m in b.members) for b in branches)
        
        self._log(f"计算所有党费: {total_members} 名党员, 总党费 {total_fee:.2f} 元")
        messagebox.showinfo("计算完成", f"党费计算完成！\n\n党员人数: {total_members}\n总党费金额: {total_fee:.2f} 元")
    
    def _generate_detail_sheet(self):
        """生成党费明细表"""
        branches = self.member_manager.get_all_branches()
        if not branches:
            messagebox.showwarning("警告", "暂无党支部数据！")
            return
        
        try:
            filepath = self.excel_generator.generate_fee_detail_sheet(branches, self.current_year, self.current_month)
            self._log(f"生成党费明细表: {filepath}")
            messagebox.showinfo("成功", f"党费明细表已成功生成！\n\n文件路径: {filepath}")
        except Exception as e:
            messagebox.showerror("错误", f"生成失败: {e}")
    
    def _generate_branch_sheets(self):
        """生成各支部党费收缴子表"""
        branches = self.member_manager.get_all_branches()
        if not branches:
            messagebox.showwarning("警告", "暂无党支部数据！")
            return
        
        try:
            filepaths = self.excel_generator.generate_all_branch_sheets(branches, self.current_year, self.current_month)
            self._log(f"生成各支部党费收缴子表: {len(filepaths)} 个文件")
            messagebox.showinfo("成功", f"成功生成 {len(filepaths)} 个支部党费收缴子表！")
        except Exception as e:
            messagebox.showerror("错误", f"生成失败: {e}")
    
    def _generate_summary_sheet(self):
        """生成党费汇总表"""
        branches = self.member_manager.get_all_branches()
        if not branches:
            messagebox.showwarning("警告", "暂无党支部数据！")
            return
        
        try:
            filepath = self.excel_generator.generate_summary_sheet(branches, self.current_year, self.current_month)
            self._log(f"生成党费汇总表: {filepath}")
            messagebox.showinfo("成功", f"党费汇总表已成功生成！\n\n文件路径: {filepath}")
        except Exception as e:
            messagebox.showerror("错误", f"生成失败: {e}")
    
    def _show_generate_report_dialog(self):
        """显示报表生成选择对话框"""
        branches = self.member_manager.get_all_branches()
        if not branches:
            messagebox.showwarning("警告", "暂无党支部数据！")
            return
        
        # 创建对话框
        dialog = tk.Toplevel(self.root)
        dialog.title("生成报表")
        dialog.geometry("580x620")
        dialog.transient(self.root)
        dialog.grab_set()
        
        # 选项变量
        self._var_detail = tk.BooleanVar(value=True)
        self._var_branch = tk.BooleanVar(value=True)
        self._var_summary = tk.BooleanVar(value=True)
        
        # 学院名称选项
        self._var_college_name_option = tk.StringVar(value="short")  # short / full / custom
        self._var_custom_college_name = tk.StringVar(value="计算机科学与技术学院")
        
        # 月份格式选项
        self._var_month_format = tk.StringVar(value="chinese")  # chinese / number
        
        # 输出路径
        self._output_path_var = tk.StringVar(value="")  # 空表示使用默认路径
        
        # ============== 第一部分：报表类型选择 ==============
        type_frame = ttk.LabelFrame(dialog, text="第一步：选择要生成的报表类型", padding=10)
        type_frame.pack(fill=tk.X, padx=10, pady=5)
        
        ttk.Checkbutton(type_frame, text="党费明细表（所有党员汇总）", variable=self._var_detail).pack(anchor=tk.W, pady=2)
        ttk.Checkbutton(type_frame, text="各支部党费收缴子表（每个支部单独文件）", variable=self._var_branch).pack(anchor=tk.W, pady=2)
        ttk.Checkbutton(type_frame, text="党费汇总表（统计汇总）", variable=self._var_summary).pack(anchor=tk.W, pady=2)
        
        # ============== 第二部分：表头样式设置 ==============
        header_frame = ttk.LabelFrame(dialog, text="第二步：表头样式设置", padding=10)
        header_frame.pack(fill=tk.X, padx=10, pady=5)
        
        # 学院名称选项
        college_frame = ttk.LabelFrame(header_frame, text="学院名称显示", padding=5)
        college_frame.pack(fill=tk.X, pady=5)
        
        ttk.Radiobutton(college_frame, text="简化名称（如：计算机科学与技术学院）", 
                        variable=self._var_college_name_option, value="short").pack(anchor=tk.W, pady=2)
        ttk.Radiobutton(college_frame, text="完整名称（如：新疆大学计算机科学与技术学院）", 
                        variable=self._var_college_name_option, value="full").pack(anchor=tk.W, pady=2)
        
        custom_frame = ttk.Frame(college_frame)
        custom_frame.pack(fill=tk.X, pady=2)
        ttk.Radiobutton(custom_frame, text="自定义:", 
                        variable=self._var_college_name_option, value="custom").pack(side=tk.LEFT)
        ttk.Entry(custom_frame, textvariable=self._var_custom_college_name, width=30).pack(side=tk.LEFT, padx=5)
        
        # 月份格式选项
        month_frame = ttk.LabelFrame(header_frame, text="月份显示格式", padding=5)
        month_frame.pack(fill=tk.X, pady=5)
        
        ttk.Radiobutton(month_frame, text="中文月份（如：二月、十二月）", 
                        variable=self._var_month_format, value="chinese").pack(side=tk.LEFT, padx=20, pady=2)
        ttk.Radiobutton(month_frame, text="数字月份（如：2月、12月）", 
                        variable=self._var_month_format, value="number").pack(side=tk.LEFT, padx=20, pady=2)
        
        # 表头预览
        preview_frame = ttk.LabelFrame(header_frame, text="表头预览（基于当前设置）", padding=5)
        preview_frame.pack(fill=tk.X, pady=5)
        
        # 预览标签
        self._preview_detail_var = tk.StringVar(value="")
        self._preview_branch_var = tk.StringVar(value="")
        
        ttk.Label(preview_frame, text="明细表表头:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=2)
        ttk.Label(preview_frame, textvariable=self._preview_detail_var, 
                  foreground="blue", font=("Arial", 9, "bold")).grid(row=0, column=1, sticky=tk.W, padx=5, pady=2)
        
        ttk.Label(preview_frame, text="子表表头示例:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=2)
        ttk.Label(preview_frame, textvariable=self._preview_branch_var, 
                  foreground="blue", font=("Arial", 9, "bold")).grid(row=1, column=1, sticky=tk.W, padx=5, pady=2)
        
        # ============== 第三部分：输出路径选择 ==============
        path_frame = ttk.LabelFrame(dialog, text="第三步：输出路径（留空使用默认路径）", padding=10)
        path_frame.pack(fill=tk.X, padx=10, pady=5)
        
        default_path = os.path.join(OUTPUT_DIR, f"{self.current_year}年{self.current_month}月")
        
        path_entry_frame = ttk.Frame(path_frame)
        path_entry_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(path_entry_frame, text="自定义路径:").pack(side=tk.LEFT, padx=5)
        path_entry = ttk.Entry(path_entry_frame, textvariable=self._output_path_var, width=45)
        path_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        def browse_path():
            selected_path = filedialog.askdirectory(
                title="选择输出目录",
                initialdir=OUTPUT_DIR
            )
            if selected_path:
                self._output_path_var.set(selected_path)
        
        ttk.Button(path_entry_frame, text="浏览...", command=browse_path).pack(side=tk.LEFT, padx=5)
        
        # 默认路径提示
        ttk.Label(path_frame, text=f"默认路径: {default_path}", foreground="gray").pack(anchor=tk.W)
        
        # ============== 更新预览 ==============
        def update_preview(*args):
            # 获取学院名称
            if self._var_college_name_option.get() == "short":
                # 简化名称：去掉"新疆大学"前缀
                college_name = COLLEGE_NAME.replace("新疆大学", "")
            elif self._var_college_name_option.get() == "full":
                college_name = COLLEGE_NAME
            else:
                college_name = self._var_custom_college_name.get().strip() or "学院"
            
            # 获取月份格式
            if self._var_month_format.get() == "chinese":
                from excel_generator import get_chinese_month
                month_str = f"{get_chinese_month(self.current_month)}月"
            else:
                month_str = f"{self.current_month}月"
            
            # 生成预览
            detail_preview = f"{college_name}{self.current_year}年{month_str}党费收缴明细表"
            branch_preview = f"{self.current_year}年{college_name}本科生第一党支部{month_str}党费收缴明细"
            
            self._preview_detail_var.set(detail_preview)
            self._preview_branch_var.set(branch_preview)
        
        # 绑定变量变化事件
        self._var_college_name_option.trace_add("write", update_preview)
        self._var_custom_college_name.trace_add("write", update_preview)
        self._var_month_format.trace_add("write", update_preview)
        
        # 初始更新预览
        update_preview()
        
        # ============== 按钮 ==============
        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(pady=15)
        
        def do_generate():
            # 检查是否至少选择了一个报表类型
            if not any([self._var_detail.get(), self._var_branch.get(), self._var_summary.get()]):
                messagebox.showwarning("警告", "请至少选择一种报表类型！")
                return
            
            # 获取学院名称
            if self._var_college_name_option.get() == "short":
                college_name = COLLEGE_NAME.replace("新疆大学", "")
            elif self._var_college_name_option.get() == "full":
                college_name = COLLEGE_NAME
            else:
                college_name = self._var_custom_college_name.get().strip() or COLLEGE_NAME
            
            # 是否使用中文月份
            use_chinese_month = (self._var_month_format.get() == "chinese")
            
            # 关闭对话框
            dialog.destroy()
            
            # 执行生成
            self._generate_selected_reports(
                generate_detail=self._var_detail.get(),
                generate_branch=self._var_branch.get(),
                generate_summary=self._var_summary.get(),
                output_path=self._output_path_var.get().strip(),
                custom_college_name=college_name,
                use_chinese_month=use_chinese_month
            )
        
        ttk.Button(btn_frame, text="确定生成", command=do_generate, 
                   width=15, style='Accent.TButton').pack(side=tk.LEFT, padx=15)
        ttk.Button(btn_frame, text="取消", command=dialog.destroy, width=15).pack(side=tk.LEFT, padx=15)
    
    def _generate_selected_reports(self, generate_detail: bool, generate_branch: bool, 
                                    generate_summary: bool, output_path: str = "",
                                    custom_college_name: Optional[str] = None,
                                    use_chinese_month: bool = False):
        """
        生成选中的报表
        
        Args:
            generate_detail: 是否生成明细表
            generate_branch: 是否生成子表
            generate_summary: 是否生成汇总表
            output_path: 输出路径，空字符串表示使用默认路径
            custom_college_name: 自定义学院名称
            use_chinese_month: 是否使用中文月份
        """
        import shutil
        
        branches = self.member_manager.get_all_branches()
        
        self.status_var.set("正在生成报表...")
        self.root.update()
        
        try:
            # 先计算党费
            FeeCalculator.calculate_all_fees(branches)
            
            # 确定输出路径
            use_default_path = (output_path == "")
            
            # 生成选中的报表到默认路径
            generated_files = []
            
            if generate_detail:
                filepath = self.excel_generator.generate_fee_detail_sheet(
                    branches, self.current_year, self.current_month,
                    custom_college_name=custom_college_name,
                    use_chinese_month=use_chinese_month
                )
                generated_files.append(filepath)
                self._log(f"生成党费明细表: {filepath}")
            
            if generate_branch:
                filepaths = self.excel_generator.generate_all_branch_sheets(
                    branches, self.current_year, self.current_month,
                    custom_college_name=custom_college_name,
                    use_chinese_month=use_chinese_month
                )
                generated_files.extend(filepaths)
                self._log(f"生成各支部党费收缴子表: {len(filepaths)} 个文件")
            
            if generate_summary:
                filepath = self.excel_generator.generate_summary_sheet(branches, self.current_year, self.current_month)
                generated_files.append(filepath)
                self._log(f"生成党费汇总表: {filepath}")
            
            # 确定默认月份目录
            default_month_dir = os.path.join(OUTPUT_DIR, f"{self.current_year}年{self.current_month}月")
            
            # 如果用户选择了自定义路径，复制文件到目标路径
            final_output_dir = default_month_dir
            if not use_default_path and os.path.exists(output_path):
                # 创建目标目录
                target_month_dir = os.path.join(output_path, f"{self.current_year}年{self.current_month}月")
                if not os.path.exists(target_month_dir):
                    os.makedirs(target_month_dir)
                
                # 复制文件
                copied_count = 0
                for src_file in generated_files:
                    if os.path.exists(src_file):
                        # 保持相对目录结构
                        rel_path = os.path.relpath(src_file, default_month_dir)
                        dest_file = os.path.join(target_month_dir, rel_path)
                        
                        # 确保目标目录存在
                        dest_dir = os.path.dirname(dest_file)
                        if not os.path.exists(dest_dir):
                            os.makedirs(dest_dir)
                        
                        # 复制文件
                        shutil.copy2(src_file, dest_file)
                        copied_count += 1
                
                final_output_dir = target_month_dir
                self._log(f"已复制 {copied_count} 个文件到: {target_month_dir}")
            
            self.status_var.set("就绪")
            
            # 显示结果
            report_types = []
            if generate_detail:
                report_types.append("党费明细表")
            if generate_branch:
                report_types.append("各支部党费收缴子表")
            if generate_summary:
                report_types.append("党费汇总表")
            
            report_str = "、".join(report_types)
            path_type = "默认路径" if use_default_path else "自定义路径"
            
            messagebox.showinfo(
                "生成成功",
                f"已成功生成以下报表：\n\n{report_str}\n\n输出位置: {path_type}\n文件目录: {final_output_dir}"
            )
            
        except Exception as e:
            self.status_var.set("就绪")
            self._log(f"生成报表失败: {e}")
            import traceback
            self._log(traceback.format_exc())
            messagebox.showerror("错误", f"生成失败: {e}")
    
    def _generate_all_reports(self):
        """生成所有报表（保留原方法用于菜单调用）"""
        branches = self.member_manager.get_all_branches()
        if not branches:
            messagebox.showwarning("警告", "暂无党支部数据！")
            return
        
        self.status_var.set("正在生成报表...")
        self.root.update()
        
        try:
            # 先计算党费
            FeeCalculator.calculate_all_fees(branches)
            
            # 生成党费明细表
            self.excel_generator.generate_fee_detail_sheet(branches, self.current_year, self.current_month)
            
            # 生成各支部党费收缴子表
            self.excel_generator.generate_all_branch_sheets(branches, self.current_year, self.current_month)
            
            # 生成党费汇总表
            self.excel_generator.generate_summary_sheet(branches, self.current_year, self.current_month)
            
            self._log(f"生成所有报表完成")
            self.status_var.set("就绪")
            messagebox.showinfo("成功", f"所有报表已生成完成！\n\n文件保存在: {os.path.join(OUTPUT_DIR, f'{self.current_year}年{self.current_month}月')}")
            
        except Exception as e:
            self.status_var.set("就绪")
            messagebox.showerror("错误", f"生成失败: {e}")
    
    def _show_statistics(self):
        """显示统计信息"""
        self.notebook.select(2)  # 统计信息标签页现在在索引2
        self._log("查看统计信息")
    
    def _save_data(self):
        """保存数据"""
        try:
            filepath = self.member_manager.save_to_file(self.current_year, self.current_month)
            self._log(f"保存数据: {filepath}")
            messagebox.showinfo("成功", f"数据已成功保存！\n\n文件路径: {filepath}")
        except Exception as e:
            messagebox.showerror("错误", f"保存失败: {e}")
    
    def _load_data(self):
        """加载数据"""
        dialog = tk.Toplevel(self.root)
        dialog.title("加载数据")
        dialog.geometry("500x400")
        dialog.transient(self.root)
        dialog.grab_set()
        
        # 查找数据文件
        if not os.path.exists(DATA_DIR):
            messagebox.showwarning("警告", "数据目录不存在！")
            dialog.destroy()
            return
        
        data_files = glob.glob(os.path.join(DATA_DIR, "party_fee_data_*.json"))
        if not data_files:
            messagebox.showwarning("警告", "未找到历史数据文件！")
            dialog.destroy()
            return
        
        available_data = []
        for filepath in data_files:
            filename = os.path.basename(filepath)
            parts = filename.replace(".json", "").split("_")
            if len(parts) >= 5:
                try:
                    year = int(parts[3])
                    month = int(parts[4])
                    available_data.append((year, month, filepath))
                except (ValueError, IndexError):
                    continue
        
        if not available_data:
            messagebox.showwarning("警告", "未找到有效的历史数据文件！")
            dialog.destroy()
            return
        
        available_data.sort(key=lambda x: (x[0], x[1]), reverse=True)
        
        ttk.Label(dialog, text="请选择要加载的数据:", font=("Arial", 10, "bold")).pack(pady=10)
        
        # 列表框
        list_frame = ttk.Frame(dialog)
        list_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        scrollbar = ttk.Scrollbar(list_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        listbox = tk.Listbox(list_frame, yscrollcommand=scrollbar.set, height=15, font=("Consolas", 10))
        listbox.pack(fill=tk.BOTH, expand=True)
        scrollbar.config(command=listbox.yview)
        
        for i, (year, month, filepath) in enumerate(available_data, 1):
            listbox.insert(tk.END, f"{i}. {year}年{month}月")
        
        if available_data:
            listbox.selection_set(0)
        
        def do_load():
            selection = listbox.curselection()
            if not selection:
                messagebox.showwarning("警告", "请选择要加载的数据！")
                return
            
            idx = selection[0]
            year, month, filepath = available_data[idx]
            
            self.current_year = year
            self.current_month = month
            
            if self.member_manager.load_from_file(year, month):
                # 重新计算党费
                branches = self.member_manager.get_all_branches()
                FeeCalculator.calculate_all_fees(branches)
                
                self._refresh_display()
                self._log(f"加载数据: {year}年{month}月")
                
                total_branches = len(branches)
                total_members = sum(len(b.members) for b in branches)
                
                messagebox.showinfo("成功", f"成功加载 {year}年{month}月 的数据！\n\n党支部数量: {total_branches}\n党员总人数: {total_members}")
                dialog.destroy()
            else:
                messagebox.showerror("错误", "加载失败！")
        
        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(pady=10)
        ttk.Button(btn_frame, text="加载", command=do_load).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="取消", command=dialog.destroy).pack(side=tk.LEFT, padx=5)
    
    def _new_month(self):
        """新建月份数据"""
        dialog = tk.Toplevel(self.root)
        dialog.title("新建月份数据")
        dialog.geometry("300x180")
        dialog.transient(self.root)
        dialog.grab_set()
        
        ttk.Label(dialog, text="请设置新的年月:", font=("Arial", 10, "bold")).pack(pady=10)
        
        # 年份
        year_frame = ttk.Frame(dialog)
        year_frame.pack(pady=5)
        ttk.Label(year_frame, text="年份:").pack(side=tk.LEFT)
        year_var = tk.StringVar(value=str(self.current_year))
        ttk.Entry(year_frame, textvariable=year_var, width=10).pack(side=tk.LEFT, padx=5)
        
        # 月份
        month_frame = ttk.Frame(dialog)
        month_frame.pack(pady=5)
        ttk.Label(month_frame, text="月份:").pack(side=tk.LEFT)
        month_var = tk.StringVar(value=str(self.current_month))
        ttk.Entry(month_frame, textvariable=month_var, width=10).pack(side=tk.LEFT, padx=5)
        
        def do_new():
            try:
                year = int(year_var.get())
                month = int(month_var.get())
                
                if not (1 <= month <= 12):
                    messagebox.showerror("错误", "月份必须在1-12之间！")
                    return
                
                # 确认是否清空当前数据
                if self.member_manager.get_all_branches():
                    if not messagebox.askyesno("确认", "新建月份将清空当前数据，是否继续？"):
                        return
                
                # 清空数据
                self.member_manager.branches = []
                self.current_year = year
                self.current_month = month
                
                self._refresh_display()
                self._log(f"新建月份数据: {year}年{month}月")
                messagebox.showinfo("成功", f"已创建 {year}年{month}月 的新数据！")
                dialog.destroy()
                
            except ValueError:
                messagebox.showerror("错误", "请输入有效的年份和月份！")
        
        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(pady=10)
        ttk.Button(btn_frame, text="确定", command=do_new).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="取消", command=dialog.destroy).pack(side=tk.LEFT, padx=5)
    
    def _set_year_month(self):
        """设置年月"""
        dialog = tk.Toplevel(self.root)
        dialog.title("设置年月")
        dialog.geometry("300x150")
        dialog.transient(self.root)
        dialog.grab_set()
        
        ttk.Label(dialog, text=f"当前年月: {self.current_year}年{self.current_month}月", font=("Arial", 10, "bold")).pack(pady=10)
        
        year_frame = ttk.Frame(dialog)
        year_frame.pack(pady=5)
        ttk.Label(year_frame, text="年份:").pack(side=tk.LEFT)
        year_var = tk.StringVar(value=str(self.current_year))
        ttk.Entry(year_frame, textvariable=year_var, width=10).pack(side=tk.LEFT, padx=5)
        
        month_frame = ttk.Frame(dialog)
        month_frame.pack(pady=5)
        ttk.Label(month_frame, text="月份:").pack(side=tk.LEFT)
        month_var = tk.StringVar(value=str(self.current_month))
        ttk.Entry(month_frame, textvariable=month_var, width=10).pack(side=tk.LEFT, padx=5)
        
        def do_set():
            try:
                year = int(year_var.get())
                month = int(month_var.get())
                
                if not (1 <= month <= 12):
                    messagebox.showerror("错误", "月份必须在1-12之间！")
                    return
                
                self.current_year = year
                self.current_month = month
                self._refresh_display()
                self._log(f"设置年月: {year}年{month}月")
                messagebox.showinfo("成功", f"已设置为 {year}年{month}月")
                dialog.destroy()
                
            except ValueError:
                messagebox.showerror("错误", "请输入有效的年份和月份！")
        
        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(pady=10)
        ttk.Button(btn_frame, text="确定", command=do_set).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="取消", command=dialog.destroy).pack(side=tk.LEFT, padx=5)
    
    def _import_excel_data(self):
        """导入Excel数据"""
        filepath = filedialog.askopenfilename(
            title="选择Excel文件",
            filetypes=[("Excel文件", "*.xlsx *.xls"), ("所有文件", "*.*")]
        )
        
        if not filepath:
            return
        
        self.status_var.set("正在导入数据...")
        self.root.update()
        
        try:
            from excel_importer import ExcelImporter
            importer = ExcelImporter()
            
            # 检测文件类型
            file_type = importer.detect_file_type(filepath)
            self._log(f"检测到文件类型: {file_type}")
            
            # 导入数据
            result = importer.import_from_file(filepath)
            
            if result.warnings:
                for warning in result.warnings:
                    self._log(f"提示: {warning}")
            
            if result.errors:
                for error in result.errors:
                    self._log(f"错误: {error}")
                messagebox.showerror("错误", "导入过程中发生错误，请查看日志！")
                self.status_var.set("就绪")
                return
            
            if not result.success:
                messagebox.showwarning("警告", "未导入任何数据！")
                self.status_var.set("就绪")
                return
            
            # 询问导入模式
            mode = messagebox.askyesnocancel("导入模式", f"检测到 {result.total_branches} 个支部，{result.total_members} 名党员\n\n选择导入模式:\n\n[是] 追加到当前数据\n[否] 替换当前数据\n[取消] 取消导入")
            
            if mode is None:
                self.status_var.set("就绪")
                return
            
            if mode:
                # 追加模式
                for branch in result.branches:
                    # 检查是否已存在同名支部
                    existing = self.member_manager.get_branch(branch.name)
                    if existing:
                        # 合并党员
                        for member in branch.members:
                            # 重新分配序号
                            member.sequence = len(existing.members) + 1
                            member.branch_sequence = existing.sequence
                            try:
                                self.member_manager.add_member(member, existing.name)
                            except ValueError:
                                # 如果党员已存在，跳过
                                pass
                    else:
                        # 新支部
                        branch.sequence = len(self.member_manager.get_all_branches()) + 1
                        self.member_manager.add_branch(branch)
                        for member in branch.members:
                            member.branch_sequence = branch.sequence
            else:
                # 替换模式
                self.member_manager.branches = []
                for i, branch in enumerate(result.branches, 1):
                    branch.sequence = i
                    self.member_manager.add_branch(branch)
                    for member in branch.members:
                        member.branch_sequence = i
            
            # 重新计算党费
            branches = self.member_manager.get_all_branches()
            FeeCalculator.calculate_all_fees(branches)
            
            self._refresh_display()
            self._log(f"导入Excel数据: {filepath}")
            self.status_var.set("就绪")
            
            messagebox.showinfo("导入成功", f"成功导入 {result.total_branches} 个支部，{result.total_members} 名党员！")
            
        except Exception as e:
            self.status_var.set("就绪")
            self._log(f"导入失败: {e}")
            import traceback
            self._log(traceback.format_exc())
            messagebox.showerror("错误", f"导入失败: {e}")
    
    def _show_about(self):
        """显示关于对话框"""
        messagebox.showinfo("关于", "党费收取系统 v1.0\n\n可视化版\n\n支持SQLite数据库\n支持历史数据对比分析")
    
    def _create_comparison_panel(self, parent):
        """创建数据对比面板"""
        # 选择月份区域
        select_frame = ttk.LabelFrame(parent, text="选择对比月份", padding=10)
        select_frame.pack(fill=tk.X, pady=5)
        
        # 基准月份
        ttk.Label(select_frame, text="基准月份:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=2)
        self.base_year_var = tk.StringVar(value=str(self.current_year))
        self.base_month_var = tk.StringVar(value=str(self.current_month))
        
        base_frame = ttk.Frame(select_frame)
        base_frame.grid(row=0, column=1, sticky=tk.W, padx=5, pady=2)
        ttk.Entry(base_frame, textvariable=self.base_year_var, width=8).pack(side=tk.LEFT)
        ttk.Label(base_frame, text="年").pack(side=tk.LEFT)
        ttk.Entry(base_frame, textvariable=self.base_month_var, width=6).pack(side=tk.LEFT)
        ttk.Label(base_frame, text="月").pack(side=tk.LEFT)
        
        # 对比月份
        ttk.Label(select_frame, text="对比月份:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=2)
        self.compare_year_var = tk.StringVar(value=str(self.current_year))
        self.compare_month_var = tk.StringVar(value=str(self.current_month))
        
        compare_frame = ttk.Frame(select_frame)
        compare_frame.grid(row=1, column=1, sticky=tk.W, padx=5, pady=2)
        ttk.Entry(compare_frame, textvariable=self.compare_year_var, width=8).pack(side=tk.LEFT)
        ttk.Label(compare_frame, text="年").pack(side=tk.LEFT)
        ttk.Entry(compare_frame, textvariable=self.compare_month_var, width=6).pack(side=tk.LEFT)
        ttk.Label(compare_frame, text="月").pack(side=tk.LEFT)
        
        # 操作按钮
        btn_frame = ttk.Frame(select_frame)
        btn_frame.grid(row=0, column=2, rowspan=2, padx=20)
        ttk.Button(btn_frame, text="执行对比", command=self._execute_comparison).pack(pady=2)
        ttk.Button(btn_frame, text="刷新数据库月份", command=self._refresh_db_months).pack(pady=2)
        
        # 可用月份列表
        months_frame = ttk.LabelFrame(parent, text="数据库中可用的月份", padding=10)
        months_frame.pack(fill=tk.X, pady=5)
        
        self.months_listbox = tk.Listbox(months_frame, height=4, font=("Consolas", 10))
        self.months_listbox.pack(fill=tk.X, pady=5)
        self.months_listbox.bind('<<ListboxSelect>>', self._on_month_select)
        
        # 对比结果显示
        result_frame = ttk.LabelFrame(parent, text="对比结果", padding=10)
        result_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        self.comparison_text = scrolledtext.ScrolledText(result_frame, wrap=tk.WORD, font=("Consolas", 9))
        self.comparison_text.pack(fill=tk.BOTH, expand=True)
        
        # 初始刷新月份列表
        self._refresh_db_months()
    
    def _refresh_db_months(self):
        """刷新数据库中的月份列表"""
        self.months_listbox.delete(0, tk.END)
        
        months = self.db_manager.get_available_months()
        if months:
            for year, month in months:
                self.months_listbox.insert(tk.END, f"{year}年{month}月")
            self._log(f"数据库中有 {len(months)} 个月份的数据")
        else:
            self.months_listbox.insert(tk.END, "暂无数据")
            self._log("数据库中暂无数据")
    
    def _on_month_select(self, event):
        """选择月份时的处理"""
        selection = self.months_listbox.curselection()
        if not selection:
            return
        
        selected = self.months_listbox.get(selection[0])
        if selected == "暂无数据":
            return
        
        # 解析年份和月份
        import re
        match = re.match(r"(\d+)年(\d+)月", selected)
        if match:
            year = match.group(1)
            month = match.group(2)
            # 如果是左键点击，设置为基准月份；右键点击设置为对比月份
            # 这里简化处理：如果基准月份已设置，则设置为对比月份
            if self.base_year_var.get() == "" or self.base_year_var.get() == str(self.current_year):
                self.base_year_var.set(year)
                self.base_month_var.set(month)
            else:
                self.compare_year_var.set(year)
                self.compare_month_var.set(month)
    
    def _execute_comparison(self):
        """执行数据对比"""
        try:
            base_year = int(self.base_year_var.get())
            base_month = int(self.base_month_var.get())
            compare_year = int(self.compare_year_var.get())
            compare_month = int(self.compare_month_var.get())
            
            if base_year == compare_year and base_month == compare_month:
                messagebox.showwarning("警告", "请选择不同的月份进行对比！")
                return
            
            self.status_var.set("正在执行数据对比...")
            self.root.update()
            
            # 执行对比
            result = self.data_comparator.compare_months(
                base_year, base_month, compare_year, compare_month
            )
            
            if result is None:
                messagebox.showwarning("警告", "无法获取对比数据，请确保两个月份的数据都已保存到数据库！")
                self.status_var.set("就绪")
                return
            
            # 生成报告
            report = self.data_comparator.generate_comparison_report(result)
            
            # 显示结果
            self.comparison_text.delete(1.0, tk.END)
            self.comparison_text.insert(tk.END, report)
            
            self._log(f"完成数据对比: {base_year}年{base_month}月 vs {compare_year}年{compare_month}月")
            self.status_var.set("就绪")
            
        except ValueError:
            messagebox.showerror("错误", "请输入有效的年份和月份！")
        except Exception as e:
            self.status_var.set("就绪")
            self._log(f"对比失败: {e}")
            messagebox.showerror("错误", f"对比失败: {e}")
    
    def _save_to_database(self):
        """保存数据到数据库"""
        branches = self.member_manager.get_all_branches()
        if not branches:
            messagebox.showwarning("警告", "暂无数据需要保存！")
            return
        
        try:
            if self.db_manager.save_monthly_data(self.current_year, self.current_month, branches):
                self._log(f"保存到数据库: {self.current_year}年{self.current_month}月")
                messagebox.showinfo("成功", f"数据已成功保存到数据库！\n\n月份: {self.current_year}年{self.current_month}月")
                self._refresh_db_months()
            else:
                messagebox.showerror("错误", "保存到数据库失败！")
        except Exception as e:
            self._log(f"保存到数据库失败: {e}")
            messagebox.showerror("错误", f"保存失败: {e}")
    
    def _load_from_database(self):
        """从数据库加载数据"""
        months = self.db_manager.get_available_months()
        if not months:
            messagebox.showwarning("警告", "数据库中暂无数据！")
            return
        
        # 创建选择对话框
        dialog = tk.Toplevel(self.root)
        dialog.title("从数据库加载数据")
        dialog.geometry("400x350")
        dialog.transient(self.root)
        dialog.grab_set()
        
        ttk.Label(dialog, text="请选择要加载的月份:", font=("Arial", 10, "bold")).pack(pady=10)
        
        list_frame = ttk.Frame(dialog)
        list_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        scrollbar = ttk.Scrollbar(list_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        listbox = tk.Listbox(list_frame, yscrollcommand=scrollbar.set, height=12, font=("Consolas", 10))
        listbox.pack(fill=tk.BOTH, expand=True)
        scrollbar.config(command=listbox.yview)
        
        for year, month in months:
            listbox.insert(tk.END, f"{year}年{month}月")
        
        if months:
            listbox.selection_set(0)
        
        def do_load():
            selection = listbox.curselection()
            if not selection:
                messagebox.showwarning("警告", "请选择要加载的数据！")
                return
            
            selected_idx = selection[0]
            year, month = months[selected_idx]
            
            branches = self.db_manager.load_monthly_data(year, month)
            if branches is not None:
                self.member_manager.branches = branches
                self.current_year = year
                self.current_month = month
                
                # 重新计算党费
                FeeCalculator.calculate_all_fees(branches)
                
                self._refresh_display()
                self._log(f"从数据库加载: {year}年{month}月")
                
                total_branches = len(branches)
                total_members = sum(len(b.members) for b in branches)
                
                messagebox.showinfo("成功", f"成功从数据库加载 {year}年{month}月 的数据！\n\n党支部数量: {total_branches}\n党员总人数: {total_members}")
                dialog.destroy()
            else:
                messagebox.showerror("错误", "加载失败！")
        
        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(pady=10)
        ttk.Button(btn_frame, text="加载", command=do_load).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="取消", command=dialog.destroy).pack(side=tk.LEFT, padx=5)
    
    def _list_database_months(self):
        """查看数据库中的月份"""
        months = self.db_manager.get_available_months()
        
        if not months:
            messagebox.showinfo("数据库信息", "数据库中暂无数据。")
            return
        
        info = f"数据库中共有 {len(months)} 个月份的数据：\n\n"
        for year, month in months:
            info += f"  - {year}年{month}月\n"
        
        messagebox.showinfo("数据库信息", info)
    
    def _show_data_comparison(self):
        """显示数据对比"""
        self.notebook.select(3)  # 数据对比标签页现在在索引3
        self._log("切换到数据对比")
    
    def _on_close(self):
        """关闭窗口"""
        if self.member_manager.get_all_branches():
            if messagebox.askyesno("确认退出", "是否保存数据后退出？"):
                try:
                    self.member_manager.save_to_file(self.current_year, self.current_month)
                    self._log("退出前保存数据")
                except:
                    pass
        
        self.root.destroy()


def main():
    """主函数"""
    root = tk.Tk()
    
    # 设置主题
    try:
        style = ttk.Style()
        style.theme_use('clam')
    except:
        pass
    
    app = PartyFeeGUI(root)
    root.protocol("WM_DELETE_WINDOW", app._on_close)
    root.mainloop()


if __name__ == "__main__":
    main()
