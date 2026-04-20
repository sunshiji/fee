# Word文件导入模块
import os
import re
from typing import List, Dict, Optional, Tuple, Any
from dataclasses import dataclass

try:
    from docx import Document
    from docx.table import Table
    from docx.text.paragraph import Paragraph
    HAS_DOCX = True
except ImportError:
    HAS_DOCX = False

from data_models import PartyMember, PartyBranch


@dataclass
class WordImportResult:
    """Word导入结果"""
    success: bool = False
    summary_data: Dict[str, Any] = None
    branches: List[PartyBranch] = None
    errors: List[str] = None
    warnings: List[str] = None
    
    def __post_init__(self):
        if self.summary_data is None:
            self.summary_data = {}
        if self.branches is None:
            self.branches = []
        if self.errors is None:
            self.errors = []
        if self.warnings is None:
            self.warnings = []


class WordImporter:
    """Word文件导入器"""
    
    def __init__(self):
        """初始化导入器"""
        if not HAS_DOCX:
            print("警告: 未安装python-docx库，无法解析Word文件")
            print("请运行: pip install python-docx")
        
        # 定义汇总表列名映射
        self.summary_column_mappings = {
            'sequence': ['序号', '编号'],
            'branch_name': ['支部名称', '党支部', '支部'],
            'member_count': ['党员人数', '人数', '数量'],
            'amount': ['金额', '总金额', '党费', '应缴金额'],
            'remark': ['备注', '说明']
        }
        
        # 反转映射
        self.reverse_summary_mapping = {}
        for key, values in self.summary_column_mappings.items():
            for value in values:
                self.reverse_summary_mapping[value] = key
    
    def _normalize_text(self, text: str) -> str:
        """规范化文本"""
        if not text:
            return ""
        return str(text).strip().replace('\n', ' ').replace('\r', '')
    
    def _parse_table(self, table: Table) -> List[List[str]]:
        """
        解析表格为二维列表
        
        Args:
            table: Word表格对象
            
        Returns:
            二维列表，每个元素是单元格的文本
        """
        data = []
        for row in table.rows:
            row_data = []
            for cell in row.cells:
                text = self._normalize_text(cell.text)
                row_data.append(text)
            data.append(row_data)
        return data
    
    def _detect_header_row(self, table_data: List[List[str]]) -> Tuple[int, Dict[str, int]]:
        """
        检测表头行和列映射
        
        Args:
            table_data: 表格数据
            
        Returns:
            (表头行号, {字段名: 列号})
        """
        # 尝试前5行作为可能的表头
        max_header_row = min(5, len(table_data))
        
        for row_idx in range(max_header_row):
            row = table_data[row_idx]
            columns_found = {}
            
            for col_idx, cell_value in enumerate(row):
                normalized_value = self._normalize_text(cell_value)
                
                # 检查是否匹配任何已知列名
                if normalized_value in self.reverse_summary_mapping:
                    field_name = self.reverse_summary_mapping[normalized_value]
                    columns_found[field_name] = col_idx
            
            # 如果找到至少2个关键字段
            essential_fields = ['branch_name', 'member_count', 'amount']
            found_essential = sum(1 for f in essential_fields if f in columns_found)
            
            if found_essential >= 2 or len(columns_found) >= 3:
                return row_idx, columns_found
        
        return 0, {}
    
    def _parse_summary_table(self, table: Table) -> List[Dict[str, Any]]:
        """
        解析汇总表
        
        Args:
            table: Word表格对象
            
        Returns:
            汇总数据列表
        """
        table_data = self._parse_table(table)
        
        if not table_data:
            return []
        
        # 检测表头
        header_row, column_mapping = self._detect_header_row(table_data)
        
        if not column_mapping:
            # 尝试使用默认映射（假设列顺序）
            if len(table_data[0]) >= 3:
                # 假设第一列是序号，第二列是支部名称，第三列是党员人数，第四列是金额
                column_mapping = {
                    'sequence': 0,
                    'branch_name': 1,
                    'member_count': 2,
                    'amount': 3
                }
        
        summary_data = []
        
        # 遍历数据行
        for row_idx in range(header_row + 1, len(table_data)):
            row = table_data[row_idx]
            
            # 跳过空行或合计行
            if not row or all(not self._normalize_text(cell) for cell in row):
                continue
            
            row_text = ''.join(row).lower()
            if '合计' in row_text or '总计' in row_text:
                continue
            
            item = {}
            
            # 解析各字段
            for field_name, col_idx in column_mapping.items():
                if col_idx < len(row):
                    value = self._normalize_text(row[col_idx])
                    
                    # 尝试转换数值
                    if field_name in ['member_count', 'amount']:
                        try:
                            # 去除可能的逗号和空格
                            clean_value = value.replace(',', '').replace('，', '').strip()
                            if field_name == 'member_count':
                                item[field_name] = int(float(clean_value))
                            else:
                                item[field_name] = float(clean_value)
                        except (ValueError, TypeError):
                            item[field_name] = value
                    else:
                        item[field_name] = value
            
            if item.get('branch_name'):
                summary_data.append(item)
        
        return summary_data
    
    def import_from_file(self, filepath: str) -> WordImportResult:
        """
        从Word文件导入数据
        
        Args:
            filepath: Word文件路径
            
        Returns:
            导入结果
        """
        result = WordImportResult()
        
        # 检查是否安装了python-docx
        if not HAS_DOCX:
            result.errors.append("未安装python-docx库，无法解析Word文件")
            result.errors.append("请运行: pip install python-docx")
            return result
        
        # 检查文件是否存在
        if not os.path.exists(filepath):
            result.errors.append(f"文件不存在: {filepath}")
            return result
        
        # 检查文件扩展名
        if not filepath.endswith(('.docx', '.doc')):
            result.errors.append(f"不支持的文件格式: {filepath}")
            result.errors.append("仅支持 .docx 格式的Word文件")
            return result
        
        try:
            # 加载文档
            doc = Document(filepath)
            result.warnings.append(f"成功加载文档: {os.path.basename(filepath)}")
            
            # 提取文档标题
            title_text = ""
            for para in doc.paragraphs[:5]:  # 检查前5段
                text = self._normalize_text(para.text)
                if text and len(text) > 5:
                    title_text = text
                    break
            
            if title_text:
                result.warnings.append(f"文档标题: {title_text}")
                result.summary_data['title'] = title_text
            
            # 解析所有表格
            all_summary_data = []
            
            for table_idx, table in enumerate(doc.tables):
                result.warnings.append(f"解析表格 {table_idx + 1}")
                
                # 尝试解析为汇总表
                summary_data = self._parse_summary_table(table)
                
                if summary_data:
                    all_summary_data.extend(summary_data)
                    result.warnings.append(f"  找到 {len(summary_data)} 条汇总数据")
            
            result.summary_data['branches'] = all_summary_data
            
            # 如果有汇总数据，尝试创建支部对象（仅包含基本信息）
            if all_summary_data:
                for item in all_summary_data:
                    branch_name = item.get('branch_name', '')
                    if branch_name:
                        branch = PartyBranch(
                            name=branch_name,
                            sequence=item.get('sequence', len(result.branches) + 1)
                        )
                        # 存储汇总信息到支部对象（虽然没有党员数据）
                        result.branches.append(branch)
                
                result.total_branches = len(result.branches)
                result.success = True
            
            # 如果没有表格，尝试从文本中提取信息
            if not all_summary_data:
                result.warnings.append("未找到可解析的表格")
                result.warnings.append("注意: Word汇总表通常仅包含统计信息，不包含详细的党员数据")
                result.warnings.append("如果需要导入党员数据，请使用Excel格式的明细表或子表")
                
                # 仍然标记为部分成功
                result.success = True
            
        except Exception as e:
            result.errors.append(f"导入失败: {str(e)}")
            import traceback
            result.errors.append(traceback.format_exc())
        
        return result
    
    def extract_summary_info(self, filepath: str) -> Dict[str, Any]:
        """
        提取汇总表信息（用于统计参考）
        
        Args:
            filepath: Word文件路径
            
        Returns:
            汇总信息字典
        """
        result = self.import_from_file(filepath)
        
        if result.success:
            return result.summary_data
        else:
            return {'error': result.errors}


def check_word_dependency() -> bool:
    """
    检查Word依赖是否安装
    
    Returns:
        是否安装了python-docx
    """
    return HAS_DOCX


def install_word_dependency_instructions() -> str:
    """
    获取安装Word依赖的说明
    
    Returns:
        安装说明文本
    """
    return """
需要安装 python-docx 库才能解析Word文件。

请运行以下命令安装:
    pip install python-docx

或者使用conda安装:
    conda install -c conda-forge python-docx

注意: 
- Word汇总表通常只包含统计信息（支部名称、党员人数、金额等）
- Word文件不包含详细的党员工资和扣款信息
- 如果需要导入完整的党员数据，建议使用Excel格式的明细表
"""
