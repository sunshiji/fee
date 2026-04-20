# 系统配置文件
import os

# 项目根目录
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# 数据目录
DATA_DIR = os.path.join(BASE_DIR, "data")
# 输出目录
OUTPUT_DIR = os.path.join(BASE_DIR, "output")
# 模板目录
TEMPLATE_DIR = os.path.join(BASE_DIR, "templates")

# 默认年份和月份
DEFAULT_YEAR = 2026
DEFAULT_MONTH = 2

# 党费计算比例
PARTY_FEE_RATES = {
    "below_3000": 0.005,  # 3000元以下(含3000元)，0.5%
    "3000_5000": 0.01,    # 3000元以上至5000元(含5000元)，1%
    "5000_10000": 0.015,  # 5000元以上至10000元(含10000元)，1.5%
    "above_10000": 0.02    # 10000元以上，2%
}

# 学院名称
COLLEGE_NAME = "新疆大学计算机科学与技术学院"
