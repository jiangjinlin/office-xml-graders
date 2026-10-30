"""网页(Dreamweaver)操作题评分器 - 213 题组第 2 题（高三计算机技能题库）"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from dw_engine import main
from dw_specs import SPECS

if __name__ == '__main__':
    sys.exit(main(SPECS['213-2']))
