# 小狮子 Office 2010 考试系统 - WSGI 入口
# 宝塔部署：Python 项目管理器 → 启动文件选此文件
# 说明：uwsgi.ini 直接加载 app.py，清理线程在 app.py 导入时自行启动
import os

from app import app

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get('EXAM_PORT', '5000')))
