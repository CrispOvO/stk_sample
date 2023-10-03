# 使用comtypes包来连接外部应用
import comtypes
from comtypes.client import CreateObject

# 创建STK新实例
app = CreateObject('STK11.Application')
app.Visible = True # Needed to view the user interface application

# 获取 IAgStkObjectRoot 接口
root = app.Personality2

# Note: When 'root=uiApplication.Personality2' is executed, the comtypes library automatically creates a gen folder that contains
# STKObjects and other Python wrappers for the STK libraries. After running this at least once on your computer, the libraries

# 查看新生成的文件
import os as os
print(comtypes.client.gen_dir, '\n')
print(os.listdir(comtypes.client.gen_dir))

# After running this cell comment out this cell.
# Use "Ctrl" + "a" to select all of the cell content, then Use "Ctrl" + "/" to toggle comments in Jupyter Notebooks