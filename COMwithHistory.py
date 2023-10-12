import win32com.client
import os


def AddToHistory(mws, Tag, Command):
    '''
    添加到历史树
    '''
    mws._FlagAsMethod("AddToHistory")
    mws.AddToHistory(Tag, Command)


def CstSaveAsProject(mws, projectName):
    mws._FlagAsMethod("SaveAs")
    mws.SaveAs(projectName + '.cst', 'false')


f_min = 5
f_max = 12


cst = win32com.client.dynamic.Dispatch("CSTStudio.Application")
# 调用COM对象的方法

path = os.getcwd()  # 获取当前py文件所在文件夹
filename = 'TestCOMWithHistorys'
fullname = os.path.join(path, filename)
print(fullname)
projectName = fullname

# 下面的被调用的几个方法都隶属于VBA Application Object，
# VBA Application Object隶属的方法会直接使用刚刚的cst的COM对象进行调用
cst.SetQuietMode(True)
# new_mws = cst.NewMWS()  # 新建并且返回一个微波工作室的对象
# 这个返回的对象跟直接打开某一个文件返回的对象是一样的
new_mws = cst.OpenFile(fullname + '.cst')
mws = cst.Active3D()  # 控制现在的被打开的微波工作室

CstSaveAsProject(mws, projectName)

# AddToHistory ( string caption, string contents ) bool
# 隶属于ProjectObject的方法，必须使用mws对象进行调用
version = mws.GetApplicationVersion
print(version)

a = 20
AddToHistory(
    mws, 'StoreParameter', 'MakeSureParameterExists("a", "%f")' % a)


# 设置边界条件

# 设置求解器类型
