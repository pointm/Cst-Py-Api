import win32com.client
import os


def quotes_to_list(text):
    lines = text.splitlines()
    quoted_lines = ["" + line.lstrip() + "" for line in lines]
    return quoted_lines


def AddToHistory(mws, Tag, Command):
    '''
    添加到历史树
    Command只能是队列哦，不能传输裸字符串！
    '''
    line_break = '\n'  # 这句话表示将每个单句用换行符链接起来
    Command = line_break.join(Command)  # 这句话表示将每个单句用换行符链接起来
    # AddToHistory ( string caption, string contents ) bool
    # 隶属于ProjectObject的方法，必须使用mws对象进行调用
    mws._FlagAsMethod("AddToHistory")
    mws.AddToHistory(Tag, Command)


def CstSaveAsProject(mws, projectName):
    mws._FlagAsMethod("SaveAs")
    mws.SaveAs(projectName + '.cst', 'false')


def BackGroundInitial(mws):
    sCommand = ['With Background',
                '.ResetBackground',
                '.Type "PEC"',
                'End With']
    AddToHistory(mws, 'BackGroundInitial', sCommand)


def BoundaryInitial(mws):
    sCommand = ['With Boundary',
                '.Xmin "electric"',
                '.Xmax "electric"',
                '.Ymin "electric"',
                '.Ymax "electric"',
                '.Zmin "electric"',
                '.Zmax "electric"',
                '.Xsymmetry "none"',
                '.Ysymmetry "none"',
                '.Zsymmetry "none"',
                'End With']
    AddToHistory(mws, 'BoundaryInitial', sCommand)


def BrickBuild(mws, Tag, Name, component, material, Xrange, Yrange, Zrange):
    sCommand = ['With Brick',
                '.Reset',
                '.Name "%s"' % Name,
                '.Component "%s"' % component,
                '.Material "%s"' % material,
                '.Xrange "%s", "%s"' % (Xrange[0], Xrange[1]),
                '.Yrange "%s", "%s"' % (Yrange[0], Yrange[1]),
                '.Zrange "%s", "%s"' % (Zrange[0], Zrange[1]),
                '.Create',
                'End With']
    AddToHistory(mws, Tag, sCommand)


def PickPort1(mws):
    sCommand = ['Pick.PickFaceFromId "RW:WaveGuided", "1"']
    AddToHistory(mws, 'pick face id 1', sCommand)
    sCommand = ['With Port',
                '.Reset',
                '.PortNumber 1',
                '.Label  ""',
                '.NumberOfModes 1',
                '.AdjustPolarization "False"',
                '.PolarizationAngle 0.0',
                '.ReferencePlaneDistance 0',
                '.TextSize 50',
                '.TextMaxLimit 0',
                '.Coordinates "Picks"',
                '.Orientation "positive"',
                '.PortOnBound "False"',
                '.ClipPickedPortToBound "False"',
                '.Xrange "-a/2","-a/2"',
                '.Yrange "-b/2","b/2"',
                '.Zrange "l/2","l/2"',
                '.XrangeAdd "0.0","0.0"',
                '.YrangeAdd "0.0","0.0"',
                '.ZrangeAdd "0.0","0.0"',
                '.SingleEnded "False"',
                '.Create',
                'End With']
    AddToHistory(mws, 'define port 1', sCommand)


def PickPort2(mws):
    sCommand = ['Pick.PickFaceFromId "RW:WaveGuided", "2"']
    AddToHistory(mws, 'pick face id 2', sCommand)
    sCommand = ['With Port',
                '.Reset',
                '.PortNumber 2',
                '.Label  ""',
                '.NumberOfModes 1',
                '.AdjustPolarization "False"',
                '.PolarizationAngle 0.0',
                '.ReferencePlaneDistance 0',
                '.TextSize 50',
                '.TextMaxLimit 0',
                '.Coordinates "Picks"',
                '.Orientation "positive"',
                '.PortOnBound "False"',
                '.ClipPickedPortToBound "False"',
                '.Xrange "-a/2","-a/2"',
                '.Yrange "-b/2","b/2"',
                '.Zrange "-l/2","-l/2"',
                '.XrangeAdd "0.0","0.0"',
                '.YrangeAdd "0.0","0.0"',
                '.ZrangeAdd "0.0","0.0"',
                '.SingleEnded "False"',
                '.Create',
                'End With']
    AddToHistory(mws, 'define port 2', sCommand)


def FreqSimulation(mws):
    Tag = 'StartSolver'
    Command = '''
    With FDSolver
        .Start
    End With
    '''
    mws._FlagAsMethod("AddToHistory")
    mws.AddToHistory(Tag, Command)


def ParaSweep(mws):
    sCommand = '''
	With ParameterSweep
        .SetSimulationType ("Frequency")
        .AddSequence ("Sweep")
        .AddParameter_Samples ("Sweep", "a", 2, 10, 3, False)
        .Start
    End With
'''
    psp = mws.ParameterSweep
    psp.SetSimulationType("Frequency")
    psp.AddSequence("Sweep")
    psp.AddParameter_Samples("Sweep", "a", 2, 10, 3, False)
    psp.Start


Frq = [5, 7]  # 工作频率，单位：GHz


cst = win32com.client.dynamic.Dispatch("CSTStudio.Application")
# 调用COM对象的方法

path = os.getcwd()  # 获取当前py文件所在文件夹
filename = 'TestCOMWithHistorys'
fullname = os.path.join(path, filename)
print(fullname)
projectName = fullname

# 下面的被调用的几个方法都隶属于VBA Application Object，
# VBA Application Object隶属的方法会直接使用刚刚的cst的COM对象进行调用
# cst.SetQuietMode(True)
new_mws = cst.NewMWS()  # 新建并且返回一个微波工作室的对象
# 这个返回的对象跟直接打开某一个文件返回的对象是一样的
# new_mws = cst.OpenFile(fullname + '.cst')
mws = cst.Active3D()  # 控制现在的被打开的微波工作室

CstSaveAsProject(mws, projectName)

# 工作频率设置
sCommand = ['Solver.FrequencyRange "%f","%f"' % (Frq[0], Frq[1])]
AddToHistory(mws, 'define frequency range', sCommand)

# 设置边界条件
BackGroundInitial(mws)

# 背景材料设置
BoundaryInitial(mws)

a = 20
b = 15
l = 5
AddToHistory(
    mws, 'StoreParameter', ['MakeSureParameterExists("a", "%f")' % a,
                            'MakeSureParameterExists("b", "%f")' % b,
                            'MakeSureParameterExists("l", "%f")' % l])

# 设置求解器类型
AddToHistory(mws, "change solver type",
             ['ChangeSolverType "HF Frequency Domain"'])

# 开始建模
BrickBuild(mws, 'Brick1', 'WaveGuided', 'RW', 'Vacuum',
           Xrange=['-a/2', 'a/2'], Yrange=['-b/2', 'b/2'], Zrange=['-l/2', 'l/2'])

# 选择端口
PickPort1(mws)
PickPort2(mws)

# 开始仿真
FreqSimulation(mws)

# 开始扫参
# ParaSweep(mws)

#
