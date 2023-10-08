from PostProcessing.CstExportTouchstone import *
import win32com.client
from Modeling.Cstbrick import *
from Home.CstSaveAsProject import *
from Simulation.CstDefineFrequencydomainSolver import *
import win32com.client
from matplotlib import pyplot as plt
from Home.CstMeshInitiator import *
from Home.CstDefineBackroundMaterial import *
from Home.CstDefaultUnits import *
from Home.CstSaveAsProject import *
from Home.CstQuitProject import *
from Materials.CstCopperAnnealedLossy import *
from Materials.CstFR4lossy import *
from Modeling.Cstbrick import *
from Modeling.Cstcylinder import *
from Modeling.CstSubtract import *
from Modeling.CstAdd import *
from Modeling.CstPickFace import *
from Simulation.CstDefineFrequencyRange import *
from Simulation.CstDefineOpenBoundary import *
from Simulation.CstWaveguidePort import *
from Simulation.CstDefineFrequencydomainSolver import *
from Simulation.CstDefineHfieldMonitor import *
from Simulation.CstDefineEfieldMonitor import *
from Simulation.CstDefineFarfieldMonitor import *
from Simulation.CstDefineTimedomainSolver import *
from PostProcessing.CstResultParameters import *


cst = win32com.client.dynamic.Dispatch("CSTStudio.Application")
# cst.SetQuietMode(True)
new_mws = cst.NewMWS()
# new_mws = cst.OpenFile(r"C:\Users\PointM2001\Documents\Demo\RWtest.cst")
mws = cst.Active3D()

f_min = 5
f_max = 12

cst = win32com.client.dynamic.Dispatch("CSTStudio.Application")
cst.SetQuietMode(True)
new_mws = cst.NewMWS()
# new_mws = cst.OpenFile(r"C:\Users\PointM2001\Documents\Demo\RWtest.cst")
mws = cst.Active3D()


CstDefaultUnits(mws)

CstDefineFrequencyRange(mws, f_min, f_max)

CstMeshInitiator(mws)

Xmin = "electric"
Xmax = "electric"
Ymin = "electric"
Ymax = "electric"
Zmin = "electric"
Zmax = "electric"
minfrequency = 1.5
CstDefineOpenBoundary(mws, minfrequency, Xmin, Xmax, Ymin, Ymax, Zmin, Zmax)

# XminSpace = 0
# XmaxSpace = 0
# YminSpace = 0
# YmaxSpace = 0
# ZminSpace = 0
# ZmaxSpace = 0
# CstDefineBackroundMaterial(
#     mws, XminSpace, XmaxSpace, YminSpace, YmaxSpace, ZminSpace, ZmaxSpace
# )
background = mws.Background
background.Type("PEC")  # 直接设置背景为PEC，现阶段仿真暂时不需要那么精确的背景材料


# project = mws.Project
# project.SaveAs("C:\\Users\\PointM2001\\Documents\\Demo\\test1.cst", False)

a = 20
b = 10
trh = 5


mws._FlagAsMethod("StoreParameter")
mws.StoreParameter("a", a)
mws._FlagAsMethod("StoreParameter")
mws.StoreParameter("b", b)
mws._FlagAsMethod("StoreParameter")
mws.StoreParameter("trh", trh)

Name = "RW"
component = "TransWaveGuide"
material = "Vacuum"
Xrange = [-0.5 * a, 0.5 * a]
Yrange = [-0.5 * b, 0.5 * b]
Zrange = [0, trh]
Cstbrick(mws, Name, component, material, Xrange, Yrange, Zrange)

Name = "RW"
component = "TransWaveGuide"
CstPickFace(mws, Name, component, id=1)

PortNumber = 1
Zrange = [trh, trh]
XrangeAdd = [0, 0]
YrangeAdd = [0, 0]
ZrangeAdd = [0, 0]
Coordinates = "Picks"
Orientation = "positive"
CstWaveguidePort(
    mws,
    PortNumber,
    Xrange,
    Yrange,
    Zrange,
    XrangeAdd,
    YrangeAdd,
    ZrangeAdd,
    Coordinates,
    Orientation,
)

Name = "RW"
component = "TransWaveGuide"
CstPickFace(mws, Name, component, id=2)

PortNumber = 2
Zrange = [0, 0]
XrangeAdd = [0, 0]
YrangeAdd = [0, 0]
ZrangeAdd = [0, 0]
Coordinates = "Picks"
Orientation = "positive"
CstWaveguidePort(
    mws,
    PortNumber,
    Xrange,
    Yrange,
    Zrange,
    XrangeAdd,
    YrangeAdd,
    ZrangeAdd,
    Coordinates,
    Orientation,
)

CstDefineFrequencydomainSolver(mws, f_min, f_max, "")


# dsp = mws.ParameterSweep
# dsp.SetSimulationType("parameter sweep")

# dsp.AddSequence("Sweep")

# dsp.AddParameter_Samples("Sweep", "a", 5, 20, 5, False)

# dsp.Start

# CstSaveAsProject(mws, r"C:\Users\PointM2001\Documents\Demo\RWtest")

ver = mws.GetApplicationVersion
print("Version:", ver)  # 打印版本号

mu = mws.Mu0
print(mu)
