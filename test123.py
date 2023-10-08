from extended import *


def BuildComponentBrick(handle):
    componentobj = handle.Component
    brick = handle.Brick
    componentobj.New("component1")
    brick.Reset
    brick.Name("solid1")
    brick.Component("component1")
    brick.Material("Vacuum")
    brick.Xrange("-a/2", "a/2")
    brick .Yrange("-b/2", "b/2")
    brick.Zrange("-trh/2", "trh/2")
    brick.Create


cst = win32com.client.dynamic.Dispatch("CSTStudio.Application")
# cst.SetQuietMode(True)
# new_mws = cst.NewMWS()
new_mws = cst.OpenFile(r"C:\Users\PointM2001\Documents\Demo\RWtest.cst")
mws = cst.Active3D()

f_min = 5
f_max = 12

CstDefaultUnits(mws)

CstDefineFrequencyRange(mws, f_min, f_max)

CstMeshInitiator(mws)

Xmin = "electric"
Xmax = "electric"
Ymin = "electric"
Ymax = "electric"
Zmin = "electric"
Zmax = "electric"

CstDefineOpenBoundary(mws, f_min, Xmin, Xmax, Ymin, Ymax, Zmin, Zmax)

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

a = 20
b = 10
trh = 5


mws._FlagAsMethod("StoreParameter")
mws.StoreParameter("a", a)
mws._FlagAsMethod("StoreParameter")
mws.StoreParameter("b", b)
mws._FlagAsMethod("StoreParameter")
mws.StoreParameter("trh", trh)

# BuildComponentBrick(mws)


pickname = "solid1"
pickcomponent = "component1"
CstPickFace(mws, pickname, pickcomponent, id=2)

PortNumber = 1
Xrange = [-0.5 * a, 0.5 * a]
Yrange = [-0.5 * b, 0.5 * b]
Zrange = [-trh/2, -trh/2]
XrangeAdd = [0, 0]
YrangeAdd = [0, 0]
ZrangeAdd = [0, 0]
Coordinates = "Picks"
Orientation = "positive"
WaveGuidePort(
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

pickname = "solid1"
pickcomponent = "component1"
CstPickFace(mws, pickname, pickcomponent, id=1)

PortNumber = 2
Xrange = [-0.5 * a, 0.5 * a]
Yrange = [-0.5 * b, 0.5 * b]
Zrange = [trh/2, trh/2]
XrangeAdd = [0, 0]
YrangeAdd = [0, 0]
ZrangeAdd = [0, 0]
Coordinates = "Picks"
Orientation = "positive"
WaveGuidePort(
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


# CstDefineFrequencydomainSolver(mws, f_min, f_max, "")


dsp = mws.ParameterSweep
dsp.SetSimulationType("Frequency")

dsp._FlagAsMethod("AddSequence")

dsp.AddParameter_Samples("Sweep", "a", 5, 20, 3, False)
# dsp.AddParameter_Samples("Sweep", "b", 2.5, 10, 3, False)

dsp.Start

# # CstSaveAsProject(mws, r"C:\Users\PointM2001\Documents\Demo\RWtest")


ver = mws.GetApplicationVersion
print("Version:", ver)  # 打印版本号

mu = mws.Mu0
print(mu)
