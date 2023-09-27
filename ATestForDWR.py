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
from PostProcessing.CstExportTouchstone import *


def WaveGuidePort(
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
):
    port = mws.Port
    port.Reset()
    port.PortNumber(str(PortNumber))
    port.Label("")
    port.Folder("")
    port.NumberOfModes("1")
    port.AdjustPolarization("False")
    port.PolarizationAngle("0.0")
    port.ReferencePlaneDistance("0")
    port.TextSize("50")
    port.TextMaxLimit("0")
    port.Coordinates(Coordinates)
    port.Orientation(Orientation)
    port.PortOnBound("False")
    port.ClipPickedPortToBound("False")
    port.Xrange(str(Xrange[0]), str(Xrange[1]))
    port.Yrange(str(Yrange[0]), str(Yrange[1]))
    port.Zrange(str(Zrange[0]), str(Zrange[1]))
    port.XrangeAdd(str(XrangeAdd[0]), str(XrangeAdd[1]))
    port.YrangeAdd(str(YrangeAdd[0]), str(YrangeAdd[1]))
    port.ZrangeAdd(str(ZrangeAdd[0]), str(ZrangeAdd[1]))
    port.SingleEnded("False")
    port.WaveguideMonitor("False")
    port.Create()


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

XminSpace = 0
XmaxSpace = 0
YminSpace = 0
YmaxSpace = 0
ZminSpace = 0
ZmaxSpace = 0
CstDefineBackroundMaterial(
    mws, XminSpace, XmaxSpace, YminSpace, YmaxSpace, ZminSpace, ZmaxSpace
)
background = mws.Background
background.Type("PEC")  # 设置背景为PEC

project = mws.Project


a = 20
b = 10
d = 5
s = 4
l = 0.6 * a


Name = "DRWWaveGuide"
component = "WaveGuide"
material = "Vacuum"
Xrange = [-0.5 * a, 0.5 * a]
Yrange = [-0.5 * b, 0.5 * b]
Zrange = [-0.5 * l, 0.5 * l]
Cstbrick(mws, Name, component, material, Xrange, Yrange, Zrange)

Name = "CutOffSpace"
component = "WaveGuide"
material = "Vacuum"
Xrange = [-0.5 * s, 0.5 * s]
Yrange = [0.5 * d, 0.5 * b]
Zrange = [-0.5 * l, 0.5 * l]
Cstbrick(mws, Name, component, material, Xrange, Yrange, Zrange)

Name = "CutOffSpace2"
component = "WaveGuide"
material = "Vacuum"
Xrange = [-0.5 * s, 0.5 * s]
Yrange = [-0.5 * b, -0.5 * d]
Zrange = [-0.5 * l, 0.5 * l]
Cstbrick(mws, Name, component, material, Xrange, Yrange, Zrange)

component1 = "WaveGuide:DRWWaveGuide"
component2 = "WaveGuide:CutOffSpace"  # component:Name的格式，不要写反了
CstSubtract(mws, component1, component2)

component1 = "WaveGuide:DRWWaveGuide"
component2 = "WaveGuide:CutOffSpace2"
CstSubtract(mws, component1, component2)

Name = "DRWWaveGuide"
CstPickFace(mws, Name, component, id=21)

PortNumber = 1
Xrange = [-0.5 * a, 0.5 * a]
Yrange = [-0.5 * b, 0.5 * b]
Zrange = [-0.5 * l, -0.5 * l]
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

Name = "DRWWaveGuide"
CstPickFace(mws, Name, component, id=14)

PortNumber = 2
Xrange = [-0.5 * a, 0.5 * a]
Yrange = [-0.5 * b, 0.5 * b]
Zrange = [0.5 * l, 0.5 * l]
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
# CstDefineEfieldMonitor(mws, ("e-field" + "10"), 10)
# CstDefineHfieldMonitor(mws, ("h-field" + "10"), 10)


CstSaveAsProject(mws, "RW")


# CstDefineFrequencydomainSolver(mws, f_min, f_max, "")

# (
#     frequencies_list,
#     [y_real, y_imag],
#     y_list,
#     [x_label, y_label, plot_title],
# ) = CstResultParameters(
#     mws, parent_path=r"1D Results\S-Parameters", run_id=0, result_id=0
# )

# plt.figure(dpi=200)
# plt.plot(frequencies_list, y_real)
# plt.plot(frequencies_list, y_imag)
# plt.xlabel(x_label)
# plt.ylabel(y_label)
# plt.title(plot_title)
# plt.show()

# plt.figure(dpi=200)
# plt.plot(frequencies_list, y_list)
# plt.xlabel(x_label)
# plt.ylabel(y_label)
# plt.title(plot_title)
# plt.show()


# export_file_path = "C:\\Users\\PointM2001\\Documents\\demo\\DRW.txt"
# CstExportTouchstone(mws, export_file_path)


# cst.Quit()
