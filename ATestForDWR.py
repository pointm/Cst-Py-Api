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


cst = win32com.client.dynamic.Dispatch("CSTStudio.Application")
cst.SetQuietMode(True)
new_mws = cst.NewMWS()
mws = cst.Active3D()

CstDefaultUnits(mws)

CstDefineFrequencyRange(mws, 5, 12)

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
background.Type("PEC")

W = 28.45
L = 28.45
Fi = 9
Wf = 1.137
Gpf = 1
Lg = 2 * L
Wg = 2 * W
Ht = 0.035
Hs = 1.6

a = 20
b = 10
l = 0.8 * a

Name = "WaveGuide"
component = "component2"
# material = 'Copper (annealed)'
material = "Vacuum"
Xrange = [-0.5 * a, 0.5 * a]
Yrange = [-0.5 * b, 0.5 * b]
Zrange = [-0.5 * l, 0.5 * l]
Cstbrick(mws, Name, component, material, Xrange, Yrange, Zrange)

Name = "WaveGuide"
CstPickFace(mws, Name, component, id=1)

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


Name = "WaveGuide"
CstPickFace(mws, Name, component, id=2)

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
CstDefineEfieldMonitor(mws, ("e-field" + "10"),10)
CstDefineHfieldMonitor(mws, ("h-field" + "10"), 10)


CstSaveAsProject(mws, "C:\\demo\\demo_result\\MicrostripAntenna")


CstDefineFrequencydomainSolver(mws, 5, 12, samples = '')


# export_file_path = "C:\\demo\\microstrip_demo.txt"
# CstExportTouchstone(mws, export_file_path)

# cst.Quit()
