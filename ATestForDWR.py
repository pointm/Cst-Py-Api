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
from PostProcessing.CstExportTouchstone import *


def CstSapphire(mws):
    material = mws.Material

    material.Reset()
    material.Name("Sapphire")
    material.FrqType("all")
    material.Type("Normal")
    material.MaterialUnit("Frequency", "GHz")
    material.MaterialUnit("Geometry", "mm")
    material.Epsilon("9.4")
    material.Mu("1.0")
    material.Kappa("0.0")
    material.TanD("0.0")
    material.TanDFreq("0.0")
    material.TanDGiven("False")
    material.TanDModel("ConstKappa")
    material.KappaM("0")
    material.TanDM("0.0")
    material.TanDMFreq("0.0")
    material.TanDMGiven("False")
    material.TanDMModel("ConstKappa")
    material.DispModelEps("None")
    material.DispModelMu("None")
    material.DispersiveFittingSchemeEps("General 1st")
    material.DispersiveFittingSchemeMu("General 1st")
    material.UseGeneralDispersionEps("False")
    material.UseGeneralDispersionMu("False")
    material.Rho("0")
    material.ThermalType("Normal")
    material.ThermalConductivity("0")
    material.HeatCapacity("0")
    material.SetActiveMaterial("all")
    material.Colour("0", "0.717647", "1")
    material.Wireframe("False")
    material.Transparency("0")
    material.Create()
    format(material)


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


def MirrorTransform(mws, transcomponent, transName):
    trans = mws.Transform
    trans.Reset
    trans.Name(transcomponent + ":" + transName)
    trans.Origin("Free")
    trans.Center("0", "0", "0")
    trans.PlaneNormal("0", "0", "-1")
    trans.MultipleObjects("True")
    trans.GroupObjects("False")
    trans.Repetitions("1")
    trans.MultipleSelection("False")
    trans.Destination("")
    trans.Material("")
    trans.AutoDestination("True")
    trans.Transform("Shape", "Mirror")


def UpdateMesh(mws):
    mesh = mws.Mesh
    meshsettings = mws.MeshSettings

    mesh.MeshType("Tetrahedral")
    mesh.SetCreator("High Frequency")

    meshsettings.SetMeshType("Tet")
    meshsettings.Set("Version", 1)
    # 'MAX CELL - WAVELENGTH REFINEMENT
    meshsettings.Set("StepsPerWaveNear", "17")
    meshsettings.Set("StepsPerWaveFar", "11")
    meshsettings.Set("PhaseErrorNear", "0.02")
    meshsettings.Set("PhaseErrorFar", "0.02")
    meshsettings.Set("CellsPerWavelengthPolicy", "cellsperwavelength")
    # 'MAX CELL - GEOMETRY REFINEMENT
    meshsettings.Set("StepsPerBoxNear", "17")
    meshsettings.Set("StepsPerBoxFar", "11")
    meshsettings.Set("ModelBoxDescrNear", "maxedge")
    meshsettings.Set("ModelBoxDescrFar", "maxedge")
    # 'MIN CELL
    meshsettings.Set("UseRatioLimit", "0")
    meshsettings.Set("RatioLimit", "100")
    meshsettings.Set("MinStep", "0")
    # 'MESHING METHOD
    meshsettings.SetMeshType("Unstr")
    meshsettings.Set("Method", "0")

    meshsettings.SetMeshType("Tet")
    meshsettings.Set("CurvatureOrder", "1")
    meshsettings.Set("CurvatureOrderPolicy", "automatic")
    meshsettings.Set("CurvRefinementControl", "NormalTolerance")
    meshsettings.Set("NormalTolerance", "22.5")
    meshsettings.Set("SrfMeshGradation", "1.5")
    meshsettings.Set("SrfMeshOptimization", "1")

    meshsettings.SetMeshType("Unstr")
    meshsettings.Set("UseMaterials", "1")
    meshsettings.Set("MoveMesh", "0")

    meshsettings.SetMeshType("All")
    meshsettings.Set("AutomaticEdgeRefinement", "0")

    meshsettings.SetMeshType("Tet")
    meshsettings.Set("UseAnisoCurveRefinement", "1")
    meshsettings.Set("UseSameSrfAndVolMeshGradation", "1")
    meshsettings.Set("VolMeshGradation", "1.5")
    meshsettings.Set("VolMeshOptimization", "1")

    meshsettings.SetMeshType("Unstr")
    meshsettings.Set("SmallFeatureSize", "0")
    meshsettings.Set("CoincidenceTolerance", "1e-06")
    meshsettings.Set("SelfIntersectionCheck", "1")
    meshsettings.Set("OptimizeForPlanarStructures", "0")

    mesh.SetParallelMesherMode("Tet", "maximum")
    mesh.SetMaxParallelMesherThreads("Tet", "1")

    mesh.Update


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

CstSapphire(mws)  # 创建蓝宝石材料

a = 20
b = 10
d = 5
s = 4
wr = 10.987247233334
L = 0.3 * a
wt = 0.87928459706468
trh = 1.8512664671501
ta = 15.83559084768
tb = 10.324854504076


windowcomponent = "Window"
windowname = "SapphireWindow"
windowmaterial = "Sapphire"
windowaxis = "Z"
Xcenter = 0
Ycenter = 0
Zrange = [-wt / 2, wt / 2]
Cstcylinder(
    mws,
    windowname,
    windowcomponent,
    windowmaterial,
    windowaxis,
    wr,
    0,
    Xcenter,
    Ycenter,
    Zrange,
)

# pickname = "SapphireWindow"
# pickcomponent = "Window"

# pick = mws.Pick

# pick.PickCenterpointFromId(pickname + ":" + pickcomponent, "3")

# x = pick.GetPickpointCoordinatesCompByIndex(0,0)
# y = pick.GetPickpointCoordinatesCompByIndex(0,1)
# z = pick.GetPickpointCoordinatesCompByIndex(0,2)#得到Pick的点的坐标
wcs = mws.WCS
x = 0
y = 0
z = wt / 2  # 得到局部WCS坐标系要移动到的点的坐标
wcs.SetOrigin(x, y, z)  # 激活局部坐标
wcs.ActivateWCS("local")

Name = "RW"
component = "TransWaveGuide"
material = "Vacuum"
Xrange = [-0.5 * ta, 0.5 * ta]
Yrange = [-0.5 * tb, 0.5 * tb]
Zrange = [0, trh]
Cstbrick(mws, Name, component, material, Xrange, Yrange, Zrange)

x = 0
y = 0
z = wt / 2 + trh  # 得到局部WCS坐标系要移动到的点的坐标
wcs.SetOrigin(x, y, z)  # 激活局部坐标
wcs.ActivateWCS("local")

Name = "DRWWaveGuide"
component = "WaveGuide"
material = "Vacuum"
Xrange = [-0.5 * a, 0.5 * a]
Yrange = [-0.5 * b, 0.5 * b]
Zrange = [0, L]
Cstbrick(mws, Name, component, material, Xrange, Yrange, Zrange)

Name = "CutOffSpace"
component = "WaveGuide"
material = "Vacuum"
Xrange = [-0.5 * s, 0.5 * s]
Yrange = [0.5 * d, 0.5 * b]
Zrange = [0, L]
Cstbrick(mws, Name, component, material, Xrange, Yrange, Zrange)

Name = "CutOffSpace2"
component = "WaveGuide"
material = "Vacuum"
Xrange = [-0.5 * s, 0.5 * s]
Yrange = [-0.5 * b, -0.5 * d]
Zrange = [0, L]
Cstbrick(mws, Name, component, material, Xrange, Yrange, Zrange)

component1 = "WaveGuide:DRWWaveGuide"
component2 = "WaveGuide:CutOffSpace"  # component:Name的格式，不要写反了
CstSubtract(mws, component1, component2)

component1 = "WaveGuide:DRWWaveGuide"
component2 = "WaveGuide:CutOffSpace2"
CstSubtract(mws, component1, component2)

wcs.ActivateWCS("global")

transcomponent = "TransWaveGuide"
transname = "RW"
MirrorTransform(mws, transcomponent, transname)

transcomponent = "WaveGuide"
transname = "DRWWaveGuide"
MirrorTransform(mws, transcomponent, transname)

pickname = "DRWWaveGuide"
pickcomponent = "WaveGuide"
CstPickFace(mws, pickname, pickcomponent, id=21)

PortNumber = 1
Xrange = [-0.5 * a, 0.5 * a]
Yrange = [-0.5 * b, 0.5 * b]
Zrange = [L + wt / 2 + trh, L + wt / 2 + trh]
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

pickname = "DRWWaveGuide_1"
pickcomponent = "WaveGuide"
CstPickFace(mws, pickname, pickcomponent, id=21)

PortNumber = 2
Xrange = [-0.5 * a, 0.5 * a]
Yrange = [-0.5 * b, 0.5 * b]
Zrange = [-(L + wt / 2 + trh), -(L + wt / 2 + trh)]
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


mws.SaveAs("C:\\Users\\PointM2001\\Documents\\Demo\\test.cst", True)

UpdateMesh(mws)
CstDefineFrequencydomainSolver(mws, f_min, f_max, "")

(
    frequencies_list,
    [y_real, y_imag],
    y_list,
    [x_label, y_label, plot_title],
) = CstResultParameters(
    mws, parent_path=r"1D Results\S-Parameters", run_id=0, result_id=0
)

plt.figure(dpi=200)
plt.plot(frequencies_list, y_real)
plt.plot(frequencies_list, y_imag)
plt.xlabel(x_label)
plt.ylabel(y_label)
plt.title(plot_title)
plt.show()

plt.figure(dpi=200)
plt.plot(frequencies_list, y_list)
plt.xlabel(x_label)
plt.ylabel(y_label)
plt.title(plot_title)
plt.show()


export_file_path = "C:\\Users\\PointM2001\\Documents\\demo\\DWRWindow.txt"
CstExportTouchstone(mws, export_file_path)


# cst.Quit()
