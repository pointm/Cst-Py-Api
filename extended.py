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
