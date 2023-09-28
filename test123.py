import win32com.client
from Modeling.Cstbrick import *
from Home.CstSaveAsProject import *

cst = win32com.client.dynamic.Dispatch("CSTStudio.Application")
cst.SetQuietMode(True)
new_mws = cst.NewMWS()
new_mws = cst.OpenFile(r"C:\Users\PointM2001\Documents\Demo\RWtest.cst")
mws = cst.Active3D()

mws._FlagAsMethod("StoreParameter")
mws.StoreParameter("a","20")

CstSaveAsProject(mws, r"C:\Users\PointM2001\Documents\Demo\test1")

ver = mws.GetApplicationVersion
print("Version:", ver)  # 打印版本号

mu = mws.Mu0
print(mu)
