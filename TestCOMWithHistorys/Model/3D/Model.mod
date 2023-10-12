'# MWS Version: Version 2022.4 - Apr 26 2022 - ACIS 31.0.1 -

'# length = mm
'# frequency = GHz
'# time = ns
'# frequency range: fmin = 5.000000 fmax = 7.000000
'# created = '[VERSION]2022.4|31.0.1|20220426[/VERSION]


'@ define frequency range

'[VERSION]2022.4|31.0.1|20220426[/VERSION]
Solver.FrequencyRange "5.000000","7.000000"

'@ BackGroundInitial

'[VERSION]2022.4|31.0.1|20220426[/VERSION]
With Background
.ResetBackground
.Type "PEC"
End With

'@ BoundaryInitial

'[VERSION]2022.4|31.0.1|20220426[/VERSION]
With Boundary
.Xmin "electric"
.Xmax "electric"
.Ymin "electric"
.Ymax "electric"
.Zmin "electric"
.Zmax "electric"
.Xsymmetry "none"
.Ysymmetry "none"
.Zsymmetry "none"
End With

'@ StoreParameter

'[VERSION]2022.4|31.0.1|20220426[/VERSION]
MakeSureParameterExists("a", "20.000000")
MakeSureParameterExists("b", "15.000000")

