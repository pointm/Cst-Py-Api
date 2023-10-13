a = 50
b = 30
Command = '''
With Port
    .Reset
    .PortNumber 2
    .Label  ""
    .NumberOfModes 1
    .AdjustPolarization "False"
    .PolarizationAngle 0.0
    .ReferencePlaneDistance 0
    .TextSize 50
    .TextMaxLimit 0
    .Coordinates "Picks"
    .Orientation "positive"
    .PortOnBound "False"
    .ClipPickedPortToBound "False"
    '''
Command = Command + '''
    .Xrange "%s","%s"
    .Yrange "%s","%s"
    .Zrange "-l/2","-l/2"
''' % ('alpha', 'alpha', 'beta', 'beta')
Command = Command + '''
    .XrangeAdd "0.0","0.0"
    .YrangeAdd "0.0","0.0"
    .ZrangeAdd "0.0","0.0"
    .SingleEnded "False"
    .Create
End With
'''

print(Command)
