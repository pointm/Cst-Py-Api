
def add_quotes_and_to_list(text):
    lines = text.splitlines()
    quoted_lines = ["" + line.lstrip() + "" for line in lines]
    return quoted_lines


text = """
Mesh.SetCreator "High Frequency" 

With FDSolver
     .Reset 
     .SetMethod "Tetrahedral", "General purpose" 
     .OrderTet "Second" 
     .OrderSrf "First" 
     .Stimulation "All", "All" 
     .ResetExcitationList 
     .AutoNormImpedance "False" 
     .NormingImpedance "50" 
     .ModesOnly "False" 
     .ConsiderPortLossesTet "True" 
     .SetShieldAllPorts "False" 
     .AccuracyHex "1e-6" 
     .AccuracyTet "1e-4" 
     .AccuracySrf "1e-3" 
     .LimitIterations "False" 
     .MaxIterations "0" 
     .SetCalcBlockExcitationsInParallel "True", "True", "" 
     .StoreAllResults "False" 
     .StoreResultsInCache "False" 
     .UseHelmholtzEquation "True" 
     .LowFrequencyStabilization "True" 
     .Type "Auto" 
     .MeshAdaptionHex "False" 
     .MeshAdaptionTet "True" 
     .AcceleratedRestart "True" 
     .FreqDistAdaptMode "Distributed" 
     .NewIterativeSolver "True" 
     .TDCompatibleMaterials "False" 
     .ExtrudeOpenBC "False" 
     .SetOpenBCTypeHex "Default" 
     .SetOpenBCTypeTet "Default" 
     .AddMonitorSamples "True" 
     .CalcPowerLoss "True" 
     .CalcPowerLossPerComponent "False" 
     .StoreSolutionCoefficients "True" 
     .UseDoublePrecision "False" 
     .UseDoublePrecision_ML "True" 
     .MixedOrderSrf "False" 
     .MixedOrderTet "False" 
     .PreconditionerAccuracyIntEq "0.15" 
     .MLFMMAccuracy "Default" 
     .MinMLFMMBoxSize "0.3" 
     .UseCFIEForCPECIntEq "True" 
     .UseEnhancedCFIE2 "True" 
     .UseFastRCSSweepIntEq "false" 
     .UseSensitivityAnalysis "False" 
     .UseEnhancedNFSImprint "False" 
     .RemoveAllStopCriteria "Hex"
     .AddStopCriterion "All S-Parameters", "0.01", "2", "Hex", "True"
     .AddStopCriterion "Reflection S-Parameters", "0.01", "2", "Hex", "False"
     .AddStopCriterion "Transmission S-Parameters", "0.01", "2", "Hex", "False"
     .RemoveAllStopCriteria "Tet"
     .AddStopCriterion "All S-Parameters", "0.01", "2", "Tet", "True"
     .AddStopCriterion "Reflection S-Parameters", "0.01", "2", "Tet", "False"
     .AddStopCriterion "Transmission S-Parameters", "0.01", "2", "Tet", "False"
     .AddStopCriterion "All Probes", "0.05", "2", "Tet", "True"
     .RemoveAllStopCriteria "Srf"
     .AddStopCriterion "All S-Parameters", "0.01", "2", "Srf", "True"
     .AddStopCriterion "Reflection S-Parameters", "0.01", "2", "Srf", "False"
     .AddStopCriterion "Transmission S-Parameters", "0.01", "2", "Srf", "False"
     .SweepMinimumSamples "3" 
     .SetNumberOfResultDataSamples "1001" 
     .SetResultDataSamplingMode "Automatic" 
     .SweepWeightEvanescent "1.0" 
     .AccuracyROM "1e-4" 
     .AddSampleInterval "", "", "1", "Automatic", "True" 
     .AddSampleInterval "", "", "", "Automatic", "False" 
     .MPIParallelization "False"
     .UseDistributedComputing "False"
     .NetworkComputingStrategy "RunRemote"
     .NetworkComputingJobCount "3"
     .UseParallelization "True"
     .MaxCPUs "1024"
     .MaximumNumberOfCPUDevices "2"
End With

With IESolver
     .Reset 
     .UseFastFrequencySweep "True" 
     .UseIEGroundPlane "False" 
     .SetRealGroundMaterialName "" 
     .CalcFarFieldInRealGround "False" 
     .RealGroundModelType "Auto" 
     .PreconditionerType "Auto" 
     .ExtendThinWireModelByWireNubs "False" 
     .ExtraPreconditioning "False" 
End With

With IESolver
     .SetFMMFFCalcStopLevel "0" 
     .SetFMMFFCalcNumInterpPoints "6" 
     .UseFMMFarfieldCalc "True" 
     .SetCFIEAlpha "0.500000" 
     .LowFrequencyStabilization "False" 
     .LowFrequencyStabilizationML "True" 
     .Multilayer "False" 
     .SetiMoMACC_I "0.0001" 
     .SetiMoMACC_M "0.0001" 
     .DeembedExternalPorts "True" 
     .SetOpenBC_XY "True" 
     .OldRCSSweepDefintion "False" 
     .SetRCSOptimizationProperties "True", "100", "0.00001" 
     .SetAccuracySetting "Custom" 
     .CalculateSParaforFieldsources "True" 
     .ModeTrackingCMA "True" 
     .NumberOfModesCMA "3" 
     .StartFrequencyCMA "-1.0" 
     .SetAccuracySettingCMA "Default" 
     .FrequencySamplesCMA "0" 
     .SetMemSettingCMA "Auto" 
     .CalculateModalWeightingCoefficientsCMA "True" 
     .DetectThinDielectrics "True" 
End With

With FDSolver
    .Start
End With
"""

print(add_quotes_and_to_list(text))

scom = ['', 'Mesh.SetCreator "High Frequency" ', '', 'With FDSolver', '.Reset ', '.SetMethod "Tetrahedral", "General purpose" ', '.OrderTet "Second" ', '.OrderSrf "First" ', '.Stimulation "All", "All" ', '.ResetExcitationList ', '.AutoNormImpedance "False" ', '.NormingImpedance "50" ', '.ModesOnly "False" ', '.ConsiderPortLossesTet "True" ', '.SetShieldAllPorts "False" ', '.AccuracyHex "1e-6" ', '.AccuracyTet "1e-4" ', '.AccuracySrf "1e-3" ', '.LimitIterations "False" ', '.MaxIterations "0" ', '.SetCalcBlockExcitationsInParallel "True", "True", "" ', '.StoreAllResults "False" ', '.StoreResultsInCache "False" ', '.UseHelmholtzEquation "True" ', '.LowFrequencyStabilization "True" ', '.Type "Auto" ', '.MeshAdaptionHex "False" ', '.MeshAdaptionTet "True" ', '.AcceleratedRestart "True" ', '.FreqDistAdaptMode "Distributed" ', '.NewIterativeSolver "True" ', '.TDCompatibleMaterials "False" ', '.ExtrudeOpenBC "False" ', '.SetOpenBCTypeHex "Default" ', '.SetOpenBCTypeTet "Default" ', '.AddMonitorSamples "True" ', '.CalcPowerLoss "True" ', '.CalcPowerLossPerComponent "False" ', '.StoreSolutionCoefficients "True" ', '.UseDoublePrecision "False" ', '.UseDoublePrecision_ML "True" ', '.MixedOrderSrf "False" ', '.MixedOrderTet "False" ', '.PreconditionerAccuracyIntEq "0.15" ', '.MLFMMAccuracy "Default" ', '.MinMLFMMBoxSize "0.3" ', '.UseCFIEForCPECIntEq "True" ', '.UseEnhancedCFIE2 "True" ', '.UseFastRCSSweepIntEq "false" ', '.UseSensitivityAnalysis "False" ', '.UseEnhancedNFSImprint "False" ', '.RemoveAllStopCriteria "Hex"', '.AddStopCriterion "All S-Parameters", "0.01", "2", "Hex", "True"', '.AddStopCriterion "Reflection S-Parameters", "0.01", "2", "Hex", "False"', '.AddStopCriterion "Transmission S-Parameters", "0.01", "2", "Hex", "False"', '.RemoveAllStopCriteria "Tet"', '.AddStopCriterion "All S-Parameters", "0.01", "2", "Tet", "True"', '.AddStopCriterion "Reflection S-Parameters", "0.01", "2", "Tet", "False"', '.AddStopCriterion "Transmission S-Parameters", "0.01", "2", "Tet", "False"',
        '.AddStopCriterion "All Probes", "0.05", "2", "Tet", "True"', '.RemoveAllStopCriteria "Srf"', '.AddStopCriterion "All S-Parameters", "0.01", "2", "Srf", "True"', '.AddStopCriterion "Reflection S-Parameters", "0.01", "2", "Srf", "False"', '.AddStopCriterion "Transmission S-Parameters", "0.01", "2", "Srf", "False"', '.SweepMinimumSamples "3" ', '.SetNumberOfResultDataSamples "1001" ', '.SetResultDataSamplingMode "Automatic" ', '.SweepWeightEvanescent "1.0" ', '.AccuracyROM "1e-4" ', '.AddSampleInterval "", "", "1", "Automatic", "True" ', '.AddSampleInterval "", "", "", "Automatic", "False" ', '.MPIParallelization "False"', '.UseDistributedComputing "False"', '.NetworkComputingStrategy "RunRemote"', '.NetworkComputingJobCount "3"', '.UseParallelization "True"', '.MaxCPUs "1024"', '.MaximumNumberOfCPUDevices "2"', 'End With', '', 'With IESolver', '.Reset ', '.UseFastFrequencySweep "True" ', '.UseIEGroundPlane "False" ', '.SetRealGroundMaterialName "" ', '.CalcFarFieldInRealGround "False" ', '.RealGroundModelType "Auto" ', '.PreconditionerType "Auto" ', '.ExtendThinWireModelByWireNubs "False" ', '.ExtraPreconditioning "False" ', 'End With', '', 'With IESolver', '.SetFMMFFCalcStopLevel "0" ', '.SetFMMFFCalcNumInterpPoints "6" ', '.UseFMMFarfieldCalc "True" ', '.SetCFIEAlpha "0.500000" ', '.LowFrequencyStabilization "False" ', '.LowFrequencyStabilizationML "True" ', '.Multilayer "False" ', '.SetiMoMACC_I "0.0001" ', '.SetiMoMACC_M "0.0001" ', '.DeembedExternalPorts "True" ', '.SetOpenBC_XY "True" ', '.OldRCSSweepDefintion "False" ', '.SetRCSOptimizationProperties "True", "100", "0.00001" ', '.SetAccuracySetting "Custom" ', '.CalculateSParaforFieldsources "True" ', '.ModeTrackingCMA "True" ', '.NumberOfModesCMA "3" ', '.StartFrequencyCMA "-1.0" ', '.SetAccuracySettingCMA "Default" ', '.FrequencySamplesCMA "0" ', '.SetMemSettingCMA "Auto" ', '.CalculateModalWeightingCoefficientsCMA "True" ', '.DetectThinDielectrics "True" ', 'End With', '', 'With FDSolver', '.Start', 'End With']
