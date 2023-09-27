def CstDefineFrequencydomainSolver(mws, startFreq, endFreq, samples):
    changeSolverType = mws.ChangeSolverType('HF Frequency Domain')
    mesh = mws.Mesh
    fdSolver = mws.FDSolver
    ieSolver = mws.IESolver

    mesh.SetCreator('High Frequency')

    fdSolver.Reset()
    fdSolver.SetMethod('Tetrahedral', 'General purpose')
    fdSolver.OrderTet('Second')
    fdSolver.OrderSrf('First')
    fdSolver.Stimulation("All", "All")
    fdSolver.ResetExcitationList()
    fdSolver.AutoNormImpedance('True')
    fdSolver.NormingImpedance('50')
    fdSolver.ModesOnly('False')
    fdSolver.ConsiderPortLossesTet('True')
    fdSolver.SetShieldAllPorts('False')
    fdSolver.AccuracyHex('1e-6')
    fdSolver.AccuracyTet('1e-4')
    fdSolver.AccuracySrf('1e-3')
    fdSolver.LimitIterations('False')
    fdSolver.MaxIterations('0')
    fdSolver.SetCalcBlockExcitationsInParallel('True', 'True', '')
    fdSolver.StoreAllResults('False')
    fdSolver.StoreResultsInCache('False')
    fdSolver.UseHelmholtzEquation('True')
    fdSolver.LowFrequencyStabilization('True')
    fdSolver.Type('Auto')
    fdSolver.MeshAdaptionHex('False')
    fdSolver.MeshAdaptionTet('True')
    fdSolver.AcceleratedRestart('True')
    fdSolver.FreqDistAdaptMode('Distributed')
    fdSolver.NewIterativeSolver('True')  # warning
    fdSolver.TDCompatibleMaterials('False')
    fdSolver.ExtrudeOpenBC('False')
    fdSolver.SetOpenBCTypeHex('Default')
    fdSolver.SetOpenBCTypeTet('Default')
    fdSolver.AddMonitorSamples('True')
    fdSolver.CalcStatBField('False')
    fdSolver.CalcPowerLoss('True')
    fdSolver.CalcPowerLossPerComponent('False')
    fdSolver.StoreSolutionCoefficients('True')
    fdSolver.UseDoublePrecision('False')
    fdSolver.UseDoublePrecision_ML('True')  # warning
    fdSolver.MixedOrderSrf('False')
    fdSolver.MixedOrderTet('False')
    fdSolver.PreconditionerAccuracyIntEq('0.15')
    fdSolver.MLFMMAccuracy('Default')
    fdSolver.MinMLFMMBoxSize('0.3')
    fdSolver.UseCFIEForCPECIntEq('True')
    fdSolver.UseFastRCSSweepIntEq('False')
    fdSolver.UseSensitivityAnalysis('False')
    fdSolver.RemoveAllStopCriteria('Hex')
    fdSolver.AddStopCriterion('All S-Parameters', '0.01', '2', 'Hex', 'True')
    fdSolver.AddStopCriterion('Reflection S-Parameters', '0.01', '2', 'Hex', 'False')
    fdSolver.AddStopCriterion('Transmission S-Parameters', '0.01', '2', 'Hex', 'False')
    fdSolver.RemoveAllStopCriteria('Tet')
    fdSolver.AddStopCriterion('All S-Parameters', '0.01', '2', 'Tet', 'True')
    fdSolver.AddStopCriterion('Reflection S-Parameters', '0.01', '2', 'Tet', 'False')
    fdSolver.AddStopCriterion('Transmission S-Parameters', '0.01', '2', 'Tet', 'False')
    fdSolver.RemoveAllStopCriteria('Srf')
    fdSolver.AddStopCriterion('All S-Parameters', '0.01', '2', 'Srf', 'True')
    fdSolver.AddStopCriterion('Reflection S-Parameters', '0.01', '2', 'Srf', 'False')
    fdSolver.AddStopCriterion('Transmission S-Parameters', '0.01', '2', 'Srf', 'False')
    fdSolver.SweepMinimumSamples('3')
    fdSolver.SetNumberOfResultDataSamples('1001')
    fdSolver.SetResultDataSamplingMode('Automatic')
    fdSolver.SweepWeightEvanescent('1.0')
    fdSolver.AccuracyROM('1e-4')
    fdSolver.AddSampleInterval('', '', '1', 'Automatic', 'True')
    fdSolver.AddSampleInterval(str(startFreq), str(endFreq), samples, 'Automatic', 'False')
    fdSolver.MPIParallelization('False')
    fdSolver.UseDistributedComputing('False')
    fdSolver.NetworkComputingStrategy('RunRemote')
    fdSolver.NetworkComputingJobCount('3')
    fdSolver.LimitCPUs('True')
    fdSolver.MaxCPUs('96')
    fdSolver.MaximumNumberOfCPUDevices('2')

    ieSolver.Reset()
    ieSolver.UseFastFrequencySweep('True')
    ieSolver.UseIEGroundPlane('False')
    ieSolver.CalcFarFieldInRealGround('False')
    ieSolver.RealGroundModelType('Auto')
    ieSolver.PreconditionerType('Auto')
    ieSolver.ExtendThinWireModelByWireNubs('False')
    ieSolver.SetFMMFFCalcStopLevel('0')
    ieSolver.SetFMMFFCalcNumInterpPoints('6')
    ieSolver.UseFMMFarfieldCalc('True')
    ieSolver.SetCFIEAlpha('0.500000')
    ieSolver.LowFrequencyStabilization('False')
    ieSolver.LowFrequencyStabilizationML('True')
    ieSolver.Multilayer('False')
    ieSolver.SetiMoMACC_I('0.0001')
    ieSolver.SetiMoMACC_M('0.0001')
    ieSolver.DeembedExternalPorts('True')
    ieSolver.SetOpenBC_XY('True')
    ieSolver.OldRCSSweepDefintion('False')
    ieSolver.SetAccuracySetting('Custom')
    ieSolver.CalculateSParaforFieldsources('True')
    ieSolver.ModeTrackingCMA('True')
    ieSolver.NumberOfModesCMA('3')
    ieSolver.StartFrequencyCMA('-1.0')
    ieSolver.SetAccuracySettingCMA('Default')
    ieSolver.FrequencySamplesCMA('0')
    ieSolver.SetMemSettingCMA('Auto')

    fdSolver.Start()
