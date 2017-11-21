using System;
using System.Runtime.InteropServices;

namespace P2PCalculations
{
    /// <summary>
    /// COM Interface for GetVoltage.
    /// </summary>
    [ComVisible(false)]
    [Guid("D287F838-4CB8-4C8B-987D-6A07DAF202E5")]
    // [InterfaceType ( ComInterfaceType.InterfaceIsDual )]
    // Following line changed due to FxCop rules.
    // "COM source interfaces should be marked ComInterfaceType.InterfaceIsIDispatch. 
    // Visual Basic 6 clients cannot receive events with non-IDispatch 
    // interfaces."
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IGetInjectionValues
    {
        /// <summary>
        /// Calculates new operating bandcenter per given DVArCurrent and LDC R, X setting.
        /// </summary>
        /// <param name="bandCenterSetting">The value of the unit's Bandcenter Setting.</param>
        /// <param name="dvarCurrent">DVAr current the unit is operating.</param>
        /// <param name="sensitivitySetting">The units Sensitivity setting.</param>
        /// <param name="reactiveIL">Reactive part of the Load Current applied to the unit.</param>
        /// <param name="appliedILPhase">Load Current Phase that applied to the unit.</param>
        /// <param name="ldcRSetting">LDC R setting the unit is currently operating.</param>
        /// <param name="ldcXSetting">LDC X setting the unit is currently operating.</param>
        /// <returns>Returns new Bandcenter that the unit should be operating.</returns>   
        [Obsolete("pu assumptions changed.", true)]
        double Voltage(double bandCenterSetting,
                         double dvarCurrent,
                         double sensitivitySetting,
                         double reactiveIL,
                         double appliedILPhase,
                         int ldcRSetting,
                         int ldcXSetting
                         );
        /// <summary>
        /// Calculates new operating bandcenter per given DVArCurrent and LDC R, X setting.
        /// </summary>
        /// <param name="bandCenterSetting">The value of the unit's Bandcenter Setting.</param>
        /// <param name="dvarCurrent">DVAr current the unit is operating.</param>
        /// <param name="sensitivitySetting">The units Sensitivity setting.</param>
        /// <param name="reactiveIL">Reactive part of the Load Current applied to the unit.</param>
        /// <param name="appliedILPhase">Load Current Phase that applied to the unit.</param>
        /// <param name="ldcRSetting">LDC R setting the unit is currently operating.</param>
        /// <param name="ldcXSetting">LDC X setting the unit is currently operating.</param>
        /// <param name="activePU">Value of the current power unit.</param>
        /// <param name="isParallel">paralleling network status of the unit.</param>
        /// <returns>Returns new Bandcenter that the unit should be operating.</returns>
        double Voltage(double bandCenterSetting,
                         double dvarCurrent,
                         double sensitivitySetting,
                         double reactiveIL,
                         double appliedILPhase,
                         int ldcRSetting,
                         int ldcXSetting,
                         double activePU,
                         bool isParallel
                         );

        /// <summary>
        /// Calculates Load Current Magnitude to inject to the unit.
        /// </summary>
        /// <param name="nominalIL">Real part of the load current to inject to the unit.</param>
        /// <param name="nominalIC">Circulating current magnitude to inject to the unit.</param>
        /// <param name="reactiveIL">Reactive part of the load current to inject to the unit.</param>
        /// <returns>Returns total Load Current Magnitude to inject to the unit.</returns>
        double ILMagnitude_ToApply(double nominalIL,
                                     double nominalIC,
                                     double reactiveIL
                                    );

        /// <summary>
        /// Calculates Load Current Magnitude to inject to the unit.
        /// </summary>
        /// <param name="lineBreaker">Line breaker status of the unit.</param>
        /// <param name="primaryIL">Primary load current value in Amps.</param>
        /// <param name="impedance">The impedance ratio of the unit vs other units in the paralleling scheme.</param>
        /// <param name="ct">The current multiplier value.</param>
        /// <param name="nominalIC">Calculated circulating current magnitude.</param>
        /// <param name="reactiveIL">The value of reactive part of the load current.</param>
        /// <returns>Returns calculated Load Current Magnitude to inject to the unit.</returns>
        double ILMagnitude(bool lineBreaker,
                             double primaryIL,
                             double impedance,
                             double ct,
                             double nominalIC,
                             double reactiveIL
                            );

        /// <summary>
        /// Calculates Load Current Phase to inject to the unit.
        /// </summary>
        /// <param name="nominalIL">Real part of the load current to inject to the unit.</param>
        /// <param name="nominalIC">Circulating current magnitude to inject to the unit.</param>
        /// <param name="reactiveIL">Reactive part of the load current to inject to the unit.</param>
        /// <returns>Returns total Load Current Phase to inject to the unit.</returns>
        double ILPhase(double nominalIL,
                         double nominalIC,
                         double reactiveIL
                        );

        /// <summary>
        /// Calculates equalize DVAr current due to CT rating difference if any.
        /// </summary>
        /// <param name="myIC">Circulating current of the unit.</param>
        /// <param name="myCT">CT rating of the unit.</param>
        /// <param name="maxCT">The highest CT rating in the paralleling scheme.</param>
        /// <returns>Returns equalized DeltaVAr Current per the highest CT rating.</returns>
        double DVARCurrent(double myIC,
                             int myCT,
                             int maxCT
                            );

        /// <summary>
        /// Detemines if the units are parallel
        /// </summary>
        /// <param name="A1">Line Breaker status - On = True, Off = False.</param>
        /// <param name="A2">Right Tie Breaker status - On = True, Off = False.</param>
        /// <param name="A3">Left Tie Breaker status - On = True, Off = False.</param>
        /// <returns>Returns True if the unit is in parallel state per specified breaker conditions.</returns>
        bool Paralleling(bool A1,
                           bool A2,
                           bool A3
                          );

        /// <summary>
        /// Calculate Circulating Current with given parameters.
        /// </summary>
        /// <param name="Unit1Status">Paralleling status of paralleling address = 1 in the current P2P scheme.</param>
        /// <param name="Unit2Status">Paralleling status of paralleling address = 2 in the current P2P scheme.</param>
        /// <param name="Unit3Status">Paralleling status of paralleling address = 3 in the current P2P scheme.</param>
        /// <param name="UnitCurrentToLearn">Target unit to calculate circulating current.</param>
        /// <param name="PrimaryIC">Total source side (primary) circulating current to calculate current.</param>
        /// <param name="Z1">Impedance of the paralleling address = 1 in the current P2P scheme.</param>
        /// <param name="Z2">Impedance of the paralleling address = 2 in the current P2P scheme.</param>
        /// <param name="CT">Target unit's CT value to calculate circulating current.</param>
        /// <param name="TotalParallelUnitNumber">Total number of the units in the paralleling scheme.</param>
        /// <returns>Returns calculated Circulating Current of the unit with given parameters.</returns>
        double CalculateIC(bool Unit1Status,
                             bool Unit2Status,
                             bool Unit3Status,
                             int UnitCurrentToLearn,
                             double PrimaryIC,
                             double Z1,
                             double Z2,
                             int CT,
                             int TotalParallelUnitNumber
                             );

        /// <summary>
        /// Calculates DVAr current to operate.
        /// </summary>
        /// <param name="IC1">Unit1 circulating current magnitude.</param>
        /// <param name="IC2">Unit2 circulating current magnitude.</param>
        /// <param name="IC3">Unit3 circulating current magnitude.</param>
        /// <param name="CT1">Unit1 CT multiplier.</param>
        /// <param name="CT2">Unit2 CT multiplier.</param>
        /// <param name="CT3">Unit2 CT multiplier.</param>
        /// <param name="MVA1">Unit1 MVA rating.</param>
        /// <param name="MVA2">Unit2 MVA rating.</param>
        /// <param name="MVA3">Unit3 MVA rating.</param>
        /// <param name="myIC">UnitX circulating current magnitude.</param>
        /// <param name="myCT">UnitX CT multiplier.</param>
        /// <param name="myMVA">UnitX MVA rating.</param>
        /// <param name="MaxCT">Maximum CT rating of the unit in the scheme.</param>
        /// <param name="MaxMVA">Maximum MVA rating of the unit in the scheme.</param>
        /// <returns>Returns DVAr current to set voltage offset in P2P Paralleling scheme.</returns>
        double calculateDVAr(double IC1,
                               double IC2,
                               double IC3,
                               double CT1,
                               double CT2,
                               double CT3,
                               double MVA1,
                               double MVA2,
                               double MVA3,
                               double myIC,
                               double myCT,
                               double myMVA,
                               double MaxMVA,
                               double MaxCT
                              );

        /// <summary>
        /// Calculates active power unit value to use in P2P Paralleling Operation Voltage calculations.
        /// </summary>
        /// <param name="primaryLineToLineVoltage">This is a setting in Paralleling Options.</param>
        /// <param name="threePhaseMaximumFullLoadCurrent">This is a setting in Paralleling Options.</param>
        /// <param name="primaryRealLoadCurrent">This is a test value.</param>
        /// <param name="primaryReactiveLoadCurrent">This is a test value.</param>
        /// <returns>Calculated power units thay used to calculate P2P Paralleling Operation Voltage.</returns>
        double calculateActivePU(double primaryLineToLineVoltage,
                                    double threePhaseMaximumFullLoadCurrent,
                                    double primaryRealLoadCurrent,
                                    double primaryReactiveLoadCurrent
                                  );


        /// <summary>
        /// Count active parallel unit number.
        /// </summary>
        /// <param name="unitParallelStatus">Paralleling status of the units.</param>
        /// <returns>Returns total number of the parallel units.</returns>
        int countParallelUnits(bool[] unitParallelStatus);

    }
}
