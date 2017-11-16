using System;
using System.Numerics;
using System.Runtime.InteropServices;

namespace P2PCalculations
{
    /// <summary>
    /// Calculates injection values for P2P DVAr3 Paralleling scheme.
    /// </summary>
    [ComVisible(false)]
    [Guid("7E45A2BC-E158-4335-83B6-8D2E4627193F")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComSourceInterfaces(typeof(IGetInjectionValues))]
    [ProgId("P2PCalculations.GetVoltage")]
    public class GetInjectionValues : IGetInjectionValues
    {
        /// <summary>
        /// Default constructor
        /// </summary>
        public GetInjectionValues() { }

        /// <summary>
        /// 1mA DVAr current is equal to 0.12V
        /// </summary>
        [ComVisible(false)]
        const double VoltageOffsetPerDVArCurrent = 0.12F * 1000F;

        /// <summary>
        /// If Sensitivity is smaller than 0 then Sensitivity minimum is 50%
        /// </summary>
        [ComVisible(false)]
        const double NegativeSensitivityMultiplier = 12.5F;

        /// <summary>
        /// If Sensitivity is greater than 0 then Sensitivity maximum is 200%
        /// </summary>
        [ComVisible(false)]
        const double PositiveSensitivityMultiplier = 25.0F;

        /// <summary>
        /// 1pu is equivalent of 200mA or 0.2A 
        /// </summary>
        [ComVisible(false)]
        const double PU = 0.2F;

        /// <summary>
        /// If LineBreaker opens what is the load current
        /// Per T.Branch this should be zero all the time 
        /// I have this variable in case in the future 
        /// this might also changes like everything else.
        /// </summary>
        [ComVisible(false)]
        const double CurrentAtLineBreakerOpen = 0.0F;

        /// <summary>
        /// Calculates new operating bandcenter per given DVArCurrent.
        /// </summary>
        /// <param name="BandCenterSetting">This is what the units Bandcenter Setting.</param>
        /// <param name="DVArCurrent">DVAr current the unit is operating.</param>
        /// <param name="SensitivitySetting">The units Sensitivity setting.</param>
        /// <returns>Returns new Bandcenter that the unit should be operating.</returns>
        [ComVisible(false)]
        private double Voltage(double BandCenterSetting, double DVArCurrent, double SensitivitySetting)
        {
            double SensitivityMultiplier;

            if (SensitivitySetting < 0)
            {
                SensitivityMultiplier = SensitivitySetting * NegativeSensitivityMultiplier;
            }
            else
            {
                SensitivityMultiplier = SensitivitySetting * PositiveSensitivityMultiplier;
            }

            // Calculating percentage of the Sensitivity 
            // Ex: Sensitivity = 2 => SensitivityMultiplier = 1.5
            SensitivityMultiplier = (SensitivityMultiplier + 100F) / 100F;

            return (BandCenterSetting - (DVArCurrent * VoltageOffsetPerDVArCurrent * SensitivityMultiplier));
        }

        /// <summary>
        /// Calculates new operating bandcenter per given DVArCurrent and LDC R, X setting.
        /// </summary>
        /// <param name="BandCenterSetting">This is what the units Bandcenter Setting.</param>
        /// <param name="DVArCurrent">DVAr current the unit is operating.</param>
        /// <param name="SensitivitySetting">The units Sensitivity setting.</param>
        /// <param name="reactiveIL">Reactive part of the Load Current applied to the unit.</param>
        /// <param name="AppliedILPhase">Load Current Phase that applied to the unit.</param>
        /// <param name="ldcRSetting">LDC R setting the unit is currently operating.</param>
        /// <param name="ldcXSetting">LDC X setting the unit is currently operating.</param>
        /// <returns>Returns new Bandcenter that the unit should be operating.</returns>
        [ComVisible(false)]
        public double Voltage(double BandCenterSetting, double DVArCurrent, double SensitivitySetting, double reactiveIL, double AppliedILPhase, int ldcRSetting, int ldcXSetting)
        {
            
            double Vload = Voltage(BandCenterSetting, DVArCurrent, SensitivitySetting);
            double activePU = reactiveIL / PU;
            double radians = Math.PI / 180 * AppliedILPhase;
            double cos = Math.Cos(radians);
            double sin = Math.Sin(radians);

            Complex Rset = new Complex(cos, (sin)); //* (-1)));
            Complex Xset = new Complex(sin, (cos)) * (-1); // * (-1)));

            Complex Vbus = new Complex();

            // Following calculation is from Tap Changer Controls Application Note #17
            // Bus Voltage conditions when using Line Drop Compensation
            // Page 14 Date: 2/98
            // VB = VL + I [ R ( cos ⁡φ - j sin⁡ φ ) + X ( sin ⁡φ + j cos ⁡φ ) ]
            Vbus = Complex.Add(Vload, Complex.Multiply(activePU, Complex.Add(Complex.Multiply(ldcRSetting, Rset), Complex.Multiply(ldcXSetting, Xset))));

            return Complex.Abs(Vbus);
        }

        /// <summary>
        /// Calculates Load Current Magnitude to inject to the unit.
        /// </summary>
        /// <param name="nominalIL">Real part of the load current to inject to the unit.</param>
        /// <param name="nominalIC">Circulating current magnitude to inject to the unit.</param>
        /// <param name="reactiveIL">Reactive part of the load current to inject to the unit.</param>
        /// <returns>Returns total Load Current Magnitude to inject to the unit.</returns>
        [ComVisible(false)]
        public double ILMagnitude_ToApply(double nominalIL, double nominalIC, double reactiveIL)
        {
            double totalReactiveCurrent = nominalIC + reactiveIL;
            return Math.Sqrt(Math.Pow(nominalIL, 2) + Math.Pow(totalReactiveCurrent, 2));
        }

        /// <summary>
        /// Calculates Load Current Magnitude to inject to the unit.
        /// </summary>
        /// <param name="A1">Line breaker status of the unit.</param>
        /// <param name="PrimaryIL">Primary load current value in Amps.</param>
        /// <param name="Z">The impedance ratio of the unit vs other units in the paralleling scheme.</param>
        /// <param name="CT">The current multiplier value.</param>
        /// <param name="nominalIC">Calculated circulating current magnitude.</param>
        /// <param name="reactiveIL">The value of reactive part of the load current.</param>
        /// <returns>Returns calculated Load Current Magnitude to inject to the unit.</returns>
        [ComVisible(false)]
        public double ILMagnitude(bool A1, double PrimaryIL, double Z, double CT, double nominalIC, double reactiveIL)
        {
            if (A1)
            {
                // the unit in paralleling scheme.
                // Calculating real part of the IL to applied.
                double nominalIL = PrimaryIL * Z / CT;

                return nominalIL;
            }
            else
            {
                // the unit is NOT paralleling.
                return PrimaryIL * CurrentAtLineBreakerOpen;
            }
        }

        /// <summary>
        /// Calculates Load Current Phase to inject to the unit.
        /// </summary>
        /// <param name="nominalIL">Real part of the load current to inject to the unit.</param>
        /// <param name="nominalIC">Circulating current magnitude to inject to the unit.</param>
        /// <param name="reactiveIL">Reactive part of the load current to inject to the unit.</param>
        /// <returns>Returns total Load Current Phase to inject to the unit.</returns>
        [ComVisible(false)]
        public double ILPhase(double nominalIL, double nominalIC, double reactiveIL)
        {
            double totalReactiveCurrent = nominalIC + reactiveIL;
            double radians;

            // If nominal IL = 0 this function returns NaN so verify nominal IL is not zero,
            // If it is return 0.0 radians.
            // if ( nominalIL != 0 )
            // Math.Atan2 function would return a NaN value if both of the quotient are ZERO.
            if (nominalIL == 0 && nominalIC + reactiveIL == 0)
            {
                radians = 0.0;
            }
            else
            {
                // As of 6/23/2017, any P2P operations must work in any quandrant,
                // hence Math.Atan2 must be used instead of Math.Atan.

                //radians = Math.Atan ( totalReactiveCurrent / nominalIL );
                radians = Math.Atan2(totalReactiveCurrent, nominalIL);
            }
            return radians * 180 / Math.PI;
        }

        /// <summary>
        /// Calculates Load Current Phase.
        /// </summary>
        /// <param name="nominalIL">Real part of the load current to inject to the unit.</param>
        /// <param name="reactiveIL">Reactive part of the load current to inject to the unit.</param>
        /// <returns>Returns total Load Current Phase to calculate voltage to inject to the unit.</returns>
        [ComVisible(false)]
        public double LoadCurrentPhase(double nominalIL, double reactiveIL)
        {
            double radians;

            if (nominalIL == 0 && reactiveIL == 0)
            {
                radians = 0.0;
            }
            else
            {
                // As of 6/23/2017, any P2P operations must work in any quandrant,
                // hence Math.Atan2 must be used instead of Math.Atan.
                radians = Math.Atan2(reactiveIL, nominalIL);
            }
            return radians * 180 / Math.PI;
        }

        /// <summary>
        /// Calculates equalize DVAr current due to CT rating difference if any.
        /// </summary>
        /// <param name="myIC">Circulating current of the unit.</param>
        /// <param name="myCT">CT rating of the unit.</param>
        /// <param name="MaxCT">The highest CT rating in the paralleling scheme.</param>
        /// <returns>Returns equalized DeltaVAr Current per the highest CT rating.</returns>
        [ComVisible(false)]
        public double DVARCurrent(double myIC, int myCT, int MaxCT)
        {
            // Delta VAr current is always reverse to the myIC (the circulating current)
            return (myIC * (double)myCT / (double)MaxCT);
        }

        /// <summary>
        /// Detemines if the units are parallel
        /// </summary>
        /// <param name="A1">Line Breaker status - On = True, Off = False.</param>
        /// <param name="A2">Right Tie Breaker status - On = True, Off = False.</param>
        /// <param name="A3">Left Tie Breaker status - On = True, Off = False.</param>
        /// <returns>Returns True if the unit is in parallel state per specified breaker conditions.</returns>
        [ComVisible(false)]
        public bool Paralleling(bool A1, bool A2, bool A3)
        {
            return A1 && ((!A2 && A3) || (A2 && !A3) || (A2 && A3));
        }

        /// <summary>
        /// Calculates DVAr correctly if only two units are parallel and unit3 is NOT in this paralleling scheme.
        /// </summary>
        /// <param name="Unit2Status">Unit 2 paralleling status.</param>
        /// <param name="Unit1Status">Unit 1 paralleling status.</param>
        /// <param name="Unit3Status">Unit 3 paralleling status.</param>
        /// <param name="PrimaryIC">This total circulating current of the system.</param>
        /// <param name="Z2">Unit 2 impedance ratio vs unit 1 after unit 3 is set to zero.</param>
        /// <param name="Z1">Unit 1 impedance ratio vs unit 2 after unit 3 is set to zero.</param>
        /// <param name="CT">CT ratio of unit 2.</param>
        /// <returns>In a 3-unit system, If unit 3 is not in the paralelling scheme there will be no opposite circulating can be calculated
        /// This function correctly calculates these currents.</returns>
        [ComVisible(false)]
        [Obsolete("This function is deprecated, please use CalculateIC method.", true)]
        public double Unit2ParallelingCurrent(bool Unit2Status, bool Unit1Status, bool Unit3Status, double PrimaryIC, double Z2, double Z1, int CT)
        {
            int currentDirection = 1;

            if (Unit3Status)
            {
                currentDirection = -1;
            }

            if (Unit2Status && (Unit1Status || Unit3Status))
            {
                return PrimaryIC * (Z2 / (Z1 + Z2) / CT) * currentDirection;
            }
            else
            {
                return 0;
            }
        }

        /// <summary>
        /// Calculate Circulating Current with given parameters.
        /// </summary>
        /// <param name="Unit1Status">Paralleling status of paralleling address = 1 in the current P2P scheme.</param>
        /// <param name="Unit2Status">Paralleling status of paralleling address = 2 in the current P2P scheme.</param>
        /// <param name="Unit3Status">Paralleling status of paralleling address = 1 in the current P2P scheme.</param>
        /// <param name="UnitCurrentToLearn">Target unit to calculate circulating current.</param>
        /// <param name="PrimaryIC">Total source side (primary) circulating current to calculate current.</param>
        /// <param name="Z1">Impedance of the paralleling address = 1 in the current P2P scheme.</param>
        /// <param name="Z2">Impedance of the paralleling address = 2 in the current P2P scheme.</param>
        /// <param name="CT">Target unit's CT value to calculate circulating current.</param>
        /// <param name="TotalParallelUnitNumber">Total number of the units in the paralleling scheme.</param>
        /// <returns>Returns calculated Circulating Current of the unit with given parameters.</returns>
        [ComVisible(false)]
        public double CalculateIC(bool Unit1Status, bool Unit2Status, bool Unit3Status, int UnitCurrentToLearn, double PrimaryIC, double Z1, double Z2, int CT, int TotalParallelUnitNumber)
        {
            double circulatingCurrent = 0.0F;
            const int inverse = -1;

            switch (TotalParallelUnitNumber)
            {
                case 0:
                case 1:
                    // No Paralleling Active.
                    // Circulating Current = 0.
                    circulatingCurrent = 0.0F;
                    break;

                case 2:
                    // Paralleling active with two units only.
                    // Both units circulating current must be same,
                    // if they have same CT rating.
                    switch (UnitCurrentToLearn)
                    {
                        case 1:
                            // Calculate Unit# 1, if Unit1 is paralleled.
                            if (Unit1Status)
                            {
                                circulatingCurrent = PrimaryIC / CT * inverse;
                            }
                            else
                            {
                                circulatingCurrent = 0.0F;
                            }
                            break;

                        case 2:
                            // Calculate Unit# 2, if Unit2 is paralleled.
                            if (Unit2Status)
                            {
                                if (Unit3Status)
                                {
                                    circulatingCurrent = PrimaryIC / CT * inverse;
                                }
                                else
                                {
                                    circulatingCurrent = PrimaryIC / CT;
                                }

                            }
                            else
                            {
                                circulatingCurrent = 0.0F;
                            }
                            break;

                        case 3:
                            // Calculate Unit# 3, if Unit3 is paralleled.
                            if (Unit3Status)
                            {
                                circulatingCurrent = PrimaryIC / CT;
                            }
                            else
                            {
                                circulatingCurrent = 0.0F;
                            }
                            break;

                        default:
                            circulatingCurrent = 0.0F;
                            break;
                    }
                    break;

                case 3:
                    // Paralleling active with three units.
                    switch (UnitCurrentToLearn)
                    {
                        case 1:
                            // Calculate Unit# 1, if Unit1 is paralleled.
                            circulatingCurrent = PrimaryIC * (Z1 / (Z1 + Z2)) / CT * inverse;
                            break;

                        case 2:
                            // Calculate Unit# 2, if Unit2 is paralleled.
                            circulatingCurrent = PrimaryIC * (Z2 / (Z1 + Z2)) / CT * inverse;
                            break;

                        case 3:
                            // Calculate Unit# 3, if Unit3 is paralleled.
                            circulatingCurrent = PrimaryIC / CT;
                            break;

                        default:
                            // returning 0 just in case of user input error.
                            circulatingCurrent = 0.0F;
                            break;
                    }
                    break;

                default:
                    // returning 0 just in case of user input error.
                    circulatingCurrent = 0;
                    break;
            }

            return circulatingCurrent;
        }

        /// <summary>
        /// Calculates DVAr current to operate.
        /// </summary>
        /// <param name="IC1">Unit1 circulating current magnitude.</param>
        /// <param name="IC2">Unit2 circulating current magnitude.</param>
        /// <param name="IC3">Unit3 circulating current magnitude.</param>
        /// <param name="unitIC">Unitx circulating current magnitude.</param>
        /// <returns>Returns average of the circulating currents minus its own circulating current.</returns>
        [ComVisible(false)]
        static double AverageDVAr(double IC1, double IC2, double IC3, double unitIC)
        {
            // to add more precision convert currents to mA.
            double[] currents = { IC1 * 1000F, IC2 * 1000F, IC3 * 1000F };
            TestCase tc = new TestCase();

            int maxNetworkDevice = (int)tc.GetSettings("maxnetworkdevice", 0);
            int totalUnitsInNetwork = 0;

            foreach (double current in currents)
            {
                if (current != 0.0F)
                {
                    totalUnitsInNetwork++;
                }
            }

            if (unitIC == 0)
            {
                return 0;
            }
            else
            {
                return ((IC1 + IC2 + IC3) / totalUnitsInNetwork) - unitIC;
            }
        }

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
        /// <param name="MaxCT">Maximum CT rating.</param>
        /// <param name="MaxMVA">Maximum MVA rating in the scheme.</param>
        /// <returns>Returns DVAr current to set voltage offset in P2P Paralleling scheme.</returns>
        [ComVisible(false)]
        public double calculateDVAr(double IC1, double IC2, double IC3, double CT1,
                                        double CT2, double CT3, double MVA1, double MVA2,
                                            double MVA3, double myIC, double myCT, double myMVA,
                                                double MaxMVA, double MaxCT)
        {
            double correctedIC1 = CorrectedIC(IC1, CT1, MVA1, MaxMVA, MaxCT);
            double correctedIC2 = CorrectedIC(IC2, CT2, MVA2, MaxMVA, MaxCT);
            double correctedIC3 = CorrectedIC(IC3, CT3, MVA3, MaxMVA, MaxCT);
            double correctedMyIC = CorrectedIC(myIC, myCT, myMVA, MaxMVA, MaxCT);

            return AverageDVAr(correctedIC1, correctedIC2, correctedIC3, correctedMyIC);
        }

        /// <summary>
        /// Calculates CT ratio corrected circulating current.
        /// </summary>
        /// <param name="myIC">UnitX circulating current magnitude.</param>
        /// <param name="myCT">UnitX CT multiplier.</param>
        /// <param name="myMVA">UnitX MVA rating.</param>
        /// <param name="MaxMVA">Maximum MVA rating in the scheme.</param>
        /// <param name="MaxCT">Maximum CT rating.</param>
        /// <returns>Returns calculated CT ratio corrected circulating current.</returns>
        [ComVisible(false)]
        private double CorrectedIC(double myIC, double myCT, double myMVA, double MaxMVA, double MaxCT)
        {
            return (myIC * myCT * MaxMVA) / (myMVA * MaxCT);
        }

        /// <summary>
        /// Count active parallel unit number.
        /// </summary>
        /// <param name="unitParallelStatus">Paralleling status of the units.</param>
        /// <returns>Returns total number of the parallel units.</returns>
        [ComVisible(false)]
        public int countParallelUnits(bool[] unitParallelStatus)
        {
            int result = 0;

            foreach (var status in unitParallelStatus)
            {
                if (status)
                    result++;
            }

            return result;
        }

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
        [ComVisible(false)]
        public double Voltage(double bandCenterSetting, double dvarCurrent, double sensitivitySetting, double reactiveIL, double appliedILPhase, int ldcRSetting, int ldcXSetting, double activePU, bool isParallel)
        {
            Complex Rset = new Complex();
            Complex Xset = new Complex();
            Complex Vbus = new Complex();
            Complex RValue = new Complex();
            Complex XValue = new Complex();
            Complex LDCTotal = new Complex();

            double Vload = Voltage(bandCenterSetting, dvarCurrent, sensitivitySetting);

            // as of 6/23/2017 pu calculations has changed. It is no longer based on the 200mA, 
            // but rather based on the user inputs in the configuration menu.

            double radians = Math.PI / 180 * appliedILPhase;
            double cos = Math.Cos(radians);
            double sin = Math.Sin(radians);

            // Following calculation is from Tap Changer Controls Application Note #17
            // Bus Voltage conditions when using Line Drop Compensation
            // Page 14 Date: 2/98
            // VB = VL + I [ R ( cos ⁡φ - j sin⁡ φ ) + X ( sin ⁡φ + j cos ⁡φ ) ]

            Rset = Complex.Add(new Complex(cos, 0.0), new Complex(0.0, sin));
            Xset = Complex.Negate(Complex.Add(new Complex(sin, 0.0), new Complex(0.0, cos)));

            RValue = Complex.Multiply(ldcRSetting, Rset);
            XValue = Complex.Multiply(ldcXSetting, Xset);

            // following code fails to calculate the bandcenter correctly, if the unit is out of paralleling network.
            // need to provide paralleling status of the unit.
            // 8/22/2017. TB.            
            // LDCTotal = Complex.Multiply(activePU, Complex.Add(RValue, XValue));
            if (isParallel)
            {
                LDCTotal = Complex.Multiply(activePU, Complex.Add(RValue, XValue));
            }

            Vbus = Complex.Add(Vload, LDCTotal);

            return Complex.Abs(Vbus);
        }


        /// <summary>
        /// Calculates active power unit value to use in P2P Paralleling Operation Voltage calculations.
        /// </summary>
        /// <param name="primaryLineToLineVoltage">This is a setting in Paralleling Options.</param>
        /// <param name="threePhaseMaximumFullLoadCurrent">This is a setting in Paralleling Options.</param>
        /// <param name="primaryRealLoadCurrent">This is a test value.</param>
        /// <param name="primaryReactiveLoadCurrent">This is a test value.</param>
        /// <returns>Calculated power units thay used to calculate P2P Paralleling Operation Voltage.</returns>
        [ComVisible(false)]
        public double calculateActivePU(double primaryLineToLineVoltage,
                                    double threePhaseMaximumFullLoadCurrent,
                                    double primaryRealLoadCurrent,
                                    double primaryReactiveLoadCurrent
                                  )
        {

            double iRated = threePhaseMaximumFullLoadCurrent / (Math.Pow(3, 0.5) * primaryLineToLineVoltage);
            double iTotal = Complex.Abs(new Complex(primaryRealLoadCurrent, primaryReactiveLoadCurrent));

            return iTotal / iRated;

        }
    }
}
