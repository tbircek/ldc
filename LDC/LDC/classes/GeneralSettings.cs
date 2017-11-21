using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Globalization;

namespace P2PCalculations
{
    /// <summary>
    /// Test cases for P2P DVAr3 Paralleling scheme.
    /// </summary>
    [ComVisible(false)]
    [Guid("5D6A60C9-033A-465D-9E95-468C0F9E224E")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComSourceInterfaces(typeof(IGeneralSettings))]
    [ProgId("P2PCalculations.TestCase")]
    public class TestCase : IGeneralSettings
    {
        /// <summary>
        /// Default constructor
        /// </summary>
        public TestCase() { }

        /// <summary>
        /// Total test numbers that fits into regular test patterns.
        /// There are 4 more cases that not include in this number, don't fit into these test patterns.
        /// In these cases Unit3 will have same CT, MVA and Impedance settings as other units.
        /// </summary>
        [ComVisible(false)]
        public int TotalRegularTest
        {
            get { return (int)GetSettings("totalregulartest", 0); }
        }

        /// <summary>
        /// Total test numbers to test for 3-Unit systems.
        /// </summary>
        [ComVisible(false)]
        public int TotalTestNumbers
        {
            get { return (int)GetSettings("totaltestnumbers", 0); }
        }

        /// <summary>
        /// P2P tests designed to only allow one unit to carry all DeltaVAr current.
        /// This is the highest paralleling address of the unit in the paralleling scheme.
        /// </summary>
        [ComVisible(false)]
        public int InitiatorPosition
        {
            get { return (int)GetSettings("maxnetworkdevice", 0); }
        }

        /// <summary>
        /// Reads settings values from"Setup" file.
        /// </summary>
        /// <param name="itemName">Setting name in the "Setup" file.</param>
        /// <param name="UnitNumber">It is 0 for if a setting is common for all the units, 
        /// and it is unit's paralleling address if not same for the all units.</param>
        /// <returns></returns>
        [ComVisible(false)]
        public double GetSettings(String itemName, int UnitNumber)
        {

            int unitNumberOffset = -1;

            if (UnitNumber == 0)
            {
                unitNumberOffset = 0;
            }

            UCL ucl = new UCL();

            return Convert.ToDouble(ucl.GetValues(itemName, (UnitNumber + unitNumberOffset)), new CultureInfo("en-US"));

        }

        /// <summary>
        /// MVA settings per Test Case is being run.
        /// </summary>
        /// <param name="CaseNumber">Test case number of the "Grouping Schemes" document.</param>
        /// <param name="UnitNumber">Unit's paralleling address.</param>
        /// <returns>Returns MVA setting value per "Setup" document.</returns>
        [ComVisible(false)]
        public double MvaCase(int CaseNumber, int UnitNumber)
        {
            // Unit 3 MVA will always change accordingly to test case.
            // Test cases > 72 will always have same MVA rate at 100.
            if (UnitNumber == InitiatorPosition && CaseNumber <= TotalRegularTest)
            {
                switch (CaseNumber % 4)
                {
                    case 3:
                        // MVA rate at 50.
                        CaseNumber = 1;
                        break;
                    default:
                        // MVA rate at 100.
                        CaseNumber = 0;
                        break;
                }
            }
            else
            {
                // MVA rate at 100.
                CaseNumber = 0;
            }

            UCL ucl = new UCL();

            return Convert.ToDouble(ucl.GetValues("mva", CaseNumber), new CultureInfo("en-US"));
        }

        /// <summary>
        /// CT settings per Test Case is being run.
        /// </summary>
        /// <param name="CaseNumber">Test case number of the "Grouping Schemes" document.</param>
        /// <param name="UnitNumber">Unit's paralleling address.</param>
        /// <returns>Returns CT setting value per P2P Test Setup.xml document.</returns>
        [ComVisible(false)]
        public double CTCase(int CaseNumber, int UnitNumber)
        {
            // Unit 3 CT will always change accordingly to test case.
            // Test cases > 72 will always have same CT ratio at 5000.
            if (UnitNumber == InitiatorPosition && CaseNumber <= TotalRegularTest)
            {
                switch (CaseNumber % 4)
                {
                    case 1:
                        // CT ratio is 4500.
                        CaseNumber = 1;
                        break;
                    case 3:
                        // CT ratio is 2500.
                        CaseNumber = 2;
                        break;
                    default:
                        // CT ratio is 5000.
                        CaseNumber = 0;
                        break;
                }
            }
            else
            {
                CaseNumber = 0;
            }

            UCL ucl = new UCL();

            return Convert.ToDouble(ucl.GetValues("ct", CaseNumber), new CultureInfo("en-US"));
        }

        /// <summary>
        /// Impedance assumption per Test Case is being run.
        /// </summary>
        /// <param name="CaseNumber">Test case number of the "Grouping Schemes" document.</param>
        /// <param name="UnitNumber">Unit's paralleling address.</param>
        /// <returns>Returns Impedance assumption value per the "Setup" document.</returns>
        [ComVisible(false)]
        public double ImpedanceCase(int CaseNumber, int UnitNumber)
        {
            // Unit 3 CT will always change accordingly to test case.
            // Test cases > 72 will always have same Impedance assumption at 10.
            if (UnitNumber == InitiatorPosition && CaseNumber <= TotalRegularTest)
            {
                switch (CaseNumber % 4)
                {
                    case 0:
                        // Impedance assumption is 8%.
                        CaseNumber = 1;
                        break;
                    default:
                        // Impedance assumption is 10%.
                        CaseNumber = 0;
                        break;
                }
            }
            else
            {
                // Impedance assumption is 10%.
                CaseNumber = 0;
            }

            UCL ucl = new UCL();

            return Convert.ToDouble(ucl.GetValues("impedance", CaseNumber), new CultureInfo("en-US"));
        }

        /// <summary>
        /// Matches Test Case to the breaker combination.
        /// </summary>
        /// <param name="CaseNumber">Test case number of the "Grouping Schemes" document.</param>
        /// <param name="UnitNumber">Unit's paralleling address.</param>
        /// <param name="Breaker">There are 3 breakers available per unit.
        /// A1 = Line breaker.
        /// A2 = Right tie breaker.
        /// A3 = Left tie breaker.</param>
        /// <returns>Returns boolean value determine if the breaker is open or close
        /// in the current Test Case.</returns>
        [ComVisible(false)]
        public bool BreakerCase(int CaseNumber, int UnitNumber, int Breaker)
        {
            try
            {
                // A1, A2 and A3 represented by only one bit.
                const int breakerBit = 1;
                string breakerSetting = "";

                // This function fails if the total unit number > 3
                Debug.WriteLine("Case #: {0}", CaseNumber);

                switch (CaseNumber)
                {
                    // 1, 3
                    case 1:
                    case 2:
                    case 3:
                    case 4:
                    case 9:
                    case 10:
                    case 11:
                    case 12:

                        breakerSetting = GetBreakers(UnitNumber, "102");
                        break;

                    // 2, 21
                    case 5:
                    case 6:
                    case 7:
                    case 8:
                    case 79:

                        breakerSetting = GetBreakers(UnitNumber, "000");
                        break;

                    // 4
                    case 13:
                    case 14:
                    case 15:
                    case 16:

                        breakerSetting = GetBreakers(UnitNumber, "210");
                        break;

                    // 5
                    case 17:
                    case 18:
                    case 19:
                    case 20:

                        breakerSetting = GetBreakers(UnitNumber, "021");
                        break;

                    // 6
                    case 21:
                    case 22:
                    case 23:
                    case 24:

                        breakerSetting = GetBreakers(UnitNumber, "111");
                        break;

                    // 7
                    case 25:
                    case 26:
                    case 27:
                    case 28:

                        breakerSetting = GetBreakers(UnitNumber, "222");
                        break;

                    // 8, 10
                    case 29:
                    case 30:
                    case 31:
                    case 32:
                    case 37:
                    case 38:
                    case 39:
                    case 40:

                        breakerSetting = GetBreakers(UnitNumber, "123");
                        break;

                    // 9,11
                    case 33:
                    case 34:
                    case 35:
                    case 36:
                    case 41:
                    case 42:
                    case 43:
                    case 44:

                        breakerSetting = GetBreakers(UnitNumber, "312");
                        break;

                    // 12
                    case 45:
                    case 46:
                    case 47:
                    case 48:

                        breakerSetting = GetBreakers(UnitNumber, "231");
                        break;

                    // 13
                    case 49:
                    case 50:
                    case 51:
                    case 52:

                        breakerSetting = GetBreakers(UnitNumber, "113");
                        break;

                    // 14
                    case 53:
                    case 54:
                    case 55:
                    case 56:

                        breakerSetting = GetBreakers(UnitNumber, "223");
                        break;

                    // 15
                    case 57:
                    case 58:
                    case 59:
                    case 60:

                        breakerSetting = GetBreakers(UnitNumber, "311");
                        break;

                    // 16
                    case 61:
                    case 62:
                    case 63:
                    case 64:

                        breakerSetting = GetBreakers(UnitNumber, "322");
                        break;

                    // 17
                    case 65:
                    case 66:
                    case 67:
                    case 68:

                        breakerSetting = GetBreakers(UnitNumber, "131");
                        break;

                    // 18
                    case 69:
                    case 70:
                    case 71:
                    case 72:

                        breakerSetting = GetBreakers(UnitNumber, "232");
                        break;

                    // 19
                    case 73:
                    case 74:
                    case 75:

                        breakerSetting = GetBreakers(UnitNumber, "333");
                        break;

                    // 20
                    case 76:
                    case 77:
                    case 78:

                        breakerSetting = GetBreakers(UnitNumber, "444");
                        break;

                    default:
                        breakerSetting = GetBreakers(UnitNumber, "000");
                        break;
                }

                bool result = Convert.ToBoolean(Convert.ToByte(breakerSetting.Substring(Breaker - 1, breakerBit), new CultureInfo("en-US")));
                
                return result;
            }
            catch (System.OverflowException ex)
            {

                throw new OverflowException("BreakerSetting byte conversion failed", ex);
            }
        }

        /// <summary>
        /// Provides interface to XML test file.
        /// </summary>
        /// <param name="UnitNumber">Unit's paralleling address.</param>
        /// <param name="BreakerCombination">This information provided by the "Grouping Schemes" document.</param>
        /// <returns>Returns string of breaker position specified in "Setup" file.</returns>
        [ComVisible(false)]
        private string GetBreakers(int UnitNumber, string BreakerCombination)
        {
            int position = Convert.ToInt16(BreakerCombination.Substring((UnitNumber - 1), 1), new CultureInfo("en-US"));
            
            UCL ucl = new UCL();

            return ucl.GetValues("breakers", position);
        }
    }
}
