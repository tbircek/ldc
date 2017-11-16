using System;
using System.Runtime.InteropServices;

namespace P2PCalculations
{
    /// <summary>
    /// COM Interface for GeneralSettings.
    /// </summary>
    [ComVisible ( false )]
    [Guid ( "FA572494-6EDC-46B1-B228-84DD7A99395F" )]
    // [InterfaceType ( ComInterfaceType.InterfaceIsDual )]
    // Following line changed due to FxCop rules.
    // "COM source interfaces should be marked ComInterfaceType.InterfaceIsIDispatch. 
    // Visual Basic 6 clients cannot receive events with non-IDispatch 
    // interfaces."
    [InterfaceType ( ComInterfaceType.InterfaceIsIDispatch )]
    public interface IGeneralSettings
    {
        /// <summary>
        /// Total test numbers that fits into regular test patterns.
        /// There are 4 more cases that not include in this number, don't fit into these test patterns.
        /// In these cases Unit3 will have same CT, MVA and Impedance settings as other units.
        /// </summary>
        int TotalRegularTest { get; }

        /// <summary>
        /// Total test numbers to test for 3-Unit systems.
        /// </summary>
        int TotalTestNumbers { get; }

        /// <summary>
        /// Reads setting values from "Setup" file".
        /// </summary>
        /// <param name="itemName">Setting name in "Setup" file.</param>
        /// <param name="unitNumber">It is 0 for if a setting is common for all the units, 
        /// and it is unit's paralleling address if not same for the all units.</param>
        /// <returns></returns>
        double GetSettings ( String itemName, int unitNumber );

        /// <summary>
        /// MVA settings per Test Case is being run.
        /// </summary>
        /// <param name="caseNumber">Test case number of the "Grouping Schemes" document.</param>
        /// <param name="unitNumber">Unit's paralleling address.</param>
        /// <returns>Returns MVA setting value per "Setup" document.</returns>
        double MvaCase ( int caseNumber, int unitNumber );

        /// <summary>
        /// CT settings per Test Case is being run.
        /// </summary>
        /// <param name="caseNumber">Test case number of the "Grouping Schemes" document.</param>
        /// <param name="unitNumber">Unit's paralleling address.</param>
        /// <returns>Returns CT setting value per "Setup" document.</returns>
        double CTCase ( int caseNumber, int unitNumber );

        /// <summary>
        /// Impedance assumption per Test Case is being run.
        /// </summary>
        /// <param name="caseNumber">Test case number of the "Grouping Schemes" document.</param>
        /// <param name="unitNumber">Unit's paralleling address.</param>
        /// <returns>Returns Impedance assumption value per "Setup" document.</returns>
        double ImpedanceCase ( int caseNumber, int unitNumber );

        /// <summary>
        /// Matches Test Case to the breaker combination.
        /// </summary>
        /// <param name="caseNumber">Test case number of the "Grouping Schemes" document.</param>
        /// <param name="unitNumber">Unit's paralleling address.</param>
        /// <param name="breaker">There are 3 breakers available per unit.
        /// A1 = Line breaker.
        /// A2 = Right tie breaker.
        /// A3 = Left tie breaker.</param>
        /// <returns>Returns boolean value determine if the breaker is open or close
        /// in the current Test Case.</returns>
        bool BreakerCase ( int caseNumber, int unitNumber, int breaker );
    }
}
