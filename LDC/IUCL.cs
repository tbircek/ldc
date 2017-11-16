using System;
using System.Runtime.InteropServices;

namespace P2PCalculations
{
    /// <summary>
    /// Provides access to "P2P Setup File.xml".
    /// </summary>
    [ComVisible ( true )]
    [Guid ( "E24CDB75-4ABF-4E6D-BC56-C906647A42D6" )]
    // [InterfaceType ( ComInterfaceType.InterfaceIsDual )]
    // Following line changed due to FxCop rules.
    // "COM source interfaces should be marked ComInterfaceType.InterfaceIsIDispatch. 
    // Visual Basic 6 clients cannot receive events with non-IDispatch 
    // interfaces."
    [InterfaceType ( ComInterfaceType.InterfaceIsIDispatch )]
    public interface IUCL
    {

        /// <summary>
        /// Access function for "Setup" file.
        /// </summary>
        /// <param name="CaseName">Xml Element to find in the file.</param>
        /// <param name="CaseNumber">Value of the Xml Element in the xml file.</param>
        /// <param name="ElementName">Specify element value type. Three possible choices. Value, Max and Min.</param>
        /// <returns>Returns value of the item requested from the xml file.</returns>
        string GetValues ( String CaseName, int CaseNumber, String ElementName );
    }
}
