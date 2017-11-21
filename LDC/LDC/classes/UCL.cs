using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.IO;
using System.Reflection;
using System.Xml.Linq;

namespace P2PCalculations
{
    /// <summary>
    /// Interface to access "Setup" file.
    /// File name specified in the project resources.
    /// </summary>
    [Guid ( "BBC4E422-38C6-420D-BD37-01B40D9E8E33" )]
    [ClassInterface ( ClassInterfaceType.None )]
    [ComSourceInterfaces ( typeof ( IUCL ) )]
    [ProgId ( "P2PCalculations.UCL" )]
    [ComVisible ( true )]
    public class UCL : IUCL
    {
        /// <summary>
        /// Access function for "Setup" file.
        /// </summary>
        /// <param name="CaseName">Xml Element to find in the file.</param>
        /// <param name="CaseNumber">Value of the Xml Element in the xml file.</param>
        /// <returns>Returns value of the item requested from the xml file.</returns>
        [ComVisible ( false )]
        public string GetValues ( String CaseName, int CaseNumber )
        {

            return GetValues ( CaseName, CaseNumber, "value" );
        }

        /// <summary>
        /// Access function for "Setup" file.
        /// </summary>
        /// <param name="CaseName">Xml Element to find in the file.</param>
        /// <param name="CaseNumber">Value of the Xml Element in the xml file.</param>
        /// <param name="ElementName">Specify element value type. Three possible choices. Value, Max and Min.</param>
        /// <returns>Returns value of the item requested from the xml file.</returns>
        [ComVisible ( true )]
        public string GetValues ( String CaseName, int CaseNumber, String ElementName )
        {

            try
            {
                string dllFolderPath = Path.GetDirectoryName ( Assembly.GetExecutingAssembly ( ).Location );
                string setupFileName = Properties.Resources.setupFileName;
                string settingFilePath = Path.Combine ( dllFolderPath, setupFileName );

                XElement root = XElement.Load ( settingFilePath );

                IEnumerable<XElement> values =
                    from el in root.Elements ( "setting" )
                    where
                    ( string ) el.Element ( "name" ) == CaseName
                    select el;

                // Element name has 3 choices: Value, Max, and Min
                string result = values.ElementAt ( CaseNumber ).Element ( ElementName ).Value;

                return result;
            }
            catch ( Exception )
            {
                // Item is not in the "Setup" file.
                return "0";
            }            
        }
    }
}
