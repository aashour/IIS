using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.DirectoryServices;

#region FileInfo
/*--------------------------------------------
 * File         : IISProp.cs
 * Description  : Export MIME Types from IIS
 * Date         : 14/01/2009
 * Author       : César Afonso
 *-------------------------------------------*/
#endregion

namespace cAfonso
{
    /// <summary>
    /// Get current IIS Properties
    /// </summary>
    public static class IISProp
    {
        /// <summary>
        /// Get the list of registered MimeTypes
        /// </summary>
        /// <param name="IISAddress">The IIS address to query</param>
        /// <param name="PropertiesName">The name of the properties to get</param>
        /// <returns>A sorted list with all the extension and mime types that are registered.</returns>
        /// <remarks>No exception handling.</remarks>
        public static SortedList<string, string> Get(string IISAddress, string PropertiesName)
        {
            SortedList<string, string> MimeTypesList = new SortedList<string, string>();

            DirectoryEntry Path = new DirectoryEntry(IISAddress);
            PropertyValueCollection PropValues = Path.Properties[PropertiesName];

            IISOle.MimeMap MimeTypeObj;
            foreach (var item in PropValues)
            {
                // IISOle -> Add reference to Active DS IIS Namespace provider
                MimeTypeObj = (IISOle.MimeMap)item;
                MimeTypesList.Add(MimeTypeObj.Extension, MimeTypeObj.MimeType);
            }

            return MimeTypesList;
        }

    }
}
