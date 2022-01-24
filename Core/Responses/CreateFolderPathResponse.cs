using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Microsoft.Exchange.WebServices.Data
{
    /// <summary>
    /// Represents the response to create a folder hierarchy operation.
    /// </summary>
    public sealed class CreateFolderPathResponse : ServiceResponse
    {
        private Folder folder;

        internal CreateFolderPathResponse(Folder folder)
            : base()
        {
            this.folder = folder;
        }

        private Folder GetObjectInstance(ExchangeService service, string xmlElementName)
        {
            if (this.folder != null)
            {
                return this.folder;
            }
            else
            {
                return EwsUtilities.CreateEwsObjectFromXmlElementName<Folder>(service, xmlElementName);
            }
        }

        /// <summary>
        /// Reads response elements from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        internal override void ReadElementsFromXml(EwsServiceXmlReader reader)
        {
            base.ReadElementsFromXml(reader);

            List<Folder> folders = reader.ReadServiceObjectsCollectionFromXml<Folder>(
                XmlElementNames.Folders,
                this.GetObjectInstance,
                false,  /* clearPropertyBag */
                null,   /* requestedPropertySet */
                false); /* summaryPropertiesOnly */

            this.folder = folders[0];
        }

        /// <summary>
        /// Clears the change log of the created folder if the creation succeeded.
        /// </summary>
        internal override void Loaded()
        {
            if (this.Result == ServiceResult.Success)
            {
                this.folder.ClearChangeLog();
            }
        }
    }
}
