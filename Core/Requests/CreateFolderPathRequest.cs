using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Microsoft.Exchange.WebServices.Data
{
    internal sealed class CreateFolderPathRequest : CreateRequest<Folder, CreateFolderPathResponse>
    {
        public CreateFolderPathRequest(ExchangeService service, ServiceErrorHandling errorHandlingMode) 
            : base(service, errorHandlingMode)
        {
        }

        internal override void Validate()
        {
            base.Validate();
            EwsUtilities.ValidateParam(this.Folders, "RelativeFolderPath");

            // Validate each folder.
            foreach (Folder folder in this.Folders)
            {
                folder.Validate();
            }
        }

        internal override string GetXmlElementName()
        {
            return XmlElementNames.CreateFolderPath;
        }

        internal override string GetResponseXmlElementName()
        {
            return XmlElementNames.CreateFolderPathResponse;
        }

        internal override ExchangeVersion GetMinimumRequiredServerVersion()
        {
            return ExchangeVersion.Exchange2013;
        }

        internal override CreateFolderPathResponse CreateServiceResponse(ExchangeService service, int responseIndex)
        {
            return new CreateFolderPathResponse((Folder)EwsUtilities.GetEnumeratedObjectAt(this.Folders, responseIndex));
        }

        internal override string GetResponseMessageXmlElementName()
        {
            return XmlElementNames.CreateFolderPathResponseMessage;
        }

        internal override string GetParentFolderXmlElementName()
        {
            return XmlElementNames.ParentFolderId;
        }

        internal override string GetObjectCollectionXmlElementName()
        {
            return XmlElementNames.RelativeFolderPath;
        }

        public IEnumerable<Folder> Folders
        {
            get { return this.Objects; }
            set { this.Objects = value; }
        }
    }
}
