using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices; 
using Microsoft.SharePoint.Administration;

namespace MarabouStork.Sharepoint.LinkChecker
{
    [GuidAttribute("C0DB8DDC-B9B0-4E45-898D-05559BA3E749")]
    public class LinkCheckerPersistedSettings: SPPersistedObject
    {
        [Persisted]
        public string DocLibraries;

        [Persisted]
        public string FieldsToCheck;

        [Persisted]
        public bool UnpublishInvalidDocs;

        public LinkCheckerPersistedSettings() { }

        public LinkCheckerPersistedSettings(string name, SPPersistedObject parent)
            : base(name, parent) { }
    }
}
