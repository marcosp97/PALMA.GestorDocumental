using System;
using System.Collections.Generic;
using System.Text;
using SP = Microsoft.SharePoint.Client;

namespace Palma.GestorDocumental.Repository.Entity
{
    public class FileBE
    {
        public SP.ListItem item { get; set; }
        public string nombreLista { get; set; }
    }
}
