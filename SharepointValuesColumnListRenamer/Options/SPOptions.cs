using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using CommandLine;

namespace SharepointValuesColumnListRenamer.Options
{
    class SPOptions
    {
        [Option('w', "website", Required = true, HelpText = "Sharepoint complete website Url to webservice list manager")]
        public string WebSite { get; set; }

        #region Credentials

        [Option('u', "username", Required = true, HelpText = "Sharepoint WebSite Username Credentials")]
        public string UserName { get; set; }

        [Option('p', "password", Required = true, HelpText = "Sharepoint WebSite Password Credentials")]
        public string Password { get; set; }

        [Option('d', "domain", Required = true, HelpText = "Sharepoint WebSite Domain Credentials")]
        public string Domain { get; set; }

        #endregion

        #region List

        [Option('l', "listname", Required = true, HelpText = "List Sharepoint Name")]
        public string ListName { get; set; }

        [Option('c', "column", Required = true, HelpText = "Column List Name That you want to update")]
        public string Column { get; set; }

        [Option('v', "value", Required = true, HelpText = "Value That you want")]
        public string Value { get; set; }

        #endregion

        [Option('s', "start", Required = false, HelpText = "Launch Update")]
        public Boolean Start { get; set; }

        [Option('m', "multi", Required = false, HelpText = "Multiple ColumnsValues")]
        public Boolean Multi { get; set; }
    }
}
