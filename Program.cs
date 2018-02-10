using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SGMLExtracter
{
    class Program
    {
        static void Main(string[] args)
        {
            SGMLHelper s = new SGMLHelper();
            s.StoreFolder = "c:\\Foo";

            string action = args[0].ToLower();
            if (action == "file")
                s.ReadSingleFile(args[1]);
            else if (action == "folder")
                s.ReadSingleFolder(args[1]);
            else if (action == "setupsp")
            {
                SharePointHelper sh = null;
                System.Net.NetworkCredential cred = null;
                if (args.Length > 2)
                {
                    cred = new System.Net.NetworkCredential(args[2], args[3]);
                    sh = new SharePointHelper(args[1], cred);
                }
                else
                    sh = new SharePointHelper(args[1]);

                sh.DeleteDocumentLibrary("Reuter");
                List list = sh.CreateDocumentLibrary("Reuter", "ReuterFiles");
                TermGroup group = sh.CreateGroup(sh.GetTermStore(), "MachineLearning");
                TermSet set = sh.CreateTermSet(group, "MachineLearning");
                sh.DeleteFieldIfExists("Topic");
                Microsoft.SharePoint.Client.Field taxField = sh.CreateTaxonomyField(sh.GetTermStore(), set, "Topic", "Topic", false, false);
                sh.AddTaxFieldToList(list, taxField, true);

            }
            else
                PrintInfo();

        }

        static void PrintInfo()
        {
            Console.WriteLine(@"Program file c:\Data\data.sgml");
            Console.WriteLine(@"Program folder c:\Data");
            Console.WriteLine(@"Program setupsp c:\Data UserName Password");
        }
    }
}
