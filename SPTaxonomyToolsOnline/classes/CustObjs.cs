using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace BandR
{

    public class SimpleExportObj
    {
        public string i { get; set; }
        public string id { get; set; }
        public string parentId { get; set; }
        public string pathOfTerm { get; set; }
        public string name { get; set; }
        public string description { get; set; }
        public string isAvailableForTagging { get; set; }
        public string isReused { get; set; }
        public string isSourceTerm { get; set; }
        public string labels { get; set; }

        public int level
        {
            get
            {
                return pathOfTerm.ToCharArray().Count(x => x == ';');
            }
        }
    }

    public class SimpleImportObj
    {
        public Guid termId { get; set; }
        public string termName { get; set; }
        public string descr { get; set; }
        public bool isAvailForTagging { get; set; }
        public List<string> labels { get; set; }
    }

    public class TermObj
    {
        public Guid termId { get; set; }
        public string termName { get; set; }
        public List<string> labels { get; set; }

        public TermObj()
        {
            termId = Guid.NewGuid();
            termName = "";
            labels = new List<string>();
        }
    }

    public class TermObjAdv
    {
        public Guid id { get; set; }
        public string termName { get; set; }
        public string path { get; set; } // every path is unique, otherwise would not be added
        public List<string> labels { get; set; }
        public int level { get; set; }
        public bool isreused { get; set; }
        public bool reusebranch { get; set; }

        public TermObjAdv()
        {
            id = Guid.NewGuid();
            termName = "";
            path = "";
            labels = new List<string>();
            level = 0;
            isreused = false;
            reusebranch = false;
        }
    }

    public class CustomRegistrySettings
    {
        public string siteUrl { get; set; }

        public string userName { get; set; }
        public string passWord { get; set; }
        public string domain { get; set; }
        public string isSPOnline { get; set; }

        public string termStore { get; set; }
        public string termGroup { get; set; }
        public string termSet { get; set; }
        public string termStoreID { get; set; }
        public string termGroupID { get; set; }
        public string termSetID { get; set; }

        public string exportFilePath { get; set; }

        public string importSourceFilePath { get; set; }
        public string importSeparator { get; set; }
        public string importDbConnString { get; set; }
        public string importSelectStmt { get; set; }

        public string updateTermsSourceFilePath { get; set; }

    }

}
