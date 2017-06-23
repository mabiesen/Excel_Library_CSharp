using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelLibrary
{
    public class ExcelWorkbook
    {
        public List<ExcelWorksheet> worksheetList = new List<ExcelWorksheet>();

        public string workbookName
        {
            get;
            set;
        }

        public string workbookPath
        {
            get;
            set;
        }
    }
}
