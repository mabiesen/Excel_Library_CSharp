using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelLibrary
{
    public class ExcelWorksheet
    {
        public List<ExcelTable> worksheetTables = new List<ExcelTable>();

        public string worksheetName
        {
            get;
            set;
        }

        public string[,] worksheetData
        {
            get;
            set;
        }

        public void CreateTableFromData(bool wholeTable = false, int startx = 0, int starty = 0, int stopx = 0, int stopy = 0)
        {
            string[,] tableData;
            if (wholeTable)
            {
                tableData = this.worksheetData;
            }
            else
            {
                int colcount = stopx - startx + 1;
                int rowcount = stopy - starty + 1;
                tableData = new string[rowcount, colcount];

                for (int ctr1 = starty; ctr1 < rowcount; ctr1++)
                {
                    for (int ctr2 = startx; ctr2 < colcount; ctr2++)
                    {
                        tableData[ctr1, ctr2] = this.worksheetData[ctr1, ctr2];
                    }
                }
                tableData = this.worksheetData;
            }

            ExcelTable newTable = new ExcelTable(tableData);
            this.worksheetTables.Add(newTable);

        }
    }
}
