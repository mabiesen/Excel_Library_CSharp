using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelLibrary
{
    public class ExcelTable
    {
        private string[,] rawData;
        public DataTable dataTable = new DataTable();

        public ExcelTable(string[,] tableData)
        {
            this.rawData = tableData;
            SetDataTable();
        }

        public void SetDataTable()
        {
            int rw = this.rawData.GetLength(0);
            int col = this.rawData.GetLength(1);

            DataColumn datacolumn;
            DataRow datarow;

            //Set the column headers from first row
            for (int ctr = 0; ctr < col; ctr++)
            {
                datacolumn = new DataColumn();
                datacolumn.DataType = Type.GetType("System.String");
                datacolumn.ColumnName = this.rawData[0, ctr];
                this.dataTable.Columns.Add(datacolumn);
            }

            //Set the rows
            for (int ctr1 = 1; ctr1 < rw; ctr1++)
            {
                datarow = this.dataTable.NewRow();
                for (int ctr2 = 0; ctr2 < col; ctr2++)
                {
                    datarow[this.rawData[0, ctr2]] = this.rawData[ctr1, ctr2];
                }
                this.dataTable.Rows.Add(datarow);
            }
        }
    }
}
