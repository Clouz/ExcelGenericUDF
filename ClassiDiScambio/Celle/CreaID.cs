using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Excel = Microsoft.Office.Interop.Excel;



namespace ClassiDiScambio.Celle
{
    public class CellID
    {
        public int Row { get; set; }
        public int Column { get; set; }


        private Excel.Application application;
        private Excel.Workbook workbook;

        public CellID(Excel.Application application)
        {
            this.application = application;
            this.workbook = application.ActiveWorkbook;
            Excel.Worksheet worksheet = workbook.ActiveSheet;
            
            Row = application.ActiveCell.Row;
            Column = application.ActiveCell.Column;

            Excel.Range range = worksheet.UsedRange.Columns[Column];

            object[,] colonna = range.get_Value(Excel.XlRangeValueDataType.xlRangeValueDefault);

            int conteggio = 0;

            for (int i = 1; i <= colonna.Length; i++)
            {
                if (colonna[i,1] is double)
                {
                    if (conteggio < (double)colonna[i, 1])
                    {
                        conteggio = int.Parse(colonna[i, 1].ToString());
                    }
                }
            }

            worksheet.Cells[Row, Column] = conteggio+1; 
        }
    }
}
