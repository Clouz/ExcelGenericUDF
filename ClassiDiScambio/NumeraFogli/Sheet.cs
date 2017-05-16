using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Excel = Microsoft.Office.Interop.Excel;



namespace ClassiDiScambio.NumeraFogli
{
    public class Sheet
    {
        public string Cell { get; set; }
        public int FromSheet { get; set; }
        public int ToSheet { get; set; }
        public int TotalSheet { get; set; }
        public int StartingNumber { get; set; }

        public int Row { get; set; }
        public int Column { get; set; }

        private Excel.Application application;

        public Sheet(Excel.Application application)
        {
            this.application = application;
            Excel.Workbook workbook = application.ActiveWorkbook;
            Excel.Worksheet worksheets = workbook.ActiveSheet;

            Cell = application.ActiveCell.Address;
            Row = application.ActiveCell.Row;
            Column = application.ActiveCell.Column;
            FromSheet = worksheets.Index;
            TotalSheet = workbook.Worksheets.Count;
            ToSheet = TotalSheet;
            StartingNumber = FromSheet;
        }

        public void Write()
        {
            Excel.Workbook workbook = application.ActiveWorkbook;
            int num = StartingNumber;
            for (int i = FromSheet; i <= ToSheet; i++)
            {
                Excel.Worksheet worksheets = workbook.Worksheets[i];
                worksheets.Cells[Row,Column] = num++;
            }

        }
    }
}
