using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Excel = Microsoft.Office.Interop.Excel;



namespace ClassiDiScambio.NumeraFogli
{
    public class ContenutoFogli
    {
        public int Id { get; set; }
        public string Description { get; set; }
        public string Value { get; set; }
    }

    public class Sheet
    {
        public string Cell { get; set; }
        public int FromSheet { get; set; }
        public int ToSheet { get; set; }
        public int TotalSheet { get; set; }
        public int StartingNumber { get; set; }

        public int Row { get; set; }
        public int Column { get; set; }

        public List<ContenutoFogli> contenuto { get; set; } = new List<ContenutoFogli>();

        private Excel.Application application;
        private Excel.Workbook workbook;

        public Sheet(Excel.Application application)
        {
            this.application = application;
            this.workbook = application.ActiveWorkbook;
            Excel.Worksheet worksheets = workbook.ActiveSheet;

            Cell = application.ActiveCell.Address;
            Row = application.ActiveCell.Row;
            Column = application.ActiveCell.Column;
            FromSheet = worksheets.Index;
            TotalSheet = workbook.Worksheets.Count;
            ToSheet = TotalSheet;
            StartingNumber = FromSheet;

            ReloadList();
        }

        public void ReloadList()
        {
            contenuto.Clear();

            for (int i = FromSheet; i <= ToSheet; i++)
            {
                Excel.Worksheet worksheets = workbook.Worksheets[i];
                contenuto.Add(new ContenutoFogli()
                {
                    Id = worksheets.Index,
                    Description = worksheets.Name,
                    Value = (worksheets.Cells[Row,Column] as Excel.Range).Value
                });
            }
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
