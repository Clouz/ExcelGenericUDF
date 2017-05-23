using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
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

        private int _FromSheet;
        public int FromSheet
        {
            get { return this._FromSheet; }
            set
            {
                if (this._FromSheet != value)
                    this._FromSheet = value;
            }
        }

        private int _ToSheet;
        public int ToSheet
        {
            get { return _ToSheet; }
            set { _ToSheet = value; }
        }

        public int TotalSheet { get; set; }

        private int _StartingNumber;
        public int StartingNumber
        {
            get { return _StartingNumber; }
            set { _StartingNumber = value; }
        }

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
            TotalSheet = workbook.Worksheets.Count;
            ToSheet = TotalSheet;
            FromSheet = worksheets.Index;
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
                    Value = Convert.ToString((worksheets.Cells[Row, Column] as Excel.Range).Value)
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
                worksheets.Cells[Row, Column] = num++;
            }
        }

    }
}
