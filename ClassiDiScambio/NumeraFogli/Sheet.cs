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

    public class Sheet : INotifyDataErrorInfo
    {
        public string Cell { get; set; }

        private int _FromSheet;
        public int FromSheet
        {
            get { return this._FromSheet; }
            set
            {
                if (IsFromSheetValid(value) && this._FromSheet != value)
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


        private const string MinValueError = "Il numero deve essere maggiore di 0";
        private const string StringLengthError = "Il nome è più lungo di 15 caratteri";
        private const string AgeError = "L'età non può essere inferiore a 18";

        private Dictionary<string, List<string>> errors = new Dictionary<string, List<string>>();

        public bool IsFromSheetValid(int value)
        {
            bool isValid = true;
            Task.Run(() =>
            {
                if (value < 1)
                {
                    AddError("FromSheetError", MinValueError, false);
                    isValid = false;
                }
                else
                {
                    RemoveError("FromSheetError", MinValueError);
                }
            });

            return isValid;
        }

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

        public void AddError(string propertyName, string error, bool isWarning)
        {
            //se la proprietà, determinata con l'indice della collection, 
            //non contiene l' errore specificato, aggiunge o mette in prima 
            //posizione l'errore a seconda di isWarning e scatena 
            //l'evento ErrorsChanged 
            if (!errors.ContainsKey(propertyName))
                errors[propertyName] = new List<string>();

            if (!errors[propertyName].Contains(error))
            {
                if (isWarning)
                {
                    errors[propertyName].Add(error);
                }
                else
                {
                    errors[propertyName].Insert(0, error);
                }

                RaiseErrorsChanged(propertyName);
            }
        }

        // Rimuove l'errore specificato dalla collezione se presente 
        // e scatena l'evento ErrorsChanged 
        public void RemoveError(string propertyName, string error)
        {
            //se la collection di errori contiene il nome 
            //della proprietà e contiene l'errore specificato, 
            //lo rimuove dalla collection, quindi scatena 
            //l'evento ErrorsChanged 
            if (errors.ContainsKey(propertyName) && errors[propertyName].Contains(error))
            {
                errors[propertyName].Remove(error);
                if (errors[propertyName].Count == 0)
                    errors.Remove(propertyName);

                RaiseErrorsChanged(propertyName);
            }
        }

        public void RaiseErrorsChanged(string propertyName) {
            if (ErrorsChanged != null) {
                ErrorsChanged(this, new DataErrorsChangedEventArgs(propertyName));
            }
        }

        public event EventHandler<DataErrorsChangedEventArgs> ErrorsChanged;

        public delegate void ErrorsChangedEventHandler(object sender, DataErrorsChangedEventArgs e);

        public IEnumerable GetErrors(string propertyName) {
            if ((string.IsNullOrEmpty(propertyName) || !errors.ContainsKey(propertyName)))
                return null;

            return errors[propertyName];
        }

        public bool HasErrors { get { return errors.Any(); } }
    }
}
