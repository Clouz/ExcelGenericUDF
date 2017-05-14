using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using ExcelDna.Integration;
using ExcelDna.IntelliSense;

using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelGenericUDF.Scheda
{
    public class Nome
    {

        [ExcelFunction(Description = "Dato il numero della scheda ne restituisce il nome", Name = Udf.nome + ".Scheda.Nome")]
        public static object NomeScheda(
        [ExcelArgument(Name = "Numero Scheda", Description = "Inserire il numero della scheda")] int i)
        {
            try
            {
                Excel.Application application = (Excel.Application)ExcelDnaUtil.Application;
                Excel.Workbook workbook = application.ActiveWorkbook;
                Excel.Worksheet worksheets = workbook.Worksheets[i];

                return worksheets.Name;
            }
            catch (Exception e)
            {
                return e.ToString();
            }
        }

    }
}
