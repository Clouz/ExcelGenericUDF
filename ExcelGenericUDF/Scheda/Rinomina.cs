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
    public class Rinomina 
    {
        [ExcelFunction(Description = "Dato il numero della scheda ne modifica il nome", Name = Udf.nome + ".Scheda.Rinomia")]
        public static object RinominaScheda(
        [ExcelArgument(Name = "Numero Scheda", Description = "Inserire il numero della scheda")] int i,
        [ExcelArgument(Name = "Nome", Description = "Inserire il nuovo nome della scheda")] string nome)
        {
            try
            {
                Excel.Application application = (Excel.Application)ExcelDnaUtil.Application;
                Excel.Workbook workbook = application.ActiveWorkbook;
                Excel.Worksheet worksheets = workbook.Worksheets[i];

                worksheets.Name = nome;

                return worksheets.Name;
            }
            catch (Exception e)
            {
                return e.ToString();
            }
        }

    }
}
