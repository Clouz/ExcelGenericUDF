using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using ExcelDna.Integration;
using ExcelDna.IntelliSense;

using ExcelDna.Integration.CustomUI;

using Excel = Microsoft.Office.Interop.Excel;

using ClassiDiScambio.NumeraFogli;
using ExcelGui;

namespace ExcelGenericUDF
{
    public class Udf : IExcelAddIn
    {
        public void AutoOpen()
        {
            IntelliSenseServer.Install();
        }

        public void AutoClose()
        {
            IntelliSenseServer.Uninstall();
        }

        public const string nome = "Claudio";

    }
}
