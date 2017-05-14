using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using ExcelDna.Integration;
using ExcelDna.IntelliSense;

using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelGenericUDF
{
    public class Udf : IExcelAddIn
    {
        public void AutoOpen()
        {
            IntelliSenseServer.Register();
        }

        public void AutoClose()
        {

        }

        public const string nome = "Claudio";
    }
}
