using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


using ExcelDna.Integration;
using ExcelDna.IntelliSense;

using ExcelDna.Integration.CustomUI;

using System.Runtime.InteropServices;

namespace ExcelGenericUDF
{
    [ComVisible(true)]
    class Ribbon : ExcelRibbon
    {
        static public void prova()
        {
            var test = new ClassiDiScambio.NumeraFogli.Sheet()
            {
                Cell = "",
                FromSheet = 2,
                StartingNumber = 1,
                ToSheet = 2,
                TotalSheet = 3
            };

            var lala = new ExcelGui.NumeraFogli(test);
            lala.ShowDialog();
        }
    }
}
