using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


using ExcelDna.Integration;
using ExcelDna.IntelliSense;

using ExcelDna.Integration.CustomUI;

using System.Runtime.InteropServices;

using System.Runtime.InteropServices;
using ExcelDna.Integration.CustomUI;
using System.Windows.Forms;

namespace ExcelGenericUDF
{
    [ComVisible(true)]
    public class Ribbon : ExcelRibbon
    {

        public override string GetCustomUI(string RibbonID)
        {
            return @"
<customUI xmlns='http://schemas.microsoft.com/office/2006/01/customui'>
    <ribbon>
    <tabs>
        <tab id='tab1' label='My Tab'>
        <group id='group1' label='My Group'>
            <button id='button1' label='My Button' onAction='OnButtonPressed'/>
        </group >
        </tab>
    </tabs>
    </ribbon>
</customUI>";
        }


        public void OnButtonPressed(IRibbonControl control)
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
