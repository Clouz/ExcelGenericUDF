﻿using System;
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

using Excel = Microsoft.Office.Interop.Excel;

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
            var sheet = new ClassiDiScambio.NumeraFogli.Sheet((Excel.Application)ExcelDnaUtil.Application);

            var excelGui = new ExcelGui.NumeraFogli(sheet);
            excelGui.ShowDialog();
        }


    }
}
