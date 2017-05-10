using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using ExcelDna.Integration;
using ExcelDna.IntelliSense;

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

        [ExcelFunction(Description = "Data una matrice con tipico, potenza carico e lunghezza del cavo restituisce la sezione la sezione")]
        public static string ClaudioCalcolaCavi(
                [ExcelArgument(Description = "Inserire una matrice dove nella prima colonna è indicato il tipico, nella seconda la potenza (kW) e nelle restanti righe di intestazione le varie lunghezze del cavo", Name = "Matrice Cavi")] object[,] matriceCavi,
                [ExcelArgument(Description = "Inserire il tipico desiderato già presente nella matrice", Name = "Tipico")] string tipico,
                [ExcelArgument(Description = "Inserire la potenza del carico in kW", Name = "Potenza")] Double potenza,
                [ExcelArgument(Description = "Inserire la lunghezza del carico in metri", Name = "Lunghezza")] Double lunghezza
            )
        {
            try
            {
                int righeMatrice = matriceCavi.GetLength(0);
                int colonneMatrice = matriceCavi.GetLength(1);

                for (int i = 1; i < righeMatrice; i++)
                {
                    if (tipico.ToLower() == matriceCavi[i,0].ToString().ToLower())
                    {
                        if (double.Parse(matriceCavi[i, 1].ToString()) >= potenza)
                        {
                            for (int ii = 2; ii < colonneMatrice; ii++)
                            {
                                if (double.Parse(matriceCavi[0, ii].ToString()) >= lunghezza)
                                {
                                    return $"[{matriceCavi[i, 1].ToString()}; {matriceCavi[0, ii].ToString()}] {matriceCavi[i, ii].ToString()}";
                                }
                            }
                        }
                    }
                }

                return "Nessuna corrispondenza trovata";
            }
            catch (Exception e)
            {
                return e.ToString();
            }
        }
    }
}
