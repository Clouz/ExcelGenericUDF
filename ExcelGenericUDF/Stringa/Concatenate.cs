using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using ExcelDna.Integration;
using ExcelDna.IntelliSense;

namespace ExcelGenericUDF.Stringa
{
    public class Concatenate
    {
        [ExcelFunction(Name = Udf.nome + ".Stringa.ConcatenateMultiple", Description = "Concatenate adiacent cells")]
        public static string ConcatenateMultiple(
        [ExcelArgument(Name = "Concatenate Range", Description = "Select multiple cell to concanenate")]  object[] Range,
        [ExcelArgument(Name = "Separator", Description = "Separator String")] string Separator,
        [ExcelArgument(Name = "Include Empty", Description = "if true show empty cells")] bool Empty)
        {
            try
            {
                string s = "";

                foreach (var cell in Range)
                {
                    if (cell is ExcelEmpty && Empty)
                        s += "" + Separator;
                    else if (cell is ExcelEmpty && Empty == false) {

                    }
                    else
                        s += cell + Separator;
                }

                return s.Substring(0, s.Length - Separator.Length);
            }
            catch (Exception e)
            {
                return e.ToString();
            }

        }
    }
}
