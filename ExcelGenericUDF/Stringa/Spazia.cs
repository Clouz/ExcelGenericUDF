using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using ExcelDna.Integration;
using ExcelDna.IntelliSense;

namespace ExcelGenericUDF.Stringa
{
    public class Spazia
    {
        [ExcelFunction(Name = Udf.nome + ".Stringa.Spazia", Description = "Data una stringa ed un pattern ne spazia il contenuto")]
        public static string spazia(
        [ExcelArgument(Name = "Stringa Iniziale", Description = "")] string nome,
        [ExcelArgument(Name = "Pattern", Description = "Indica con quanti caratteri separatori riempiere la sottostringa. es. 10-10-5")] string Pattern,
        [ExcelArgument(Name = "Carattere Separatore", Description = "Indica con quale carattere la stringa viene separata")] string Separatore,
        [ExcelArgument(Name = "Carattere di Riempimento", Description = "Indica con quale carattere riempire la stringa")] string CarattereRiempimento)
        {
            try
            {
                string[] PatternDiviso = Pattern.Split(char.Parse(Separatore));
                var lista = nome.Split(char.Parse(Separatore));

                string nuovaStringa = "";

                for (int i = 0; i < lista.Length; i++)
                {
                    int quantitaSpazi;
                    try
                    {
                        if (int.Parse(PatternDiviso[i]) == 0)
                        {
                            quantitaSpazi = 0;
                        }
                        else
                        {
                            quantitaSpazi = int.Parse(PatternDiviso[i]) - lista[i].Length;
                        }
                    }
                    catch (Exception)
                    {
                        quantitaSpazi = 0;
                    }

                    nuovaStringa = nuovaStringa + new string(char.Parse(CarattereRiempimento), quantitaSpazi) + lista[i] + Separatore;
                }
                nuovaStringa = nuovaStringa.Substring(0, nuovaStringa.Length - 1);

                return nuovaStringa;
            }
            catch (Exception e)
            {
                return $"Stringa non corrispondente al pattern: {e.ToString()}";
            }
        }
    }
}
