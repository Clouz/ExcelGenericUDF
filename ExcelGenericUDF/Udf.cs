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

        [ExcelFunction(Description = "Data una matrice con tipico, potenza e lunghezza del cavo restituisce la sezione", Name ="Claudio.CalcolaCavi")]
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

                if (tipico == "") return "";

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
                                    return $"{matriceCavi[i, ii].ToString()} [{matriceCavi[i, 1].ToString()}; {matriceCavi[0, ii].ToString()}]";
                                }
                            }
                        }
                    }
                }

                return "NO MATCH FOUND";
            }
            catch (Exception e)
            {
                return e.ToString();
            }
        }

        [ExcelFunction(Description = "Dato il numero della scheda ne restituisce il nome",Name = "Claudio.Scheda.Nome")]
        public static object ClaudioNomeScheda(
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

        [ExcelFunction(Description = "Dato il numero della scheda ne modifica il nome", Name ="Claudio.Scheda.Rinomia")]
        public static object ClaudioRinomicaScheda(
            [ExcelArgument(Name ="Numero Scheda", Description ="Inserire il numero della scheda")] int i, 
            [ExcelArgument(Name ="Nome", Description ="Inserire il nuovo nome della scheda")] string nome)
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
        [ExcelFunction(Name ="Claudio.Stringa.Spazia", Description ="Data una stringa ed un pattern ne spazia il contenuto")]
        public static string ClaudioSepara(
            [ExcelArgument(Name ="Stringa Iniziale", Description ="")] string nome,
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
