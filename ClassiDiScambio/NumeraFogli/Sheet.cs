using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ClassiDiScambio.NumeraFogli
{
    public class Sheet
    {
        public string Cell { get; set; }
        public int FromSheet { get; set; }
        public int ToSheet { get; set; }
        public int TotalSheet { get; set; }
        public int StartingNumber { get; set; }
    }
}
