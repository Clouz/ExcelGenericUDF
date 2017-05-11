using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using ExcelGenericUDF;

namespace UnitTestProject1
{
    [TestClass]
    public class UdfTest
    {
        [TestMethod]
        public void TestMethod1()
        {

            int expected = 3;
            int actual = 3;

            Assert.AreEqual(expected, actual, 0, "Risultato errato");
        }
    }
}
