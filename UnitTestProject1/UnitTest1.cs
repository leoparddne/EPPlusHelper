using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;

namespace UnitTestProject1
{
    [TestClass]
    public class UnitTest1
    {
        [TestMethod]
        public void TestNPOI()
        {
            using (var tools = new ConsoleApp1.EPPlusHelper())
            {
                tools.Open(new System.IO.FileInfo("test.xls"));
                tools.SetData<Entity.Table>(new List<Entity.Table>() {
                    new Entity.Table { A = "f", B = "fff" },
                    new Entity.Table { A = "1111f", B = "111111fff" }
                });
                tools.Dispose();
            } 
        }
    }
}
