using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;
using WrapperNetPOI;

namespace MsTestWrapper
{
    [TestClass]
    public class UnitTestWord
    {
        [TestMethod]
        public void ReadCellValueTest()
        {
            var path = "..//..//..//srcTest//listView2.docx";
            TableView exchangeClass = new(ExchangeOperation.Read);
            WrapperWord wrapper = new(path, exchangeClass);
            wrapper.Exchange();
            var d=exchangeClass.ExchangeValue;
            Assert.AreEqual(36, exchangeClass.ExchangeValue.ToList().Count);
        }

        [TestMethod]
        public void ReadTableValueTest()
        {
            var path = "..//..//..//srcTest//listView2.docx";
            List<string> listS = new()
            {
                "1",
                "2",
                "3"
            };
            TableView exchangeClass = new(ExchangeOperation.Read, null);
            WrapperWord wrapper = new(path, exchangeClass, null);
            wrapper.Exchange();
            var d = exchangeClass.ExchangeValue;

            CollectionAssert.AreEqual(listS, exchangeClass.ExchangeValue.ToList());
        }
    }
}
