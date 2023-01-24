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
        public void ListViewTestCreateInsert()
        {
            var path = "..//..//..//srcTest//listView2.docx";
            List<string> listS = new()
            {
                "1",
                "2",
                "3"
            };
            WordExchange exchangeClass = new(ExchangeOperation.Read, null);
            WrapperWord wrapper = new(path, exchangeClass, null);
            wrapper.Exchange();
            var d=exchangeClass.ExchangeValue;
            CollectionAssert.AreEqual(listS, exchangeClass.ExchangeValue.ToList());
        }
    }
}
