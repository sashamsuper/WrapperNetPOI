using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;
using WrapperNetPOI;

namespace MsTestWrapper
{
    internal class UnitTestWord
    {
        [TestMethod]
        public void ListViewTestCreateInsert()
        {
            var path = "..//..//..//srcTest//listView.docx";
            if (File.Exists(path))
            {
                File.Delete(path);
            }

            List<string> listS = new()
            {
                "1",
                "2",
                "3"
            };
            WordExchange exchangeClass = new(ExchangeOperation.Read, null);
            WrapperWord wrapper = new(path, exchangeClass, null);
            wrapper.Exchange();
            List<string> listGet = new();
            exchangeClass = new(ExchangeOperation.Read, null);
            wrapper = new(path, exchangeClass, null);
            wrapper.Exchange();
            CollectionAssert.AreEqual(listS, exchangeClass.ExchangeValue.ToList());
        }
    }
}
