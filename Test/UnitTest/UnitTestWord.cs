using System.Collections;
using WrapperNetPOI;
using WrapperNetPOI.Word;

namespace MsTestWrapper
{
    [TestClass]
    public class UnitTestWord
    {
        [TestMethod]
        public void ReadTableValueTest()
        {
            const string path = "..//..//..//srcTest//listView2.docx";
            List<string[]> listS = new()
            {
                new string[]{"1", "2", "7" },
                new string[]{"3", "4", "8"}
            };
            List<TableValue> sample = new();
            var tableValue = new TableValue(listS, 0, 0);
            sample.Add(tableValue);
            TableView exchangeClass = new(ExchangeOperation.Read, null);
            WrapperWord wrapper = new(path, exchangeClass, null);
            wrapper.Exchange();
            CollectionAssert.AreEqual(sample.ToList(), exchangeClass.ExchangeValue.ToList(), new ListComparerClass());
        }

        [TestMethod]
        public void ReadParagraphValueTest()
        {
            const string path = "..//..//..//srcTest//listView2.docx";
            ParagraphView exchangeClass = new(ExchangeOperation.Read, null);
            WrapperWord wrapper = new(path, exchangeClass, null);
            wrapper.Exchange();
            //CollectionAssert.AreEqual(sample.ToList(), exchangeClass.ExchangeValue.ToList(), new ListComparerClass());
        }

        public class ListComparerClass : IComparer
        {
            // Call CaseInsensitiveComparer.Compare with the parameters reversed.
            public int Compare(object? x, object? y)
            {
                if (x is IEnumerable _x && y is IEnumerable _y)
                {
                    IEnumerator enumeratorX = _x.GetEnumerator();
                    IEnumerator enumeratorY = _y.GetEnumerator();
                    while (enumeratorX.MoveNext() && enumeratorY.MoveNext())
                    {
                        if (new ListComparerClass().Compare(enumeratorX.Current, enumeratorY.Current) != 0)
                        {
                            return -1;
                        }
                    }
                    return 0;
                }
                else if (x is TableValue _xTable && y is TableValue _yTable)
                {
                    if (_xTable.tableNumber == _yTable.tableNumber
                        &&
                        _xTable.level == _yTable.level)
                    {
                        if (new ListComparerClass().Compare(_xTable.Value, _yTable.Value) != 0)
                        {
                            return -1;
                        }
                        else
                        {
                            return 0;
                        }
                    }
                    else
                    {
                        return -1;
                    }
                }
                else
                {
                    if (x == null || y == null)
                    {
                        return 0;
                    }
                    else
                    {
                        return x.Equals(y) ? 0 : -1;
                    }
                }
            }
        }
    }
}