using NPOI.OpenXmlFormats.Wordprocessing;
using System.Collections;
using System.Diagnostics;
using WrapperNetPOI;
using WrapperNetPOI.Word;
using Microsoft.Data.Analysis;

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
                new string[]{"133", "244", "7555" },
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
        public void SimpleReadTableValueTest()
        {
            const string path = "..//..//..//srcTest//listView2.docx";
            List<string[]> listS = new()
            {
                new string[]{"133", "244", "7555" },
                new string[]{"3", "4", "8"}
            };
            Simple.GetFromWord(out List<TableValue> sample, path);
            CollectionAssert.AreEqual(sample.First().Value, listS, new ListComparerClass());
        }

        [TestMethod]
        public void SimpleReadTableValueDataFrameTest()
        {
            const string path = "..//..//..//srcTest//listView2.docx";
            List<string[]> listS = new()
            {
                new string[]{"133", "244", "7555" },
                new string[]{"3", "4", "8"}
            };
            Simple.GetFromWord(out List<DataFrame> sample, path);
            var value=sample.First().Rows.ToList().Select(x=>x.ToArray()).ToList();
            CollectionAssert.AreEqual(value, listS, new ListComparerClass());
        }

        [TestMethod]
        public void ReadWord2003()
        {
            const string path = "..//..//..//srcTest//word2003.doc";
            ParagraphView exchangeClass = new(ExchangeOperation.Read, null);
            WrapperWord wrapper = new(path, exchangeClass, null);
            wrapper.Exchange();
            var value = exchangeClass.Document.Paragraphs[0].ToString();
            Assert.AreEqual("gffgn1sdfsdfsdfsdfàâðïààïðàïðàïð\r", value);
        }

        [TestMethod]
        public void ReadParagraphValueTestInCell()
        {
            const string path = "..//..//..//srcTest//listView2.docx";
            TableView exchangeClass = new(ExchangeOperation.Read, null);
            WrapperWord wrapper = new(path, exchangeClass, null);
            wrapper.Exchange();
            List<string[]> right = new()
            {   new[]{"133", "244", "7555"},
                new[]{"3",   "4",   "8"}
            };
            CollectionAssert.AreEqual(right, exchangeClass.ExchangeValue.First().Value, new ListComparerClass());
        }


        [TestMethod]
        public void ReadParagraphValueTestInCellSimple()
        {
            const string path = "..//..//..//srcTest//listView2.docx";
            Simple.GetFromWord(out List<TableValue> listValue, path);
            List<string[]> right = new()
            {   new[]{"133", "244", "7555"},
                new[]{"3",   "4",   "8"}
            };
            CollectionAssert.AreEqual(right, listValue.First().Value, new ListComparerClass());
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
                    var move = true;
                    while (move)
                    {
                        var enumeratorXMoveNext = enumeratorX.MoveNext();
                        var enumeratorYMoveNext = enumeratorY.MoveNext();
                        if (enumeratorXMoveNext && enumeratorYMoveNext)
                        {
                            move = true;
                        }
                        else if (enumeratorXMoveNext==false && enumeratorYMoveNext==false)
                        {
                            move = false;
                            return 0;
                        }
                        else if (enumeratorXMoveNext ^ enumeratorYMoveNext)
                        {
                            move = false;
                            return -1;
                        }
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