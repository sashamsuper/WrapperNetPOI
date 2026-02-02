using Microsoft.Data.Analysis;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Text;

namespace WrapperNetPOI
{

    public class DataColumn
    {
        public string Name { set; get; }
        public int Number { set; get; }
        public Type Type { set; get; }

        public override string ToString()
        {
            return Name;
        }

        public DataColumn(string name, int columnNumber, Type columnType)
        {
            Name = name;
            Number = columnNumber;
            Type = columnType;
        }
    }

    public interface IHeader
    {
        int[] Rows { get; set; }
        DataColumn[] DataColumns { get; set; }
        void CreateHeaderType(Dictionary<int, Type> columns);
        void RenameDoubleHeaderColumn();
    }

    public interface IBorder
    {
        ISheet ActiveSheet { get; set; }
        int FirstRow { get; set; }
        int FirstColumn { get; set; }
        int LastRow { get; set; }
        int LastColumn { get; set; }
        bool IsCorrected { get; }
        int Row(int i);
        int Column(int i);
        void CorrectBorder(int? firstRow = null, int? firstColumn = null, int? lastRow = null, int? lastColumn = null);
    }

    public interface IDataFrameView: IExchange
    {
        IHeader DataHeader { set; get; }

        void ReadHeader();

        void CreateColumns();
        
    }
}
