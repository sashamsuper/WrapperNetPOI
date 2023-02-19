using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.Text;
using Mapster;

using NPOI.HSSF.Record;
using NPOI.XWPF.UserModel;
using System.Runtime.CompilerServices;

[assembly: InternalsVisibleTo("MsTestWrapper")]
namespace WrapperNetPOI
{
    public class WrapperCell
    {
        NPOI.SS.UserModel.ICell Cell { set; get; }
        public CellType CellType { set; get; }

        public WrapperCell(NPOI.SS.UserModel.ICell cell)
        {
            Cell = cell;
            CellType = cell.CellType;
        }

        public CellType CachedFormulaResultType 
        {
            get
            {
                try 
                {
                    return Cell.CachedFormulaResultType;
                }
                catch (InvalidOperationException ex)
                {

                    return default;
                }
            }
        }

        public double NumericCellValue
        {
            get
            {
                double value;

                try
                {
                    value = Cell.NumericCellValue;
                }
                catch (InvalidOperationException ex)
                {
                    value = 0;
                }
                catch (FormatException ex)
                {
                    value = 0;
                }

                return value;
            }
        }

        public string StringCellValue
        {
            get
            {
                string value;
                value = Cell.StringCellValue;
                return value;
            }
        }

        public DateTime DateCellValue
        {
            get
            {
                DateTime value;

                try
                {
                    value = Cell.DateCellValue;
                }
                catch (InvalidOperationException ex)
                {
                    value = default;
                }
                catch (FormatException ex)
                {
                    value = default;
                }

                return value;
            }
        }



    }

    public class ConvertType
    {

        CultureInfo ThisCultureInfo { get; set; } = CultureInfo.CurrentCulture;
        NumberStyles ThisNumberStyle { get; set; } = NumberStyles.Number;
        DateTimeStyles ThisDateTimeStyle { get; set; } = DateTimeStyles.AssumeUniversal;
        TypeAdapterConfig Config { get; set; }


        public ConvertType()
        {
            CreateMapster();
        }

        private void CreateMapster()
        {
            Config = new TypeAdapterConfig();

            //Config.ForType<WrapperCell, string>()
            //.Map(dest => dest,      
            //src => GetValueString(src));

            Config.ForType<WrapperCell, double>()
            .Map(dest => dest,      
            src => GetValueDouble(src));

            Config.ForType<WrapperCell, DateTime>()
            .Map(dest => dest,      
            src => GetValueDateTime(src));

        }

        public T GetValue<T>(WrapperCell cell) 
        {
            if (typeof(T) ==
                typeof(string)
                ||
                typeof(T) ==
                typeof(double)
                ||
                typeof(T) ==
                typeof(double))
            {
                return cell.Adapt<T>(Config);
            }
            else
            {
                throw new NotImplementedException("Do not have handler");
            }
        }


        protected internal DateTime GetValueDateTime(WrapperCell cell) => cell switch
        {
            {
                CellType: var cellType,
                StringCellValue: var stringCellValue,
            }
            when cellType == CellType.String => GetDateTime(stringCellValue),
            {
                CellType: var cellType,
                NumericCellValue: var numericCellValue,
            }
            when cellType == CellType.Numeric => GetDateTime(numericCellValue),
            {
                CellType: var cellType,
                StringCellValue: var stringCellValue,
                CachedFormulaResultType: var cachedFormulaResultType
            }
            when cellType == CellType.Formula &&
            cachedFormulaResultType == CellType.String
            => GetDateTime(stringCellValue),
            {
                CellType: var cellType,
                NumericCellValue: var numericCellValue,
                CachedFormulaResultType: var cachedFormulaResultType
            }
            when cellType == CellType.Formula &&
            cachedFormulaResultType == CellType.Numeric
            => GetDateTime(numericCellValue),
            _
            => cell.DateCellValue
        };

        protected internal DateTime GetDateTime(string value)
        {
            DateTime.TryParse(value, ThisCultureInfo,ThisDateTimeStyle, out var doubleValue);
            return doubleValue;
        }

        public DateTime GetDateTime(double value)
        {
            return Convert.ToDateTime(value);
        }


        


        protected internal string GetValueString(WrapperCell cell) => cell switch
        {
            {
                CellType: var cellType,
                StringCellValue: var stringCellValue,
            }
            when cellType == CellType.String => stringCellValue,
            {
                CellType: var cellType,
                NumericCellValue: var numericCellValue,
            }
            when cellType == CellType.Numeric => numericCellValue.ToString(),
            {
                CellType: var cellType,
                StringCellValue: var stringCellValue,
                CachedFormulaResultType: var cachedFormulaResultType
            }
            when cellType == CellType.Formula &&
            cachedFormulaResultType == CellType.String
            => stringCellValue,
            {
                CellType: var cellType,
                NumericCellValue: var numericCellValue,
                CachedFormulaResultType: var cachedFormulaResultType
            }
            when cellType == CellType.Formula &&
            cachedFormulaResultType == CellType.Numeric
            => numericCellValue.ToString(),
            _
            => cell.StringCellValue
        };


        protected internal double GetDouble(string value)
        {
            double.TryParse(value,ThisNumberStyle, ThisCultureInfo, out var doubleValue);
            return doubleValue;
        }

        protected internal double GetValueDouble(WrapperCell cell) => cell switch
        {
            {
                CellType: var cellType,
                NumericCellValue: var numericCellValue,
            }
            when cellType == CellType.Numeric => numericCellValue,
            {
                CellType: var cellType,
                StringCellValue: var stringCellValue,
            }
            when cellType == CellType.String => GetDouble(stringCellValue),
            {
                CellType: var cellType,
                StringCellValue: var stringCellValue,
                CachedFormulaResultType: var cachedFormulaResultType
            }
            when cellType == CellType.Formula &&
            cachedFormulaResultType == CellType.String
            => GetDouble(stringCellValue),
            {
                CellType: var cellType,
                NumericCellValue: var numericCellValue,
                CachedFormulaResultType: var cachedFormulaResultType
            }
            when cellType == CellType.Formula &&
            cachedFormulaResultType == CellType.Numeric
            => numericCellValue,
            _
            => GetDouble(cell.StringCellValue)
        };
    }
}
