using NPOI.SS.UserModel;
using System;
using System.Globalization;
using Mapster;
using System.Runtime.CompilerServices;

[assembly: InternalsVisibleTo("UnitTest")]
namespace WrapperNetPOI
{
    public class WrapperCell
    {
        ICell Cell { set; get; }
        public CellType CellType { set; get; }

        public WrapperCell(NPOI.SS.UserModel.ICell cell)
        {
            Cell = cell;
            if (cell == null)
            {
                CellType = CellType.Unknown;
            }
            else
            {
                CellType = cell.CellType;
            }
        }

        public CellType CachedFormulaResultType 
        {
            get
            {
                try 
                {
                    if (Cell == null)
                    {
                        return CellType.Unknown;
                    }
                    else
                    {
                        return Cell.CachedFormulaResultType;
                    }
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
                    if (Cell == null)
                    {
                        value = default;
                    }
                    else 
                    {
                        value = Cell.NumericCellValue;
                    }
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

                try
                {
                    value = Cell?.StringCellValue;
                }
                catch (InvalidOperationException ex)
                {
                    value = null;
                }
                catch (FormatException ex)
                {
                    value = null;
                }

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
                    if (Cell == null)
                    {
                        value = default;
                    }
                    else
                    {
                        value = Cell.DateCellValue;
                    }
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
            //CreateMapster();
        }

        private void CreateMapster()
        {
            Config = new TypeAdapterConfig();

           Config.ForType<WrapperCell, string>()
           .Map(dest => dest,
           src => GetValueString(src));

            Config.ForType<WrapperCell, double>()
            .Map(dest => dest,
            src => GetValueDouble(src));

           Config.ForType<WrapperCell, DateTime>()
            .Map(dest => dest,
            src => GetValueDateTime(src));

            Config.Compile();
        }

        public T MapGetValue<T>(NPOI.SS.UserModel.ICell cell)
        {
            WrapperCell wrapperCell = new(cell);
            return wrapperCell.Adapt<T>(Config);
            //return default;

        }

        

        public dynamic GetValue (NPOI.SS.UserModel.ICell cell, Type type)
        {
            switch (type.Name)
            {
                case "String":
                    GetValue(cell, out string tmp);
                    return tmp;
                 case "Double":
                    GetValue(cell, out double tmp1);
                    return tmp1;
                case "DateTime":
                    GetValue(cell, out DateTime tmp2);
                    return tmp2;
                default:
                    throw new NotImplementedException("Do not have handler");
            }
        }

        public void GetValue<T>(NPOI.SS.UserModel.ICell cell, out T value)
        {
            WrapperCell wrapperCell = new(cell);
            value = typeof(T).Name switch
            {
                "String" => (T)Convert.ChangeType(GetValueString(wrapperCell), typeof(T)),
                "Double" => (T)Convert.ChangeType(GetValueDouble(wrapperCell), typeof(T)),
                "DateTime" => (T)Convert.ChangeType(GetValueDateTime(wrapperCell), typeof(T)),
                _ => throw new NotImplementedException("Do not have handler"),
            };
        }

        public T GetValue<T>(NPOI.SS.UserModel.ICell cell)
        {
            GetValue<T>(cell,out T value);
            return value;
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
            try
            {
                return Convert.ToDateTime(value, ThisCultureInfo);
            }
            catch 
            {
                return default;
            }
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
