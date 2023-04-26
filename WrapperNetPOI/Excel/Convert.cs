/* ==================================================================
Copyright 2020-2023 sashamsuper

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

    http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.
==========================================================================*/

using NPOI.SS.UserModel;
using System;
using System.Globalization;
using System.Runtime.CompilerServices;

[assembly: InternalsVisibleTo("UnitTest")]

namespace WrapperNetPOI.Excel
{
    public class WrapperCell
    {
        private ICell Cell { get; }
        public CellType CellType { set; get; }

        public WrapperCell(ICell cell)
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
                catch (InvalidOperationException)
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
                catch (InvalidOperationException)
                {
                    value = 0;
                }
                catch (FormatException)
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
                    if (Cell?.CellType == CellType.String)
                    {
                        value = Cell?.StringCellValue;
                    }
                    else
                    {
                        value = null;
                    }
                }
                catch (InvalidOperationException)
                {
                    value = null;
                }
                catch (FormatException)
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
                catch (InvalidOperationException)
                {
                    value = default;
                }
                catch (FormatException)
                {
                    value = default;
                }

                return value;
            }
        }
    }

    public class ConvertType
    {
        private CultureInfo ThisCultureInfo { get; } = CultureInfo.CurrentCulture;
        private NumberStyles ThisNumberStyle { get; } = NumberStyles.Number;
        private DateTimeStyles ThisDateTimeStyle { get; } = DateTimeStyles.AssumeUniversal;

        public ConvertType()
        {
            //CreateMapster();
        }

        public dynamic GetValue(ICell cell, Type type)
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

                case "Int32":
                    GetValue(cell, out int tmp3);
                    return tmp3;

                default:
                    throw new NotImplementedException("Do not have handler");
            }
        }

        public void GetValue<T>(ICell cell, out T value)
        {
            WrapperCell wrapperCell = new(cell);
            value = typeof(T).Name switch
            {
                "String" => (T)Convert.ChangeType(GetValueString(wrapperCell), typeof(T)),
                "Double" => (T)Convert.ChangeType(GetValueDouble(wrapperCell), typeof(T)),
                "DateTime" => (T)Convert.ChangeType(GetValueDateTime(wrapperCell), typeof(T)),
                "Int32" => (T)Convert.ChangeType(GetValueInt32(wrapperCell), typeof(T)),
                _ => throw new NotImplementedException("Do not have handler"),
            };
        }

        public T GetValue<T>(ICell cell)
        {
            GetValue(cell, out T value);
            return value;
        }

        protected internal DateTime GetValueDateTime(WrapperCell cell) => cell switch
        {
            {
                CellType: var cellType,
            }
            when cellType == CellType.Blank => default,
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
            DateTime.TryParse(value, ThisCultureInfo, ThisDateTimeStyle, out var outValue);
            return outValue;
        }

        public DateTime GetDateTime(double value)
        {
            try
            {
                return DateTime.FromOADate(value);
                //return default;
                //Convert.ToDateTime(value, ThisCultureInfo);
            }
            catch (Exception e)
            {
                Wrapper.Logger?.Error(e.Message);
                Wrapper.Logger?.Error(e.StackTrace);
                return default;
            }
        }

        protected internal static string GetValueString(WrapperCell cell) => cell switch
        {
            {
                CellType: var cellType,
            }
            when cellType == CellType.Blank => null,
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
            => cell.StringCellValue,
        };

        protected internal double GetDouble(string value)
        {
            double.TryParse(value, ThisNumberStyle, ThisCultureInfo, out var doubleValue);
            return doubleValue;
        }

        protected internal int GetInt32(string value)
        {
            int.TryParse(value, ThisNumberStyle, ThisCultureInfo, out var intValue);
            return intValue;
        }

        protected internal double GetValueDouble(WrapperCell cell) => cell switch
        {
            {
                CellType: var cellType,
            }
            when cellType == CellType.Blank => 0.0,
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

        protected internal int GetValueInt32(WrapperCell cell) => cell switch
        {
            {
                CellType: var cellType,
            }
            when cellType == CellType.Blank => 0,
            {
                CellType: var cellType,
                NumericCellValue: var numericCellValue,
            }
            when cellType == CellType.Numeric => (int)numericCellValue,
            {
                CellType: var cellType,
                StringCellValue: var stringCellValue,
            }
            when cellType == CellType.String => GetInt32(stringCellValue),
            {
                CellType: var cellType,
                StringCellValue: var stringCellValue,
                CachedFormulaResultType: var cachedFormulaResultType
            }
            when cellType == CellType.Formula &&
            cachedFormulaResultType == CellType.String
            => GetInt32(stringCellValue),
            {
                CellType: var cellType,
                NumericCellValue: var numericCellValue,
                CachedFormulaResultType: var cachedFormulaResultType
            }
            when cellType == CellType.Formula &&
            cachedFormulaResultType == CellType.Numeric
            => (int)numericCellValue,
            _
            => GetInt32(cell.StringCellValue)
        };
    }
}