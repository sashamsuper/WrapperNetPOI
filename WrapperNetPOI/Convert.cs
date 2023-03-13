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

namespace WrapperNetPOI
{
    public class WrapperCell
    {
        private ICell Cell { get; }
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
                    value = Cell?.StringCellValue;
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

        public dynamic GetValue(NPOI.SS.UserModel.ICell cell, Type type)
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
            GetValue<T>(cell, out T value);
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
            DateTime.TryParse(value, ThisCultureInfo, ThisDateTimeStyle, out var doubleValue);
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
            double.TryParse(value, ThisNumberStyle, ThisCultureInfo, out var doubleValue);
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