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
using System.Configuration;
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
            set
            {
                Cell.SetCellValue(value);
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
            set
            {
                Cell.SetCellValue(value);
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
            set
            {
                Cell.SetCellValue(value);
            }
        }

    }

    public class ConvertType
    {
        public CultureInfo ThisCultureInfo { get; } = CultureInfo.CurrentCulture;
        public NumberStyles ThisNumberStyle { get; } = NumberStyles.Number;
        public DateTimeStyles ThisDateTimeStyle { get; } = DateTimeStyles.AssumeUniversal;
        public ConvertType() {}

        public static void SetValue<T>(ICell cell,T value)
        {
            Action b = value switch
            {
                String when value is string str => new Action(() => cell.SetCellValue(str)),
                Double when value is string dbl => new Action(() => cell.SetCellValue(dbl)),
                DateTime when value is DateTime dateTime => new Action(() => cell.SetCellValue(dateTime)),
                Int32 when value is Int32 int32 => new Action(() => cell.SetCellValue(int32)),
                Boolean when value is Boolean boolean => new Action(() => cell.SetCellValue(boolean)),
                null when value is null => new Action(()=> cell.SetCellValue("")), 
                _ => new Action(()=>throw new NotImplementedException("Do not have handler"))
            }; ;
            b.Invoke();
        }

        public dynamic GetValue(ICell cell, Type type)
        {
            dynamic value;
            WrapperCell wrapperCell = new(cell);
            value = type.Name switch
            {
                "String" => Convert.ChangeType(GetValueString(wrapperCell), type),
                "Double" => Convert.ChangeType(GetValueDouble(wrapperCell), type),
                "DateTime" => Convert.ChangeType(GetValueDateTime(wrapperCell), type),
                "Int32" => Convert.ChangeType(GetValueInt32(wrapperCell), type),
                "Boolean" => Convert.ChangeType(GetValueBoolean(wrapperCell), type),
                _ => throw new NotImplementedException("Do not have handler"),
            };
            return value;
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
                "Boolean" => (T)Convert.ChangeType(GetValueBoolean(wrapperCell), typeof(T)),
                _ => throw new NotImplementedException("Do not have handler"),
            };
        }
        public T GetValue<T>(ICell cell)
        {
            GetValue(cell, out T value);
            return value;
        }
        private DateTime GetValueDateTime(WrapperCell cell) => cell switch
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

        private DateTime GetDateTime(string value)
        {
            DateTime.TryParse(value, ThisCultureInfo, ThisDateTimeStyle, out var outValue);
            return outValue;
        }

        private DateTime GetDateTime(double value)
        {
            try
            {
                return DateTime.FromOADate(value);
            }
            catch (Exception e)
            {
                Wrapper.Logger?.Error(e.Message);
                Wrapper.Logger?.Error(e.StackTrace);
                return default;
            }
        }

        private static string GetValueString(WrapperCell cell) => cell switch
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

        private double GetDouble(string value)
        {
            double.TryParse(value, ThisNumberStyle, ThisCultureInfo, out var doubleValue);
            return doubleValue;
        }

        private int GetInt32(string value)
        {
            int.TryParse(value, ThisNumberStyle, ThisCultureInfo, out var intValue);
            return intValue;
        }

        private bool GetBoolean(string value)
        {
            bool.TryParse(value, out var boolValue);
            return boolValue;
        }

        private bool GetBoolean(double value)
        {
            return value switch
            {
                0.0 => false,
                _ => true,
            };
        }

        private double GetValueDouble(WrapperCell cell) => cell switch
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

        private int GetValueInt32(WrapperCell cell) => cell switch
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

        private bool GetValueBoolean(WrapperCell cell) => cell switch
        {
            {
                CellType: var cellType,
            }
            when cellType == CellType.Blank => false,
            {
                CellType: var cellType,
                NumericCellValue: var numericCellValue,
            }
            when cellType == CellType.Numeric => GetBoolean(numericCellValue),
            {
                CellType: var cellType,
                StringCellValue: var stringCellValue,
            }
            when cellType == CellType.String => GetBoolean(stringCellValue),
            {
                CellType: var cellType,
                StringCellValue: var stringCellValue,
                CachedFormulaResultType: var cachedFormulaResultType
            }
            when cellType == CellType.Formula &&
            cachedFormulaResultType == CellType.String
            => GetBoolean(stringCellValue),
            {
                CellType: var cellType,
                NumericCellValue: var numericCellValue,
                CachedFormulaResultType: var cachedFormulaResultType
            }
            when cellType == CellType.Formula &&
            cachedFormulaResultType == CellType.Numeric
            => GetBoolean(numericCellValue),
            _
            => GetBoolean(cell.StringCellValue)
        };
    }
}