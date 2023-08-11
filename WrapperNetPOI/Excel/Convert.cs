using System.ComponentModel;
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

using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using System;
using System.Configuration;
using System.Globalization;
using System.Runtime.CompilerServices;
using System.Reflection;
using System.Diagnostics.CodeAnalysis;
using NPOI.OpenXmlFormats.Dml;

[assembly: InternalsVisibleTo("UnitTest")]

namespace WrapperNetPOI.Excel
{


    public class WrapperCell : IConvertible
    {
        public CultureInfo ThisCultureInfo { get; } = CultureInfo.CurrentCulture;
        public NumberStyles ThisNumberStyle { get; } = NumberStyles.Number;
        public DateTimeStyles ThisDateTimeStyle { get; } = DateTimeStyles.AssumeUniversal;
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

        public TypeCode GetTypeCode()
        {
            return Type.GetTypeCode(Cell.GetType());
        }

        public bool ToBoolean(IFormatProvider provider)
        {
            return GetValueBoolean(this);
        }

        public byte ToByte(IFormatProvider provider)
        {
            return Convert.ToByte(GetValueInt32(this));
        }

        public char ToChar(IFormatProvider provider)
        {
            var value = GetValueString(this);
            return value == null
                ? new char()
                : value[0];
        }

        public DateTime ToDateTime(IFormatProvider provider)
        {
            return GetValueDateTime(this);
        }

        public decimal ToDecimal(IFormatProvider provider)
        {
            return Convert.ToDecimal(GetValueInt32(this));
        }

        public double ToDouble(IFormatProvider provider)
        {
            return GetValueDouble(this);
        }

        public short ToInt16(IFormatProvider provider)
        {
            return Convert.ToInt16(GetValueInt32(this));
        }

        public int ToInt32(IFormatProvider provider)
        {
            return GetValueInt32(this);
        }

        public long ToInt64(IFormatProvider provider)
        {
            return GetValueInt32(this);
        }

        public sbyte ToSByte(IFormatProvider provider)
        {
            return Convert.ToSByte(GetValueInt32(this));
        }

        public float ToSingle(IFormatProvider provider)
        {
            return Convert.ToSingle(GetValueInt32(this));
        }

        public string ToString(IFormatProvider provider)
        {
            return GetValueString(this);
        }

        public override string ToString()
        {
            return GetValueString(this);
        }

        public object ToType(Type conversionType, IFormatProvider provider)
        {
            throw new NotImplementedException();
        }

        public ushort ToUInt16(IFormatProvider provider)
        {
            return Convert.ToUInt16(GetValueInt32(this));
        }

        public uint ToUInt32(IFormatProvider provider)
        {
            return Convert.ToUInt32(GetValueInt32(this));
        }

        public ulong ToUInt64(IFormatProvider provider)
        {
            return Convert.ToUInt64(GetValueInt32(this));
        }


        //public class ConvertType
        //{

        //public ConvertType() { }

        public static void SetValue<T>(ICell cell, T value)
        {
            Action b = value switch
            {
                String when value is string str => new Action(() => cell.SetCellValue(str)),
                Double when value is string dbl => new Action(() => cell.SetCellValue(dbl)),
                DateTime when value is DateTime dateTime => new Action(() => cell.SetCellValue(dateTime)),
                Int32 when value is Int32 int32 => new Action(() => cell.SetCellValue(int32)),
                Boolean when value is Boolean boolean => new Action(() => cell.SetCellValue(boolean)),
                null when value is null => new Action(() => cell.SetCellValue("")),
                _ => new Action(() => throw new NotImplementedException("Do not have handler"))
            };
            b.Invoke();
        }

        public void SetValue(object value, Type type)
        {
            Action b = type.Name switch
            {
                "String" when value is string str => new Action(() => Cell.SetCellValue(str)),
                "Double" when value is double dbl => new Action(() => Cell.SetCellValue(dbl)),
                "DateTime" when value is DateTime dateTime => new Action(() => Cell.SetCellValue(dateTime)),
                "Int32" when value is Int32 int32 => new Action(() => Cell.SetCellValue(int32)),
                "Boolean" when value is Boolean boolean => new Action(() => Cell.SetCellValue(boolean)),
                null when value is null => new Action(() => Cell.SetCellValue("")),
                _ => new Action(() => throw new NotImplementedException("Do not have handler"))
            };
            b.Invoke();
            var convertedvalue = Convert.ChangeType(value, type);
        }


        public dynamic GetValue(ICell cell, Type type)
        {
            //var value= Convert.ChangeType(Activator.CreateInstance(type),type);
            //GetValue(cell, out value);
            //return value;
            //GetValue(cell, out value);

            WrapperCell wrapperCell = new(cell);
            dynamic value = type.Name switch
            {
                "String" => wrapperCell.ToString(),
                "Double" => wrapperCell.ToDouble(ThisCultureInfo),
                "DateTime" => wrapperCell.ToDateTime(ThisCultureInfo),
                "Int32" => wrapperCell.ToInt32(ThisCultureInfo),
                "Boolean" => wrapperCell.ToBoolean(ThisCultureInfo),
                _ => throw new NotImplementedException("Do not have handler"),
            };


            return value;
        }
        public void GetValue<T>(out T value)
        {
            value = GetValue(Cell, typeof(T));
            value = typeof(T).Name switch
            {
                "String" => (T)Convert.ChangeType(this.ToString(), typeof(T)),
                "Double" => (T)Convert.ChangeType(this.ToDouble(ThisCultureInfo), typeof(T)),
                "DateTime" => (T)Convert.ChangeType(this.ToDateTime(ThisCultureInfo), typeof(T)),
                "Int32" => (T)Convert.ChangeType(this.ToInt32(ThisCultureInfo), typeof(T)),
                "Boolean" => (T)Convert.ChangeType(this.ToBoolean(ThisCultureInfo), typeof(T)),
                _ => throw new NotImplementedException("Do not have handler"),
            };

        }
        public T GetValue<T>()
        {
            GetValue(out T value);
            return value;
        }
        protected DateTime GetValueDateTime(WrapperCell cell) => cell switch
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

        protected static string GetValueString(WrapperCell cell) => cell switch
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

        public static CellType ReturnCellType(Type type) => type switch
        {
            {
                Name: var nameof,
            } when nameof == "String" => CellType.String,
            _
            => CellType.Numeric
            /*int=>CellType.Numeric,
            double=>CellType.Numeric,
            bool=>CellType.Numeric,
            DateTime=>CellType.Numeric,
            null=>CellType.Blank*/
        };
    }
}
//}
