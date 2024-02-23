/* Необъединенное слияние из проекта "WrapperNetPOI (net6.0)"
До:
using System.Dynamic;
using System.Security.Authentication.ExtendedProtection;
using System.ComponentModel;
/* ==================================================================
После:
using Internal;
using NPOI.OpenXmlFormats.Dml;
/* ==================================================================
*/
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
/* Необъединенное слияние из проекта "WrapperNetPOI (net6.0)"
До:
using NPOI.SS.UserModel;
using NPOI.XSSF.Streaming.Values;
/* Необъединенное слияние из проекта "WrapperNetPOI (net6.0)"
После:
/* Необъединенное слияние из проекта "WrapperNetPOI (net6.0)"
*/
using NPOI.SS.UserModel;
/* Необъединенное слияние из проекта "WrapperNetPOI (net6.0)"
До:
using System.Configuration;
После:
using System.XWPF.UserModel;
*/
using
/* Необъединенное слияние из проекта "WrapperNetPOI (net6.0)"
До:
using System;
После:
using NPOI.SS.UserModel;
using NPOI.XSSF.Streaming.Values;
using System;
*/
System;
using System.ComponentModel;
using System.Globalization;
/* Необъединенное слияние из проекта "WrapperNetPOI (net6.0)"
До:
using System.Runtime.CompilerServices;
using System.Reflection;
После:
using System.ComponentModel;
using System.Configuration;
*/
using System.Linq;
using System.Runtime.CompilerServices;
/* Необъединенное слияние из проекта "WrapperNetPOI (net6.0)"
До:
using NPOI.OpenXmlFormats.Dmlkih/g,
using Internal;
using NPOI.XSSF.Streaming.Values;
using static NPOI.HSSF.Util.HSSFColor;
using NPOI.XWPF.UserModel;
using System.Linq;
После:
using System.Dynamic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Security.Authentication.ExtendedProtection;
using static NPOI.HSSF.Util.HSSFColor;
*/
[assembly: InternalsVisibleTo("UnitTest")]

namespace WrapperNetPOI.Excel
{
    //[TypeConverter(typeof(WrapperCellConverter))]
    public class WrapperCell : IConvertible
    {
        public CultureInfo ThisCultureInfo
        {
            get { return thisCultureInfo; }
            set { thisCultureInfo = value; }
        }
        public NumberStyles ThisNumberStyle { get; } = NumberStyles.Number;
        private CultureInfo thisCultureInfo = CultureInfo.CurrentCulture;
        public DateTimeStyles ThisDateTimeStyle { get; } = DateTimeStyles.AssumeUniversal;
        private NPOI.SS.UserModel.ICell Cell { get; }
        public CellType CellType { set; get; }
        /// <summary>
        /// return type for Auto Find Type in DataFrame
        /// </summary>
        public Type AutoType
        {
            get
            {
                return GetCellType(Cell);
            }
        }
        private Type GetCellType(NPOI.SS.UserModel.ICell cell) =>
            cell switch
            {
                { CellType: var cellType } when cellType == CellType.Blank => null,
                { CellType: var cellType } when cellType == CellType.Unknown => null,
                {
                    CellType: var cellType,
                }
                    when cellType == CellType.String
                    => FindTypeInString(Cell),
                { CellType: var cellType }
                    when cellType == CellType.Numeric
                    => FindTypeInNumeric(Cell),
                {
                    CellType: var cellType,
                    CachedFormulaResultType: var cachedFormulaResultType,
                } when cellType == CellType.Formula && cachedFormulaResultType == CellType.String
                    => FindTypeInString(Cell),
                {
                    CellType: var cellType,
                    CachedFormulaResultType: var cachedFormulaResultType
                } when cellType == CellType.Formula && cachedFormulaResultType == CellType.Numeric
                    => FindTypeInNumeric(Cell),
                _ => null
            };
        private Type GetCellType2(NPOI.SS.UserModel.ICell cell)
        {
            foreach (var x in new Type[] { typeof(DateTime), typeof(int), typeof(double), typeof(string) })
            {
                if (x == typeof(DateTime) && (DateTime)this.ToType(x, thisCultureInfo) != DateTime.FromOADate(default))
                {
                    return x;
                }
                else if (this.ToType(x, thisCultureInfo) != default)
                {
                    return x;
                }
            }
            return typeof(String);
        }
        public Type FindTypeInString(NPOI.SS.UserModel.ICell Cell)
        {
            var style = Cell.CellStyle;
            if (style.DataFormat != 0)
                return typeof(DateTime);
            else if (Cell.StringCellValue == "")
                return null;
            else
                return typeof(String);
        }
        public Type FindTypeInNumeric(NPOI.SS.UserModel.ICell Cell)
        {
            var style = Cell.CellStyle;
            var b = style.GetDataFormatString().All(x => new Char[] { 'H', 'D', 's', 'Y', 'M', 'm' }.Contains(x) == true);
            if (b & style.GetDataFormatString() != "General")
                return typeof(DateTime);
            else if (Cell.NumericCellValue % 1 == 0)
                return typeof(int);
            else
                return typeof(double);
        }
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
            set { Cell.SetCellValue(value); }
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
            set { Cell.SetCellValue(value); }
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
            set { Cell.SetCellValue(value); }
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
            return value == null ? new char() : value[0];
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
            switch (Type.GetTypeCode(conversionType))
            {
                case TypeCode.Boolean:
                    return this.ToBoolean(provider);
                case TypeCode.Byte:
                    return this.ToByte(provider);
                case TypeCode.Char:
                    return this.ToChar(provider);
                case TypeCode.DateTime:
                    return this.ToDateTime(provider);
                case TypeCode.Decimal:
                    return this.ToDecimal(provider);
                case TypeCode.Double:
                    return this.ToDouble(provider);
                case TypeCode.Int16:
                    return this.ToInt16(provider);
                case TypeCode.Int32:
                    return this.ToInt32(provider);
                case TypeCode.Int64:
                    return this.ToInt64(provider);
                case TypeCode.SByte:
                    return this.ToSByte(provider);
                case TypeCode.Single:
                    return this.ToSingle(provider);
                case TypeCode.String:
                    return this.ToString(provider);
                case TypeCode.UInt16:
                    return this.ToUInt16(provider);
                case TypeCode.UInt32:
                    return this.ToUInt32(provider);
                case TypeCode.UInt64:
                    return this.ToUInt64(provider);
                case TypeCode.DBNull:
                    break;
                case TypeCode.Empty:
                    break;
                case TypeCode.Object:
                    break;
                default:
                    throw new InvalidCastException(
                        String.Format("Conversion to {0} is not supported.", conversionType.Name)
                    );
            }
            return null;
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
        public void SetValue<T>(T value)
        {
            var val = value;
            Action b = value switch
            {
                String when value is string str => new Action(() => Cell.SetCellValue(str)),
                Double when value is double dbl => new Action(() => Cell.SetCellValue(dbl)),
                DateTime when value is DateTime dateTime
                    => new Action(() => Cell.SetCellValue(dateTime)),
                Int32 when value is Int32 int32 => new Action(() => Cell.SetCellValue(int32)),
                Boolean when value is Boolean boolean
                    => new Action(() => Cell.SetCellValue(boolean)),
                null when value is null => new Action(() => Cell.SetCellValue("")),
                _ => new Action(() => throw new NotImplementedException("Do not have handler"))
                //_=>new Action(() => Console.WriteLine(value))
            };
            b.Invoke();
        }
        private void SetCellValue<T>(T value)
        {
        }
        public dynamic GetValue(Type type)
        {
            return Convert.ChangeType(this, type);
        }
        public void GetValue<T>(out T value)
        {
            value = (T)Convert.ChangeType(this, typeof(T));
        }
        public T GetValue<T>()
        {
            GetValue(out T value);
            return value;
        }
        protected DateTime GetValueDateTime(WrapperCell cell) =>
            cell switch
            {
                { CellType: var cellType, } when cellType == CellType.Blank => default,
                { CellType: var cellType, StringCellValue: var stringCellValue, }
                    when cellType == CellType.String
                    => GetDateTime(stringCellValue),
                { CellType: var cellType, NumericCellValue: var numericCellValue, }
                    when cellType == CellType.Numeric
                    => GetDateTime(numericCellValue),
                {
                    CellType: var cellType,
                    StringCellValue: var stringCellValue,
                    CachedFormulaResultType: var cachedFormulaResultType
                } when cellType == CellType.Formula && cachedFormulaResultType == CellType.String
                    => GetDateTime(stringCellValue),
                {
                    CellType: var cellType,
                    NumericCellValue: var numericCellValue,
                    CachedFormulaResultType: var cachedFormulaResultType
                } when cellType == CellType.Formula && cachedFormulaResultType == CellType.Numeric
                    => GetDateTime(numericCellValue),
                _ => cell.DateCellValue
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
        protected static string GetValueString(WrapperCell cell) =>
            cell switch
            {
                { CellType: var cellType, } when cellType == CellType.Blank => null,
                { CellType: var cellType, StringCellValue: var stringCellValue, }
                    when cellType == CellType.String
                    => stringCellValue,
                { CellType: var cellType, NumericCellValue: var numericCellValue, }
                    when cellType == CellType.Numeric
                    => numericCellValue.ToString(),
                {
                    CellType: var cellType,
                    StringCellValue: var stringCellValue,
                    CachedFormulaResultType: var cachedFormulaResultType
                } when cellType == CellType.Formula && cachedFormulaResultType == CellType.String
                    => stringCellValue,
                {
                    CellType: var cellType,
                    NumericCellValue: var numericCellValue,
                    CachedFormulaResultType: var cachedFormulaResultType
                } when cellType == CellType.Formula && cachedFormulaResultType == CellType.Numeric
                    => numericCellValue.ToString(),
                _ => cell.StringCellValue,
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
        private double GetValueDouble(WrapperCell cell) =>
            cell switch
            {
                { CellType: var cellType, } when cellType == CellType.Blank => 0.0,
                { CellType: var cellType, NumericCellValue: var numericCellValue, }
                    when cellType == CellType.Numeric
                    => numericCellValue,
                { CellType: var cellType, StringCellValue: var stringCellValue, }
                    when cellType == CellType.String
                    => GetDouble(stringCellValue),
                {
                    CellType: var cellType,
                    StringCellValue: var stringCellValue,
                    CachedFormulaResultType: var cachedFormulaResultType
                } when cellType == CellType.Formula && cachedFormulaResultType == CellType.String
                    => GetDouble(stringCellValue),
                {
                    CellType: var cellType,
                    NumericCellValue: var numericCellValue,
                    CachedFormulaResultType: var cachedFormulaResultType
                } when cellType == CellType.Formula && cachedFormulaResultType == CellType.Numeric
                    => numericCellValue,
                _ => GetDouble(cell.StringCellValue)
            };
        private int GetValueInt32(WrapperCell cell) =>
            cell switch
            {
                { CellType: var cellType, } when cellType == CellType.Blank => 0,
                { CellType: var cellType, NumericCellValue: var numericCellValue, }
                    when cellType == CellType.Numeric
                    => (int)numericCellValue,
                { CellType: var cellType, StringCellValue: var stringCellValue, }
                    when cellType == CellType.String
                    => GetInt32(stringCellValue),
                {
                    CellType: var cellType,
                    StringCellValue: var stringCellValue,
                    CachedFormulaResultType: var cachedFormulaResultType
                } when cellType == CellType.Formula && cachedFormulaResultType == CellType.String
                    => GetInt32(stringCellValue),
                {
                    CellType: var cellType,
                    NumericCellValue: var numericCellValue,
                    CachedFormulaResultType: var cachedFormulaResultType
                } when cellType == CellType.Formula && cachedFormulaResultType == CellType.Numeric
                    => (int)numericCellValue,
                _ => GetInt32(cell.StringCellValue)
            };
        private bool GetValueBoolean(WrapperCell cell) =>
            cell switch
            {
                { CellType: var cellType, } when cellType == CellType.Blank => false,
                { CellType: var cellType, NumericCellValue: var numericCellValue, }
                    when cellType == CellType.Numeric
                    => GetBoolean(numericCellValue),
                { CellType: var cellType, StringCellValue: var stringCellValue, }
                    when cellType == CellType.String
                    => GetBoolean(stringCellValue),
                {
                    CellType: var cellType,
                    StringCellValue: var stringCellValue,
                    CachedFormulaResultType: var cachedFormulaResultType
                } when cellType == CellType.Formula && cachedFormulaResultType == CellType.String
                    => GetBoolean(stringCellValue),
                {
                    CellType: var cellType,
                    NumericCellValue: var numericCellValue,
                    CachedFormulaResultType: var cachedFormulaResultType
                } when cellType == CellType.Formula && cachedFormulaResultType == CellType.Numeric
                    => GetBoolean(numericCellValue),
                _ => GetBoolean(cell.StringCellValue)
            };
        public static CellType ReturnCellType(Type type) =>
            type switch
            {
                { Name: var nameof, } when nameof == "String" => CellType.String,
                _ => CellType.Numeric
                /*int=>CellType.Numeric,
                double=>CellType.Numeric,
                bool=>CellType.Numeric,
                DateTime=>CellType.Numeric,
                null=>CellType.Blank*/
            };
    }
    public class WrapperCellConverter : TypeConverter
    {
        public override bool CanConvertFrom(ITypeDescriptorContext context, Type destinationType)
        {
            switch (destinationType.Name)
            {
                case "String":
                    return true;
                case "Int32":
                    return true;
                case "Double":
                    return true;
                case "DateTime":
                    return true;
                default:
                    return false;
            }
        }
        public override object ConvertFrom(
            ITypeDescriptorContext descriptorContext,
            CultureInfo cultureInfo,
            Object value
        )
        {
            if (value is string valueStr)
            {
                ((WrapperCell)descriptorContext.Instance).StringCellValue = valueStr;
                return (WrapperCell)descriptorContext.Instance;
            }
            else
            {
                return default;
            }
        }
    }
}
//}
