/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at
       http://www.apache.org/licenses/LICENSE-2.0
   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */

using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System;

namespace WrapperNetPOI.Excel
{
    /// <summary>
    /// Source code copied from https://github.com/nissl-lab/npoi/blob/master/main/SS/Util/SheetUtil.cs
    /// </summary>
    public static class ChangedNPOI
    {
        /// <summary>
        /// Source file has been taken from npoi/main/SS/Util/SheetUtil.cs
        /// </summary>
        /// <param name="oldCell"></param>
        /// <param name="newCell"></param>
        public static void CopyValue(ICell oldCell, ICell newCell)
        {
            switch (oldCell.CellType)
            {
                case CellType.Blank:
                    newCell.SetCellValue(oldCell.StringCellValue);
                    break;

                case CellType.Boolean:
                    newCell.SetCellValue(oldCell.BooleanCellValue);
                    break;

                case CellType.Error:
                    newCell.SetCellErrorValue(oldCell.ErrorCellValue);
                    break;

                case CellType.Formula:
                    newCell.SetCellFormula(oldCell.CellFormula);
                    break;

                case CellType.Numeric:
                    newCell.SetCellValue(oldCell.NumericCellValue);
                    break;

                case CellType.String:
                    if (oldCell.GetType() != newCell.GetType())
                    {
                        newCell.SetCellValue(oldCell.RichStringCellValue.String);
                    }
                    else
                    {
                        newCell.SetCellValue(oldCell.RichStringCellValue);
                    }
                    break;
            }
        }

        /// <summary>
        /// Source file has been taken from npoi/main/SS/Util/SheetUtil.cs
        /// </summary>
        /// <param name="sourceSheet"></param>
        /// <param name="sourceRowIndex"></param>
        /// <param name="targetSheet"></param>
        /// <param name="targetRowIndex"></param>
        /// <returns></returns>
        /// <exception cref="ArgumentNullException"></exception>
        public static IRow ChangedCopyRow(ISheet sourceSheet, int sourceRowIndex, ISheet targetSheet, int targetRowIndex)
        {
            // Get the source / new row
            IRow newRow = targetSheet.GetRow(targetRowIndex);
            IRow sourceRow = sourceSheet.GetRow(sourceRowIndex);

            // If the row exist in destination, push down all rows by 1 else create a new row
            if (newRow != null)
            {
                targetSheet.RemoveRow(newRow);
            }
            newRow = targetSheet.CreateRow(targetRowIndex);
            if (sourceRow == null)
            {
                ArgumentNullException argumentNullException = new(nameof(ChangedCopyRow), " ChangedCopyRow source row doesn't exist");
                throw argumentNullException;
            }
            // Loop through source columns to add to new row
            for (int i = sourceRow.FirstCellNum; i < sourceRow.LastCellNum; i++)
            {
                // Grab a copy of the old/new cell
                ICell oldCell = sourceRow.GetCell(i);

                // If the old cell is null jump to next cell
                if (oldCell == null)
                {
                    continue;
                }
                ICell newCell = newRow.CreateCell(i);

                if (oldCell.CellStyle != null)
                {
                    var CellStyle = targetSheet.Workbook.CreateCellStyle();

                    if (oldCell.GetType() != newCell.GetType())
                    {
                        CopyStyle(oldCell.CellStyle, CellStyle);
                        // not copy all format. Until I came up with something else
                    }
                    else
                    {
                        CellStyle.CloneStyleFrom(oldCell.CellStyle);
                    }
                    newCell.CellStyle = CellStyle;
                }

                // If there is a cell comment, copy
                if (oldCell.CellComment != null)
                {
                    newCell.CellComment = oldCell.CellComment;
                }

                // If there is a cell hyperlink, copy
                if (oldCell.Hyperlink != null)
                {
                    newCell.Hyperlink = oldCell.Hyperlink;
                }

                // Set the cell data type
                newCell.SetCellType(oldCell.CellType);
                // Set the cell data value
                CopyValue(oldCell, newCell);
            }

            // If there are are any merged regions in the source row, copy to new row
            for (int i = 0; i < sourceSheet.NumMergedRegions; i++)
            {
                CellRangeAddress cellRangeAddress = sourceSheet.GetMergedRegion(i);

                if (cellRangeAddress != null && cellRangeAddress.FirstRow == sourceRow.RowNum)
                {
                    CellRangeAddress newCellRangeAddress = new(newRow.RowNum,
                            newRow.RowNum +
                                    (cellRangeAddress.LastRow - cellRangeAddress.FirstRow
                                            ),
                            cellRangeAddress.FirstColumn,
                            cellRangeAddress.LastColumn);
                    targetSheet.AddMergedRegion(newCellRangeAddress);
                }
            }
            return newRow;
        }

        public static void CopyStyle(ICellStyle source, ICellStyle recipient)
        {
            var propertiesS = source.GetType().GetProperties();
            foreach (var propertyS in propertiesS)
            {
                var sourceValue = propertyS.GetValue(source);
                var sourceName = propertyS.Name;
                var sourceValueType = propertyS.PropertyType;
                var propertiesR = recipient.GetType().GetProperties();

                foreach (var propertyR in propertiesR)
                {
                    if (propertyR.Name == propertyS.Name)
                    {
                        if (propertyR.SetMethod != null)
                        {
                            try
                            {
                                propertyR.SetValue(recipient, sourceValue);
                            }
                            catch (Exception e)
                            {
#if DEBUG
                                Console.WriteLine(e.Message);
#endif
                            }
                        }
                    }
                }
            }
        }
    }
}