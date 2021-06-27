using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Reflection.Metadata.Ecma335;
using System.Runtime.InteropServices;
using System.Text;
using FluentEpplus.Common.Base;
using FluentEpplus.Common.Processor;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace FluentEpplus.Tables.Processor
{
    public class ExcelComplexTableProcessor : IExcelHeaderProcessor, IExcelRowProcessor
    {
        public int ProcessHeader(ExcelWorksheet worksheet, int row, IExcelContainerCells container)
        {
            return ProcessHeaderWithCurrentRow(worksheet, row, row, container);
        }

        private int ProcessHeaderWithCurrentRow(ExcelWorksheet worksheet, int row, int currentRow, IExcelContainerCells container)
        {
            var incrementRow = 1;

            var groupCells = container.Cells.OfType<IExcelContainerCells>();
            var maxCurrentRow = 0;

            foreach (var cell in groupCells)
            {
                if (cell is IExcelComplexGroupCell) continue;
                incrementRow = container.ShowCaption ? 1 : 0;
                maxCurrentRow = ProcessHeaderWithCurrentRow(worksheet, row + incrementRow, currentRow, cell);

                currentRow = (maxCurrentRow > currentRow) ? maxCurrentRow : currentRow;
            }

            var isEnableToWriteCaption = !(container is IExcelTable);
            isEnableToWriteCaption |= container is IExcelRootTable;

            if (container.ShowCaption && isEnableToWriteCaption)
            {
                worksheet.Cells[row, container.ColumnOrder, row, container.GetMaxOrder()].Value = container.Caption;
                worksheet.Cells[row, container.ColumnOrder, row, container.GetMaxOrder()].Merge = true;
                worksheet.Cells[row, container.ColumnOrder, row, container.ColumnOrder].Style.Font.Bold = true;
                worksheet.Cells[row, container.ColumnOrder, row, container.ColumnOrder].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                container.ApplyHeaderStyle(worksheet.Cells[row, container.ColumnOrder, row, container.GetMaxOrder()]);

                row++;
                currentRow = currentRow < row ? row : currentRow;
            }

            foreach (var cell in container.Cells)
            {
                var groupCell = cell as IExcelGroupCell;

                if (groupCell == null || cell is IExcelComplexGroupCell)
                {
                    var columnOrder = cell is IExcelComplexGroupCell ? ((IExcelComplexGroupCell) cell).GetMaxOrder() : cell.ColumnOrder;
                    worksheet.Cells[row, cell.ColumnOrder, currentRow, columnOrder].Value = cell.Caption;
                    cell.ApplyHeaderStyle(worksheet.Cells[row, cell.ColumnOrder, currentRow, columnOrder]);

                    if (cell.ColumnOrder != columnOrder || row != currentRow)
                    {
                        worksheet.Cells[row, cell.ColumnOrder, currentRow, columnOrder].Merge = true;
                        if (cell.IsAutoFit)
                            worksheet.Cells[row, cell.ColumnOrder, currentRow, columnOrder].AutoFitColumns();
                    }
                    else
                    {
                        if (cell.IsAutoFit)
                            worksheet.Column(cell.ColumnOrder).AutoFit();
                    }

                    worksheet.Cells[row, cell.ColumnOrder, currentRow, columnOrder].Style.Font.Bold = true;
                    worksheet.Cells[row, cell.ColumnOrder, currentRow, columnOrder].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                }

            }

            return currentRow;
        }

        public int ProcessRowFromEnumerable(ExcelWorksheet worksheet, int row, IExcelContainerCells container, IEnumerable data)
        {
            foreach (var item in data)
            {
                row = ProcessRowWithCurrentRow(worksheet, row, row, container, item);
            }

            return row;
        }

        public int ProcessRowFromEnumerableWithCurrentRow(ExcelWorksheet worksheet, int row, int currentRow, IExcelContainerCells container, IEnumerable data)
        {
            foreach (var item in data)
            {
                row = ProcessRowWithCurrentRow(worksheet, row, currentRow, container, item);
            }

            return row;
        }

        public int ProcessRow(ExcelWorksheet worksheet, int row, IExcelContainerCells container, object data)
        {
            return ProcessRowWithCurrentRow(worksheet, row, row, container, data);
        }

        private int ProcessRowWithCurrentRow(ExcelWorksheet worksheet, int row, int currentRow, IExcelContainerCells container, object data)
        {
            var complexCells = container.Cells.OfType<IExcelComplexGroupCell>();
            var maxCurrentRow = 0;

            foreach (var cell in complexCells)
            {
                var lastHeaderRow = row;
                if (cell.IsShowHeaderPerRow)
                {
                    lastHeaderRow = ProcessHeader(worksheet, lastHeaderRow, cell) + 1;
                }

                var newData = cell.GetValue(data);
                if (newData is Enumerable)
                {
                    maxCurrentRow = ProcessRowFromEnumerable(worksheet, lastHeaderRow, cell, (IEnumerable)newData) - 1;
                }
                else
                {
                    maxCurrentRow = ProcessRow(worksheet, lastHeaderRow, cell, newData) - 1;
                }

                currentRow = (maxCurrentRow > currentRow) ? maxCurrentRow : currentRow;
            }

            foreach (var cell in container.Cells)
            {
                var groupComplexCells = cell as IExcelComplexGroupCell;

                if (groupComplexCells == null && cell is IExcelGroupCell)
                {
                    if (data is IEnumerable)
                        ProcessRowFromEnumerableWithCurrentRow(worksheet, row, currentRow, (IExcelGroupCell) cell, (IEnumerable) cell.GetValue(data));
                    else
                        ProcessRowWithCurrentRow(worksheet, row, currentRow, (IExcelGroupCell) cell, cell.GetValue(data));
                    continue;
                }

                if (groupComplexCells == null)
                {
                    worksheet.Cells[row, cell.ColumnOrder, currentRow, cell.ColumnOrder].Value = cell.GetValue(data);
                    cell.ApplyStyle(worksheet.Cells[row, cell.ColumnOrder, currentRow, cell.ColumnOrder], data);

                    if (cell.IsAutoFit)
                        worksheet.Column(cell.ColumnOrder).AutoFit();

                    if (row != currentRow)
                    {
                        worksheet.Cells[row, cell.ColumnOrder, currentRow, cell.ColumnOrder].Merge = true;
                        worksheet.Cells[row, cell.ColumnOrder, currentRow, cell.ColumnOrder].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    }

                    worksheet.Cells[row, cell.ColumnOrder, currentRow, cell.ColumnOrder].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                }
            }

            return ++currentRow;
        }
    }
}
