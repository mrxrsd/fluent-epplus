using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Text;
using FluentEpplus.Common.Base;
using FluentEpplus.Common.Processor;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace FluentEpplus.Forms.Processor
{
    public class ExcelComplexFormProcessor : IExcelHeaderProcessor, IExcelRowProcessor
    {
        public int ProcessHeader(ExcelWorksheet worksheet, int row, IExcelContainerCells container)
        {
            var incrementRow = row;

            if (container.ShowCaption)
            {
                var captionRow = row + container.Cells.Count - 1;
                worksheet.Cells[row, container.ColumnOrder, captionRow, container.ColumnOrder].Value = container.Caption;
                worksheet.Cells[row, container.ColumnOrder, captionRow, container.ColumnOrder].Merge = true;
                worksheet.Cells[row, container.ColumnOrder, captionRow, container.ColumnOrder].Style.Font.Bold = true;
                worksheet.Cells[row, container.ColumnOrder, captionRow, container.ColumnOrder].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                container.ApplyHeaderStyle(worksheet.Cells[row, container.ColumnOrder, captionRow, container.ColumnOrder]);

                if (container.IsAutoFit)
                    worksheet.Column(container.ColumnOrder).AutoFit();
            }

            foreach (var cell in container.Cells)
            {
                worksheet.Cells[incrementRow, cell.ColumnOrder].Value = cell.Caption;
                worksheet.Cells[incrementRow, cell.ColumnOrder].Style.Font.Bold = true;
                worksheet.Cells[incrementRow, cell.ColumnOrder].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                cell.ApplyHeaderStyle(worksheet.Cells[incrementRow, cell.ColumnOrder]);

                if (cell.IsAutoFit)
                    worksheet.Column(cell.ColumnOrder).AutoFit();
                incrementRow++;
            }


            return row - 1;
        }

        public int ProcessRowFromEnumerable(ExcelWorksheet worksheet, int row, IExcelContainerCells container, IEnumerable data)
        {
            var isFirstRow = true;
            foreach (var item in data)
            {
                if (!isFirstRow)
                    row = ProcessHeader(worksheet, row, container) + 1;
                row = ProcessRow(worksheet, row, container, item);
                isFirstRow = false;
            }

            return row;
        }

        public int ProcessRow(ExcelWorksheet worksheet, int row, IExcelContainerCells container, object data)
        {
            foreach (var cell in container.Cells)
            {
                var cellOrderValue = cell.ColumnOrder + 1;
                worksheet.Cells[row, cellOrderValue].Value = cell.GetValue(data);

                if (cell.IsAutoFit)
                    worksheet.Column(cellOrderValue).AutoFit();

                cell.ApplyStyle(worksheet.Cells[row,cellOrderValue],data);
                row++;
            }

            return row;
        }
    }
}
