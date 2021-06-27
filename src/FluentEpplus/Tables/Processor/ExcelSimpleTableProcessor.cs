using System.Collections;
using System.Data;
using FluentEpplus.Common.Base;
using FluentEpplus.Common.Processor;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace FluentEpplus.Tables.Processor
{
    public class ExcelSimpleTableProcessor : IExcelHeaderProcessor, IExcelRowProcessor
    {
        public int ProcessHeader(ExcelWorksheet worksheet, int row, IExcelContainerCells container)
        {
            var incrementRow = 1;

            if (container.ShowCaption)
            {
                worksheet.Cells[row, container.ColumnOrder, row, container.GetMaxOrder()].Value = container.Caption;
                worksheet.Cells[row, container.ColumnOrder, row, container.GetMaxOrder()].Merge = true;
                worksheet.Cells[row, container.ColumnOrder, row, container.ColumnOrder].Style.Font.Bold = true;
                worksheet.Cells[row, container.ColumnOrder, row, container.ColumnOrder].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                container.ApplyHeaderStyle(worksheet.Cells[row, container.ColumnOrder, row, container.ColumnOrder]);
            }
            else
            {
                incrementRow--;
            }

            foreach (var cell in container.Cells)
            {
                worksheet.Cells[row + incrementRow, cell.ColumnOrder].Value = cell.Caption;
                worksheet.Cells[row + incrementRow, cell.ColumnOrder].Style.Font.Bold = true;
                worksheet.Cells[row + incrementRow, cell.ColumnOrder].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                cell.ApplyHeaderStyle(worksheet.Cells[row + incrementRow, cell.ColumnOrder]);
            }

            return row++ + incrementRow;
        }

        public int ProcessRowFromEnumerable(ExcelWorksheet worksheet, int row, IExcelContainerCells container,
            IEnumerable data)
        {
            foreach (var item in data)
            {
                row = ProcessRow(worksheet, row, container, item);
            }

            return row;
        }

        public int ProcessRow(ExcelWorksheet worksheet, int row, IExcelContainerCells container, object data)
        {
            foreach (var property in container.Cells)
            {
                if (property is IExcelGroupCell)
                {
                    ProcessRow(worksheet, row, (IExcelGroupCell) property, property.GetValue(data));
                    continue;
                }

                worksheet.Cells[row, property.ColumnOrder].Value = property.GetValue(data);
                property.ApplyStyle(worksheet.Cells[row, property.ColumnOrder],data);

                if (property.IsAutoFit)
                    worksheet.Column(property.ColumnOrder).AutoFit();
            }

            return ++row;
        }
    }
}