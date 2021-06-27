using System.Collections;
using FluentEpplus.Common.Base;
using OfficeOpenXml;

namespace FluentEpplus.Common.Processor
{
    public interface IExcelRowProcessor
    {
        int ProcessRowFromEnumerable(ExcelWorksheet worksheet, int row, IExcelContainerCells container,
            IEnumerable data);

        int ProcessRow(ExcelWorksheet worksheet, int row, IExcelContainerCells container, object data);
    }
}