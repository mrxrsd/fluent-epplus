using FluentEpplus.Common.Base;
using OfficeOpenXml;

namespace FluentEpplus.Common.Processor
{
    public interface IExcelHeaderProcessor
    {
        int ProcessHeader(ExcelWorksheet worksheet, int row, IExcelContainerCells container);
    }
}