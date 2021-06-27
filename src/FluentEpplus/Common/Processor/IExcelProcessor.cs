using FluentEpplus.Common.Base;
using OfficeOpenXml;

namespace FluentEpplus.Common.Processor
{
    public interface IExcelProcessor
    {
        void Process(ExcelWorksheet worksheet, IExcelContainerCells container, object data);
    }
}