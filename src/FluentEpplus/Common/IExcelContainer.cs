using OfficeOpenXml;

namespace FluentEpplus.Common
{
    public interface IExcelContainer
    {
        int? StartRow { get; set; }
        void Process(ExcelWorksheet worksheet, bool printHeaders = true);
        void Validate();
    }
}