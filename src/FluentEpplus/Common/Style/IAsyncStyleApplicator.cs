using OfficeOpenXml;

namespace FluentEpplus.Common.Style
{
    public interface IAsyncStyleApplicator
    {
        void Apply(ExcelRange cells, IExcelCellStyle style, object data = null);
    }
}