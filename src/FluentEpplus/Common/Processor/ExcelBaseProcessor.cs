using System.Collections;
using FluentEpplus.Common.Base;
using OfficeOpenXml;

namespace FluentEpplus.Common.Processor
{
    public class ExcelBaseProcessor : IExcelProcessor
    {
        private IExcelHeaderProcessor Header { get; set; }
        private IExcelRowProcessor Row { get; set; }

        public ExcelBaseProcessor(IExcelHeaderProcessor header, IExcelRowProcessor row)
        {
            Header = header;
            Row = row;
        }

        public void Process(ExcelWorksheet worksheet, IExcelContainerCells container, object data)
        {
            var startRow = Header.ProcessHeader(worksheet, container.StartAt.HasValue ? container.StartAt.Value : 1, container) + 1;
            var enumerableData = (data as IEnumerable);
            if (enumerableData != null)
            {
                Row.ProcessRowFromEnumerable(worksheet, startRow, container, enumerableData);
            }
            else
            {
                Row.ProcessRow(worksheet, startRow, container, data);
            }
        }

       
    }
}
