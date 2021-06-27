using System;
using System.Collections.Generic;
using System.IO.Compression;
using System.Text;
using FluentEpplus.Common.Style;
using OfficeOpenXml;

namespace FluentEpplus.Common.Base
{
    public interface IExcelCellConfigurationMappingFluent
    {
        IExcelCellConfigurationMappingFluent SetOrder(int columnOrder);
        IExcelCellConfigurationMappingFluent SetCaption(string caption);
        IExcelCellConfigurationMappingFluent SetStyle(Action<IExcelCellStyle> configuration);
        IExcelCellConfigurationMappingFluent SetHeaderStyle(Action<IExcelCellStyle> configuration);
        IExcelCellConfigurationMappingFluent AutoFit();
        IExcelCellConfigurationMappingFluent WordWrap();
    }

    public interface IExcelCell : IExcelCellConfigurationMappingFluent
    {
        IExcelContainerCells Parent { get; set; }
        int ColumnOrder { get; }
        string Caption { get; }
        bool IsAutoFit { get; }
        bool IsWordWrap { get; }

        int GetNumberParents();
        object GetValue(object fromData);

        void ApplyStyle(ExcelRange cells, object data);
        void ApplyHeaderStyle(ExcelRange cells);

    }
}
