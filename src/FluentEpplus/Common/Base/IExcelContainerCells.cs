using System;
using System.Collections.Generic;

namespace FluentEpplus.Common.Base
{
    public interface IExcelContainerCellsConfigurationMappingFluent : IExcelCellConfigurationMappingFluent
    {
        IExcelContainerCellsConfigurationMappingFluent StartAtRow(int startRow);
        IExcelContainerCellsConfigurationMappingFluent HideCaption();
    }

    public interface IExcelContainerCells : IExcelCell, IExcelContainerCellsConfigurationMappingFluent
    {
        bool IsAutoIncrementOrder { get; set; }
        int? StartAt { get; }
        bool ShowCaption { get; }
        List<IExcelCell> Cells { get; }

        int GetMaxOrder();
    }

    public interface IExcelRootContainerCells : IExcelContainerCells
    {
        object Data { get; }
        Type SimpleProcessorType { get; }
        Type ComplexProcessorType { get; }
    }
}