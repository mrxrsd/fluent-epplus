using OfficeOpenXml.ConditionalFormatting.Contracts;

namespace FluentEpplus.Common.Base
{
    public interface IExcelGroupCellConfigurationMappingFluent : IExcelContainerCellsConfigurationMappingFluent
    {
    }


    public interface IExcelGroupCell : IExcelContainerCells, IExcelGroupCellConfigurationMappingFluent
    {
        
    }
}