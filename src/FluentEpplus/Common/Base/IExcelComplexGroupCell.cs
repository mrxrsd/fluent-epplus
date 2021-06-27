namespace FluentEpplus.Common.Base
{
    public interface IExcelComplexGroupCellConfigurationMappingFluent : IExcelGroupCellConfigurationMappingFluent
    {
        IExcelComplexGroupCellConfigurationMappingFluent ShowHeaderPerRow();
    }

    public interface IExcelComplexGroupCell : IExcelGroupCell, IExcelComplexGroupCellConfigurationMappingFluent
    {
        bool IsShowHeaderPerRow { get; }
    }
}