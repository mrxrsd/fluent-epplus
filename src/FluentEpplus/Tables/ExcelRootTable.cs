using System;
using System.Collections.Generic;
using System.Text;
using FluentEpplus.Common.Base;
using FluentEpplus.Tables.Processor;
using OfficeOpenXml.Table;

namespace FluentEpplus.Tables
{
    public interface IExcelRootTableConfigurationMappingFluent<TDto> : IExcelTableConfigurationMappingFluent<TDto>
    {
        IExcelRootTableConfigurationMappingFluent<TDto> Bind(object fromData);
    }

    public interface IExcelRootTable<TDto> : IExcelTable<TDto>, IExcelRootTable, IExcelRootTableConfigurationMappingFluent<TDto>
    {
    }


    public class ExcelRootTable<TDto> : ExcelTable<TDto>, IExcelRootTable<TDto>
    {
        public Type ComplexProcessorType { get; set; }
        public object Data { get; set; }
        public Type SimpleProcessorType { get; set; }

        public ExcelRootTable() : base()
        {
            IsAutoIncrementOrder = true;
            SimpleProcessorType = typeof(ExcelSimpleTableProcessor);
            ComplexProcessorType = typeof(ExcelComplexTableProcessor);
        }

        public IExcelRootTableConfigurationMappingFluent<TDto> Bind(object fromData)
        {
            Data = fromData;
            return this;
        }
    }
}
