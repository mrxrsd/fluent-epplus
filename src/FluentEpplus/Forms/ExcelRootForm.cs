using System;
using System.Collections.Generic;
using System.Text;
using FluentEpplus.Common.Base;
using FluentEpplus.Forms.Processor;

namespace FluentEpplus.Forms
{

    public interface IExcelRootFormConfigurationMappingFluent<TDto> : IExcelFormConfigurationMappingFluent<TDto>
    {
        IExcelRootFormConfigurationMappingFluent<TDto> Bind(object fromData);
    }

    public interface IExcelRootForm<TDto> : IExcelForm<TDto>, IExcelRootForm, IExcelRootFormConfigurationMappingFluent<TDto>
    {

    }

    public class ExcelRootForm<TDto> : ExcelForm<TDto>, IExcelRootForm<TDto>
    {
        public object Data { get; set; }
        public Type SimpleProcessorType { get; set; }
        public Type ComplexProcessorType { get; set; }

        public ExcelRootForm() : base()
        {
            IsAutoIncrementOrder = false;
            SimpleProcessorType = typeof(ExcelComplexFormProcessor);
            ComplexProcessorType = typeof(ExcelComplexFormProcessor);
        }

        public IExcelRootFormConfigurationMappingFluent<TDto> Bind(object fromData)
        {
            Data = fromData;
            return this;
        }
    }
}
