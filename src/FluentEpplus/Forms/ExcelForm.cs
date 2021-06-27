using System;
using System.Collections.Generic;
using System.Text;
using FluentEpplus.Common;
using FluentEpplus.Common.Base;

namespace FluentEpplus.Forms
{
    public interface IExcelFormMappingFluent<TDto> : IExcelGroupCellMappingFluent<TDto>
    {
    }

    public interface IExcelFormConfigurationMappingFluent<TDto> : IExcelComplexGroupCellConfigurationMappingFluent
    {
    }

    public interface IExcelForm<TDto> : IExcelTable, IExcelFormMappingFluent<TDto>,
        IExcelFormConfigurationMappingFluent<TDto>
    {
    }

    public class ExcelForm<TDto> : ExcelGroupCell<TDto>, IExcelForm<TDto>
    {
        public ExcelForm(IExcelContainerCells parent) : base(parent)
        {
        }

        public ExcelForm() : base(null)
        {
            IsShowHeaderPerRow = false;
        }

        public IExcelComplexGroupCellConfigurationMappingFluent ShowHeaderPerRow()
        {
            IsShowHeaderPerRow = true;
            return this;
        }

        public bool IsShowHeaderPerRow { get; set;  }
    }
    
    
}
