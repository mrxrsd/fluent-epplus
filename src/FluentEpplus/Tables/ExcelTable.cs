using System;
using System.Collections.Generic;
using System.Linq.Expressions;
using System.Text;
using FluentEpplus.Common;
using FluentEpplus.Common.Base;

namespace FluentEpplus.Tables
{
    public interface IExcelTableMappingFluent<TDto> : IExcelGroupCellMappingFluent<TDto>
    {
        IExcelTableMappingFluent<TDto> MapTable(Action<IExcelTableMappingFluent<TDto>> relationship);
        IExcelTableMappingFluent<TNewDto> MapTable<TNewDto>(Expression<Func<TDto, TNewDto>> propertyExpression, Action<IExcelTableMappingFluent<TNewDto>> relationship);
        IExcelTableMappingFluent<TNewDto> MapTable<TNewDto>(Expression<Func<TDto, List<TNewDto>>> propertyExpression, Action<IExcelTableMappingFluent<TNewDto>> relationship);
    }

    public interface IExcelTableConfigurationMappingFluent<TDto> : IExcelComplexGroupCellConfigurationMappingFluent {}

    public interface IExcelTable<TDto> : IExcelTable, IExcelTableMappingFluent<TDto>,
        IExcelTableConfigurationMappingFluent<TDto>
    {

    }

    public class ExcelTable<TDto> : ExcelGroupCell<TDto>, IExcelTable<TDto>
    {
        public bool IsShowHeaderPerRow { get; set; }

        public ExcelTable(IExcelContainerCells parent) : base(parent)
        {
        }

        public ExcelTable() : base(null)
        {
            IsShowHeaderPerRow = false;
            ShowCaption = false;
        }

        public IExcelComplexGroupCellConfigurationMappingFluent ShowHeaderPerRow()
        {
            IsShowHeaderPerRow = true;
            return this;
        }


        public IExcelTableMappingFluent<TDto> MapTable(Action<IExcelTableMappingFluent<TDto>> relationship)
        {
            var newTable = new ExcelTable<TDto>(this);
            newTable.BuildDataExtractor(data=>data);
            relationship(newTable);
            Cells.Add(newTable);

            return newTable;
        }

        public IExcelTableMappingFluent<TNewDto> MapTable<TNewDto>(Expression<Func<TDto, TNewDto>> propertyExpression, Action<IExcelTableMappingFluent<TNewDto>> relationship)
        {
            var newTable = new ExcelTable<TNewDto>(this);
            newTable.BuildDataExtractor(propertyExpression);
            relationship(newTable);
            Cells.Add(newTable);

            return newTable;
        }

        public IExcelTableMappingFluent<TNewDto> MapTable<TNewDto>(Expression<Func<TDto, List<TNewDto>>> propertyExpression, Action<IExcelTableMappingFluent<TNewDto>> relationship)
        {
            var newTable = new ExcelTable<TNewDto>(this);
            newTable.BuildDataExtractor(propertyExpression);
            relationship(newTable);
            Cells.Add(newTable);

            return newTable;
        }
    }
}
