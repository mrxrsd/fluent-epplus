using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using FluentEpplus.Common.Base;
using FluentEpplus.Common.Extractor;

namespace FluentEpplus.Common
{
    public interface IExcelGroupCellMappingFluent<TDto> : IExcelSingleCellMappingFluent<TDto>
    {
        IExcelGroupCellConfigurationMappingFluent<TDto> MapGroup(Action<IExcelGroupCellMappingFluent<TDto>> relationship);
        IExcelGroupCellConfigurationMappingFluent<TNewDto> MapGroup<TNewDto>(Expression<Func<TDto, TNewDto>> propertyExpression, Action<IExcelGroupCellMappingFluent<TNewDto>> relationship);
    }

    public interface IExcelGroupCellConfigurationMappingFluent<TDto> : IExcelGroupCellConfigurationMappingFluent
    {
        
    }

    public interface IExcelGroupCell<TDto> : IExcelGroupCell, IExcelGroupCellConfigurationMappingFluent<TDto>, IExcelGroupCellMappingFluent<TDto>
    {
        void BuildDataExtractor<TParentDto>(Expression<Func<TParentDto, TDto>> propertyExpression);
        void BuildDataExtractor<TParentDto>(Expression<Func<TParentDto, List<TDto>>> propertyExpression);
    }

    public class ExcelGroupCell<TDto> : ExcelSingleCell<TDto>, IExcelGroupCell<TDto>
    {
        public int? StartAt { get; set; }
        public bool IsAutoIncrementOrder { get; set; }
        public bool ShowCaption { get; set; }

        public List<IExcelCell> Cells { get; set; }


        public ExcelGroupCell(IExcelContainerCells parent) : base(parent)
        {
            if (Parent != null)
                IsAutoIncrementOrder = parent.IsAutoIncrementOrder;
            Cells = new List<IExcelCell>();
            ShowCaption = true;
            Extractor = new DataExtractor<TDto, TDto>(data => data);
        }

        public void BuildDataExtractor<TParentDto>(Expression<Func<TParentDto, TDto>> propertyExpression)
        {
            Extractor = new DataExtractor<TParentDto, TDto>(propertyExpression);
        }

        public void BuildDataExtractor<TParentDto>(Expression<Func<TParentDto, List<TDto>>> propertyExpression)
        {
            Extractor = new DataExtractor<TParentDto, List<TDto>>(propertyExpression);
        }


        public int GetMaxOrder()
        {
            var maxOrder = 0;
            if (Cells.Count > 0)
            {
                maxOrder = Cells.Max(p => p.ColumnOrder);
                foreach (var cell in Cells.OfType<IExcelContainerCells>())
                {
                    var currentMaxOrder = cell.GetMaxOrder();
                    maxOrder = maxOrder < currentMaxOrder ? currentMaxOrder : maxOrder;
                }
            }

            if (Parent != null && maxOrder == 0)
            {
                var currentMaxOrder = Parent.GetMaxOrder();
                maxOrder = maxOrder < currentMaxOrder ? currentMaxOrder : maxOrder;
            }

            return maxOrder;
        }

        public IExcelContainerCellsConfigurationMappingFluent StartAtRow(int startRow)
        {
            StartAt = startRow;
            return this;
        }

        public IExcelContainerCellsConfigurationMappingFluent HideCaption()
        {
            ShowCaption = false;
            return this;
        }

        public override IExcelSingleCellConfigurationMappingFluent<TDto> MapProperty<TValue>(Expression<Func<TDto, TValue>> propertyExpression, Expression<Func<TValue, object>> formatExpression = null)
        {
            var newCell = new ExcelSingleCell<TDto>(this);
            Cells.Add(newCell);
            return newCell.MapProperty(propertyExpression, formatExpression);
        }

        public IExcelGroupCellConfigurationMappingFluent<TDto> MapGroup(Action<IExcelGroupCellMappingFluent<TDto>> relationship)
        {
            var newGroupProperty = new ExcelGroupCell<TDto>(this);
            relationship(newGroupProperty);
            Cells.Add(newGroupProperty);

            return newGroupProperty;
        }

        public IExcelGroupCellConfigurationMappingFluent<TNewDto> MapGroup<TNewDto>(Expression<Func<TDto, TNewDto>> propertyExpression, Action<IExcelGroupCellMappingFluent<TNewDto>> relationship)
        {
            var newGroupProperty = new ExcelGroupCell<TNewDto>(this);
            newGroupProperty.BuildDataExtractor(propertyExpression);
            relationship(newGroupProperty);
            Cells.Add(newGroupProperty);

            return newGroupProperty;
        }


    }
    
    
}
