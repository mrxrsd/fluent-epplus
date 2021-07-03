using System;
using System.Dynamic;
using FluentEpplus.Common.Base;
using FluentEpplus.Common.Style;
using System.Linq.Expressions;
using FluentEpplus.Common.Extractor;
using OfficeOpenXml;

namespace FluentEpplus.Common
{
    public interface IExcelSingleCellMappingFluent<TDto>
    {
        IExcelSingleCellConfigurationMappingFluent<TDto> MapProperty<TValue>(Expression<Func<TDto, TValue>> propertyExpression, Expression<Func<TValue, object>> formatExpression = null);
    }

    public interface IExcelSingleCellConfigurationMappingFluent<Dto> : IExcelCellConfigurationMappingFluent
    {
        IExcelSingleCellConfigurationMappingFluent<Dto> SetStyle(Action<IExcelCellStyle, Dto> configuration);
        new IExcelSingleCellConfigurationMappingFluent<Dto> SetStyle(Action<IExcelCellStyle> configuration);
        new IExcelSingleCellConfigurationMappingFluent<Dto> SetHeaderStyle(Action<IExcelCellStyle> configuration);
        new IExcelSingleCellConfigurationMappingFluent<Dto> SetOrder(int columnOrder);
        new IExcelSingleCellConfigurationMappingFluent<Dto> SetCaption(string caption);
        new IExcelSingleCellConfigurationMappingFluent<Dto> AutoFit();
        new IExcelSingleCellConfigurationMappingFluent<Dto> WordWrap();
    }

    public interface IExcelSingleCell<TDto> : IExcelCell, IExcelSingleCellMappingFluent<TDto>, IExcelSingleCellConfigurationMappingFluent<TDto>
    {
        void BuildDataExtractor<TValue>(Expression<Func<TDto, TValue>> propertyExpression, Expression<Func<TValue, object>> formatExpression = null);
    }


    public class ExcelSingleCell<TDto> : IExcelSingleCell<TDto>
    {
        protected IDataExtractor Extractor { get; set; }
        public IExcelContainerCells Parent { get; set; }
        public string Caption { get; set; }
        public int ColumnOrder { get; set; }
        public bool IsAutoFit { get; set; }
        protected IExcelCellStyleApplicator<TDto> ValueStyleApplicator { get; private set; }
        protected IExcelCellStyleApplicator HeaderStyleApplicator { get; private set; }
        public bool IsWordWrap { get; private set; }

        public ExcelSingleCell(IExcelContainerCells parent)
        {
            
            Parent = parent;
            if (Parent != null)
            {
                if (parent.IsAutoIncrementOrder)
                    ColumnOrder = Parent != null ? Parent.GetMaxOrder() + 1 : 1;
                else
                {
                    ColumnOrder = Parent.ColumnOrder + GetNumberParents();
                }
            }
        }

        public virtual void BuildDataExtractor<TValue>(Expression<Func<TDto, TValue>> propertyExpression, Expression<Func<TValue, object>> formatExpression = null)
        {
            Extractor = new DataExtractor<TDto, TValue>(propertyExpression);
            if (formatExpression != null)
                Extractor = new DecoratorFormatExtractor<TDto, TValue>(Extractor, formatExpression);
        }

        public object GetValue(object data)
        {
            return Extractor.GetValue(data);
        }

        public IExcelCellConfigurationMappingFluent SetOrder(int columnOrder)
        {
            ColumnOrder = columnOrder;
            return this;
        }

        public IExcelCellConfigurationMappingFluent SetCaption(string caption)
        {
            Caption = caption;
            return this;
        }

        public virtual IExcelSingleCellConfigurationMappingFluent<TDto> MapProperty<TValue>(Expression<Func<TDto, TValue>> propertyExpression, Expression<Func<TValue, object>> formatExpression = null)
        {
            BuildDataExtractor(propertyExpression, formatExpression);
            return this;
        }
        public int GetNumberParents()
        {
            int height = 0;
            var upParent = Parent;
            while (upParent != null)
            {
                height++;
                upParent = upParent.Parent;
            }

            return height;
        }

        public IExcelCellConfigurationMappingFluent AutoFit()
        {
            IsAutoFit = true;
            return this;
        }

        public IExcelCellConfigurationMappingFluent SetStyle(Action<IExcelCellStyle> configuration)
        {
            ValueStyleApplicator = new ExcelCellStyleApplicator<TDto>();
            configuration.Invoke(ValueStyleApplicator);
            return this;
        }

        public void ApplyStyle(ExcelRange cells, object data)
        {
            ValueStyleApplicator?.Apply(cells, ValueStyleApplicator, data);
        }

        public IExcelSingleCellConfigurationMappingFluent<TDto> SetStyle(Action<IExcelCellStyle, TDto> configuration)
        {
            ValueStyleApplicator = new ExcelCellStyleApplicator<TDto>();
            ValueStyleApplicator.AttachConfiguration(configuration);
            return this;
        }

        public void ApplyHeaderStyle(ExcelRange cells)
        {
            this.HeaderStyleApplicator?.Apply(cells, HeaderStyleApplicator);
        }
        public IExcelCellConfigurationMappingFluent SetHeaderStyle(Action<IExcelCellStyle> configuration)
        {
            HeaderStyleApplicator = new ExcelCellStyleApplicator<TDto>();
            configuration.Invoke(HeaderStyleApplicator);
            return this;
        }

        IExcelSingleCellConfigurationMappingFluent<TDto> IExcelSingleCellConfigurationMappingFluent<TDto>.SetStyle(Action<IExcelCellStyle> configuration)
        {
            return (IExcelSingleCellConfigurationMappingFluent<TDto>) SetStyle(configuration);
        }

        IExcelSingleCellConfigurationMappingFluent<TDto> IExcelSingleCellConfigurationMappingFluent<TDto>.SetHeaderStyle(Action<IExcelCellStyle> configuration)
        {
            return (IExcelSingleCellConfigurationMappingFluent<TDto>)SetHeaderStyle(configuration);
        }

        IExcelSingleCellConfigurationMappingFluent<TDto> IExcelSingleCellConfigurationMappingFluent<TDto>.SetOrder(int columnOrder)
        {
            return (IExcelSingleCellConfigurationMappingFluent<TDto>) SetOrder(columnOrder);
        }

        IExcelSingleCellConfigurationMappingFluent<TDto> IExcelSingleCellConfigurationMappingFluent<TDto>.SetCaption(string caption)
        {
            return (IExcelSingleCellConfigurationMappingFluent<TDto>)SetCaption(caption);
        }

        IExcelSingleCellConfigurationMappingFluent<TDto> IExcelSingleCellConfigurationMappingFluent<TDto>.AutoFit()
        {
            return (IExcelSingleCellConfigurationMappingFluent<TDto>) AutoFit();


        }

        IExcelSingleCellConfigurationMappingFluent<TDto> IExcelSingleCellConfigurationMappingFluent<TDto>.WordWrap()
        {
            return (IExcelSingleCellConfigurationMappingFluent<TDto>) WordWrap();
        }

        public IExcelCellConfigurationMappingFluent WordWrap()
        {
            IsWordWrap = true;
            return this;
        }




    }
}
