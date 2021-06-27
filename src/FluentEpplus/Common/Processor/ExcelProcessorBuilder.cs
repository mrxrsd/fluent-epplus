using System;
using System.Collections.Generic;
using System.Text;
using System.Linq;
using FluentEpplus.Common.Base;

namespace FluentEpplus.Common.Processor
{
    public interface IExcelProcessorBuilder
    {
        IExcelProcessorBuilder WithRootContainer(IExcelRootContainerCells container);
        IExcelProcessor Build();
    }

    public class ExcelProcessorBuilder : IExcelProcessorBuilder
    {
        private IExcelHeaderProcessor _header;
        private IExcelRowProcessor _row;
        private Type _simpleProcessorType;
        private Type _complexProcessorType;

        private ExcelProcessorBuilder() {}

        public IExcelProcessor Build()
        {
            return new ExcelBaseProcessor(_header, _row);
        }

        public IExcelProcessorBuilder WithRootContainer(IExcelRootContainerCells container)
        {
            _complexProcessorType = container.ComplexProcessorType;
            _simpleProcessorType  = container.SimpleProcessorType;

            _header = BuildHeader(container.Cells);
            _row = BuildRow(container.Cells);

            return this;
        }

        private IExcelHeaderProcessor BuildHeader(List<IExcelCell> properties)
        {
            var hasGroupProperty = (properties != null && properties.Count(p => p is IExcelGroupCell) > 0);
            return (IExcelHeaderProcessor) Activator.CreateInstance(hasGroupProperty ? _complexProcessorType : _simpleProcessorType);
        }

        private IExcelRowProcessor BuildRow(List<IExcelCell> properties)
        {
            var hasComplexGroupProperty = (properties != null && properties.Count(p => p is IExcelComplexGroupCell) > 0);
            return (IExcelRowProcessor)Activator.CreateInstance(hasComplexGroupProperty ? _complexProcessorType : _simpleProcessorType);
        }

        public static IExcelProcessorBuilder GetDefaultBuilder()
        {
            return new ExcelProcessorBuilder();
        }
    }
}
