using System;
using System.Collections.Generic;
using System.Text;
using FluentEpplus.Common.Base;
using FluentEpplus.Common.Processor;
using FluentEpplus.Forms;
using FluentEpplus.Tables;
using OfficeOpenXml;

namespace FluentEpplus
{
    public interface IExcelWorksheetFluent
    {
        IExcelWorksheetFluent Label(string label);
        IExcelWorksheetFluent BlankAsNew();
        IExcelTableMappingFluent<T> AddTable<T>(Action<IExcelRootTableConfigurationMappingFluent<T>> relationship);
        IExcelFormMappingFluent<T> AddForm<T>(Action<IExcelRootFormConfigurationMappingFluent<T>> relationship);

    }

    public interface IExcelWorksheet
    {
        string WorksheetLabel { get; set; }
        void Process(ExcelWorksheet worksheet);
        void Validate();
    }

    public class ExcelMapWorksheet : IExcelWorksheet, IExcelWorksheetFluent
    {
        private List<IExcelRootContainerCells> MappedExcelContainers { get; set; }

        public ExcelMapWorksheet()
        {
            MappedExcelContainers = new List<IExcelRootContainerCells>();
        }

        public string WorksheetLabel { get; set; }
        public bool HideGridLines { get; set; }

        public void Process(ExcelWorksheet worksheet)
        {
            if (HideGridLines) worksheet.View.ShowGridLines = false;

            foreach (var rootContainer in MappedExcelContainers)
            {
                var processor = ExcelProcessorBuilder
                    .GetDefaultBuilder()
                    .WithRootContainer(rootContainer)
                    .Build();
                processor.Process(worksheet, rootContainer, rootContainer.Data);
            }
        }

        public void Validate()
        {

        }

        public IExcelWorksheetFluent Label(string label)
        {
            WorksheetLabel = label;
            return this;
        }

        public IExcelWorksheetFluent BlankAsNew()
        {
            HideGridLines = true;
            return this;
        }

        public IExcelTableMappingFluent<T> AddTable<T>(Action<IExcelRootTableConfigurationMappingFluent<T>> relationship)
        {
            var newTable = new ExcelRootTable<T>();
            relationship(newTable);
            MappedExcelContainers.Add(newTable);
            return newTable;
        }

        public IExcelFormMappingFluent<T> AddForm<T>(Action<IExcelRootFormConfigurationMappingFluent<T>> relationship)
        {
            var newTable = new ExcelRootForm<T>();
            relationship(newTable);
            MappedExcelContainers.Add(newTable);
            return newTable;
        }
    }
}
