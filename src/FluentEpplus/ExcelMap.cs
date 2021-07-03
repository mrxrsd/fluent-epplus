using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;
using OfficeOpenXml;

namespace FluentEpplus
{
    public class ExcelMap
    {
        private List<IExcelWorksheet> MappedWorksheets { get; set; }

        public ExcelMap()
        {
            MappedWorksheets = new List<IExcelWorksheet>();
        }

        public IExcelWorksheetFluent AddWorksheet()
        {
            var newWs = new ExcelMapWorksheet();
            MappedWorksheets.Add(newWs);
            return newWs;
        }

        private void Validate()
        {
            foreach (var mappedWorksheet in MappedWorksheets)
            {
                mappedWorksheet.Validate();
            }
        }

        public ExcelPackage GetExcelPackage()
        {
            ExcelPackage pack = null;
            try
            {
                pack = new ExcelPackage();
                foreach (var mappedWorksheet in MappedWorksheets)
                {
                    var wsLabel = (string.IsNullOrEmpty(mappedWorksheet.WorksheetLabel))
                        ? String.Format("extract{0}", MappedWorksheets.IndexOf(mappedWorksheet))
                        : mappedWorksheet.WorksheetLabel;

                    var ws = pack.Workbook.Worksheets.Add(wsLabel);

                    mappedWorksheet.Process(ws);
                }
            }finally{}

            return pack;
        }

        private byte[] Process()
        {
            byte[] result = null;
            var pack = GetExcelPackage();
            result = pack.GetAsByteArray();

            return result;
        }
    }
}
