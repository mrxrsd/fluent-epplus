using System;
using System.Collections.Generic;
using System.Text;

namespace FluentEpplus.Tables
{
    public interface IExcelColumn
    {
        int GetHeight();
        string HeaderCaption { get; set; }
    }
}
