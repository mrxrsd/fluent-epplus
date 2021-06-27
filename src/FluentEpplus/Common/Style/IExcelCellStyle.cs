using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text;

namespace FluentEpplus.Common.Style
{
    public interface IExcelCellStyle
    {
        Color FontColor { get; set; }
        Color BackgroundColor { get; set; }
        bool IsBold { get; set; }
        bool IsItalic { get; set; }
    }
}
