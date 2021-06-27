using System;
using System.Drawing;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace FluentEpplus.Common.Style
{
    public interface IExcelCellStyleApplicator : IAsyncStyleApplicator, IExcelCellStyle {}

    public interface IExcelCellStyleApplicator<TDto> : IExcelCellStyleApplicator
    {
        void AttachConfiguration(Action<IExcelCellStyle, TDto> configuration);
    }

    public class ExcelCellStyleApplicator<TDto> : IExcelCellStyleApplicator<TDto>
    {
        private Action<IExcelCellStyle, TDto> _attachedConfiguration;
        public Color FontColor { get; set; } = Color.Empty;
        public Color BackgroundColor { get; set; } = Color.Empty;
        public bool IsBold { get; set; }
        public bool IsItalic { get; set; }

        public void AttachConfiguration(Action<IExcelCellStyle, TDto> configuration)
        {
            _attachedConfiguration = configuration;
        }

        public void Apply(ExcelRange cells, IExcelCellStyle style, object data = null)
        {
            _attachedConfiguration?.Invoke(style, (TDto) data);
            cells.Style.Font.Bold = style.IsBold;
            cells.Style.Font.Italic = style.IsItalic;

            if (style.FontColor != Color.Empty)
                cells.Style.Font.Color.SetColor(style.FontColor);

            if (style.BackgroundColor != Color.Empty)
            {
                cells.Style.Fill.PatternType = ExcelFillStyle.Solid;
                cells.Style.Fill.BackgroundColor.SetColor(style.BackgroundColor);
            }
        }
    }
    
    
}
