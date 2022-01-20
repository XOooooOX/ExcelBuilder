using Aspose.Cells;
using ExcelBuilder.Attributes;
using System.Drawing;
using System.Reflection;

namespace ExcelBuilder.Helper
{
    public static class ExcelBuilder
    {
        public const string EXCEL_CONTENT_TYPE = "application/vnd.ms-exel";
        public const string EXCEL_EXTENSION = ".xls";

        public static (MemoryStream Stream, byte[] File, string ContentType, string Extension) ExportExcel<T>(
        this IList<T> list,
        int rowPadding = 0,
        int columnPadding = 0,
        Color borderColor = default,
        bool displayRightToLeft = true)
        {
            if (!list.Any())
                throw new NullReferenceException();

            var props = (typeof(T).GetProperties()).ToList();

            RemoveExcelIgnore(props);

            Workbook workbook = new();
            Worksheet cellSheet = workbook.Worksheets[0];

            cellSheet.DisplayRightToLeft = displayRightToLeft;

            SetDefualtSheetStyle(workbook);

            int column = columnPadding;

            int row = rowPadding;

            foreach (PropertyInfo prop in props)
            {
                SetCellValue(cellSheet, row, column, GetColumnName(prop), borderColor);
                column++;
            }

            row++;

            foreach (T item in list)
            {
                column = columnPadding;

                foreach (PropertyInfo prop in props)
                {
                    SetCellValue(cellSheet, row, column, prop.GetValue(item), borderColor);

                    if (list.Last().Equals(item) && IsExcelTotla(prop))
                    {
                        var total = list.Sum(o
                            => Convert.ToDecimal(o.GetType().GetProperty(prop.Name).GetValue(o, null)));

                        SetCellValue(cellSheet, row + 1, column, total, borderColor);
                    }
                    column++;
                }
                row++;
            }

            cellSheet.AutoFitColumns();
            using MemoryStream memoryStream = new();
            workbook.Save(memoryStream, SaveFormat.Excel97To2003);
            return (memoryStream, memoryStream.ToArray(), EXCEL_CONTENT_TYPE, EXCEL_EXTENSION);
        }

        private static string GetColumnName(PropertyInfo prop)
                                        => prop.GetCustomAttributes(typeof(ExcelDisplayName), false).Length > 0
                                        ? ((ExcelDisplayName)prop.GetCustomAttributes(typeof(ExcelDisplayName), false)[0]).Title
                                        : prop.Name;

        private static bool IsExcelTotla(PropertyInfo prop)
                                        => prop.GetCustomAttributes(typeof(ExcelTotal), false).Length > 0;

        private static void SetCellValue(Worksheet worksheet, int row, int column, object value, Color borderColor)
        {
            worksheet.Cells[row, column].PutValue(value);

            worksheet.Cells[row, column].SetStyle(SetCellColor(worksheet.Cells[row, column].GetStyle(), borderColor));
        }

        private static void RemoveExcelIgnore(List<PropertyInfo> props)
        {
            foreach (PropertyInfo prop in props.Reverse<PropertyInfo>())
                if (prop.GetCustomAttributes(typeof(ExcelIgnore), false).Length > 0)
                    _ = props.Remove(prop);
        }

        private static Style SetCellColor(Style style, Color color)
        {
            _ = style.SetBorder(BorderType.BottomBorder, CellBorderType.Thin, color);

            _ = style.SetBorder(BorderType.LeftBorder, CellBorderType.Thin, color);

            _ = style.SetBorder(BorderType.RightBorder, CellBorderType.Thin, color);

            _ = style.SetBorder(BorderType.TopBorder, CellBorderType.Thin, color);

            return style;
        }

        private static void SetDefualtSheetStyle(Workbook workbook)
        {
            Style style = workbook.DefaultStyle;

            style.TextDirection = TextDirectionType.RightToLeft;

            style.HorizontalAlignment = TextAlignmentType.Center;

            style.IsTextWrapped = true;
        }
    }
}
