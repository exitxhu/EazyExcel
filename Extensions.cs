using ClosedXML.Excel;
using System.Reflection;

namespace EazyExcel
{
    public static class Extensions
    {
        /// <summary>
        /// build an excel file and save it to disk
        /// </summary>
        /// <param name="filePathAndName"></param>
        /// <param name="sheetName"></param>
        public static void ToExcel<T>(this IEnumerable<T> @this, string filePathAndName = "SampleWorkbook.xlsx", string sheetName = "Sample Sheet")
        {
            using var workbook = new XLWorkbook();
            using var mem = new MemoryStream();
            var worksheet = workbook.Worksheets.Add(sheetName);

            worksheet.Cell(1, 1).InsertTable(@this);
            worksheet.Columns().AdjustToContents();
            workbook.SaveAs(filePathAndName);
        }
        /// <summary>
        /// build an excel file as stream
        /// </summary>
        /// <param name="mem"></param>
        /// <param name="sheetName"></param>
        public static void ToExcel<T>(this IEnumerable<T> @this, Stream mem, string sheetName = "Sample Sheet")
        {
            using var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add(sheetName);

            worksheet.Cell(1, 1).InsertTable(@this);
            worksheet.Columns().AdjustToContents();
            workbook.SaveAs(mem);
        }
        public static IXLWorkbook ToExcelWorkbook<T>(this byte[] source) where T : new()
        {
            var result = new XLWorkbook(new MemoryStream(source));
            return result;
        }
        static bool TryChangeType(object? value, Type conversionType, out object result)
        {
            try
            {
                result = Convert.ChangeType(value, conversionType);
                return true;
            }
            catch
            {
                result = null;
                return false;
            }
        }
    }
}