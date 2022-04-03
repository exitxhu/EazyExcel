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
        public static List<T> ToList<T>(this IXLWorksheet ws) where T : new()
        {
            var result = new List<T>();


            var props = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance);
            var tits = new List<(string, int, PropertyInfo)>();
            for (int i = 0; i < props.Count(); i++)
            {
                var contit = ws.Cell(1, i + 1).Value?.ToString()!;
                if (props.Any(n => n.Name == contit))
                    tits.Add((contit, i + 1, props[i]));

            }
            for (int i = 1; i < ws.LastRowUsed().RowNumber(); i++)
            {
                T temp = new();
                var c = false;
                foreach (var tit in tits)
                {
                    var fh = ws.Cell(i + 1, tit.Item2).Value;
                    if (TryChangeType(fh, Nullable.GetUnderlyingType(tit.Item3.PropertyType) ?? tit.Item3.PropertyType, out var obje))
                    {
                        c = true;
                        tit.Item3.SetValue(temp, obje);
                    }
                }
                if (c)
                    result.Add(temp);
            }
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