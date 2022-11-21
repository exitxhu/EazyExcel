using ClosedXML;
using ClosedXML.Excel;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.EMMA;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Data;
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
            //worksheet.Columns().AdjustToContents();
            workbook.SaveAs(mem);
        }
        public static IXLWorkbook ToExcelWorkbook(this byte[] source)
        {
            var result = new XLWorkbook(new MemoryStream(source));
            return result;
        }
        public static List<T> ToList<T>(this IXLWorksheet ws) where T : new()
        {
            var result = new List<T>();
            var props = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance);
            var tits = new List<(string, int, PropertyInfo)>();
            for (int i = 0; i < props.Count(); i++)
            {
                var contit = ws.Cell(1, i + 1).Value?.ToString()!;
                var t = props.SingleOrDefault(n => n.Name == contit);
                if (t is not null)
                    tits.Add((contit, i + 1, t));

            }
            for (int i = 2; i <= ws.LastRowUsed().RowNumber(); i++)
            {
                T temp = new();
                var c = false;
                foreach (var tit in tits)
                {
                    var fh = ws.Cell(i, tit.Item2).Value;
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
                if (conversionType.IsEnum)
                {
                    result = Enum.Parse(conversionType, value.ToString());
                }
                else
                    result = Convert.ChangeType(value, conversionType);
                return true;
            }
            catch
            {
                result = null;
                return false;
            }
        }

        public static DataTable CreateDataTableFromAnyCollection<T>(IEnumerable<T> list)
        {
            Type type = typeof(T);
            var properties = type.GetProperties();

            DataTable dataTable = new DataTable();
            foreach (PropertyInfo info in properties)
            {
                dataTable.Columns.Add(new DataColumn(info.Name, Nullable.GetUnderlyingType(info.PropertyType) ?? info.PropertyType));
            }

            foreach (T entity in list)
            {
                object[] values = new object[properties.Length];
                for (int i = 0; i < properties.Length; i++)
                {
                    values[i] = properties[i].GetValue(entity, null);
                }

                dataTable.Rows.Add(values);
            }

            return dataTable;
        }


        /// <summary>
        /// build an excel file as stream
        /// </summary>
        /// <param name="mem"></param>
        /// <param name="sheetName"></param>
        public static void ToExcelTable<T>(this IEnumerable<T> @this, Stream mem, string sheetName = "Sample Sheet")
        {

            using var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add(sheetName);


            var properties = (typeof(T)).GetProperties()
                .Select(a => new
                {
                    Attribute = (typeof(T).GetProperty(a.Name).GetCustomAttributes(false).Where(x => x.GetType().Name == nameof(ColumnNameOrderAttribute)).FirstOrDefault() as ColumnNameOrderAttribute) ?? new ColumnNameOrderAttribute(a.Name, 0),
                    Property = a
                })

                .OrderBy(a => a.Attribute.Order)
                .ToList();


            DataTable dataTable = new DataTable();

            foreach (var info in properties)
            {
                string displayName = "";
                if (info.Attribute != null && info.Attribute.ResourceType != null)
                {
                    var resourceType = (info.Attribute.ResourceType != null) ? (Type)info.Attribute.ResourceType : null;
                    var decorationResx = new ComponentResourceManager(resourceType);
                    displayName = decorationResx.GetString(info.Attribute.DisplayName);
                    if (string.IsNullOrEmpty(displayName))
                    {
                        displayName = decorationResx.GetString(info.Attribute.DisplayName.Replace("_", " "));
                    }
                }
                else
                {
                    displayName = info.Attribute.DisplayName;
                }


                if (info.Property.PropertyType.BaseType == typeof(Enum))
                {
                    dataTable.Columns.Add(new DataColumn(displayName));
                }
                else
                    dataTable.Columns.Add(new DataColumn(displayName,
                        Nullable.GetUnderlyingType(info.Property.PropertyType) ?? info.Property.PropertyType));
            }

            foreach (var entity in @this)
            {
                object[] values = new object[properties.Count];
                for (int i = 0; i < properties.Count; i++)
                {
                    if (properties[i].Property.PropertyType.BaseType == typeof(Enum))
                    {
                        var enumValue = GetDescription(properties[i].Property.PropertyType, properties[i].Property.GetValue(entity, null).ToString()) ?? properties[i].Property.GetValue(entity, null).ToString();

                        values[i] = enumValue;
                    }
                    else
                    {
                        values[i] = properties[i].Property.GetValue(entity, null);
                    }

                }

                dataTable.Rows.Add(values);
            }
            worksheet.Cell(1, 1).InsertTable(dataTable);
            worksheet.Columns().AdjustToContents();
            workbook.SaveAs(mem);
        }



        static string GetDescription(Type enumType, string field)
        {
            FieldInfo fieldInfo = enumType.GetField(field);
            if (fieldInfo == null)
                return string.Empty;

            foreach (var attribute in fieldInfo.GetCustomAttributes())
            {
                if (attribute == null)
                    continue;
                if (attribute is DescriptionAttribute descAtt)
                    return descAtt.Description;
                else if (attribute.ToString().IndexOf("Display", StringComparison.Ordinal) > 0)
                {
                    var prop = attribute.GetType().GetProperty("Name");
                    if (prop == null)
                        continue;
                    return Convert.ToString(prop.GetValue(attribute));
                }
            }

            return null;
        }

    }
}
